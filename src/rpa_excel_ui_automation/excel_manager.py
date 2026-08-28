"""Gestión de la instancia de Excel y despliegue de los diálogos "Abrir"/"Guardar como".

Responsabilidad única: inicializar la aplicación, asegurar una cuadrícula/cinta
disponible, cerrar el modal de activación de Office si apareciera y desplegar los
diálogos nativos de archivo mediante estrategias en cascada:

1. Atajo global de teclado (Ctrl+F12 / F12), el enfoque canónico del README.
2. Cinta de Excel (Pestaña Archivo -> Abrir/Guardar como -> Examinar) vía
   InvokePattern (independiente del foco y del ratón).

Sin `time.sleep`: toda espera usa los mecanismos nativos de `uiautomation`
(`.Exists(maxSearchSeconds=...)`, `SetActive(waitTime=...)`, `SendKeys(waitTime=...)`).

El diálogo nativo (#32770) puede registrarse como ventana raíz o anidado bajo
XLMAIN en el árbol UIA; se busca con suficiente profundidad desde el escritorio.
"""

from __future__ import annotations

import ctypes
import logging
import subprocess
import time
from ctypes import wintypes
from pathlib import Path

import uiautomation as auto

logger = logging.getLogger(__name__)

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

HWND_TOPMOST = -1
HWND_NOTOPMOST = -2
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
SW_RESTORE = 9
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
WAIT_OBJECT_0 = 0

_WAIT_AFTER_CLOSE = 6          # segundos nativos esperando el cierre tras intentar WM_CLOSE
_CLOSE_BUDGET = 15             # presupuesto nativo total para el cierre elegante

EXCEL_MAIN_CLASS = "XLMAIN"
DIALOG_CLASS = "#32770"
DEFAULT_EXCEL_PATH = Path(r"C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE")

OPEN_SHORTCUT = "{Ctrl}{F12}"  # despliega el diálogo "Abrir"
SAVE_AS_SHORTCUT = "{F12}"     # despliega el diálogo "Guardar como"

_MAIN_TIMEOUT = 40    # segundos nativos esperando la ventana XLMAIN
_DIALOG_TIMEOUT = 8   # segundos nativos esperando el diálogo de archivo
_MODAL_TIMEOUT = 3   # segundos nativos sondeando el modal de activación (red de seguridad)
_MAX_DIALOG_DEPTH = 6  # profundidad de búsqueda de #32770 desde el escritorio


class ExcelManager:
    """Inicializa Excel y despliega los diálogos "Abrir" / "Guardar como"."""

    def __init__(self, excel_path: str | Path | None = None, initial_workbook: str | Path | None = None) -> None:
        self.excel_path = Path(excel_path) if excel_path else DEFAULT_EXCEL_PATH
        self.initial_workbook = Path(initial_workbook) if initial_workbook else None
        self._window: auto.WindowControl | None = None

    # ------------------------------------------------------------------ #
    #  API pública
    # ------------------------------------------------------------------ #

    def open_file(self) -> None:
        logger.info("[%s] Iniciando apertura", type(self).__name__)
        self._ensure_running()
        self._deploy_dialog(OPEN_SHORTCUT, "Abrir", "Abrir")
        logger.info("[%s] Apertura completada", type(self).__name__)

    def save_as(self) -> None:
        logger.info("[%s] Iniciando guardado", type(self).__name__)
        self._ensure_running()
        self._deploy_dialog(SAVE_AS_SHORTCUT, "Guardar como", "Guardar como")
        logger.info("[%s] Guardado completado", type(self).__name__)

    def close(self) -> None:
        """Cierra Excel solo cuando no quedan modales ni diálogos abiertos, y verifica
        la salida real del proceso. Si Excel no termina dentro del presupuesto nativo,
        finaliza el PID que arrancó esta instancia como último recurso."""
        if self._window is None:
            return
        logger.info("[%s] Cerrando Excel (modales/diálogos previos ya atendidos)", type(self).__name__)
        pid = self._hwnd_pid(self._window.NativeWindowHandle)
        deadline = time.monotonic() + _CLOSE_BUDGET
        while time.monotonic() < deadline:
            self._close_pending()
            try:
                self._window.GetWindowPattern().Close()
            except Exception as exc:  # noqa: BLE001
                logger.debug("[ExcelManager] WindowPattern.Close: %s", exc)
            if not self._window.Exists(maxSearchSeconds=_WAIT_AFTER_CLOSE):
                if self._process_exited(pid, _WAIT_AFTER_CLOSE):
                    logger.info("[%s] Excel cerrado", type(self).__name__)
                else:
                    self._terminate_pid(pid)
                    logger.info("[%s] Excel cerrado (proceso finalizado)", type(self).__name__)
                return
        self._close_pending()
        try:
            self._window.GetWindowPattern().Close()
        except Exception:  # noqa: BLE001
            pass
        self._terminate_pid(pid)
        logger.info("[%s] Excel cerrado (proceso finalizado tras presupuesto)", type(self).__name__)

    @staticmethod
    def _terminate_pid(pid: int) -> None:
        if pid <= 0:
            return
        subprocess.run(["taskkill", "/F", "/T", "/PID", str(pid)], capture_output=True, check=False)

    @staticmethod
    def _hwnd_pid(hwnd: int) -> int:
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        return int(pid.value)

    @classmethod
    def _process_exited(cls, pid: int, timeout: float) -> bool:
        """Espera nativa (WaitForSingleObject) a que el proceso termine."""
        if pid <= 0:
            return True
        handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if not handle:
            return True
        try:
            return kernel32.WaitForSingleObject(handle, int(timeout * 1000)) == WAIT_OBJECT_0
        finally:
            kernel32.CloseHandle(handle)

    # ------------------------------------------------------------------ #
    #  Arranque
    # ------------------------------------------------------------------ #

    def _ensure_running(self) -> None:
        if self._window and self._window.Exists(maxSearchSeconds=1):
            return
        args = [str(self.excel_path)]
        if self.initial_workbook is not None:
            args.append(str(self.initial_workbook))
        logger.info("[%s] Lanzando Excel con args: %s", type(self).__name__, args)
        subprocess.Popen(args)
        self._window = self._find_main_window()
        self._settle_modals(_MODAL_TIMEOUT)

    def _find_main_window(self) -> auto.WindowControl:
        control = self._main_control()
        if not control.Exists(maxSearchSeconds=_MAIN_TIMEOUT):
            raise RuntimeError("No se localizó la ventana principal de Excel.")
        return control

    def _main_control(self) -> auto.WindowControl:
        return auto.WindowControl(searchFromControl=auto.GetRootControl(), ClassName=EXCEL_MAIN_CLASS, searchDepth=1)

    # ------------------------------------------------------------------ #
    #  Modal de activación
    # ------------------------------------------------------------------ #

    def _settle_modals(self, timeout: float) -> None:
        modal = self._modal_control()
        if modal.Exists(maxSearchSeconds=timeout):
            self._dismiss_modal(modal)

    def _modal_control(self) -> auto.WindowControl:
        return auto.WindowControl(searchFromControl=auto.GetRootControl(), ClassName="NUIDialog", searchDepth=_MAX_DIALOG_DEPTH)

    def _dismiss_modal_if_present(self) -> None:
        modal = self._modal_control()
        if modal.Exists(maxSearchSeconds=2):
            self._dismiss_modal(modal)

    @staticmethod
    def _dismiss_modal(modal: auto.WindowControl) -> None:
        logger.info("[ExcelManager] Modal de activación presente: %r", modal.Name)
        button = ExcelManager._find_close_button(modal)
        if button is not None:
            button.Click()
            logger.info("[ExcelManager] Modal cerrado con su botón %r", button.Name)
            return
        try:
            modal.GetWindowPattern().Close()
            logger.info("[ExcelManager] Modal cerrado con WindowPattern.Close")
        except Exception as exc:  # noqa: BLE001
            logger.warning("[ExcelManager] No se pudo cerrar el modal: %s", exc)

    @staticmethod
    def _find_close_button(modal: auto.Control) -> auto.ButtonControl | None:
        stack = [modal]
        while stack:
            node = stack.pop()
            for child in ExcelManager._children(node):
                try:
                    is_close = (
                        child.ControlTypeName == "ButtonControl"
                        and (child.Name or "").strip().lower() in ("cerrar", "close")
                    )
                except Exception:  # noqa: BLE001
                    is_close = False
                if is_close:
                    return child
                stack.append(child)
        return None

    # ------------------------------------------------------------------ #
    #  Despliegue del diálogo (estrategia en cascada)
    # ------------------------------------------------------------------ #

    def _deploy_dialog(self, shortcut: str, shortcut_dialog: str, ribbon_item: str) -> None:
        for attempt in range(1, 4):
            self._dismiss_modal_if_present()
            self._focus_main_window()
            auto.SendKeys(shortcut, waitTime=0.6)
            logger.info("[%s] Atajo %s enviado (intento %d)", type(self).__name__, shortcut, attempt)
            if self._wait_for_dialog(shortcut_dialog):
                return
        logger.info("[%s] Atajo no efectivo; probando cinta de Excel para %r", type(self).__name__, ribbon_item)
        self._open_dialog_via_ribbon(ribbon_item)
        if not self._wait_for_dialog(shortcut_dialog):
            raise RuntimeError(
                f"El diálogo {shortcut_dialog!r} no apareció. "
                "Verifique que Microsoft Office está activado y que Excel arranca correctamente."
            )

    def _open_dialog_via_ribbon(self, ribbon_item: str) -> None:
        """Despliega el diálogo desde la cinta: Pestaña Archivo -> item -> Examinar."""
        if self._window is None:
            self._window = self._find_main_window()
        file_tab = self._find_in(self._window, lambda c: c.AutomationId == "FileTabButton")
        if file_tab is None:
            logger.warning("[%s] No se encontró la cinta (Pestaña Archivo) en Excel.", type(self).__name__)
            return
        self._invoke(file_tab, "Pestaña Archivo")
        nav_item = self._find_in(self._window, lambda c: (c.Name or "").strip() == ribbon_item and c.ControlTypeName == "ListItemControl")
        if nav_item is None:
            logger.warning("[%s] No se encontró el elemento %r en la pestaña Archivo.", type(self).__name__, ribbon_item)
            return
        self._invoke(nav_item, f"elemento {ribbon_item!r}")
        browse = self._find_in(self._window, lambda c: (c.Name or "").strip() == "Examinar" and c.ControlTypeName == "ButtonControl")
        if browse is None:
            logger.warning("[%s] No se encontró el botón Examinar.", type(self).__name__)
            return
        self._invoke(browse, "Examinar")

    # ------------------------------------------------------------------ #
    #  Utilidades de árbol UIA
    # ------------------------------------------------------------------ #

    @staticmethod
    def _walk(node: auto.Control, max_nodes: int = 20000):
        queue = [node]
        count = 0
        while queue and count < max_nodes:
            current = queue.pop(0)
            count += 1
            yield current
            try:
                queue.extend(current.GetChildren())
            except Exception:  # noqa: BLE001
                pass

    @classmethod
    def _find_in(cls, node: auto.Control, predicate) -> auto.Control | None:
        for control in cls._walk(node):
            try:
                if predicate(control):
                    return control
            except Exception:  # noqa: BLE001
                pass
        return None

    @staticmethod
    def _children(node: auto.Control) -> list[auto.Control]:
        try:
            return list(node.GetChildren())
        except Exception:  # noqa: BLE001
            return []

    @staticmethod
    def _invoke(control: auto.Control, label: str) -> None:
        try:
            control.GetInvokePattern().Invoke()
            logger.info("[ExcelManager] Invoke OK: %s", label)
            return
        except Exception:  # noqa: BLE001
            pass
        try:
            control.Click()
            logger.info("[ExcelManager] Click OK (fallback): %s", label)
        except Exception as exc:  # noqa: BLE001
            logger.warning("[ExcelManager] No se pudo accionar %s: %s", label, exc)

    # ------------------------------------------------------------------ #
    #  Detección nativa del diálogo (#32770) en cualquier profundidad
    # ------------------------------------------------------------------ #

    def _wait_for_dialog(self, dialog_name: str) -> bool:
        dialog = auto.WindowControl(
            searchFromControl=auto.GetRootControl(),
            ClassName=DIALOG_CLASS,
            Name=dialog_name,
            searchDepth=_MAX_DIALOG_DEPTH,
        )
        if dialog.Exists(maxSearchSeconds=_DIALOG_TIMEOUT):
            try:
                dialog.SetActive(waitTime=0.5)
            except Exception:  # noqa: BLE001
                pass
            logger.info("[%s] Diálogo %r detectado.", type(self).__name__, dialog_name)
            return True
        return False

    # ------------------------------------------------------------------ #
    #  Foco y cierre
    # ------------------------------------------------------------------ #

    def _focus_main_window(self) -> None:
        if self._window is None:
            self._window = self._find_main_window()
        hwnd = self._window.NativeWindowHandle
        self._force_foreground(hwnd)
        self._window.SetActive(waitTime=0.5)
        try:
            self._window.SetFocus(waitTime=0.5)
        except Exception:  # noqa: BLE001
            pass

    @staticmethod
    def _force_foreground(hwnd: int) -> None:
        if user32.IsIconic(hwnd):
            user32.ShowWindow(hwnd, SW_RESTORE)
        user32.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE)
        user32.SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE)
        user32.BringWindowToTop(hwnd)
        user32.SetForegroundWindow(hwnd)

    def _close_pending(self) -> None:
        """Cierra diálogos (#32770, cualquier título) y el modal de activación
        si aún estuvieran abiertos, tanto en raíz como anidados bajo XLMAIN."""
        for _ in range(4):
            closed = False
            dialog = self._any_dialog_control()
            if dialog.Exists(maxSearchSeconds=1):
                if self._close_window_completely(dialog):
                    closed = True
                else:
                    return
            modal = self._modal_control()
            if modal.Exists(maxSearchSeconds=1):
                self._dismiss_modal(modal)
                closed = True
            if not closed:
                return

    def _any_dialog_control(self) -> auto.WindowControl:
        return auto.WindowControl(
            searchFromControl=auto.GetRootControl(),
            ClassName=DIALOG_CLASS,
            searchDepth=_MAX_DIALOG_DEPTH,
        )

    def _close_window_completely(self, window: auto.WindowControl) -> bool:
        try:
            window.GetWindowPattern().Close()
            return True
        except Exception:  # noqa: BLE001
            pass
        for name in ("Cancelar", "Cancel", "取消"):
            button = window.ButtonControl(Name=name)
            if button.Exists(maxSearchSeconds=1):
                ExcelManager._invoke(button, f"botón {name!r}")
                return True
        return False