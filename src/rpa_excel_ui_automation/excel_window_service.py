"""Servicio de gestión de la ventana principal de Excel y sus modales.

Responsabilidad única: enfocar la ventana XLMAIN, detectar y cerrar el modal de
activación de Office, y cerrar diálogos nativos (#32770) pendientes antes de
intentar cerrar Excel.

No sabe nada del proceso (PID) ni de los diálogos de archivo. Esa lógica
pertenece a ``ExcelProcessService`` y ``ExcelDialogOperation`` respectivamente.
"""

from __future__ import annotations

import ctypes

import uiautomation as auto

from rpa_excel_ui_automation.logger import logger

user32 = ctypes.windll.user32

HWND_TOPMOST = -1
HWND_NOTOPMOST = -2
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
SW_RESTORE = 9

DIALOG_CLASS = "#32770"
_MAX_DIALOG_DEPTH = 6  # profundidad de búsqueda de #32770 desde el escritorio
_MODAL_TIMEOUT = 3     # segundos sondeando el modal de activación
_ITEM_TIMEOUT = 1

_CLOSE_BUTTON_NAMES = frozenset({"cerrar", "close"})


class ExcelWindowService:
    """Gestiona el foco, los modales y el cierre de ventanas de Excel."""

    # ── Foco ─────────────────────────────────────────────────────────────────

    def focus(self, window: auto.WindowControl) -> None:
        """Trae la ventana al primer plano y le asigna el foco."""
        hwnd = window.NativeWindowHandle
        self._force_foreground(hwnd)
        window.SetActive(waitTime=0.5)
        try:
            window.SetFocus(waitTime=0.5)
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

    # ── Modal de activación ───────────────────────────────────────────────────

    def settle_modals(self, timeout: float = _MODAL_TIMEOUT) -> None:
        """Espera y descarta el modal de activación de Office si aparece."""
        modal = self._modal_control()
        if modal.Exists(maxSearchSeconds=timeout):
            self.dismiss_modal(modal)

    def dismiss_modal_if_present(self) -> None:
        modal = self._modal_control()
        if modal.Exists(maxSearchSeconds=2):
            self.dismiss_modal(modal)

    def dismiss_modal(self, modal: auto.WindowControl) -> None:
        logger.info("Modal de activación presente: {!r}", modal.Name)
        button = self._find_close_button(modal)
        if button is not None:
            button.Click()
            logger.info("Modal cerrado con botón {!r}", button.Name)
            return
        try:
            modal.GetWindowPattern().Close()
            logger.info("Modal cerrado con WindowPattern.Close")
        except Exception as exc:  # noqa: BLE001
            logger.warning("No se pudo cerrar el modal: {}", exc)

    @staticmethod
    def _modal_control() -> auto.WindowControl:
        return auto.WindowControl(
            searchFromControl=auto.GetRootControl(),
            ClassName="NUIDialog",
            searchDepth=_MAX_DIALOG_DEPTH,
        )

    @staticmethod
    def _find_close_button(modal: auto.Control) -> auto.ButtonControl | None:
        stack = [modal]
        while stack:
            node = stack.pop()
            try:
                children = list(node.GetChildren())
            except Exception:  # noqa: BLE001
                children = []
            for child in children:
                try:
                    is_close = (
                        child.ControlTypeName == "ButtonControl"
                        and (child.Name or "").strip().lower() in _CLOSE_BUTTON_NAMES
                    )
                except Exception:  # noqa: BLE001
                    is_close = False
                if is_close:
                    return child  # type: ignore[return-value]
                stack.append(child)
        return None

    # ── Cierre de diálogos pendientes ─────────────────────────────────────────

    def close_pending(self) -> None:
        """Cierra diálogos (#32770) y modales pendientes en hasta 4 iteraciones."""
        for _ in range(4):
            closed = False
            dialog = self._any_dialog_control()
            if dialog.Exists(maxSearchSeconds=_ITEM_TIMEOUT):
                if self._close_window_completely(dialog):
                    closed = True
                else:
                    return
            modal = self._modal_control()
            if modal.Exists(maxSearchSeconds=_ITEM_TIMEOUT):
                self.dismiss_modal(modal)
                closed = True
            if not closed:
                return

    @staticmethod
    def _any_dialog_control() -> auto.WindowControl:
        return auto.WindowControl(
            searchFromControl=auto.GetRootControl(),
            ClassName=DIALOG_CLASS,
            searchDepth=_MAX_DIALOG_DEPTH,
        )

    @staticmethod
    def _close_window_completely(window: auto.WindowControl) -> bool:
        try:
            window.GetWindowPattern().Close()
            return True
        except Exception:  # noqa: BLE001
            pass
        for name in ("Cancelar", "Cancel", "取消"):
            button = window.ButtonControl(Name=name)
            if button.Exists(maxSearchSeconds=_ITEM_TIMEOUT):
                try:
                    button.GetInvokePattern().Invoke()
                except Exception:  # noqa: BLE001
                    button.Click()
                return True
        return False
