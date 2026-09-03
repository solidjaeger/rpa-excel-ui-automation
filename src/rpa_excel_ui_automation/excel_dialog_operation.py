"""Operación de despliegue de diálogos de Excel (Abrir / Guardar como).

Responsabilidad única: orquestar la secuencia de acciones UIA para conseguir
que aparezca el diálogo nativo de archivo (#32770), ya sea mediante atajos de
teclado o mediante la cinta de Excel como fallback.

Depende de ``ExcelWindowService`` para el foco previo al envío de teclas, pero
no gestiona el ciclo de vida del proceso ni la interacción con el contenido del
diálogo de archivo.
"""

from __future__ import annotations

import uiautomation as auto

from rpa_excel_ui_automation.excel_window_service import ExcelWindowService
from rpa_excel_ui_automation.logger import logger

DIALOG_CLASS = "#32770"
_DIALOG_TIMEOUT = 8   # segundos esperando el diálogo de archivo
_MAX_DIALOG_DEPTH = 6
_MAX_SHORTCUT_ATTEMPTS = 3

_SHORTCUT_WAIT = 0.6   # segundos después de SendKeys


class ExcelDialogOperation:
    """Despliega los diálogos 'Abrir' y 'Guardar como' de Excel.

    Estrategia en cascada:
    1. Atajo global de teclado (Ctrl+F12 / F12) hasta ``_MAX_SHORTCUT_ATTEMPTS`` veces.
    2. Cinta de Excel: Pestaña Archivo → item → Examinar.
    """

    def __init__(self, window_service: ExcelWindowService) -> None:
        self._win = window_service

    # ── API pública ──────────────────────────────────────────────────────────

    def deploy(
        self,
        window: auto.WindowControl,
        shortcut: str,
        dialog_name: str,
        ribbon_item: str,
    ) -> None:
        """Despliega el diálogo nativo o lanza RuntimeError si no aparece."""
        for attempt in range(1, _MAX_SHORTCUT_ATTEMPTS + 1):
            self._win.dismiss_modal_if_present()
            self._win.focus(window)
            auto.SendKeys(shortcut, waitTime=_SHORTCUT_WAIT)
            logger.info("Atajo {} enviado (intento {})", shortcut, attempt)
            if self._wait_for_dialog(dialog_name):
                return

        logger.info("Atajo no efectivo; probando cinta de Excel para {!r}", ribbon_item)
        self._open_via_ribbon(window, ribbon_item)
        if not self._wait_for_dialog(dialog_name):
            raise RuntimeError(
                f"El diálogo {dialog_name!r} no apareció. "
                "Verifique que Microsoft Office está activado y que Excel arranca correctamente."
            )

    # ── Cinta de Excel ────────────────────────────────────────────────────────

    @staticmethod
    def _open_via_ribbon(window: auto.WindowControl, ribbon_item: str) -> None:
        """Pestaña Archivo → ribbon_item → Examinar."""
        file_tab = ExcelDialogOperation._find_in(
            window, lambda c: c.AutomationId == "FileTabButton"
        )
        if file_tab is None:
            logger.warning("No se encontró la cinta (Pestaña Archivo) en Excel.")
            return
        ExcelDialogOperation._invoke(file_tab, "Pestaña Archivo")

        nav_item = ExcelDialogOperation._find_in(
            window,
            lambda c: (c.Name or "").strip() == ribbon_item
            and c.ControlTypeName == "ListItemControl",
        )
        if nav_item is None:
            logger.warning("No se encontró el elemento {!r} en la pestaña Archivo.", ribbon_item)
            return
        ExcelDialogOperation._invoke(nav_item, f"elemento {ribbon_item!r}")

        browse = ExcelDialogOperation._find_in(
            window,
            lambda c: (c.Name or "").strip() == "Examinar"
            and c.ControlTypeName == "ButtonControl",
        )
        if browse is None:
            logger.warning("No se encontró el botón Examinar.")
            return
        ExcelDialogOperation._invoke(browse, "Examinar")

    # ── Detección del diálogo ─────────────────────────────────────────────────

    @staticmethod
    def _wait_for_dialog(dialog_name: str) -> bool:
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
            logger.info("Diálogo {!r} detectado.", dialog_name)
            return True
        return False

    # ── Utilidades de árbol UIA ───────────────────────────────────────────────

    @staticmethod
    def _walk(node: auto.Control, max_nodes: int = 20_000):
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
    def _invoke(control: auto.Control, label: str) -> None:
        try:
            control.GetInvokePattern().Invoke()
            logger.info("Invoke OK: {}", label)
            return
        except Exception:  # noqa: BLE001
            pass
        try:
            control.Click()
            logger.info("Click OK (fallback): {}", label)
        except Exception as exc:  # noqa: BLE001
            logger.warning("No se pudo accionar {}: {}", label, exc)
