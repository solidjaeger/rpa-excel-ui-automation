"""Fachada de alto nivel para controlar Excel.

Responsabilidad única: orquestar los servicios y operaciones necesarios para
abrir un libro de Excel, desplegar diálogos de archivo y cerrar la aplicación
de forma segura.

Los detalles de cada tarea están delegados a:
- ``ExcelProcessService``  → ciclo de vida del proceso (lanzar, terminar, PID).
- ``ExcelWindowService``   → foco, modales y diálogos pendientes.
- ``ExcelDialogOperation`` → despliegue de los diálogos Abrir / Guardar como.
"""

from __future__ import annotations

import time
from pathlib import Path

import uiautomation as auto

from rpa_excel_ui_automation.excel_dialog_operation import ExcelDialogOperation
from rpa_excel_ui_automation.excel_process_service import ExcelProcessService
from rpa_excel_ui_automation.excel_window_service import ExcelWindowService
from rpa_excel_ui_automation.logger import logger

OPEN_SHORTCUT = "{Ctrl}{F12}"   # despliega el diálogo "Abrir"
SAVE_AS_SHORTCUT = "{F12}"      # despliega el diálogo "Guardar como"

_WAIT_AFTER_CLOSE = 6    # segundos esperando el cierre tras WM_CLOSE
_CLOSE_BUDGET = 15       # presupuesto total para el cierre elegante


class ExcelManager:
    """Fachada que inicializa Excel y despliega los diálogos 'Abrir' / 'Guardar como'."""

    def __init__(
        self,
        excel_path: str | Path | None = None,
        initial_workbook: str | Path | None = None,
    ) -> None:
        self._process_svc = ExcelProcessService(excel_path, initial_workbook)
        self._window_svc = ExcelWindowService()
        self._dialog_op = ExcelDialogOperation(self._window_svc)
        self._window: auto.WindowControl | None = None

    # ── API pública ──────────────────────────────────────────────────────────

    def open_file(self) -> None:
        logger.info("ExcelManager: iniciando apertura")
        self._ensure_running()
        self._dialog_op.deploy(
            self._window,  # type: ignore[arg-type]
            OPEN_SHORTCUT,
            "Abrir",
            "Abrir",
        )
        logger.info("ExcelManager: apertura completada")

    def save_as(self) -> None:
        logger.info("ExcelManager: iniciando guardado")
        self._ensure_running()
        self._dialog_op.deploy(
            self._window,  # type: ignore[arg-type]
            SAVE_AS_SHORTCUT,
            "Guardar como",
            "Guardar como",
        )
        logger.info("ExcelManager: guardado completado")

    def close(self) -> None:
        """Cierra Excel de forma segura; como último recurso finaliza el PID."""
        if self._window is None:
            return
        logger.info("ExcelManager: cerrando Excel")
        pid = self._process_svc.get_pid(self._window.NativeWindowHandle)
        deadline = time.monotonic() + _CLOSE_BUDGET

        while time.monotonic() < deadline:
            self._window_svc.close_pending()
            try:
                self._window.GetWindowPattern().Close()
            except Exception as exc:  # noqa: BLE001
                logger.debug("WindowPattern.Close: {}", exc)
            if not self._window.Exists(maxSearchSeconds=_WAIT_AFTER_CLOSE):
                if self._process_svc.wait_for_exit(pid, _WAIT_AFTER_CLOSE):
                    logger.info("ExcelManager: Excel cerrado")
                else:
                    self._process_svc.terminate(pid)
                    logger.info("ExcelManager: Excel cerrado (proceso finalizado)")
                return

        self._window_svc.close_pending()
        try:
            self._window.GetWindowPattern().Close()
        except Exception:  # noqa: BLE001
            pass
        self._process_svc.terminate(pid)
        logger.info("ExcelManager: Excel cerrado (proceso finalizado tras presupuesto)")

    # ── Internos ─────────────────────────────────────────────────────────────

    def _ensure_running(self) -> None:
        if self._window and self._window.Exists(maxSearchSeconds=1):
            return
        self._window = self._process_svc.launch()
        self._window_svc.settle_modals()