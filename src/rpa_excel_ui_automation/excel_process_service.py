"""Servicio de gestión del proceso de Excel.

Responsabilidad única: arrancar el proceso de Excel, esperar a que la ventana
principal esté disponible y terminar el proceso de forma limpia cuando sea
necesario.

No interactúa directamente con la UI más allá de localizar la ventana principal
(XLMAIN). Toda la lógica de foco, diálogos y modales vive en otros servicios.
"""

from __future__ import annotations

import ctypes
import subprocess
from ctypes import wintypes
from pathlib import Path

import uiautomation as auto

from rpa_excel_ui_automation.logger import logger

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

EXCEL_MAIN_CLASS = "XLMAIN"
DEFAULT_EXCEL_PATH = Path(r"C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE")

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
WAIT_OBJECT_0 = 0
_MAIN_TIMEOUT = 40  # segundos esperando la ventana XLMAIN


class ExcelProcessService:
    """Arranca Excel, localiza su ventana principal y, si es necesario, termina el proceso."""

    def __init__(
        self,
        excel_path: str | Path | None = None,
        initial_workbook: str | Path | None = None,
    ) -> None:
        self.excel_path = Path(excel_path) if excel_path else DEFAULT_EXCEL_PATH
        self.initial_workbook = Path(initial_workbook) if initial_workbook else None

    # ── API pública ──────────────────────────────────────────────────────────

    def launch(self) -> auto.WindowControl:
        """Lanza Excel y retorna el control de la ventana principal (XLMAIN)."""
        args = [str(self.excel_path)]
        if self.initial_workbook is not None:
            args.append(str(self.initial_workbook))
        logger.info("Lanzando Excel con args: {}", args)
        subprocess.Popen(args)
        window = self._find_main_window()
        logger.info("Ventana principal de Excel localizada.")
        return window

    def terminate(self, pid: int) -> None:
        """Termina el proceso de Excel por PID como último recurso."""
        if pid <= 0:
            return
        logger.warning("Terminando proceso Excel (PID={}) por taskkill.", pid)
        subprocess.run(
            ["taskkill", "/F", "/T", "/PID", str(pid)],
            capture_output=True,
            check=False,
        )

    # ── Utilidades de proceso ────────────────────────────────────────────────

    @staticmethod
    def get_pid(hwnd: int) -> int:
        """Devuelve el PID del proceso dueño de una ventana dado su HWND."""
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        return int(pid.value)

    @classmethod
    def wait_for_exit(cls, pid: int, timeout: float) -> bool:
        """Espera (WaitForSingleObject) a que el proceso termine; retorna True si salió."""
        if pid <= 0:
            return True
        handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if not handle:
            return True
        try:
            return kernel32.WaitForSingleObject(handle, int(timeout * 1000)) == WAIT_OBJECT_0
        finally:
            kernel32.CloseHandle(handle)

    # ── Internos ─────────────────────────────────────────────────────────────

    def _find_main_window(self) -> auto.WindowControl:
        control = auto.WindowControl(
            searchFromControl=auto.GetRootControl(),
            ClassName=EXCEL_MAIN_CLASS,
            searchDepth=1,
        )
        if not control.Exists(maxSearchSeconds=_MAIN_TIMEOUT):
            raise RuntimeError("No se localizó la ventana principal de Excel (XLMAIN).")
        return control
