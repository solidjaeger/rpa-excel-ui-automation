"""Caso de prueba end-to-end: apertura de origen y exportación segura ('Guardar como').

Flujo según README:
  1. ExcelManager.open_file() -> diálogo 'Abrir' + FileExplorer.inyectar ruta origen.
  2. ExcelManager.save_as()   -> diálogo 'Guardar como' + FileExplorer.inyectar ruta destino
     y confirmar reemplazo si el archivo ya existe.
"""

from __future__ import annotations

import sys
from pathlib import Path

from rpa_excel_ui_automation.excel_manager import ExcelManager
from rpa_excel_ui_automation.file_explorer import FileExplorer

# Importar logger inicializa Loguru con los handlers de consola y archivo.
from rpa_excel_ui_automation.logger import logger

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DATA_DIR = PROJECT_ROOT / ".data"
SOURCE_XLSX = DATA_DIR / "input" / "origen.xlsx"
DEST_XLSX = DATA_DIR / "output" / "destino.xlsx"


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):  # noqa: PERF203
            pass

    if not SOURCE_XLSX.exists():
        raise FileNotFoundError(f"Archivo de origen no encontrado: {SOURCE_XLSX}")

    excel = ExcelManager(initial_workbook=SOURCE_XLSX)
    explorer = FileExplorer()

    try:
        logger.info("=== Caso 01: apertura de {} ===", SOURCE_XLSX)
        excel.open_file()
        explorer.open_file(SOURCE_XLSX)

        logger.info("=== Caso 02: exportación a {} ===", DEST_XLSX)
        excel.save_as()
        explorer.save_as(DEST_XLSX)
    finally:
        excel.close()

    if DEST_XLSX.exists():
        logger.info("FINALIZADO: se generó {}", DEST_XLSX)
        return 0

    logger.error("FINALIZADO CON ERROR: no se generó {}", DEST_XLSX)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
