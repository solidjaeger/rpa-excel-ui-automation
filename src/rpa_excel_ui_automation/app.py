"""Caso de prueba end-to-end: apertura de origen y exportación segura ('Guardar como').

Flujo según README:
  1. ExcelManager.open_file() -> diálogo 'Abrir' + FileExplorer.inyectar ruta origen.
  2. ExcelManager.save_as()   -> diálogo 'Guardar como' + FileExplorer.inyectar ruta destino
     y confirmar reemplazo si el archivo ya existe.
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path

from rpa_excel_ui_automation.excel_manager import ExcelManager
from rpa_excel_ui_automation.file_explorer import FileExplorer

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
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        stream=sys.stdout,
    )

    if not SOURCE_XLSX.exists():
        raise FileNotFoundError(f"Archivo de origen no encontrado: {SOURCE_XLSX}")

    excel = ExcelManager(initial_workbook=SOURCE_XLSX)
    explorer = FileExplorer()

    try:
        logger = logging.getLogger("rpa")
        logger.info("=== Caso 01: apertura de %s ===", SOURCE_XLSX)
        excel.open_file()
        explorer.open_file(SOURCE_XLSX)

        logger.info("=== Caso 02: exportación a %s ===", DEST_XLSX)
        excel.save_as()
        explorer.save_as(DEST_XLSX)
    finally:
        excel.close()

    if DEST_XLSX.exists():
        logging.getLogger("rpa").info("FINALIZADO: se generó %s", DEST_XLSX)
        return 0
    logging.getLogger("rpa").error("FINALIZADO CON ERROR: no se generó %s", DEST_XLSX)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
