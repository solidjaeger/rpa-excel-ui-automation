"""Configuración centralizada de Loguru para el paquete rpa_excel_ui_automation.

Uso en cualquier módulo del paquete::

    from rpa_excel_ui_automation.logger import logger

    logger.info("Mensaje de información")
    logger.debug("Mensaje de depuración")

La configuración añade:
- Un handler a stdout con nivel INFO y formato legible para humanos.
- Un handler a un archivo rotativo en ``logs/rpa.log`` con nivel DEBUG.

Para silenciar el logger de ``uiautomation`` (que puede ser muy verboso),
se filtra con ``loguru``'s ``filter``.
"""

from __future__ import annotations

import sys
from pathlib import Path

from loguru import logger

# Elimina el handler por defecto de Loguru para configurar los propios.
logger.remove()

# ── Handler de consola (stdout, nivel INFO) ──────────────────────────────────
logger.add(
    sys.stdout,
    level="INFO",
    colorize=True,
    format=(
        "<green>{time:YYYY-MM-DD HH:mm:ss}</green> "
        "<level>[{level}]</level> "
        "<cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> "
        "- <level>{message}</level>"
    ),
)

# ── Handler de archivo rotativo (nivel DEBUG) ────────────────────────────────
_LOG_DIR = Path(__file__).resolve().parents[3] / "logs"
_LOG_DIR.mkdir(parents=True, exist_ok=True)

logger.add(
    _LOG_DIR / "rpa.log",
    level="DEBUG",
    rotation="5 MB",
    retention="10 days",
    compression="zip",
    encoding="utf-8",
    format=(
        "{time:YYYY-MM-DD HH:mm:ss.SSS} [{level}] "
        "{name}:{function}:{line} - {message}"
    ),
)

__all__ = ["logger"]
