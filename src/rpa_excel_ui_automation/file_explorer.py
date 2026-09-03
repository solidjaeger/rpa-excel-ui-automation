"""Fachada de alto nivel para interactuar con diálogos nativos del explorador de archivos.

Responsabilidad única: ofrecer una API limpia de dos métodos (``open_file`` /
``save_as``) y delegar todos los detalles de localización de controles e
inyección de rutas a los servicios y operaciones especializados:

- ``FileDialogService``   → localización de diálogos y controles UIA.
- ``FileDialogOperation`` → inyección de ruta y confirmación de acciones.
"""

from __future__ import annotations

from pathlib import Path

from rpa_excel_ui_automation.file_dialog_operation import FileDialogOperation
from rpa_excel_ui_automation.file_dialog_service import FileDialogService
from rpa_excel_ui_automation.logger import logger

_DEFAULT_TIMEOUT = 10


class FileExplorer:
    """Gestiona las ventanas de diálogo de archivo del sistema operativo."""

    def __init__(self, timeout: int = _DEFAULT_TIMEOUT) -> None:
        self._service = FileDialogService(timeout=timeout)
        self._operation = FileDialogOperation(self._service)

    # ── API pública ──────────────────────────────────────────────────────────

    def open_file(self, path: str | Path) -> None:
        """Inyecta ``path`` en el diálogo 'Abrir' y lo confirma."""
        logger.info("FileExplorer: open_file '{}'", path)
        self._operation.open_file(path)

    def save_as(self, path: str | Path) -> None:
        """Inyecta ``path`` en el diálogo 'Guardar como', confirma y gestiona sobreescritura."""
        logger.info("FileExplorer: save_as '{}'", path)
        self._operation.save_as(path)