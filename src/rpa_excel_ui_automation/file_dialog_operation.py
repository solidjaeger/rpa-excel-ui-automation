"""Operación de inyección de rutas y confirmación en diálogos de archivo.

Responsabilidad única: dado un diálogo UIA ya localizado (proporcionado por
``FileDialogService``), inyectar la ruta en el campo de nombre de archivo,
accionar el botón primario y confirmar el diálogo de sobreescritura cuando
sea necesario.

No busca diálogos por su cuenta; esa responsabilidad es de ``FileDialogService``.
"""

from __future__ import annotations

import time
from pathlib import Path

import uiautomation as auto

from rpa_excel_ui_automation.file_dialog_service import FileDialogService
from rpa_excel_ui_automation.logger import logger

_OPEN_DIALOG_NAMES = ("Abrir", "Open", "打开")
_SAVE_DIALOG_NAMES = ("Guardar como", "Save As", "另存为")
_OPEN_BUTTON_NAMES = ("Abrir", "Open", "打开")
_SAVE_BUTTON_NAMES = ("Guardar", "Save", "保存")
_YES_BUTTON_NAMES = ("Sí", "Yes", "是")

_OVERWRITE_TIMEOUT = 30
_ITEM_TIMEOUT = 1


class FileDialogOperation:
    """Inyecta rutas y confirma acciones en diálogos de archivo del SO.

    Depende de ``FileDialogService`` para localizar los controles UIA.
    """

    def __init__(self, service: FileDialogService) -> None:
        self._service = service

    # ── API pública ──────────────────────────────────────────────────────────

    def open_file(self, path: str | Path) -> None:
        """Inyecta la ruta en el diálogo 'Abrir' y confirma."""
        logger.info("Iniciando apertura: inyectando ruta '{}'.", path)
        dialog = self._service.find_dialog(_OPEN_DIALOG_NAMES)
        self._set_filename(dialog, path)
        button = self._service.find_action_button(dialog, _OPEN_BUTTON_NAMES)
        self._activate(button, f"botón primario ({', '.join(_OPEN_BUTTON_NAMES)})")
        logger.info("Apertura confirmada para '{}'.", path)

    def save_as(self, path: str | Path) -> None:
        """Inyecta la ruta en el diálogo 'Guardar como' y confirma (incluyendo sobreescritura)."""
        logger.info("Iniciando guardado: inyectando ruta '{}'.", path)
        dialog = self._service.find_dialog(_SAVE_DIALOG_NAMES)
        self._set_filename(dialog, path)
        button = self._service.find_action_button(dialog, _SAVE_BUTTON_NAMES)
        self._activate(button, f"botón primario ({', '.join(_SAVE_BUTTON_NAMES)})")

        # Solo esperar confirmación de sobreescritura si el archivo ya existe.
        was_replaced = Path(path).exists() and self._confirm_overwrite()
        logger.info(
            "Guardado completado para '{}'{} .",
            path,
            " (reemplazado)" if was_replaced else "",
        )

    # ── Internos ─────────────────────────────────────────────────────────────

    def _set_filename(self, dialog: auto.WindowControl, path: str | Path) -> None:
        edit = self._service.find_filename_edit(dialog)
        edit.SetFocus()
        edit.GetValuePattern().SetValue(str(path))

    def _confirm_overwrite(self) -> bool:
        deadline = time.monotonic() + _OVERWRITE_TIMEOUT
        confirm = self._service.find_overwrite_dialog_with_yes(_YES_BUTTON_NAMES, deadline)
        if confirm is None:
            return False
        logger.info("Advertencia de sobreescritura detectada: {!r}", confirm.Name)
        try:
            confirm.SetActive(waitTime=0.5)
        except Exception:  # noqa: BLE001
            pass
        # Intento directo con los nombres de botón conocidos.
        for name in _YES_BUTTON_NAMES:
            button = confirm.ButtonControl(Name=name)
            if button.Exists(maxSearchSeconds=_ITEM_TIMEOUT):
                self._activate(button, f"botón {name!r}")
                logger.info("Reemplazo confirmado.")
                return True
        # Fallback: barrido profundo dentro del diálogo.
        stack = [confirm]
        while stack:
            node = stack.pop()
            try:
                if node.ControlTypeName == "ButtonControl" and (
                    node.Name or ""
                ).strip() in _YES_BUTTON_NAMES:
                    self._activate(node, f"botón {node.Name!r}")
                    logger.info("Reemplazo confirmado (barrido profundo).")
                    return True
                stack.extend(node.GetChildren())
            except Exception:  # noqa: BLE001
                pass
        raise RuntimeError("Advertencia de sobreescritura presente sin botón 'Sí'.")

    @staticmethod
    def _activate(control: auto.Control, label: str) -> None:
        try:
            control.GetInvokePattern().Invoke()
            logger.info("Accionado por Invoke: {}", label)
            return
        except Exception:  # noqa: BLE001
            pass
        try:
            control.Click()
            logger.info("Accionado por Click: {}", label)
        except Exception as exc:  # noqa: BLE001
            raise RuntimeError(f"No se pudo accionar {label}: {exc}") from exc
