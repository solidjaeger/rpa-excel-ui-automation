"""Interacción exclusiva con diálogos nativos del explorador de archivos.

Responsabilidad única: inyectar rutas en el campo de nombre de archivo y
confirmar las acciones (Abrir / Guardar / Reemplazar) de forma directa,
sin navegación con Tab ni clics por coordenadas.

Sin pausas estáticas: toda espera usa los mecanismos nativos de `uiautomation`
(`.Exists(maxSearchSeconds=...)`), que sondean internamente.
"""

from __future__ import annotations

import logging
import re
import time
from pathlib import Path

import uiautomation as auto

logger = logging.getLogger(__name__)

_DIALOG_CLASS = "#32770"

_FILENAME_FIELD_IDS = ("1148", "1001")
_FILENAME_FIELD_NAMES = ("Nombre de archivo", "File name", "Name", "文件名")

_OPEN_DIALOG_NAMES = ("Abrir", "Open", "打开")
_SAVE_DIALOG_NAMES = ("Guardar como", "Save As", "另存为")

_OPEN_BUTTON_NAMES = ("Abrir", "Open", "打开")
_SAVE_BUTTON_NAMES = ("Guardar", "Save", "保存")
_YES_BUTTON_NAMES = ("Sí", "Yes", "是")
_OVERWRITE_TITLE_NAMES = ("Confirmar guardar como", "Confirm Save As", "确认另存为")

_PRIMARY_BUTTON_ID = "1"
_DROPDOWN_BUTTON_ID = "DropDown"
_DIALOG_TIMEOUT = 10
_OVERWRITE_TIMEOUT = 30
_ITEM_TIMEOUT = 1
_MAX_DIALOG_DEPTH = 6  # profundidad de búsqueda de #32770 desde el escritorio

_OVERWRITE_TITLE_REGEX = re.compile(r"^(" + "|".join(map(re.escape, _OVERWRITE_TITLE_NAMES)) + r")$")


class FileExplorer:
    """Gestiona las ventanas de diálogo de archivo del sistema operativo."""

    def __init__(self, timeout: int = _DIALOG_TIMEOUT) -> None:
        self.timeout = timeout

    def open_file(self, path: str | Path) -> None:
        logger.info("[%s] Iniciando apertura: inyectando ruta '%s'.", type(self).__name__, path)
        dialog = self._find_dialog(_OPEN_DIALOG_NAMES)
        self._set_filename(dialog, path)
        self._invoke_button(dialog, _OPEN_BUTTON_NAMES)
        logger.info("[%s] Apertura confirmada para '%s'.", type(self).__name__, path)

    def save_as(self, path: str | Path) -> None:
        logger.info("[%s] Iniciando guardado: inyectando ruta '%s'.", type(self).__name__, path)
        dialog = self._find_dialog(_SAVE_DIALOG_NAMES)
        self._set_filename(dialog, path)
        self._invoke_button(dialog, _SAVE_BUTTON_NAMES)
        # Solo tiene sentido esperar la confirmación si el destino ya existe.
        was_replaced = Path(path).exists() and self._confirm_overwrite()
        logger.info(
            "[%s] Guardado completado para '%s'%s.",
            type(self).__name__,
            path,
            " (reemplazado)" if was_replaced else "",
        )

    def _find_dialog(self, names: tuple[str, ...]) -> auto.WindowControl:
        for name in names:
            dialog = auto.WindowControl(
                searchFromControl=auto.GetRootControl(),
                ClassName=_DIALOG_CLASS,
                Name=name,
                searchDepth=_MAX_DIALOG_DEPTH,
            )
            if dialog.Exists(maxSearchSeconds=self.timeout):
                try:
                    dialog.SetActive(waitTime=0.5)
                except Exception:  # noqa: BLE001
                    pass
                return dialog
        raise RuntimeError(f"No se detectó el diálogo de archivo ({', '.join(names)}).")

    @staticmethod
    def _set_filename(dialog: auto.WindowControl, path: str | Path) -> None:
        edit = FileExplorer._find_filename_edit(dialog)
        edit.SetFocus()
        edit.GetValuePattern().SetValue(str(path))

    @staticmethod
    def _find_filename_edit(dialog: auto.Control) -> auto.EditControl:
        for control_id in _FILENAME_FIELD_IDS:
            edit = dialog.EditControl(AutomationId=control_id)
            if edit.Exists(maxSearchSeconds=_ITEM_TIMEOUT):
                return edit
        for name in _FILENAME_FIELD_NAMES:
            edit = dialog.EditControl(NameContains=name)
            if edit.Exists(maxSearchSeconds=_ITEM_TIMEOUT):
                return edit
        raise RuntimeError("No se localizó el campo de nombre de archivo en el diálogo.")

    @staticmethod
    def _invoke_button(dialog: auto.Control, names: tuple[str, ...]) -> None:
        primary = dialog.SplitButtonControl(AutomationId=_PRIMARY_BUTTON_ID)
        if not primary.Exists(maxSearchSeconds=_ITEM_TIMEOUT):
            primary = dialog.ButtonControl(AutomationId=_PRIMARY_BUTTON_ID)
        if primary.Exists(maxSearchSeconds=_ITEM_TIMEOUT):
            FileExplorer._activate(primary, f"botón primario ({', '.join(names)})")
            return
        for ctor in (auto.SplitButtonControl, auto.ButtonControl):
            for name in names:
                button = ctor(searchFromControl=dialog, Name=name)
                if button.Exists(maxSearchSeconds=_ITEM_TIMEOUT) and (button.AutomationId or "") != _DROPDOWN_BUTTON_ID:
                    FileExplorer._activate(button, f"botón {name!r}")
                    return
        raise RuntimeError(f"No se localizó el botón esperado: {', '.join(names)}.")

    @staticmethod
    def _activate(control: auto.Control, label: str) -> None:
        try:
            control.GetInvokePattern().Invoke()
            logger.info("[%s] Accionado por Invoke: %s", FileExplorer.__name__, label)
            return
        except Exception:  # noqa: BLE001
            pass
        try:
            control.Click()
            logger.info("[%s] Accionado por Click: %s", FileExplorer.__name__, label)
        except Exception as exc:  # noqa: BLE001
            raise RuntimeError(f"No se pudo accionar {label}: {exc}") from exc

    def _confirm_overwrite(self) -> bool:
        confirm = self._find_overwrite_confirm()
        if confirm is None:
            return False
        title = confirm.Name
        logger.info("[%s] Advertencia de sobreescritura detectada: '%s'.", type(self).__name__, title)
        try:
            confirm.SetActive(waitTime=0.5)
        except Exception:  # noqa: BLE001
            pass
        for name in _YES_BUTTON_NAMES:
            button = confirm.ButtonControl(Name=name)
            if button.Exists(maxSearchSeconds=_ITEM_TIMEOUT):
                FileExplorer._activate(button, f"botón {name!r}")
                logger.info("[%s] Reemplazo confirmado.", type(self).__name__)
                return True
        # Fallback: barrido profundo dentro del diálogo por si el botón no es hijo directo.
        stack = [confirm]
        while stack:
            node = stack.pop()
            try:
                if node.ControlTypeName == "ButtonControl" and (node.Name or "").strip() in _YES_BUTTON_NAMES:
                    FileExplorer._activate(node, f"botón {node.Name!r}")
                    logger.info("[%s] Reemplazo confirmado.", type(self).__name__)
                    return True
                stack.extend(node.GetChildren())
            except Exception:  # noqa: BLE001
                pass
        raise RuntimeError("Advertencia de sobreescritura presente sin botón 'Sí'.")

    def _find_overwrite_confirm(self) -> auto.WindowControl | None:
        """Sondea hasta `_OVERWRITE_TIMEOUT` s el diálogo de reemplazo (títulos conocidos o
        cualquier #32770 con botón 'Sí'), usando únicamente esperas nativas."""
        exact = auto.WindowControl(
            searchFromControl=auto.GetRootControl(),
            ClassName=_DIALOG_CLASS,
            RegexName=_OVERWRITE_TITLE_REGEX,
            searchDepth=_MAX_DIALOG_DEPTH,
        )
        any_dialog = auto.WindowControl(
            searchFromControl=auto.GetRootControl(),
            ClassName=_DIALOG_CLASS,
            searchDepth=_MAX_DIALOG_DEPTH,
        )
        deadline = time.monotonic() + _OVERWRITE_TIMEOUT
        while time.monotonic() < deadline:
            if exact.Exists(maxSearchSeconds=0.5):
                return exact
            if any_dialog.Exists(maxSearchSeconds=0.5):
                stack = [any_dialog]
                while stack:
                    node = stack.pop()
                    try:
                        if node.ControlTypeName == "ButtonControl" and (node.Name or "").strip() in _YES_BUTTON_NAMES:
                            return any_dialog
                        stack.extend(node.GetChildren())
                    except Exception:  # noqa: BLE001
                        pass
        return None