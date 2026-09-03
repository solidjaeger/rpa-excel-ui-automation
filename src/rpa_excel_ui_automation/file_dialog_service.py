"""Servicio de localización de diálogos y controles UIA en el explorador de archivos.

Responsabilidad única: encontrar el diálogo nativo del sistema operativo (#32770)
correspondiente a "Abrir" o "Guardar como", localizar el campo de nombre de
archivo y los botones de acción dentro de ese diálogo.

No inyecta rutas ni confirma acciones; esa lógica pertenece a
``FileDialogOperation``.
"""

from __future__ import annotations

import uiautomation as auto

from rpa_excel_ui_automation.logger import logger

_DIALOG_CLASS = "#32770"
_MAX_DIALOG_DEPTH = 6

_FILENAME_FIELD_IDS = ("1148", "1001")
_FILENAME_FIELD_NAMES = ("Nombre de archivo", "File name", "Name", "文件名")
_PRIMARY_BUTTON_ID = "1"
_DROPDOWN_BUTTON_ID = "DropDown"
_ITEM_TIMEOUT = 1


class FileDialogService:
    """Localiza el diálogo de archivo del SO y sus controles internos."""

    def __init__(self, timeout: int = 10) -> None:
        self.timeout = timeout

    # ── Diálogo ───────────────────────────────────────────────────────────────

    def find_dialog(self, names: tuple[str, ...]) -> auto.WindowControl:
        """Retorna el diálogo (#32770) cuyo nombre coincida con alguno de ``names``."""
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
                logger.debug("Diálogo encontrado: {!r}", name)
                return dialog
        raise RuntimeError(f"No se detectó el diálogo de archivo ({', '.join(names)}).")

    # ── Campo de nombre de archivo ────────────────────────────────────────────

    @staticmethod
    def find_filename_edit(dialog: auto.Control) -> auto.EditControl:
        """Localiza el campo de nombre de archivo dentro del diálogo."""
        for control_id in _FILENAME_FIELD_IDS:
            edit = dialog.EditControl(AutomationId=control_id)
            if edit.Exists(maxSearchSeconds=_ITEM_TIMEOUT):
                logger.debug("Campo de nombre de archivo encontrado por ID: {}", control_id)
                return edit
        for name in _FILENAME_FIELD_NAMES:
            edit = dialog.EditControl(NameContains=name)
            if edit.Exists(maxSearchSeconds=_ITEM_TIMEOUT):
                logger.debug("Campo de nombre de archivo encontrado por nombre: {}", name)
                return edit
        raise RuntimeError("No se localizó el campo de nombre de archivo en el diálogo.")

    # ── Botones de acción ─────────────────────────────────────────────────────

    @staticmethod
    def find_action_button(
        dialog: auto.Control, names: tuple[str, ...]
    ) -> auto.Control:
        """Localiza el botón principal de acción (Abrir / Guardar / Sí)."""
        primary = dialog.SplitButtonControl(AutomationId=_PRIMARY_BUTTON_ID)
        if not primary.Exists(maxSearchSeconds=_ITEM_TIMEOUT):
            primary = dialog.ButtonControl(AutomationId=_PRIMARY_BUTTON_ID)
        if primary.Exists(maxSearchSeconds=_ITEM_TIMEOUT):
            logger.debug("Botón primario encontrado por ID: {}", _PRIMARY_BUTTON_ID)
            return primary

        for ctor in (auto.SplitButtonControl, auto.ButtonControl):
            for name in names:
                button = ctor(searchFromControl=dialog, Name=name)
                if button.Exists(maxSearchSeconds=_ITEM_TIMEOUT) and (
                    button.AutomationId or ""
                ) != _DROPDOWN_BUTTON_ID:
                    logger.debug("Botón encontrado por nombre: {!r}", name)
                    return button

        raise RuntimeError(f"No se localizó el botón esperado: {', '.join(names)}.")

    # ── Diálogo de sobreescritura ─────────────────────────────────────────────

    @staticmethod
    def find_overwrite_dialog_with_yes(
        yes_names: tuple[str, ...], deadline_mono: float
    ) -> auto.WindowControl | None:
        """Sondea hasta ``deadline_mono`` (monotonic) el diálogo de reemplazo.

        Retorna el control del diálogo si lo encuentra con un botón 'Sí',
        o ``None`` si expira el tiempo.
        """
        import re
        import time

        _OVERWRITE_TITLES = ("Confirmar guardar como", "Confirm Save As", "确认另存为")
        title_regex = re.compile(
            r"^(" + "|".join(map(re.escape, _OVERWRITE_TITLES)) + r")$"
        )

        exact = auto.WindowControl(
            searchFromControl=auto.GetRootControl(),
            ClassName=_DIALOG_CLASS,
            RegexName=title_regex,
            searchDepth=_MAX_DIALOG_DEPTH,
        )
        any_dialog = auto.WindowControl(
            searchFromControl=auto.GetRootControl(),
            ClassName=_DIALOG_CLASS,
            searchDepth=_MAX_DIALOG_DEPTH,
        )

        while time.monotonic() < deadline_mono:
            if exact.Exists(maxSearchSeconds=0.5):
                return exact
            if any_dialog.Exists(maxSearchSeconds=0.5):
                stack = [any_dialog]
                while stack:
                    node = stack.pop()
                    try:
                        if node.ControlTypeName == "ButtonControl" and (
                            node.Name or ""
                        ).strip() in yes_names:
                            return any_dialog
                        stack.extend(node.GetChildren())
                    except Exception:  # noqa: BLE001
                        pass
        return None
