# AGENTS.md

## Proyecto: rpa-excel-ui-automation

### Estado actual del repositorio

Proyecto PDM (Python >=3.12) de automatización RPA sobre la interfaz de usuario de Microsoft Excel. Código en `src/rpa_excel_ui_automation/`.

**Estructura del repositorio:**
- `pyproject.toml` - Configuración del proyecto PDM
- `src/rpa_excel_ui_automation/__init__.py` - Paquete con `ExcelManager` y `FileExplorer`
- `src/rpa_excel_ui_automation/excel_manager.py` - Gestor de Excel, atajos (Ctrl+F12 / F12) y cierre seguro
- `src/rpa_excel_ui_automation/file_explorer.py` - Interacción con diálogos nativos (Abrir / Guardar / Reemplazar)
- `src/rpa_excel_ui_automation/app.py` - Punto de entrada `main()` (flujo end-to-end)
- `.data/input/origen.xlsx` - Archivo de prueba base (Excel no activado)
- `.data/output/destino.xlsx` - Archivo destino generado por las ejecuciones
- `pdm.lock` - Bloqueo de dependencias

### Dependencias principales

- `uiautomation>=2.0.29` (automatización UI)
- `openpyxl>=3.1.5` (dependencia de desarrollo, para manipulación de Excel)
- Python >=3.12

### Clases principales (implementación funcional verificada)

**ExcelManager** (`src/rpa_excel_ui_automation/excel_manager.py`):
- `ExcelManager(excel_path=None, initial_workbook=None)`: lanza Excel. `initial_workbook` (ruta del libro) abre rejilla + cinta: condición indispensable en este Office no activado (arrancar sin libro muestra una pantalla fullpage sin cinta y sin atajos útiles).
- `open_file()` / `save_as()`: despliegan "Abrir" / "Guardar como" con **estrategia en cascada**: primero atajo global (`Ctrl+F12` / `F12`); si el diálogo no aparece, vía cinta UIA (`Pestaña Archivo` -> `Abrir/Guardar como` -> `Examinar`) usando `GetInvokePattern().Invoke()` (nada de clics por coordenadas: hay ventanas a pantalla completa encima de Excel).
- El diálogo `#32770` se detecta **en profundidad** (`searchDepth=6`) desde el escritorio: en este entorno aparece **anidado bajo `XLMAIN`**, no como ventana raíz; las búsquedas con `searchDepth=1` fallan.
- `_settle_modals()` / `_dismiss_modal()`: detecta y cierra el modal "Asistente para la activación de Microsoft Office" (`NUIDialog`, aparición no determinista).
- `close()`: cierra primero cualquier `#32770` pendiente y el modal (regla del usuario: nunca cerrar Excel con ventanas abiertas), intenta `WindowPattern.Close()` y verifica la salida real del PID con `WaitForSingleObject`; si no termina dentro del presupuesto nativo, finaliza el PID propio con `taskkill` (Excel sin activar no termina solo tras aceptar una guardar/confirmación).

**FileExplorer** (`src/rpa_excel_ui_automation/file_explorer.py`):
- `open_file(path)` / `save_as(path)`: buscan el diálogo `#32770` por nombre (búsqueda profunda) e inyectan la ruta con `ValuePattern.SetValue` en el campo `AutomationId=1148`.
- El botón primario del diálogo es un **`SplitButtonControl`** (`AutomationId=1`), no un `ButtonControl`; se acciona por Invoke con Click como fallback.
- Confirmación de reemplazo: en Excel 2016 no activado la advertencia se muestra **dentro del propio diálogo "Guardar como"** (título "Guardar como") con botón "Sí"/"No". Se detecta sondeando cualquier `#32770` con un botón `Sí/Yes` (fallback) o los títulos clásicos "Confirmar guardar como" (regex). El sondeo es nativo, hasta 30 s y **solo se activa si `Path(path).exists()`**.
- Selectores basados en `AutomationId` y `Name` (nativos de uiautomation).

### Flujo de trabajo (según README)

1. `ExcelManager(initial_workbook=origen).open_file()` -> diálogo "Abrir" -> `FileExplorer().open_file(origen)`
2. `ExcelManager.save_as()` -> diálogo "Guardar como" -> `FileExplorer().save_as(destino)` (confirma "Sí" si el destino ya existe)
3. `ExcelManager.close()` -> cierra diálogos/modales, cierra Excel y verifica la salida del proceso

**Verificado el 2026-08-27 con `pdm run start`** en esta máquina:
- **Destino inexistente**: guarda directo, sin espera de confirmación. FINALIZADO.
- **Destino existente**: detecta la advertencia, pulsa "Sí" (log "(reemplazado)"), archivo reemplazado; contenido idéntico al origen (comparado con openpyxl).
- Tras `close()` no quedan procesos `EXCEL.EXE` residuales.

### Limitaciones del entorno

- **Office no activado**: Microsoft Office 2016 requiere activación corporativa. Aparece un modal no determinista ("Asistente para la activación") y los atajos de teclado solo a veces despliegan los diálogos (latencias de 10-25 s en las pruebas). La cinta UIA es la vía de respaldo.
- **Foco**: `SetActive()` / `SetForegroundWindow` no siempre retienen el foco; los atajos se envían a la ventana enfocada del sistema.
- **Diálogos anidados**: los `#32770` se registran bajo `XLMAIN`; toda búsqueda debe usar profundidad >= 2 (el código usa 6).
- Los atajos `Ctrl+F12`/`F12` no abren diálogo si Excel arranca sin libro (estado fullpage); por eso `app.py` pasa `initial_workbook=origen`.

### ¿Qué falta por hacer?

- Ejecutar en una máquina con Excel activado para validar el flujo en condiciones normales.
- Suite de pruebas completa según los casos del README.
- `git commit` + `push origin main`: pendientes de validación explícita del usuario (identidad git ya configurada en el repo: `antho <antho@example.com>`; `origin` = `solidjaeger/rpa-excel-ui-automation`).