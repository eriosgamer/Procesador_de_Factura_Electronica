@echo off

REM Exclude patterns from .exclude
setlocal enabledelayedexpansion
set EXCLUDE=
for /F "usebackq delims=" %%i in (".exclude") do (
    set line=%%i
    if not "!line!"=="" (
        set EXCLUDE=!EXCLUDE! --exclude=!line!
    )
)

REM Name of the output executable
set OUTPUT=FacturaElectronica

REM Find main Python file (adjust if needed)
set MAIN=ui.py

REM Instalar dependencias si existe requirements.txt
set HIDDEN_IMPORTS=
if exist requirements.txt (
    pip install -r requirements.txt
    for /F "usebackq delims=" %%i in (requirements.txt) do (
        set line=%%i
        REM Ignorar líneas vacías y comentarios
        if not "!line!"=="" if not "!line:~0,1!"=="#" (
            REM Extraer solo el nombre del paquete (sin versión)
            for /F "tokens=1 delims==><" %%a in ("!line!") do (
                set HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=%%a
            )
        )
    )
)

REM Compilar con PyInstaller del entorno activo incluyendo hidden-imports y submódulos
set HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=PySide6 --hidden-import=PySide6.QtCore --hidden-import=PySide6.QtGui --hidden-import=PySide6.QtWidgets --hidden-import=PySide6.QtNetwork --hidden-import=PySide6.QtPrintSupport --hidden-import=rich.progress --hidden-import=rich.console --hidden-import=openpyxl.utils --hidden-import=openpyxl.cell._writer

pyinstaller --onefile --name %OUTPUT% %EXCLUDE% !HIDDEN_IMPORTS! %MAIN%
