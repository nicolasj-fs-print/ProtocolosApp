@echo off
setlocal

REM ============================================================
REM Build .exe one-file con PyInstaller
REM Empaqueta TODO (assets, tools, .env) en un único ProtocolosApp.exe
REM Requiere: venv activable y dependencias ya instaladas.
REM ============================================================

if not exist .venv\Scripts\activate.bat (
    echo [ERROR] No existe .venv. Ejecutar primero:
    echo    python -m venv .venv
    echo    .venv\Scripts\activate
    echo    pip install -r requirements.txt
    exit /b 1
)

call .venv\Scripts\activate.bat

if not exist .env (
    echo [ERROR] No existe .env. Copia .env.example a .env y completa los valores.
    exit /b 1
)

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist ProtocolosApp.spec del /q ProtocolosApp.spec

pyinstaller --noconfirm --onedir --windowed --name ProtocolosApp ^
    --icon "assets\Icono.ico" ^
    --add-data "assets;assets" ^
    --add-data "tools\SumatraPDF.exe;tools" ^
    --add-data "tools\msodbcsql18.msi;tools" ^
    --add-data ".env;." ^
    --collect-all customtkinter ^
    --collect-all tkcalendar ^
    --collect-submodules pyodbc ^
    --collect-submodules app ^
    --hidden-import win32com.client ^
    --hidden-import PIL._tkinter_finder ^
    --hidden-import msal ^
    --hidden-import msal_extensions ^
    run.py

if errorlevel 1 (
    echo [ERROR] Build fallido.
    exit /b 1
)

echo.
echo [OK] Build completo: dist\ProtocolosApp\ProtocolosApp.exe
echo     Para distribuir: copia la CARPETA ENTERA dist\ProtocolosApp\
echo     (contiene el .exe + _internal\ con todas las dependencias).
endlocal
