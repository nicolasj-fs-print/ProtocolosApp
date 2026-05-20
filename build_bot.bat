@echo off
setlocal

REM ============================================================
REM Build .exe dedicado del modo automatico (bot, sin GUI)
REM Empaqueta TODO en una carpeta dist\ProtocolosBot\
REM Doble click sobre ProtocolosBot.exe corre el bot.
REM Para programar con Task Scheduler: apunta a ProtocolosBot.exe.
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

if exist build_bot rmdir /s /q build_bot
if exist dist\ProtocolosBot rmdir /s /q dist\ProtocolosBot
if exist ProtocolosBot.spec del /q ProtocolosBot.spec

REM --console: dejamos consola visible para que se vean los logs del bot.
REM Si despues queres ocultarla en el server (Task Scheduler la maneja),
REM cambiar a --windowed o bien correr con "Run hidden" en Task Scheduler.
pyinstaller --noconfirm --onedir --console --name ProtocolosBot ^
    --icon "assets\Icono.ico" ^
    --workpath build_bot ^
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
    run_bot.py

if errorlevel 1 (
    echo [ERROR] Build fallido.
    exit /b 1
)

echo.
echo [OK] Build completo: dist\ProtocolosBot\ProtocolosBot.exe
echo     Doble click sobre el .exe corre el modo automatico.
echo     Para Task Scheduler: action = ese .exe, sin argumentos.
echo     Para login inicial en el server: ProtocolosBot.exe --setup-auth
endlocal
