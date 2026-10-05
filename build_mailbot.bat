@echo off
setlocal

REM ============================================================
REM Build .exe dedicado del MailBot Fedrigoni (sin GUI)
REM Empaqueta TODO en una carpeta dist\MailBot\
REM Doble click sobre MailBot.exe corre el mailbot.
REM Para programar con Task Scheduler: apunta a MailBot.exe.
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

if exist build_mailbot rmdir /s /q build_mailbot
if exist dist\MailBot rmdir /s /q dist\MailBot
if exist MailBot.spec del /q MailBot.spec

REM --console: dejamos consola visible para que se vean los logs del bot.
REM Si despues queres ocultarla en el server (Task Scheduler la maneja),
REM cambiar a --windowed o bien correr con "Run hidden" en Task Scheduler.
pyinstaller --noconfirm --onedir --console --name MailBot ^
    --icon "assets\Icono.ico" ^
    --workpath build_mailbot ^
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
    run_mailbot.py

if errorlevel 1 (
    echo [ERROR] Build fallido.
    exit /b 1
)

echo.
echo [OK] Build completo: dist\MailBot\MailBot.exe
echo     Doble click sobre el .exe corre el MailBot Fedrigoni.
echo     Para Task Scheduler: action = ese .exe, sin argumentos.
echo                          Trigger: Daily + repeat every 10 minutes indefinitely.
echo     Para login inicial AGENTE en el server: MailBot.exe --setup-auth
endlocal
