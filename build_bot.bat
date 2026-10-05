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
    echo         (NO se empaqueta en el .exe: se copia al lado del ejecutable.)
    exit /b 1
)

REM ============================================================
REM APP2_PG_* es OBLIGATORIO desde el 2026-09-11, y se chequea ACA
REM ARRIBA para fallar en un segundo y no despues de dos minutos de
REM PyInstaller.
REM
REM Sin esas variables el bot NO FALLA: `esta_configurado()` da False
REM y `load_ingresos()` / `load_clientes_protocolos()` caen al Excel.
REM O sea que un .env incompleto produce un .exe que anda, corre a las
REM 08:00, manda los mails... leyendo las planillas viejas, que es
REM justo lo que estos cambios vinieron a jubilar. Degradar en
REM silencio es peor que no compilar.
REM ============================================================
findstr /b /c:"APP2_PG_PASSWORD=" .env >nul 2>&1
if errorlevel 1 (
    echo [ERROR] El .env no tiene el bloque APP2_PG_*.
    echo         Sin el, el bot lee los Excel viejos EN SILENCIO en vez de app_2_db:
    echo         ni los certificados cargados desde la app ni las altas de clientes
    echo         se van a ver, y no va a dar ningun error.
    echo         Copia el bloque desde .env.example y completa APP2_PG_PASSWORD
    echo         con la credencial de agente_dsk: la misma que ya usa el server.
    exit /b 1
)
for /f "tokens=1,* delims==" %%a in ('findstr /b /c:"APP2_PG_PASSWORD=" .env') do set "_APP2PW=%%b"
if "%_APP2PW%"=="" (
    echo [ERROR] APP2_PG_PASSWORD esta vacia en el .env. Ver el mensaje de arriba.
    exit /b 1
)
set "_APP2PW="

if exist build_bot rmdir /s /q build_bot
if exist dist\ProtocolosBot rmdir /s /q dist\ProtocolosBot
if exist ProtocolosBot.spec del /q ProtocolosBot.spec

REM --console: dejamos consola visible para que se vean los logs del bot.
REM Si despues queres ocultarla en el server (Task Scheduler la maneja),
REM cambiar a --windowed o bien correr con "Run hidden" en Task Scheduler.
REM
REM --collect-all psycopg2 NO es opcional: el espejo a app_2_db hace el
REM "import psycopg2" DENTRO de una funcion y dentro de un try (ver
REM app/sharepoint/lists.py, _espejo_app2). PyInstaller suele detectarlo, pero si
REM no lo hace el bot corre igual y el espejo falla en silencio a las 6 AM sin
REM que nadie lo vea. Se declara explicito, que sale gratis.
REM    --add-data ".env;." <-- SACADO A PROPOSITO (2026-09-04)
REM
REM    El .env va SIEMPRE al lado del .exe, nunca adentro del bundle.
REM    `runtime_dir()` (app/utils/paths.py) busca primero el .env junto al
REM    ejecutable y recien despues el del bundle, asi que empaquetarlo creaba un
REM    FALLBACK SILENCIOSO al .env de DESARROLLO: si en el server faltara el
REM    archivo, el bot arrancaria con TEST_MODE_OVERRIDE_EMAIL puesto y
REM    redirigiria TODOS los mails de clientes a una casilla interna sin que nadie
REM    se entere. Sin el bundle, si falta el .env el bot falla al conectarse a SQL
REM    y se ve en el log.
pyinstaller --noconfirm --onedir --console --name ProtocolosBot ^
    --icon "assets\Icono.ico" ^
    --workpath build_bot ^
    --add-data "assets;assets" ^
    --add-data "tools\SumatraPDF.exe;tools" ^
    --add-data "tools\msodbcsql18.msi;tools" ^
    --collect-all customtkinter ^
    --collect-all tkcalendar ^
    --collect-submodules pyodbc ^
    --collect-submodules app ^
    --collect-all psycopg2 ^
    --hidden-import win32com.client ^
    --hidden-import PIL._tkinter_finder ^
    --hidden-import msal ^
    --hidden-import msal_extensions ^
    run_bot.py

if errorlevel 1 (
    echo [ERROR] Build fallido.
    exit /b 1
)

REM ============================================================
REM El .env AL LADO DEL .EXE. No es un detalle: `runtime_dir()`
REM (app/utils/paths.py) lo busca ahi primero, y como NO va adentro
REM del bundle (ver arriba), sin este copy el .exe arranca sin
REM configuracion: no encuentra SQL Server ni app_2_db.
REM
REM Hasta el 2026-09-11 el comentario de arriba decia "se copia al
REM lado del ejecutable" y el copy NO EXISTIA: habia que acordarse
REM de hacerlo a mano en cada deploy.
REM ============================================================
copy /y .env dist\ProtocolosBot\.env >nul
if errorlevel 1 (
    echo [ERROR] No se pudo copiar el .env al lado del .exe.
    exit /b 1
)

echo.
echo [OK] Build completo: dist\ProtocolosBot\ProtocolosBot.exe
echo     .env copiado al lado del .exe (obligatorio: no va en el bundle).
echo     Doble click sobre el .exe corre el modo automatico (cron diario).
echo     Para Task Scheduler: action = ese .exe, sin argumentos.
echo     Para login inicial en el server: ProtocolosBot.exe --setup-auth
echo.
echo Para buildear el MailBot Fedrigoni: build_mailbot.bat
endlocal
