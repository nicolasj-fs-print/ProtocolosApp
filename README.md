# Generador de Protocolos de Calidad

App de escritorio Python que reemplaza un flujo legacy de Power BI + Power Automate + Encodian para generar PDFs consolidados de protocolos de calidad por comprobante/remito a partir de SQL Server, leyendo y escribiendo archivos directo en SharePoint vía Microsoft Graph.

---

## Pre-requisitos

- Windows 10/11.
- Python 3.11+ (solo para desarrollo / build; los usuarios finales reciben el `.exe`).
- Cuenta corporativa con acceso al sitio SharePoint donde viven los archivos.
- IP autorizada en el servidor SQL (si la DB tiene whitelist por IP).
- Outlook instalado (para los borradores de mail).
- **ODBC Driver 18 for SQL Server**: la app lo instala sola al primer arranque desde el `.msi` embebido (`tools/msodbcsql18.msi`). No hace falta hacerlo manualmente.

---

## Setup local (dev)

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env
# Editá .env con credenciales y paths reales (ver tabla abajo).
python -m app.main
```

### Modo MOCK (sin SQL)

Para probar el pipeline PDF sin DB, en `.env`:
```
MOCK_DATA=true
```
Trae un dataset hardcoded en `data_service._mock_dataframe()` para validar la generación sin tocar SQL.

---

## Configurar `.env`

| Variable | Descripción |
|---|---|
| `SQL_DRIVER` | Default `ODBC Driver 18 for SQL Server`. |
| `SQL_SERVER` | Host del SQL Server. Formato `host\instancia` o `host,puerto`. |
| `SQL_DATABASE` | Nombre de la base de datos principal. |
| `SQL_DATABASE_TRAZABILIDAD` | DB secundaria para el lookup de trazabilidad. |
| `SQL_TRUSTED_CONNECTION` | `false` para SQL auth, `true` para Windows auth. |
| `SQL_USER` / `SQL_PASSWORD` | Credenciales SQL si `TRUSTED_CONNECTION=false`. |
| `SQL_TRUST_SERVER_CERT` | `yes` (default, evita "self-signed certificate"). |
| `SQL_ENCRYPT` | `yes` (default). |
| `STORAGE_BACKEND` | `sharepoint` (recomendado) o `local`. |
| `MS_TENANT_ID` | GUID del tenant en Azure AD (pedir a TI). |
| `MS_CLIENT_ID` | GUID de la App Registration creada para esta app. |
| `SHAREPOINT_HOSTNAME` | `tutenant.sharepoint.com`. |
| `SHAREPOINT_SITE` | Nombre del sitio SP donde están los datos. |
| `SP_CLIENTS_FILE` | Path SP del Excel de clientes. |
| `SP_INGRESOS_FILE` | Path SP del Excel de ingresos. |
| `SP_PROTOCOLS_FOLDER` | Carpeta SP de PDFs originales. |
| `SP_OUTPUT_FOLDER` | Carpeta SP donde van los PDFs generados. |
| `SP_LOGS_FOLDER` / `SP_EJECUCIONES_FOLDER` / `SP_USUARIOS_FILE` | Telemetría en una sub-carpeta dedicada. |
| `MOCK_DATA` | `true` para usar dataset de prueba. |
| `LOG_LEVEL` | `INFO` por default. |

> Ver `.env.example` para el listado completo con valores placeholder.

---

## Setup Azure AD (App Registration)

La app necesita una **App Registration** en el tenant Azure AD para autenticarse vía Microsoft Graph y acceder a SharePoint.

1. Portal: https://entra.microsoft.com → **Aplicaciones** → **Registros de aplicaciones** → **Nuevo registro**.
2. Configurar:
   - Nombre: `ProtocolosApp` (o lo que prefieras).
   - Cuentas: **Solo cuentas de este directorio organizativo (Inquilino único)**.
   - URI de redirección: **Cliente público / nativo (móvil y escritorio)** → `http://localhost`.
3. Copiar de la pantalla **Información general**:
   - `Id. de aplicación (cliente)` → `MS_CLIENT_ID`.
   - `Id. de directorio (inquilino)` → `MS_TENANT_ID`.
4. Permisos de API → **Microsoft Graph** → **Permisos delegados**:
   - `Files.ReadWrite.All`
   - `Sites.ReadWrite.All`
   - `User.Read`
5. Conceder consentimiento de administrador para el tenant (si la política lo requiere).
6. Autenticación → **Permitir flujos de cliente público** = **Sí**.

---

## Flujo de la app

1. **Arranque** (splash screen `SplashWindow`):
   - Si falta el ODBC Driver 18 → ofrece instalarlo desde el `.msi` embebido.
   - Test rápido de conexión SQL (5s). Si falla → dialog "Verificar IP / credenciales" y la app no abre.
   - Auth Microsoft 365 vía MSAL (browser interactivo la 1ª vez, después token cacheado en `%APPDATA%`).
   - Resuelve site + drive de SharePoint.
   - Pre-crea carpeta de salida si no existe.
   - Pre-descarga Excels (clientes + ingresos) al cache en `%APPDATA%`.
   - Pre-popula el index de PDFs originales en memoria.

2. **GUI principal** (`MainWindow`):
   - Selector **Desde / Hasta** (DateEntry).
   - Listas seleccionables **Clientes** y **Comprobantes** con cross-filter (multi-select).
   - Tabla **Resumen** por comprobante (M2 total, observaciones).
   - Tabla **Detalle SAFED** con marcador rojo `⚠ FALTA PROTOCOLO` cuando no hay match.
   - Botones: Buscar / Borrar filtros / Abrir carpeta / Generar protocolos.

3. **Buscar**: query SQL → enrich con trazabilidad → cruce con Excel ingresos → filtro Excel clientes.

4. **Generar protocolos**:
   - Si hay items sin protocolo → dialog modal con lista + botón Exportar (CSV/Excel) + opciones "Generar todos" / "Omitir problemáticos" / "Cancelar".
   - Worker en thread genera 1 PDF por comprobante = carátula + (detalle por protocolo + PDF original con header overlay) + paginación `X de N`.
   - Sube cada PDF a la carpeta SP de salida vía Graph.
   - Si el PDF ya existe en SP → dialog "Remitos ya generados" con botón "Abrir carpeta".
   - Al final: dialog "¿Abrir N borradores en Outlook?" → opcional.

5. **Telemetría** (thread daemon, post-done):
   - Sube `resumen_ejecucion_*.xlsx` a la carpeta de Ejecuciones.
   - Sube el `.log` de la corrida a la carpeta de Logs.
   - Actualiza `usuarios.xlsx` con Email / Nombre / Sesiones / Ejecuciones.

---

## Cache local (en `%APPDATA%\ProtocolosApp\`)

- `token.cache` — token MSAL refresh.
- `logs\run_YYYYMMDD_HHMMSS.log` — log local de cada ejecución (se sube a SP al final).
- `cache\` — Excels descargados de SharePoint con metadata (`lastModifiedDateTime`).
  - En cada Buscar, 1 GET de metadata a SP. Si el Excel no cambió → se lee del disco (instantáneo).
  - Si cambió → re-descarga y reemplaza el cache.

---

## Empaquetar como `.exe` (one-folder)

Genera **carpeta `dist\ProtocolosApp\`** con el `.exe` + `_internal\` (deps).
Arranque rápido (sin descompresión en runtime). Para distribuir, copiá la carpeta entera (o zippeala).

### Opción A — Script (recomendado)

```powershell
.venv\Scripts\activate
build_exe.bat
```

### Opción B — Comando manual

```powershell
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
```

**Salida**: `dist\ProtocolosApp\ProtocolosApp.exe` + `dist\ProtocolosApp\_internal\`.

### Notas
- El `.env` está adentro del bundle. Para cambiarlo hay que re-buildear.
- Cache de tokens, logs y Excels van a `%APPDATA%\ProtocolosApp\` (fuera de la carpeta del .exe).
- ODBC Driver 18 se auto-instala al primer arranque si falta.
- Compatible con cualquier user del tenant con acceso al sitio SP.

---

## Estructura del proyecto

```
app/
  main.py                  entrypoint
  config.py                carga .env
  gui/
    main_window.py         ventana principal
    splash.py              splash inicial (test SQL + auth + warm-up)
    dialogs.py             MissingProtocols + AlreadyGenerated
    widgets.py             DataTable, MultiSelectListbox
  db/
    connection.py          pyodbc + quick_test_connection
    queries.py             queries SQL (cabecera + items + clientes + productos + trazabilidad)
  services/
    data_service.py        fetch SQL + Excel clientes + enrich protocolos + cache disco
    trazabilidad_service.py  lookup DB secundaria de trazabilidad
    ingresos_service.py    Excel de cruce de ingresos
    protocol_service.py    búsqueda PDFs originales en SP
    pdf_service.py         orquesta carátula + detalle + merge por comprobante
    mail_service.py        borradores Outlook vía win32com
  pdf/
    caratula.py, detalle.py, header.py, merger.py, common.py, styles.py
  sharepoint/
    auth.py                MSAL public client + token cache
    client.py              wrapper Graph API (download, upload, ensure_folder)
    users.py               tracking usuarios.xlsx
  utils/
    paths.py, logger.py, validators.py
    disk_cache.py          cache en %APPDATA% con eTag/lastModified
    odbc_check.py          detecta + auto-instala ODBC Driver 18
    io.py, format.py
  workers/
    generation_worker.py   thread + queue para PDFs + telemetría post-done
assets/
  Icono.ico, logo_fs.png, logo_fs_dark.png, logo_fedrigoni.png, logo_fedrigoni_dark.png
tools/
  SumatraPDF.exe (visor), msodbcsql18.msi (auto-instalación)
run.py                     entrypoint para PyInstaller (imports absolutos)
build_exe.bat              script de build
.env.example, requirements.txt
```

---

## Mail

`mail_service.open_outlook_draft` abre un **borrador en Outlook** (no envía):
- Destinatarios: campo `Mail Protocolos` del Excel de clientes.
- Asunto: `Protocolos de calidad - {RazonSocial} | Remito {Comprobante}`.
- Cuerpo HTML con lista de protocolos detectados / no encontrados.
- Adjunto: el PDF generado (copia local en `%TEMP%\ProtocolosApp\`).

Fallback a `mailto:` si Outlook no responde (sin adjunto).

---

## Modo automático (cron diario)

A partir de Iteración 7, el mismo `.exe` admite un modo **automático** sin GUI, pensado para correr en un Windows Server vía Task Scheduler. Genera y envía los protocolos del día anterior sin intervención humana.

### Diferencias con la GUI

| | GUI manual | Modo `--auto` |
|--|--|--|
| Interfaz | customtkinter | Sin ventana (consola + log) |
| Mail | Borrador en Outlook (Classic) | Enviado directo vía Graph Mail API |
| Items sin protocolo | Dialog interactivo | Mail a `CONTROL_EMAIL` |
| Tracking | No | SharePoint List (idempotente) |
| Rango procesado | Lo elige el usuario | `today - AUTO_DAYS_BACK` → ayer |

### Comportamiento

Para cada remito del rango:
- Si **todos** los items SAFED tienen protocolo → genera el PDF, lo sube a SP y manda el mail al cliente. Marca `enviado_cliente` en el tracking.
- Si **falta al menos uno** → NO manda al cliente. Manda aviso a `CONTROL_EMAIL` con la lista de faltantes. Marca `pending_control`.
- Si **falla** → log de error, no escribe tracking → reintenta al día siguiente.

Cuando el operador resuelve un caso `pending_control` desde la GUI manual (genera PDF + abre borrador Outlook), la GUI marca automáticamente el tracking como `resuelto_manual`. El bot al día siguiente NO lo re-procesa.

### Setup en el server

1. **Buildear `.exe`** en tu maquina dev y copiarlo al server (ej. `C:\ProtocolosApp\`).
2. **Completar `.env` del server** con las variables del modo auto:
   ```
   CONTROL_EMAIL=mail-de-control@fs-print.com
   AUTO_DAYS_BACK=7
   SP_TRACKING_LIST_NAME=Tracking Envios
   ```
3. **Login inicial** (una sola vez, con interacción humana):
   ```powershell
   C:\ProtocolosApp\ProtocolosApp.exe --setup-auth
   ```
   Abre browser, logueás con la cuenta del bot, aceptás permisos (incluye `Mail.Send`). Token cache queda en `%APPDATA%\ProtocolosApp\token.cache`.

4. **Dry run** (opcional, no manda nada):
   ```powershell
   C:\ProtocolosApp\ProtocolosApp.exe --auto --dry-run
   ```
   Verifica que la auth, fetch SQL y tracking funcionen. Loguea qué mails se mandarían sin enviarlos.

5. **Programar Task Scheduler**:
   - **Action**: `C:\ProtocolosApp\ProtocolosApp.exe`
   - **Arguments**: `--auto`
   - **Trigger**: Daily, hora a elección (ej. 06:00).
   - **Run whether user is logged on or not**: ✓ (con la cuenta del bot).
   - **Settings**: "If task fails, restart every 10 min" / 3 intentos.

### Exit codes
- `0` — OK (procesó lo que correspondía, incluso si no había nada para procesar).
- `1` — Error fatal. Revisar `%APPDATA%\ProtocolosApp\logs\run_*.log`.
- `2` — Auth falló. Reloggear: `ProtocolosApp.exe --setup-auth`.

### SharePoint List `Tracking Envios`

El bot la crea sola la primera vez en el sitio configurado en `SP_APP_SITE`. Columnas:

| Columna | Descripción |
|--|--|
| `Title` | `#Comprobante` (clave única) |
| `Cliente` | Código de cliente |
| `RazonSocial` | Razón social |
| `FechaRemito` | Fecha del remito en SQL |
| `FechaProcesado` | Datetime ISO de cuándo lo procesó el bot |
| `Estado` | `enviado_cliente` / `pending_control` / `resuelto_manual` / `error` |
| `MailDestino` | A quién se mandó |
| `PdfUrl` | URL del PDF en SP |
| `ItemsFaltantes` | JSON con los items SAFED sin protocolo (solo en `pending_control`) |
| `RunId` | Timestamp de la corrida que escribió este row |

Se puede consultar/editar a mano desde el browser para casos puntuales.

---

## Troubleshooting

- **"No se pudo establecer conexión a SQL"** al arrancar → verificar que tu IP esté autorizada en el servidor SQL y las credenciales en `.env`.
- **"Falta ODBC Driver 18"** → aceptá el prompt para que se instale solo (puede pedir UAC). Si decís No, la app cierra.
- **"Sin conexión a SharePoint"** → revisar internet o que tu cuenta tenga acceso al sitio SharePoint configurado.
- **El .exe se cierra al abrirse** → revisar `%APPDATA%\ProtocolosApp\logs\run_*.log`.
- **Outlook abre el viejo "Classic" y no el "New"** → el New Outlook no soporta COM/adjuntos por API. La app usa Classic Outlook (clásico). Es la única opción técnica.

---

## Licencia / Uso

Proyecto interno. Las credenciales reales, paths y datos del tenant van en `.env` (no se commitea — ver `.gitignore`).
