"""Borrador de mail en Outlook (no envía: solo abre la ventana)."""
from __future__ import annotations

import os
import urllib.parse
from pathlib import Path
from typing import Sequence

from ..utils.logger import get_logger

log = get_logger(__name__)


def _find_classic_outlook_exe() -> str | None:
    """Localiza el .exe de Classic Outlook. Primero registro, después paths
    típicos de Office 365/2019/2016 (todos usan carpeta Office16)."""
    # 1) Registro: App Paths\OUTLOOK.EXE.
    try:
        import winreg
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(
                    hive,
                    r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE",
                ) as key:
                    path, _ = winreg.QueryValueEx(key, None)
                    if path and Path(path).is_file():
                        return path
            except OSError:
                continue
    except Exception as e:
        log.debug("Registro App Paths OUTLOOK.EXE no encontrado: %s", e)

    # 2) Paths típicos.
    candidates = [
        r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
        r"C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE",
    ]
    for c in candidates:
        if Path(c).is_file():
            return c
    return None


def _outlook_is_running() -> bool:
    """True si hay al menos un proceso OUTLOOK.EXE activo."""
    try:
        import subprocess
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq OUTLOOK.EXE", "/NH"],
            capture_output=True, text=True, timeout=5,
            creationflags=0x08000000,  # CREATE_NO_WINDOW
        )
        return "OUTLOOK.EXE" in result.stdout.upper()
    except Exception:
        return False


def _ensure_outlook_running(timeout: float = 8.0) -> bool:
    """Si Classic Outlook NO está corriendo, lo arranca y espera a que esté
    listo. Devuelve True si quedó disponible, False si no se pudo."""
    import time as _t
    if _outlook_is_running():
        return True

    exe = _find_classic_outlook_exe()
    if not exe:
        log.warning("No se encontró OUTLOOK.EXE de Classic Outlook.")
        return False

    log.info("Outlook no estaba corriendo. Arrancando: %s", exe)
    try:
        import subprocess
        subprocess.Popen([exe], close_fds=True)
    except Exception as e:
        log.warning("No se pudo arrancar Outlook: %s", e)
        return False

    t0 = _t.time()
    while _t.time() - t0 < timeout:
        if _outlook_is_running():
            _t.sleep(0.5)  # margen extra para que Outlook complete su init
            return True
        _t.sleep(0.3)
    return False


def _dispatch_outlook_with_retry(max_attempts: int = 3):
    """Intenta `Dispatch("Outlook.Application")` con reintentos.
    Si falla la primera vez, intenta arrancar Outlook y reintentar.
    """
    import time as _t
    import win32com.client  # type: ignore
    last_exc: Exception | None = None
    for attempt in range(1, max_attempts + 1):
        try:
            return win32com.client.Dispatch("Outlook.Application")
        except Exception as e:
            last_exc = e
            log.warning("Dispatch Outlook falló (intento %d/%d): %s", attempt, max_attempts, e)
            # Tras el primer fallo, intentar arrancar Outlook explícitamente.
            if attempt == 1:
                _ensure_outlook_running()
            _t.sleep(1.0 * attempt)
    raise last_exc or RuntimeError("No se pudo conectar con Outlook")


def _read_outlook_signature() -> str:
    """Lee la firma default del usuario directo desde el filesystem de Outlook.

    Outlook guarda las firmas en `%APPDATA%\\Microsoft\\Signatures\\*.htm`.
    El nombre de la firma default está en el registro:
        HKCU\\Software\\Microsoft\\Office\\<ver>\\Common\\MailSettings\\NewSignature
    """
    base = Path(os.getenv("APPDATA", "")) / "Microsoft" / "Signatures"
    if not base.is_dir():
        return ""

    # 1) Intentar leer el nombre de la firma default del registro.
    sig_name: str | None = None
    try:
        import winreg
        for office_ver in ("16.0", "15.0", "14.0"):
            try:
                key_path = rf"Software\Microsoft\Office\{office_ver}\Common\MailSettings"
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                    val, _ = winreg.QueryValueEx(key, "NewSignature")
                    if val:
                        sig_name = str(val)
                        break
            except OSError:
                continue
    except Exception as e:
        log.debug("No se pudo leer firma default del registro: %s", e)

    # 2) Armar lista de candidatos (firma default primero, fallback a cualquiera).
    candidates: list[Path] = []
    if sig_name:
        candidates.append(base / f"{sig_name}.htm")
    for f in base.iterdir():
        if f.is_file() and f.suffix.lower() == ".htm" and f not in candidates:
            candidates.append(f)

    # 3) Leer el primer candidato que abra OK.
    for path in candidates:
        if not path.is_file():
            continue
        for encoding in ("utf-8-sig", "utf-8", "windows-1252", "latin-1"):
            try:
                return path.read_text(encoding=encoding)
            except UnicodeDecodeError:
                continue
        try:
            return path.read_bytes().decode("latin-1", errors="ignore")
        except Exception:
            continue
    return ""


def _build_html_body(
    razon_social: str,
    comprobante: str,
    encontrados: Sequence[str],
    no_encontrados: Sequence[str],
) -> str:
    encontrados_html = "".join(f"<li>{p}</li>" for p in encontrados) or "<li>(sin protocolos)</li>"
    nf_block = ""
    if no_encontrados:
        items = "".join(f"<li>{p}</li>" for p in no_encontrados)
        nf_block = (
            "<p><b>Protocolos no encontrados en archivo (revisar):</b></p>"
            f"<ul>{items}</ul>"
        )

    return (
        "<p>Estimados,</p>"
        f"<p>Se adjuntan los protocolos correspondientes al remito <b>{comprobante}</b> "
        f"del cliente <b>{razon_social}</b>.</p>"
        "<p><b>Detalle de protocolos:</b></p>"
        f"<ul>{encontrados_html}</ul>"
        f"{nf_block}"
        "<p>Quedamos a disposición ante cualquier consulta.</p>"
        "<p>Saludos cordiales.</p>"
    )


def open_outlook_draft(
    to: str,
    razon_social: str,
    comprobante: str,
    pdf_path: Path,
    protocolos_encontrados: Sequence[str],
    protocolos_no_encontrados: Sequence[str] = (),
) -> bool:
    """Abre Outlook con un borrador listo para revisar y enviar manualmente.

    Devuelve True si se abrió Outlook, False si cayó al fallback mailto:.
    """
    subject = f"Protocolos de calidad - {razon_social} | Remito {comprobante}"
    html_body = _build_html_body(razon_social, comprobante, protocolos_encontrados, protocolos_no_encontrados)

    try:
        import time
        # Asegurar que Classic Outlook esté corriendo y conectar vía COM con
        # reintentos (cubre el caso de cold start o que Outlook esté colgado).
        _ensure_outlook_running()
        outlook = _dispatch_outlook_with_retry()
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = to or ""
        mail.Subject = subject

        # Forzar la creación del Inspector SIN mostrar ventana → Outlook carga
        # el editor y empieza a renderizar la firma (incluyendo cloud signatures).
        try:
            _ = mail.GetInspector
        except Exception:
            pass

        # Esperar a que Outlook inyecte la firma. Polling rápido (hasta ~1.5s).
        firma = ""
        for _ in range(15):
            try:
                current = mail.HTMLBody or ""
                if current.strip():
                    firma = current
                    break
            except Exception:
                pass
            time.sleep(0.1)

        # Fallback: si la firma cloud no cargó, intentar leer del filesystem.
        if not firma.strip():
            firma = _read_outlook_signature()

        # Setear el body completo (cuerpo + firma) ANTES de mostrar la ventana.
        # Importante: modificar HTMLBody DESPUÉS del Display puede dejar el
        # mail en un estado raro donde el Send no se ejecuta y queda en Drafts.
        mail.HTMLBody = html_body + firma

        if pdf_path and Path(pdf_path).exists():
            mail.Attachments.Add(str(Path(pdf_path).resolve()))

        # Recién ahora mostramos la ventana al user.
        mail.Display(False)

        log.info("Borrador Outlook abierto para %s", comprobante)
        return True
    except Exception as e:
        log.warning("Outlook no disponible (%s) → fallback mailto:", e)

    plain_body = (
        f"Estimados,\n\n"
        f"Se adjuntan los protocolos del remito {comprobante} - {razon_social}.\n"
        f"PDF: {pdf_path}\n\n"
        f"Detalle:\n" + "\n".join(f"- {p}" for p in protocolos_encontrados)
    )
    params = urllib.parse.urlencode({
        "subject": subject,
        "body": plain_body,
    }, quote_via=urllib.parse.quote)
    url = f"mailto:{to}?{params}"
    try:
        os.startfile(url)
    except OSError as e:
        log.error("No se pudo abrir mailto: %s", e)
    return False
