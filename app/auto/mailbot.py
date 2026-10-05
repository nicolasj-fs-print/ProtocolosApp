"""Orquestador del MailBot Fedrigoni (Iteración 8).

Trigger: Task Scheduler cada 10 min en el server. Polea la mailbox AGENTE
(token cache separado) y procesa mails entrantes con PDFs de Fedrigoni.

Flujo:
  Fase A — Re-chequear filas con LF="No ingresado a softland":
    - Si ahora hay match → actualizar la celda LF.
    - Si pasaron MAILBOT_LF_ESCALATE_DAYS y sigue sin match → escalar.

  Fase B — Procesar mails nuevos del Inbox:
    - Filtrar por sender (F1 o F2).
    - Parsear PDF (autodetect si el sender está en ambas listas).
    - Calcular protocolo, buscar LF, buscar producto (F2), escribir fila.
    - Renombrar PDF, subir a SharePoint, mover mail a Procesados.
    - En caso de error → mover a Errores + mail a CONTROL_EMAIL.

  Fase C — Cierre: subir log, exit.

MODO TEST LOCAL (MAILBOT_TEST_LOCAL=true en .env):
  - SKIP Fase A (no toca el Excel para nada).
  - Fase B sí corre, pero en lugar de escribir al Excel + subir a SP,
    guarda el PDF renombrado en `test_pdfs/` junto al .exe y manda mail
    al MAILBOT_TEST_EMAIL con los datos de la fila.
  - Los mails procesados igual se mueven a ProtocolosProcesados / Errores.
"""
from __future__ import annotations

import sys
from datetime import date, datetime
from pathlib import Path
from typing import Iterable

from ..config import get_settings
from ..services.ingresos_service import (
    LF_ESCALADO,
    LF_PENDIENTE,
    IngresosExcelWriter,
)
from ..services.trazabilidad_service import (
    find_artcod_by_producto,
    lookup_trazabilidad_by_nf,
)
from ..sharepoint.auth import AuthError, MAIL_READWRITE_SCOPE, ensure_authenticated, get_token
from ..sharepoint.client import SharePointError, get_client
from ..utils.logger import current_run_id, get_logger, upload_run_log_to_sharepoint
from . import mail_reader
from . import pdf_parser
from .mailer_graph import send_mail

log = get_logger(__name__)

EXIT_OK = 0
EXIT_FATAL = 1
EXIT_AUTH = 2


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _split_csv(s: str) -> list[str]:
    return [p.strip().lower() for p in (s or "").replace(",", ";").split(";") if p.strip()]


def _classify_sender(sender: str, sf1: list[str], sf2: list[str]) -> str:
    """'F1' si el sender está SOLO en sf1, 'F2' si solo en sf2,
    'AMBIGUOUS' si está en ambos, 'UNKNOWN' si no aparece."""
    s = (sender or "").strip().lower()
    in1 = s in sf1
    in2 = s in sf2
    if in1 and not in2:
        return "F1"
    if in2 and not in1:
        return "F2"
    if in1 and in2:
        return "AMBIGUOUS"
    return "UNKNOWN"


def _build_lf_value(traza: dict | None) -> str:
    """Convierte el dict de lookup_trazabilidad_by_nf a string para la celda."""
    if not traza:
        return LF_PENDIENTE
    fc = traza.get("FormularioCodigo", "")
    fn = traza.get("FormularioNumero", "")
    pr = traza.get("Producto", "")
    return f"{fc} - {fn} - {pr}".strip(" -")


def _parse_fecha_str(s: str) -> date | None:
    """Acepta 'dd/MM/yyyy', 'dd-MM-yyyy', etc. Devuelve `date` o None."""
    s = (s or "").strip()
    for sep in ("/", "-", "."):
        parts = s.split(sep)
        if len(parts) == 3:
            try:
                d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
                if y < 100:
                    y += 2000
                return date(y, m, d)
            except ValueError:
                pass
    return None


def _notify_control(subject: str, html: str, control_email: str) -> None:
    """Best-effort: si falla el send_mail, log warning y seguir."""
    if not control_email:
        log.warning("CONTROL_EMAIL vacío, no se manda aviso: %s", subject)
        return
    try:
        send_mail(to=[control_email], subject=subject, html_body=html)
    except Exception as e:
        log.warning("No se pudo mandar aviso a control: %s", e)


def _build_new_pdf_name(pr: str, nf: str, original_name: str, formato: str) -> str:
    """Genera el nombre final del PDF.

    F1: `{PR} - {NF} - {original}.pdf`
    F2: el original ya empieza con NF, agregamos solo `{PR} - {original}`.
    """
    orig = Path(original_name).name
    if formato == "F1":
        # Si el original ya empieza con PR/NF por re-procesos, no duplicar.
        return f"{pr} - {nf} - {orig}"
    # Formato 2: original tipo "91684 - THERMAL TOP BPA FREE FSC P7 YG55.pdf"
    return f"{pr} - {orig}"


# ---------------------------------------------------------------------------
# Test mode helpers
# ---------------------------------------------------------------------------
def _test_pdfs_dir() -> Path:
    """Carpeta `test_pdfs/` adyacente al .exe (o a la raíz del proyecto en dev)."""
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).resolve().parent
    else:
        # Dev: raíz del repo (subimos desde app/auto/mailbot.py).
        base = Path(__file__).resolve().parent.parent.parent
    d = base / "test_pdfs"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _build_test_email_html(
    *,
    nf: str,
    fecha: str,
    lf: str,
    producto: str,
    pr: str,
    formato: str,
    original_name: str,
    new_name: str,
    local_path: Path,
) -> str:
    """Cuerpo HTML del mail que se manda en modo test con los datos de la fila."""
    return (
        "<p>El <b>MailBot</b> procesó un PDF de Fedrigoni en <b>MODO TEST LOCAL</b>.</p>"
        "<p>Datos que se <b>HABRÍAN</b> escrito al Excel "
        "<i>Protocolos x Ingreso OK.xlsx</i>:</p>"
        "<table border='1' cellpadding='6' style='border-collapse:collapse;font-family:Segoe UI,sans-serif;font-size:13px'>"
        "<tr style='background:#f0f0f0'><th align='left'>Columna</th><th align='left'>Valor</th></tr>"
        f"<tr><td>Packing List</td><td>{nf}</td></tr>"
        f"<tr><td>FC Fedrigoni</td><td>{nf}</td></tr>"
        f"<tr><td>Fecha</td><td>{fecha}</td></tr>"
        f"<tr><td>LF</td><td>{lf}</td></tr>"
        f"<tr><td>Producto</td><td>{producto}</td></tr>"
        f"<tr><td>Protocolo</td><td><b>{pr}</b></td></tr>"
        "</table>"
        "<br/>"
        f"<p><b>Formato detectado:</b> {formato}</p>"
        f"<p><b>PDF original:</b> {original_name}</p>"
        f"<p><b>PDF renombrado guardado en:</b><br/><code>{local_path}</code></p>"
        "<p style='color:#666;font-size:12px'>"
        "⚠ <b>NO</b> se subió el PDF a SharePoint ni se modificó el Excel real. "
        "Para producción, vaciar MAILBOT_TEST_LOCAL en el .env."
        "</p>"
    )


# ---------------------------------------------------------------------------
# Fase A — Re-chequeo de filas pendientes de LF
# ---------------------------------------------------------------------------
def _process_pending_lf(
    writer: IngresosExcelWriter,
    today: date,
    escalate_days: int,
    control_email: str,
) -> dict:
    """Re-chequea filas con LF=No ingresado a softland. Devuelve resumen."""
    stats = {"reviewed": 0, "resolved": 0, "escalated": 0, "still_pending": 0}
    try:
        pending = writer.find_pending_lf_rows()
    except Exception as e:
        log.exception("Error leyendo filas pendientes de LF: %s", e)
        return stats

    stats["reviewed"] = len(pending)
    if not pending:
        return stats

    log.info("Fase A: %d fila(s) pendiente(s) de LF a re-chequear.", len(pending))

    for row in pending:
        ri = row["row_index"]
        nf = row["fc_fedrigoni"] or row["packing_list"]
        if not nf:
            log.warning("Fila %d sin NF/Packing → skip.", ri)
            continue
        try:
            traza = lookup_trazabilidad_by_nf(nf)
        except Exception as e:
            log.exception("Lookup trazabilidad falló para NF=%s: %s", nf, e)
            continue

        if traza:
            new_lf = _build_lf_value(traza)
            try:
                writer.update_lf(ri, new_lf)
                stats["resolved"] += 1
                log.info("Fila %d (NF=%s): LF resuelto → %s", ri, nf, new_lf)
            except Exception as e:
                log.exception("update_lf falló para fila %d: %s", ri, e)
            continue

        # No hay match todavía. Chequear si hay que escalar.
        fecha_pdf = _parse_fecha_str(row["fecha"])
        if fecha_pdf is None:
            log.warning("Fila %d: Fecha inválida '%s' → no se puede escalar.", ri, row["fecha"])
            stats["still_pending"] += 1
            continue

        days_old = (today - fecha_pdf).days
        if days_old >= escalate_days:
            try:
                writer.update_lf(ri, LF_ESCALADO)
                stats["escalated"] += 1
                log.warning(
                    "Fila %d (NF=%s): escalada (%d días sin LF).",
                    ri, nf, days_old,
                )
                # Aviso a control.
                html = (
                    f"<p>El protocolo <b>{row['protocolo']}</b> sigue sin LF en Trazabilidad "
                    f"después de <b>{days_old} días</b> desde la fecha del PDF "
                    f"({row['fecha']}).</p>"
                    f"<ul>"
                    f"<li>Packing List: {row['packing_list']}</li>"
                    f"<li>NF (FC Fedrigoni): {row['fc_fedrigoni']}</li>"
                    f"<li>Producto: {row['producto']}</li>"
                    f"</ul>"
                    f"<p>La fila quedó marcada como '⚠ Sin LF - escalado' y el bot "
                    f"no la volverá a chequear. Resolver manualmente.</p>"
                )
                _notify_control(
                    subject=f"[MailBot] Escalado - {row['protocolo']} sin LF tras {days_old}d",
                    html=html,
                    control_email=control_email,
                )
            except Exception as e:
                log.exception("Escalado falló para fila %d: %s", ri, e)
        else:
            stats["still_pending"] += 1
            log.info(
                "Fila %d (NF=%s): pendiente (%d/%d días).",
                ri, nf, days_old, escalate_days,
            )
    return stats


# ---------------------------------------------------------------------------
# Fase B — Procesar mails nuevos
# ---------------------------------------------------------------------------
def _process_inbox(
    writer: IngresosExcelWriter,
    senders_f1: list[str],
    senders_f2: list[str],
    processed_folder_id: str,
    errors_folder_id: str,
    control_email: str,
    sp_protocols_folder: str,
    dry_run: bool,
    test_local: bool = False,
    test_email: str = "",
) -> dict:
    stats = {"checked": 0, "ok": 0, "errors": 0, "skipped_existing": 0, "ignored_sender": 0}

    # Mails de los senders configurados, no leídos.
    # top=200: si el Inbox tiene mucho ruido acumulado (notifs, alertas, etc.),
    # con 50 corremos riesgo de que los mails reales de Fedrigoni queden fuera
    # del top. Fedrigoni manda pocos mails; 200 cubre cualquier caso razonable.
    senders_all = list(set(senders_f1) | set(senders_f2))
    try:
        msgs = mail_reader.list_unread_in_inbox(senders=senders_all, top=200)
    except Exception as e:
        log.exception("No se pudo listar mails del Inbox: %s", e)
        return stats

    if not msgs:
        log.info("Fase B: no hay mails nuevos.")
        return stats

    log.info("Fase B: %d mail(s) por procesar.", len(msgs))
    sp_client = get_client()

    for msg in msgs:
        stats["checked"] += 1
        log.info("[%s] subject=%r from=%s", msg.id[:8], msg.subject, msg.from_address)

        formato_by_sender = _classify_sender(msg.from_address, senders_f1, senders_f2)
        if formato_by_sender == "UNKNOWN":
            log.info("[%s] sender desconocido → ignoro.", msg.id[:8])
            stats["ignored_sender"] += 1
            continue

        # --- Descargar adjunto ---
        try:
            attachments = mail_reader.get_message_attachments(msg.id)
        except Exception as e:
            log.exception("[%s] no se pudo listar adjuntos: %s", msg.id[:8], e)
            stats["errors"] += 1
            _move_to_errors(msg.id, errors_folder_id, dry_run)
            _notify_control(
                f"[MailBot] Error listando adjuntos - {msg.subject}",
                f"<p>Mail de {msg.from_address}</p><p>{e}</p>",
                control_email,
            )
            continue

        pdf_attachments = [
            a for a in attachments
            if (a.get("name") or "").lower().endswith(".pdf")
        ]
        if not pdf_attachments:
            # Sin PDF → NO es un mail válido de Fedrigoni. Decisión:
            # - Si el sender es de Fedrigoni → eso es raro, ESCALAR (Fedrigoni
            #   siempre manda PDF; sin PDF probablemente es un error en origen).
            # - Si el sender es "de prueba" (ej. tu mail nicolas.jerez puesto
            #   en las listas) o un mail que se coló → marcar leído y skipear
            #   en silencio. Sin esto, cada notificación de OneDrive/alertas
            #   spamearía a control.
            from_domain = (msg.from_address.split("@")[-1] if "@" in msg.from_address else "").lower()
            is_fedrigoni = "fedrigoni" in from_domain
            if is_fedrigoni:
                log.warning("[%s] sender Fedrigoni SIN PDF → errores.", msg.id[:8])
                stats["errors"] += 1
                _move_to_errors(msg.id, errors_folder_id, dry_run)
                _notify_control(
                    f"[MailBot] Mail Fedrigoni sin PDF - {msg.subject}",
                    f"<p>El mail de <b>{msg.from_address}</b> no trae PDF adjunto. Revisar manualmente.</p>",
                    control_email,
                )
            else:
                log.info(
                    "[%s] sin PDF y sender no-Fedrigoni (%s) → marcar leído y skip.",
                    msg.id[:8], msg.from_address,
                )
                stats["ignored_sender"] += 1
                if not dry_run:
                    try:
                        mail_reader.mark_as_read(msg.id)
                    except Exception as e:
                        log.warning("[%s] mark_as_read falló: %s", msg.id[:8], e)
            continue

        att = pdf_attachments[0]
        att_name = str(att.get("name") or "attachment.pdf")
        try:
            pdf_bytes = mail_reader.download_attachment(msg.id, att["id"])
        except Exception as e:
            log.exception("[%s] no se pudo descargar adjunto: %s", msg.id[:8], e)
            stats["errors"] += 1
            _move_to_errors(msg.id, errors_folder_id, dry_run)
            _notify_control(
                f"[MailBot] Error descargando PDF - {msg.subject}",
                f"<p>{e}</p>",
                control_email,
            )
            continue

        # --- Detectar formato ---
        formato: str | None = None
        if formato_by_sender in ("F1", "F2"):
            formato = formato_by_sender
        else:  # AMBIGUOUS → autodetect por contenido
            try:
                text = pdf_parser.extract_text(pdf_bytes)
                formato = pdf_parser.autodetect_format(text)
                log.info("[%s] autodetect → %s", msg.id[:8], formato)
            except pdf_parser.PdfParseError as e:
                log.warning("[%s] autodetect falló: %s", msg.id[:8], e)
                stats["errors"] += 1
                _move_to_errors(msg.id, errors_folder_id, dry_run)
                _notify_control(
                    f"[MailBot] Autodetect formato falló - {msg.subject}",
                    f"<p>De: {msg.from_address}</p><p>{e}</p>",
                    control_email,
                )
                continue

        # --- Parsear ---
        try:
            if formato == "F1":
                parsed = pdf_parser.parse_formato_1(pdf_bytes)
                nf = parsed.nota_fiscal
                fecha = parsed.fecha
                producto = parsed.cod_ref
            else:
                parsed2 = pdf_parser.parse_formato_2(pdf_bytes)
                nf = parsed2.nf
                fecha = parsed2.fecha
                # Buscar el ARTCOD por producto (match fuzzy de palabras).
                artcod = find_artcod_by_producto(parsed2.producto)
                if not artcod:
                    log.warning(
                        "[%s] F2: producto no hallado en SQL (producto=%r)",
                        msg.id[:8], parsed2.producto,
                    )
                    stats["errors"] += 1
                    _move_to_errors(msg.id, errors_folder_id, dry_run)
                    _notify_control(
                        f"[MailBot] Producto no encontrado - NF {nf}",
                        f"<p>Mail: {msg.subject}</p>"
                        f"<p>Producto del PDF: <b>{parsed2.producto}</b></p>"
                        "<p>No matchea ningún ARTCOD SAFED en STMPDH. "
                        "Resolver manualmente o agregar el producto a la BD.</p>",
                        control_email,
                    )
                    continue
                producto = artcod
        except pdf_parser.PdfParseError as e:
            log.warning("[%s] parse falló: %s", msg.id[:8], e)
            stats["errors"] += 1
            _move_to_errors(msg.id, errors_folder_id, dry_run)
            _notify_control(
                f"[MailBot] Parse PDF falló - {msg.subject}",
                f"<p>De: {msg.from_address}</p><p>{e}</p>",
                control_email,
            )
            continue
        except Exception as e:
            log.exception("[%s] error inesperado parseando: %s", msg.id[:8], e)
            stats["errors"] += 1
            _move_to_errors(msg.id, errors_folder_id, dry_run)
            _notify_control(
                f"[MailBot] Error inesperado - {msg.subject}",
                f"<p>{e}</p>",
                control_email,
            )
            continue

        # --- Idempotencia (solo en producción; en test reprocesa siempre) ---
        # Un mismo Packing List puede aparecer varias veces con productos
        # distintos. Por eso chequeamos PL + Producto.
        if not test_local:
            try:
                if writer.packing_list_producto_exists(nf, producto):
                    log.info(
                        "[%s] NF=%s + producto=%s ya existe en Excel → skip.",
                        msg.id[:8], nf, producto,
                    )
                    stats["skipped_existing"] += 1
                    _move_to_processed(msg.id, processed_folder_id, dry_run)
                    continue
            except Exception as e:
                log.exception("[%s] packing_list_producto_exists falló: %s", msg.id[:8], e)
                stats["errors"] += 1
                _move_to_errors(msg.id, errors_folder_id, dry_run)
                continue

        # --- Lookup LF (puede no estar todavía) ---
        try:
            traza = lookup_trazabilidad_by_nf(nf)
        except Exception as e:
            log.exception("[%s] lookup LF falló: %s", msg.id[:8], e)
            traza = None
        lf_value = _build_lf_value(traza)

        # --- Asignar protocolo ---
        try:
            pr = writer.next_protocolo()
        except Exception as e:
            log.exception("[%s] next_protocolo falló: %s", msg.id[:8], e)
            stats["errors"] += 1
            _move_to_errors(msg.id, errors_folder_id, dry_run)
            continue

        log.info(
            "[%s] NF=%s | PR=%s | producto=%s | fecha=%s | LF=%s",
            msg.id[:8], nf, pr, producto, fecha, lf_value,
        )

        if dry_run:
            log.info("[DRY-RUN] [%s] no se escribe Excel ni se sube PDF ni se mueve mail.", msg.id[:8])
            stats["ok"] += 1
            continue

        new_name = _build_new_pdf_name(pr, nf, att_name, formato)

        # =================================================================
        # MODO TEST LOCAL: guardar PDF en test_pdfs/ + mandar mail con datos.
        # NO se escribe el Excel real, NO se sube a SP.
        # =================================================================
        if test_local:
            try:
                local_path = _test_pdfs_dir() / new_name
                local_path.write_bytes(pdf_bytes)
                log.info("[TEST] [%s] PDF guardado local: %s", msg.id[:8], local_path)
            except Exception as e:
                log.exception("[TEST] [%s] no se pudo guardar PDF local: %s", msg.id[:8], e)
                stats["errors"] += 1
                _move_to_errors(msg.id, errors_folder_id, dry_run)
                continue

            html = _build_test_email_html(
                nf=nf, fecha=fecha, lf=lf_value, producto=producto, pr=pr,
                formato=formato, original_name=att_name, new_name=new_name,
                local_path=local_path,
            )
            destino = test_email or control_email
            if destino:
                try:
                    send_mail(
                        to=[destino],
                        subject=f"[MailBot TEST] {pr} - NF {nf}",
                        html_body=html,
                    )
                except Exception as e:
                    log.warning("[TEST] [%s] no se pudo mandar mail de test: %s", msg.id[:8], e)
            else:
                log.warning("[TEST] No hay MAILBOT_TEST_EMAIL ni CONTROL_EMAIL — no se manda mail.")

            _move_to_processed(msg.id, processed_folder_id, dry_run)
            stats["ok"] += 1
            continue

        # =================================================================
        # PRODUCCIÓN: escribir Excel + subir PDF a SharePoint.
        # =================================================================

        # --- Escribir fila ---
        try:
            writer.append_row({
                "packing_list": nf,
                "fc_fedrigoni": nf,
                "fecha": fecha,
                "lf": lf_value,
                "producto": producto,
                "protocolo": pr,
            })
        except Exception as e:
            log.exception("[%s] append_row falló: %s", msg.id[:8], e)
            stats["errors"] += 1
            _move_to_errors(msg.id, errors_folder_id, dry_run)
            _notify_control(
                f"[MailBot] Error escribiendo Excel - NF {nf}",
                f"<p>{e}</p>",
                control_email,
            )
            continue

        # --- Subir PDF ---
        sp_path = f"{sp_protocols_folder.rstrip('/')}/{new_name}"
        try:
            sp_client.upload_file(sp_path, pdf_bytes, overwrite=False)
            log.info("[%s] PDF subido: %s", msg.id[:8], sp_path)
        except SharePointError as e:
            log.exception("[%s] upload PDF falló: %s", msg.id[:8], e)
            # La fila ya se escribió; loguear y avisar pero seguir adelante
            # (mover a Errores para que se revise manualmente).
            stats["errors"] += 1
            _move_to_errors(msg.id, errors_folder_id, dry_run)
            _notify_control(
                f"[MailBot] PDF NO subido (fila ya escrita) - {pr}",
                f"<p>La fila se escribió pero NO se pudo subir el PDF a SP.</p>"
                f"<p>Path destino: {sp_path}</p><p>{e}</p>",
                control_email,
            )
            continue

        # --- Mover mail a procesados ---
        _move_to_processed(msg.id, processed_folder_id, dry_run)
        stats["ok"] += 1

    return stats


def _move_to_processed(msg_id: str, folder_id: str, dry_run: bool) -> None:
    if dry_run:
        return
    try:
        mail_reader.move_message(msg_id, folder_id)
    except Exception as e:
        log.warning("No se pudo mover mail %s a procesados: %s", msg_id[:8], e)


def _move_to_errors(msg_id: str, folder_id: str, dry_run: bool) -> None:
    if dry_run:
        return
    try:
        mail_reader.move_message(msg_id, folder_id)
    except Exception as e:
        log.warning("No se pudo mover mail %s a errores: %s", msg_id[:8], e)


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------
def run(setup_auth: bool = False, dry_run: bool = False) -> int:
    log.info("=" * 60)
    log.info("MAILBOT arrancando (dry_run=%s, setup_auth=%s)", dry_run, setup_auth)
    log.info("=" * 60)

    s = get_settings()

    # --- Setup auth (login interactivo una vez con AGENTE) ---
    if setup_auth:
        try:
            ensure_authenticated(
                extra_scopes=[MAIL_READWRITE_SCOPE],
                cache_name=mail_reader.AGENTE_CACHE,
            )
            log.info("Auth OK. Cache 'agente' guardado.")
            return EXIT_OK
        except AuthError as e:
            log.error("Setup auth (agente) falló: %s", e)
            return EXIT_AUTH

    # --- Auth silenciosa AGENTE ---
    try:
        get_token(
            interactive=False,
            extra_scopes=[MAIL_READWRITE_SCOPE],
            cache_name=mail_reader.AGENTE_CACHE,
        )
    except AuthError as e:
        log.error(
            "Auth AGENTE silenciosa falló: %s. "
            "Correr con --mailbot --setup-auth para reloggear.",
            e,
        )
        return EXIT_AUTH

    # Log diagnóstico: con qué cuenta quedó autenticado el token de AGENTE.
    try:
        me = mail_reader.whoami()
        upn = me.get("userPrincipalName") or me.get("mail") or "(desconocido)"
        display = me.get("displayName") or ""
        log.info("MailBot autenticado como: %s (%s)", upn, display)
        # Si la cuenta autenticada NO coincide con MAILBOT_INBOX_USER → warning grande.
        expected = (s.mailbot_inbox_user or "").strip().lower()
        actual = str(upn or "").strip().lower()
        if expected and actual and expected != actual:
            log.warning("=" * 60)
            log.warning("⚠ ATENCIÓN: el token está autenticado como '%s'", upn)
            log.warning("⚠ pero MAILBOT_INBOX_USER en .env dice '%s'.", s.mailbot_inbox_user)
            log.warning("⚠ El bot va a leer la Inbox de '%s' (la cuenta logueada).", upn)
            log.warning("⚠ Si querés cambiar de cuenta: borrá %%APPDATA%%\\ProtocolosApp\\token_agente.cache")
            log.warning("⚠ y volvé a correr: MailBot.exe --setup-auth")
            log.warning("=" * 60)
    except Exception as e:
        log.warning("No se pudo obtener whoami: %s", e)

    senders_f1 = _split_csv(s.mailbot_senders_f1)
    senders_f2 = _split_csv(s.mailbot_senders_f2)
    if not senders_f1 and not senders_f2:
        log.error("MAILBOT_SENDERS_F1 y MAILBOT_SENDERS_F2 vacíos en .env. Nada para procesar.")
        return EXIT_FATAL

    # --- Carpetas en mailbox ---
    try:
        processed_folder_id = mail_reader.ensure_mail_folder(s.mailbot_processed_folder)
        errors_folder_id = mail_reader.ensure_mail_folder(s.mailbot_errors_folder)
    except Exception as e:
        log.exception("No se pudieron asegurar las mail folders: %s", e)
        return EXIT_FATAL

    today = date.today()
    overall_stats: dict = {}
    test_local = bool(s.mailbot_test_local)

    if test_local:
        log.warning("=" * 60)
        log.warning("⚠ MODO TEST LOCAL activo")
        log.warning("⚠ NO se escribirá el Excel ni se subirán PDFs a SharePoint")
        log.warning("⚠ Los PDFs renombrados van a `test_pdfs/` junto al .exe")
        log.warning("⚠ Aviso por mail a: %s", s.mailbot_test_email or s.control_email or "(NINGUNO)")
        log.warning("=" * 60)

    # --- Fase A + B con la sesión Excel abierta ---
    try:
        with IngresosExcelWriter() as writer:
            if test_local:
                log.info("Fase A salteada en modo test local.")
                overall_stats["fase_a"] = {"skipped": True}
            else:
                log.info("Excel session abierta. Fase A (re-check LF pendientes)...")
                stats_a = _process_pending_lf(
                    writer, today, s.mailbot_lf_escalate_days, s.control_email,
                )
                log.info("Fase A stats: %s", stats_a)
                overall_stats["fase_a"] = stats_a

            log.info("Fase B (mails nuevos)...")
            stats_b = _process_inbox(
                writer=writer,
                senders_f1=senders_f1,
                senders_f2=senders_f2,
                processed_folder_id=processed_folder_id,
                errors_folder_id=errors_folder_id,
                control_email=s.control_email,
                sp_protocols_folder=s.sp_protocols_folder,
                dry_run=dry_run,
                test_local=test_local,
                test_email=s.mailbot_test_email,
            )
            log.info("Fase B stats: %s", stats_b)
            overall_stats["fase_b"] = stats_b
    except Exception as e:
        log.exception("Error fatal en MAILBOT: %s", e)
        return EXIT_FATAL

    # --- Fase C ---
    if not dry_run:
        try:
            upload_run_log_to_sharepoint()
        except Exception as e:
            log.warning("upload_run_log falló: %s", e)

    log.info("MAILBOT terminó. Stats: %s", overall_stats)
    return EXIT_OK
