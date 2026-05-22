"""Orquestador del modo automático (cron diario).

Flujo:
1. Auth MSAL silent (si falla → exit 2).
2. Buscar remitos del rango [today - AUTO_DAYS_BACK, today - 1].
3. Filtrar los que ya están en el tracking con estado terminal.
4. Por cada remito:
   - Si tiene items SAFED sin protocolo → mail a CONTROL_EMAIL + tracking pending_control.
   - Si limpio → generar PDF + mail al cliente + tracking enviado_cliente.
   - Si error → log, NO trackea (reintento al día siguiente).
5. Resumen .xlsx + log → SharePoint.
"""
from __future__ import annotations

import io
import json
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

from ..config import get_settings
from ..services import data_service
from ..services.mail_service import build_control_html_body, build_html_body
from ..services.pdf_service import ComprobanteResult, generate_for_comprobante
from ..services.protocol_service import ProtocolFinder
from ..sharepoint.auth import AuthError, get_token
from ..sharepoint.client import SharePointError, get_client
from ..sharepoint.lists import (
    ESTADO_ENVIADO_CLIENTE,
    ESTADO_ENVIADO_MANUAL,
    ESTADO_PENDING_CONTROL,
    filter_unprocessed,
    get_lists,
)
from ..utils.logger import (
    current_run_id,
    get_logger,
    logs_dir,
    upload_run_log_to_sharepoint,
)
from .mailer_graph import send_mail

log = get_logger(__name__)

# Exit codes
EXIT_OK = 0
EXIT_FATAL = 1
EXIT_AUTH = 2


def _read_pdf_local(result: ComprobanteResult) -> bytes | None:
    """Lee el PDF generado para adjuntarlo al mail.

    `pdf_local_path` es una copia local en %TEMP% que dejó pdf_service tras
    subir a SharePoint. Si no está, hace fallback al webUrl.
    """
    if result.pdf_local_path:
        p = Path(result.pdf_local_path)
        if p.exists():
            try:
                return p.read_bytes()
            except Exception as e:
                log.warning("No se pudo leer pdf_local_path %s: %s", p, e)
    # Fallback: descargar del SP. result.pdf_path puede ser webUrl o path.
    # Solo intentamos si parece path SharePoint (no URL https).
    if result.pdf_path and not result.pdf_path.startswith("http"):
        try:
            return get_client().download_file(result.pdf_path)
        except Exception as e:
            log.warning("Fallback download de SP falló: %s", e)
    return None


def _subject_cliente(razon_social: str, comprobante: str) -> str:
    return f"Protocolos de calidad - {razon_social} | Remito {comprobante}"


def _subject_control(comprobante: str) -> str:
    return f"[Protocolos] Falta info — Remito {comprobante}"


def _pdf_filename(result: ComprobanteResult) -> str:
    if result.pdf_path:
        name = Path(result.pdf_path).name
        if name:
            return name
    return f"{result.comprobante}.pdf"


def _process_comprobante(
    comp: str,
    sub_df: pd.DataFrame,
    output_dir,
    finder: ProtocolFinder,
    *,
    list_name: str,
    control_email: str,
    test_override_email: str,
    run_id: str,
    dry_run: bool,
    missing_per_comp: dict,
) -> dict:
    """Procesa UN remito. Devuelve dict con resumen del resultado para el .xlsx."""
    settings = get_settings()
    lists = get_lists()
    first = sub_df.iloc[0]
    cliente = str(first.get("Código de Cliente", "")).strip()
    razon_social = str(first.get("Razón Social", "")).strip()

    # ╔══════════════════════════════════════════════════════════════════════╗
    # ║  DESTINATARIO DEL CLIENTE                                            ║
    # ║                                                                      ║
    # ║  En PRODUCCIÓN: usa el `Mail Protocolos` real del Excel.            ║
    # ║  En PRUEBA: si TEST_MODE_OVERRIDE_EMAIL está set en .env, redirige  ║
    # ║  todos los mails a esa dirección (clientes reales NO reciben nada). ║
    # ╚══════════════════════════════════════════════════════════════════════╝
    mail_cliente_real = str(first.get("Mail Protocolos", "")).strip()
    if test_override_email:
        mail_cliente = test_override_email
        log.warning(
            "[%s] TEST MODE: mail cliente redirigido '%s' → '%s'",
            comp, mail_cliente_real or "(sin mail)", test_override_email,
        )
    else:
        mail_cliente = mail_cliente_real

    fecha = first.get("Fecha")
    fecha_str = ""
    try:
        fecha_str = pd.to_datetime(fecha).date().isoformat()
    except Exception:
        pass

    missing_items = missing_per_comp.get(comp, {}).get("items") if missing_per_comp.get(comp) else None

    # =========================================================
    # CASO 1: Hay items SAFED sin protocolo → control
    # =========================================================
    if missing_items:
        # En TEST MODE, también se redirige el aviso de control al override.
        effective_control = test_override_email or control_email
        log.info("[%s] tiene %d item(s) sin protocolo → control (%s)",
                 comp, len(missing_items), effective_control or "(sin destino)")
        if not effective_control:
            log.error("[%s] no se puede avisar a control: CONTROL_EMAIL vacío en .env", comp)
            return {
                "Comprobante": comp, "Cliente": cliente, "RazonSocial": razon_social,
                "Estado": "error", "Detalle": "CONTROL_EMAIL no configurado",
            }
        html = build_control_html_body(comp, cliente, razon_social, missing_items)
        if dry_run:
            log.info("[DRY-RUN] [%s] mail control → %s", comp, effective_control)
        else:
            try:
                send_mail(
                    to=[effective_control],
                    subject=_subject_control(comp),
                    html_body=html,
                )
            except SharePointError as e:
                log.exception("[%s] send_mail control falló: %s", comp, e)
                return {
                    "Comprobante": comp, "Cliente": cliente, "RazonSocial": razon_social,
                    "Estado": "error", "Detalle": f"send_mail control: {e}",
                }
            try:
                lists.upsert(
                    list_name,
                    comprobante=comp,
                    cliente=cliente,
                    razon_social=razon_social,
                    fecha_remito=fecha_str,
                    estado=ESTADO_PENDING_CONTROL,
                    mail_destino=effective_control,
                    items_faltantes=json.dumps(
                        [{"producto": i.get("producto"), "serie": i.get("serie"), "m2": i.get("m2")}
                         for i in missing_items], ensure_ascii=False,
                    ),
                    run_id=run_id,
                )
            except SharePointError as e:
                log.warning("[%s] upsert tracking falló (control): %s", comp, e)
        return {
            "Comprobante": comp, "Cliente": cliente, "RazonSocial": razon_social,
            "Estado": ESTADO_PENDING_CONTROL, "Detalle": f"{len(missing_items)} item(s) sin protocolo",
        }

    # =========================================================
    # CASO 2: Remito limpio → generar PDF + mandar al cliente
    # =========================================================
    if not mail_cliente:
        log.warning("[%s] sin Mail Protocolos → no se puede enviar al cliente", comp)
        return {
            "Comprobante": comp, "Cliente": cliente, "RazonSocial": razon_social,
            "Estado": "error", "Detalle": "Cliente sin Mail Protocolos",
        }

    # En dry-run, NO generamos el PDF (no sube nada a SharePoint) ni mandamos
    # mail ni escribimos tracking. Solo simulamos.
    if dry_run:
        log.info("[DRY-RUN] [%s] generaría PDF + mail al cliente → %s",
                 comp, mail_cliente)
        return {
            "Comprobante": comp, "Cliente": cliente, "RazonSocial": razon_social,
            "Estado": "dry_run", "Detalle": f"sería enviado a {mail_cliente}",
        }

    try:
        result = generate_for_comprobante(sub_df.reset_index(drop=True), output_dir, finder)
    except Exception as e:
        log.exception("[%s] generate_for_comprobante falló: %s", comp, e)
        return {
            "Comprobante": comp, "Cliente": cliente, "RazonSocial": razon_social,
            "Estado": "error", "Detalle": f"generate: {e}",
        }

    if result.estado == "error":
        return {
            "Comprobante": comp, "Cliente": cliente, "RazonSocial": razon_social,
            "Estado": "error", "Detalle": result.error,
        }

    # ╔══════════════════════════════════════════════════════════════════════╗
    # ║  PDF ya existía en SharePoint → asumimos envío manual desde la GUI. ║
    # ║  NO re-enviamos mail (evita duplicados al cliente). Solo registramos║
    # ║  en tracking como `enviado_manual` para que futuras corridas tampoco║
    # ║  lo toquen.                                                          ║
    # ╚══════════════════════════════════════════════════════════════════════╝
    if result.estado == "ya_existe":
        log.info("[%s] PDF ya existía en SP → asumimos envío manual, NO se re-manda mail", comp)
        try:
            lists.upsert(
                list_name,
                comprobante=comp,
                cliente=cliente,
                razon_social=razon_social,
                fecha_remito=fecha_str,
                estado=ESTADO_ENVIADO_MANUAL,
                mail_destino="(manual desde GUI)",
                pdf_url=result.pdf_path,
                run_id=run_id,
            )
        except SharePointError as e:
            log.warning("[%s] upsert tracking falló (manual): %s", comp, e)
        return {
            "Comprobante": comp, "Cliente": cliente, "RazonSocial": razon_social,
            "Estado": ESTADO_ENVIADO_MANUAL,
            "Detalle": "PDF ya existía en SP — no se re-envió",
        }

    # PDF generado por primera vez. Mandar mail al cliente.
    pdf_bytes = _read_pdf_local(result)
    if not pdf_bytes:
        log.error("[%s] no se pudo recuperar el PDF para adjuntar", comp)
        return {
            "Comprobante": comp, "Cliente": cliente, "RazonSocial": razon_social,
            "Estado": "error", "Detalle": "No se pudo leer el PDF generado",
        }

    pdf_name = _pdf_filename(result)
    subject = _subject_cliente(razon_social, comp)
    html = build_html_body(
        razon_social, comp,
        result.protocolos_encontrados,
        result.protocolos_no_encontrados,
    )

    # CC fijo (settings.protocols_cc) — solo en PRODUCCIÓN. Si estamos en
    # test mode con override, ignoramos el CC para no contaminar la prueba.
    cc_list: list[str] = []
    if not test_override_email and settings.protocols_cc:
        cc_list = [settings.protocols_cc]

    try:
        send_mail(
            to=[mail_cliente],
            subject=subject,
            html_body=html,
            attachments=[(pdf_name, pdf_bytes)],
            cc=cc_list,
        )
    except SharePointError as e:
        log.exception("[%s] send_mail cliente falló: %s", comp, e)
        return {
            "Comprobante": comp, "Cliente": cliente, "RazonSocial": razon_social,
            "Estado": "error", "Detalle": f"send_mail cliente: {e}",
        }
    try:
        lists.upsert(
            list_name,
            comprobante=comp,
            cliente=cliente,
            razon_social=razon_social,
            fecha_remito=fecha_str,
            estado=ESTADO_ENVIADO_CLIENTE,
            mail_destino=mail_cliente,
            pdf_url=result.pdf_path,
            run_id=run_id,
        )
    except SharePointError as e:
        log.warning("[%s] upsert tracking falló (cliente): %s", comp, e)

    return {
        "Comprobante": comp, "Cliente": cliente, "RazonSocial": razon_social,
        "Estado": ESTADO_ENVIADO_CLIENTE, "Detalle": f"{len(result.protocolos_encontrados)} protocolo(s) OK",
    }


def _write_summary(rows: list[dict]) -> None:
    """Sube resumen .xlsx a SP_EJECUCIONES_FOLDER (best-effort)."""
    if not rows:
        return
    s = get_settings()
    ts = current_run_id()
    filename = f"resumen_auto_{ts}.xlsx"
    try:
        buf = io.BytesIO()
        pd.DataFrame(rows).to_excel(buf, index=False, engine="openpyxl")
        buf.seek(0)
        if s.storage_backend == "sharepoint" and s.sp_ejecuciones_folder:
            from ..sharepoint.client import get_app_client
            sp_path = f"{s.sp_ejecuciones_folder.rstrip('/')}/{filename}"
            get_app_client().upload_file(sp_path, buf.getvalue(), overwrite=True)
            log.info("Resumen subido: %s", sp_path)
        else:
            fallback = logs_dir().parent / filename
            pd.DataFrame(rows).to_excel(fallback, index=False, engine="openpyxl")
            log.info("Resumen guardado local: %s", fallback)
    except Exception as e:
        log.exception("No se pudo escribir resumen: %s", e)


def run(setup_auth: bool = False, dry_run: bool = False) -> int:
    """Entrypoint del modo --auto. Devuelve exit code."""
    log.info("=" * 60)
    log.info("Modo AUTOMÁTICO arrancando (dry_run=%s, setup_auth=%s)", dry_run, setup_auth)
    log.info("=" * 60)

    s = get_settings()

    # --- Setup auth (login interactivo una vez) ---
    if setup_auth:
        try:
            from ..sharepoint.auth import ensure_authenticated
            ensure_authenticated()
            log.info("Auth OK. Cache guardado, próximas corridas pueden ser silenciosas.")
            return EXIT_OK
        except AuthError as e:
            log.error("Setup auth falló: %s", e)
            return EXIT_AUTH

    # --- Auth silenciosa ---
    try:
        get_token(interactive=False)
    except AuthError as e:
        log.error("Auth silenciosa falló: %s. Correr con --setup-auth para reloggear.", e)
        return EXIT_AUTH

    # --- Rango de fechas ---
    # date_to = HOY (inclusive) → captura remitos cargados en el día corriente.
    # date_from = HOY - AUTO_DAYS_BACK. Si AUTO_DAYS_BACK=0 → solo procesa HOY.
    today = date.today()
    date_to = today
    date_from = today - timedelta(days=max(0, s.auto_days_back))
    log.info("Rango a procesar: %s → %s", date_from, date_to)

    # --- Fetch SQL ---
    try:
        # Misma política que la GUI: prioriza LF > IRP y descarta IR
        # en el lookup de trazabilidad.
        df = data_service.fetch_data(date_from, date_to, prioritize_lf=True)
    except Exception as e:
        log.exception("fetch_data falló: %s", e)
        return EXIT_FATAL

    if df is None or df.empty:
        log.info("Sin remitos en el rango → exit OK.")
        return EXIT_OK
    log.info("fetch_data devolvió %d filas / %d remitos únicos",
             len(df), df["#Comprobante"].nunique())

    # --- Filtrar contra tracking ---
    list_name = s.sp_tracking_list_name
    try:
        lists = get_lists()
        tracking = lists.fetch_in_range(list_name, date_from, date_to)
    except SharePointError as e:
        log.exception("No se pudo leer tracking: %s", e)
        # Sin tracking arriesgaríamos duplicar. Mejor abortar.
        return EXIT_FATAL

    comprobantes_all = list(df["#Comprobante"].astype(str).unique())
    pendientes = filter_unprocessed(comprobantes_all, tracking)
    n_saltados = len(comprobantes_all) - len(pendientes)
    log.info("Pendientes: %d (saltados por tracking previo: %d)", len(pendientes), n_saltados)
    if not pendientes:
        log.info("Todos los remitos ya están procesados → exit OK.")
        return EXIT_OK

    df = df[df["#Comprobante"].astype(str).isin(pendientes)].copy()

    # --- Pre-análisis: missing por comprobante ---
    missing_per_comp = data_service.analyze_missing_protocols(df)

    # --- Procesar cada remito ---
    finder = ProtocolFinder(s.protocols_folder if s.storage_backend == "local" else None)
    if s.storage_backend == "sharepoint":
        finder.index()

    out_dir = s.sp_output_folder if s.storage_backend == "sharepoint" else s.default_output_folder
    run_id = current_run_id()
    summary_rows: list[dict] = []

    test_override = (s.test_mode_override_email or "").strip()
    if test_override:
        log.warning("=" * 60)
        log.warning("⚠ TEST MODE activo: TODOS los mails van a '%s'", test_override)
        log.warning("⚠ Para producción: vaciar TEST_MODE_OVERRIDE_EMAIL en .env")
        log.warning("=" * 60)

    for comp, sub in df.groupby("#Comprobante", sort=False):
        try:
            row = _process_comprobante(
                str(comp), sub, out_dir, finder,
                list_name=list_name,
                control_email=s.control_email,
                test_override_email=test_override,
                run_id=run_id,
                dry_run=dry_run,
                missing_per_comp=missing_per_comp,
            )
            summary_rows.append(row)
        except Exception as e:
            log.exception("Error fatal procesando %s: %s", comp, e)
            summary_rows.append({
                "Comprobante": str(comp), "Cliente": "", "RazonSocial": "",
                "Estado": "error", "Detalle": f"unhandled: {e}",
            })

    # --- Resumen + log ---
    if not dry_run:
        _write_summary(summary_rows)
        try:
            upload_run_log_to_sharepoint()
        except Exception as e:
            log.warning("upload_run_log falló: %s", e)

    log.info("Modo AUTOMÁTICO terminó. Procesados: %d remito(s).", len(summary_rows))
    return EXIT_OK
