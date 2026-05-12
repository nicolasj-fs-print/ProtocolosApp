"""Orquestación de generación de PDF por comprobante."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from ..config import get_settings
from ..pdf.caratula import build_caratula
from ..pdf.detalle import build_detalle
from ..pdf.header import add_header_to_pdf
from ..pdf.merger import merge_pdfs, merge_pdfs_to_bytes
from ..pdf.common import CaratulaMeta
from ..utils.logger import get_logger
from ..utils.paths import sanitize_filename
from .protocol_service import ProtocolFinder

log = get_logger(__name__)


@dataclass
class ComprobanteResult:
    comprobante: str = ""
    cliente: str = ""
    razon_social: str = ""
    numero_oc: str = ""
    mail: str = ""
    estado: str = ""           # "ok" | "ya_existe" | "solo_caratula" | "error" | "omitido_por_usuario"
    pdf_path: str = ""         # webUrl SharePoint o path local
    pdf_local_path: str = ""   # copia local en %TEMP% para adjuntar al mail
    protocolos_detectados: list[str] = field(default_factory=list)
    protocolos_encontrados: list[str] = field(default_factory=list)
    protocolos_no_encontrados: list[str] = field(default_factory=list)
    error: str = ""

    def to_dict(self) -> dict:
        return {
            "Comprobante": self.comprobante,
            "Cliente": self.cliente,
            "Razón Social": self.razon_social,
            "Número OC": self.numero_oc,
            "Mail Protocolos": self.mail,
            "Estado": self.estado,
            "PDF generado": Path(self.pdf_path).name if self.pdf_path else "",
            "Ruta PDF": self.pdf_path,
            "Protocolos detectados": ", ".join(self.protocolos_detectados),
            "Protocolos encontrados": ", ".join(self.protocolos_encontrados),
            "Protocolos no encontrados": ", ".join(self.protocolos_no_encontrados),
            "Error": self.error,
        }


def _output_filename(comprobante: str, cliente: str, numero_oc: str) -> str:
    base = f"{sanitize_filename(comprobante)}_{sanitize_filename(cliente)}_{sanitize_filename(numero_oc or 'SIN_OC')}"
    return f"{base}.pdf"


def _unique_protocols(rows: pd.DataFrame) -> list[str]:
    seen: set[str] = set()
    ordered: list[str] = []
    for v in rows["#Protocolo PDF"]:
        if not isinstance(v, str):
            continue
        p = v.strip()
        if not p or p in seen:
            continue
        seen.add(p)
        ordered.append(p)
    return ordered


def _output_already_exists(pdf_name: str, output_dir, settings) -> tuple[bool, str]:
    """Devuelve (existe, path_str) para el PDF en disco o SharePoint."""
    if settings.storage_backend == "sharepoint":
        from ..sharepoint.client import get_client
        sp_path = f"{output_dir.rstrip('/')}/{pdf_name}"
        try:
            return get_client().file_exists(sp_path), sp_path
        except Exception as e:
            log.warning("file_exists SP falló (%s) → asumimos que no existe", e)
            return False, sp_path
    pdf_path = (output_dir if isinstance(output_dir, Path) else Path(output_dir)) / pdf_name
    return pdf_path.exists(), str(pdf_path)


def generate_for_comprobante(
    rows_df: pd.DataFrame,
    output_dir,
    finder: ProtocolFinder,
) -> ComprobanteResult:
    """`output_dir` es Path local o string SharePoint según `storage_backend`."""
    log.info("generate_for_comprobante: ENTRADA (df=%d filas)", len(rows_df))
    settings = get_settings()
    result = ComprobanteResult()

    if rows_df.empty:
        result.estado = "error"
        result.error = "DataFrame vacío"
        return result

    first = rows_df.iloc[0]
    comprobante = str(first["#Comprobante"]).strip()
    cliente = str(first["Código de Cliente"]).strip()
    razon_social = str(first.get("Razón Social", "")).strip()
    numero_oc = str(first.get("Número OC", "") or "").strip()
    mail = str(first.get("Mail Protocolos", "") or "").strip()

    result.comprobante = comprobante
    result.cliente = cliente
    result.razon_social = razon_social
    result.numero_oc = numero_oc
    result.mail = mail

    pdf_name = _output_filename(comprobante, cliente, numero_oc)
    log.info("Verificando si ya existe: %s", pdf_name)
    exists, pdf_path_str = _output_already_exists(pdf_name, output_dir, settings)
    log.info("Verificación existencia: existe=%s", exists)
    result.pdf_path = pdf_path_str

    if exists:
        log.info("PDF ya existe, se omite: %s", pdf_name)
        result.estado = "ya_existe"
        return result

    meta = CaratulaMeta(
        cliente=cliente,
        razon_social=razon_social,
        comprobante=comprobante,
        numero_oc=numero_oc,
        fecha=first.get("Fecha"),
    )
    rows_dicts = rows_df.to_dict(orient="records")
    log.info("Generando carátula PDF (%d filas)...", len(rows_dicts))
    caratula_bytes = build_caratula(meta, rows_dicts)
    log.info("Carátula generada (%.1f KB)", len(caratula_bytes) / 1024)

    protocolos = _unique_protocols(rows_df)
    log.info("Protocolos únicos detectados: %d → %s", len(protocolos), protocolos)
    result.protocolos_detectados = protocolos
    parts: list[bytes] = [caratula_bytes]

    estado_final = "solo_caratula" if not protocolos else "ok"

    if protocolos:
        # Si hay 1 solo protocolo, no incluyo el "detalle" (la info ya está
        # completa en la carátula → no tiene sentido duplicar).
        single_proto = len(protocolos) == 1

        for proto in protocolos:
            proto_rows = rows_df[rows_df["#Protocolo PDF"] == proto].to_dict(orient="records")

            if not single_proto:
                try:
                    detalle_bytes = build_detalle(meta, proto_rows, proto)
                    parts.append(detalle_bytes)
                except Exception as e:
                    log.exception("Error generando detalle de %s: %s", proto, e)
                    continue

            original = finder.read_bytes(proto)
            if original is None:
                log.warning("Protocolo %s NO encontrado", proto)
                result.protocolos_no_encontrados.append(proto)
                continue

            try:
                with_header = add_header_to_pdf(original, proto)
                parts.append(with_header)
                result.protocolos_encontrados.append(proto)
            except Exception as e:
                log.exception("Error agregando header a %s: %s", proto, e)
                result.protocolos_no_encontrados.append(proto)

    # Persistir
    if settings.storage_backend == "sharepoint":
        from ..sharepoint.client import get_client
        merged_bytes = merge_pdfs_to_bytes(parts)
        sp_path = f"{str(output_dir).rstrip('/')}/{pdf_name}"
        try:
            uploaded = get_client().upload_file(sp_path, merged_bytes, overwrite=False)
            result.pdf_path = uploaded.get("webUrl") or sp_path
            log.info("PDF subido a SharePoint: %s", sp_path)
        except Exception as e:
            log.exception("Falló upload a SharePoint para %s", comprobante)
            result.estado = "error"
            result.error = f"Upload SP: {e}"
            return result

        # Copia local en %TEMP% para poder adjuntarla al mail después.
        try:
            import os
            import tempfile
            temp_dir = Path(tempfile.gettempdir()) / "ProtocolosApp"
            temp_dir.mkdir(parents=True, exist_ok=True)
            temp_pdf = temp_dir / pdf_name
            temp_pdf.write_bytes(merged_bytes)
            result.pdf_local_path = str(temp_pdf)
        except Exception as e:
            log.warning("No se pudo guardar copia local temp: %s", e)
    else:
        out_dir = output_dir if isinstance(output_dir, Path) else Path(output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        pdf_path = out_dir / pdf_name
        merge_pdfs(parts, pdf_path)
        result.pdf_path = str(pdf_path)
        result.pdf_local_path = str(pdf_path)
        log.info("PDF generado local: %s", pdf_path.name)

    result.estado = estado_final
    return result
