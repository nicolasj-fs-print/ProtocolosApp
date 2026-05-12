"""Thread worker que itera comprobantes y genera PDFs sin bloquear la GUI."""
from __future__ import annotations

import io
import os
import queue
import threading
import traceback
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from ..config import get_settings
from ..services.pdf_service import ComprobanteResult, generate_for_comprobante
from ..services.protocol_service import ProtocolFinder
from ..utils.logger import current_run_id, get_logger, logs_dir, upload_run_log_to_sharepoint

log = get_logger(__name__)


@dataclass
class ProgressEvent:
    """Evento que se mete en la queue. La GUI los consume con `after()`."""
    kind: str  # "log" | "progress" | "result" | "done" | "error"
    message: str = ""
    current: int = 0
    total: int = 0
    payload: object | None = None


class GenerationWorker(threading.Thread):
    def __init__(
        self,
        df: pd.DataFrame,
        output_dir,                        # Path local | str (path SharePoint)
        protocols_folder: Path | None,
        events: "queue.Queue[ProgressEvent]",
        cancel_event: threading.Event,
        pre_results: list[ComprobanteResult] | None = None,
    ):
        super().__init__(daemon=True)
        self.df = df
        self.output_dir = output_dir
        self.events = events
        self.cancel_event = cancel_event
        self.finder = ProtocolFinder(protocols_folder)
        self.results: list[ComprobanteResult] = list(pre_results or [])
        self.settings = get_settings()

    def _emit(self, ev: ProgressEvent) -> None:
        self.events.put(ev)

    def _run_post_done_safe(self) -> None:
        """Resumen.xlsx + log + telemetría usuario en background tras done."""
        try:
            self._write_summary()
        except Exception as e:
            log.warning("write_summary falló: %s", e)

        if self.settings.storage_backend != "sharepoint":
            return

        try:
            upload_run_log_to_sharepoint()
        except Exception as e:
            log.warning("upload_run_log falló: %s", e)
        try:
            from ..sharepoint.users import current_user, register_login, register_run
            if current_user() is None:
                register_login()
            register_run(executions=1)
        except Exception as e:
            log.warning("Telemetría de usuario falló: %s", e)

    def _ensure_output_dir(self) -> None:
        if self.settings.storage_backend == "local":
            out = self.output_dir if isinstance(self.output_dir, Path) else Path(self.output_dir)
            out.mkdir(parents=True, exist_ok=True)

    def run(self) -> None:
        log.info("Worker.run() arrancó")
        try:
            self._ensure_output_dir()
            log.info("Worker: indexando protocolos en SharePoint...")
            self.finder.index()
            log.info("Worker: index OK")

            n_pre = len(self.results)
            if n_pre:
                self._emit(ProgressEvent(
                    "log",
                    f"{n_pre} comprobante(s) omitido(s) por el usuario (registrados en el resumen).",
                ))

            comprobantes = list(self.df.groupby("#Comprobante", dropna=False, sort=False))
            total = len(comprobantes)
            self._emit(ProgressEvent("log", f"Procesando {total} comprobante(s)..."))
            self._emit(ProgressEvent("progress", current=0, total=total))

            for i, (comp, sub) in enumerate(comprobantes, 1):
                if self.cancel_event.is_set():
                    self._emit(ProgressEvent("log", "Cancelado por el usuario."))
                    break

                self._emit(ProgressEvent("log", f"[{i}/{total}] Comprobante {comp}..."))
                log.info("Worker: pre-call generate_for_comprobante (%s)", comp)
                try:
                    result = generate_for_comprobante(
                        sub.reset_index(drop=True),
                        self.output_dir,
                        self.finder,
                    )
                    log.info("Worker: post-call generate_for_comprobante (%s) → %s",
                             comp, result.estado)
                except Exception as e:
                    log.exception("Error procesando comprobante %s", comp)
                    result = ComprobanteResult(
                        comprobante=str(comp),
                        estado="error",
                        error=f"{type(e).__name__}: {e}",
                    )

                self.results.append(result)
                self._emit(ProgressEvent(
                    "result",
                    message=f"  → {result.estado} ({len(result.protocolos_encontrados)} protocolos OK, "
                            f"{len(result.protocolos_no_encontrados)} faltantes)",
                    payload=result,
                ))
                self._emit(ProgressEvent("progress", current=i, total=total))

            # Emitir done INMEDIATAMENTE → la GUI muestra los dialogs
            # (ya_existe / Outlook) sin esperar a uploads ni telemetría.
            self._emit(ProgressEvent(
                "done",
                message=f"Listo. {len(self.results)} comprobantes procesados.",
                payload=self.results,
            ))

            # Resumen.xlsx + log + telemetría en thread daemon — si el user
            # cierra la app antes que termine, se pierden estos uploads (no
            # críticos). Las carpetas ya se crearon en el splash al inicio.
            threading.Thread(
                target=self._run_post_done_safe,
                daemon=True,
                name="ProtocolosApp-Telemetry",
            ).start()
        except Exception as e:
            log.exception("Error fatal en worker")
            self._emit(ProgressEvent(
                "error",
                message=f"Error fatal: {e}\n{traceback.format_exc()}",
            ))

    def _write_summary(self) -> str | None:
        if not self.results:
            return None
        rows = [r.to_dict() for r in self.results]
        ts = current_run_id()
        filename = f"resumen_ejecucion_{ts}.xlsx"

        if self.settings.storage_backend == "sharepoint":
            try:
                buf = io.BytesIO()
                pd.DataFrame(rows).to_excel(buf, index=False, engine="openpyxl")
                buf.seek(0)
                from ..sharepoint.client import get_app_client
                # Los resúmenes viven en el sitio Desarrollo / ProtocolosApp/Ejecuciones,
                # NO en la carpeta de PDFs (Reportes Finales del sitio VentasPowerBI).
                target_folder = (
                    self.settings.sp_ejecuciones_folder.rstrip("/")
                    if self.settings.sp_ejecuciones_folder
                    else str(self.output_dir).rstrip("/")
                )
                sp_path = f"{target_folder}/{filename}"
                get_app_client().upload_file(sp_path, buf.getvalue(), overwrite=True)
                self._emit(ProgressEvent("log", f"Resumen subido a SharePoint: {filename}"))
                return sp_path
            except Exception as e:
                log.exception("No se pudo subir resumen a SharePoint")
                # Fallback: guardarlo local en %APPDATA% para que no se pierda
                try:
                    fallback = logs_dir().parent / filename
                    pd.DataFrame(rows).to_excel(fallback, index=False, engine="openpyxl")
                    self._emit(ProgressEvent("log", f"Resumen guardado local (fallback): {fallback}"))
                    return str(fallback)
                except Exception as e2:
                    self._emit(ProgressEvent("log", f"No se pudo escribir resumen: {e2}"))
                    return None

        # Backend local
        out_dir = self.output_dir if isinstance(self.output_dir, Path) else Path(self.output_dir)
        out = out_dir / filename
        try:
            pd.DataFrame(rows).to_excel(out, index=False, engine="openpyxl")
            self._emit(ProgressEvent("log", f"Resumen escrito: {out.name}"))
            return str(out)
        except Exception as e:
            log.exception("No se pudo escribir resumen")
            self._emit(ProgressEvent("log", f"No se pudo escribir resumen: {e}"))
            return None
