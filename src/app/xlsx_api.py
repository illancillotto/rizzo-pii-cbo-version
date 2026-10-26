"""Local workbook API. Capability IDs + owner cookies, bounded volatile storage."""
from __future__ import annotations

import hmac
import io
import json
import os
import secrets
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from urllib.parse import urlsplit

from flask import Blueprint, Response, jsonify, request, send_file
from werkzeug.exceptions import BadRequest, RequestEntityTooLarge

from xlsx_io import Limits, WorkbookJob, XlsxError, restore, xml_bytes

MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
COOKIE = "rizzo_xlsx_owner"


class Store:
    def __init__(self, maximum=6, memory=512 * 1024 * 1024, ttl=900, now=time.monotonic):
        self.maximum, self.memory, self.ttl, self.now = maximum, memory, ttl, now
        self.entries = {}
        self.lock = threading.RLock()

    def purge(self):
        with self.lock:
            for identifier, entry in list(self.entries.items()):
                if self.now() - entry["created"] > self.ttl:
                    entry["cancelled"] = True
                    if entry["state"] not in {"analysing", "exporting"}:
                        self.entries.pop(identifier, None)

    def add(self, job, owner):
        self.purge()
        roots = job.package.roots.values()
        # Reserve source + working copy + verified output + dictionary, including
        # XML object overhead. Admission is bounded before any background job.
        cost = max(1024 * 1024, sum(len(xml_bytes(r)) * 12 + sum(1 for _ in r.iter()) * 256 for r in roots) * 3)
        with self.lock:
            if len(self.entries) >= self.maximum or sum(e["cost"] for e in self.entries.values()) + cost > self.memory:
                raise XlsxError("Archivio temporaneo pieno: eliminare documenti o attendere la scadenza.")
            identifier = secrets.token_urlsafe(32)
            entry = {"id": identifier, "job": job, "owner": owner, "created": self.now(), "cost": cost,
                     "state": "uploaded", "error": None, "cancelled": False}
            self.entries[identifier] = entry
            return identifier, entry

    def get(self, identifier, owner):
        self.purge()
        with self.lock:
            entry = self.entries.get(identifier)
            if entry is None or entry["cancelled"] or not owner or not hmac.compare_digest(entry["owner"], owner):
                return None
            return entry

    def delete(self, identifier, owner):
        with self.lock:
            entry = self.get(identifier, owner)
            if entry is not None:
                entry["cancelled"] = True
                if entry["state"] not in {"analysing", "exporting"}:
                    self.entries.pop(identifier, None)
                return True
            return False


def register_xlsx(app, recognize_batch, preferences=lambda: ([], True), store=None, start_sweeper=True):
    limits = Limits.environment()
    store = store or Store(maximum=int(os.getenv("PII_XLSX_SESSIONS", "6")),
                           memory=int(os.getenv("PII_XLSX_MEMORY", str(512 * 1024 * 1024))),
                           ttl=int(os.getenv("PII_XLSX_TTL", "900")))
    if min(store.maximum, store.memory, store.ttl) <= 0:
        raise XlsxError("Limiti delle sessioni XLSX non validi.")
    bp = Blueprint("xlsx", __name__)
    pool = ThreadPoolExecutor(max_workers=1, thread_name_prefix="xlsx-local")
    ingest = threading.BoundedSemaphore(1)
    bp.store = store
    bp.executor = pool

    class MemoryUpload(io.BytesIO):
        def write(self, data):
            if self.tell() + len(data) > max(limits.compressed, limits.characters * 8 + 1024):
                raise RequestEntityTooLarge()
            return super().write(data)

    class LocalXlsxRequest(app.request_class):
        def _get_file_stream(self, total_content_length, content_type, filename=None, content_length=None):
            # Werkzeug otherwise spools uploads >500 KiB to a temporary disk file.
            # Keep existing PDF/text request behavior unchanged.
            if self.path.startswith("/api/xlsx/"):
                return MemoryUpload()
            return super()._get_file_stream(total_content_length, content_type, filename, content_length)

    app.request_class = LocalXlsxRequest

    if start_sweeper:
        def sweep():
            while True:
                threading.Event().wait(min(30, store.ttl))
                store.purge()
        threading.Thread(target=sweep, daemon=True, name="xlsx-expiry").start()

    @bp.before_request
    def protect():
        store.purge()
        if request.path.startswith("/api/xlsx"):
            host = urlsplit("//" + request.host).hostname
            trusted = {"localhost", "127.0.0.1", "::1"} | set(
                os.getenv("PII_XLSX_TRUSTED_HOSTS", "").split(","))
            if host not in trusted:
                return jsonify(error="Host locale non consentito per le API XLSX."), 403
            origin = request.headers.get("Origin")
            if origin and (urlsplit(origin).netloc != request.host or urlsplit(origin).scheme != request.scheme):
                return jsonify(error="Origine della richiesta non consentita."), 403
            if request.headers.get("Sec-Fetch-Site") in {"cross-site", "same-site"}:
                return jsonify(error="Richiesta da un altro sito non consentita."), 403
            if request.method in {"POST", "PATCH", "DELETE"} and request.headers.get("X-Rizzo-XLSX") != "1":
                return jsonify(error="Header di protezione XLSX mancante."), 403
            if (request.method in {"POST", "PATCH"} and not request.path.endswith(("/upload", "/restore"))
                    and request.content_length is not None and request.content_length > 1024 * 1024):
                return jsonify(error="Configurazione o revisione XLSX troppo grande."), 413

    @bp.after_request
    def private(response):
        response.headers["Cache-Control"] = "no-store, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["Referrer-Policy"] = "no-referrer"
        response.headers["Cross-Origin-Resource-Policy"] = "same-origin"
        return response

    @bp.errorhandler(XlsxError)
    def user_error(error):
        return jsonify(error=str(error)), 422

    @bp.errorhandler(Exception)
    def internal_error(error):
        # Some XML/ZIP/reader errors may contain cell text. Do not let Flask
        # serialize/log their exception messages or stack traces.
        from werkzeug.exceptions import HTTPException
        code = error.code if isinstance(error, HTTPException) else 500
        return jsonify(error="Richiesta o elaborazione XLSX non riuscita."), code

    @bp.errorhandler(BadRequest)
    @bp.errorhandler(RequestEntityTooLarge)
    def bad_request(error):
        return jsonify(error="Richiesta XLSX non valida o troppo grande."), error.code

    def entry_for(identifier, idle=False):
        entry = store.get(identifier, request.cookies.get(COOKIE))
        if entry is None:
            return None, (jsonify(error="Documento assente, scaduto o appartenente a un'altra sessione."), 404)
        if idle and entry["state"] in {"analysing", "exporting"}:
            return None, (jsonify(error="Elaborazione già in corso."), 409)
        return entry, None

    def detector(entry):
        def budget():
            if entry["cancelled"] or store.now() - entry["created"] > store.ttl:
                raise XlsxError("Documento eliminato o scaduto.")
            entry["job"].package.check_time()
        def detect(texts):
            budget()
            result = recognize_batch(texts, budget)
            budget()
            return result
        return detect

    def start(entry, operation, config=None):
        entry.update(state="analysing" if operation == "analyze" else "exporting", error=None)
        def run():
            try:
                if operation == "analyze":
                    entry["job"].analyze(detector(entry), config)
                else:
                    entry["job"].export(detector(entry))
                with store.lock:
                    if not entry["cancelled"]:
                        entry["state"] = "review" if operation == "analyze" else "ready"
            except Exception as error:
                # Never log exceptions with library messages that may contain
                # user text. Only our controlled, content-free messages leave here.
                message = str(error) if isinstance(error, XlsxError) else "Elaborazione locale XLSX non riuscita."
                with store.lock:
                    entry.update(state="review" if entry["job"].analysed else "uploaded", error=message)
                    entry["job"].output = entry["job"].dictionary = None
            finally:
                with store.lock:
                    if entry["cancelled"]:
                        store.entries.pop(entry["id"], None)
        pool.submit(run)

    @bp.get("/xlsx")
    def page():
        source = Path(__file__).with_name("xlsx_ui.html")
        # _MEIPASS bundle data is placed beside the module's source entry point.
        if not source.exists():
            import sys
            source = Path(getattr(sys, "_MEIPASS", "")) / "xlsx_ui.html"
        response = Response(source.read_text(encoding="utf-8"), mimetype="text/html")
        response.set_cookie(COOKIE, request.cookies.get(COOKIE) or secrets.token_urlsafe(32),
                            httponly=True, samesite="Strict", max_age=store.ttl)
        return response

    @bp.post("/api/xlsx/upload")
    def upload():
        if not ingest.acquire(blocking=False):
            return jsonify(error="Un altro caricamento è in corso."), 429
        try:
            fs = request.files.get("file")
            if fs is None:
                raise XlsxError("Selezionare un file .xlsx.")
            raw = fs.stream.read(limits.compressed + 1)
            job = WorkbookJob(raw, fs.filename or "", limits)
            owner = request.cookies.get(COOKIE) or secrets.token_urlsafe(32)
            identifier, _ = store.add(job, owner)
            response = jsonify(id=identifier, **job.summary(), ttl_seconds=store.ttl)
            response.set_cookie(COOKIE, owner, httponly=True, samesite="Strict", max_age=store.ttl)
            return response, 201
        finally:
            ingest.release()

    @bp.get("/api/xlsx/<identifier>/status")
    def status(identifier):
        with store.lock:
            entry, error = entry_for(identifier)
            if error:
                return error
            job = entry["job"]
            return jsonify(state=entry["state"], progress=job.progress, error=entry["error"],
                           changes=len(job.changes), mapping_enabled=bool(job.session and job.session.mapping_enabled))

    @bp.post("/api/xlsx/<identifier>/analyze")
    def analyze(identifier):
        with store.lock:
            entry, error = entry_for(identifier, idle=True)
            if error:
                return error
            config = request.get_json()
            if not isinstance(config, dict):
                raise XlsxError("Configurazione XLSX non valida.")
            excluded, mapping = preferences()
            config.setdefault("exclude_tags", list(excluded))
            config.setdefault("include_mapping", mapping)
            start(entry, "analyze", config)
            return jsonify(state="analysing"), 202

    @bp.get("/api/xlsx/<identifier>/preview")
    def preview(identifier):
        with store.lock:
            entry, error = entry_for(identifier, idle=True)
            if error:
                return error
            try:
                page = int(request.args.get("page", "0"))
                size = int(request.args.get("size", "50"))
            except ValueError:
                raise XlsxError("Numero di pagina non valido.") from None
            return jsonify(entry["job"].preview(page, size))

    @bp.patch("/api/xlsx/<identifier>/review")
    def review(identifier):
        with store.lock:
            entry, error = entry_for(identifier, idle=True)
            if error:
                return error
            entry["job"].correct(request.get_json())
            entry.update(state="review", error=None)
            return jsonify(ok=True)

    @bp.get("/api/xlsx/<identifier>/fields")
    def fields(identifier):
        with store.lock:
            entry, error = entry_for(identifier, idle=True)
            if error:
                return error
            try:
                page = int(request.args.get("page", "0"))
                size = int(request.args.get("size", "25"))
            except ValueError:
                raise XlsxError("Numero di pagina non valido.") from None
            return jsonify(entry["job"].fields(page, size, request.args.get("scope", "metadata")))

    @bp.post("/api/xlsx/<identifier>/review")
    def add(identifier):
        with store.lock:
            entry, error = entry_for(identifier, idle=True)
            if error:
                return error
            entry["job"].add(request.get_json())
            entry.update(state="review", error=None)
            return jsonify(ok=True), 201

    @bp.post("/api/xlsx/<identifier>/export")
    def export(identifier):
        with store.lock:
            entry, error = entry_for(identifier, idle=True)
            if error:
                return error
            if not entry["job"].analysed:
                raise XlsxError("Analizzare prima il documento.")
            start(entry, "export")
            return jsonify(state="exporting"), 202

    @bp.get("/api/xlsx/<identifier>/download")
    def download(identifier):
        with store.lock:
            entry, error = entry_for(identifier, idle=True)
            if error:
                return error
            if entry["state"] != "ready" or entry["job"].output is None:
                raise XlsxError("Download bloccato: esportazione verificata non disponibile.")
            return send_file(io.BytesIO(entry["job"].output), mimetype=MIME, as_attachment=True,
                             download_name="documento_pseudonimizzato.xlsx", max_age=0)

    @bp.get("/api/xlsx/<identifier>/mapping")
    def mapping(identifier):
        with store.lock:
            entry, error = entry_for(identifier, idle=True)
            if error:
                return error
            dictionary = entry["job"].dictionary
            if entry["state"] != "ready" or dictionary is None:
                return jsonify(error="Dizionario non abilitato o esportazione non verificata."), 403
            return send_file(io.BytesIO(json.dumps(dictionary, ensure_ascii=False).encode()),
                             mimetype="application/json", as_attachment=True,
                             download_name="dizionario_ripristino_xlsx.json", max_age=0)

    @bp.delete("/api/xlsx/<identifier>")
    def delete(identifier):
        if not store.delete(identifier, request.cookies.get(COOKIE)):
            return jsonify(error="Documento assente o scaduto."), 404
        return "", 204

    @bp.post("/api/xlsx/restore")
    def restore_route():
        if not ingest.acquire(blocking=False):
            return jsonify(error="Un'altra importazione è in corso."), 429
        try:
            file, dictionary = request.files.get("file"), request.files.get("mapping")
            if file is None or dictionary is None or not (file.filename or "").lower().endswith(".xlsx"):
                raise XlsxError("Caricare il file .xlsx elaborato e il dizionario .json.")
            raw = file.stream.read(limits.compressed + 1)
            mapping_bytes = dictionary.stream.read(limits.characters * 8 + 1025)
            if len(mapping_bytes) > limits.characters * 8 + 1024:
                raise XlsxError("Dizionario troppo grande.")
            def unique(pairs):
                result = {}
                for key, value in pairs:
                    if key in result:
                        raise ValueError
                    result[key] = value
                return result
            try:
                parsed = json.loads(mapping_bytes, object_pairs_hook=unique)
            except (ValueError, UnicodeError, RecursionError):
                raise XlsxError("Dizionario JSON non valido o con chiavi duplicate.") from None
            output = restore(raw, parsed, limits)
            return send_file(io.BytesIO(output), mimetype=MIME, as_attachment=True,
                             download_name="documento_ripristinato.xlsx", max_age=0)
        finally:
            ingest.release()

    app.register_blueprint(bp)
    return bp
