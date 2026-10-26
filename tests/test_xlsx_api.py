"""API lifecycle tests with a simulated detector, without torch or model weights."""
import io
import json
import sys
import threading
import time
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src/app"))
try:
    from flask import Flask, request
    from xlsx_api import Store, register_xlsx
    from test_xlsx_io import EMAIL, NAME, fake_batch, load_workbook, save, workbook
    AVAILABLE = True
except ImportError:
    AVAILABLE = False

HEADERS = {"X-Rizzo-XLSX": "1"}


@unittest.skipUnless(AVAILABLE, "Installare requirements-xlsx-test.txt (job XLSX obbligatorio in CI)")
class XlsxApi(unittest.TestCase):
    def setUp(self):
        self.app = Flask(__name__)
        self.app.config.update(TESTING=True, MAX_CONTENT_LENGTH=50 * 1024 * 1024)
        self.clock = [100.0]
        self.store = Store(maximum=3, ttl=60, now=lambda: self.clock[0])
        self.blueprint = register_xlsx(self.app, lambda texts, budget: fake_batch(texts),
                                       store=self.store, start_sweeper=False)
        self.client = self.app.test_client()

    def tearDown(self):
        self.blueprint.executor.shutdown(wait=True)

    def upload(self, book=None, client=None):
        client = client or self.client
        response = client.post("/api/xlsx/upload", headers=HEADERS,
                               data={"file": (io.BytesIO(save(book or workbook())), "synthetic.xlsx")})
        self.assertEqual(response.status_code, 201, response.get_json())
        return "/api/xlsx/" + response.json["id"]

    def wait(self, url):
        deadline = time.monotonic() + 8
        while time.monotonic() < deadline:
            response = self.client.get(url + "/status")
            self.assertEqual(response.status_code, 200)
            if response.json["state"] not in {"analysing", "exporting"}:
                self.assertIsNone(response.json["error"], response.json["error"])
                return response.json
            time.sleep(.005)
        self.fail("Background operation timed out")

    def analyze(self, url, config=None):
        r = self.client.post(url + "/analyze", json=config or {}, headers=HEADERS)
        self.assertEqual(r.status_code, 202)
        return self.wait(url)

    def export(self, url):
        self.assertEqual(self.client.post(url + "/export", headers=HEADERS).status_code, 202)
        return self.wait(url)

    def test_complete_flow_review_export_verify_and_restore(self):
        book = workbook({"A1": "Identificativo", "A2": 5432178, "B2": NAME, "C2": EMAIL})
        book.create_sheet("Altro")["A1"] = NAME
        url = self.upload(book)
        self.analyze(url, {"columns": [{"sheet": 0, "column": 1, "category": "ACCOUNT"}]})
        preview = self.client.get(url + "/preview?size=2").json
        self.assertEqual(len(preview["items"]), 2)
        identifier = preview["items"][1]["id"]
        review = self.client.patch(url + "/review", headers=HEADERS,
                                   json={"id": identifier, "category": "PERSON", "scope": "associated"})
        self.assertEqual(review.status_code, 200)
        self.export(url)
        output = self.client.get(url + "/download")
        mapping = self.client.get(url + "/mapping")
        self.assertEqual(output.status_code, 200)
        self.assertEqual(mapping.status_code, 200)
        masked = load_workbook(io.BytesIO(output.data))
        self.assertEqual(masked.active["B2"].value, masked.worksheets[1]["A1"].value)
        self.assertNotEqual(masked.active["A2"].value, 5432178)
        restored = self.client.post("/api/xlsx/restore", headers=HEADERS,
            data={"file": (io.BytesIO(output.data), "masked.xlsx"),
                  "mapping": (io.BytesIO(mapping.data), "dictionary.json")})
        self.assertEqual(restored.status_code, 200, restored.get_json())
        result = load_workbook(io.BytesIO(restored.data))
        for coordinate in ["A1", "A2", "B2", "C2"]:
            self.assertEqual(result.active[coordinate].value, book.active[coordinate].value)
            self.assertEqual(result.active[coordinate].data_type, book.active[coordinate].data_type)

    def test_owner_isolation_even_with_known_id(self):
        url = self.upload()
        other = self.app.test_client()
        for endpoint in ["status", "preview", "download", "mapping"]:
            self.assertEqual(other.get(url + "/" + endpoint).status_code, 404)
        self.assertEqual(other.delete(url, headers=HEADERS).status_code, 404)

    def test_csrf_host_origin_and_no_cache(self):
        response = self.client.post("/api/xlsx/upload")
        self.assertEqual(response.status_code, 403)
        for headers in [{**HEADERS, "Origin": "http://evil.invalid"},
                        {**HEADERS, "Sec-Fetch-Site": "cross-site"},
                        {**HEADERS, "Host": "rebind.invalid"}]:
            response = self.client.post("/api/xlsx/upload", headers=headers)
            self.assertEqual(response.status_code, 403)
        url = self.upload()
        response = self.client.get(url + "/status")
        self.assertIn("no-store", response.headers["Cache-Control"])
        self.assertEqual(response.headers["X-Content-Type-Options"], "nosniff")

    def test_expiry_and_deletion(self):
        url = self.upload()
        self.clock[0] += 61
        self.assertEqual(self.client.get(url + "/status").status_code, 404)
        self.assertEqual(len(self.store.entries), 0)
        url = self.upload()
        self.assertEqual(self.client.delete(url, headers=HEADERS).status_code, 204)
        self.assertEqual(self.client.get(url + "/status").status_code, 404)

    def test_store_capacity_memory_admission_and_expiry(self):
        self.store.maximum = 1
        self.upload()
        response = self.client.post("/api/xlsx/upload", headers=HEADERS,
            data={"file": (io.BytesIO(save(workbook())), "synthetic.xlsx")})
        self.assertEqual(response.status_code, 422)
        self.clock[0] += 61
        self.upload()
        self.store.memory = 1
        self.clock[0] += 61
        response = self.client.post("/api/xlsx/upload", headers=HEADERS,
            data={"file": (io.BytesIO(save(workbook())), "synthetic.xlsx")})
        self.assertEqual(response.status_code, 422)

    def test_without_mapping_no_originals_in_responses(self):
        url = self.upload()
        self.analyze(url, {"include_mapping": False})
        preview = self.client.get(url + "/preview")
        self.assertNotIn(EMAIL.encode(), preview.data)
        fields = self.client.get(url + "/fields?scope=all")
        self.assertNotIn(EMAIL.encode(), fields.data)
        self.assertNotIn(b'"value"', fields.data)
        self.export(url)
        self.assertEqual(self.client.get(url + "/mapping").status_code, 403)
        self.assertEqual(self.client.get(url + "/download").status_code, 200)

    def test_add_review_invalidates_verified_download(self):
        url = self.upload(workbook({"A1": EMAIL, "B1": "record-78"}))
        self.analyze(url)
        self.export(url)
        response = self.client.post(url + "/review", headers=HEADERS,
                                    json={"sheet": 0, "cell": "B1", "category": "ID"})
        self.assertEqual(response.status_code, 201)
        self.assertEqual(self.client.get(url + "/download").status_code, 422)
        self.assertEqual(self.client.get(url + "/mapping").status_code, 403)
        self.export(url)

    def test_errors_never_include_document_content(self):
        response = self.client.post("/api/xlsx/upload", headers=HEADERS,
                                    data={"file": (io.BytesIO(EMAIL.encode()), NAME + ".xlsx")})
        self.assertEqual(response.status_code, 422)
        self.assertNotIn(EMAIL.encode(), response.data)
        self.assertNotIn(NAME.encode(), response.data)
        url = self.upload()
        self.analyze(url)
        response = self.client.patch(url + "/review", headers=HEADERS, json={"id": 0, "category": EMAIL})
        self.assertEqual(response.status_code, 422)
        self.assertNotIn(EMAIL.encode(), response.data)

    def test_download_blocked_before_analysis_and_before_export(self):
        url = self.upload()
        self.assertEqual(self.client.get(url + "/download").status_code, 422)
        self.analyze(url)
        self.assertEqual(self.client.get(url + "/download").status_code, 422)

    def test_bad_pages_and_json(self):
        url = self.upload(); self.analyze(url)
        for query in ["page=no", "page=-1", "size=999"]:
            self.assertEqual(self.client.get(url + "/preview?" + query).status_code, 422)
            self.assertEqual(self.client.get(url + "/fields?" + query).status_code, 422)
        self.assertEqual(self.client.patch(url + "/review", headers=HEADERS,
                                          data="{bad", content_type="application/json").status_code, 400)

    def test_metadata_inventory_add_and_pagination(self):
        book = workbook(); book.properties.description = "record-metadata-82"
        url = self.upload(book); self.analyze(url)
        fields = self.client.get(url + "/fields?scope=metadata&size=100")
        self.assertEqual(fields.status_code, 200)
        field = next(f for f in fields.json["items"] if f["field"] == "description")
        self.assertEqual(self.client.post(url + "/review", headers=HEADERS,
                                         json={"slot": field["slot"], "category": "META"}).status_code, 201)
        self.export(url)
        output = load_workbook(io.BytesIO(self.client.get(url + "/download").data))
        self.assertNotEqual(output.properties.description, "record-metadata-82")

    def test_restore_duplicate_json_keys_and_wrong_dictionary(self):
        url = self.upload(); self.analyze(url); self.export(url)
        raw = self.client.get(url + "/download").data
        for mapping in [b'{"version":1,"version":1}', b"not json", b"{}"]:
            response = self.client.post("/api/xlsx/restore", headers=HEADERS,
                data={"file": (io.BytesIO(raw), "masked.xlsx"), "mapping": (io.BytesIO(mapping), "bad.json")})
            self.assertEqual(response.status_code, 422)

    def test_ui_offline_no_external_scripts_and_owner_cookie(self):
        response = self.client.get("/xlsx")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"/api/xlsx", response.data)
        self.assertNotIn(b'<script src="http', response.data)
        cookie = response.headers["Set-Cookie"]
        self.assertIn("HttpOnly", cookie)
        self.assertIn("SameSite=Strict", cookie)

    def test_upload_parser_never_spools_xlsx_documents_to_disk(self):
        body = (b'--synthetic-boundary\r\nContent-Disposition: form-data; name="file"; '
                b'filename="synthetic.xlsx"\r\nContent-Type: application/octet-stream\r\n\r\n' +
                b"synthetic" * 80000 + b'\r\n--synthetic-boundary--\r\n')
        with self.app.test_request_context("/api/xlsx/upload", method="POST",
             data=body, content_type="multipart/form-data; boundary=synthetic-boundary"):
            self.assertIsInstance(request.files["file"].stream, io.BytesIO)

    def test_concurrent_operations_rejected_and_delete_cancels(self):
        # Hold a real background worker at a deterministic batch boundary.
        ready, release = threading.Event(), threading.Event()
        previous = self.blueprint.executor.submit
        def delayed(fn):
            def wrap():
                ready.set()
                release.wait(5)
                return fn()
            return previous(wrap)
        self.blueprint.executor.submit = delayed
        url = self.upload()
        try:
            response = self.client.post(url + "/analyze", headers=HEADERS, json={})
            self.assertEqual(response.status_code, 202)
            self.assertTrue(ready.wait(2))
            self.assertEqual(self.client.post(url + "/analyze", headers=HEADERS, json={}).status_code, 409)
            self.assertEqual(self.client.patch(url + "/review", headers=HEADERS, json={}).status_code, 409)
            self.assertEqual(self.client.delete(url, headers=HEADERS).status_code, 204)
            self.assertEqual(self.client.get(url + "/status").status_code, 404)
        finally:
            release.set()


if __name__ == "__main__":
    unittest.main()
