"""Request-local, exact-value pseudonyms. No Flask or model imports."""
import hashlib
import hmac
import json
import re
import secrets


class PiiSession:
    def __init__(self, inputs, mapping_enabled=True):
        self.mapping_enabled = mapping_enabled
        self.namespace = secrets.token_hex(8)
        while any(self.namespace in str(value) for value in inputs):
            self.namespace = secrets.token_hex(8)
        self._secret = secrets.token_bytes(32)
        self._seen = {}
        self._counters = {}
        self.mapping = {}

    def placeholder(self, category, value, kind="text"):
        if not isinstance(category, str) or not re.fullmatch(r"[A-Z][A-Z0-9_]{0,31}", category):
            raise ValueError("Categoria/prefisso non valido (A-Z, 0-9, underscore; max 32).")
        # Exact spelling, whitespace, Unicode and type are significant. The index
        # contains keyed digests rather than an exportable original-value key.
        raw = json.dumps([category, kind, value], ensure_ascii=False).encode()
        key = hmac.digest(self._secret, raw, hashlib.sha256)
        if key not in self._seen:
            number = self._counters.get(category, 0) + 1
            self._counters[category] = number
            ph = f"[[RZ_{self.namespace}_{category}_{number}]]"
            self._seen[key] = ph
            if self.mapping_enabled:
                self.mapping[ph] = {"kind": kind, "value": value}
        return self._seen[key]

    def discard_index(self):
        self._seen.clear()
        self._secret = b""
