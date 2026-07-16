"""
store.py
Persistent storage for saved sprint details and per-project report configs
(buckets, KPIs, status mapping).

On Streamlit Community Cloud the local filesystem is wiped whenever the app
sleeps, reboots, or redeploys, so anything written to local JSON files is
lost. This module persists data to a **Zoho Sheet** when Zoho credentials are
present in Streamlit secrets, and falls back to local JSON files otherwise
(so local development keeps working with no credentials).

Design note - the Zoho backend is intentionally **append-only** and uses only
the two Zoho Sheet operations that are unambiguously documented,
`worksheet.records.fetch` and `worksheet.records.add`. It never relies on the
fragile update / delete-by-criteria API:

  * every save appends a row  (store, project, data, updated_at)
  * on load, the newest row per (store, project) wins
  * a delete appends a tombstone row (data = "__deleted__")
"""

from __future__ import annotations

import json
import time
from datetime import datetime, timezone
from pathlib import Path

import streamlit as st

APP_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = APP_DIR / "data"

_TOMBSTONE = "__deleted__"
_FETCH_PAGE = 1000


def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%fZ")


class LocalFileBackend:
    """Stores each logical store as a dict[project] -> obj in a JSON file."""

    def _path(self, store: str) -> Path:
        return DATA_DIR / f"{store}.local.json"

    def load_all(self, store: str) -> dict:
        path = self._path(store)
        if not path.exists():
            return {}
        try:
            with path.open("r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except (OSError, json.JSONDecodeError):
            return {}

    def _write(self, store: str, data: dict) -> None:
        path = self._path(store)
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(".tmp")
        with tmp.open("w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, sort_keys=True)
        tmp.replace(path)

    def save_one(self, store: str, project: str, obj) -> None:
        data = self.load_all(store)
        data[project] = obj
        self._write(store, data)

    def delete_one(self, store: str, project: str) -> None:
        data = self.load_all(store)
        data.pop(project, None)
        self._write(store, data)


class ZohoSheetBackend:
    """Append-only persistence backed by a single Zoho Sheet worksheet."""

    def __init__(self, cfg: dict):
        self.client_id = cfg["client_id"]
        self.client_secret = cfg["client_secret"]
        self.refresh_token = cfg["refresh_token"]
        self.resource_id = cfg["resource_id"]
        self.worksheet = cfg.get("worksheet_name", "app_data")
        dc = str(cfg.get("location", "com")).strip().lstrip(".") or "com"
        self.accounts_host = cfg.get("accounts_domain") or f"https://accounts.zoho.{dc}"
        self.sheet_host = cfg.get("sheet_domain") or f"https://sheet.zoho.{dc}"
        self._token = None
        self._token_exp = 0.0

    def _access_token(self) -> str:
        import requests

        if self._token and time.time() < self._token_exp - 60:
            return self._token
        resp = requests.post(
            f"{self.accounts_host}/oauth/v2/token",
            data={
                "refresh_token": self.refresh_token,
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "grant_type": "refresh_token",
            },
            timeout=20,
        )
        resp.raise_for_status()
        body = resp.json()
        if "access_token" not in body:
            raise RuntimeError(f"Zoho token error: {body}")
        self._token = body["access_token"]
        self._token_exp = time.time() + int(body.get("expires_in", 3600))
        return self._token

    def _api(self, method: str, extra: dict) -> dict:
        import requests

        headers = {"Authorization": f"Zoho-oauthtoken {self._access_token()}"}
        data = {"method": method, "worksheet_name": self.worksheet}
        data.update(extra)
        resp = requests.post(
            f"{self.sheet_host}/api/v2/{self.resource_id}",
            headers=headers,
            data=data,
            timeout=30,
        )
        try:
            body = resp.json()
        except ValueError:
            body = None
        failed = resp.status_code >= 400 or (
            isinstance(body, dict) and body.get("status") == "failure"
        )
        if failed:
            detail = json.dumps(body) if body is not None else (resp.text or "")[:400]
            raise RuntimeError(f"Zoho Sheet {method} failed (HTTP {resp.status_code}): {detail}")
        return body or {}

    def _fetch_records(self) -> list:
        records = []
        start = 1
        while True:
            body = self._api(
                "worksheet.records.fetch",
                {"records_start_index": str(start), "count": str(_FETCH_PAGE)},
            )
            page = body.get("records", []) or []
            records.extend(page)
            if len(page) < _FETCH_PAGE:
                break
            start += len(page)
        return records

    def _add_record(self, row: dict) -> None:
        self._api("worksheet.records.add", {"json_data": json.dumps([row])})

    def load_all(self, store: str) -> dict:
        latest: dict[str, tuple[str, str]] = {}
        for rec in self._fetch_records():
            if rec.get("store") != store:
                continue
            project = rec.get("project")
            if not project:
                continue
            ts = str(rec.get("updated_at", ""))
            if project not in latest or ts >= latest[project][0]:
                latest[project] = (ts, rec.get("data", ""))
        out = {}
        for project, (_, raw) in latest.items():
            if raw in (_TOMBSTONE, "", None):
                continue
            try:
                out[project] = json.loads(raw)
            except (json.JSONDecodeError, TypeError):
                continue
        return out

    def save_one(self, store: str, project: str, obj) -> None:
        self._add_record(
            {"store": store, "project": project,
             "data": json.dumps(obj), "updated_at": _now_iso()}
        )

    def delete_one(self, store: str, project: str) -> None:
        self._add_record(
            {"store": store, "project": project,
             "data": _TOMBSTONE, "updated_at": _now_iso()}
        )


def _zoho_cfg():
    try:
        section = st.secrets.get("zoho")
    except Exception:
        return None
    if not section:
        return None
    required = ["client_id", "client_secret", "refresh_token", "resource_id"]
    if all(section.get(k) for k in required):
        return dict(section)
    return None


@st.cache_resource(show_spinner=False)
def get_backend():
    cfg = _zoho_cfg()
    if cfg:
        return ZohoSheetBackend(cfg)
    return LocalFileBackend()


def backend_kind() -> str:
    """'ZohoSheetBackend' or 'LocalFileBackend' - for a status indicator."""
    return type(get_backend()).__name__
