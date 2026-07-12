"""
cliq_client.py
Posts a summary message and uploads the generated report files (Excel/PDF/PNG)
to a Zoho Cliq channel.

Configured via st.secrets["cliq"]:
    [cliq]
    client_id = "..."
    client_secret = "..."
    refresh_token = "..."   # self-client generated with ZohoCliq.Channels.CREATE scope
    location = "in"         # Zoho data center (in / com / eu / ...)

Design mirrors the storage layer: OAuth refresh-token flow, and every API error
surfaces Zoho's response body so a wrong endpoint/scope is easy to diagnose. The
message is posted first (best-documented call); each file upload is attempted
independently and failures are returned as warnings rather than aborting.
"""

from __future__ import annotations

import time

import streamlit as st


def _cfg():
    try:
        section = st.secrets.get("cliq")
    except Exception:
        return None
    if not section:
        return None
    if all(section.get(k) for k in ("client_id", "client_secret", "refresh_token")):
        return dict(section)
    return None


def is_configured() -> bool:
    return _cfg() is not None


class CliqClient:
    def __init__(self, cfg: dict):
        self.client_id = cfg["client_id"]
        self.client_secret = cfg["client_secret"]
        self.refresh_token = cfg["refresh_token"]
        dc = str(cfg.get("location", "in")).strip().lstrip(".") or "in"
        self.accounts_host = cfg.get("accounts_domain") or f"https://accounts.zoho.{dc}"
        self.api_host = cfg.get("cliq_domain") or f"https://cliq.zoho.{dc}"
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
            raise RuntimeError(f"Cliq token error: {body}")
        self._token = body["access_token"]
        self._token_exp = time.time() + int(body.get("expires_in", 3600))
        return self._token

    def _headers(self) -> dict:
        return {"Authorization": f"Zoho-oauthtoken {self._access_token()}"}

    def post_message(self, channel: str, text: str) -> None:
        import requests

        url = f"{self.api_host}/api/v2/channelsbyname/{channel}/message"
        resp = requests.post(url, headers=self._headers(), json={"text": text}, timeout=30)
        if resp.status_code >= 400:
            raise RuntimeError(f"Cliq message failed (HTTP {resp.status_code}): {(resp.text or '')[:400]}")

    def upload_file(self, channel: str, filename: str, data: bytes, mime: str) -> None:
        import requests

        url = f"{self.api_host}/api/v2/channelsbyname/{channel}/files"
        resp = requests.post(
            url, headers=self._headers(),
            files={"file": (filename, data, mime)}, timeout=60,
        )
        if resp.status_code >= 400:
            raise RuntimeError(f"Cliq upload '{filename}' failed (HTTP {resp.status_code}): {(resp.text or '')[:400]}")


@st.cache_resource(show_spinner=False)
def _client():
    cfg = _cfg()
    return CliqClient(cfg) if cfg else None


def post_report(channel: str, files: list, message: str) -> list:
    """files: list of (filename, bytes, mime). Posts the message, then uploads
    each file. Returns a list of warning strings (empty = all succeeded).
    Raises only if Cliq is unconfigured, the channel is missing, or the message
    itself cannot be posted."""
    client = _client()
    if client is None:
        raise RuntimeError("Zoho Cliq is not configured. Add a [cliq] section to Streamlit secrets.")
    if not channel:
        raise RuntimeError("No Zoho Cliq channel is set for this project.")
    client.post_message(channel, message)
    warnings = []
    for filename, data, mime in files:
        try:
            client.upload_file(channel, filename, data, mime)
        except Exception as exc:
            warnings.append(str(exc))
    return warnings
