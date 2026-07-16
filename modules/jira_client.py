"""
jira_client.py
Fetches issues from Jira Cloud and returns them as a pandas DataFrame using the
SAME column names as the Jira CSV export, so the existing parser/report pipeline
consumes them unchanged (see modules/parser.py for the expected schema).

Configured via st.secrets["jira"]:
    [jira]
    base_url = "https://<yourco>.atlassian.net"
    email = "<you>@zigram.tech"
    api_token = "..."
    target_start_field = "customfield_XXXXX"   # optional; else auto-discovered
    target_end_field   = "customfield_YYYYY"   # optional
"""

from __future__ import annotations

import pandas as pd
import streamlit as st

# Standard Jira fields we always request. Custom Target start/end fields are
# appended once resolved.
_BASE_FIELDS = [
    "summary", "issuetype", "status", "priority",
    "assignee", "parent", "created", "updated", "comment", "labels",
]


def _cfg():
    try:
        section = st.secrets.get("jira")
    except Exception:
        return None
    if not section:
        return None
    if all(section.get(k) for k in ("base_url", "email", "api_token")):
        return dict(section)
    return None


def is_configured() -> bool:
    return _cfg() is not None


def _auth(cfg):
    return (cfg["email"], cfg["api_token"])


def _base(cfg) -> str:
    return cfg["base_url"].rstrip("/")


def build_jql(query: dict) -> str:
    """query = {"mode": "filter"|"jql", "value": "<filter id or JQL>"}."""
    mode = (query or {}).get("mode", "jql")
    value = str((query or {}).get("value", "")).strip()
    if not value:
        raise ValueError("No Jira saved-filter ID or JQL provided.")
    return f"filter={value}" if mode == "filter" else value


def discover_target_fields(cfg) -> tuple:
    """Resolve the Target-start / Target-end custom field ids: use the ones in
    secrets if given, else look them up by name via /rest/api/3/field."""
    start = cfg.get("target_start_field")
    end = cfg.get("target_end_field")
    if start and end:
        return start, end
    import requests

    resp = requests.get(f"{_base(cfg)}/rest/api/3/field", auth=_auth(cfg), timeout=30)
    resp.raise_for_status()
    fields = resp.json()

    def find(substr):
        for f in fields:
            if substr in str(f.get("name", "")).lower():
                return f.get("id")
        return None

    return (start or find("target start")), (end or find("target end"))


def _adf_to_text(node) -> str:
    """Flatten an Atlassian Document Format node (comment body) to plain text."""
    if node is None:
        return ""
    if isinstance(node, str):
        return node
    if isinstance(node, list):
        return "".join(_adf_to_text(c) for c in node)
    if isinstance(node, dict):
        text = node.get("text", "") if node.get("type") == "text" else ""
        return text + "".join(_adf_to_text(c) for c in (node.get("content") or []))
    return ""


def _latest_comment_text(fields: dict) -> str:
    comments = ((fields.get("comment") or {}).get("comments")) or []
    if not comments:
        return ""
    body = comments[-1].get("body")
    if isinstance(body, (dict, list)):
        return _adf_to_text(body).strip()
    return str(body or "").strip()


def _issue_to_row(issue: dict, start_field, end_field) -> dict:
    f = issue.get("fields", {}) or {}
    assignee = f.get("assignee") or {}
    return {
        "Issue key": issue.get("key", ""),
        "Issue Type": (f.get("issuetype") or {}).get("name", ""),
        "Summary": f.get("summary", "") or "",
        "Status": (f.get("status") or {}).get("name", ""),
        "Priority": (f.get("priority") or {}).get("name", "") if f.get("priority") else "",
        "Assignee": assignee.get("displayName", "Unassigned") if assignee else "Unassigned",
        "Parent key": (f.get("parent") or {}).get("key", "") if f.get("parent") else "",
        "Custom field (Target start)": f.get(start_field) if start_field else None,
        "Custom field (Target end)": f.get(end_field) if end_field else None,
        "Created": f.get("created"),
        "Updated": f.get("updated"),
        "Comment": _latest_comment_text(f),
        "Labels": ", ".join(str(x).strip() for x in (f.get("labels") or []) if str(x).strip()),
    }


def fetch_issues(query: dict, cfg: dict | None = None) -> pd.DataFrame:
    """Run the saved filter / JQL and return a DataFrame in the CSV-export schema."""
    import requests

    cfg = cfg or _cfg()
    if not cfg:
        raise RuntimeError(
            "Jira is not configured. Add a [jira] section (base_url, email, api_token) to Streamlit secrets."
        )
    jql = build_jql(query)
    start_field, end_field = discover_target_fields(cfg)
    fields = list(_BASE_FIELDS)
    for f in (start_field, end_field):
        if f and f not in fields:
            fields.append(f)

    rows = []
    next_token = None
    for _ in range(1000):  # safety cap on pages
        payload = {"jql": jql, "fields": fields, "maxResults": 100}
        if next_token:
            payload["nextPageToken"] = next_token
        resp = requests.post(
            f"{_base(cfg)}/rest/api/3/search/jql",
            json=payload, auth=_auth(cfg), timeout=60,
        )
        if resp.status_code >= 400:
            raise RuntimeError(f"Jira search failed (HTTP {resp.status_code}): {(resp.text or '')[:400]}")
        body = resp.json()
        for issue in body.get("issues", []) or []:
            rows.append(_issue_to_row(issue, start_field, end_field))
        next_token = body.get("nextPageToken")
        if body.get("isLast") or not next_token:
            break

    return pd.DataFrame(rows)
