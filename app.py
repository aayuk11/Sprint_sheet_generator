"""
app.py - Sprint Report Generator
Run with: streamlit run app.py
"""

import json
import base64
import html
from pathlib import Path
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from datetime import date, datetime
from zoneinfo import ZoneInfo

from modules.parser          import (
    parse_jira_csv,
    BUCKET_KEYS, BUCKET_LABELS,
    PCT_KEYS, PCT_LABELS, DEFAULT_PCT_BUCKETS,
    PROJECT_STATUS_MAP, PROJECT_PCT_BUCKETS,
)
from modules.excel_generator import build_excel
from modules.pdf_generator   import build_pdf
from modules.image_generator import build_summary_image
from modules import store
from modules import jira_client
from modules import cliq_client

PROJECTS = [
    "PreScreening.io",
    "Transact Comply",
    "Entity Hero",
    "DueDiliger",
    "ZiZi",
    "SATOC",
    "WMP",
    "Fraud Fighter",
    "Profile Builder",
]

APP_DIR = Path(__file__).resolve().parent
SPRINT_DETAILS_PATH = APP_DIR / "data" / "sprint_details.local.json"
LEGACY_SPRINT_DETAILS_PATH = APP_DIR / "data" / "sprint_details.json"
STATUS_MAPPINGS_PATH = APP_DIR / "data" / "status_mappings.local.json"


def _today_ist() -> date:
    return datetime.now(ZoneInfo("Asia/Kolkata")).date()


def _load_store(name: str) -> dict:
    """Load a store's data via the configured backend (Zoho Sheet or local
    files). Never raises - records a message and returns {} on failure so the
    app still runs."""
    try:
        return store.get_backend().load_all(name)
    except Exception as exc:  # network / auth / API errors
        st.session_state["_store_error"] = f"Could not load saved data: {exc}"
        return {}


def _save_to_store(name: str, project: str, obj) -> bool:
    try:
        store.get_backend().save_one(name, project, obj)
        return True
    except Exception as exc:
        st.session_state["_store_error"] = f"Saved for this session, but could not persist to storage: {exc}"
        return False


def _delete_from_store(name: str, project: str) -> bool:
    try:
        store.get_backend().delete_one(name, project)
        return True
    except Exception as exc:
        st.session_state["_store_error"] = f"Cleared for this session, but could not update storage: {exc}"
        return False


def _load_saved_sprint_details() -> dict:
    return _load_store("sprint_details")


def _parse_saved_date(value, fallback: date) -> date:
    if not value:
        return fallback
    try:
        return date.fromisoformat(str(value))
    except ValueError:
        return fallback


def _project_defaults(project_name: str) -> dict:
    today = _today_ist()
    saved = st.session_state.saved_sprint_details.get(
        project_name,
        st.session_state.saved_sprint_details.get(project_name.replace(" ", "_"), {}),
    )
    return {
        "sprint_number_input": int(saved.get("sprint_number", 1) or 1),
        "sprint_start_input": _parse_saved_date(saved.get("sprint_start"), today),
        "dev_release_input": _parse_saved_date(saved.get("dev_release"), today),
        "qa_release_input": _parse_saved_date(saved.get("qa_release"), today),
        "prod_release_input": _parse_saved_date(saved.get("prod_release"), today),
        "sprint_end_input": _parse_saved_date(saved.get("sprint_end"), today),
        "scrum_master_input": saved.get("scrum_master", ""),
        "sprint_goal_input": saved.get("sprint_goal", ""),
        "major_item_1_input": saved.get("major_item_1", ""),
        "major_item_2_input": saved.get("major_item_2", ""),
        "major_item_3_input": saved.get("major_item_3", ""),
    }


def _hydrate_project_form(project_name: str) -> None:
    for key, value in _project_defaults(project_name).items():
        st.session_state[key] = value


def _project_options() -> list:
    """Base project list plus any custom projects the user has created/saved,
    so newly added projects keep appearing in the dropdown across sessions."""
    options = list(PROJECTS)
    for name in st.session_state.get("saved_sprint_details", {}):
        if name not in options:
            options.append(name)
    return options


def _load_selected_project_details() -> None:
    project_name = st.session_state.project_selector
    st.session_state.active_project = project_name
    _hydrate_project_form(project_name)
    st.session_state.uploaded_df = None
    st.session_state.excel_bytes = None
    st.session_state.pdf_bytes = None
    st.session_state.image_bytes = None
    st.session_state.parsed_report = None
    st.session_state.parsed_kpis = None


def _serialize_form_data(form_data: dict) -> dict:
    serialized = form_data.copy()
    for key in ["sprint_start", "dev_release", "qa_release", "prod_release", "sprint_end"]:
        serialized[key] = serialized[key].isoformat()
    return serialized


def _save_project_form_data(project_name: str, form_data: dict) -> None:
    serialized = _serialize_form_data(form_data)
    saved = st.session_state.saved_sprint_details.copy()
    saved[project_name] = serialized
    st.session_state.saved_sprint_details = saved
    _save_to_store("sprint_details", project_name, serialized)


def _load_status_mappings() -> dict:
    return _load_store("status_mappings")


def _resolve_status_config(project_name: str):
    """Config handed to the parser: the saved user mapping if one exists,
    else None so the parser falls back to its built-in map / global keywords."""
    return st.session_state.get("status_mappings", {}).get(project_name)


def _effective_status_config(project_name: str) -> dict:
    """Starting values for the Settings editor: saved mapping, else the
    built-in map for a known project, else an empty map with default rollups.
    `known_statuses` is the persisted comprehensive list of Jira statuses for
    the project (including ones mapped to 'ignore'), so the full mapping table
    survives across sessions and CSV uploads."""
    saved = st.session_state.get("status_mappings", {}).get(project_name)
    if saved:
        smap = dict(saved.get("status_map", {}))
        return {
            "status_map": smap,
            "pct_buckets": {k: list(v) for k, v in (saved.get("pct_buckets") or DEFAULT_PCT_BUCKETS).items()},
            "known_statuses": list(saved.get("known_statuses") or smap.keys()),
        }
    if project_name in PROJECT_STATUS_MAP:
        smap = dict(PROJECT_STATUS_MAP[project_name])
        return {
            "status_map": smap,
            "pct_buckets": {k: list(v) for k, v in PROJECT_PCT_BUCKETS[project_name].items()},
            "known_statuses": list(smap.keys()),
        }
    return {
        "status_map": {},
        "pct_buckets": {k: list(v) for k, v in DEFAULT_PCT_BUCKETS.items()},
        "known_statuses": [],
    }


def _effective_report_map(project_name: str):
    """The status->bucket map the parser will actually use for this project:
    saved user map if any, else the built-in map, else None (global fallback)."""
    saved = st.session_state.get("status_mappings", {}).get(project_name)
    if saved and saved.get("status_map"):
        return {str(k).lower().strip(): v for k, v in saved["status_map"].items()}
    if project_name in PROJECT_STATUS_MAP:
        return PROJECT_STATUS_MAP[project_name]
    return None


def _invalidate_generated_report() -> None:
    """Force a re-parse/re-build next time the report page is opened."""
    for key in ("excel_bytes", "pdf_bytes", "image_bytes", "parsed_report", "parsed_kpis"):
        st.session_state[key] = None


def _ingest_dataframe(df, source_label: str) -> bool:
    """Validate and accept a DataFrame (from CSV upload OR a Jira fetch) as the
    working dataset, then invalidate any previously generated report."""
    missing = [c for c in ["Issue key", "Issue Type", "Summary", "Status"] if c not in df.columns]
    if missing:
        st.error(f"Missing required columns: {', '.join(missing)}")
        return False
    st.session_state.uploaded_df = df
    st.session_state.uploaded_source = source_label
    _invalidate_generated_report()
    return True


def render_settings_page() -> None:
    st.markdown('<div class="section-title">Status Mapping</div>', unsafe_allow_html=True)
    st.caption("Map each project's Jira statuses to the report fields. Set once - reused automatically.")

    project = st.selectbox(
        "Project",
        _project_options(),
        key="settings_project",
        accept_new_options=True,
        help="Pick a project, or type a new name to configure it.",
    )

    current = _effective_status_config(project)
    label_to_key = {BUCKET_LABELS[b]: b for b in BUCKET_KEYS}
    bucket_labels = [BUCKET_LABELS[b] for b in BUCKET_KEYS]

    # Candidate statuses: uploaded CSV + already-saved + manual.
    status_display = {}
    for s in current.get("known_statuses", []):
        status_display.setdefault(str(s).lower().strip(), str(s))
    up_df = st.session_state.get("uploaded_df")
    if up_df is not None and "Status" in up_df.columns:
        source = up_df[up_df["Issue Type"] != "Epic"] if "Issue Type" in up_df.columns else up_df
        for s in source["Status"].dropna().unique():
            s = str(s).strip()
            if s:
                status_display.setdefault(s.lower(), s)
    for lower_key in current["status_map"]:
        status_display.setdefault(lower_key, lower_key)

    # Adding statuses is tucked away - open automatically only when there are none yet.
    with st.expander("＋ Add Jira statuses (for statuses not in the current CSV)",
                     expanded=not status_display):
        manual_raw = st.text_area(
            "One status per line",
            key=f"settings_manual_{project}",
            placeholder="In Work\nGrooming\nReady for QA",
        )
    for s in [x.strip() for x in manual_raw.splitlines() if x.strip()]:
        status_display.setdefault(s.lower(), s)
    ordered = sorted(status_display.items(), key=lambda kv: kv[1].lower())

    # Core: one dropdown per Jira status.
    st.markdown("**Map each Jira status to a report field**")
    if not ordered:
        st.info("Upload a Jira CSV on the report page, or add statuses above, to start mapping.")
    new_status_map = {}
    choices = ["(ignore)"] + bucket_labels
    for lower_key, disp in ordered:
        cur_bucket = current["status_map"].get(lower_key)
        default_label = BUCKET_LABELS.get(cur_bucket, "(ignore)")
        idx = choices.index(default_label) if default_label in choices else 0
        chosen = st.selectbox(disp, choices, index=idx, key=f"map_{project}_{lower_key}")
        if chosen != "(ignore)":
            new_status_map[lower_key] = label_to_key[chosen]

    unmapped = [disp for lk, disp in ordered if lk not in new_status_map]
    if unmapped:
        st.caption("⚠️ Not mapped (won't be counted): " + ", ".join(unmapped))

    # Advanced % rollups - hidden by default; defaults work for most projects.
    new_pct = {pk: list(current["pct_buckets"].get(pk, [])) for pk in PCT_KEYS}
    with st.expander("Advanced: which statuses roll into each % KPI (optional)"):
        for pk in PCT_KEYS:
            default_labels = [BUCKET_LABELS[b] for b in new_pct[pk] if b in BUCKET_LABELS]
            sel = st.multiselect(
                PCT_LABELS[pk], bucket_labels, default=default_labels, key=f"pct_{project}_{pk}"
            )
            new_pct[pk] = [label_to_key[l] for l in sel]

    # Optional preview - hidden by default.
    with st.expander("Preview: what feeds each report field"):
        by_bucket = {b: [] for b in BUCKET_KEYS}
        for lower_key, bucket in new_status_map.items():
            if bucket in by_bucket:
                by_bucket[bucket].append(status_display.get(lower_key, lower_key))
        rows = ""
        for b in BUCKET_KEYS:
            fed = ", ".join(_html_escape(s) for s in sorted(by_bucket[b])) or "<span style='color:#B45309;'>—</span>"
            rows += (f"<tr><td style='padding:5px 8px;border-bottom:1px solid #E2E8F0;font-size:12px;"
                     f"font-weight:600;'>{BUCKET_LABELS[b]}</td><td style='padding:5px 8px;"
                     f"border-bottom:1px solid #E2E8F0;font-size:12px;'>{fed}</td></tr>")
        for pk in PCT_KEYS:
            fed = ", ".join(BUCKET_LABELS[b] for b in new_pct.get(pk, [])) or "<span style='color:#B45309;'>—</span>"
            rows += (f"<tr><td style='padding:5px 8px;border-bottom:1px solid #E2E8F0;font-size:12px;"
                     f"font-weight:600;color:#375623;'>{PCT_LABELS[pk]}</td><td style='padding:5px 8px;"
                     f"border-bottom:1px solid #E2E8F0;font-size:12px;'>{fed}</td></tr>")
        st.markdown(
            "<table style='width:100%;border-collapse:collapse;'>"
            "<thead><tr style='background:#1F3864;color:white;'>"
            "<th style='padding:6px 8px;text-align:left;font-size:11px;'>Report field</th>"
            "<th style='padding:6px 8px;text-align:left;font-size:11px;'>Fed by</th>"
            f"</tr></thead><tbody>{rows}</tbody></table>",
            unsafe_allow_html=True,
        )

    save_col, reset_col = st.columns([1, 1])
    with save_col:
        if st.button("Save", type="primary", use_container_width=True):
            cfg = {
                "status_map": new_status_map,
                "pct_buckets": new_pct,
                "known_statuses": [disp for _, disp in ordered],
            }
            mappings = dict(st.session_state.get("status_mappings", {}))
            mappings[project] = cfg
            st.session_state.status_mappings = mappings
            _save_to_store("status_mappings", project, cfg)
            _invalidate_generated_report()
            st.success(f"Saved. {len(new_status_map)} status(es) mapped for {project}.")
    with reset_col:
        if st.button("Reset", use_container_width=True):
            mappings = dict(st.session_state.get("status_mappings", {}))
            mappings.pop(project, None)
            st.session_state.status_mappings = mappings
            _delete_from_store("status_mappings", project)
            _invalidate_generated_report()
            st.rerun()


def _report_base_name(form_data: dict) -> str:
    today = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d%b%Y")
    project_name = form_data['project_name'].replace("_", " ")
    return f"{project_name}_Sprint{form_data['sprint_number']}_{today}"


def _html_escape(value) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def _render_multi_file_download_button(files: list[tuple[str, bytes, str]]) -> None:
    downloads = [
        {
            "filename": filename,
            "mime": mime,
            "data": base64.b64encode(data).decode("ascii"),
        }
        for filename, data, mime in files
    ]
    downloads_json = json.dumps(downloads)
    file_count = len(downloads)
    label = "Download Selected File" if file_count == 1 else f"Download {file_count} Selected Files"
    components.html(
        f"""
        <button id="download-selected" type="button">{label}</button>
        <div id="download-status" aria-live="polite"></div>
        <script>
        const downloads = {downloads_json};
        const button = document.getElementById("download-selected");
        const status = document.getElementById("download-status");

        function base64ToBlob(base64, mime) {{
            const binary = window.atob(base64);
            const chunkSize = 8192;
            const chunks = [];
            for (let offset = 0; offset < binary.length; offset += chunkSize) {{
                const slice = binary.slice(offset, offset + chunkSize);
                const bytes = new Uint8Array(slice.length);
                for (let i = 0; i < slice.length; i += 1) {{
                    bytes[i] = slice.charCodeAt(i);
                }}
                chunks.push(bytes);
            }}
            return new Blob(chunks, {{ type: mime }});
        }}

        button.addEventListener("click", () => {{
            downloads.forEach((file, index) => {{
                window.setTimeout(() => {{
                    const blob = base64ToBlob(file.data, file.mime);
                    const url = URL.createObjectURL(blob);
                    const link = document.createElement("a");
                    link.href = url;
                    link.download = file.filename;
                    document.body.appendChild(link);
                    link.click();
                    link.remove();
                    window.setTimeout(() => URL.revokeObjectURL(url), 5000);
                    status.textContent = `Started ${{index + 1}} of ${{downloads.length}} downloads.`;
                }}, index * 500);
            }});
        }});
        </script>
        <style>
        #download-selected {{
            width: 100%;
            border: 0;
            border-radius: 8px;
            padding: 12px 32px;
            color: #ffffff;
            cursor: pointer;
            font-size: 14px;
            font-weight: 700;
            background: linear-gradient(135deg, #1F3864, #2E75B6);
            font-family: Arial, sans-serif;
        }}
        #download-selected:hover {{
            filter: brightness(1.06);
        }}
        #download-status {{
            min-height: 18px;
            margin-top: 8px;
            color: #64748B;
            font: 12px Arial, sans-serif;
        }}
        </style>
        """,
        height=78,
    )

st.set_page_config(
    page_title="Sprint Report Generator",
    page_icon=":bar_chart:",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
    .stApp { background-color: #F5F7FA; }
    footer { visibility: hidden; }
    .header-banner {
        background: linear-gradient(135deg, #1F3864 0%, #2E75B6 100%);
        padding: 28px 36px; border-radius: 12px; margin-bottom: 28px;
    }
    .header-banner h1 { color: white !important; font-size: 26px; font-weight: 700; margin: 0 0 6px 0; }
    .header-banner p  { color: #BDD7EE; font-size: 13px; margin: 0; }
    .section-title {
        font-size: 12px; font-weight: 700; color: #1F3864;
        text-transform: uppercase; letter-spacing: 0.8px;
        margin-bottom: 14px; padding-bottom: 8px; border-bottom: 2px solid #E2E8F0;
    }
    .info-pill { background: #EFF6FF; border-left: 3px solid #2E75B6; padding: 10px 14px; border-radius: 0 6px 6px 0; font-size: 12px; color: #1E40AF; margin: 8px 0; }
    .val-error { background: #FEF2F2; border: 1px solid #FCA5A5; border-radius: 6px; padding: 10px 14px; color: #DC2626; font-size: 12px; margin-top: 8px; }
    .divider { height: 1px; background: #E2E8F0; margin: 18px 0; }
    label, .stMarkdown, .stMetric label, [data-testid="stWidgetLabel"] p {
        color: #1F3864 !important;
    }
    [data-testid="stNumberInput"] input,
    [data-testid="stTextInput"] input,
    [data-testid="stDateInput"] input,
    [data-testid="stSelectbox"] div[data-baseweb="select"] > div {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border-color: #CBD5E1 !important;
    }
    [data-testid="stNumberInput"] input::placeholder,
    [data-testid="stTextInput"] input::placeholder,
    [data-testid="stDateInput"] input::placeholder {
        color: #64748B !important;
        opacity: 1 !important;
    }
    [data-testid="stDateInput"] svg,
    [data-testid="stSelectbox"] svg {
        color: #1F3864 !important;
        fill: #1F3864 !important;
    }
    [data-testid="stMetric"] {
        background: #FFFFFF !important;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 12px 14px;
    }
    [data-testid="stMetricValue"],
    [data-testid="stMetricLabel"],
    [data-testid="stMetricLabel"] p {
        color: #0F172A !important;
    }
    .stButton > button { background: linear-gradient(135deg, #1F3864, #2E75B6) !important; color: white !important; border: none !important; border-radius: 8px !important; padding: 12px 32px !important; font-size: 14px !important; font-weight: 600 !important; width: 100% !important; }
    .download-box { background: #F0FDF4; border: 2px solid #86EFAC; border-radius: 12px; padding: 28px; text-align: center; margin: 20px 0; }
    .download-title { font-size: 20px; font-weight: 700; color: #166534; margin-bottom: 8px; }
    .download-sub { font-size: 13px; color: #15803D; }
    .detail-box {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-left: 4px solid #2E75B6;
        border-radius: 8px;
        padding: 12px 14px;
        min-height: 72px;
        color: #0F172A;
        font-size: 12px;
        line-height: 1.5;
    }
    .detail-label {
        color: #64748B;
        font-size: 11px;
        font-weight: 700;
        text-transform: uppercase;
        margin-bottom: 4px;
    }
    .detail-value {
        color: #0F172A;
        font-size: 13px;
        font-weight: 600;
        overflow-wrap: anywhere;
    }
</style>
""", unsafe_allow_html=True)

for key, default in [
    ('step',1),
    ('form_data',{}),
    ('uploaded_df',None),
    ('excel_bytes',None),
    ('pdf_bytes',None),
    ('image_bytes',None),
    ('parsed_report',None),
    ('parsed_kpis',None),
]:
    if key not in st.session_state:
        st.session_state[key] = default

if "saved_sprint_details" not in st.session_state:
    st.session_state.saved_sprint_details = _load_saved_sprint_details()

if "status_mappings" not in st.session_state:
    st.session_state.status_mappings = _load_status_mappings()

if "project_selector" not in st.session_state:
    st.session_state.project_selector = PROJECTS[0]

if "active_project" not in st.session_state:
    st.session_state.active_project = st.session_state.project_selector
    _hydrate_project_form(st.session_state.active_project)

# SIDEBAR - Status Mapping Reference
STATUS_MAPPING_REFERENCE = {
    "WMP": [
        ("To Do",              "Not Initiated",  "Not Initiated %"),
        ("Grooming Completed", "Not Initiated",  "Not Initiated %"),
        ("In Progress",        "In Progress",    "Pending %"),
        ("Staging Deployed",   "Staging",        "Pending %"),
        ("QA",                 "Completed - QA", "Completion - QA %"),
        ("Done",               "Production",     "Production Release %"),
    ],
    "SATOC": [
        ("To Do",          "Not Initiated",  "Not Initiated %"),
        ("In Progress",    "In Progress",    "Pending %"),
        ("Stage Deployed", "Staging",        "Completion - QA %"),
        ("QA Review",      "QA Review",      "Completion - QA %"),
        ("Done",           "Production",     "Production Release %"),
    ],
    "PreScreening.io": [
        ("Grooming Completed", "Not Initiated",  "Not Initiated %"),
        ("To Do",              "Not Initiated",  "Not Initiated %"),
        ("In Progress",        "In Progress",    "Pending %"),
        ("Stage Deployed",     "Staging",        "Pending %"),
        ("QA Deployed",        "Completed - QA", "Completion - QA %"),
        ("Done",               "Production",     "Production Release %"),
    ],
}

BUCKET_COLORS = {
    "Not Initiated":  "#ED7D31",
    "In Progress":    "#00B0F0",
    "Staging":        "#BF8F00",
    "QA Review":      "#FFC000",
    "Completed - QA": "#00B050",
    "Production":     "#375623",
}

PCT_COLORS = {
    "Not Initiated %":    "#ED7D31",
    "Pending %":          "#FFC000",
    "Completion - QA %":  "#00B050",
    "Production Release %": "#375623",
}

with st.sidebar:
    st.radio(
        "Page",
        ["Generate Report", "Status Mapping Settings"],
        key="page",
    )
    _persistent = store.backend_kind() == "ZohoSheetBackend"
    st.caption(
        "💾 Storage: **Zoho Sheet** (saved across restarts)" if _persistent
        else "💾 Storage: **local files** (not persistent on Streamlit Cloud)"
    )
    st.markdown("---")
    st.caption("Configure how Jira statuses fill the report on the **Status Mapping Settings** page.")

st.markdown('<div class="header-banner"><h1>Sprint Report Generator</h1><p>Fill in sprint details - Upload your Jira CSV - Download formatted Excel and PDF reports</p></div>', unsafe_allow_html=True)

if st.session_state.get("_store_error"):
    st.warning(st.session_state["_store_error"])

if st.session_state.get("page") == "Status Mapping Settings":
    render_settings_page()
    st.stop()

step = st.session_state.step


def _go_to_step(target_step: int) -> None:
    if target_step == 1:
        st.session_state.step = 1
    elif target_step == 2 and st.session_state.form_data:
        st.session_state.step = 2
    elif target_step == 3 and st.session_state.uploaded_df is not None:
        st.session_state.step = 3

# Jira CSV export guide
@st.dialog("How to Export Your Jira CSV", width="large")
def show_jira_guide():
    st.markdown("""
<style>
.guide-step {
    display:flex; gap:16px; align-items:flex-start;
    background:white; border-radius:10px; padding:16px 18px;
    margin-bottom:12px; box-shadow:0 1px 4px rgba(0,0,0,0.07);
    border-left: 4px solid #2E75B6;
}
.guide-num {
    background:#1F3864; color:white; border-radius:50%;
    width:28px; height:28px; min-width:28px;
    display:flex; align-items:center; justify-content:center;
    font-size:13px; font-weight:700;
}
.guide-body { flex:1; }
.guide-title { font-size:14px; font-weight:700; color:#1F3864; margin-bottom:4px; }
.guide-desc  { font-size:12px; color:#374151; line-height:1.6; }
.jql-box {
    background:#1E293B; color:#7DD3FC; font-family:monospace;
    font-size:12px; padding:10px 14px; border-radius:6px;
    margin-top:8px; word-break:break-all;
}
.tip-box {
    background:#EFF6FF; border-left:3px solid #2E75B6;
    padding:10px 14px; border-radius:0 6px 6px 0;
    font-size:12px; color:#1E40AF; margin-top:12px;
}
</style>

<div class="guide-step">
  <div class="guide-num">1</div>
  <div class="guide-body">
    <div class="guide-title">Open your Jira Project Board</div>
    <div class="guide-desc">Go to <b>Jira</b> and navigate to the relevant project (e.g. <b>WMP -> Data Management &amp; Scraping</b>).</div>
  </div>
</div>

<div class="guide-step">
  <div class="guide-num">2</div>
  <div class="guide-body">
    <div class="guide-title">Click on the "All Work" Tab</div>
    <div class="guide-desc">In the top navigation of your project board, click on <b>All work</b>.</div>
  </div>
</div>

<div class="guide-step">
  <div class="guide-num">3</div>
  <div class="guide-body">
    <div class="guide-title">Open Filters -> Switch to JQL</div>
    <div class="guide-desc">Click the <b>Filter</b> button in the top bar. In the filter panel, switch to the <b>JQL</b> tab.</div>
  </div>
</div>

<div class="guide-step">
  <div class="guide-num">4</div>
  <div class="guide-body">
    <div class="guide-title">Enter the JQL Query</div>
    <div class="guide-desc">Type the following JQL - replace the project key and sprint name with your sprint's details:</div>
    <div class="jql-box">project = WM AND sprint = "WM Sprint 29" ORDER BY issuetype ASC, priority DESC</div>
    <div class="guide-desc" style="margin-top:8px;">
        &bull; <b>project</b> = your Jira project key (e.g. <code>WM</code>, <code>SAT</code>, <code>PS</code>)<br>
        &bull; <b>sprint</b> = exact sprint name as it appears in Jira<br>
        &bull; Keep the <code>ORDER BY</code> clause as-is for best results
    </div>
  </div>
</div>

<div class="guide-step">
  <div class="guide-num">5</div>
  <div class="guide-body">
    <div class="guide-title">Export -> CSV (All Fields)</div>
    <div class="guide-desc">Click the <b>more</b> menu icon on the top-right of the work list  hover over <b>Export</b>  select <b>CSV - all fields</b>.</div>
  </div>
</div>

<div class="guide-step">
  <div class="guide-num">6</div>
  <div class="guide-body">
    <div class="guide-title">Upload the CSV Here</div>
    <div class="guide-desc">Come back to this app, complete <b>Step 1</b> (Sprint Details), then upload the downloaded CSV file in <b>Step 2</b>.</div>
  </div>
</div>

<div class="tip-box">
<b>Tip:</b> Always use <b>CSV - all fields</b> (not "my defaults") to ensure all required columns like
<code>Parent key</code>, <code>Custom field (Target start)</code>, and <code>Comment</code> are included.
</div>
""", unsafe_allow_html=True)

nav1, nav2, nav3, help_col = st.columns([1, 1, 1, 0.18])
with nav1:
    st.button(
        "Sprint Details",
        key="nav_step_1",
        disabled=(step == 1),
        use_container_width=True,
        on_click=_go_to_step,
        args=(1,),
    )
with nav2:
    st.button(
        "Upload Jira CSV",
        key="nav_step_2",
        disabled=(step == 2 or not st.session_state.form_data),
        use_container_width=True,
        on_click=_go_to_step,
        args=(2,),
    )
with nav3:
    st.button(
        "Download Report",
        key="nav_step_3",
        disabled=(step == 3 or st.session_state.uploaded_df is None),
        use_container_width=True,
        on_click=_go_to_step,
        args=(3,),
    )
with help_col:
    if st.button("?", help="How to export Jira CSV", use_container_width=True):
        show_jira_guide()

# STEP 1

if step == 1:
    st.markdown('<div class="section-title">Sprint Information</div>', unsafe_allow_html=True)

    project_name = st.selectbox(
        "Project",
        _project_options(),
        key="project_selector",
        on_change=_load_selected_project_details,
        accept_new_options=True,
        help="Pick a project, or type a new name and choose the \"Add ...\" option to create it.",
    )

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        sprint_number = st.number_input(
            "Sprint Number",
            min_value=1,
            max_value=999,
            step=1,
            key="sprint_number_input",
        )
        sprint_start = st.date_input("Sprint Start Date", key="sprint_start_input")
    with c2:
        dev_release = st.date_input("Sprint Development Release", key="dev_release_input")
        qa_release = st.date_input("Sprint QA Release", key="qa_release_input")
    with c3:
        prod_release = st.date_input("Production Release Date", key="prod_release_input")
        sprint_end = st.date_input("Sprint End Date", key="sprint_end_input")
    with c4:
        scrum_master = st.text_input(
            "Scrum Master",
            placeholder="e.g. Rishav Kumar",
            key="scrum_master_input",
        )

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Auto-Calculated (live preview)</div>', unsafe_allow_html=True)
    total_days = (sprint_end - sprint_start).days + 1
    days_left = max((sprint_end - _today_ist()).days + 1, 0)
    ac1, ac2, ac3 = st.columns(3)
    ac1.metric("Total No. of Days", total_days)
    ac2.metric("Days Left in Sprint", days_left)
    ac3.metric("Sprint End Date", sprint_end.strftime("%d %b %Y"))
    st.markdown('<div class="info-pill">Total Days = Sprint End - Sprint Start + 1 (inclusive). Days Left = Sprint End - Today + 1.</div>', unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Sprint Goal & Major Items</div>', unsafe_allow_html=True)
    sprint_goal = st.text_input(
        "Sprint Goal",
        placeholder="e.g. Rule Engine Module Enhancements (Phase 1)",
        key="sprint_goal_input",
    )
    mg1, mg2, mg3 = st.columns(3)
    with mg1:
        major1 = st.text_input("Major Sprint Item 1", placeholder="Item 1", key="major_item_1_input")
    with mg2:
        major2 = st.text_input("Major Sprint Item 2", placeholder="Item 2", key="major_item_2_input")
    with mg3:
        major3 = st.text_input("Major Sprint Item 3", placeholder="Item 3", key="major_item_3_input")

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    if st.button("Next -> Upload Jira CSV"):
        errors = []
        if not scrum_master.strip():
            errors.append("Scrum Master name is required.")
        if dev_release < sprint_start:
            errors.append("Dev Release cannot be before Sprint Start.")
        if qa_release < dev_release:
            errors.append("QA Release cannot be before Dev Release.")
        if prod_release < qa_release:
            errors.append("Production Release cannot be before QA Release.")
        if sprint_end < sprint_start:
            errors.append("Sprint End Date cannot be before Sprint Start.")

        if errors:
            for e in errors:
                st.markdown(f'<div class="val-error">{e}</div>', unsafe_allow_html=True)
        else:
            form_data = dict(
                sprint_number=sprint_number,
                sprint_start=sprint_start,
                dev_release=dev_release,
                qa_release=qa_release,
                prod_release=prod_release,
                sprint_end=sprint_end,
                total_days=total_days,
                days_left=days_left,
                scrum_master=scrum_master.strip(),
                sprint_goal=sprint_goal.strip(),
                major_item_1=major1.strip(),
                major_item_2=major2.strip(),
                major_item_3=major3.strip(),
                project_name=project_name,
            )
            _save_project_form_data(project_name, form_data)
            st.session_state.form_data = form_data
            st.session_state.last_saved_project = project_name
            st.session_state.uploaded_df = None
            st.session_state.excel_bytes = None
            st.session_state.pdf_bytes = None
            st.session_state.image_bytes = None
            st.session_state.parsed_report = None
            st.session_state.parsed_kpis = None
            st.session_state.step = 2
            st.rerun()

elif step == 2:
    fd = st.session_state.form_data
    st.markdown('<div class="section-title">Sprint Details Confirmed</div>', unsafe_allow_html=True)
    sc = st.columns(5)
    sc[0].metric("Sprint",       f"#{fd['sprint_number']}")
    sc[1].metric("Start",        fd['sprint_start'].strftime("%d %b %Y"))
    sc[2].metric("Prod Release", fd['prod_release'].strftime("%d %b %Y"))
    sc[3].metric("Total Days",   fd['total_days'])
    sc[4].metric("Scrum Master", fd['scrum_master'])

    proj = fd['project_name']
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Get Jira Data</div>', unsafe_allow_html=True)

    # ---- Option 1: fetch directly from Jira (if configured) ----
    if jira_client.is_configured():
        saved_q = _load_store("jira_config").get(proj, {})
        with st.expander("🔗 Fetch from Jira (saved filter or JQL)", expanded=bool(saved_q)):
            mode_label = st.radio(
                "Query type", ["Saved filter ID", "JQL"],
                index=0 if saved_q.get("mode", "filter") == "filter" else 1,
                horizontal=True, key=f"jira_mode_{proj}",
            )
            value = st.text_input(
                "Filter ID or JQL", value=saved_q.get("value", ""), key=f"jira_value_{proj}",
                placeholder="e.g.  12345   (filter id)   —or—   project = WMP AND sprint in openSprints()",
            )
            query = {"mode": "filter" if mode_label == "Saved filter ID" else "jql", "value": value}
            fc1, fc2 = st.columns([1, 1])
            with fc1:
                if st.button("Fetch from Jira", type="primary", use_container_width=True,
                             disabled=not value.strip()):
                    _save_to_store("jira_config", proj, query)  # remember for next time
                    try:
                        with st.spinner("Fetching issues from Jira..."):
                            jdf = jira_client.fetch_issues(query)
                        if _ingest_dataframe(jdf, f"Jira • {len(jdf)} issues"):
                            st.success(f"Fetched {len(jdf)} issues from Jira.")
                    except Exception as exc:
                        st.error(f"Jira fetch failed: {exc}")
            with fc2:
                if st.button("Save query", use_container_width=True, disabled=not value.strip()):
                    _save_to_store("jira_config", proj, query)
                    st.success("Saved - this query is remembered for this project.")
    else:
        st.caption("💡 Add a `[jira]` section to Streamlit secrets to enable one-click **Fetch from Jira**.")

    # ---- Option 2: upload a CSV export (always available) ----
    st.markdown('<div class="info-pill">Or upload a Jira CSV export (all fields). Required: <b>Issue key, Issue Type, Summary, Status</b>.</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Drop your Jira CSV here", type=["csv"], label_visibility="collapsed")
    if uploaded_file:
        try:
            _ingest_dataframe(pd.read_csv(uploaded_file), uploaded_file.name)
        except Exception as e:
            st.error(f"Could not read CSV: {e}")

    # ---- Shared preview (whichever source loaded the data) ----
    df = st.session_state.uploaded_df
    if df is not None:
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.success(f"Loaded: **{st.session_state.get('uploaded_source', 'data')}** — {len(df)} rows")
        epics   = len(df[df['Issue Type'] == 'Epic'])
        stories = len(df[df['Issue Type'].isin(['Story', 'Task'])])
        subs    = len(df[df['Issue Type'] == 'Sub-task'])
        pc = st.columns(4)
        pc[0].metric("Total Rows", len(df)); pc[1].metric("Epics", epics)
        pc[2].metric("Stories/Tasks", stories); pc[3].metric("Sub-tasks", subs)

        st.markdown('<div class="section-title">Status Breakdown Preview</div>', unsafe_allow_html=True)
        non_epic = df[df['Issue Type'] != 'Epic']
        sdf = non_epic['Status'].value_counts().reset_index()
        sdf.columns = ['Status', 'Count']
        st.dataframe(sdf, use_container_width=True, hide_index=True)

        # Warn only about genuinely new statuses (not deliberately ignored ones).
        emap = _effective_report_map(proj)
        if emap is not None:
            known_lower = {
                str(s).lower().strip()
                for s in _effective_status_config(proj).get("known_statuses", [])
            }
            unmapped = sorted({
                str(s).strip() for s in non_epic['Status'].dropna().unique()
                if str(s).strip()
                and str(s).lower().strip() not in emap
                and str(s).lower().strip() not in known_lower
            })
            if unmapped:
                st.warning(
                    f"New Jira statuses not yet mapped for **{proj}** (won't be counted): "
                    f"**{', '.join(unmapped)}**. Map them on the **Status Mapping Settings** page."
                )

        st.markdown('<div class="section-title">Data Preview (first 5 rows)</div>', unsafe_allow_html=True)
        pcols = [c for c in ['Issue key', 'Issue Type', 'Summary', 'Status', 'Priority', 'Assignee', 'Parent key'] if c in df.columns]
        st.dataframe(df[pcols].head(5), use_container_width=True, hide_index=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    b1, b2 = st.columns([1, 3])
    with b1:
        if st.button("<- Back"):
            st.session_state.step = 1; st.rerun()
    with b2:
        if st.button("Generate Excel/PDF Report", disabled=(st.session_state.uploaded_df is None)):
            st.session_state.step = 3; st.rerun()
    if st.session_state.uploaded_df is None:
        st.markdown('<div class="val-error">Fetch from Jira or upload a CSV to continue.</div>', unsafe_allow_html=True)

# STEP 3
elif step == 3:
    fd  = st.session_state.form_data
    df  = st.session_state.uploaded_df

    if st.session_state.excel_bytes is None or st.session_state.pdf_bytes is None or st.session_state.image_bytes is None:
        with st.spinner("Parsing Jira data and building your Excel/PDF/Image report..."):
            parsed = parse_jira_csv(df, fd['project_name'], _resolve_status_config(fd['project_name']))
            st.session_state.excel_bytes = build_excel(fd, parsed)
            st.session_state.pdf_bytes = build_pdf(fd, parsed)
            st.session_state.image_bytes = build_summary_image(fd, parsed)
            st.session_state.parsed_report = parsed
            st.session_state.parsed_kpis = parsed['kpis']

    kpis        = st.session_state.parsed_kpis
    excel_bytes = st.session_state.excel_bytes
    pdf_bytes   = st.session_state.pdf_bytes
    image_bytes = st.session_state.image_bytes
    base_filename = _report_base_name(fd)
    filename = f"{base_filename}.xlsx"
    pdf_filename = f"{base_filename}.pdf"
    image_filename = f"{base_filename}.png"

    st.markdown('<div class="download-box"><div class="download-title">Report Ready</div><div class="download-sub">Your Sprint Report has been generated successfully.</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Sprint Details</div>', unsafe_allow_html=True)
    detail_items = [
        ("Project", fd['project_name']),
        ("Sprint", f"#{fd['sprint_number']}"),
        ("Start Date", fd['sprint_start'].strftime("%d %b %Y")),
        ("Dev Release", fd['dev_release'].strftime("%d %b %Y")),
        ("QA Release", fd['qa_release'].strftime("%d %b %Y")),
        ("Production Release", fd['prod_release'].strftime("%d %b %Y")),
        ("Sprint End", fd['sprint_end'].strftime("%d %b %Y")),
        ("Scrum Master", fd['scrum_master']),
    ]
    for offset in range(0, len(detail_items), 4):
        detail_cols = st.columns(4)
        for col, (label, value) in zip(detail_cols, detail_items[offset:offset + 4]):
            col.markdown(
                f'<div class="detail-box"><div class="detail-label">{_html_escape(label)}</div>'
                f'<div class="detail-value">{_html_escape(value)}</div></div>',
                unsafe_allow_html=True,
            )

    goal_cols = st.columns([1, 2])
    goal_cols[0].markdown(
        f'<div class="detail-box"><div class="detail-label">Sprint Goal</div>'
        f'<div class="detail-value">{_html_escape(fd.get("sprint_goal", "") or "-")}</div></div>',
        unsafe_allow_html=True,
    )
    major_items = [fd.get('major_item_1', ''), fd.get('major_item_2', ''), fd.get('major_item_3', '')]
    major_text = "<br>".join(_html_escape(item) for item in major_items if item) or "-"
    goal_cols[1].markdown(
        f'<div class="detail-box"><div class="detail-label">Major Sprint Items</div>'
        f'<div class="detail-value">{major_text}</div></div>',
        unsafe_allow_html=True,
    )

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Sprint KPI Summary</div>', unsafe_allow_html=True)
    k = st.columns(5)
    k[0].metric("Sprint",          f"#{fd['sprint_number']}")
    k[1].metric("Action Items",    kpis['action_items'])
    k[2].metric("Pending %",       kpis['pending_pct'])
    k[3].metric("Not Initiated %", kpis['not_initiated_pct'])
    k[4].metric("Days Left",       fd['days_left'])

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Status Breakdown</div>', unsafe_allow_html=True)

    status_items = [
        ("Not Initiated",  kpis['not_initiated'],  "#ED7D31"),
        ("In Progress",    kpis['in_progress'],    "#00B0F0"),
        ("Staging",        kpis['staging'],        "#BF8F00"),
        ("QA Review",      kpis['qa_review'],      "#FFC000"),
        ("QA Deployed",    kpis['qa_deployed'],    "#70AD47"),
        ("QA Approved",    kpis['qa_approved'],    "#00B050"),
        ("Production",     kpis['production'],     "#375623"),
        ("On Hold",        kpis['on_hold'],        "#A6A6A6"),
        ("Another Sprint", kpis['to_be_picked'],   "#7030A0"),
    ]
    sc = st.columns(5)
    for i, (label, val, color) in enumerate(status_items):
        sc[i % 5].markdown(f"""
        <div style="background:white;border-radius:8px;padding:14px;
                    border-left:4px solid {color};margin-bottom:8px;
                    box-shadow:0 1px 3px rgba(0,0,0,0.08);">
            <div style="font-size:11px;color:#64748B;font-weight:600;">{label}</div>
            <div style="font-size:24px;font-weight:700;color:{color};">{val}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Summary Image Preview</div>', unsafe_allow_html=True)
    st.image(image_bytes, use_container_width=True)

    download_files = {
        "Excel": (filename, excel_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        "PDF": (pdf_filename, pdf_bytes, "application/pdf"),
        "Image": (image_filename, image_bytes, "image/png"),
    }
    selected_formats = st.multiselect(
        "Select files to download",
        options=list(download_files.keys()),
        default=list(download_files.keys()),
    )
    if selected_formats:
        _render_multi_file_download_button([download_files[option] for option in selected_formats])
    else:
        st.markdown('<div class="val-error">Select at least one file type to download.</div>', unsafe_allow_html=True)

    # ---- Post to Zoho Cliq channel ----
    if cliq_client.is_configured():
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Post to Zoho Cliq</div>', unsafe_allow_html=True)
        saved_channel = _load_store("zoho_channels").get(fd['project_name'], "")
        ch1, ch2 = st.columns([3, 1])
        with ch1:
            channel = st.text_input(
                "Cliq channel (unique name)", value=saved_channel,
                key=f"cliq_channel_{fd['project_name']}",
                placeholder="e.g. wmp-sprint-reports",
                help="The channel's unique name from its URL/settings. Saved per project.",
            )
        with ch2:
            st.write("")
            if st.button("Save channel", use_container_width=True, disabled=not channel.strip()):
                _save_to_store("zoho_channels", fd['project_name'], channel.strip())
                st.success("Channel saved.")
        if st.button("Post report to Cliq", type="primary", disabled=not channel.strip()):
            _save_to_store("zoho_channels", fd['project_name'], channel.strip())
            message = (
                f"*Sprint {fd['sprint_number']} report — {fd['project_name']}*\n"
                f"Action Items: {kpis['action_items']} | Pending: {kpis['pending_pct']} | "
                f"Not Initiated: {kpis['not_initiated_pct']} | Production: {kpis['production_release_pct']}\n"
                f"Scrum Master: {fd['scrum_master']}"
            )
            files = [download_files["Excel"], download_files["PDF"], download_files["Image"]]
            try:
                with st.spinner("Posting to Zoho Cliq..."):
                    warnings = cliq_client.post_report(channel.strip(), files, message)
                if warnings:
                    st.warning("Message posted, but some files failed:\n\n" + "\n\n".join(warnings))
                else:
                    st.success(f"Posted the report (message + 3 files) to #{channel.strip()}.")
            except Exception as exc:
                st.error(f"Could not post to Cliq: {exc}")

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    _, reset_col = st.columns([2, 1])
    with reset_col:
        if st.button("Generate Another Report", key="report_generate_another", use_container_width=True):
            for key in ['form_data','uploaded_df','excel_bytes','pdf_bytes','image_bytes','parsed_report','parsed_kpis']:
                st.session_state.pop(key, None)
            st.session_state.step = 1
            _hydrate_project_form(st.session_state.get("project_selector", PROJECTS[0]))
            st.rerun()

# CHANGELOG - visible on every step
CHANGELOG = [
    {
        "version": "v1.5",
        "date":    "27 Mar 2026",
        "commit":  "c9e4117",
        "title":   "Project-Specific Jira Status Mapping & Sidebar Reference",
        "tag":     "Feature",
        "tag_color": "#2E75B6",
        "changes": [
            "Added project-specific Jira -> Sprint Sheet status mapping for WMP, SATOC, and PreScreening.io",
            "Fixed WMP 'Grooming Completed' typo - 59 items were unmapped causing % to not sum to 100%",
            "Staging Deployed now maps to Staging -> Pending % for WMP",
            "Added Completed-QA (absolute count) and Completion-QA% columns to the KPI row in Excel",
            "Added sidebar hamburger menu showing color-coded Jira -> Sprint Sheet status reference per project",
        ],
    },
    {
        "version": "v1.4",
        "date":    "25 Feb 2026",
        "commit":  "2db237a",
        "title":   "S.No Column, Renamed Columns & Revised Date Columns",
        "tag":     "Feature",
        "tag_color": "#2E75B6",
        "changes": [
            "Added S.No (serial number) column to the task table in Excel output",
            "Renamed columns for better clarity in the sprint sheet",
            "Added Revised Target Start and Revised Target End date columns",
        ],
    },
    {
        "version": "v1.3",
        "date":    "25 Feb 2026",
        "commit":  "c1cdb76",
        "title":   "Project Dropdown & IST-Based Filename",
        "tag":     "Feature",
        "tag_color": "#2E75B6",
        "changes": [
            "Added project selector dropdown (WMP, SATOC, PreScreening.io, etc.)",
            "Excel filename now auto-generated using project name + current IST date",
        ],
    },
    {
        "version": "v1.2",
        "date":    "23 Feb 2026",
        "commit":  "3818aa4",
        "title":   "Excel Colour Scheme & Formatting Overhaul",
        "tag":     "UI",
        "tag_color": "#7030A0",
        "changes": [
            "Updated Excel output with a new colour scheme matching sprint sheet standards",
            "Improved cell formatting, font weights, and header row styling",
        ],
    },
    {
        "version": "v1.1",
        "date":    "20 Feb 2026",
        "commit":  "3d7ad77",
        "title":   "Comment Column, Hierarchy Fix & Sprint End Date Input",
        "tag":     "Fix",
        "tag_color": "#C00000",
        "changes": [
            "Added latest Jira comment column to the task table",
            "Fixed Epic -> Story -> Sub-task hierarchy logic (external epics & orphan sub-tasks now handled)",
            "Added Sprint End Date as a separate input field",
        ],
    },
    {
        "version": "v1.0",
        "date":    "19 Feb 2026",
        "commit":  "98165e7",
        "title":   "Initial Release",
        "tag":     "Launch",
        "tag_color": "#00B050",
        "changes": [
            "Sprint Report Generator launched",
            "3-step UI: Sprint Details -> Upload Jira CSV -> Download Excel Report",
            "Auto-calculated KPIs: Action Items, Pending %, Not Initiated %, Production Release %",
            "Jira CSV parsed into Epic -> Story/Task -> Sub-task hierarchy in Excel",
        ],
    },
]

st.markdown("---")
with st.expander("Changelog - What's New", expanded=False):
    html_parts = []
    for i, entry in enumerate(CHANGELOG):
        latest = '<span style="background:#FFD700;color:#1F3864;font-size:10px;font-weight:700;padding:2px 8px;border-radius:10px;margin-left:8px;">LATEST</span>' if i == 0 else ''
        bullets = "".join(f'<li style="font-size:12px;color:#374151;margin-bottom:4px;">{_html_escape(c)}</li>' for c in entry["changes"])
        html_parts.append(
            f'<div style="background:white;border-radius:10px;padding:16px 20px;margin-bottom:12px;'
            f'box-shadow:0 1px 4px rgba(0,0,0,0.08);border-left:4px solid {entry["tag_color"]};">'
            f'<p style="font-size:15px;font-weight:700;color:#1F3864;margin:0 0 8px 0;">'
            f'{_html_escape(entry["version"])} - {_html_escape(entry["title"])}{latest}</p>'
            f'<p style="margin:0 0 10px 0;">'
            f'<span style="background:{entry["tag_color"]};color:white;font-size:11px;font-weight:600;padding:2px 10px;border-radius:10px;">{_html_escape(entry["tag"])}</span>&nbsp;'
            f'<span style="background:#F1F5F9;color:#64748B;font-size:11px;padding:2px 10px;border-radius:10px;">&#128197; {_html_escape(entry["date"])}</span>&nbsp;'
            f'<span style="background:#F1F5F9;color:#64748B;font-size:11px;padding:2px 10px;border-radius:10px;font-family:monospace;">#{_html_escape(entry["commit"])}</span>'
            f'</p>'
            f'<ul style="margin:0;padding-left:18px;">{bullets}</ul>'
            f'</div>'
        )
    st.markdown("".join(html_parts), unsafe_allow_html=True)


