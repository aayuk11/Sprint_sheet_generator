"""
app.py - Sprint Report Generator
Run with: streamlit run app.py

Two pages (sidebar):
  - Generate Report: 3-step wizard (Sprint Details -> Upload/Fetch Jira -> Download)
  - Settings: per-project configuration - status buckets, % KPI cards, their
    colors, and the Jira-status -> bucket mapping. Nothing about the buckets
    or KPIs is fixed; a project can have any number of each.
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

from modules.parser import parse_jira_csv, distinct_statuses, STORY_LEVEL_TYPES
from modules.excel_generator import build_excel
from modules.pdf_generator import build_pdf
from modules.image_generator import build_summary_image
from modules import store
from modules import jira_client
from modules import cliq_client
from modules import columns

PROJECTS = [
    "Product Design",
    "PreScreening.io", "Transact Comply", "Entity Hero", "DueDiliger",
    "ZiZi", "SATOC", "WMP", "Fraud Fighter", "Profile Builder",
]
DESIGN_PROJECT = "Product Design"

PALETTE = ['#ED7D31', '#00B0F0', '#BF8F00', '#FFC000', '#70AD47', '#00B050',
           '#375623', '#A6A6A6', '#7030A0', '#2E75B6', '#C00000', '#1F4E79']


def _palette_color(idx: int) -> str:
    return PALETTE[idx % len(PALETTE)]


def _today_ist() -> date:
    return datetime.now(ZoneInfo("Asia/Kolkata")).date()


def _load_store(name: str) -> dict:
    try:
        return store.get_backend().load_all(name)
    except Exception as exc:
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


# ---- Sprint detail persistence (per project) --------------------------------

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
    saved = st.session_state.saved_sprint_details.get(project_name, {})
    default_type = "Design Sprint" if project_name == DESIGN_PROJECT else "Product Sprint"
    return {
        "sprint_type_input": saved.get("sprint_type", default_type),
        "sprint_number_input": int(saved.get("sprint_number", 1) or 1),
        "sprint_start_input": _parse_saved_date(saved.get("sprint_start"), today),
        "dev_release_input": _parse_saved_date(saved.get("dev_release"), today),
        "qa_release_input": _parse_saved_date(saved.get("qa_release"), today),
        "prod_release_input": _parse_saved_date(saved.get("prod_release"), today),
        "sprint_end_input": _parse_saved_date(saved.get("sprint_end"), today),
        "sprint_release_input": _parse_saved_date(saved.get("sprint_release"), today),
        "scrum_master_input": saved.get("scrum_master", ""),
        "sprint_goal_input": saved.get("sprint_goal", ""),
        "major_item_1_input": saved.get("major_item_1", ""),
        "major_item_2_input": saved.get("major_item_2", ""),
        "major_item_3_input": saved.get("major_item_3", ""),
    }


def _sprint_type_for_project(project_name: str) -> str:
    """The sprint type a project is normally run as - the saved one if any,
    else Design for the design project and Product for everything else.
    Used to pick default column order in Settings."""
    saved = st.session_state.get("saved_sprint_details", {}).get(project_name, {})
    default_type = "Design Sprint" if project_name == DESIGN_PROJECT else "Product Sprint"
    return saved.get("sprint_type", default_type)


def _hydrate_project_form(project_name: str) -> None:
    for key, value in _project_defaults(project_name).items():
        st.session_state[key] = value


def _on_sprint_type_change() -> None:
    # _hydrate_project_form() also resets sprint_type_input from saved data,
    # which would immediately undo the switch the user just made - so capture
    # the chosen type first and re-apply it after hydrating the other fields.
    chosen_type = st.session_state.sprint_type_input
    if chosen_type == "Design Sprint":
        _hydrate_project_form(DESIGN_PROJECT)
    else:
        _hydrate_project_form(st.session_state.get("project_selector", PROJECTS[0]))
    st.session_state.sprint_type_input = chosen_type


def _project_options() -> list:
    options = list(PROJECTS)
    for name in list(st.session_state.get("saved_sprint_details", {})) + list(st.session_state.get("project_configs", {})):
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


def _serialize_form_data(form_data: dict) -> dict:
    """Generic: isoformat any date fields present. Product and Design sprints
    have different field sets (e.g. only one has dev_release/qa_release/
    prod_release), so this doesn't hardcode which keys exist."""
    serialized = {}
    for k, v in form_data.items():
        if k in ("meta_items", "label_header", "task_columns"):
            continue  # display-only, rebuilt fresh each time a report is generated
        serialized[k] = v.isoformat() if isinstance(v, date) else v
    return serialized


def _build_meta_items(fd: dict) -> list:
    """The Sprint Summary meta row shown in the report (Excel/PDF/Image) and
    the Step 3 detail cards - differs by sprint type."""
    if fd.get("sprint_type") == "Design Sprint":
        return [
            ("Project", fd["project_name"], "#000000"),
            ("Sprint Number", fd["sprint_number"], "#1F4E79"),
            ("Sprint Start Date", fd["sprint_start"].strftime("%d %b %Y"), "#2E75B6"),
            ("Sprint End Date", fd["sprint_end"].strftime("%d %b %Y"), "#375623"),
            ("Sprint Release Date", fd["sprint_release"].strftime("%d %b %Y"), "#C00000"),
            ("Total No. of Days", fd["total_days"], "#595959"),
            ("Scrum Master", fd["scrum_master"], "#7030A0"),
        ]
    return [
        ("Project", fd["project_name"], "#000000"),
        ("Sprint Number", fd["sprint_number"], "#1F4E79"),
        ("Sprint Start Date", fd["sprint_start"].strftime("%d %b %Y"), "#2E75B6"),
        ("Sprint Development Release", fd["dev_release"].strftime("%d %b %Y"), "#00B0F0"),
        ("Sprint QA Release", fd["qa_release"].strftime("%d %b %Y"), "#BF8F00"),
        ("Production Release", fd["prod_release"].strftime("%d %b %Y"), "#C00000"),
        ("Sprint End Date", fd["sprint_end"].strftime("%d %b %Y"), "#375623"),
        ("Total No. of Days", fd["total_days"], "#595959"),
        ("Scrum Master", fd["scrum_master"], "#7030A0"),
    ]


def _save_project_form_data(project_name: str, form_data: dict) -> None:
    serialized = _serialize_form_data(form_data)
    saved = st.session_state.saved_sprint_details.copy()
    saved[project_name] = serialized
    st.session_state.saved_sprint_details = saved
    _save_to_store("sprint_details", project_name, serialized)


def _invalidate_generated_report() -> None:
    for key in ("excel_bytes", "pdf_bytes", "image_bytes", "parsed_report"):
        st.session_state[key] = None


def _ingest_dataframe(df, source_label: str) -> bool:
    missing = [c for c in ["Issue key", "Issue Type", "Summary", "Status"] if c not in df.columns]
    if missing:
        st.error(f"Missing required columns: {', '.join(missing)}")
        return False
    st.session_state.uploaded_df = df
    st.session_state.uploaded_source = source_label
    _invalidate_generated_report()
    return True


# ---- Project config (buckets / kpis / status map) ---------------------------

def _load_project_configs() -> dict:
    return _load_store("project_configs")


# ---- One-time migration: old fixed-bucket status_mappings -> project_configs -
# The previous version used 9 fixed buckets + 4 fixed % KPIs. Convert any saved
# "status_mappings" into the new configurable "project_configs" format so existing
# projects keep their mapping (labels/colors below match the old report output).
_LEGACY_BUCKETS = [
    ("not_initiated", "Not Initiated",              "#ED7D31"),
    ("in_progress",   "In Progress",                "#00B0F0"),
    ("staging",       "Staging",                    "#BF8F00"),
    ("qa_review",     "QA Review",                  "#FFC000"),
    ("qa_deployed",   "QA Deployed",                "#70AD47"),
    ("qa_approved",   "QA Approved (Completed-QA)", "#00B050"),
    ("production",    "Production",                 "#375623"),
    ("on_hold",       "On Hold",                    "#A6A6A6"),
    ("to_be_picked",  "To Be Picked (Another Sprint)", "#7030A0"),
]
_LEGACY_PCTS = [
    ("not_initiated_pct",      "Not Initiated %",      "#ED7D31", ["not_initiated"]),
    ("pending_pct",            "Pending %",            "#FFC000", ["in_progress", "staging"]),
    ("completion_qa_pct",      "Completion - QA %",    "#00B050", ["qa_review", "qa_deployed", "qa_approved"]),
    ("production_release_pct", "Production Release %", "#375623", ["production"]),
]


def _migrate_one_status_mapping(old: dict) -> dict:
    pct_buckets = old.get("pct_buckets") or {}
    buckets = [{"key": k, "label": lbl, "color": col} for (k, lbl, col) in _LEGACY_BUCKETS]
    kpis = [
        {"key": pk, "label": lbl, "color": col, "bucket_keys": list(pct_buckets.get(pk, default_bks))}
        for (pk, lbl, col, default_bks) in _LEGACY_PCTS
    ]
    return {
        "buckets": buckets,
        "kpis": kpis,
        "status_map": dict(old.get("status_map") or {}),
        "known_statuses": list(old.get("known_statuses") or []),
    }


def _migrate_legacy_status_mappings() -> None:
    """Run once per session: convert legacy status_mappings -> project_configs
    for any project that doesn't already have a new config."""
    if st.session_state.get("_migrated_status_mappings"):
        return
    st.session_state["_migrated_status_mappings"] = True
    try:
        legacy = store.get_backend().load_all("status_mappings")
    except Exception:
        legacy = {}
    if not legacy:
        return
    configs = dict(st.session_state.get("project_configs", {}))
    changed = False
    for project, old in legacy.items():
        if project in configs:
            continue
        if not isinstance(old, dict) or not old.get("status_map"):
            continue
        new_cfg = _migrate_one_status_mapping(old)
        configs[project] = new_cfg
        _save_to_store("project_configs", project, new_cfg)
        changed = True
    if changed:
        st.session_state.project_configs = configs


def _new_bucket(idx: int, label: str = "", color: str | None = None) -> dict:
    return {"key": f"bucket_{idx}", "label": label, "color": color or _palette_color(idx)}


def _new_kpi(idx: int, label: str = "", color: str | None = None, bucket_keys=None) -> dict:
    return {"key": f"kpi_{idx}", "label": label, "color": color or _palette_color(idx + 3),
            "bucket_keys": list(bucket_keys or [])}


# ---- Built-in defaults (used until a project is explicitly saved in Settings) -
# So existing/unconfigured projects keep the full standard KPIs & buckets - no
# disruption. Editing a project in Settings overrides these.

# Standard Product-sprint status map (mirrors the old keyword defaults).
_PRODUCT_DEFAULT_STATUS_MAP = {
    "to do": "not_initiated", "not initiated": "not_initiated", "open": "not_initiated",
    "in progress": "in_progress",
    "staging deployed": "staging", "staging": "staging", "stage deployed": "staging",
    "qa review": "qa_review", "in review": "qa_review",
    "qa deployed": "qa_deployed",
    "qa approved": "qa_approved", "qa": "qa_approved",
    "done": "production", "production": "production", "released": "production", "closed": "production",
    "on hold": "on_hold", "blocked": "on_hold",
    "to be picked in another sprint": "to_be_picked", "deferred": "to_be_picked",
}

# The three projects whose maps used to be hard-coded - preserved exactly.
_BUILTIN_PROJECT_MAPS = {
    "WMP": {
        "status_map": {"to do": "not_initiated", "grooming completed": "not_initiated",
                       "in progress": "in_progress", "staging deployed": "staging",
                       "qa": "qa_approved", "done": "production"},
        "pct_buckets": {"pending_pct": ["in_progress", "staging"], "not_initiated_pct": ["not_initiated"],
                        "completion_qa_pct": ["qa_approved"], "production_release_pct": ["production"]},
    },
    "SATOC": {
        "status_map": {"to do": "not_initiated", "in progress": "in_progress",
                       "stage deployed": "staging", "qa review": "qa_review", "done": "production"},
        "pct_buckets": {"pending_pct": ["in_progress"], "not_initiated_pct": ["not_initiated"],
                        "completion_qa_pct": ["staging", "qa_review"], "production_release_pct": ["production"]},
    },
    "PreScreening.io": {
        "status_map": {"grooming completed": "not_initiated", "to do": "not_initiated",
                       "in progress": "in_progress", "stage deployed": "staging",
                       "qa deployed": "qa_deployed", "done": "production"},
        "pct_buckets": {"pending_pct": ["in_progress", "staging"], "not_initiated_pct": ["not_initiated"],
                        "completion_qa_pct": ["qa_deployed"], "production_release_pct": ["production"]},
    },
}

# Product Design (Design sprint) default bucket/KPI set.
_DESIGN_DEFAULT = {
    "buckets": [
        {"key": "bucket_0", "label": "Not Initiated",   "color": "#ED7D31"},
        {"key": "bucket_1", "label": "In Progress",     "color": "#f0e600"},
        {"key": "bucket_2", "label": "Completed",       "color": "#3bbf00"},
        {"key": "bucket_3", "label": "Document Pending","color": "#bdbbbb"},
        {"key": "bucket_4", "label": "On Hold",         "color": "#515251"},
        {"key": "bucket_5", "label": "Rework",          "color": "#00b0a6"},
        {"key": "bucket_6", "label": "Wireframe Ready", "color": "#5edecf"},
        {"key": "bucket_7", "label": "In Review",       "color": "#3386d8"},
    ],
    "kpis": [
        {"key": "kpi_0", "label": "Completion %",    "color": "#3bbf00", "bucket_keys": ["bucket_2"]},
        {"key": "kpi_1", "label": "Pending %",       "color": "#e6e200",
         "bucket_keys": ["bucket_1", "bucket_3", "bucket_4", "bucket_5", "bucket_6"]},
        {"key": "kpi_2", "label": "Not Initiated %", "color": "#ed7d31", "bucket_keys": ["bucket_0"]},
        {"key": "kpi_3", "label": "In Review %",     "color": "#3386d8", "bucket_keys": ["bucket_7"]},
    ],
    "status_map": {
        "completed": "bucket_2", "document pending": "bucket_3", "documentation pending": "bucket_3",
        "in progress": "bucket_1", "in review": "bucket_7", "not initiated": "bucket_0",
        "on hold": "bucket_4", "rework": "bucket_5", "wireframe ready": "bucket_6",
    },
    "known_statuses": ["Completed", "Document Pending", "Documentation Pending", "In Progress",
                       "In Review", "Not Initiated", "On Hold", "Rework", "Wireframe Ready"],
}


def _product_default_config() -> dict:
    """Full standard Product-sprint config (old 9 buckets / 4 % KPIs)."""
    known = [s.title() for s in dict.fromkeys(_PRODUCT_DEFAULT_STATUS_MAP)]
    return _migrate_one_status_mapping(
        {"status_map": dict(_PRODUCT_DEFAULT_STATUS_MAP), "pct_buckets": {}, "known_statuses": known}
    )


def _default_config_for(project_name: str) -> dict:
    """The default config for a project that hasn't been saved in Settings yet.
    Design project -> Design set; the 3 legacy projects -> their old maps;
    everything else -> the full standard Product set. Never the empty seed."""
    if project_name == DESIGN_PROJECT:
        cfg = {
            "buckets": [dict(b) for b in _DESIGN_DEFAULT["buckets"]],
            "kpis": [dict(k) for k in _DESIGN_DEFAULT["kpis"]],
            "status_map": dict(_DESIGN_DEFAULT["status_map"]),
            "known_statuses": list(_DESIGN_DEFAULT["known_statuses"]),
        }
    else:
        builtin = _BUILTIN_PROJECT_MAPS.get(project_name)
        if builtin:
            known = [s.title() for s in builtin["status_map"]]
            cfg = _migrate_one_status_mapping({**builtin, "known_statuses": known})
        else:
            cfg = _product_default_config()
    # None = no explicit column order saved; the sprint-type default is used.
    cfg["task_columns"] = None
    return cfg


def _effective_project_config(project_name: str) -> dict:
    cfg = st.session_state.get("project_configs", {}).get(project_name)
    if cfg:
        return {
            "buckets": [dict(b) for b in cfg.get("buckets", [])],
            "kpis": [dict(k) for k in cfg.get("kpis", [])],
            "status_map": dict(cfg.get("status_map", {})),
            "known_statuses": list(cfg.get("known_statuses", [])),
            # None (not a default) when unsaved, so the sprint-type default applies.
            "task_columns": [dict(c) for c in cfg.get("task_columns", [])] or None,
        }
    return _default_config_for(project_name)


def _next_seq(rows: list, prefix: str) -> int:
    nums = []
    for r in rows:
        parts = str(r.get("key", "")).rsplit("_", 1)
        if len(parts) == 2 and parts[0] == prefix and parts[1].isdigit():
            nums.append(int(parts[1]))
    return (max(nums) + 1) if nums else 0


def _ensure_editor(project_name: str) -> dict:
    """Per-project working copy of buckets/kpis, kept in session_state while
    the Settings page is open so add/remove/edit doesn't require re-fetching
    from storage on every widget interaction. Reset explicitly on Save/Reset."""
    if "editor" not in st.session_state:
        st.session_state.editor = {}
    if project_name not in st.session_state.editor:
        cfg = _effective_project_config(project_name)
        st.session_state.editor[project_name] = {
            "buckets": cfg["buckets"],
            "kpis": cfg["kpis"],
            "task_columns": columns.resolve(
                cfg.get("task_columns"), _sprint_type_for_project(project_name)
            ),
        }
    return st.session_state.editor[project_name]


def render_settings_page() -> None:
    st.markdown('<div class="section-title">Project Configuration</div>', unsafe_allow_html=True)
    st.caption("Define the status buckets and % KPI cards for a project, map Jira statuses to them, "
               "and pick colors - all reused automatically every time you generate a report.")

    project = st.selectbox(
        "Project", _project_options(), key="settings_project",
        accept_new_options=True, help="Pick a project, or type a new name to configure it.",
    )

    ed = _ensure_editor(project)

    # ---- Buckets ----
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown("**Status Buckets** — the stage categories counted in your report (e.g. In Review, Done)")
    for i, b in enumerate(list(ed["buckets"])):
        c1, c2, c3 = st.columns([3, 1, 0.4])
        b["label"] = c1.text_input("Label", value=b["label"], key=f"bl_{project}_{b['key']}", label_visibility="collapsed")
        b["color"] = c2.color_picker("Color", value=b["color"], key=f"bc_{project}_{b['key']}", label_visibility="collapsed")
        if c3.button("✕", key=f"bd_{project}_{b['key']}", help="Remove bucket"):
            ed["buckets"] = [x for x in ed["buckets"] if x["key"] != b["key"]]
            st.rerun()
    if st.button("+ Add bucket", key=f"add_bucket_{project}"):
        ed["buckets"].append(_new_bucket(_next_seq(ed["buckets"], "bucket")))
        st.rerun()

    bucket_labels = [b["label"] for b in ed["buckets"] if b["label"].strip()]
    label_to_key = {b["label"]: b["key"] for b in ed["buckets"]}
    key_to_label = {b["key"]: b["label"] for b in ed["buckets"]}

    # ---- KPI cards ----
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown("**% KPI Cards** — the percentage metrics shown at the top of the report")
    for i, k in enumerate(list(ed["kpis"])):
        c1, c2, c3, c4 = st.columns([2.2, 1, 2.6, 0.4])
        k["label"] = c1.text_input("Label", value=k["label"], key=f"kl_{project}_{k['key']}", label_visibility="collapsed")
        k["color"] = c2.color_picker("Color", value=k["color"], key=f"kc_{project}_{k['key']}", label_visibility="collapsed")
        current_labels = [key_to_label.get(bk, bk) for bk in k["bucket_keys"] if key_to_label.get(bk, bk) in bucket_labels]
        chosen = c3.multiselect("Buckets rolled up", bucket_labels, default=current_labels,
                                 key=f"kb_{project}_{k['key']}", label_visibility="collapsed",
                                 placeholder="Which buckets sum into this %?")
        k["bucket_keys"] = [label_to_key[l] for l in chosen]
        if c4.button("✕", key=f"kd_{project}_{k['key']}", help="Remove KPI"):
            ed["kpis"] = [x for x in ed["kpis"] if x["key"] != k["key"]]
            st.rerun()
    if st.button("+ Add % KPI card", key=f"add_kpi_{project}"):
        ed["kpis"].append(_new_kpi(_next_seq(ed["kpis"], "kpi")))
        st.rerun()

    # ---- Task table columns (order + visibility) ----
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown("**Task Table Columns** — order and show/hide the columns of the sprint sheet")
    st.caption("Applies to the Excel and PDF task table. Use ▲ / ▼ to reorder; untick to hide.")
    cols_ed = ed["task_columns"]
    label_hdr = ("Requirement Type" if _sprint_type_for_project(project) == "Design Sprint" else "Label")
    for i, entry in enumerate(list(cols_ed)):
        cdef = columns.BY_KEY.get(entry["key"])
        if not cdef:
            continue
        name = label_hdr if entry["key"] == columns.LABELS_KEY else cdef["label"]
        c1, c2, c3, c4 = st.columns([4, 1.2, 0.5, 0.5])
        c1.markdown(f"<div style='padding-top:6px;font-size:13px;'>{i + 1}. {_html_escape(name)}</div>",
                    unsafe_allow_html=True)
        entry["visible"] = c2.checkbox("Show", value=entry.get("visible", True),
                                        key=f"cv_{project}_{entry['key']}")
        if c3.button("▲", key=f"cu_{project}_{entry['key']}", help="Move up", disabled=(i == 0)):
            cols_ed[i - 1], cols_ed[i] = cols_ed[i], cols_ed[i - 1]
            st.rerun()
        if c4.button("▼", key=f"cd_{project}_{entry['key']}", help="Move down",
                     disabled=(i == len(cols_ed) - 1)):
            cols_ed[i + 1], cols_ed[i] = cols_ed[i], cols_ed[i + 1]
            st.rerun()
    if not any(e.get("visible", True) for e in cols_ed):
        st.warning("At least one column should stay visible.")
    if st.button("Reset column order", key=f"cols_reset_{project}"):
        ed["task_columns"] = columns.default_columns(_sprint_type_for_project(project))
        st.rerun()

    # ---- Status mapping ----
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown("**Map Jira statuses to buckets**")
    st.caption("Statuses are picked up automatically from the CSV you upload / fetch from Jira. "
               "You can also add them manually below before you have a CSV.")

    cfg = _effective_project_config(project)
    status_display = {}
    for s in cfg.get("known_statuses", []):
        status_display.setdefault(str(s).lower().strip(), str(s))
    up_df = st.session_state.get("uploaded_df")
    if up_df is not None:
        for s in distinct_statuses(up_df):
            status_display.setdefault(s.lower(), s)

    with st.expander("＋ Add Jira statuses manually", expanded=not status_display):
        manual_raw = st.text_area("One status per line", key=f"manual_{project}",
                                   placeholder="In Design Review\nReady for Dev\nDone")
    for s in [x.strip() for x in manual_raw.splitlines() if x.strip()]:
        status_display.setdefault(s.lower(), s)
    ordered = sorted(status_display.items(), key=lambda kv: kv[1].lower())

    saved_status_map = {str(k).lower().strip(): v for k, v in cfg.get("status_map", {}).items()}
    new_status_map = {}
    choices = ["(ignore)"] + bucket_labels
    if not ordered:
        st.info("Upload a Jira CSV on the Generate Report page, or add statuses above, to start mapping.")
    for lower_key, disp in ordered:
        cur_bucket_key = saved_status_map.get(lower_key)
        default_label = key_to_label.get(cur_bucket_key, "(ignore)")
        idx = choices.index(default_label) if default_label in choices else 0
        chosen = st.selectbox(disp, choices, index=idx, key=f"map_{project}_{lower_key}")
        if chosen != "(ignore)":
            new_status_map[lower_key] = label_to_key[chosen]

    unmapped = [disp for lk, disp in ordered if lk not in new_status_map]
    if unmapped:
        st.caption("⚠️ Not mapped (won't be counted in any bucket): " + ", ".join(unmapped))

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    save_col, reset_col = st.columns([1, 1])
    with save_col:
        if st.button("Save", type="primary", use_container_width=True):
            clean_buckets = [b for b in ed["buckets"] if b["label"].strip()]
            clean_kpis = [k for k in ed["kpis"] if k["label"].strip()]
            final_cfg = {
                "buckets": clean_buckets,
                "kpis": clean_kpis,
                "status_map": new_status_map,
                "known_statuses": [disp for _, disp in ordered],
                "task_columns": [
                    {"key": c["key"], "visible": bool(c.get("visible", True))}
                    for c in ed["task_columns"]
                ],
            }
            configs = dict(st.session_state.get("project_configs", {}))
            configs[project] = final_cfg
            st.session_state.project_configs = configs
            _save_to_store("project_configs", project, final_cfg)
            st.session_state.editor.pop(project, None)
            _invalidate_generated_report()
            st.success(f"Saved: {len(clean_buckets)} bucket(s), {len(clean_kpis)} KPI card(s), "
                       f"{len(new_status_map)} status(es) mapped for {project}.")
            st.rerun()
    with reset_col:
        if st.button("Reset to defaults", use_container_width=True):
            configs = dict(st.session_state.get("project_configs", {}))
            configs.pop(project, None)
            st.session_state.project_configs = configs
            _delete_from_store("project_configs", project)
            st.session_state.editor.pop(project, None)
            _invalidate_generated_report()
            st.rerun()


def _report_base_name(form_data: dict) -> str:
    today = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d%b%Y")
    project_name = form_data['project_name'].replace("_", " ")
    return f"{project_name}_Sprint{form_data['sprint_number']}_{today}"


def _html_escape(value) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def _render_multi_file_download_button(files: list) -> None:
    downloads = [
        {"filename": filename, "mime": mime, "data": base64.b64encode(data).decode("ascii")}
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
            width: 100%; border: 0; border-radius: 8px; padding: 12px 32px;
            color: #ffffff; cursor: pointer; font-size: 14px; font-weight: 700;
            background: linear-gradient(135deg, #1F3864, #2E75B6); font-family: Arial, sans-serif;
        }}
        #download-selected:hover {{ filter: brightness(1.06); }}
        #download-status {{ min-height: 18px; margin-top: 8px; color: #64748B; font: 12px Arial, sans-serif; }}
        </style>
        """,
        height=78,
    )


st.set_page_config(page_title="Sprint Report Generator", page_icon=":bar_chart:", layout="wide",
                    initial_sidebar_state="collapsed")

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
    label, .stMarkdown, .stMetric label, [data-testid="stWidgetLabel"] p { color: #1F3864 !important; }
    [data-testid="stNumberInput"] input, [data-testid="stTextInput"] input,
    [data-testid="stDateInput"] input, [data-testid="stSelectbox"] div[data-baseweb="select"] > div {
        background-color: #FFFFFF !important; color: #0F172A !important; border-color: #CBD5E1 !important;
    }
    [data-testid="stMetric"] { background: #FFFFFF !important; border: 1px solid #E2E8F0; border-radius: 8px; padding: 12px 14px; }
    [data-testid="stMetricValue"], [data-testid="stMetricLabel"], [data-testid="stMetricLabel"] p { color: #0F172A !important; }
    .stButton > button { background: linear-gradient(135deg, #1F3864, #2E75B6) !important; color: white !important; border: none !important; border-radius: 8px !important; padding: 12px 32px !important; font-size: 14px !important; font-weight: 600 !important; width: 100% !important; }
    .download-box { background: #F0FDF4; border: 2px solid #86EFAC; border-radius: 12px; padding: 28px; text-align: center; margin: 20px 0; }
    .download-title { font-size: 20px; font-weight: 700; color: #166534; margin-bottom: 8px; }
    .download-sub { font-size: 13px; color: #15803D; }
    .detail-box { background: #FFFFFF; border: 1px solid #E2E8F0; border-left: 4px solid #2E75B6; border-radius: 8px; padding: 12px 14px; min-height: 72px; color: #0F172A; font-size: 12px; line-height: 1.5; }
    .detail-label { color: #64748B; font-size: 11px; font-weight: 700; text-transform: uppercase; margin-bottom: 4px; }
    .detail-value { color: #0F172A; font-size: 13px; font-weight: 600; overflow-wrap: anywhere; }
</style>
""", unsafe_allow_html=True)

for key, default in [
    ('step', 1), ('form_data', {}), ('uploaded_df', None),
    ('excel_bytes', None), ('pdf_bytes', None), ('image_bytes', None), ('parsed_report', None),
]:
    if key not in st.session_state:
        st.session_state[key] = default

if "saved_sprint_details" not in st.session_state:
    st.session_state.saved_sprint_details = _load_saved_sprint_details()

if "project_configs" not in st.session_state:
    st.session_state.project_configs = _load_project_configs()

_migrate_legacy_status_mappings()

if "project_selector" not in st.session_state:
    st.session_state.project_selector = PROJECTS[0]

if "active_project" not in st.session_state:
    st.session_state.active_project = st.session_state.project_selector
    _hydrate_project_form(st.session_state.active_project)

with st.sidebar:
    st.radio("Page", ["Generate Report", "Settings"], key="page")
    _persistent = store.backend_kind() == "ZohoSheetBackend"
    st.caption("💾 Storage: **Zoho Sheet** (saved across restarts)" if _persistent
               else "💾 Storage: **local files** (not persistent on Streamlit Cloud)")
    st.markdown("---")
    st.caption("Configure buckets, % KPIs, colors, and Jira status mapping on the **Settings** page.")

st.markdown('<div class="header-banner"><h1>Sprint Report Generator</h1>'
            '<p>Fill in sprint details - Upload your Jira CSV - Download formatted Excel, PDF and Image reports</p></div>',
            unsafe_allow_html=True)

if st.session_state.get("_store_error"):
    st.warning(st.session_state["_store_error"])

if st.session_state.get("page") == "Settings":
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


@st.dialog("How to Export Your Jira CSV", width="large")
def show_jira_guide():
    st.markdown("""
<style>
.guide-step { display:flex; gap:16px; align-items:flex-start; background:white; border-radius:10px; padding:16px 18px; margin-bottom:12px; box-shadow:0 1px 4px rgba(0,0,0,0.07); border-left: 4px solid #2E75B6; }
.guide-num { background:#1F3864; color:white; border-radius:50%; width:28px; height:28px; min-width:28px; display:flex; align-items:center; justify-content:center; font-size:13px; font-weight:700; }
.guide-body { flex:1; }
.guide-title { font-size:14px; font-weight:700; color:#1F3864; margin-bottom:4px; }
.guide-desc  { font-size:12px; color:#374151; line-height:1.6; }
.jql-box { background:#1E293B; color:#7DD3FC; font-family:monospace; font-size:12px; padding:10px 14px; border-radius:6px; margin-top:8px; word-break:break-all; }
</style>
<div class="guide-step"><div class="guide-num">1</div><div class="guide-body">
<div class="guide-title">Open your Jira Project Board</div>
<div class="guide-desc">Go to Jira and open the relevant project.</div></div></div>
<div class="guide-step"><div class="guide-num">2</div><div class="guide-body">
<div class="guide-title">Click "All Work" -> Filter -> JQL</div>
<div class="guide-desc">Switch the filter panel to the JQL tab.</div></div></div>
<div class="guide-step"><div class="guide-num">3</div><div class="guide-body">
<div class="guide-title">Enter your JQL query</div>
<div class="guide-desc">Example:</div>
<div class="jql-box">project = "Products Design" AND sprint = 4319 ORDER BY issuetype ASC, priority DESC</div>
</div></div>
<div class="guide-step"><div class="guide-num">4</div><div class="guide-body">
<div class="guide-title">Export -> CSV (all fields)</div>
<div class="guide-desc">Top-right "more" menu -> Export -> CSV - all fields.</div></div></div>
<div class="guide-step"><div class="guide-num">5</div><div class="guide-body">
<div class="guide-title">Upload it here</div>
<div class="guide-desc">Complete Step 1 (Sprint Details), then upload the CSV in Step 2.</div></div></div>
""", unsafe_allow_html=True)


nav1, nav2, nav3, help_col = st.columns([1, 1, 1, 0.18])
with nav1:
    st.button("Sprint Details", key="nav_step_1", disabled=(step == 1), use_container_width=True,
              on_click=_go_to_step, args=(1,))
with nav2:
    st.button("Upload Jira CSV", key="nav_step_2", disabled=(step == 2 or not st.session_state.form_data),
              use_container_width=True, on_click=_go_to_step, args=(2,))
with nav3:
    st.button("Download Report", key="nav_step_3", disabled=(step == 3 or st.session_state.uploaded_df is None),
              use_container_width=True, on_click=_go_to_step, args=(3,))
with help_col:
    if st.button("?", help="How to export Jira CSV", use_container_width=True):
        show_jira_guide()

# STEP 1 -----------------------------------------------------------------
if step == 1:
    st.markdown('<div class="section-title">Sprint Type</div>', unsafe_allow_html=True)
    sprint_type = st.radio(
        "Sprint Type", ["Product Sprint", "Design Sprint"], key="sprint_type_input",
        on_change=_on_sprint_type_change, horizontal=True,
        help="Product and Design sprints track different fields.",
    )

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Sprint Information</div>', unsafe_allow_html=True)

    if sprint_type == "Design Sprint":
        project_name = DESIGN_PROJECT
        st.text_input("Project", value=DESIGN_PROJECT, disabled=True)

        c1, c2, c3 = st.columns(3)
        with c1:
            sprint_number = st.number_input("Sprint Number", min_value=1, max_value=999, step=1, key="sprint_number_input")
            sprint_start = st.date_input("Sprint Start Date", key="sprint_start_input")
        with c2:
            sprint_end = st.date_input("Sprint End Date", key="sprint_end_input")
            sprint_release = st.date_input("Sprint Release Date", key="sprint_release_input")
        with c3:
            scrum_master = st.text_input("Scrum Master", placeholder="e.g. Siddharth Verma", key="scrum_master_input")

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Auto-Calculated (live preview)</div>', unsafe_allow_html=True)
        total_days = (sprint_release - sprint_start).days + 1
        days_left = max((sprint_release - _today_ist()).days + 1, 0)
        ac1, ac2, ac3 = st.columns(3)
        ac1.metric("Total No. of Days", total_days)
        ac2.metric("Days Left in Sprint", days_left)
        ac3.metric("Sprint Release Date", sprint_release.strftime("%d %b %Y"))
        st.markdown('<div class="info-pill">Total Days = Release Date - Sprint Start + 1 (inclusive). Days Left = Release Date - Today + 1.</div>', unsafe_allow_html=True)
    else:
        project_name = st.selectbox("Project", _project_options(), key="project_selector",
                                     on_change=_load_selected_project_details, accept_new_options=True,
                                     help="Pick a project, or type a new name and choose the \"Add ...\" option.")

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            sprint_number = st.number_input("Sprint Number", min_value=1, max_value=999, step=1, key="sprint_number_input")
            sprint_start = st.date_input("Sprint Start Date", key="sprint_start_input")
        with c2:
            dev_release = st.date_input("Sprint Development Release", key="dev_release_input")
            qa_release = st.date_input("Sprint QA Release", key="qa_release_input")
        with c3:
            prod_release = st.date_input("Production Release Date", key="prod_release_input")
            sprint_end = st.date_input("Sprint End Date", key="sprint_end_input")
        with c4:
            scrum_master = st.text_input("Scrum Master", placeholder="e.g. Siddharth Verma", key="scrum_master_input")

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
    sprint_goal = st.text_input("Sprint Goal", placeholder="e.g. Profile Builder Dashboard v1", key="sprint_goal_input")
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
        if sprint_type == "Design Sprint":
            if sprint_end < sprint_start:
                errors.append("Sprint End Date cannot be before Sprint Start.")
            if sprint_release < sprint_start:
                errors.append("Sprint Release Date cannot be before Sprint Start.")
        else:
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
                sprint_type=sprint_type, sprint_number=sprint_number, sprint_start=sprint_start,
                sprint_end=sprint_end, total_days=total_days, days_left=days_left,
                scrum_master=scrum_master.strip(), sprint_goal=sprint_goal.strip(),
                major_item_1=major1.strip(), major_item_2=major2.strip(), major_item_3=major3.strip(),
                project_name=project_name,
            )
            if sprint_type == "Design Sprint":
                form_data["sprint_release"] = sprint_release
            else:
                form_data.update(dev_release=dev_release, qa_release=qa_release, prod_release=prod_release)

            _save_project_form_data(project_name, form_data)
            st.session_state.form_data = form_data
            st.session_state.uploaded_df = None
            st.session_state.excel_bytes = None
            st.session_state.pdf_bytes = None
            st.session_state.image_bytes = None
            st.session_state.parsed_report = None
            st.session_state.step = 2
            st.rerun()

# STEP 2 -------------------------------------------------------------------
elif step == 2:
    fd = st.session_state.form_data
    st.markdown('<div class="section-title">Sprint Details Confirmed</div>', unsafe_allow_html=True)
    release_label = "Release Date" if fd["sprint_type"] == "Design Sprint" else "Prod Release"
    release_value = fd["sprint_release"] if fd["sprint_type"] == "Design Sprint" else fd["prod_release"]
    sc = st.columns(5)
    sc[0].metric("Sprint", f"#{fd['sprint_number']}")
    sc[1].metric("Start", fd['sprint_start'].strftime("%d %b %Y"))
    sc[2].metric(release_label, release_value.strftime("%d %b %Y"))
    sc[3].metric("Total Days", fd['total_days'])
    sc[4].metric("Scrum Master", fd['scrum_master'])

    proj = fd['project_name']
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Get Jira Data</div>', unsafe_allow_html=True)

    if jira_client.is_configured():
        saved_q = _load_store("jira_config").get(proj, {})
        with st.expander("🔗 Fetch from Jira (saved filter or JQL)", expanded=True):
            mode_label = st.radio("Query type", ["JQL", "Saved filter ID"],
                                   index=0 if saved_q.get("mode", "jql") == "jql" else 1,
                                   horizontal=True, key=f"jira_mode_{proj}")
            value = st.text_input("Filter ID or JQL", value=saved_q.get("value", ""), key=f"jira_value_{proj}",
                                   placeholder='e.g.  project = "Products Design" AND sprint = 4319 ORDER BY issuetype ASC, priority DESC')
            query = {"mode": "jql" if mode_label == "JQL" else "filter", "value": value}
            st.caption("Saved automatically per project - update just the sprint number/name each time and click Fetch.")
            fc1, fc2 = st.columns([1, 1])
            with fc1:
                if st.button("Fetch from Jira", type="primary", use_container_width=True, disabled=not value.strip()):
                    _save_to_store("jira_config", proj, query)
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
        st.caption("💡 Add a `[jira]` section to Streamlit secrets (base_url, email, api_token) to enable one-click **Fetch from Jira**.")

    st.markdown('<div class="info-pill">Or upload a Jira CSV export (all fields). Required: <b>Issue key, Issue Type, Summary, Status</b>.</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Drop your Jira CSV here", type=["csv"], label_visibility="collapsed")
    if uploaded_file:
        try:
            _ingest_dataframe(pd.read_csv(uploaded_file), uploaded_file.name)
        except Exception as e:
            st.error(f"Could not read CSV: {e}")

    df = st.session_state.uploaded_df
    if df is not None:
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.success(f"Loaded: **{st.session_state.get('uploaded_source', 'data')}** — {len(df)} rows")
        epics = len(df[df['Issue Type'] == 'Epic'])
        stories = len(df[df['Issue Type'].isin(['Story', 'Task'])])
        subs = len(df[df['Issue Type'] == 'Sub-task'])
        pc = st.columns(4)
        pc[0].metric("Total Rows", len(df)); pc[1].metric("Epics", epics)
        pc[2].metric("Stories/Tasks", stories); pc[3].metric("Sub-tasks", subs)

        st.markdown('<div class="section-title">Status Breakdown Preview</div>', unsafe_allow_html=True)
        non_epic = df[df['Issue Type'] != 'Epic']
        sdf = non_epic['Status'].value_counts().reset_index()
        sdf.columns = ['Status', 'Count']
        st.dataframe(sdf, use_container_width=True, hide_index=True)

        cfg = _effective_project_config(proj)
        status_map = {str(k).lower().strip(): v for k, v in cfg.get("status_map", {}).items()}
        unmapped = sorted({s for s in distinct_statuses(non_epic) if s.lower().strip() not in status_map})
        if unmapped:
            st.warning(f"Jira statuses not yet mapped for **{proj}** (won't be counted in any bucket): "
                       f"**{', '.join(unmapped)}**. Map them on the **Settings** page.")

        # ---- Data health: dates + issue types actually usable ----
        start_series = pd.to_datetime(df.get('Custom field (Target start)'), errors='coerce') \
            if 'Custom field (Target start)' in df.columns else pd.Series(pd.NaT, index=df.index)
        end_series = pd.to_datetime(df.get('Custom field (Target end)'), errors='coerce') \
            if 'Custom field (Target end)' in df.columns else pd.Series(pd.NaT, index=df.index)
        with_dates = int((start_series.notna() & end_series.notna()).sum())
        sf, ef = jira_client.last_date_fields()
        if sf or ef:
            st.caption(f"📅 Dates read from — Start: **{sf}** · End: **{ef}**")
        if with_dates == 0:
            st.warning("No issue has both a start and an end date. Rows will show '-'. "
                       "If your team does set dates, the field names may differ - set "
                       "`target_start_field` / `target_end_field` in the `[jira]` secrets.")
        else:
            st.caption(f"📅 {with_dates} of {len(df)} issues have both dates; the rest show '-'.")

        known_types = set(STORY_LEVEL_TYPES) | {'Epic', 'Sub-task'}
        dropped = sorted({str(t) for t in df['Issue Type'].dropna().unique() if str(t) not in known_types})
        if dropped:
            st.warning(f"These issue types won't appear in the task table (they still count in KPIs): "
                       f"**{', '.join(dropped)}**.")
        # Parent epics referenced but not fetched -> rendered as [EXTERNAL EPIC]
        if 'Parent key' in df.columns:
            keys = set(df['Issue key'].astype(str))
            story_rows = df[df['Issue Type'].isin(STORY_LEVEL_TYPES)]
            missing_epics = sorted({str(p) for p in story_rows['Parent key'].dropna()
                                    if str(p).strip() and str(p) not in keys})
            if missing_epics:
                st.warning(f"{len(missing_epics)} parent epic(s) are not in this result set, so they show as "
                           f"**[EXTERNAL EPIC]** with no name. Include them in the query to show full detail.")

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

# STEP 3 --------------------------------------------------------------------
elif step == 3:
    fd = st.session_state.form_data
    df = st.session_state.uploaded_df
    proj = fd['project_name']

    fd["meta_items"] = _build_meta_items(fd)
    # Jira label column header depends on sprint type.
    fd["label_header"] = "Requirement Type" if fd.get("sprint_type") == "Design Sprint" else "Label"

    if st.session_state.parsed_report is None:
        with st.spinner("Parsing Jira data and building your Excel/PDF/Image report..."):
            cfg = _effective_project_config(proj)
            # Column order/visibility: saved config, else the sprint-type default.
            fd["task_columns"] = columns.resolve(cfg.get("task_columns"), fd.get("sprint_type"))
            parsed = parse_jira_csv(df, proj, cfg)
            st.session_state.excel_bytes = build_excel(fd, parsed)
            st.session_state.pdf_bytes = build_pdf(fd, parsed)
            st.session_state.image_bytes = build_summary_image(fd, parsed)
            st.session_state.parsed_report = parsed

    parsed = st.session_state.parsed_report
    kpis = parsed['kpis']
    buckets = parsed['buckets']
    kpi_defs = parsed['kpi_defs']
    excel_bytes = st.session_state.excel_bytes
    pdf_bytes = st.session_state.pdf_bytes
    image_bytes = st.session_state.image_bytes
    base_filename = _report_base_name(fd)
    filename = f"{base_filename}.xlsx"
    pdf_filename = f"{base_filename}.pdf"
    image_filename = f"{base_filename}.png"

    st.markdown('<div class="download-box"><div class="download-title">Report Ready</div><div class="download-sub">Your Sprint Report has been generated successfully.</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Sprint Details</div>', unsafe_allow_html=True)
    detail_items = [(label, value) for label, value, _ in fd["meta_items"]]
    for offset in range(0, len(detail_items), 4):
        detail_cols = st.columns(4)
        for col, (label, value) in zip(detail_cols, detail_items[offset:offset + 4]):
            col.markdown(f'<div class="detail-box"><div class="detail-label">{_html_escape(label)}</div>'
                         f'<div class="detail-value">{_html_escape(value)}</div></div>', unsafe_allow_html=True)

    goal_cols = st.columns([1, 2])
    goal_cols[0].markdown(f'<div class="detail-box"><div class="detail-label">Sprint Goal</div>'
                          f'<div class="detail-value">{_html_escape(fd.get("sprint_goal", "") or "-")}</div></div>', unsafe_allow_html=True)
    major_items = [fd.get('major_item_1', ''), fd.get('major_item_2', ''), fd.get('major_item_3', '')]
    major_text = "<br>".join(_html_escape(item) for item in major_items if item) or "-"
    goal_cols[1].markdown(f'<div class="detail-box"><div class="detail-label">Major Sprint Items</div>'
                          f'<div class="detail-value">{major_text}</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Sprint KPI Summary</div>', unsafe_allow_html=True)
    daily_task = round(kpis["action_items"] / fd["total_days"], 2) if fd["total_days"] > 0 else 0
    k = st.columns(4)
    k[0].metric("Sprint", f"#{fd['sprint_number']}")
    k[1].metric("Action Items", kpis['action_items'])
    k[2].metric("Days Left", fd['days_left'])
    k[3].metric("Daily Task Count", daily_task)
    if kpis['unmapped_count']:
        st.caption(f"⚠️ {kpis['unmapped_count']} item(s) have a status not mapped to any bucket - "
                   f"map them on the **Settings** page so they're counted.")

    if kpi_defs:
        st.markdown('<div class="section-title">% KPI Cards</div>', unsafe_allow_html=True)
        kc = st.columns(min(len(kpi_defs), 5))
        for i, kdef in enumerate(kpi_defs):
            v = kpis['kpi_values'].get(kdef['key'], {'pct_display': '0%'})
            kc[i % len(kc)].markdown(f"""
            <div style="background:white;border-radius:8px;padding:14px;
                        border-left:4px solid {kdef['color']};margin-bottom:8px;
                        box-shadow:0 1px 3px rgba(0,0,0,0.08);">
                <div style="font-size:11px;color:#64748B;font-weight:600;">{_html_escape(kdef['label'])}</div>
                <div style="font-size:24px;font-weight:700;color:{kdef['color']};">{v['pct_display']}</div>
            </div>""", unsafe_allow_html=True)

    if buckets:
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Status Breakdown</div>', unsafe_allow_html=True)
        sc = st.columns(min(len(buckets), 5))
        for i, b in enumerate(buckets):
            val = kpis['bucket_counts'].get(b['key'], 0)
            sc[i % len(sc)].markdown(f"""
            <div style="background:white;border-radius:8px;padding:14px;
                        border-left:4px solid {b['color']};margin-bottom:8px;
                        box-shadow:0 1px 3px rgba(0,0,0,0.08);">
                <div style="font-size:11px;color:#64748B;font-weight:600;">{_html_escape(b['label'])}</div>
                <div style="font-size:24px;font-weight:700;color:{b['color']};">{val}</div>
            </div>""", unsafe_allow_html=True)
    else:
        st.info("No status buckets configured for this project yet. Add some on the **Settings** page.")

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Summary Image Preview</div>', unsafe_allow_html=True)
    st.image(image_bytes, use_container_width=True)

    download_files = {
        "Excel": (filename, excel_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        "PDF": (pdf_filename, pdf_bytes, "application/pdf"),
        "Image": (image_filename, image_bytes, "image/png"),
    }
    selected_formats = st.multiselect("Select files to download", options=list(download_files.keys()),
                                       default=list(download_files.keys()))
    if selected_formats:
        _render_multi_file_download_button([download_files[option] for option in selected_formats])
    else:
        st.markdown('<div class="val-error">Select at least one file type to download.</div>', unsafe_allow_html=True)

    if cliq_client.is_configured():
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Post to Zoho Cliq</div>', unsafe_allow_html=True)
        saved_channel = _load_store("zoho_channels").get(fd['project_name'], "")
        ch1, ch2 = st.columns([3, 1])
        with ch1:
            channel = st.text_input("Cliq channel (unique name)", value=saved_channel,
                                     key=f"cliq_channel_{fd['project_name']}",
                                     placeholder="e.g. design-sprint-reports",
                                     help="The channel's unique name from its URL/settings. Saved per project.")
        with ch2:
            st.write("")
            if st.button("Save channel", use_container_width=True, disabled=not channel.strip()):
                _save_to_store("zoho_channels", fd['project_name'], channel.strip())
                st.success("Channel saved.")
        include_summary = st.checkbox("Also post a summary message", value=False,
                                       help="Off = post only the selected files. On = also post the KPI summary text.")
        st.caption('Posts the files ticked above in "Select files to download".')
        post_disabled = (not channel.strip()) or (not selected_formats)
        if st.button("Post selected files to Cliq", type="primary", disabled=post_disabled):
            _save_to_store("zoho_channels", fd['project_name'], channel.strip())
            message = ""
            if include_summary:
                kpi_bits = " | ".join(f"{kd['label']}: {kpis['kpi_values'].get(kd['key'], {}).get('pct_display', '0%')}"
                                     for kd in kpi_defs)
                message = (f"*Sprint {fd['sprint_number']} report — {fd['project_name']}*\n"
                          f"Action Items: {kpis['action_items']}" + (f" | {kpi_bits}" if kpi_bits else "") +
                          f"\nScrum Master: {fd['scrum_master']}")
            files = [download_files[option] for option in selected_formats]
            try:
                with st.spinner("Posting to Zoho Cliq..."):
                    warnings = cliq_client.post_report(channel.strip(), files, message)
                if warnings:
                    st.warning("Posted, but some files failed:\n\n" + "\n\n".join(warnings))
                else:
                    st.success(f"Posted {len(files)} file(s) to #{channel.strip()}: {', '.join(selected_formats)}"
                              + (" (with summary)" if include_summary else "") + ".")
            except Exception as exc:
                st.error(f"Could not post to Cliq: {exc}")
        if not selected_formats:
            st.caption("Select at least one file above to enable posting.")

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    _, reset_col = st.columns([2, 1])
    with reset_col:
        if st.button("Generate Another Report", key="report_generate_another", use_container_width=True):
            for key in ['form_data', 'uploaded_df', 'excel_bytes', 'pdf_bytes', 'image_bytes', 'parsed_report']:
                st.session_state.pop(key, None)
            st.session_state.step = 1
            _hydrate_project_form(st.session_state.get("project_selector", PROJECTS[0]))
            st.rerun()
