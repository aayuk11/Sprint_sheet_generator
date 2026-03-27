"""
app.py  —  Sprint Report Generator
Run with: streamlit run app.py
"""

import streamlit as st
import pandas as pd
from datetime import date, datetime
from zoneinfo import ZoneInfo

from modules.parser          import parse_jira_csv
from modules.excel_generator import build_excel

PROJECTS = [
    "PreScreening.io",
    "Transact_Comply",
    "Entity_Hero",
    "DueDiliger",
    "ZiZi",
    "SATOC",
    "WMP",
    "Profile_Builder",
]

st.set_page_config(
    page_title="Sprint Report Generator",
    page_icon="🚀",
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
    .step-bar { display: flex; margin-bottom: 24px; }
    .step { flex: 1; padding: 10px; text-align: center; font-size: 12px; font-weight: 600; background: #E2E8F0; color: #94A3B8; }
    .step.active { background: #1F3864; color: white; }
    .step.done   { background: #00B050; color: white; }
    .step:first-child { border-radius: 8px 0 0 8px; }
    .step:last-child  { border-radius: 0 8px 8px 0; }
    .info-pill { background: #EFF6FF; border-left: 3px solid #2E75B6; padding: 10px 14px; border-radius: 0 6px 6px 0; font-size: 12px; color: #1E40AF; margin: 8px 0; }
    .val-error { background: #FEF2F2; border: 1px solid #FCA5A5; border-radius: 6px; padding: 10px 14px; color: #DC2626; font-size: 12px; margin-top: 8px; }
    .divider { height: 1px; background: #E2E8F0; margin: 18px 0; }
    .stButton > button { background: linear-gradient(135deg, #1F3864, #2E75B6) !important; color: white !important; border: none !important; border-radius: 8px !important; padding: 12px 32px !important; font-size: 14px !important; font-weight: 600 !important; width: 100% !important; }
    .download-box { background: #F0FDF4; border: 2px solid #86EFAC; border-radius: 12px; padding: 28px; text-align: center; margin: 20px 0; }
    .download-title { font-size: 20px; font-weight: 700; color: #166534; margin-bottom: 8px; }
    .download-sub { font-size: 13px; color: #15803D; }
</style>
""", unsafe_allow_html=True)

for key, default in [('step',1),('form_data',{}),('uploaded_df',None),('excel_bytes',None),('parsed_kpis',None)]:
    if key not in st.session_state:
        st.session_state[key] = default

# ═══════════════════════════════════════════
# SIDEBAR — Status Mapping Reference
# ═══════════════════════════════════════════
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
    st.markdown("## 📋 Status Mapping Reference")
    st.markdown("How Jira statuses map to the Sprint Sheet for each project.")
    st.markdown("---")

    for project, mappings in STATUS_MAPPING_REFERENCE.items():
        st.markdown(f"### 🔷 {project}")
        rows_html = ""
        for jira_status, bucket, pct_col in mappings:
            b_color  = BUCKET_COLORS.get(bucket,  "#94A3B8")
            p_color  = PCT_COLORS.get(pct_col,    "#94A3B8")
            rows_html += f"""
            <tr>
                <td style="padding:6px 8px;font-size:12px;border-bottom:1px solid #E2E8F0;">{jira_status}</td>
                <td style="padding:6px 8px;font-size:12px;border-bottom:1px solid #E2E8F0;">
                    <span style="background:{b_color};color:white;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;">{bucket}</span>
                </td>
                <td style="padding:6px 8px;font-size:12px;border-bottom:1px solid #E2E8F0;">
                    <span style="background:{p_color};color:white;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;">{pct_col}</span>
                </td>
            </tr>"""
        st.markdown(f"""
        <table style="width:100%;border-collapse:collapse;margin-bottom:8px;">
            <thead>
                <tr style="background:#1F3864;">
                    <th style="padding:7px 8px;font-size:11px;color:white;text-align:left;">Jira Status</th>
                    <th style="padding:7px 8px;font-size:11px;color:white;text-align:left;">Sprint Sheet Status</th>
                    <th style="padding:7px 8px;font-size:11px;color:white;text-align:left;">Contributes To</th>
                </tr>
            </thead>
            <tbody>{rows_html}</tbody>
        </table>
        """, unsafe_allow_html=True)
        st.markdown("---")

st.markdown('<div class="header-banner"><h1>🚀 Sprint Report Generator</h1><p>Fill in sprint details · Upload your Jira CSV · Download the formatted Excel report</p></div>', unsafe_allow_html=True)

step = st.session_state.step
def _sc(n): return "active" if step==n else ("done" if step>n else "")
st.markdown(f'<div class="step-bar"><div class="step {_sc(1)}">① Sprint Details</div><div class="step {_sc(2)}">② Upload Jira CSV</div><div class="step {_sc(3)}">③ Download Report</div></div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════
# STEP 1
# ═══════════════════════════════════════════
if step == 1:
    with st.form("sprint_form"):
        st.markdown('<div class="section-title">📋 Sprint Information</div>', unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            sprint_number = st.number_input("Sprint Number", min_value=1, max_value=999, value=27, step=1)
            sprint_start  = st.date_input("Sprint Start Date", value=date(2026, 2, 2))
        with c2:
            dev_release = st.date_input("Sprint Development Release", value=date(2026, 2, 18))
            qa_release  = st.date_input("Sprint QA Release",          value=date(2026, 2, 20))
        with c3:
            prod_release = st.date_input("Production Release Date", value=date(2026, 2, 22))
            sprint_end   = st.date_input("Sprint End Date",          value=date(2026, 2, 22))
        with c4:
            project_name = st.selectbox("Project", PROJECTS)
            scrum_master = st.text_input("Scrum Master", placeholder="e.g. Rishav Kumar")

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.markdown('<div class="section-title">⚡ Auto-Calculated (live preview)</div>', unsafe_allow_html=True)
        total_days = (sprint_end - sprint_start).days + 1
        days_left  = max((sprint_end - date.today()).days + 1, 0)
        ac1, ac2, ac3 = st.columns(3)
        ac1.metric("Total No. of Days",   total_days)
        ac2.metric("Days Left in Sprint", days_left)
        ac3.metric("Sprint End Date",     sprint_end.strftime("%d %b %Y"))
        st.markdown('<div class="info-pill">💡 Total Days = Sprint End − Sprint Start + 1 (inclusive). Days Left = Sprint End − Today + 1.</div>', unsafe_allow_html=True)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.markdown('<div class="section-title">🎯 Sprint Goal & Major Items</div>', unsafe_allow_html=True)
        sprint_goal = st.text_input("Sprint Goal", placeholder="e.g. Rule Engine Module Enhancements (Phase 1)")
        mg1, mg2, mg3 = st.columns(3)
        with mg1: major1 = st.text_input("Major Sprint Item 1", placeholder="Item 1")
        with mg2: major2 = st.text_input("Major Sprint Item 2", placeholder="Item 2")
        with mg3: major3 = st.text_input("Major Sprint Item 3", placeholder="Item 3")

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        submitted = st.form_submit_button("Next → Upload Jira CSV")

        if submitted:
            errors = []
            if not scrum_master.strip():            errors.append("Scrum Master name is required.")
            if dev_release  < sprint_start:         errors.append("Dev Release cannot be before Sprint Start.")
            if qa_release   < dev_release:          errors.append("QA Release cannot be before Dev Release.")
            if prod_release < qa_release:           errors.append("Production Release cannot be before QA Release.")
            if sprint_end   < sprint_start:         errors.append("Sprint End Date cannot be before Sprint Start.")
            if errors:
                for e in errors: st.markdown(f'<div class="val-error">⚠️ {e}</div>', unsafe_allow_html=True)
            else:
                st.session_state.form_data = dict(
                    sprint_number=sprint_number, sprint_start=sprint_start,
                    dev_release=dev_release,     qa_release=qa_release,
                    prod_release=prod_release,   sprint_end=sprint_end,
                    total_days=total_days,       days_left=days_left,
                    scrum_master=scrum_master.strip(), sprint_goal=sprint_goal.strip(),
                    major_item_1=major1.strip(), major_item_2=major2.strip(), major_item_3=major3.strip(),
                    project_name=project_name,
                )
                st.session_state.step = 2
                st.rerun()

# ═══════════════════════════════════════════
# STEP 2
# ═══════════════════════════════════════════
elif step == 2:
    fd = st.session_state.form_data
    st.markdown('<div class="section-title">✅ Sprint Details Confirmed</div>', unsafe_allow_html=True)
    sc = st.columns(5)
    sc[0].metric("Sprint",       f"#{fd['sprint_number']}")
    sc[1].metric("Start",        fd['sprint_start'].strftime("%d %b %Y"))
    sc[2].metric("Prod Release", fd['prod_release'].strftime("%d %b %Y"))
    sc[3].metric("Total Days",   fd['total_days'])
    sc[4].metric("Scrum Master", fd['scrum_master'])

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📁 Upload Jira CSV Export</div>', unsafe_allow_html=True)
    st.markdown('<div class="info-pill">💡 In Jira: Board → Export Issues → CSV (all fields). Required columns: <b>Issue key, Issue Type, Summary, Status</b>. Recommended: Priority, Assignee, Parent key, Target start/end.</div>', unsafe_allow_html=True)

    uploaded_file = st.file_uploader("Drop your Jira CSV here", type=["csv"], label_visibility="collapsed")

    if uploaded_file:
        try:
            df = pd.read_csv(uploaded_file)
            missing = [c for c in ['Issue key','Issue Type','Summary','Status'] if c not in df.columns]
            if missing:
                st.error(f"⚠️ Missing columns: {', '.join(missing)}")
            else:
                st.session_state.uploaded_df = df
                epics   = len(df[df['Issue Type']=='Epic'])
                stories = len(df[df['Issue Type'].isin(['Story','Task'])])
                subs    = len(df[df['Issue Type']=='Sub-task'])
                st.success(f"✅ **{uploaded_file.name}** uploaded — {len(df)} rows")
                pc = st.columns(4)
                pc[0].metric("Total Rows",len(df)); pc[1].metric("Epics",epics)
                pc[2].metric("Stories/Tasks",stories); pc[3].metric("Sub-tasks",subs)

                st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
                st.markdown('<div class="section-title">📊 Status Breakdown Preview</div>', unsafe_allow_html=True)
                non_epic = df[df['Issue Type']!='Epic']
                sdf = non_epic['Status'].value_counts().reset_index()
                sdf.columns = ['Status','Count']
                st.dataframe(sdf, use_container_width=True, hide_index=True)

                st.markdown('<div class="section-title">🔍 Data Preview (first 5 rows)</div>', unsafe_allow_html=True)
                pcols = [c for c in ['Issue key','Issue Type','Summary','Status','Priority','Assignee','Parent key'] if c in df.columns]
                st.dataframe(df[pcols].head(5), use_container_width=True, hide_index=True)
        except Exception as e:
            st.error(f"❌ Could not read CSV: {e}")

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    b1, b2 = st.columns([1,3])
    with b1:
        if st.button("← Back"):
            st.session_state.step = 1; st.rerun()
    with b2:
        if st.button("🚀 Generate Excel Report", disabled=(st.session_state.uploaded_df is None)):
            st.session_state.step = 3; st.rerun()
    if st.session_state.uploaded_df is None:
        st.markdown('<div class="val-error">⚠️ Please upload a Jira CSV file first.</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════
# STEP 3
# ═══════════════════════════════════════════
elif step == 3:
    fd  = st.session_state.form_data
    df  = st.session_state.uploaded_df

    if st.session_state.excel_bytes is None:
        with st.spinner("⏳ Parsing Jira data and building your Excel report..."):
            parsed = parse_jira_csv(df, fd['project_name'])
            st.session_state.excel_bytes = build_excel(fd, parsed)
            st.session_state.parsed_kpis = parsed['kpis']

    kpis        = st.session_state.parsed_kpis
    excel_bytes = st.session_state.excel_bytes
    ist_date = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d.%m.%Y")
    filename = f"{fd['project_name']}_{ist_date}.xlsx"

    st.markdown('<div class="download-box"><div class="download-title">✅ Excel Report Ready!</div><div class="download-sub">Your Sprint Report has been generated successfully.</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">📊 Sprint KPI Summary</div>', unsafe_allow_html=True)
    k = st.columns(5)
    k[0].metric("Sprint",          f"#{fd['sprint_number']}")
    k[1].metric("Action Items",    kpis['action_items'])
    k[2].metric("Pending %",       kpis['pending_pct'])
    k[3].metric("Not Initiated %", kpis['not_initiated_pct'])
    k[4].metric("Days Left",       fd['days_left'])

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📋 Status Breakdown</div>', unsafe_allow_html=True)

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
    d1, d2, d3 = st.columns([1,2,1])
    with d2:
        st.download_button(
            label     = f"⬇️  Download {filename}",
            data      = excel_bytes,
            file_name = filename,
            mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    if st.button("🔄 Generate Another Report"):
        for key in ['step','form_data','uploaded_df','excel_bytes','parsed_kpis']:
            st.session_state.pop(key, None)
        st.rerun()

# ═══════════════════════════════════════════
# CHANGELOG  — visible on every step
# ═══════════════════════════════════════════
CHANGELOG = [
    {
        "version": "v1.5",
        "date":    "27 Mar 2026",
        "commit":  "c9e4117",
        "title":   "Project-Specific Jira Status Mapping & Sidebar Reference",
        "tag":     "✨ Feature",
        "tag_color": "#2E75B6",
        "changes": [
            "Added project-specific Jira → Sprint Sheet status mapping for WMP, SATOC, and PreScreening.io",
            "Fixed WMP 'Grooming Completed' typo — 59 items were unmapped causing % to not sum to 100%",
            "Staging Deployed now maps to Staging → Pending % for WMP",
            "Added Completed-QA (absolute count) and Completion-QA% columns to the KPI row in Excel",
            "Added sidebar hamburger menu showing color-coded Jira → Sprint Sheet status reference per project",
        ],
    },
    {
        "version": "v1.4",
        "date":    "25 Feb 2026",
        "commit":  "2db237a",
        "title":   "S.No Column, Renamed Columns & Revised Date Columns",
        "tag":     "✨ Feature",
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
        "tag":     "✨ Feature",
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
        "tag":     "🎨 UI",
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
        "tag":     "🐛 Fix",
        "tag_color": "#C00000",
        "changes": [
            "Added latest Jira comment column to the task table",
            "Fixed Epic → Story → Sub-task hierarchy logic (external epics & orphan sub-tasks now handled)",
            "Added Sprint End Date as a separate input field",
        ],
    },
    {
        "version": "v1.0",
        "date":    "19 Feb 2026",
        "commit":  "98165e7",
        "title":   "Initial Release",
        "tag":     "🚀 Launch",
        "tag_color": "#00B050",
        "changes": [
            "Sprint Report Generator launched",
            "3-step UI: Sprint Details → Upload Jira CSV → Download Excel Report",
            "Auto-calculated KPIs: Action Items, Pending %, Not Initiated %, Production Release %",
            "Jira CSV parsed into Epic → Story/Task → Sub-task hierarchy in Excel",
        ],
    },
]

st.markdown("---")
with st.expander("📜 Changelog — What's New", expanded=False):
    html_parts = []
    for i, entry in enumerate(CHANGELOG):
        latest = '<span style="background:#FFD700;color:#1F3864;font-size:10px;font-weight:700;padding:2px 8px;border-radius:10px;margin-left:8px;">LATEST</span>' if i == 0 else ''
        bullets = "".join(f'<li style="font-size:12px;color:#374151;margin-bottom:4px;">{c}</li>' for c in entry["changes"])
        html_parts.append(
            f'<div style="background:white;border-radius:10px;padding:16px 20px;margin-bottom:12px;'
            f'box-shadow:0 1px 4px rgba(0,0,0,0.08);border-left:4px solid {entry["tag_color"]};">'
            f'<p style="font-size:15px;font-weight:700;color:#1F3864;margin:0 0 8px 0;">'
            f'{entry["version"]} &mdash; {entry["title"]}{latest}</p>'
            f'<p style="margin:0 0 10px 0;">'
            f'<span style="background:{entry["tag_color"]};color:white;font-size:11px;font-weight:600;padding:2px 10px;border-radius:10px;">{entry["tag"]}</span>&nbsp;'
            f'<span style="background:#F1F5F9;color:#64748B;font-size:11px;padding:2px 10px;border-radius:10px;">&#128197; {entry["date"]}</span>&nbsp;'
            f'<span style="background:#F1F5F9;color:#64748B;font-size:11px;padding:2px 10px;border-radius:10px;font-family:monospace;">#{entry["commit"]}</span>'
            f'</p>'
            f'<ul style="margin:0;padding-left:18px;">{bullets}</ul>'
            f'</div>'
        )
    st.markdown("".join(html_parts), unsafe_allow_html=True)
