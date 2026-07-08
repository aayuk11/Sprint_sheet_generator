"""
parser.py
Reads the Jira CSV export, builds the Epic -> Story/Task/Bug -> Subtask hierarchy,
and calculates all KPI values for the sprint summary block.

Fixes applied:
  1. Bug issue type is now treated like Story/Task
  2. External Epics (parent key not in CSV) are created as synthetic groups
  3. Standalone items (no parent key) are collected into an UNLINKED group
  4. Cascade sub-task loss fixed as a result of (2) and (3)
"""

import pandas as pd
from datetime import date


# Global fallback status mappings
STATUS_NOT_INITIATED   = ['To Do', 'Not Initiated', 'Open']
STATUS_IN_PROGRESS     = ['In Progress']
STATUS_STAGING         = ['Staging Deployed', 'Staging']
STATUS_QA_REVIEW       = ['QA Review', 'In Review']
STATUS_QA_DEPLOYED     = ['QA Deployed']
STATUS_QA_APPROVED     = ['QA Approved']
STATUS_PRODUCTION      = ['Done', 'Production', 'Released', 'Closed']
STATUS_ON_HOLD         = ['On Hold', 'Blocked']
STATUS_ANOTHER_SPRINT  = ['To Be Picked In Another Sprint', 'Deferred']

# Issue types treated as Stories/Tasks (children of Epics)
STORY_LEVEL_TYPES = ['Story', 'Task', 'Bug', 'Improvement', 'New Feature']
REQUIRED_COLUMNS = ['Issue key', 'Issue Type', 'Summary', 'Status']

# Project-specific: Jira status (lowercase) -> Section 2 bucket
PROJECT_STATUS_MAP = {
    'WMP': {
        'to do':             'not_initiated',
        'grooming completed': 'not_initiated',
        'in progress':       'in_progress',
        'staging deployed':  'staging',
        'qa':                'qa_approved',    # WMP's QA = Completed-QA stage
        'done':              'production',
    },
    'SATOC': {
        'to do':             'not_initiated',
        'in progress':       'in_progress',
        'stage deployed':    'staging',
        'qa review':         'qa_review',
        'done':              'production',
    },
    'PreScreening.io': {
        'grooming completed': 'not_initiated',
        'to do':              'not_initiated',
        'in progress':        'in_progress',
        'stage deployed':     'staging',
        'qa deployed':        'qa_deployed',
        'done':               'production',
    },
}

# Canonical Sprint-Sheet buckets (internal keys) and their display labels.
# These are the target categories a Jira status can be mapped to.
BUCKET_KEYS = [
    'not_initiated', 'in_progress', 'staging', 'qa_review',
    'qa_deployed', 'qa_approved', 'production', 'on_hold', 'to_be_picked',
]
BUCKET_LABELS = {
    'not_initiated': 'Not Initiated',
    'in_progress':   'In Progress',
    'staging':       'Staging',
    'qa_review':     'QA Review',
    'qa_deployed':   'QA Deployed',
    'qa_approved':   'QA Approved (Completed-QA)',
    'production':    'Production',
    'on_hold':       'On Hold',
    'to_be_picked':  'To Be Picked (Another Sprint)',
}

# The four % KPIs and which buckets roll into each by default.
PCT_KEYS = ['not_initiated_pct', 'pending_pct', 'completion_qa_pct', 'production_release_pct']
PCT_LABELS = {
    'not_initiated_pct':      'Not Initiated %',
    'pending_pct':            'Pending %',
    'completion_qa_pct':      'Completion - QA %',
    'production_release_pct': 'Production Release %',
}
DEFAULT_PCT_BUCKETS = {
    'not_initiated_pct':      ['not_initiated'],
    'pending_pct':            ['in_progress', 'staging'],
    'completion_qa_pct':      ['qa_review', 'qa_deployed', 'qa_approved'],
    'production_release_pct': ['production'],
}

# Project-specific: which Section 2 buckets roll up into each % KPI
PROJECT_PCT_BUCKETS = {
    'WMP': {
        'pending_pct':            ['in_progress', 'staging'],  # Staging Deployed = still pending
        'not_initiated_pct':      ['not_initiated'],
        'completion_qa_pct':      ['qa_approved'],
        'production_release_pct': ['production'],
    },
    'SATOC': {
        'pending_pct':            ['in_progress'],
        'not_initiated_pct':      ['not_initiated'],
        'completion_qa_pct':      ['staging', 'qa_review'],  # both stages combined
        'production_release_pct': ['production'],
    },
    'PreScreening.io': {
        'pending_pct':            ['in_progress', 'staging'],  # Stage Deployed = still pending
        'not_initiated_pct':      ['not_initiated'],
        'completion_qa_pct':      ['qa_deployed'],
        'production_release_pct': ['production'],
    },
}


def _ensure_text_column(df: pd.DataFrame, column: str, default: str = '') -> None:
    if column not in df.columns:
        df[column] = default
    df[column] = df[column].fillna(default).astype(str).str.strip()


def _ensure_datetime_column(df: pd.DataFrame, source_column: str, target_column: str | None = None) -> None:
    target = target_column or source_column
    if source_column in df.columns:
        df[target] = pd.to_datetime(df[source_column], errors='coerce')
    else:
        df[target] = pd.NaT


def _match_status(status_val, status_list):
    if pd.isna(status_val):
        return False
    return any(s.lower() in str(status_val).lower() for s in status_list)


def parse_jira_csv(df: pd.DataFrame, project_name: str = '', status_config: dict | None = None) -> dict:
    """
    Main entry point. Takes the raw Jira DataFrame and returns a dict with:
      - hierarchy: ordered list of dicts for the Excel task table
      - kpis: all calculated KPI values for the sprint summary block
    """
    # Normalise columns and supply safe defaults for optional Jira export fields.
    df = df.copy()
    missing_required = [column for column in REQUIRED_COLUMNS if column not in df.columns]
    if missing_required:
        raise ValueError(f"Missing required columns: {', '.join(missing_required)}")

    _ensure_text_column(df, 'Issue key')
    _ensure_text_column(df, 'Issue Type')
    _ensure_text_column(df, 'Summary')
    _ensure_text_column(df, 'Status')
    _ensure_text_column(df, 'Priority')
    _ensure_text_column(df, 'Assignee', 'Unassigned')
    _ensure_text_column(df, 'Parent key')
    _ensure_datetime_column(df, 'Due date')

    target_start_col = 'Custom field (Target start)'
    target_end_col = 'Custom field (Target end)'
    _ensure_datetime_column(df, target_start_col, 'Target Start')
    _ensure_datetime_column(df, target_end_col, 'Target End')

    # Jira Created / Updated dates - used as a fallback when the Target start/end
    # custom fields are empty, so the report's start/end columns are never blank.
    _ensure_datetime_column(df, 'Created')
    _ensure_datetime_column(df, 'Updated')

    # Split by type
    epics_df    = df[df['Issue Type'] == 'Epic']
    stories_df  = df[df['Issue Type'].isin(STORY_LEVEL_TYPES)]
    subtasks_df = df[df['Issue Type'] == 'Sub-task']

    # Build a lookup: issue_key -> row for all epics present in CSV
    epic_keys_in_csv = set(epics_df['Issue key'].tolist())

    # Build ordered hierarchy
    hierarchy = []
    processed_stories = set()  # track which story-level items we've placed

    # PASS 1: Walk through epics that ARE in the CSV
    for _, epic in epics_df.iterrows():
        ek = epic['Issue key']
        hierarchy.append(_make_row(epic, level=0))

        children = stories_df[stories_df['Parent key'] == ek]
        for _, child in children.iterrows():
            ck = child['Issue key']
            hierarchy.append(_make_row(child, level=1))
            processed_stories.add(ck)

            subs = subtasks_df[subtasks_df['Parent key'] == ck]
            for _, sub in subs.iterrows():
                hierarchy.append(_make_row(sub, level=2))

    # PASS 2: Stories/Tasks/Bugs whose parent Epic is NOT in the CSV (external epics)
    # Group them by their parent key so they appear under a synthetic epic header
    unresolved_stories = stories_df[~stories_df['Issue key'].isin(processed_stories)]
    external_stories   = unresolved_stories[unresolved_stories['Parent key'] != '']

    external_epic_groups = {}
    for _, row in external_stories.iterrows():
        pk = row['Parent key']
        external_epic_groups.setdefault(pk, []).append(row)

    for ext_epic_key, children in external_epic_groups.items():
        # Insert a synthetic Epic header row
        synthetic_epic = {
            'level':        0,
            'issue_key':    ext_epic_key,
            'issue_type':   'Epic',
            'summary':      f'[EXTERNAL EPIC] {ext_epic_key}',
            'status':       '',
            'priority':     '',
            'assignee':     '',
            'target_start': '-',
            'target_end':   '-',
        }
        hierarchy.append(synthetic_epic)

        for child_row in children:
            ck = child_row['Issue key']
            hierarchy.append(_make_row(child_row, level=1))
            processed_stories.add(ck)

            subs = subtasks_df[subtasks_df['Parent key'] == ck]
            for _, sub in subs.iterrows():
                hierarchy.append(_make_row(sub, level=2))

    # PASS 3: Standalone items - story-level with NO parent key at all
    standalone_stories = unresolved_stories[
        ~unresolved_stories['Issue key'].isin(processed_stories)
    ]

    # Also catch orphan sub-tasks whose parent story was never found
    placed_story_keys = processed_stories
    orphan_subtasks = subtasks_df[~subtasks_df['Parent key'].isin(placed_story_keys)]

    if len(standalone_stories) > 0 or len(orphan_subtasks) > 0:
        unlinked_header = {
            'level':        0,
            'issue_key':    '-',
            'issue_type':   'Group',
            'summary':      'UNLINKED / STANDALONE ITEMS',
            'status':       '',
            'priority':     '',
            'assignee':     '',
            'target_start': '-',
            'target_end':   '-',
        }
        hierarchy.append(unlinked_header)

        for _, row in standalone_stories.iterrows():
            hierarchy.append(_make_row(row, level=1))

        for _, row in orphan_subtasks.iterrows():
            hierarchy.append(_make_row(row, level=2))

    # KPI calculation (exclude Epics)
    non_epic = df[df['Issue Type'] != 'Epic']
    total    = len(non_epic)

    def count_status(status_list):
        return int(non_epic['Status'].apply(lambda s: _match_status(s, status_list)).sum())

    # Resolve which status->bucket map and %-rollup to use, in priority order:
    #   1. user-defined settings passed in via status_config (Settings page)
    #   2. built-in map for a known project
    #   3. global keyword fallback (the else branch below)
    resolved_map = None
    resolved_pct = None
    if status_config and status_config.get('status_map'):
        resolved_map = {str(k).lower().strip(): v for k, v in status_config['status_map'].items()}
        resolved_pct = status_config.get('pct_buckets') or DEFAULT_PCT_BUCKETS
    elif project_name in PROJECT_STATUS_MAP:
        resolved_map = PROJECT_STATUS_MAP[project_name]
        resolved_pct = PROJECT_PCT_BUCKETS[project_name]

    if resolved_map is not None:
        # Project-specific counting (built-in or user-defined)
        pct_map = resolved_pct

        # Zero-out all buckets, then count each item into its mapped bucket
        buckets = {b: 0 for b in BUCKET_KEYS}
        for status_val in non_epic['Status']:
            key = str(status_val).lower().strip()
            bucket = resolved_map.get(key)
            if bucket in buckets:
                buckets[bucket] += 1

        not_initiated  = buckets['not_initiated']
        in_progress    = buckets['in_progress']
        staging        = buckets['staging']
        qa_review      = buckets['qa_review']
        qa_deployed    = buckets['qa_deployed']
        qa_approved    = buckets['qa_approved']
        production     = buckets['production']
        on_hold        = buckets['on_hold']
        another_sprint = buckets['to_be_picked']

        def _pct(bucket_list):
            return round(sum(buckets[b] for b in bucket_list) / total * 100, 2) if total else 0

        completed_qa      = sum(buckets[b] for b in pct_map.get('completion_qa_pct', []))
        pending_pct       = _pct(pct_map.get('pending_pct', []))
        not_initiated_pct = _pct(pct_map.get('not_initiated_pct', []))
        completion_qa_pct = _pct(pct_map.get('completion_qa_pct', []))
        production_pct    = _pct(pct_map.get('production_release_pct', []))

    else:
        # Global fallback counting
        not_initiated  = count_status(STATUS_NOT_INITIATED)
        in_progress    = count_status(STATUS_IN_PROGRESS)
        staging        = count_status(STATUS_STAGING)
        qa_review      = count_status(STATUS_QA_REVIEW)
        qa_deployed    = count_status(STATUS_QA_DEPLOYED)
        qa_approved    = count_status(STATUS_QA_APPROVED)
        production     = count_status(STATUS_PRODUCTION)
        on_hold        = count_status(STATUS_ON_HOLD)
        another_sprint = count_status(STATUS_ANOTHER_SPRINT)

        completed_qa      = qa_approved
        pending_pct       = round((in_progress / total * 100), 2) if total else 0
        not_initiated_pct = round((not_initiated / total * 100), 2) if total else 0
        completion_qa_pct = round((completed_qa  / total * 100), 2) if total else 0
        production_pct    = round((production    / total * 100), 2) if total else 0

    pending_action = not_initiated

    kpis = {
        'action_items':           total,
        'completed_qa':           completed_qa,
        'completion_qa_pct':      f"{completion_qa_pct}%",
        'pending_pct':            f"{pending_pct}%",
        'not_initiated_pct':      f"{not_initiated_pct}%",
        'production_release_pct': f"{production_pct}%",
        'pending_action_items':   pending_action,
        'not_initiated':          not_initiated,
        'in_progress':            in_progress,
        'staging':                staging,
        'qa_review':              qa_review,
        'qa_deployed':            qa_deployed,
        'qa_approved':            qa_approved,
        'production':             production,
        'on_hold':                on_hold,
        'to_be_picked':           another_sprint,
    }

    return {
        'hierarchy': hierarchy,
        'kpis':      kpis,
        'raw_df':    df,
    }


def _extract_latest_comment(row) -> str:
    """
    Looks through Comment, Comment.1 ... Comment.6 columns,
    finds the last non-empty one, and returns only the text after
    the second semicolon (i.e. the actual comment text).
    Format: "DD/Mon/YY HH:MM AM/PM ; AUTHOR_UUID ; COMMENT_TEXT"
    """
    comment_cols = ['Comment', 'Comment.1', 'Comment.2', 'Comment.3',
                    'Comment.4', 'Comment.5', 'Comment.6']
    latest = ''
    for col in comment_cols:
        val = row.get(col, '')
        if pd.notna(val) and str(val).strip():
            latest = str(val).strip()
    if not latest:
        return ''
    # Split on semicolon and take everything after the 2nd semicolon
    parts = latest.split(';', 2)
    if len(parts) == 3:
        return parts[2].strip()
    # Fallback: return the whole value if format is unexpected
    return latest


def _make_row(row, level: int) -> dict:
    """Convert a DataFrame row into a clean dict for the hierarchy."""
    ts = row.get('Target Start')
    te = row.get('Target End')
    # Fall back to Jira Created / Updated when Target start/end are empty.
    if pd.isna(ts):
        ts = row.get('Created')
    if pd.isna(te):
        te = row.get('Updated')
    return {
        'level':          level,
        'issue_key':      row['Issue key'],
        'issue_type':     row['Issue Type'],
        'summary':        row['Summary'],
        'status':         row['Status'],
        'priority':       row['Priority'],
        'assignee':       row['Assignee'],
        'target_start':   ts.strftime('%d %b %Y') if pd.notna(ts) else '-',
        'target_end':     te.strftime('%d %b %Y') if pd.notna(te) else '-',
        'latest_comment': _extract_latest_comment(row),
    }


def calculate_daily_task_count(total_action_items: int, total_days: int) -> str:
    """Daily Task Count = Total Action Items / Total Sprint Days"""
    if total_days <= 0:
        return '0'
    return str(round(total_action_items / total_days, 2))
