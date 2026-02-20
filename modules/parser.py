"""
parser.py
Reads the Jira CSV export, builds the Epic → Story/Task/Bug → Subtask hierarchy,
and calculates all KPI values for the sprint summary block.

Fixes applied:
  1. Bug issue type is now treated like Story/Task
  2. External Epics (parent key not in CSV) are created as synthetic groups
  3. Standalone items (no parent key) are collected into an UNLINKED group
  4. Cascade sub-task loss fixed as a result of (2) and (3)
"""

import pandas as pd
from datetime import date


# ── Status mappings (adjust if your Jira uses different labels) ──────────────
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


def _match_status(status_val, status_list):
    if pd.isna(status_val):
        return False
    return any(s.lower() in str(status_val).lower() for s in status_list)


def parse_jira_csv(df: pd.DataFrame) -> dict:
    """
    Main entry point. Takes the raw Jira DataFrame and returns a dict with:
      - hierarchy: ordered list of dicts for the Excel task table
      - kpis: all calculated KPI values for the sprint summary block
    """
    # ── Normalise columns ────────────────────────────────────────────────────
    df = df.copy()
    df['Summary']     = df['Summary'].fillna('').str.strip()
    df['Status']      = df['Status'].fillna('').str.strip()
    df['Priority']    = df['Priority'].fillna('').str.strip()
    df['Assignee']    = df['Assignee'].fillna('Unassigned').str.strip()
    df['Issue Type']  = df['Issue Type'].fillna('').str.strip()
    df['Parent key']  = df['Parent key'].fillna('').str.strip() if 'Parent key' in df.columns else ''
    df['Due date']    = pd.to_datetime(df.get('Due date'), errors='coerce')

    target_start_col = 'Custom field (Target start)'
    target_end_col   = 'Custom field (Target end)'
    df['Target Start'] = pd.to_datetime(df.get(target_start_col), errors='coerce') \
        if target_start_col in df.columns else pd.NaT
    df['Target End']   = pd.to_datetime(df.get(target_end_col), errors='coerce') \
        if target_end_col in df.columns else pd.NaT

    # ── Split by type ────────────────────────────────────────────────────────
    epics_df    = df[df['Issue Type'] == 'Epic']
    stories_df  = df[df['Issue Type'].isin(STORY_LEVEL_TYPES)]
    subtasks_df = df[df['Issue Type'] == 'Sub-task']

    # Build a lookup: issue_key → row for all epics present in CSV
    epic_keys_in_csv = set(epics_df['Issue key'].tolist())

    # ── Build ordered hierarchy ──────────────────────────────────────────────
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

    # PASS 3: Standalone items — story-level with NO parent key at all
    standalone_stories = unresolved_stories[
        ~unresolved_stories['Issue key'].isin(processed_stories)
    ]

    # Also catch orphan sub-tasks whose parent story was never found
    placed_story_keys = processed_stories
    orphan_subtasks = subtasks_df[~subtasks_df['Parent key'].isin(placed_story_keys)]

    if len(standalone_stories) > 0 or len(orphan_subtasks) > 0:
        unlinked_header = {
            'level':        0,
            'issue_key':    '—',
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

    # ── KPI calculation (exclude Epics) ─────────────────────────────────────
    non_epic = df[df['Issue Type'] != 'Epic']
    total    = len(non_epic)

    def count_status(status_list):
        return int(non_epic['Status'].apply(lambda s: _match_status(s, status_list)).sum())

    not_initiated  = count_status(STATUS_NOT_INITIATED)
    in_progress    = count_status(STATUS_IN_PROGRESS)
    staging        = count_status(STATUS_STAGING)
    qa_review      = count_status(STATUS_QA_REVIEW)
    qa_deployed    = count_status(STATUS_QA_DEPLOYED)
    qa_approved    = count_status(STATUS_QA_APPROVED)
    production     = count_status(STATUS_PRODUCTION)
    on_hold        = count_status(STATUS_ON_HOLD)
    another_sprint = count_status(STATUS_ANOTHER_SPRINT)

    pending_action = not_initiated

    pending_pct       = round((not_initiated / total * 100), 2) if total else 0
    not_initiated_pct = round((not_initiated / total * 100), 2) if total else 0
    production_pct    = round((production    / total * 100), 2) if total else 0

    kpis = {
        'action_items':           total,
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
    # Split on semicolon — take everything after the 2nd semicolon
    parts = latest.split(';', 2)
    if len(parts) == 3:
        return parts[2].strip()
    # Fallback: return the whole value if format is unexpected
    return latest


def _make_row(row, level: int) -> dict:
    """Convert a DataFrame row into a clean dict for the hierarchy."""
    ts = row.get('Target Start')
    te = row.get('Target End')
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
