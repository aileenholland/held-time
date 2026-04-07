import os
import re
import json
import requests
from datetime import datetime

WRIKE_TOKEN    = os.environ.get('WRIKE_TOKEN')
WRIKE_BASE_URL = 'https://www.wrike.com/api/v4'

# ── Folder IDs ──────────────────────────────────────────────────────────────
FOLDER_B_DESIGN       = 'IEAFNYJDI44OSHOZ'
FOLDER_C_CONSTRUCTION = 'IEAFNYJDI44W4VI7'
FOLDER_D_COMPLETION   = 'IEAFNYJDI44WLZ53'
FOLDER_Z_CLOSED       = 'IEAFNYJDI44W42LG'

# ── Custom field IDs ────────────────────────────────────────────────────────
CF = {
    'project_number': 'IEAFNYJDJUAJZVBN',
    'sqft':           'IEAFNYJDJUADEAVQ',
    'designer':       'IEAFNYJDJUADEJY2',   # Lead Designer (contact ID)
    'design_fees':    'IEAFNYJDJUAIMAPR',   # Design Fees Sold
    'gross_fees':     'IEAFNYJDJUAIOVHK',   # Gross Fees Sold
    'cm_fees':        'IEAFNYJDJUAIOVHP',   # CM Fees
    'pod':            'IEAFNYJDJUAGZFRZ',
    'current_sp':     'IEAFNYJDJUAIRIFR',   # Current SP = Ready For Takeover
    'current_cc':     'IEAFNYJDJUAIRIFQ',   # Current CC
    'current_cstart': 'IEAFNYJDJUAIRIFN',   # Current Construction Start
}


def _headers():
    return {'Authorization': f'Bearer {WRIKE_TOKEN}'}


def _cf(custom_fields, field_id):
    for cf in custom_fields:
        if cf['id'] == field_id:
            return cf.get('value')
    return None


def _parse_date(s):
    if not s:
        return None
    try:
        return datetime.strptime(str(s)[:10], '%Y-%m-%d').date()
    except Exception:
        return None


def _parse_pod(pod_raw):
    if not pod_raw:
        return ''
    raw = str(pod_raw).strip()
    # Stored as JSON array e.g. '["1"]' or plain '1' or 'POD 1' or '1-Name/Name'
    if raw.startswith('['):
        try:
            vals = json.loads(raw)
            raw = str(vals[0]) if vals else raw
        except Exception:
            pass
    # Extract leading digit
    m = re.match(r'^(\d+)', raw)
    if m:
        return f'POD {m.group(1)}'
    return raw


_PHASE_PCT = {
    1:  0.05,
    2:  0.09,
    3:  0.25,
    4:  0.25,
    5:  0.68,
    6:  0.71,
    7:  0.71,
    8:  0.74,
    9:  0.91,
    10: 0.91,
    11: 1.00,
}

def _phase_to_pct(phase_name):
    """Return the correct % complete for a D&B phase name.
    'X - Completed' → 1.0.  Returns None if blank or unrecognised.
    """
    if not phase_name:
        return None
    s = str(phase_name).strip()
    if s.upper().startswith('X'):
        return 1.0
    m = re.match(r'^(\d+)', s)
    if m:
        return _PHASE_PCT.get(int(m.group(1)))
    return None


def _get_contact_map():
    resp = requests.get(f'{WRIKE_BASE_URL}/contacts', headers=_headers())
    resp.raise_for_status()
    return {
        c['id']: f"{c.get('firstName', '')} {c.get('lastName', '')}".strip()
        for c in resp.json().get('data', [])
    }


def _get_status_map():
    """Return {customStatusId: statusName} from all workspace workflows."""
    resp = requests.get(f'{WRIKE_BASE_URL}/workflows', headers=_headers())
    resp.raise_for_status()
    mapping = {}
    for wf in resp.json().get('data', []):
        for s in wf.get('customStatuses', []):
            mapping[s['id']] = s['name']
    return mapping


def _get_child_ids(folder_id):
    resp = requests.get(f'{WRIKE_BASE_URL}/folders/{folder_id}', headers=_headers())
    resp.raise_for_status()
    data = resp.json().get('data', [])
    if not data:
        return []
    return [cid for cid in data[0].get('childIds', []) if cid != folder_id]


def _fetch_projects_batch(ids):
    """Fetch up to 100 folder/project records in one API call."""
    if not ids:
        return []
    chunk = ','.join(ids[:100])
    resp = requests.get(f'{WRIKE_BASE_URL}/folders/{chunk}', headers=_headers())
    if resp.status_code != 200:
        return []
    return resp.json().get('data', [])


def _parse_project(proj, contacts, status_map=None):
    """Convert a Wrike folder record into a dashboard dict."""
    cfl = proj.get('customFields', [])
    title = proj.get('title', '')

    number = _cf(cfl, CF['project_number']) or ''
    if not number:
        m = re.match(r'\[(\d+)\]', title)
        number = m.group(1) if m else ''
    clean_name = re.sub(r'^\[\d+\]\s*', '', title).strip()

    sqft_raw    = _cf(cfl, CF['sqft'])
    sqft        = int(float(sqft_raw)) if sqft_raw else None
    design_fees = float(_cf(cfl, CF['design_fees']) or 0) or None
    cm_fees     = float(_cf(cfl, CF['cm_fees']) or 0) or None
    gross_fees  = float(_cf(cfl, CF['gross_fees']) or 0) or None

    designer_id = _cf(cfl, CF['designer'])
    designer    = contacts.get(designer_id, '') if designer_id else ''

    pod         = _parse_pod(_cf(cfl, CF['pod']))
    current_sp  = _parse_date(_cf(cfl, CF['current_sp']))
    current_cc  = _parse_date(_cf(cfl, CF['current_cc']))
    current_cstart = _parse_date(_cf(cfl, CF['current_cstart']))

    # Phase + % Phase from Wrike workflow custom status
    custom_status_id = (proj.get('project') or {}).get('customStatusId')
    phase     = (status_map or {}).get(custom_status_id, '') if custom_status_id else ''
    pct_phase = _phase_to_pct(phase)

    return {
        'id':             proj['id'],
        'number':         number,
        'name':           clean_name,
        'designer':       designer,
        'pod':            pod,
        'sqft':           sqft,
        'fees_sold':      design_fees,
        'cm_fees':        cm_fees,
        'gross_fees':     gross_fees,
        # These come from time tracking — not available via folder API
        'actual_fees':    None,
        'sold_vs_spent':  None,
        'cos':            None,
        'fee_plus_co':    None,
        'pct_phase':      pct_phase,
        'pct_fee':        None,
        'held_time':      None,
        'ht_category':    '',
        'notes':          '',
        'phase':          phase,
        # Dates
        'current_sp':     current_sp,
        'current_cc':     current_cc,
        'current_cstart': current_cstart,
    }


def _merge_excel_data(projects, billing_data, active_tracker, completed_tracker=None):
    """Merge Excel-sourced fee data into a list of project dicts (in-place)."""
    from services.excel_parser import load_overrides
    overrides = load_overrides()

    for p in projects:
        num = p['number']

        # ── Billing v2 export → actual fees & sold vs spent ──────────────────
        if num in billing_data:
            b = billing_data[num]
            p['actual_fees']   = b.get('actual_fees')
            p['sold_vs_spent'] = b.get('sold_vs_spent')
            # Phase comes from Wrike; only fall back to Excel if Wrike has none
            if not p.get('phase'):
                p['phase'] = b.get('status_xl', '')
            # Fill blanks from Wrike with Excel values
            if not p.get('designer') and b.get('designer_xl'):
                p['designer'] = b['designer_xl']
            if not p.get('sqft') and b.get('sqft_xl'):
                p['sqft'] = b['sqft_xl']

        # ── MASTER TRACKING → COs, held time, categories ─────────────────────
        tracker = active_tracker if completed_tracker is None else completed_tracker
        if num in tracker:
            t = tracker[num]
            p['cos']          = t.get('cos')
            # pct_phase is now calculated from the Wrike phase — don't overwrite
            # with the stale Excel value; only use Excel as fallback if Wrike gave none
            if p.get('pct_phase') is None:
                p['pct_phase'] = t.get('pct_phase')
            # pct_fee is now calculated from actual_fees / fees_sold — not from Excel
            # held_time is now calculated as Fee+CO − Actual Fees
            p['ht_category']  = t.get('ht_category', '')
            p['notes']        = t.get('notes', '')
            # Completed Projects sheet also carries actual fees, sold vs spent,
            # fees sold, POD and date — use these when billing export has no data
            if completed_tracker is not None:
                if p.get('actual_fees') is None and t.get('actual_fees') is not None:
                    p['actual_fees'] = t['actual_fees']
                if p.get('sold_vs_spent') is None and t.get('sold_vs_spent') is not None:
                    p['sold_vs_spent'] = t['sold_vs_spent']
                if p.get('fees_sold') is None and t.get('fees_sold') is not None:
                    p['fees_sold'] = t['fees_sold']
                if t.get('pod'):
                    p['pod'] = t['pod']
                if t.get('date_completed') and not p.get('date_completed'):
                    p['date_completed'] = t['date_completed']

        # ── Dashboard overrides (editable COs & Notes) ───────────────────────
        if num in overrides:
            ov = overrides[num]
            if 'cos' in ov and ov['cos'] is not None:
                try:
                    p['cos'] = float(ov['cos']) if ov['cos'] != '' else p.get('cos')
                except (ValueError, TypeError):
                    pass
            if 'notes' in ov:
                p['notes'] = ov['notes']

        # ── Fee + CO calculated from Fees Sold + COs ─────────────────────────
        fees = p.get('fees_sold')
        cos  = p.get('cos')
        p['fee_plus_co'] = round(fees + (cos or 0), 2) if fees is not None else None

        # ── Sold vs Spent = Fees Sold − Actual Fees ──────────────────────────
        actual = p.get('actual_fees')
        p['sold_vs_spent'] = round(fees - actual, 2) if (fees is not None and actual is not None) else p.get('sold_vs_spent')

        # ── % Fee = Actual Fees / Fees Sold ──────────────────────────────────
        p['pct_fee'] = round(actual / fees, 4) if (actual is not None and fees) else None

        # ── Held Time = Fee + CO − Actual Fees ───────────────────────────────
        fee_plus_co = p.get('fee_plus_co')
        p['held_time'] = round(fee_plus_co - actual, 2) if (fee_plus_co is not None and actual is not None) else None


def get_active_projects():
    """Return all projects in B-Design and C-Construction, enriched with Excel fee data."""
    from services.excel_parser import get_latest_billing_export, get_tracker_data

    contacts = _get_contact_map()

    status_map = _get_status_map()

    all_ids = []
    for folder_id in [FOLDER_B_DESIGN, FOLDER_C_CONSTRUCTION]:
        all_ids.extend(_get_child_ids(folder_id))

    records  = []
    for i in range(0, len(all_ids), 100):
        records.extend(_fetch_projects_batch(all_ids[i:i + 100]))

    projects = [_parse_project(p, contacts, status_map) for p in records]

    _, _, billing_data  = get_latest_billing_export()
    active_tracker, _   = get_tracker_data()
    _merge_excel_data(projects, billing_data, active_tracker)

    projects.sort(key=lambda p: p['number'])
    return projects


def get_completed_projects():
    """
    Return completed projects.

    Source of truth: the 'Completed Projects' sheet in the 2026 Held Time
    Tracker Excel file.  Every row in that sheet = one completed project.

    Each entry is enriched with Wrike data (dates, designer, POD) where the
    project can be found by number in the D-Completion or Z-Closed folders.

    Auto-detection: projects that have disappeared from the active Billing v2
    export AND are confirmed in the Wrike D-Completion folder but are NOT yet
    in the Excel completed sheet are appended with  new_completion=True  so
    the dashboard can flag them for the user to add to the tracker.
    """
    from services.excel_parser import get_latest_billing_export, get_tracker_data

    contacts   = _get_contact_map()
    status_map = _get_status_map()

    # ── Build Wrike lookup from D-Completion (+ Z-Closed) ────────────────────
    d_ids = set(_get_child_ids(FOLDER_D_COMPLETION))
    z_ids = set(_get_child_ids(FOLDER_Z_CLOSED))
    all_ids = list(d_ids | z_ids)

    records = []
    for i in range(0, len(all_ids), 100):
        records.extend(_fetch_projects_batch(all_ids[i:i + 100]))

    wrike_by_num   = {}
    in_d_completion = set()
    for rec in records:
        p   = _parse_project(rec, contacts, status_map)
        num = p['number']
        if num and num not in wrike_by_num:
            wrike_by_num[num] = p
        if rec['id'] in d_ids and num:
            in_d_completion.add(num)

    _, _, billing_data        = get_latest_billing_export()
    _, completed_tracker      = get_tracker_data()
    active_billing_nums       = set(billing_data.keys())

    def _make_stub(num, t):
        """Build a minimal project dict from Excel tracker data alone."""
        return {
            'id': '', 'number': num,
            'name':           t.get('name', f'Project {num}'),
            'designer':       t.get('designer', ''),
            'pod':            t.get('pod', ''),
            'sqft':           t.get('sqft'),
            'fees_sold':      t.get('fees_sold'),
            'cm_fees': None, 'gross_fees': None,
            'actual_fees':    None, 'sold_vs_spent': None,
            'cos': None, 'fee_plus_co': None,
            'pct_phase': None, 'pct_fee': None,
            'held_time': None, 'ht_category': '', 'notes': '',
            'current_sp': None, 'current_cc': None, 'current_cstart': None,
            'phase': '',
        }

    def _assign_date(p, num):
        sp = p.get('current_sp')
        if sp:
            p['date_completed'] = sp.isoformat()
            p['date_source']    = 'wrike'
        elif num in completed_tracker and completed_tracker[num].get('date_completed'):
            p['date_completed'] = completed_tracker[num]['date_completed']
            p['date_source']    = 'excel'
        else:
            p['date_completed'] = ''
            p['date_source']    = ''

    projects = []
    seen     = set()

    # ── Pass 1: Excel Completed Projects sheet is the definitive list ─────────
    for num, t in completed_tracker.items():
        if num in wrike_by_num:
            p = wrike_by_num[num]
        else:
            p = _make_stub(num, t)
        p['new_completion'] = False
        _assign_date(p, num)
        projects.append(p)
        seen.add(num)

    # ── Pass 2: Auto-detect newly completed (in D-Completion, off billing) ────
    for num in in_d_completion:
        if num in seen:
            continue
        if num in active_billing_nums:
            continue   # still showing as active → skip
        p = wrike_by_num[num]
        p['new_completion'] = True   # flag for dashboard highlight
        _assign_date(p, num)
        projects.append(p)
        seen.add(num)

    _merge_excel_data(projects, billing_data, {}, completed_tracker)

    projects.sort(key=lambda p: p.get('date_completed', ''))
    return projects
