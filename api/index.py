import sys
import os
from flask import Flask, render_template, jsonify, request
from datetime import date
from dotenv import load_dotenv

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
load_dotenv()

app = Flask(
    __name__,
    template_folder='../templates',
    static_folder='../static'
)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'dev-secret-key')


def compute_summary(projects):
    total      = len(projects)
    over_fee   = sum(1 for p in projects if p.get('held_time') is not None and p['held_time'] < 0)
    # Held time = only the over-fee portion (negative values = $ spent beyond budget)
    total_held = sum(p['held_time'] for p in projects if p.get('held_time') is not None and p['held_time'] < 0)
    total_fees = sum(p['fees_sold'] for p in projects if p.get('fees_sold') is not None)
    return {
        'total':           total,
        'over_fee':        over_fee,
        'under_fee':       total - over_fee,
        'total_held':      total_held,
        'total_fees_sold': total_fees,
    }


def compute_pod_summary(projects):
    """Group projects by POD.
    held_time_over = sum of only negative held_time (over-fee / at-risk $).
    held_time_net  = net of all held_time values (positive + negative).
    """
    pods = {}
    for p in projects:
        pod = p.get('pod') or 'Unassigned'
        if pod not in pods:
            pods[pod] = {
                'held_time_over': 0.0,  # sum of negatives only
                'held_time_net':  0.0,  # full net
                'fees_sold': 0.0,
                'count': 0,
                'over_count': 0,
                'has_ht': False, 'has_fees': False,
            }
        ht = p.get('held_time')
        if ht is not None:
            pods[pod]['held_time_net'] += ht
            if ht < 0:
                pods[pod]['held_time_over'] += ht
                pods[pod]['over_count'] += 1
            pods[pod]['has_ht'] = True
        if p.get('fees_sold') is not None:
            pods[pod]['fees_sold'] += p['fees_sold']
            pods[pod]['has_fees'] = True
        pods[pod]['count'] += 1

    result = []
    for pod, d in sorted(pods.items()):
        over  = d['held_time_over'] if d['has_ht'] else None
        net   = d['held_time_net']  if d['has_ht'] else None
        fs    = d['fees_sold']      if d['has_fees'] else None
        pct   = (over / fs) if (over is not None and fs) else None
        result.append({
            'pod':            pod,
            'count':          d['count'],
            'over_count':     d['over_count'],
            'held_time_over': over,   # $ over-fee (negative values summed)
            'held_time_net':  net,    # net for reference
            'fees_sold':      fs,
            'pct_held':       pct,    # over / fees_sold
        })
    return result


@app.route('/')
def index():
    from services.wrike_client import get_active_projects, get_completed_projects
    from services.excel_parser import get_file_info

    pod_filter      = request.args.get('pod', 'all')
    designer_filter = request.args.get('designer', 'all')

    active    = get_active_projects()
    completed = get_completed_projects()
    file_info = get_file_info()

    # Collect filter options before applying filters
    pods      = sorted(set(p['pod'] for p in active + completed if p.get('pod')))
    designers = sorted(set(p['designer'] for p in active + completed if p.get('designer')))

    if pod_filter != 'all':
        active    = [p for p in active    if p.get('pod') == pod_filter]
        completed = [p for p in completed if p.get('pod') == pod_filter]
    if designer_filter != 'all':
        active    = [p for p in active    if p.get('designer') == designer_filter]
        completed = [p for p in completed if p.get('designer') == designer_filter]

    return render_template(
        'index.html',
        active_projects      = active,
        completed_projects   = completed,
        active_summary       = compute_summary(active),
        completed_summary    = compute_summary(completed),
        completed_pod_summary= compute_pod_summary(completed),
        pods                 = pods,
        designers            = designers,
        pod_filter           = pod_filter,
        designer_filter      = designer_filter,
        report_date          = date.today().strftime('%B %d, %Y'),
        is_dummy             = False,
        file_info            = file_info,
    )


@app.route('/api/override', methods=['POST'])
def api_override():
    """Save a dashboard-editable field (cos or notes) for a project."""
    from services.excel_parser import save_override
    data    = request.get_json(force=True)
    num     = str(data.get('project_number', '')).strip()
    field   = data.get('field', '')
    value   = data.get('value', '')
    if not num or field not in ('cos', 'notes'):
        return jsonify({'ok': False, 'error': 'invalid request'}), 400
    save_override(num, field, value)
    return jsonify({'ok': True})


@app.route('/api/projects')
def api_projects():
    from services.wrike_client import get_active_projects, get_completed_projects
    active    = get_active_projects()
    completed = get_completed_projects()
    # Convert date objects to strings for JSON
    for p in active + completed:
        for k in ('current_sp', 'current_cc', 'current_cstart'):
            if p.get(k):
                p[k] = p[k].isoformat()
    return jsonify({'active': active, 'completed': completed})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f'\n  Held Time Report running at: http://localhost:{port}\n')
    app.run(debug=True, port=port)
