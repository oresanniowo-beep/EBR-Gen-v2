import sys, json, re, os, subprocess, shutil, tempfile

# ── Font scaling ────────────────────────────────────────────────────────────
def smart_sz(text, original_sz):
    n = len(str(text))
    sz = int(original_sz)
    if n <= 3:    return sz
    elif n <= 5:  return int(sz * 0.90)
    elif n <= 8:  return int(sz * 0.75)
    elif n <= 12: return int(sz * 0.55)
    elif n <= 16: return int(sz * 0.42)
    else:         return int(sz * 0.32)

def split_runs(xml):
    segments = []
    last = 0
    for m in re.finditer(r'<a:r\b', xml):
        start = m.start()
        end = xml.find('</a:r>', start)
        if end == -1: continue
        end += len('</a:r>')
        segments.append(('text', xml[last:start]))
        segments.append(('run', xml[start:end]))
        last = end
    segments.append(('text', xml[last:]))
    return segments

def get_run_text(run):
    m = re.search(r'<a:t[^>]*>(.*?)</a:t>', run, re.DOTALL)
    return m.group(1) if m else None

def set_run_text_and_sz(run, new_text, scale=False):
    esc = new_text.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
    run = re.sub(r'(<a:t[^>]*>).*?(</a:t>)', r'\g<1>' + esc + r'\g<2>', run, flags=re.DOTALL)
    if scale:
        sz_match = re.search(r'\bsz="(\d+)"', run)
        if sz_match:
            new_sz = smart_sz(new_text, sz_match.group(1))
            run = run[:sz_match.start()] + f'sz="{new_sz}"' + run[sz_match.end():]
    return run

def replace_in_runs(xml, targets, scale=False):
    segs = split_runs(xml)
    for old, new in targets:
        esc_old = old.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
        for i, (kind, seg) in enumerate(segs):
            if kind != 'run': continue
            txt = get_run_text(seg)
            if txt is not None and (txt == esc_old or txt == old):
                segs[i] = ('run', set_run_text_and_sz(seg, new, scale=scale))
                break
    return ''.join(s for _, s in segs)

def rt(xml, old, new, count=1):
    esc_new = new.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
    esc_old = old.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
    pat = r'(<a:t[^>]*>)' + re.escape(esc_old) + r'(</a:t>)'
    result, n = re.subn(pat, r'\g<1>' + esc_new + r'\g<2>', xml, count=count)
    if n == 0:
        pat2 = r'(<a:t[^>]*>)' + re.escape(old) + r'(</a:t>)'
        result, _ = re.subn(pat2, r'\g<1>' + esc_new + r'\g<2>', xml, count=count)
        return result
    return result

def read_slide(n, base):
    with open(f'{base}/ppt/slides/slide{n}.xml', encoding='utf-8') as f:
        return f.read()

def write_slide(n, content, base):
    with open(f'{base}/ppt/slides/slide{n}.xml', 'w', encoding='utf-8') as f:
        f.write(content)

# ── Field accessor — works with BOTH flat Notion fields and nested dicts ────
def get(d, *keys, default=''):
    """
    Accepts either:
      - Flat Notion format: get(d, 'Customer Name')
      - Nested format:      get(d, 'account', 'customer')
    Falls back to default if not found or empty.
    """
    # Try flat lookup first (Notion direct export)
    if len(keys) == 1:
        return d.get(keys[0]) or default
    # Try nested lookup (legacy format)
    obj = d
    for k in keys:
        if isinstance(obj, dict):
            obj = obj.get(k, {})
        else:
            return default
    return obj or default

# ── Trend narrative map ─────────────────────────────────────────────────────
TREND_MAP = {
    'Increase YTD':  'Year-to-date figures indicate an increase in duplicate invoices',
    'Decrease YTD':  'Year-to-date figures show a reduction in duplicate invoices',
    'Stable':        'Duplicate invoice volumes have remained consistent month-on-month',
}

# ── Main populate function ──────────────────────────────────────────────────
def populate(data, base):
    """
    data can be:
      - Flat dict with Notion property names (from form submission)
      - Nested dict with sections (legacy format)
    Both are supported.
    """
    # ── Normalise: support both flat Notion and legacy nested format ──
    is_flat = 'Customer Name' in data

    def f(notion_key, *nested_keys, default=''):
        if is_flat:
            return data.get(notion_key) or default
        else:
            obj = data
            for k in nested_keys:
                obj = obj.get(k, {}) if isinstance(obj, dict) else {}
            return obj or default

    # Core fields
    customer       = f('Customer Name',        'account', 'customer')      or '{CUSTOMER}'
    ebr_date       = f('EBR Date',             'account', 'ebr_date')      or 'DATE'
    prev_ebr_date  = f('Previous EBR Date',    'account', 'prev_ebr_date') or ''
    am             = f('Account Manager',      'account', 'am')            or ''
    contact        = f('Customer Primary Contact', 'account', 'contact')   or ''

    # ── SLIDE 1: Title ──────────────────────────────────────────────────────
    s = read_slide(1, base)
    s = rt(s, 'DATE', ebr_date)
    s = rt(s, '{CUSTOMER}', customer)
    write_slide(1, s, base)

    # ── SLIDE 4: Opening quote ──────────────────────────────────────────────
    s = read_slide(4, base)
    quote      = f('Customer Quote',   'recap', 'quote')
    quote_attr = f('Quote Attribution','recap', 'quote_attr')
    if quote:
        s = rt(s, 'Quote text goes here. Replace with a compelling testimonial or key insight.', quote)
    if quote_attr:
        parts = quote_attr.split(',', 1)
        s = rt(s, 'FULL NAME', parts[0].strip())
        if len(parts) > 1:
            s = rt(s, 'Title, Company Name', parts[1].strip())
    write_slide(4, s, base)

    # ── SLIDE 5: Recap points ───────────────────────────────────────────────
    s = read_slide(5, base)
    if is_flat:
        recap_points = [
            (f('Recap Point 1'), f('Recap Point 1 Detail')),
            (f('Recap Point 2'), f('Recap Point 2 Detail')),
            (f('Recap Point 3'), f('Recap Point 3 Detail')),
            (f('Xelix Commitments'), f('Customer Commitments')),
            (f('Carry Forward Items'), f('Progress Rating')),
        ]
    else:
        pts = data.get('recap', {}).get('points', [])
        recap_points = [(pt.get('h',''), pt.get('d','')) for pt in pts[:5]]

    for i, (h, d) in enumerate(recap_points[:5], 1):
        if h: s = rt(s, f'POINT {i}', h)
        if d: s = rt(s, 'Supporting detail line 1', d)
    write_slide(5, s, base)

    # ── SLIDE 7: Xelix company updates ─────────────────────────────────────
    s = read_slide(7, base)
    xelix_updates_raw = f('Xelix Company Updates', 'exec', 'xelix_updates')
    updates = [u.strip() for u in xelix_updates_raw.split('\n') if u.strip()]
    for i in range(1, 7):
        if i - 1 < len(updates):
            s = rt(s, f'Feature {i} title', updates[i-1])
    write_slide(7, s, base)

    # ── SLIDE 8: Customer company updates ───────────────────────────────────
    s = read_slide(8, base)
    s = rt(s, '{CUSTOMER} COMPANY UPDATES', f'{customer} COMPANY UPDATES')
    cust_updates = f('Customer Company Updates', 'exec', 'customer_updates')
    if cust_updates:
        s = rt(s, ' ', cust_updates)
    write_slide(8, s, base)

    # ── SLIDE 9: High-level snapshot ───────────────────────────────────────
    s = read_slide(9, base)
    s = replace_in_runs(s, [
        ('XX', f('Monthly Active Users',  'exec', 'mau')            or 'XX'),
        ('XX', f('Average Weekly Users',  'exec', 'awu')            or 'XX'),
        ('XX', f('Busiest Module',        'exec', 'busiest_module') or 'XX'),
    ], scale=True)
    write_slide(9, s, base)

    # ── SLIDE 10: Transactions snapshot ────────────────────────────────────
    s = read_slide(10, base)
    s = replace_in_runs(s, [
        ('£', f('Dup Value Confirmed (YTD)',       'exec', 'dup_val')        or '£'),
        ('X', f('Dup Caught Ahead of Pay Run (#)', 'exec', 'dup_payrun_ct') or 'X'),
        ('£', f('Dup Caught Ahead of Pay Run (£)', 'exec', 'dup_payrun_val')or '£'),
        ('X', f('Invoice Errors Ahead of Pay Run', 'exec', 'inv_err')       or 'X'),
        ('X', f('Most Common Error Type',          'exec', 'err_type')      or 'X'),
        ('X', f('Most Active User - Transactions', 'exec', 'tx_user')       or 'X'),
    ], scale=True)
    write_slide(10, s, base)

    # ── SLIDE 11: Helpdesk snapshot ────────────────────────────────────────
    s = read_slide(11, base)
    s = replace_in_runs(s, [
        ('X',  f('Avg Handling Time - Current',  'exec', 'aht')         or 'X'),
        ('X%', f('% Handling Time Decrease',     'exec', 'aht_dec')     or 'X%'),
        ('X%', f('% Tickets via Gen AI',         'exec', 'genai_pct')   or 'X%'),
        ('X%', f('% Tickets via Triggers',       'exec', 'trigger_pct') or 'X%'),
        ('X',  f('Most Active User - Helpdesk',  'exec', 'hd_user')     or 'X'),
    ], scale=True)
    write_slide(11, s, base)

    # ── SLIDE 12: Statements snapshot ──────────────────────────────────────
    s = read_slide(12, base)
    s = replace_in_runs(s, [
        ('X%', f('% Spend Reconciled',           'exec', 'spend_rec')  or 'X%'),
        ('X%', f('% Invoices Reconciled',        'exec', 'inv_rec')    or 'X%'),
        ('X%', f('% Vendors Reconciled',         'exec', 'vend_rec')   or 'X%'),
        ('X%', f('% Statements Reconciled',      'exec', 'stmt_rec')   or 'X%'),
        ('£',  f('Missing Credits Recovered (£)','exec', 'credits')    or '£'),
        ('X%', f('% Statements Automated',       'exec', 'stmt_auto')  or 'X%'),
        ('X',  f('Most Active User - Statements','exec', 'stmt_user')  or 'X'),
    ], scale=True)
    write_slide(12, s, base)

    # ── SLIDE 14: Transactions customer value ───────────────────────────────
    s = read_slide(14, base)
    if is_flat:
        g1     = f('Goal 1')
        g1k    = f('Goal 1 Target')
        g1p    = ''
        hist_dup = f('Value Recovered - Hist Dupes')
        hist_err = f('Value Recovered - Hist Errors')
    else:
        goals    = data.get('transactions', {}).get('goals', [])
        g1       = goals[0].get('g','') if goals else ''
        g1k      = goals[0].get('k','') if goals else ''
        g1p      = goals[0].get('p','') if goals else ''
        hist_dup = data.get('transactions', {}).get('hist_dup', '')
        hist_err = data.get('transactions', {}).get('hist_err', '')

    if g1:  s = rt(s, 'Increase Statement reconciliation coverage', g1)
    if g1k: s = rt(s, 'Reduction in audit resource', g1k)
    if g1p: s = rt(s, 'XX%', g1p)
    if hist_dup: s = rt(s, '£XXXX', hist_dup)
    if hist_err: s = rt(s, '£XXXX', hist_err)
    write_slide(14, s, base)

    # ── SLIDE 15: Duplicate invoices chart ─────────────────────────────────
    s = read_slide(15, base)
    trend_raw = f('Dup Trend Direction', 'transactions', 'trend')
    trend = TREND_MAP.get(trend_raw, trend_raw) or 'Year-to-date figures indicate an increase in duplicate invoices'
    s = rt(s, 'Year-to-date figures indicate an increase in duplicate invoices', trend)
    dup_ct    = f('Dup Count Confirmed (YTD)', 'transactions', 'dup_ct')
    dup_worth = f('Dup Value Confirmed (YTD)', 'transactions', 'dup_worth')
    if dup_ct:    s = rt(s, 'X', dup_ct)
    if dup_worth: s = rt(s, '£', dup_worth)
    write_slide(15, s, base)

    # ── SLIDE 17: Detected ahead of payrun ─────────────────────────────────
    s = read_slide(17, base)
    payrun = f('Dup Caught Ahead of Pay Run (#)', 'transactions', 'inv_payrun') or 'X'
    err_ct = f('Inv Errors Confirmed (Count)',    'transactions', 'inv_err_ct') or 'X'
    s = rt(s, 'X detected ahead of the payment run', f'{payrun} detected ahead of the payment run')
    s = rt(s, 'X CORRECTED ', f'{err_ct} CORRECTED ')
    write_slide(17, s, base)

    # ── SLIDE 22: Statements customer value ────────────────────────────────
    s = read_slide(22, base)
    if is_flat:
        sg1  = f('Goal 2')
        sg1k = f('Goal 2 Target')
        sg1p = ''
        miss = f('Missed Credits Count')
    else:
        sgoals = data.get('statements', {}).get('goals', [])
        sg1    = sgoals[0].get('g','') if sgoals else ''
        sg1k   = sgoals[0].get('k','') if sgoals else ''
        sg1p   = sgoals[0].get('p','') if sgoals else ''
        miss   = data.get('statements', {}).get('miss_cred_ct', '')

    if sg1:  s = rt(s, 'Increase Statement reconciliation coverage', sg1)
    if sg1k: s = rt(s, '% of spend being reconciled on a monthly/quarterly/annual basis ', sg1k)
    if sg1p: s = rt(s, 'XX%', sg1p)
    if miss: s = rt(s, 'XXX', miss)
    write_slide(22, s, base)

    # ── SLIDE 23: Reconciliation growth ────────────────────────────────────
    s = read_slide(23, base)
    miss_cred   = f('Missed Credits Count',     'statements', 'miss_cred_ct')
    vend_raw    = f('Vendors % of Total Spend', 'statements', 'vend_spend_pct')
    if miss_cred: s = rt(s, 'X', miss_cred)
    if vend_raw:
        m = re.search(r'(\d+)[^\d]*(\d+%)', vend_raw)
        if m:
            s = rt(s, 'X', m.group(1))
            s = rt(s, 'X%', m.group(2))
    write_slide(23, s, base)

    # ── SLIDE 25: Invoice coverage ──────────────────────────────────────────
    s = read_slide(25, base)
    inv_cov = f('Invoice Coverage %', 'statements', 'inv_coverage')
    if inv_cov: s = rt(s, 'X', inv_cov)
    write_slide(25, s, base)

    # ── SLIDE 26: Automation evolution ─────────────────────────────────────
    s = read_slide(26, base)
    auto_prev = f('Full Automation % Previous EBR', 'statements', 'auto_prev') or 'Data from previous EBR'
    auto_now  = f('Full Automation % Now',          'statements', 'auto_now')  or 'Statement Automation Now'
    s = rt(s, 'Data from previous EBR', f'Previous EBR: {auto_prev}')
    s = rt(s, 'Statement Automation Now', f'Now: {auto_now}')
    write_slide(26, s, base)

    # ── SLIDE 29: Helpdesk customer value ──────────────────────────────────
    s = read_slide(29, base)
    aht_curr = f('Avg Handling Time - Current', 'helpdesk', 'aht_curr') or 'X'
    aht_prev = f('Avg Handling Time - Previous','helpdesk', 'aht_prev') or 'X'
    if is_flat:
        hg1  = f('Goal 3')
        hg1k = f('Goal 3 Target')
    else:
        hgoals = data.get('helpdesk', {}).get('goals', [])
        hg1    = hgoals[0].get('g','') if hgoals else ''
        hg1k   = hgoals[0].get('k','') if hgoals else ''

    if hg1:  s = rt(s, 'Spend less FTE time managing vendor queries', hg1)
    if hg1k: s = rt(s, 'Average handling time decreased from A -&gt; B', hg1k)
    s = rt(s, ' X - X', f' {aht_prev} -> {aht_curr}')
    write_slide(29, s, base)

    # ── SLIDE 30: Gen AI helpdesk ───────────────────────────────────────────
    s = read_slide(30, base)
    genai_deep = f('Gen AI % Deep Dive', 'helpdesk', 'genai_deep')
    if genai_deep: s = rt(s, 'X%', genai_deep)
    write_slide(30, s, base)

    # ── SLIDE 35: Support ticket volumes ───────────────────────────────────
    s = read_slide(35, base)
    s = rt(s, '{CUSTOMER}', customer)
    tkt_new    = f('New Tickets Raised',           'helpdesk', 'tkt_new')
    tkt_open   = f('Open Tickets',                 'helpdesk', 'tkt_open')
    tkt_xelix  = f('Tickets Waiting on Xelix',     'helpdesk', 'tkt_xelix')
    tkt_cust   = f('Tickets Waiting on Customer',  'helpdesk', 'tkt_cust')
    tkt_closed = f('Closed Tickets',               'helpdesk', 'tkt_closed')
    for val in [tkt_new, tkt_new, tkt_open, tkt_xelix, tkt_cust, tkt_closed]:
        if val: s = rt(s, 'X', val)
    write_slide(35, s, base)

    # ── SLIDE 37: Next steps timeline ──────────────────────────────────────
    s = read_slide(37, base)
    if is_flat:
        steps = [
            (f('Action 1'), f('Action 1 Owner')),
            (f('Action 2'), f('Action 2 Owner')),
            (f('Action 3'), f('Action 3 Owner')),
            (f('Action 4'), f('Action 4 Owner')),
            (f('Action 5'), f('Action 5 Owner')),
        ]
    else:
        raw_steps = data.get('next_steps', {}).get('steps', [])
        steps = [(st.get('s',''), st.get('d','')) for st in raw_steps[:5]]

    labels = ['STEP 01','STEP 02','STEP 03','STEP 04','STEP 05']
    for i, (step_title, step_desc) in enumerate(steps[:5]):
        if step_title: s = rt(s, labels[i], step_title)
        if step_desc:  s = rt(s, 'Description text for this step', step_desc)
    write_slide(37, s, base)

    # ── SLIDE 38: Renewal ──────────────────────────────────────────────────
    s = read_slide(38, base)
    rd = f('Renewal Display Date', 'next_steps', 'renewal_display')
    rs = f('Renewal Stakeholders', 'next_steps', 'renewal_stakeholders')
    if rd:
        parts = rd.split()
        if len(parts) >= 3:
            s = rt(s, 'DAY ',   parts[0] + ' ')
            s = rt(s, 'MONTH ', parts[1] + ' ')
            s = rt(s, '2026',   parts[-1])
        elif len(parts) == 2:
            s = rt(s, 'DAY ',   parts[0] + ' ')
            s = rt(s, 'MONTH ', parts[1] + ' ')
    if rs: s = rt(s, ' ', rs)
    write_slide(38, s, base)

    # ── SLIDE 39: Scope ────────────────────────────────────────────────────
    s = read_slide(39, base)
    regions = f('Live Regions', 'next_steps', 'regions')
    erp     = f('ERP Systems',  'next_steps', 'erp')
    if regions:
        rl = [r.strip() for r in re.split(r'[,\n]', regions) if r.strip()]
        if len(rl) >= 1: s = rt(s, 'North America', rl[0])
        if len(rl) >= 2: s = rt(s, 'Europe', rl[1])
    write_slide(39, s, base)

    return True


# ═══════════════════════════════════════════════════════════════════════════
# CHART GENERATION — appended to populate_ebr.py
# Requires: matplotlib (pip install matplotlib)
# ═══════════════════════════════════════════════════════════════════════════

def _charts_available():
    try:
        import matplotlib
        return True
    except ImportError:
        return False

def generate_and_embed_charts(data, base, tmp_dir):
    """Generate PNG charts and embed them into slides 15 and 26."""
    if not _charts_available():
        print("WARNING: matplotlib not available, skipping charts")
        return

    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches
    import re, shutil

    DARK_BG  = '#1A0533'
    PINK     = '#E91E8C'
    PURPLE   = '#7B2D8B'
    GRAY     = '#CCBBFF'

    is_flat = 'Customer Name' in data

    def f(notion_key, *nested_keys, default=''):
        if is_flat:
            return data.get(notion_key) or default
        obj = data
        for k in nested_keys:
            obj = obj.get(k, {}) if isinstance(obj, dict) else {}
        return obj or default

    def save(fig, path):
        fig.savefig(path, dpi=180, bbox_inches='tight',
                    facecolor=fig.get_facecolor(), edgecolor='none')
        plt.close(fig)

    # ── helpers ──────────────────────────────────────────────────────────

    def next_rId(rels_path):
        with open(rels_path) as f:
            c = f.read()
        ids = [int(m) for m in re.findall(r'Id="rId(\d+)"', c)]
        return f'rId{max(ids)+1}' if ids else 'rId1'

    def add_rel(rels_path, rId, fname):
        with open(rels_path) as f:
            c = f.read()
        entry = (f'<Relationship Id="{rId}" '
                 f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                 f'Target="../media/{fname}"/>')
        c = c.replace('</Relationships>', entry + '</Relationships>')
        with open(rels_path, 'w') as f:
            f.write(c)

    def copy_media(src, base, fname):
        media = os.path.join(base, 'ppt', 'media')
        os.makedirs(media, exist_ok=True)
        shutil.copy(src, os.path.join(media, fname))

    def ensure_png_ct(base):
        ct = os.path.join(base, '[Content_Types].xml')
        with open(ct) as f:
            c = f.read()
        png = '<Default Extension="png" ContentType="image/png"/>'
        if png not in c:
            c = c.replace('</Types>', png + '</Types>')
            with open(ct, 'w') as f:
                f.write(c)

    def pic_xml(rId, x, y, cx, cy, pid):
        return (f'<p:pic>'
                f'<p:nvPicPr>'
                f'<p:cNvPr id="{pid}" name="ChartImg{pid}"/>'
                f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
                f'<p:nvPr/></p:nvPicPr>'
                f'<p:blipFill>'
                f'<a:blip r:embed="{rId}"/>'
                f'<a:stretch><a:fillRect/></a:stretch>'
                f'</p:blipFill>'
                f'<p:spPr>'
                f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
                f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
                f'</p:spPr></p:pic>')

    def embed(slide_n, src_png, fname, x, y, cx, cy, pid):
        slide_path = os.path.join(base, 'ppt', 'slides', f'slide{slide_n}.xml')
        rels_path  = os.path.join(base, 'ppt', 'slides', '_rels', f'slide{slide_n}.xml.rels')
        copy_media(src_png, base, fname)
        ensure_png_ct(base)
        rId = next_rId(rels_path)
        add_rel(rels_path, rId, fname)
        with open(slide_path, encoding='utf-8') as f:
            xml = f.read()
        if 'xmlns:r=' not in xml:
            xml = xml.replace('<p:sld ', '<p:sld xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships/2006/relationships" ')
        # Replace first matching INSERT IMAGE HERE placeholder shape
        pat = r'<p:sp\b(?:(?!<p:sp\b).)*?INSERT IMAGE HERE(?:(?!<p:sp\b).)*?</p:sp>'
        xml = re.sub(pat, pic_xml(rId, x, y, cx, cy, pid), xml, count=1, flags=re.DOTALL)
        with open(slide_path, 'w', encoding='utf-8') as f:
            f.write(xml)

    # ── SLIDE 15: Dup trend stat card ────────────────────────────────────
    dup_ct    = f('Dup Count Confirmed (YTD)',  'transactions', 'dup_ct')    or 'X'
    dup_worth = f('Dup Value Confirmed (YTD)',  'transactions', 'dup_worth') or '£'
    trend_raw = f('Dup Trend Direction',         'transactions', 'trend')    or 'Increase YTD'

    if trend_raw == 'Decrease YTD':
        arrow, arrow_col, trend_label = '↓', '#4CAF50', 'Decrease YTD'
    elif trend_raw == 'Stable':
        arrow, arrow_col, trend_label = '→', GRAY,     'Stable'
    else:
        arrow, arrow_col, trend_label = '↑', PINK,     'Increase YTD'

    fig, ax = plt.subplots(figsize=(7, 5.5), facecolor=DARK_BG)
    ax.set_facecolor(DARK_BG); ax.axis('off')
    ax.text(0.5, 0.92, arrow,       transform=ax.transAxes, fontsize=72,
            color=arrow_col, ha='center', va='top', fontweight='bold')
    ax.text(0.5, 0.70, trend_label, transform=ax.transAxes, fontsize=16,
            color=arrow_col, ha='center', va='top', fontweight='bold', fontstyle='italic')
    ax.axhline(y=0.60, xmin=0.1, xmax=0.9, color=GRAY, alpha=0.3, linewidth=1)
    ax.text(0.25, 0.55, str(dup_ct),    transform=ax.transAxes, fontsize=48,
            color=PINK, ha='center', va='top', fontweight='bold')
    ax.text(0.25, 0.22, 'Duplicates\nconfirmed', transform=ax.transAxes, fontsize=11,
            color=GRAY, ha='center', va='top', linespacing=1.4)
    ax.axvline(x=0.5, ymin=0.05, ymax=0.58, color=GRAY, alpha=0.3, linewidth=1)
    ax.text(0.75, 0.55, str(dup_worth), transform=ax.transAxes, fontsize=38,
            color=PINK, ha='center', va='top', fontweight='bold')
    ax.text(0.75, 0.22, 'Worth of\nduplicates', transform=ax.transAxes, fontsize=11,
            color=GRAY, ha='center', va='top', linespacing=1.4)
    dup_png = os.path.join(tmp_dir, 'chart_dup_trend.png')
    save(fig, dup_png)
    # x=260000,y=702452, cx=6912131,cy=5557606
    embed(15, dup_png, 'chart_dup_trend.png', 260000, 702452, 6912131, 5557606, 500)

    # ── SLIDE 26: Automation comparison bars ─────────────────────────────
    def parse_pct(v):
        try: return float(str(v).replace('%','').strip())
        except: return 0.0

    prev_pct = parse_pct(f('Full Automation % Previous EBR', 'statements', 'auto_prev'))
    now_pct  = parse_pct(f('Full Automation % Now',          'statements', 'auto_now'))

    def make_gauge(pct, subtitle, color, out_path):
        fig2, ax2 = plt.subplots(figsize=(5.5, 3.8), facecolor=DARK_BG)
        ax2.set_facecolor(DARK_BG); ax2.axis('off')
        bg = mpatches.FancyBboxPatch((0.05,0.35),0.90,0.22,
             boxstyle='round,pad=0.02', facecolor='#2D1550', edgecolor='none')
        ax2.add_patch(bg)
        fw = 0.90 * min(pct/100, 1.0)
        if fw > 0:
            bar = mpatches.FancyBboxPatch((0.05,0.35),fw,0.22,
                  boxstyle='round,pad=0.02', facecolor=color, edgecolor='none')
            ax2.add_patch(bar)
        ax2.text(0.5, 0.82, f'{pct:.0f}%', transform=ax2.transAxes,
                 fontsize=52, color='white', ha='center', va='top', fontweight='bold')
        ax2.text(0.5, 0.30, 'Full automation rate', transform=ax2.transAxes,
                 fontsize=12, color=GRAY, ha='center', va='top')
        ax2.text(0.5, 0.14, subtitle, transform=ax2.transAxes,
                 fontsize=10, color=GRAY, ha='center', va='top', alpha=0.7)
        ax2.set_xlim(0,1); ax2.set_ylim(0,1)
        save(fig2, out_path)

    prev_png = os.path.join(tmp_dir, 'chart_auto_prev.png')
    now_png  = os.path.join(tmp_dir, 'chart_auto_now.png')
    make_gauge(prev_pct, 'Data from previous EBR', PURPLE, prev_png)
    make_gauge(now_pct,  'Current period',          PINK,   now_png)
    # Left placeholder:  x=507999,y=1977483, cx=5426269,cy=3681922
    # Right placeholder: x=6178938,y=1977483, cx=5426269,cy=3681922
    embed(26, prev_png, 'chart_auto_prev.png', 507999,  1977483, 5426269, 3681922, 501)
    embed(26, now_png,  'chart_auto_now.png',  6178938, 1977483, 5426269, 3681922, 502)

    print("Charts embedded: slides 15, 26")

# ── Entry point ─────────────────────────────────────────────────────────────
if __name__ == '__main__':
    data = json.loads(sys.stdin.read())

    base_pptx  = '/mnt/user-data/uploads/EBR_WIP.pptx'
    tmp_dir    = tempfile.mkdtemp()
    unpacked   = os.path.join(tmp_dir, 'unpacked')
    customer   = (data.get('Customer Name') or data.get('account', {}).get('customer') or 'EBR').replace(' ', '_')
    output     = f'/mnt/user-data/outputs/EBR_{customer}.pptx'

    try:
        subprocess.run(
            ['python3', '/mnt/skills/public/pptx/scripts/office/unpack.py', base_pptx, unpacked],
            check=True, capture_output=True
        )
        populate(data, unpacked)
        generate_and_embed_charts(data, unpacked, tmp_dir)
        subprocess.run(
            ['python3', '/mnt/skills/public/pptx/scripts/clean.py', unpacked],
            check=True, capture_output=True
        )
        subprocess.run(
            ['python3', '/mnt/skills/public/pptx/scripts/office/pack.py',
             unpacked, output, '--original', base_pptx],
            check=True, capture_output=True
        )
        print(f'SUCCESS:{output}')
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

