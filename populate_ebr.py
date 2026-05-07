#!/usr/bin/env python3
"""
populate_ebr.py
Reads a JSON payload of EBR data and replaces all {{PLACEHOLDER}} tokens
in EBR_Template.pptx, then writes the result to stdout as base64.

Usage (called by the GitHub Pages generator via Anthropic API):
    echo '<json>' | python3 populate_ebr.py

Field names match the verified Notion database schema exactly.
Last updated: May 2026
"""

import sys
import json
import re
import os
import base64
import tempfile

# ─── PLACEHOLDER MAP ─────────────────────────────────────────────────────────
# Maps {{TOKEN}} → exact Notion property name
# WARNING: Notion field names are case-sensitive and some have trailing spaces.
# Do not "tidy" these strings — they must match the database exactly.

FIELD_MAP = {
    # ── Cover / header ────────────────────────────────────────────────────────
    "CUSTOMER_NAME":            "Customer",                         # relation field — proxy reads display name
    "ACCOUNT_MANAGER":          "Account Manager ",                 # trailing space — match exactly
    "EBR_DATE":                 "EBR Date",
    "EBR_PERIOD":               "EBR Period ? ",                    # trailing space
    "PREVIOUS_EBR_DATE":        "Previous EBR Date ? ",             # trailing space
    "CONTRACT_VALUE":           "Contract Value (ARR)",
    "RENEWAL_DISPLAY_DATE":     "Contract Renewal Date",
    "RENEWAL_STAKEHOLDERS":     "Renewal Stakeholders",
    "LIVE_REGIONS":             "Live Regions",
    "ERP_SYSTEMS":              "ERP Systems",                      # rollup — proxy resolves
    "CUSTOMER_PRIMARY_CONTACT": "Customer Primary Contact",
    "EXECUTIVE_SPONSOR":        "Executive Sponsor ",               # trailing space

    # ── Recap ─────────────────────────────────────────────────────────────────
    "CUSTOMER_QUOTE":           "Customer Quote",
    "QUOTE_ATTRIBUTION":        "Quote Attribution",
    "RECAP_POINT_1":            "Recap Point 1",
    "RECAP_POINT_1_DETAIL":     "Recap Point 1 Detail",
    "RECAP_POINT_2":            "Recap Point 2",
    "RECAP_POINT_2_DETAIL":     "Recap Point 2 Detail",
    "RECAP_POINT_3":            "Recap Point 3",
    "RECAP_POINT_3_DETAIL":     "Recap Point 3 Detail",

    # ── Summary & Recommendations (slides 20, 27, 33) ─────────────────────────
    # These four fields are combined into one text block on each summary slide.
    # See build_replacements() for how they are merged.
    "XELIX_COMMITMENTS":        "Xelix Commitments",
    "CUSTOMER_COMMITMENTS":     "Customer Commitments",
    "RISKS_OR_COMPLAINTS":      "Risks or Complaints",
    "RECOMMENDED_ACTION":       "Recommended Action",

    # ── Carry forward / progress (legacy fields kept for template compat) ─────
    "CARRY_FORWARD_ITEMS":      "Carry Forward Items",
    "PROGRESS_RATING":          "Progress Rating",

    # ── Company updates ───────────────────────────────────────────────────────
    "XELIX_COMPANY_UPDATES":    "Xelix Company Updates",
    "CUSTOMER_COMPANY_UPDATES": "Customer Company Updates (1)",     # form uses (1) version

    # ── Platform snapshot ─────────────────────────────────────────────────────
    "MONTHLY_ACTIVE_USERS":             "Monthly Active Users",
    "AVERAGE_WEEKLY_USERS":             "Average Weekly Users",
    "BUSIEST_MODULE":                   "Busiest Module",
    "PCT_AIV_USED_YTD":                 "% AIV Used YTD",
    "MOST_ACTIVE_USER_TRANSACTIONS":    "Most Active User - Transactions",
    "MOST_ACTIVE_USER_STATEMENTS":      "Most Active User - Statements",
    "MOST_ACTIVE_USER_HELPDESK":        "Most Active User - Helpdesk",

    # ── Transactions ──────────────────────────────────────────────────────────
    "DUP_VALUE_CONFIRMED":          "Dup Value Confirmed (YTD) - Transactions",
    "DUP_COUNT_CONFIRMED":          "Dup Count Confirmed (YTD) - Transactions",
    "DUP_CAUGHT_AHEAD_NUM":         "Dup Caught Ahead of Pay Run (#)- Transactions",   # note: no space before -
    "DUP_CAUGHT_AHEAD_VAL":         "Dup Caught Ahead of Pay Run (\u00a3) - Transactions",
    "DUP_TREND_DIRECTION":          "Dup Trend Direction - Transactions",
    "DUP_TREND_NARRATIVE":          "Dup Trend Direction - Transactions",               # converted to sentence below
    "VALUE_RECOVERED_HIST_DUPES":   "Value Recovered - Historical Dupes - Transactions",
    "VALUE_RECOVERED_HIST_ERRORS":  "Value Recovered - Historical Errors - Transactions",
    "INV_ERRORS_CONFIRMED_COUNT":   "Inv Errors Confirmed (Count) - Transactions",
    "INVOICE_ERRORS_AHEAD_PAY_RUN": "Invoice Errors Ahead of Pay Run - Transactions",
    "MOST_COMMON_ERROR_TYPE":       "Most Common Error Type - Transactions",
    "ERROR_TYPE_COUNT":             "Error Type Count - Transactions",

    # ── Statements ────────────────────────────────────────────────────────────
    "PCT_SPEND_RECONCILED":         "% Spend Reconciled - Statements",
    "PCT_INVOICES_RECONCILED":      "% Invoices Reconciled",
    "PCT_VENDORS_RECONCILED":       "% Vendors Reconciled - Statements",
    "PCT_STATEMENTS_RECONCILED":    "% Statements Reconciled ",                        # trailing space
    "PCT_STATEMENTS_AUTOMATED":     "% Statements Automated",
    "FULL_AUTOMATION_NOW":          "Full Automation % Now - Statements ",              # trailing space
    "FULL_AUTOMATION_PREV":         "Full Automation %  Previous EBR",                 # double space
    "INVOICE_COVERAGE_PCT":         "Invoice Coverage % - Statements ",                # trailing space
    "INDUSTRY_BENCHMARK":           "Industry Benchmark - Statements ",                # trailing space
    "MISSING_CREDITS_RECOVERED":    "Missing Credits Recovered (\u00a3) - Statements",
    "MISSED_CREDITS_COUNT":         "Missed Credits Count - Statements",
    "VENDORS_PCT_TOTAL_SPEND":      "Vendors % of Total Spend - Statements",
    "MOM_RECONCILIATION_GROWTH":    "MoM Reconciliation Growth - Statements",

    # ── Helpdesk ──────────────────────────────────────────────────────────────
    "NEW_TICKETS_RAISED":           "New Tickets Raised - Helpdesk",
    "OPEN_TICKETS":                 "Open Tickets - Helpdesk",
    "CLOSED_TICKETS":               "Closed Tickets - Helpdesk",
    "TICKETS_WAITING_XELIX":        "Tickets Waiting on Xelix - Helpdesk ",            # trailing space
    "TICKETS_WAITING_CUSTOMER":     "Tickets Waiting on Customer - Helpdesk",
    "TOTAL_TICKETS_YTD":            "Total Tickets (YTD) - Helpdesk",
    "AVG_HANDLING_TIME_CURRENT":    "Avg Handling Time - Current - Helpdesk",
    "AVG_HANDLING_TIME_PREV":       "Avg Handling Time - Previous - Helpdesk",
    "PCT_HANDLING_TIME_DECREASE":   "% Handling Time Decrease - Helpdesk",
    "PCT_TICKETS_VIA_GEN_AI":       "% Tickets via Gen AI - Helpdesk",
    "GEN_AI_PCT_DEEP_DIVE":         "Gen AI % Deep Dive - Helpdesk",
    "PCT_TICKETS_VIA_TRIGGERS":     "% Tickets via Triggers - Helpdesk",
    "SLA_COMPLIANCE_RATE":          "SLA compliance rate (%) - Helpdesk ",             # trailing space

    # ── Goals & KPIs (multi-select — joined as bullet list) ───────────────────
    # Form uses (1) versions for Statements, Helpdesk, Vendors Goals
    "TRANSACTIONS_GOALS":   "Transactions Goals",
    "TRANSACTIONS_KPIS":    "Transactions KPIs",
    "STATEMENTS_GOALS":     "Statements Goals (1)",
    "STATEMENTS_KPIS":      "Statements KPIs",
    "HELPDESK_GOALS":       "Helpdesk Goals (1)",
    "HELPDESK_KPIS":        "Helpdesk KPIs",
    "VENDORS_GOALS":        "Vendors Goals (1)",
    "VENDORS_KPIS":         "Vendors KPIs",

    # ── Voice of customer / health ────────────────────────────────────────────
    "CUSTOMER_SENTIMENT":       "Customer Sentiment",
    "TOP_THEME_FROM_CALLS":     "Top Theme from Calls",
    "KEY_CUSTOMER_QUOTE":       "Key Customer Quote",
    "KEY_NOTES_LAST_TOUCHPOINT":"Key Notes from Last Touchpoint",
    "OVERALL_RISK_RATING":      "Overall Risk Rating",
    "HEALTH_TREND":             "Health Trend",
    "CHURN_RISK_FLAG":          "Churn Risk Flag",

    # ── Actions (top 5 used in deck; action 6 tracked in DB only) ─────────────
    "ACTION_1":        "Action 1",
    "ACTION_1_OWNER":  "Action 1 Owner",
    "ACTION_1_DUE":    "Action 1 Due",
    "ACTION_1_STATUS": "Action 1 Status",
    "ACTION_2":        "Action 2",
    "ACTION_2_OWNER":  "Action 2 Owner",
    "ACTION_2_DUE":    "Action 2 Due",
    "ACTION_2_STATUS": "Action 2 Status",
    "ACTION_3":        "Action 3",
    "ACTION_3_OWNER":  "Action 3 Owner",
    "ACTION_3_DUE":    "Action 3 Due",
    "ACTION_3_STATUS": "Action 3 Status",
    "ACTION_4":        "Action 4",
    "ACTION_4_OWNER":  "Action 4 Owner",
    "ACTION_4_DUE":    "Action 4 Due",
    "ACTION_4_STATUS": "Action 4 Status",
    "ACTION_5":        "Action 5",
    "ACTION_5_OWNER":  "Action 5 Owner",
    "ACTION_5_DUE":    "Action 5 Due",
    "ACTION_5_STATUS": "Action 5 Status",
    # Action 6 is tracked in Notion but the deck only has 5 step shapes.
    # Kept here so the proxy can pass it through without errors.
    "ACTION_6":        "Action 6",
    "ACTION_6_OWNER":  "Action 6 Owner",
    "ACTION_6_DUE":    "Action 6 Due",
    "ACTION_6_STATUS": "Action 6 Status",
}

# ─── FIELDS NOT MAPPED TO DECK ────────────────────────────────────────────────
# These are read from Notion but never written into the PPTX.
# Kept here for reference — Phase 2 will use them for auto-population.
NOT_IN_DECK = [
    "EBR Period ? ",
    "Previous EBR Date ? ",
    "Goals set at last EBR ?",
    "Dup Count Confirmed (YTD) - Transactions",
    "Dup Value Confirmed (YTD) - Transactions",
    "Inv Errors Confirmed (Count) - Transactions",
    "Error Type Count - Transactions",
    "Value Recovered - Historical Dupes - Transactions",
    "Value Recovered - Historical Errors - Transactions",
    "Value Recovered - Duplicates -Transactions",
    "Value Recovered - Errors -Transactions",
    "Value Recovered - Duplicates",
    "Value Recovered - Errors",
]

# ─── TREND DIRECTION → NARRATIVE SENTENCE ─────────────────────────────────────
TREND_SENTENCES = {
    "Increase YTD": "Year-to-date figures indicate an increase in duplicate invoices.",
    "Decrease YTD": "The number of duplicate invoices has declined year-to-date.",
    "Stable":       "Duplicate invoice volumes have remained consistent month-on-month.",
}


def format_multiselect(val) -> str:
    """Convert a JSON array (multi-select) to a bullet list string."""
    if isinstance(val, list):
        return "\n".join(f"• {item}" for item in val if item)
    return str(val).strip() if val else ""


def format_summary_block(data: dict) -> str:
    """
    Combine the four summary fields into one text block for slides 20, 27, 33.
    Only includes sections where the AM has entered content.
    """
    parts = []
    xelix = str(data.get("Xelix Commitments", "") or "").strip()
    customer = str(data.get("Customer Commitments", "") or "").strip()
    risks = str(data.get("Risks or Complaints", "") or "").strip()
    action = str(data.get("Recommended Action", "") or "").strip()

    if xelix:
        parts.append(f"Xelix Commitments:\n{xelix}")
    if customer:
        parts.append(f"Customer Commitments:\n{customer}")
    if risks:
        parts.append(f"Risks / Complaints:\n{risks}")
    if action:
        parts.append(f"Recommended Action:\n{action}")

    return "\n\n".join(parts) if parts else "—"


def build_replacements(data: dict) -> dict:
    """Build token → value mapping from Notion property data."""
    repl = {}

    # Build the combined summary block once — used on slides 20, 27, 33
    summary_block = format_summary_block(data)

    for token, notion_key in FIELD_MAP.items():
        val = data.get(notion_key, "")
        if val is None:
            val = ""

        # Multi-select fields (Goals / KPIs) → bullet list
        if isinstance(val, list):
            val = format_multiselect(val)

        val = str(val).strip()

        # DUP_TREND_NARRATIVE → convert select value to natural language
        if token == "DUP_TREND_NARRATIVE":
            val = TREND_SENTENCES.get(val, val)

        # Summary slides — replace the combined block token if it exists in template
        if token in ("XELIX_COMMITMENTS",):
            # If template has {{SUMMARY_BLOCK}} use that; otherwise fall through
            # to individual field replacement below
            repl["{{SUMMARY_BLOCK}}"] = summary_block

        repl[f"{{{{{token}}}}}"] = val if val else "—"

    return repl


def replace_in_xml(content: str, replacements: dict) -> str:
    """
    Replace placeholder tokens in XML.
    First tries direct replacement (token in a single run).
    Then handles tokens split across multiple XML <a:t> runs.
    """
    # Pass 1: direct replacement
    for token, value in replacements.items():
        content = content.replace(token, value)

    # Pass 2: handle split tokens
    # Some PowerPoint XML editors split a token like {{FOO}} across
    # multiple <a:t> runs. We collapse adjacent text runs, replace, then
    # leave the XML as a single run (PowerPoint handles this fine).
    def collapse_runs_and_replace(xml):
        # Match a paragraph's worth of runs and collapse text
        para_pattern = re.compile(r'(<a:p\b[^>]*>)(.*?)(</a:p>)', re.DOTALL)

        def fix_para(m):
            open_tag, body, close_tag = m.group(1), m.group(2), m.group(3)
            run_pattern = re.compile(r'(<a:r\b[^>]*>)(.*?)(</a:r>)', re.DOTALL)
            runs = run_pattern.findall(body)
            if not runs:
                return m.group(0)

            # Collect all text across runs
            full_text = ""
            for _, run_body, _ in runs:
                t_match = re.search(r'<a:t[^>]*>(.*?)</a:t>', run_body, re.DOTALL)
                if t_match:
                    full_text += t_match.group(1)

            # Check if full_text contains any token
            replaced = full_text
            for token, value in replacements.items():
                replaced = replaced.replace(token, value)

            if replaced == full_text:
                return m.group(0)  # nothing changed

            # Rebuild: keep first run's rPr, replace its text, drop other runs
            first_run_match = run_pattern.search(body)
            if not first_run_match:
                return m.group(0)

            first_run_open, first_run_body, first_run_close = (
                first_run_match.group(1),
                first_run_match.group(2),
                first_run_match.group(3)
            )
            rpr_match = re.search(r'<a:rPr[^>]*/>', first_run_body)
            rpr = rpr_match.group(0) if rpr_match else ""
            new_run = f"{first_run_open}{rpr}<a:t>{replaced}</a:t>{first_run_close}"

            # Replace all runs in body with the new single run
            new_body = run_pattern.sub("", body, count=len(runs))
            new_body = new_run + new_body

            return f"{open_tag}{new_body}{close_tag}"

        return para_pattern.sub(fix_para, xml)

    content = collapse_runs_and_replace(content)
    return content


def process_pptx(template_path: str, replacements: dict, output_path: str):
    """Unpack pptx, replace tokens in all slide XMLs, repack."""
    import zipfile

    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(template_path, 'r') as z:
            z.extractall(tmpdir)

        for root, dirs, files in os.walk(tmpdir):
            for fname in files:
                if fname.endswith('.xml') or fname.endswith('.rels'):
                    fpath = os.path.join(root, fname)
                    try:
                        with open(fpath, 'r', encoding='utf-8') as f:
                            content = f.read()
                        new_content = replace_in_xml(content, replacements)
                        if new_content != content:
                            with open(fpath, 'w', encoding='utf-8') as f:
                                f.write(new_content)
                    except (UnicodeDecodeError, PermissionError):
                        pass  # skip binary files

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root, dirs, files in os.walk(tmpdir):
                for fname in files:
                    fpath = os.path.join(root, fname)
                    arcname = os.path.relpath(fpath, tmpdir)
                    zout.write(fpath, arcname)


def main():
    raw = sys.stdin.read().strip()
    if not raw:
        print("ERROR: no input", file=sys.stderr)
        sys.exit(1)

    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        print(f"ERROR: invalid JSON: {e}", file=sys.stderr)
        sys.exit(1)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(script_dir, "EBR_Template.pptx")
    if not os.path.exists(template_path):
        print(f"ERROR: EBR_Template.pptx not found at {template_path}", file=sys.stderr)
        sys.exit(1)

    replacements = build_replacements(data)

    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        output_path = tmp.name

    try:
        process_pptx(template_path, replacements, output_path)
        with open(output_path, 'rb') as f:
            b64 = base64.b64encode(f.read()).decode('utf-8')
        print(b64)
    finally:
        os.unlink(output_path)


if __name__ == "__main__":
    main()
