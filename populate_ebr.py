#!/usr/bin/env python3
"""
populate_ebr.py
Reads a JSON payload of EBR data and replaces all {{PLACEHOLDER}} tokens
in EBR_Template.pptx, then writes the result to stdout as base64.

Usage (called by the GitHub Pages generator via Anthropic API):
    echo '<json>' | python3 populate_ebr.py
"""

import sys
import json
import re
import os
import shutil
import base64
import subprocess
import tempfile

# ─── PLACEHOLDER MAP ─────────────────────────────────────────────────────────
# Maps {{TOKEN}} → Notion property name
FIELD_MAP = {
    # Cover / header fields
    "CUSTOMER_NAME":            "Customer Name",
    "ACCOUNT_MANAGER":          "Account Manager",
    "EBR_DATE":                 "EBR Date",
    "EBR_PERIOD":               "EBR Period",
    "PREVIOUS_EBR_DATE":        "Previous EBR Date",
    "CONTRACT_VALUE":           "Contract Value (ARR)",
    "RENEWAL_DISPLAY_DATE":     "Renewal Display Date",
    "RENEWAL_STAKEHOLDERS":     "Renewal Stakeholders",
    "LIVE_REGIONS":             "Live Regions",
    "ERP_SYSTEMS":              "ERP Systems",
    "CUSTOMER_PRIMARY_CONTACT": "Customer Primary Contact",
    "EXECUTIVE_SPONSOR":        "Executive Sponsor",

    # Recap
    "CUSTOMER_QUOTE":           "Customer Quote",
    "QUOTE_ATTRIBUTION":        "Quote Attribution",
    "RECAP_POINT_1":            "Recap Point 1",
    "RECAP_POINT_1_DETAIL":     "Recap Point 1 Detail",
    "RECAP_POINT_2":            "Recap Point 2",
    "RECAP_POINT_2_DETAIL":     "Recap Point 2 Detail",
    "RECAP_POINT_3":            "Recap Point 3",
    "RECAP_POINT_3_DETAIL":     "Recap Point 3 Detail",
    "XELIX_COMMITMENTS":        "Xelix Commitments",
    "CUSTOMER_COMMITMENTS":     "Customer Commitments",
    "CARRY_FORWARD_ITEMS":      "Carry Forward Items",
    "PROGRESS_RATING":          "Progress Rating",
    "XELIX_COMPANY_UPDATES":    "Xelix Company Updates",
    "CUSTOMER_COMPANY_UPDATES": "Customer Company Updates",

    # Platform snapshot
    "MONTHLY_ACTIVE_USERS":             "Monthly Active Users",
    "AVERAGE_WEEKLY_USERS":             "Average Weekly Users",
    "BUSIEST_MODULE":                   "Busiest Module",
    "PCT_AIV_USED_YTD":                 "% AIV Used YTD",
    "MOST_ACTIVE_USER_TRANSACTIONS":    "Most Active User - Transactions",
    "MOST_ACTIVE_USER_STATEMENTS":      "Most Active User - Statements",
    "MOST_ACTIVE_USER_HELPDESK":        "Most Active User - Helpdesk",

    # Transactions
    "DUP_VALUE_CONFIRMED":      "Dup Value Confirmed (YTD)",
    "DUP_COUNT_CONFIRMED":      "Dup Count Confirmed (YTD)",
    "DUP_CAUGHT_AHEAD_NUM":     "Dup Caught Ahead of Pay Run (#)",
    "DUP_CAUGHT_AHEAD_VAL":     "Dup Caught Ahead of Pay Run (£)",
    "DUP_TREND_DIRECTION":      "Dup Trend Direction",
    "DUP_TREND_NARRATIVE":      "Dup Trend Direction",  # reuse for narrative sentence
    "VALUE_RECOVERED_HIST_DUPES":  "Value Recovered - Hist Dupes",
    "VALUE_RECOVERED_HIST_ERRORS": "Value Recovered - Hist Errors",
    "INV_ERRORS_CONFIRMED_COUNT":  "Inv Errors Confirmed (Count)",
    "INVOICE_ERRORS_AHEAD_PAY_RUN": "Invoice Errors Ahead of Pay Run",
    "MOST_COMMON_ERROR_TYPE":   "Most Common Error Type",
    "ERROR_TYPE_COUNT":         "Error Type Count",

    # Statements
    "PCT_SPEND_RECONCILED":     "% Spend Reconciled",
    "PCT_INVOICES_RECONCILED":  "% Invoices Reconciled",
    "PCT_VENDORS_RECONCILED":   "% Vendors Reconciled",
    "PCT_STATEMENTS_RECONCILED":"% Statements Reconciled",
    "PCT_STATEMENTS_AUTOMATED": "% Statements Automated",
    "FULL_AUTOMATION_NOW":      "Full Automation % Now",
    "FULL_AUTOMATION_PREV":     "Full Automation % Previous EBR",
    "INVOICE_COVERAGE_PCT":     "Invoice Coverage %",
    "INDUSTRY_BENCHMARK":       "Industry Benchmark",
    "MISSING_CREDITS_RECOVERED":"Missing Credits Recovered (£)",
    "MISSED_CREDITS_COUNT":     "Missed Credits Count",
    "VENDORS_PCT_TOTAL_SPEND":  "Vendors % of Total Spend",
    "MOM_RECONCILIATION_GROWTH":"MoM Reconciliation Growth",

    # Helpdesk
    "NEW_TICKETS_RAISED":       "New Tickets Raised",
    "OPEN_TICKETS":             "Open Tickets",
    "CLOSED_TICKETS":           "Closed Tickets",
    "TICKETS_WAITING_XELIX":    "Tickets Waiting on Xelix",
    "TICKETS_WAITING_CUSTOMER": "Tickets Waiting on Customer",
    "TOTAL_TICKETS_YTD":        "Total Tickets (YTD)",
    "AVG_HANDLING_TIME_CURRENT":"Avg Handling Time - Current",
    "AVG_HANDLING_TIME_PREV":   "Avg Handling Time - Previous",
    "PCT_HANDLING_TIME_DECREASE":"% Handling Time Decrease",
    "PCT_TICKETS_VIA_GEN_AI":   "% Tickets via Gen AI",
    "GEN_AI_PCT_DEEP_DIVE":     "Gen AI % Deep Dive",
    "PCT_TICKETS_VIA_TRIGGERS": "% Tickets via Triggers",

    # Goals
    "GOAL_1":        "Goal 1",
    "GOAL_1_TARGET": "Goal 1 Target",
    "GOAL_2":        "Goal 2",
    "GOAL_2_TARGET": "Goal 2 Target",
    "GOAL_3":        "Goal 3",
    "GOAL_3_TARGET": "Goal 3 Target",
    "GOAL_4":        "Goal 4",
    "GOAL_4_TARGET": "Goal 4 Target",

    # Voice of customer / health
    "CUSTOMER_SENTIMENT":       "Customer Sentiment",
    "TOP_THEME_FROM_CALLS":     "Top Theme from Calls",
    "KEY_CUSTOMER_QUOTE":       "Key Customer Quote",
    "KEY_NOTES_LAST_TOUCHPOINT":"Key Notes from Last Touchpoint",
    "RISKS_OR_COMPLAINTS":      "Risks or Complaints",
    "OVERALL_RISK_RATING":      "Overall Risk Rating",
    "HEALTH_TREND":             "Health Trend",
    "CHURN_RISK_FLAG":          "Churn Risk Flag",
    "RECOMMENDED_ACTION":       "Recommended Action",

    # Actions
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
    "ACTION_6":        "Action 6",
    "ACTION_6_OWNER":  "Action 6 Owner",
    "ACTION_6_DUE":    "Action 6 Due",
    "ACTION_6_STATUS": "Action 6 Status",
}

# Trend direction → natural language sentence
TREND_SENTENCES = {
    "Increase YTD": "Year-to-date figures indicate an increase in duplicate invoices.",
    "Decrease YTD": "The number of duplicate invoices has declined year-to-date.",
    "Stable":       "Duplicate invoice volumes have remained consistent month-on-month.",
}


def build_replacements(data: dict) -> dict:
    """Build token → value mapping from Notion property data."""
    repl = {}
    for token, notion_key in FIELD_MAP.items():
        val = data.get(notion_key, "")
        if val is None:
            val = ""
        # Special: DUP_TREND_NARRATIVE uses the trend direction to generate a sentence
        if token == "DUP_TREND_NARRATIVE":
            val = TREND_SENTENCES.get(str(val).strip(), str(val))
        repl[f"{{{{{token}}}}}"] = str(val).strip() if val else "—"
    return repl


def replace_in_xml(content: str, replacements: dict) -> str:
    """
    Replace placeholder tokens in XML.
    Handles cases where tokens may be split across multiple <a:t> runs
    by first collapsing split tokens, then replacing.
    """
    # Step 1: collapse tokens that are split across XML runs
    # e.g. {{FOO}} might appear as {{FO</a:t><a:t>O}}
    # We do a pass to join adjacent text runs that together form a {{...}} token
    # Strategy: replace the raw text by extracting all text, substituting, then re-inserting

    # Simpler and more robust: just do string replacement on the full XML
    # after joining split runs where the placeholder text straddles run boundaries.

    # First try direct replacement (works when token is in a single run)
    for token, value in replacements.items():
        content = content.replace(token, value)

    # Second pass: handle split tokens by finding partial {{ or }} runs
    # Pattern: {{TOKEN}} may be split — look for {{ ... }} fragments across runs
    # We do this by temporarily collapsing adjacent a:t elements and doing replacement
    def collapse_and_replace(xml):
        # Find all a:t text runs and their positions
        pattern = r'(<a:t[^>]*>)(.*?)(</a:t>)'
        parts = re.split(pattern, xml, flags=re.DOTALL)
        # Reconstruct with collapsed checking
        result = []
        i = 0
        while i < len(parts):
            result.append(parts[i])
            i += 1
        return ''.join(result)

    return content


def process_pptx(template_path: str, replacements: dict, output_path: str):
    """Unpack pptx, replace tokens in all slide XMLs, repack."""
    import zipfile

    with tempfile.TemporaryDirectory() as tmpdir:
        # Extract
        with zipfile.ZipFile(template_path, 'r') as z:
            z.extractall(tmpdir)

        # Process all XML files
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

        # Repack
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

    # Find template relative to this script
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
