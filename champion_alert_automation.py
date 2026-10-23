"""
Champion Dark Alert Automation
Author: Valentino Cotte

Signal: Champion login frequency drops to Never/Quarterly
        AND renewal is within 90 days
        AND last CSM note is stale (>21 days)

Why this signal?
  Any one of these alone is noise. Together they represent the exact blind spot
  that kills NRR silently: CS finds out about churn when the customer doesn't renew.
  This compound trigger catches accounts that LOOK fine in aggregate but have a
  specific human risk that no health score will surface.

This script:
  1. Reads the account dataset (or in production: queries your CRM/data warehouse)
  2. Evaluates the compound trigger for each account
  3. Calls Claude API to generate a context-aware, actionable Slack alert per account
  4. Posts the alert to the assigned CSM's Slack channel via webhook
"""

import pandas as pd
import numpy as np
import json
import requests
import os
from datetime import datetime
from dotenv import load_dotenv

# Paths (relative to this script's location)
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = os.path.join(BASE_DIR, 'data', 'accounts_dataset.xlsx')
load_dotenv(os.path.join(BASE_DIR, '.env'))

# Config
TODAY = pd.Timestamp('2026-02-26')

# Credentials loaded automatically from .env file
ANTHROPIC_API_KEY = os.getenv('ANTHROPIC_API_KEY', '')
SLACK_WEBHOOK_URL = os.getenv('SLACK_WEBHOOK_URL', '')

# Trigger thresholds - tunable without code changes
RENEWAL_WINDOW_DAYS    = 90    # days to renewal that triggers watch
STALE_NOTE_DAYS        = 21    # days since last CSM note = stale
DARK_LOGIN_STATES      = ['Never', 'Quarterly']  # login frequencies = "dark"
CHAMPION_ROLES         = ['Champion']

# Load data
def load_data():
    xl = pd.ExcelFile(INPUT_FILE)
    accounts = pd.read_excel(xl, 'Accounts')
    contacts = pd.read_excel(xl, 'Contacts')
    notes    = pd.read_excel(xl, 'CSM Notes')

    accounts['renewal_date']     = pd.to_datetime(accounts['renewal_date'])
    accounts['days_to_renewal']  = (accounts['renewal_date'] - TODAY).dt.days
    contacts['last_active_date'] = pd.to_datetime(contacts['last_active_date'], errors='coerce')
    notes['note_date']           = pd.to_datetime(notes['note_date'])
    return accounts, contacts, notes

# Evaluate trigger
def find_triggered_accounts(accounts, contacts, notes):
    """
    Returns list of accounts that fire the compound trigger.
    Each dict contains everything needed to generate the alert.
    """
    triggered = []

    # Latest CSM note per account
    latest_note = notes.groupby('account_id').agg(
        last_note_date=('note_date', 'max'),
        last_note_text=('note', 'last'),
        last_note_csm=('csm', 'last')
    ).reset_index()
    latest_note['days_since_note'] = (TODAY - latest_note['last_note_date']).dt.days

    # Champions per account
    champions = contacts[contacts['role_type'].isin(CHAMPION_ROLES)].copy()
    champions['is_dark'] = champions['login_frequency'].isin(DARK_LOGIN_STATES)

    for _, acct in accounts.iterrows():
        acct_id = acct['account_id']

        # GATE 1: Renewal within window (or overdue)
        days_left = acct['days_to_renewal']
        if days_left > RENEWAL_WINDOW_DAYS:
            continue

        # GATE 2: Has at least one dark champion
        acct_champs = champions[champions['account_id'] == acct_id]
        if len(acct_champs) == 0:
            continue  # No champions tracked - different alert, different workflow

        dark_champs = acct_champs[acct_champs['is_dark']]
        if len(dark_champs) == 0:
            continue  # Champions are engaged

        # GATE 3: CSM note is stale
        note_row = latest_note[latest_note['account_id'] == acct_id]
        if len(note_row) == 0:
            days_since_note = 999
            last_note_text  = "No notes on record"
            last_note_csm   = "Unknown"
        else:
            note_data       = note_row.iloc[0]
            days_since_note = note_data['days_since_note']
            last_note_text  = note_data['last_note_text']
            last_note_csm   = note_data['last_note_csm']

        if days_since_note < STALE_NOTE_DAYS:
            continue  # CSM is actively engaged - no alert needed

        # All three gates passed: build context payload
        all_champ_names = ', '.join(acct_champs['name'].tolist()) if 'name' in acct_champs.columns else "Unknown"
        dark_champ_names = ', '.join(dark_champs['name'].tolist()) if 'name' in dark_champs.columns else "Unknown"

        # Calculate usage trend
        usage_cols = ['usage_m1_sep','usage_m2_oct','usage_m3_nov',
                      'usage_m4_dec','usage_m5_jan','usage_m6_feb_current']
        usage_vals = [acct[c] for c in usage_cols if pd.notna(acct[c])]
        if len(usage_vals) >= 2:
            slope = np.polyfit(range(len(usage_vals)), usage_vals, 1)[0]
            usage_trend = "increasing" if slope > 5 else ("declining" if slope < -5 else "flat")
        else:
            usage_trend = "unknown"

        # Detect "Never Activated" vs "Champion Dark":
        # If champion feature adoption score is 0 AND no notes ever -> onboarding failure, not churn risk
        champ_adoption_avg = dark_champs['feature_adoption_score'].mean() if 'feature_adoption_score' in dark_champs.columns else None
        never_activated = (
            days_since_note == 999 and
            champ_adoption_avg is not None and
            champ_adoption_avg == 0
        )

        triggered.append({
            'account_id':       acct_id,
            'company_name':     acct['company_name'],
            'csm_assigned':     acct['csm_assigned'],
            'mrr':              acct['mrr'],
            'arr':              acct['arr'],
            'plan_tier':        acct['plan_tier'],
            'days_to_renewal':  int(days_left),
            'renewal_date':     str(acct['renewal_date'].date()),
            'nps_score':        acct['nps_score'] if pd.notna(acct['nps_score']) else 'N/A',
            'contraction_flag': bool(acct['contraction_flag']),
            'competitor':       acct['competitor_detected'] if pd.notna(acct['competitor_detected']) else None,
            'all_champions':    all_champ_names,
            'dark_champions':   dark_champ_names,
            'dark_champ_count': len(dark_champs),
            'total_champ_count':len(acct_champs),
            'days_since_note':  int(days_since_note),
            'note_display':     'No notes on record' if days_since_note == 999 else f"{int(days_since_note)}d ago",
            'last_note_text':   str(last_note_text),
            'last_note_csm':    last_note_csm,
            'usage_trend':      usage_trend,
            'never_activated':  never_activated,
            'escalated_tickets':int(acct.get('escalated_tickets', 0)) if pd.notna(acct.get('escalated_tickets', 0)) else 0,
        })

    return triggered

# Generate alert via Claude API
def generate_alert_with_claude(account_context: dict) -> str:
    """
    Calls Claude API to generate a specific, actionable Slack alert.
    The prompt is engineered to produce concise, CSM-actionable output.
    """
    alert_type = "Never Activated" if account_context.get('never_activated') else "Champion Dark Alert"
    prompt = f"""You are a GTM intelligence system. Generate a Slack alert for a Customer Success Manager.

Account context:
- Company: {account_context['company_name']} ({account_context['plan_tier']} plan)
- Alert type: {alert_type}
- MRR: ${account_context['mrr']:,.0f} | ARR: ${account_context['arr']:,.0f}
- Renewal: {account_context['days_to_renewal']} days away ({account_context['renewal_date']})
- Champion(s): {account_context['all_champions']}
- Dark champion(s) (Never/Quarterly login): {account_context['dark_champions']}
- Last CSM note: {account_context['note_display']} — "{account_context['last_note_text']}"
- NPS: {account_context['nps_score']}
- Contraction flag: {account_context['contraction_flag']}
- Competitor detected: {account_context['competitor'] or 'None'}
- Usage trend: {account_context['usage_trend']}
- Escalated tickets open: {account_context['escalated_tickets']}

{"This is a NEVER ACTIVATED account — the champion has never used a single feature and there are no CSM notes. This is an onboarding failure, not a churn risk. The action is different: focus on activation, not retention." if account_context.get('never_activated') else "This is a CHAMPION DARK alert — the champion was previously engaged but has gone quiet near renewal."}

Write a Slack message for the assigned CSM. Format:
{"⚠️" if account_context.get('never_activated') else "🚨"} *[Company Name]* — {alert_type}
Then 3-4 bullet points maximum covering: what's happening, why it matters, one specific recommended action.
Keep it under 150 words. Be direct. Do not say "I recommend" — say what to do.
Do not use phrases like "it appears" or "it seems". State facts from the context.
End with a single bold *Suggested next step:* line."""

    headers = {
        'x-api-key': ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json'
    }
    payload = {
        'model': 'claude-sonnet-4-20250514',
        'max_tokens': 400,
        'messages': [{'role': 'user', 'content': prompt}]
    }

    try:
        resp = requests.post(
            'https://api.anthropic.com/v1/messages',
            headers=headers,
            json=payload,
            timeout=15
        )
        if resp.status_code == 200:
            return resp.json()['content'][0]['text']
        else:
            # Fallback: generate alert locally without API
            return generate_fallback_alert(account_context)
    except Exception as e:
        return generate_fallback_alert(account_context)

def generate_fallback_alert(ctx: dict) -> str:
    """Deterministic fallback alert if Claude API is unavailable."""
    renewal_status = "OVERDUE" if ctx['days_to_renewal'] < 0 else f"{ctx['days_to_renewal']}d away"
    competitor_line = f"\n• ⚠️ Competitor *{ctx['competitor']}* has been detected in this account." if ctx['competitor'] else ""
    contraction_line = "\n• 🔻 Contraction flag is active — budget pressure may already be in play." if ctx['contraction_flag'] else ""
    note_display = ctx.get('note_display', f"{ctx['days_since_note']}d ago")

    if ctx.get('never_activated'):
        return f"""⚠️ *{ctx['company_name']}* — Never Activated

• *Champion(s) never engaged:* {ctx['dark_champions']} — zero feature adoption, no login activity
• *Renewal:* {renewal_status} | MRR at risk: ${ctx['mrr']:,.0f}
• *CSM notes:* {note_display} — this account has never been properly onboarded
• *Usage trend:* {ctx['usage_trend'].capitalize()}

*Suggested next step:* Do not treat this as a renewal conversation. Schedule an onboarding call this week — start from scratch on activation before renewal comes up."""

    return f"""🚨 *{ctx['company_name']}* — Champion Dark Alert

• *Champion(s) gone dark:* {ctx['dark_champions']} ({ctx['dark_champ_count']} of {ctx['total_champ_count']} champions showing Never/Quarterly login)
• *Renewal:* {renewal_status} | MRR at risk: ${ctx['mrr']:,.0f}{competitor_line}{contraction_line}
• *Last CSM note:* {note_display} — "{ctx['last_note_text'][:80]}..."
• *Usage trend:* {ctx['usage_trend'].capitalize()}

*Suggested next step:* Do not email the champion. Call their manager or the exec sponsor today. The champion may have left, been reassigned, or is actively avoiding renewal conversations."""

# Post to Slack
def post_to_slack(message: str, webhook_url: str, account_context: dict) -> bool:
    """Posts formatted alert to Slack via incoming webhook."""
    renewal_color = "#FF0000" if account_context['days_to_renewal'] < 0 else \
                    "#FF6B35" if account_context['days_to_renewal'] < 30 else "#FFC107"

    slack_payload = {
        "attachments": [
            {
                "color": renewal_color,
                "text": message,
                "footer": f"GTM Intelligence System • {TODAY.strftime('%Y-%m-%d')} • {account_context['account_id']}",
                "footer_icon": "https://platform.slack-edge.com/img/default_application_icon.png"
            }
        ]
    }

    try:
        resp = requests.post(webhook_url, json=slack_payload, timeout=10)
        return resp.status_code == 200
    except Exception as e:
        print(f"    Slack post failed: {e}")
        return False

# --- Main execution ---
def run_automation(dry_run=True):
    """
    dry_run=True: print alerts to console (default for demo)
    dry_run=False: actually post to Slack
    """
    print("=" * 60)
    print("  CHAMPION DARK ALERT — Running detection")
    print(f"  Date: {TODAY.date()}")
    print(f"  Thresholds: renewal <{RENEWAL_WINDOW_DAYS}d, note stale >{STALE_NOTE_DAYS}d")
    print("=" * 60)

    accounts, contacts, notes = load_data()
    triggered = find_triggered_accounts(accounts, contacts, notes)

    print(f"\nDetection complete. {len(triggered)} account(s) triggered.\n")

    results = []

    for ctx in triggered:
        print(f"\n{'─'*60}")
        print(f"Processing: {ctx['company_name']} (${ctx['mrr']:,.0f} MRR)")
        print(f"  CSM: {ctx['csm_assigned']} | Renewal: {ctx['days_to_renewal']}d")
        print(f"  Alert type: {'Never Activated' if ctx.get('never_activated') else 'Champion Dark'}")
        print(f"  Dark champion(s): {ctx['dark_champions']}")
        print(f"  Last note: {ctx.get('note_display', str(ctx['days_since_note']) + 'd ago')}")

        # Generate alert
        print("  Generating alert via Claude API...")
        alert_text = generate_alert_with_claude(ctx)

        print("\n  ALERT PREVIEW:")
        print("  " + "\n  ".join(alert_text.split("\n")))

        if not dry_run and SLACK_WEBHOOK_URL != 'YOUR_SLACK_WEBHOOK_HERE':
            success = post_to_slack(alert_text, SLACK_WEBHOOK_URL, ctx)
            status = "Sent to Slack" if success else "Slack post failed"
        else:
            status = "Dry run — not sent"

        print(f"\n  Status: {status}")

        results.append({
            'account_id':    ctx['account_id'],
            'company_name':  ctx['company_name'],
            'csm_assigned':  ctx['csm_assigned'],
            'mrr':           ctx['mrr'],
            'days_to_renewal': ctx['days_to_renewal'],
            'alert_generated': alert_text,
            'sent_to_slack': status
        })

    # Save run log
    log_df = pd.DataFrame(results)
    log_path = os.path.join(BASE_DIR, 'outputs', 'champion_alert_run_log.csv')
    log_df.to_csv(log_path, index=False)
    print(f"\n\nRun log saved to: {log_path}")
    print(f"   {len(triggered)} alerts generated | Run time: {TODAY.date()}")

    return results

if __name__ == '__main__':
    results = run_automation(dry_run=False)
