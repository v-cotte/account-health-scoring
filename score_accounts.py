"""
Account Intelligence Scoring Engine
Author: Valentino Cotte

Methodology:
  Three signal tiers -> composite 0-100 score -> letter grade -> segment -> next action
  Contradictory signals are flagged explicitly rather than averaged away.

Usage:
  1. Run: python score_accounts.py
  2. Output will be saved to outputs/account_intelligence_list.xlsx
"""

import os
import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')
from dotenv import load_dotenv

# Paths (relative to this script's location)
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE  = os.path.join(BASE_DIR, 'data', 'accounts_dataset.xlsx')
OUTPUT_FILE = os.path.join(BASE_DIR, 'outputs', 'account_intelligence_list.xlsx')
load_dotenv(os.path.join(BASE_DIR, '.env'))

TODAY = pd.Timestamp('2026-02-26')

# Load data
xl = pd.ExcelFile(INPUT_FILE)
accounts  = pd.read_excel(xl, 'Accounts')
contacts  = pd.read_excel(xl, 'Contacts')
tickets   = pd.read_excel(xl, 'Support Tickets')
notes     = pd.read_excel(xl, 'CSM Notes')
expansion = pd.read_excel(xl, 'Expansion History')

# Date parsing
accounts['renewal_date']    = pd.to_datetime(accounts['renewal_date'])
accounts['contract_start']  = pd.to_datetime(accounts['contract_start'])
accounts['days_to_renewal'] = (accounts['renewal_date'] - TODAY).dt.days
notes['note_date']          = pd.to_datetime(notes['note_date'])
contacts['last_active_date']= pd.to_datetime(contacts['last_active_date'], errors='coerce')

# FEATURE ENGINEERING

# --- Usage trend (linear slope over 6 months) ---
usage_cols = ['usage_m1_sep','usage_m2_oct','usage_m3_nov',
              'usage_m4_dec','usage_m5_jan','usage_m6_feb_current']
accounts['usage_trend_slope'] = accounts[usage_cols].apply(
    lambda row: np.polyfit(range(6), row.fillna(row.mean()), 1)[0], axis=1
)
accounts['usage_last3_avg'] = accounts[['usage_m4_dec','usage_m5_jan','usage_m6_feb_current']].mean(axis=1)
accounts['usage_first3_avg']= accounts[['usage_m1_sep','usage_m2_oct','usage_m3_nov']].mean(axis=1)
accounts['usage_pct_change'] = (
    (accounts['usage_last3_avg'] - accounts['usage_first3_avg'])
    / accounts['usage_first3_avg'].replace(0, np.nan)
).fillna(0)

# --- Seat utilization ---
accounts['seat_utilization'] = (
    accounts['seats_active'] / accounts['seats_purchased'].replace(0, np.nan)
).fillna(0)

# --- Feature adoption breadth ---
feature_cols = [c for c in accounts.columns if c.startswith('feature_')]
accounts['feature_adoption_count'] = (accounts[feature_cols] > 0).sum(axis=1)
accounts['feature_adoption_pct']   = accounts['feature_adoption_count'] / len(feature_cols)

# --- Champion signals (per account) ---
champs = contacts[contacts['role_type'] == 'Champion'].copy()
login_score_map = {'Daily': 5, 'Weekly': 4, 'Monthly': 3, 'Quarterly': 2, 'Never': 0}
champs['login_score'] = champs['login_frequency'].map(login_score_map).fillna(2)
champ_agg = champs.groupby('account_id').agg(
    champion_count=('contact_id','count'),
    champion_login_score_avg=('login_score','mean'),
    champion_login_min=('login_score','min'),
    champion_last_active=('last_active_date','max')
).reset_index()
champ_agg['champion_days_since_active'] = (TODAY - champ_agg['champion_last_active']).dt.days
accounts = accounts.merge(champ_agg, on='account_id', how='left')
accounts['champion_dark'] = accounts['champion_login_min'].fillna(3) <= 2

# --- Exec sponsor count ---
exec_contacts = contacts[contacts['role_type'] == 'Executive Sponsor']
exec_agg = exec_contacts.groupby('account_id').agg(
    exec_count=('contact_id','count')
).reset_index()
accounts = accounts.merge(exec_agg, on='account_id', how='left')
accounts['exec_count'] = accounts['exec_count'].fillna(0)

# --- CSM note recency & quality ---
latest_note = notes.groupby('account_id').agg(
    last_note_date=('note_date','max'),
    note_count=('note_date','count')
).reset_index()
latest_note['days_since_note'] = (TODAY - latest_note['last_note_date']).dt.days
notes['note_len'] = notes['note'].str.len()
note_quality = notes.groupby('account_id')['note_len'].mean().reset_index()
note_quality.columns = ['account_id','avg_note_len']
accounts = accounts.merge(latest_note, on='account_id', how='left')
accounts = accounts.merge(note_quality, on='account_id', how='left')
accounts['days_since_note'] = accounts['days_since_note'].fillna(999)
accounts['note_count']      = accounts['note_count'].fillna(0)
accounts['avg_note_len']    = accounts['avg_note_len'].fillna(0)

# --- Support burden ---
ticket_agg = tickets.groupby('account_id').agg(
    ticket_count_90d=('ticket_id','count'),
    escalated_tickets=('status', lambda x: (x=='Escalated').sum()),
    open_tickets=('status', lambda x: (x.isin(['Open','In Progress'])).sum()),
    avg_severity=('severity','mean')
).reset_index()
accounts = accounts.merge(ticket_agg, on='account_id', how='left')
accounts['ticket_count_90d']   = accounts['ticket_count_90d'].fillna(0)
accounts['escalated_tickets']  = accounts['escalated_tickets'].fillna(0)
accounts['open_tickets']       = accounts['open_tickets'].fillna(0)
accounts['avg_severity']       = accounts['avg_severity'].fillna(0)

# --- Expansion signals ---
recent_expansion = expansion[expansion['event_type'].isin(['Expansion','Upsell'])]
exp_agg = recent_expansion.groupby('account_id').agg(
    recent_expansion_count=('event_type','count'),
    total_expansion_mrr=('mrr_change','sum')
).reset_index()
accounts = accounts.merge(exp_agg, on='account_id', how='left')
accounts['recent_expansion_count'] = accounts['recent_expansion_count'].fillna(0)
accounts['total_expansion_mrr']    = accounts['total_expansion_mrr'].fillna(0)

# SCORING - THREE TIERS -> COMPOSITE

def normalize(series, low_is_bad=True):
    mn, mx = series.min(), series.max()
    if mx == mn:
        return pd.Series([50.0]*len(series), index=series.index)
    norm = (series - mn) / (mx - mn) * 100
    return norm if low_is_bad else 100 - norm

df = accounts.copy()

# TIER 1: Engagement Health (40%)
t1_usage_trend  = normalize(df['usage_trend_slope'], low_is_bad=True)
t1_usage_change = normalize(df['usage_pct_change'], low_is_bad=True)
df['seat_util_capped'] = df['seat_utilization'].clip(0, 1.2)
t1_seat_util    = normalize(df['seat_util_capped'], low_is_bad=True)
t1_feature      = normalize(df['feature_adoption_pct'], low_is_bad=True)
t1_champion     = normalize(df['champion_login_score_avg'].fillna(2), low_is_bad=True)

tier1 = (
    t1_usage_trend  * 0.25 +
    t1_usage_change * 0.20 +
    t1_seat_util    * 0.20 +
    t1_feature      * 0.15 +
    t1_champion     * 0.20
)

# TIER 2: Relationship Health (30%)
nps_filled = df['nps_score'].fillna(df['nps_score'].median())
t2_nps          = normalize(nps_filled, low_is_bad=True)
t2_note_recency = normalize(df['days_since_note'], low_is_bad=False)
t2_note_quality = normalize(df['avg_note_len'], low_is_bad=True)
t2_exec         = normalize(df['exec_count'], low_is_bad=True)
t2_exec_changes = normalize(df['exec_sponsor_changes'].fillna(0), low_is_bad=False)

tier2 = (
    t2_nps          * 0.35 +
    t2_note_recency * 0.25 +
    t2_note_quality * 0.15 +
    t2_exec         * 0.15 +
    t2_exec_changes * 0.10
)

# TIER 3: Risk Signals (30%)
t3_contraction = 100 - df['contraction_flag'].astype(int) * 100
t3_competitor  = 100 - df['competitor_detected'].notna().astype(int) * 100
t3_escalated   = normalize(df['escalated_tickets'], low_is_bad=False)
t3_severity    = normalize(df['avg_severity'].fillna(0), low_is_bad=False)
t3_discount    = normalize(df['discount_pct'].fillna(0), low_is_bad=False)

tier3 = (
    t3_contraction * 0.35 +
    t3_competitor  * 0.25 +
    t3_escalated   * 0.20 +
    t3_severity    * 0.10 +
    t3_discount    * 0.10
)

# COMPOSITE SCORE
df['score_engagement']   = tier1.round(1)
df['score_relationship'] = tier2.round(1)
df['score_risk']         = tier3.round(1)
df['health_score']       = (tier1*0.40 + tier2*0.30 + tier3*0.30).round(1)

def letter_grade(score):
    if score >= 80: return 'A'
    if score >= 65: return 'B'
    if score >= 50: return 'C'
    if score >= 35: return 'D'
    return 'F'

df['grade'] = df['health_score'].apply(letter_grade)

# FLAGS
df['is_overdue']     = df['days_to_renewal'] < 0
df['renewal_urgent'] = (df['days_to_renewal'] >= 0) & (df['days_to_renewal'] < 30)
df['renewal_soon']   = (df['days_to_renewal'] >= 30) & (df['days_to_renewal'] < 90)
df['champion_dark']  = df['champion_login_min'].fillna(3) <= 2
df['stale_notes']    = df['days_since_note'] > 21
df['expansion_ready'] = (
    (df['seat_utilization'] > 0.75) &
    (df['pricing_page_visits_30d'] > 0) &
    (df['usage_trend_slope'] > 0)
)

# CONTRADICTION DETECTION
df['tier_divergence']  = (df['score_engagement'] - df['score_relationship']).abs()
df['is_contradictory'] = df['tier_divergence'] > 25

# SEGMENTATION

def assign_segment(row):
    if row['is_overdue']:
        return 'OVERDUE - Escalate Today'
    if row['health_score'] < 50 and (row['renewal_urgent'] or row['renewal_soon']):
        return 'Save Now'
    if row['champion_dark'] and (row['renewal_urgent'] or row['renewal_soon']) and row['health_score'] >= 50:
        return 'Champion Risk'
    if row['expansion_ready'] and row['health_score'] >= 60:
        return 'Expand Now'
    if row['health_score'] < 50 and not row['renewal_soon'] and not row['renewal_urgent']:
        return 'Defend'
    if row['health_score'] >= 60 and (row['renewal_urgent'] or row['renewal_soon']):
        return 'Renew & Grow'
    if row['is_contradictory']:
        return 'Ambiguous - Review'
    return 'Healthy - Monitor'

df['segment'] = df.apply(assign_segment, axis=1)

# NEXT ACTION

def next_action(row):
    seg  = row['segment']
    days = int(row['days_to_renewal']) if not pd.isna(row['days_to_renewal']) else 999

    if seg == 'OVERDUE - Escalate Today':
        return f"URGENT: Contract expired {abs(days)}d ago. Escalate to VP CS + loop in AE today. Do not send auto-renewal email."
    if seg == 'Save Now':
        if row['contraction_flag']:
            return f"Renewal in {days}d + contraction flag active. Schedule exec-to-exec call this week - skip CSM-only outreach."
        return f"Renewal in {days}d + health score {row['health_score']:.0f}. Run emergency health review; prepare concession options."
    if seg == 'Champion Risk':
        champ_days = int(row['champion_days_since_active']) if not pd.isna(row.get('champion_days_since_active')) else 99
        return f"Champion inactive {champ_days}d. Renewal in {days}d. Identify new internal champion before renewal conversation."
    if seg == 'Expand Now':
        util = row['seat_utilization']
        if util > 1.1:
            return f"Over seat capacity ({util:.0%} utilization). Send expansion proposal this week - natural forcing function."
        return f"Pricing page visits + positive usage trend. Hand off to AE for expansion conversation within 7 days."
    if seg == 'Defend':
        if row['competitor_detected'] and pd.notna(row['competitor_detected']):
            return f"Competitor ({row['competitor_detected']}) detected. Schedule product roadmap call + send competitive battle card."
        if row['escalated_tickets'] > 0:
            return f"{int(row['escalated_tickets'])} escalated ticket(s) open. Resolve before any commercial conversation."
        return "Schedule proactive health check call. Focus on driving feature adoption to deepen stickiness."
    if seg == 'Renew & Grow':
        return f"Renewal in {days}d. Schedule QBR now - prepare ROI summary + 1 expansion ask based on usage growth."
    if seg == 'Ambiguous - Review':
        if row['nps_score'] >= 8 and row['contraction_flag']:
            return "NPS high but contraction active - likely budget/procurement issue, not product. Escalate to finance champion."
        if row['nps_score'] <= 4 and row['expansion_events_count'] > 0:
            return "Low NPS but still expanding - dependent but dissatisfied. Prioritize support resolution before renewal."
        return "Conflicting signals across engagement and relationship data. CSM to review manually before next touchpoint."
    if row['stale_notes'] and row['days_since_note'] > 60:
        return "No CSM note in 60+ days. Log a touchpoint this week to maintain coverage data integrity."
    return "Healthy. Quarterly check-in sufficient. Watch for seat utilization growth as expansion trigger."

df['next_action'] = df.apply(next_action, axis=1)

# CONTRADICTION FLAGS
def contradiction_flag(row):
    flags = []
    if row['nps_score'] >= 8 and row['contraction_flag']:
        flags.append(f"NPS={row['nps_score']:.0f} but contraction active")
    if pd.notna(row['nps_score']) and row['nps_score'] <= 4 and row['expansion_events_count'] > 0:
        flags.append(f"NPS={row['nps_score']:.0f} but {int(row['expansion_events_count'])} expansion events")
    if row['usage_trend_slope'] < -30 and (row['nps_score'] >= 8 if pd.notna(row['nps_score']) else False):
        flags.append("Usage declining sharply despite high NPS")
    if row['is_overdue']:
        flags.append("Renewal date passed - status unknown")
    if row['days_since_note'] > 90:
        flags.append(f"No CSM note in {int(row['days_since_note'])}d")
    if row['champion_dark'] and row['health_score'] > 65:
        flags.append("Champion inactive despite healthy score")
    return "; ".join(flags) if flags else ""

df['judgment_flags']     = df.apply(contradiction_flag, axis=1)
df['needs_human_review'] = df['judgment_flags'] != ""

# PRIORITY RANK within each CSM's book
segment_priority = {
    'OVERDUE - Escalate Today': 1,
    'Save Now': 2,
    'Champion Risk': 3,
    'Renew & Grow': 4,
    'Expand Now': 5,
    'Defend': 6,
    'Ambiguous - Review': 7,
    'Healthy - Monitor': 8,
}
df['segment_rank'] = df['segment'].map(segment_priority).fillna(9)
df['csm_priority_rank'] = (
    df.groupby('csm_assigned')
    .apply(lambda g: g.sort_values(['segment_rank', 'health_score']).assign(
        csm_priority_rank=range(1, len(g)+1)
    ))
    ['csm_priority_rank']
    .reset_index(level=0, drop=True)
)

# BUILD OUTPUT SPREADSHEET

output_cols = [
    'csm_priority_rank',
    'account_id', 'company_name', 'csm_assigned',
    'plan_tier', 'mrr', 'arr',
    'health_score', 'grade',
    'score_engagement', 'score_relationship', 'score_risk',
    'segment', 'next_action',
    'days_to_renewal', 'renewal_date',
    'seat_utilization', 'usage_trend_slope',
    'feature_adoption_pct',
    'champion_login_score_avg', 'champion_dark',
    'days_since_note', 'nps_score',
    'contraction_flag', 'competitor_detected',
    'escalated_tickets', 'expansion_events_count',
    'pricing_page_visits_30d',
    'needs_human_review', 'judgment_flags',
    'is_overdue', 'expansion_ready',
]

output = df[output_cols].copy()
output = output.sort_values(['csm_assigned', 'csm_priority_rank']).reset_index(drop=True)
output['seat_utilization']     = (output['seat_utilization'] * 100).round(1).astype(str) + '%'
output['feature_adoption_pct'] = (output['feature_adoption_pct'] * 100).round(1).astype(str) + '%'

writer = pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter')
wb = writer.book

fmt_header   = wb.add_format({'bold': True, 'bg_color': '#1a1a2e', 'font_color': 'white', 'border': 1, 'font_size': 10})
fmt_overdue  = wb.add_format({'bg_color': '#FF0000', 'font_color': 'white', 'bold': True})
fmt_save_now = wb.add_format({'bg_color': '#FF6B6B', 'font_color': 'white'})
fmt_champ_r  = wb.add_format({'bg_color': '#FFB347', 'font_color': '#1a1a1a'})
fmt_expand   = wb.add_format({'bg_color': '#90EE90', 'font_color': '#1a1a1a'})
fmt_renew    = wb.add_format({'bg_color': '#87CEEB', 'font_color': '#1a1a1a'})
fmt_defend   = wb.add_format({'bg_color': '#FFD700', 'font_color': '#1a1a1a'})
fmt_ambig    = wb.add_format({'bg_color': '#D3D3D3', 'font_color': '#1a1a1a'})
fmt_healthy  = wb.add_format({'bg_color': '#FFFFFF', 'font_color': '#1a1a1a'})
fmt_default  = wb.add_format({'font_size': 10, 'border': 1})
fmt_wrap     = wb.add_format({'text_wrap': True, 'font_size': 9, 'valign': 'top', 'border': 1})
fmt_score_a  = wb.add_format({'bg_color': '#28a745', 'font_color': 'white', 'bold': True, 'align': 'center'})
fmt_score_b  = wb.add_format({'bg_color': '#5cb85c', 'font_color': 'white', 'align': 'center'})
fmt_score_c  = wb.add_format({'bg_color': '#f0ad4e', 'font_color': 'white', 'align': 'center'})
fmt_score_d  = wb.add_format({'bg_color': '#d9534f', 'font_color': 'white', 'align': 'center'})
fmt_score_f  = wb.add_format({'bg_color': '#8B0000', 'font_color': 'white', 'bold': True, 'align': 'center'})

sheet_name = 'Account Intelligence List'
output.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
ws = writer.sheets[sheet_name]

for col_num, col_name in enumerate(output.columns):
    ws.write(0, col_num, col_name, fmt_header)

col_widths = {
    'csm_priority_rank': 6, 'account_id': 10, 'company_name': 22,
    'csm_assigned': 14, 'plan_tier': 9, 'mrr': 9, 'arr': 10,
    'health_score': 9, 'grade': 6,
    'score_engagement': 10, 'score_relationship': 12, 'score_risk': 8,
    'segment': 24, 'next_action': 50,
    'days_to_renewal': 10, 'renewal_date': 12,
    'seat_utilization': 10, 'usage_trend_slope': 10,
    'feature_adoption_pct': 12, 'champion_login_score_avg': 14,
    'champion_dark': 10, 'days_since_note': 11, 'nps_score': 8,
    'contraction_flag': 10, 'competitor_detected': 16,
    'escalated_tickets': 10, 'expansion_events_count': 12,
    'pricing_page_visits_30d': 13,
    'needs_human_review': 12, 'judgment_flags': 45,
    'is_overdue': 8, 'expansion_ready': 10,
}
for i, col in enumerate(output.columns):
    ws.set_column(i, i, col_widths.get(col, 12))

segment_fmt_map = {
    'OVERDUE - Escalate Today': fmt_overdue,
    'Save Now': fmt_save_now,
    'Champion Risk': fmt_champ_r,
    'Expand Now': fmt_expand,
    'Renew & Grow': fmt_renew,
    'Defend': fmt_defend,
    'Ambiguous - Review': fmt_ambig,
    'Healthy - Monitor': fmt_healthy,
}

for row_idx, row in output.iterrows():
    excel_row = row_idx + 2
    seg = row['segment']
    row_fmt = segment_fmt_map.get(seg, fmt_default)
    for col_idx, val in enumerate(row):
        col_name = output.columns[col_idx]
        if col_name in ['next_action', 'judgment_flags']:
            ws.write(excel_row, col_idx, str(val) if pd.notna(val) else '', fmt_wrap)
        elif col_name == 'grade':
            grade_fmt = {'A': fmt_score_a, 'B': fmt_score_b, 'C': fmt_score_c,
                         'D': fmt_score_d, 'F': fmt_score_f}.get(str(val), fmt_default)
            ws.write(excel_row, col_idx, str(val) if pd.notna(val) else '', grade_fmt)
        else:
            display_val = val
            if isinstance(val, float) and pd.isna(val): display_val = ''
            elif isinstance(val, bool): display_val = str(val)
            ws.write(excel_row, col_idx, display_val, row_fmt)

ws.freeze_panes(2, 3)

# Segment Summary
seg_summary = df.groupby('segment').agg(
    account_count=('account_id','count'),
    total_arr=('arr','sum'),
    avg_health_score=('health_score','mean'),
    avg_mrr=('mrr','mean'),
    avg_days_to_renewal=('days_to_renewal','mean'),
    accounts_needing_review=('needs_human_review','sum')
).reset_index().sort_values('account_count', ascending=False)
seg_summary['total_arr'] = seg_summary['total_arr'].round(0)
seg_summary['avg_health_score'] = seg_summary['avg_health_score'].round(1)
seg_summary['avg_mrr'] = seg_summary['avg_mrr'].round(0)
seg_summary['avg_days_to_renewal'] = seg_summary['avg_days_to_renewal'].round(0)
seg_summary.to_excel(writer, sheet_name='Segment Summary', index=False)

# CSM Workload View
csm_summary = df.groupby('csm_assigned').agg(
    accounts=('account_id','count'),
    total_arr=('arr','sum'),
    avg_health=('health_score','mean'),
    overdue=('is_overdue','sum'),
    save_now=('segment', lambda x: (x=='Save Now').sum()),
    champion_risk=('segment', lambda x: (x=='Champion Risk').sum()),
    expand_now=('segment', lambda x: (x=='Expand Now').sum()),
    pct_stale_notes=('stale_notes', lambda x: f"{x.mean()*100:.0f}%"),
    accounts_review=('needs_human_review','sum'),
    total_escalated_tickets=('escalated_tickets','sum'),
).reset_index()
csm_summary['avg_health'] = csm_summary['avg_health'].round(1)
csm_summary['total_arr'] = csm_summary['total_arr'].round(0)
csm_summary.to_excel(writer, sheet_name='CSM Workload View', index=False)

# Human Judgment Queue
review_queue = df[df['needs_human_review']][
    ['account_id','company_name','csm_assigned','mrr','health_score','grade','segment',
     'days_to_renewal','nps_score','contraction_flag','judgment_flags','next_action']
].sort_values('mrr', ascending=False)
review_queue.to_excel(writer, sheet_name='Human Judgment Queue', index=False)

# Expansion Pipeline
expand_df = df[df['expansion_ready']].sort_values('mrr', ascending=False)
expand_pipeline = expand_df[[
    'account_id','company_name','csm_assigned','mrr','arr','plan_tier',
    'seat_utilization','pricing_page_visits_30d','usage_trend_slope',
    'health_score','grade','days_to_renewal','recent_expansion_count',
    'feature_adoption_pct','next_action'
]].copy()
expand_pipeline['seat_utilization'] = (expand_df['seat_utilization'] * 100).round(1).astype(str) + '%'
expand_pipeline.to_excel(writer, sheet_name='Expansion Pipeline', index=False)

writer.close()

# Add Methodology tab
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb2 = load_workbook(OUTPUT_FILE)
if 'Methodology' in wb2.sheetnames:
    del wb2['Methodology']
ws2 = wb2.create_sheet('Methodology', 0)

dark_fill  = PatternFill("solid", fgColor="1a1a2e")
light_fill = PatternFill("solid", fgColor="EBF4FF")
grey_fill  = PatternFill("solid", fgColor="F7F7F7")
title_font = Font(name="Arial", bold=True, size=16, color="1a1a2e")
h2_font    = Font(name="Arial", bold=True, size=12, color="2C5282")
body_font  = Font(name="Arial", size=10, color="333333")
bold_font  = Font(name="Arial", bold=True, size=10, color="333333")
small_font = Font(name="Arial", size=9, color="666666")
white_font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
thin       = Side(style="thin", color="CCCCCC")
box        = Border(top=thin, bottom=thin, left=thin, right=thin)
wrap       = Alignment(wrap_text=True, vertical="top")
center     = Alignment(horizontal="center", vertical="center")

col_widths = [2, 22, 12, 12, 12, 45, 2, 22, 35, 2]
for i, w in enumerate(col_widths, 1):
    ws2.column_dimensions[get_column_letter(i)].width = w

def section_header(ws, row, col, text, width=8):
    c = ws.cell(row=row, column=col, value=text)
    c.font = h2_font
    c.fill = light_fill
    c.alignment = Alignment(vertical="center", indent=1)
    c.border = Border(left=Side(style="thick", color="2C5282"))
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col+width-1)
    ws.row_dimensions[row].height = 22

ws2.row_dimensions[2].height = 36
ws2.row_dimensions[3].height = 20
ws2.merge_cells("B2:I2")
c = ws2["B2"]
c.value = "Account Health Score - Methodology"
c.font = title_font
c.alignment = Alignment(vertical="center")
ws2.merge_cells("B3:I3")
c = ws2["B3"]
c.value = ("Every account is scored 0–100 and assigned a letter grade (A–F). "
           "The score is built from three signal tiers. When tiers contradict each other, "
           "the account is flagged for human review rather than averaged.")
c.font = small_font
c.alignment = wrap

row = 6
section_header(ws2, row, 2, "Score Structure - Three Tiers", width=8)
row += 1
for i, h in enumerate(["Tier", "Name", "Weight", "Signals Included", "What It Measures"]):
    c = ws2.cell(row=row, column=i+2, value=h)
    c.font = white_font; c.fill = dark_fill; c.alignment = center; c.border = box
ws2.row_dimensions[row].height = 18

tiers = [
    ("Tier 1","Engagement Health","40%",
     "Usage trend slope (6mo), usage % change, seat utilization, feature adoption breadth, champion login frequency",
     "Is the customer actively using the product?"),
    ("Tier 2","Relationship Health","30%",
     "NPS score, CSM note recency, CSM note quality, exec sponsor presence, exec sponsor change count",
     "Is the human relationship solid?"),
    ("Tier 3","Risk Signals","30%",
     "Contraction flag, competitor detected, escalated tickets, avg ticket severity, discount %",
     "Are there active danger signs?"),
]
tier_colors = ["DBEAFE","D1FAE5","FEE2E2"]
for i,(tier,name,weight,signals,measures) in enumerate(tiers):
    row += 1
    ws2.row_dimensions[row].height = 42
    fill = PatternFill("solid", fgColor=tier_colors[i])
    for j, val in enumerate([tier,name,weight,signals,measures]):
        c = ws2.cell(row=row, column=j+2, value=val)
        c.fill = fill
        c.font = bold_font if j < 3 else body_font
        c.alignment = Alignment(horizontal="center",vertical="top",wrap_text=True) if j<3 else wrap
        c.border = box

row += 2
section_header(ws2, row, 2, "Segments - What Each One Means", width=8)
row += 1
for i, h in enumerate(["Segment","Trigger Logic","Who Acts","What Kind of Action"]):
    c = ws2.cell(row=row, column=i+2, value=h)
    c.font = white_font; c.fill = dark_fill; c.alignment = center; c.border = box
ws2.row_dimensions[row].height = 18

segments = [
    ("🔴 OVERDUE - Escalate Today","Renewal date passed","VP CS + AE","Immediate escalation. Do not send auto-renewal email."),
    ("🔴 Save Now","Health <50 AND renewal <90d","CSM + CS Manager","Emergency review. Exec-to-exec call if contraction active."),
    ("🟠 Champion Risk","Champion dark AND renewal <90d AND health ≥50","CSM","Find new champion before any renewal conversation."),
    ("🟢 Expand Now","Seat util >75% AND pricing page visit AND usage growing","AE (CS hands off)","Expansion proposal this week - they're already over capacity."),
    ("🔵 Renew & Grow","Health ≥60 AND renewal <90d","CSM","QBR now. ROI summary + one expansion ask."),
    ("🟡 Defend","Health <50 AND renewal >90d","CSM","Fix root problem first. No commercial conversation yet."),
    ("⚪ Ambiguous - Review","Tier 1 vs Tier 2 diverge >25pts OR NPS contradicts behavior","CSM (manual)","See judgment_flags column for specific contradiction."),
    ("⬜ Healthy - Monitor","Health ≥60, no urgency","CSM","Quarterly check-in. Watch seat utilization."),
]
seg_colors = ["FFC7CE","FFC7CE","FFEB9C","C6EFCE","DBEAFE","FFD700","D3D3D3","FFFFFF"]
for i,(seg,trigger,who,action) in enumerate(segments):
    row += 1
    ws2.row_dimensions[row].height = 34
    fill = PatternFill("solid", fgColor=seg_colors[i])
    for j, val in enumerate([seg,trigger,who,action]):
        c = ws2.cell(row=row, column=j+2, value=val)
        c.fill = fill; c.font = bold_font if j==0 else body_font
        c.alignment = wrap; c.border = box

row += 2
section_header(ws2, row, 2, "Key Design Decisions", width=8)
decisions = [
    ("Why three tiers instead of a flat weighted average?",
     "Engagement and relationship health can contradict each other. A flat average hides the contradiction. Three tiers make it visible - when they diverge >25 points the account is flagged for human review instead of silently averaged."),
    ("How are contradictory accounts handled?",
     "30 accounts have NPS scores that contradict behavioral signals. High NPS + contraction = budget issue, not relationship problem. Low NPS + expanding = dependent but dissatisfied. Both land in 'Ambiguous - Review' with the specific contradiction in the judgment_flags column."),
    ("Why is there an urgency override on top of the score?",
     "A score of 65 does not mean safe if renewal was yesterday. The urgency layer overrides segment assignment for overdue accounts, renewal <30d, and champion dark scenarios."),
    ("How are accounts with no CSM notes scored?",
     "34 accounts have zero notes. Missing note ≠ healthy. They get a neutral relationship sub-score and a judgment flag. Next action prompts the CSM to log a touchpoint - creating the data that makes diagnosis possible."),
]
for q, a in decisions:
    row += 1
    ws2.row_dimensions[row].height = 15
    c = ws2.cell(row=row, column=2, value=q)
    c.font = bold_font; c.fill = light_fill
    c.alignment = Alignment(vertical="center", indent=1)
    c.border = Border(left=Side(style="thick", color="2C5282"))
    ws2.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)
    row += 1
    lines = len(a)//80+1
    ws2.row_dimensions[row].height = max(30, lines*14)
    c = ws2.cell(row=row, column=2, value=a)
    c.font = body_font; c.fill = grey_fill
    c.alignment = Alignment(wrap_text=True, vertical="top", indent=2)
    c.border = Border(left=Side(style="thick", color="2C5282"), bottom=thin)
    ws2.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)

row += 2
c = ws2.cell(row=row, column=2,
    value="Scores computed by score_accounts.py")
c.font = Font(name="Arial", size=8, color="999999", italic=True)
ws2.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)

wb2.save(OUTPUT_FILE)

# FORMAT ALL REMAINING TABS
wb3 = load_workbook(OUTPUT_FILE)

# Shared styles
dark_fill2   = PatternFill("solid", fgColor="1a1a2e")
light_fill2  = PatternFill("solid", fgColor="EBF4FF")
grey_fill2   = PatternFill("solid", fgColor="F7F7F7")
alt_fill2    = PatternFill("solid", fgColor="FAFAFA")
white_fill2  = PatternFill("solid", fgColor="FFFFFF")
red_fill2    = PatternFill("solid", fgColor="FFC7CE")
orange_fill2 = PatternFill("solid", fgColor="FFEB9C")

header_font2 = Font(name="Arial", bold=True, size=10, color="FFFFFF")
title_font2  = Font(name="Arial", bold=True, size=13, color="1a1a2e")
body_font2   = Font(name="Arial", size=10, color="333333")
bold_font2   = Font(name="Arial", bold=True, size=10, color="333333")
small_font2  = Font(name="Arial", size=9, color="666666")

thin2   = Side(style="thin",   color="DDDDDD")
box2    = Border(top=thin2, bottom=thin2, left=thin2, right=thin2)
wrap2   = Alignment(wrap_text=True, vertical="top")
center2 = Alignment(horizontal="center", vertical="center")
vcenter2= Alignment(vertical="center")
center_wrap2 = Alignment(horizontal="center", vertical="top", wrap_text=True)

grade_colors2 = {'A':'C6EFCE','B':'90EE90','C':'FFEB9C','D':'FFB347','F':'FFC7CE'}
seg_colors2 = {
    'OVERDUE - Escalate Today': 'FFC7CE',
    'Save Now':                 'FFC7CE',
    'Champion Risk':            'FFEB9C',
    'Expand Now':               'C6EFCE',
    'Renew & Grow':             'DBEAFE',
    'Defend':                   'FFD700',
    'Ambiguous - Review':       'D3D3D3',
    'Healthy - Monitor':        'FFFFFF',
}

def style_header_row2(ws, row, col_widths, fill=None):
    fill = fill or dark_fill2
    for col_idx, (col_name, width) in enumerate(col_widths.items(), 1):
        c = ws.cell(row=row, column=col_idx)
        c.value = col_name
        c.font  = header_font2
        c.fill  = fill
        c.alignment = center_wrap2
        c.border = box2
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[row].height = 32

def add_sheet_title2(ws, title, subtitle=None):
    ws.insert_rows(1, 2)
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 16
    max_col = ws.max_column
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    c = ws.cell(row=1, column=1, value=title)
    c.font = title_font2
    c.fill = light_fill2
    c.alignment = Alignment(vertical="center", indent=1)
    c.border = Border(bottom=Side(style="medium", color="2C5282"))
    if subtitle:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
        c2 = ws.cell(row=2, column=1, value=subtitle)
        c2.font = small_font2
        c2.fill = light_fill2
        c2.alignment = Alignment(vertical="center", indent=1)

# Segment Summary
ws = wb3['Segment Summary']
col_widths_seg = {
    'Segment': 28, 'Account Count': 14, 'Total ARR ($)': 16,
    'Avg Health Score': 16, 'Avg MRR ($)': 14,
    'Avg Days to Renewal': 18, 'Accounts Needing Review': 22,
}
for i, h in enumerate(col_widths_seg.keys(), 1):
    ws.cell(row=1, column=i).value = h
add_sheet_title2(ws,
    "Segment Summary - ARR & Account Distribution by Segment",
    "Use this to understand where ARR is concentrated and which segments need leadership attention")
style_header_row2(ws, 3, col_widths_seg)
for row_idx in range(4, ws.max_row + 1):
    seg_val = ws.cell(row=row_idx, column=1).value
    fill = PatternFill("solid", fgColor=seg_colors2.get(str(seg_val), 'FFFFFF'))
    ws.row_dimensions[row_idx].height = 22
    for col in range(1, 8):
        c = ws.cell(row=row_idx, column=col)
        c.fill = fill
        c.font = bold_font2 if col == 1 else body_font2
        c.border = box2
        c.alignment = center2 if col > 1 else Alignment(vertical="center", indent=1)
        if col in [3, 5] and c.value: c.number_format = '$#,##0'
        if col == 4 and c.value:      c.number_format = '0.0'
ws3_freeze = wb3['Segment Summary']
ws3_freeze.freeze_panes = ws3_freeze.cell(row=4, column=1)

# CSM Workload View
ws = wb3['CSM Workload View']
col_widths_csm = {
    'CSM': 18, 'Accounts': 10, 'Total ARR ($)': 16, 'Avg Health Score': 16,
    'Overdue': 10, 'Save Now': 10, 'Champion Risk': 14, 'Expand Now': 12,
    '% Stale Notes': 14, 'Accounts Need Review': 20, 'Escalated Tickets': 16,
}
for i, h in enumerate(col_widths_csm.keys(), 1):
    ws.cell(row=1, column=i).value = h
add_sheet_title2(ws,
    "CSM Workload View - Book of Business by Rep",
    "Identify which CSMs have the most at-risk accounts and where to focus manager attention")
style_header_row2(ws, 3, col_widths_csm)
for row_idx in range(4, ws.max_row + 1):
    ws.row_dimensions[row_idx].height = 24
    overdue_val = ws.cell(row=row_idx, column=5).value
    row_fill = PatternFill("solid", fgColor="FFF0F0") if overdue_val and int(overdue_val) > 0 else (
               white_fill2 if row_idx % 2 == 0 else alt_fill2)
    for col in range(1, 12):
        c = ws.cell(row=row_idx, column=col)
        c.fill = row_fill
        c.font = bold_font2 if col == 1 else body_font2
        c.border = box2
        c.alignment = center2 if col > 1 else Alignment(vertical="center", indent=1)
        if col == 3 and c.value: c.number_format = '$#,##0'
        if col == 4 and c.value: c.number_format = '0.0'
        if col in [5, 6, 7] and c.value and int(c.value) > 0:
            c.fill = red_fill2
            c.font = Font(name="Arial", bold=True, size=10, color="9B0000")
ws_csm = wb3['CSM Workload View']
ws_csm.freeze_panes = ws_csm.cell(row=4, column=1)

# Human Judgment Queue
ws = wb3['Human Judgment Queue']
col_widths_hjq = {
    'Account ID': 12, 'Company': 22, 'CSM': 14, 'MRR ($)': 10,
    'Health Score': 12, 'Grade': 8, 'Segment': 24, 'Days to Renewal': 15,
    'NPS Score': 10, 'Contraction': 12,
    'Judgment Flags (Why Review Required)': 48, 'Suggested Next Action': 48,
}
for i, h in enumerate(col_widths_hjq.keys(), 1):
    ws.cell(row=1, column=i).value = h
add_sheet_title2(ws,
    "Human Judgment Queue - Accounts With Contradictory Signals",
    "These accounts have conflicting data. CSM must review the Judgment Flags column before taking action.")
style_header_row2(ws, 3, col_widths_hjq, fill=PatternFill("solid", fgColor="8B3A0F"))
for row_idx in range(4, ws.max_row + 1):
    ws.row_dimensions[row_idx].height = 52
    row_fill = white_fill2 if row_idx % 2 == 0 else alt_fill2
    for col in range(1, 13):
        c = ws.cell(row=row_idx, column=col)
        c.border = box2
        if col in [4] and c.value:
            c.number_format = '$#,##0'; c.font = bold_font2; c.alignment = center2; c.fill = row_fill
        elif col == 5 and c.value:
            c.number_format = '0.0'; c.font = body_font2; c.alignment = center2; c.fill = row_fill
        elif col == 6:
            c.fill = PatternFill("solid", fgColor=grade_colors2.get(str(c.value), 'FFFFFF'))
            c.font = bold_font2; c.alignment = center2
        elif col == 7:
            c.fill = PatternFill("solid", fgColor=seg_colors2.get(str(c.value), 'FFFFFF'))
            c.font = body_font2; c.alignment = wrap2
        elif col in [11, 12]:
            c.font = Font(name="Arial", size=9, color="333333"); c.alignment = wrap2
            c.fill = PatternFill("solid", fgColor="FFF8E7") if col == 11 else row_fill
        elif col == 8 and c.value:
            val = c.value; c.alignment = center2
            if isinstance(val, (int, float)):
                if val < 0:
                    c.fill = PatternFill("solid", fgColor="FF0000")
                    c.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
                elif val < 30:
                    c.fill = red_fill2; c.font = bold_font2
                elif val < 90:
                    c.fill = orange_fill2; c.font = body_font2
                else:
                    c.fill = row_fill; c.font = body_font2
            else:
                c.fill = row_fill; c.font = body_font2
        elif col == 10:
            val = str(c.value)
            c.fill = red_fill2 if val == 'True' else row_fill
            c.font = bold_font2 if val == 'True' else body_font2; c.alignment = center2
        else:
            c.fill = row_fill; c.font = body_font2
            c.alignment = wrap2 if col in [2, 3] else vcenter2
ws_hjq = wb3['Human Judgment Queue']
ws_hjq.freeze_panes = ws_hjq.cell(row=4, column=3)

# Expansion Pipeline
ws = wb3['Expansion Pipeline']
col_widths_exp = {
    'Account ID': 12, 'Company': 22, 'CSM': 14, 'MRR ($)': 10, 'ARR ($)': 12,
    'Plan Tier': 10, 'Seat Utilization': 15, 'Pricing Page Visits (30d)': 22,
    'Usage Trend Slope': 16, 'Health Score': 12, 'Grade': 8,
    'Days to Renewal': 15, 'Prior Expansions': 14,
    'Feature Adoption %': 16, 'Recommended Action': 48,
}
for i, h in enumerate(col_widths_exp.keys(), 1):
    ws.cell(row=1, column=i).value = h
add_sheet_title2(ws,
    "Expansion Pipeline - 13 Accounts Showing Active Buying Signals",
    "Hand off to AE. Seat capacity pressure + pricing page visits + growing usage. Sorted by MRR.")
style_header_row2(ws, 3, col_widths_exp, fill=PatternFill("solid", fgColor="1a6b1a"))
for row_idx in range(4, ws.max_row + 1):
    ws.row_dimensions[row_idx].height = 44
    row_fill = PatternFill("solid", fgColor="F0FFF0") if row_idx % 2 == 0 else PatternFill("solid", fgColor="E8F5E8")
    for col in range(1, 16):
        c = ws.cell(row=row_idx, column=col)
        c.border = box2
        if col in [4, 5] and c.value:
            c.number_format = '$#,##0'; c.font = bold_font2; c.alignment = center2; c.fill = row_fill
        elif col == 6:
            c.fill = PatternFill("solid", fgColor=grade_colors2.get(str(c.value), 'FFFFFF'))
            c.font = bold_font2; c.alignment = center2
        elif col == 7:
            val = str(c.value)
            pct = float(val.replace('%','')) if '%' in val else 0
            if pct > 100:
                c.fill = PatternFill("solid", fgColor="C6EFCE")
                c.font = Font(name="Arial", bold=True, size=10, color="1a6b1a")
            else:
                c.fill = row_fill; c.font = body_font2
            c.alignment = center2
        elif col == 8:
            val = c.value
            if val and int(val) > 2:
                c.fill = PatternFill("solid", fgColor="C6EFCE"); c.font = bold_font2
            else:
                c.fill = row_fill; c.font = body_font2
            c.alignment = center2
        elif col == 15:
            c.font = Font(name="Arial", size=9, color="333333"); c.alignment = wrap2; c.fill = row_fill
        else:
            c.fill = row_fill; c.font = body_font2; c.alignment = vcenter2
ws_exp = wb3['Expansion Pipeline']
ws_exp.freeze_panes = ws_exp.cell(row=4, column=3)

wb3.save(OUTPUT_FILE)
print(f"Output written to: {OUTPUT_FILE}")
print(f"\nSEGMENT BREAKDOWN:")
for _, r in seg_summary.iterrows():
    print(f"  {r['segment']:<30} {int(r['account_count']):>4} accts  ${r['total_arr']:>12,.0f} ARR")
