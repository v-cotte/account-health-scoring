# Account Health Scoring & Churn Alerts

A customer-success intelligence system for a B2B SaaS book of business: it scores every account 0–100, segments it into a concrete next action, and pushes AI-written alerts to Slack when a renewal is at risk because the customer's champion has gone dark.

Built in Python (pandas) with the Claude API, Slack webhooks and an n8n workflow for scheduling. A GTM engineering case study on a synthetic dataset of ~190 accounts, 830 contacts, 480 support tickets and 370 CSM notes.

## What it does

**1. Health scoring engine** (`score_accounts.py`)

Each account gets a composite score from three signal tiers:

| Tier | Weight | Signals |
|---|---|---|
| Engagement | 40% | Usage trend (6 mo), seat utilization, feature adoption, champion login frequency |
| Relationship | 30% | NPS, CSM note recency and quality, exec sponsor presence and changes |
| Risk | 30% | Contraction, competitor mentions, escalated tickets, ticket severity, discounting |

Two design choices make it more than a weighted average:

- **Contradictions are flagged, not averaged away.** When engagement and relationship tiers diverge by more than 25 points, or NPS contradicts behavior (high NPS + contraction, low NPS + expansion), the account goes to a human-review queue with the specific contradiction spelled out.
- **Urgency overrides the score.** A score of 65 isn't "safe" if the renewal was yesterday. Overdue renewals, renewals under 30 days and dark-champion cases override the segment.

Accounts land in one of eight segments (Overdue, Save Now, Champion Risk, Expand Now, Renew & Grow, Defend, Ambiguous, Healthy), each with an owner and a next action.

**Output:** [`outputs/account_intelligence_list.xlsx`](outputs/account_intelligence_list.xlsx), with a prioritized account list, segment summary, CSM workload view, human-judgment queue, expansion pipeline and a methodology tab.

**2. Champion dark alerts** (`champion_alert_automation.py`)

A compound trigger that catches accounts that look fine in aggregate but carry a specific human risk:

> champion login frequency is Never/Quarterly **AND** renewal is within 90 days **AND** the last CSM note is older than 21 days

Any one of these alone is noise. Together they're the blind spot where CS finds out about churn when the customer doesn't renew. For each triggered account, the script calls the Claude API to write a context-specific Slack alert with a suggested next step, falling back to a deterministic template if the API is unavailable.

**3. n8n workflow** (`champion_dark_alert_n8n_workflow.json`)

The same alert logic as a scheduled production workflow: weekday 8am trigger → fetch accounts from CRM → evaluate trigger → Claude → Slack → audit log, with an error branch that notifies engineering.

See [ARCHITECTURE.md](ARCHITECTURE.md) for the production design, failure modes and roadmap.

## Running it

```bash
pip install -r requirements.txt

# Scoring engine → outputs/account_intelligence_list.xlsx
python score_accounts.py

# Alerts: dry run by default (prints alerts, sends nothing)
python champion_alert_automation.py

# Post to Slack
cp .env.example .env   # add ANTHROPIC_API_KEY and SLACK_WEBHOOK_URL
python champion_alert_automation.py --send
```

Both credentials are optional: without an API key alerts use the fallback template, and without a webhook nothing is posted.

To use the n8n workflow: *New workflow → ⋯ → Import from file*, then set the Anthropic and Slack credentials in the HTTP nodes.

## Project structure

```
├── score_accounts.py                      scoring engine
├── champion_alert_automation.py           alert detection + Claude + Slack
├── champion_dark_alert_n8n_workflow.json  scheduled workflow version
├── data/accounts_dataset.xlsx             synthetic input data (5 sheets)
├── outputs/account_intelligence_list.xlsx sample output
└── ARCHITECTURE.md                        production design and roadmap
```
