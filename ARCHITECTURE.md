# Architecture

*How this system would run in production, where it breaks, and what comes next.*

## From script to system

Today the source of truth is an Excel export: run the script, it processes the data, you get the output. In production that changes. The CRM becomes the system of record for account identity, renewal dates, CSM ownership and notes. Product telemetry syncs daily from the product database, support tickets update every few hours via webhook, and CSM notes sync in real time as they are written. The scoring pipeline runs nightly after all feeds complete and writes fresh scores to a database, triggered on a schedule (n8n or cron) so the output refreshes every morning without anyone running anything. The spreadsheet in `outputs/` simulates that nightly output: a structured view of the customer base that any CSM can open and act on.

## Identity vs. signals vs. computed data

The key design decision is separating three kinds of data:

| Layer | Examples | Behavior |
|---|---|---|
| Identity | Accounts, contacts | Stable records in the CRM, change slowly |
| Signals | Usage events, support tickets, CSM notes | Time-stamped, append-only, never overwritten |
| Computed | Health scores | Stored as a time series, one row per account per day |

Keeping scores as a time series shows whether an account is improving or declining, not just where it stands today. The separation keeps source data clean, makes the scoring logic auditable, and lets you recalibrate the model without touching raw inputs: change a weight or add a signal, rerun against the same data, and history is preserved.

## Failure modes

1. **Silent pipeline failure.** A batch job stops but scores look fine because nobody checks. Fix: a freshness check that alerts engineering if any feed hasn't updated in 36 hours. Never serve stale scores without surfacing the staleness.
2. **CSM notes not being written.** 34 accounts in the sample data have zero notes. That's a behavioral problem, not a technical one: make notes easier to write and surface coverage gaps to CS managers.
3. **New accounts without history.** A month-one account has no usage trend, no NPS and no notes, so scoring it like a mature account produces noise. These belong in an onboarding segment with completion rate as the primary signal.
4. **Champion misclassification.** If contacts aren't tagged correctly in the CRM, the dark-champion alert fails silently and the most important signal disappears without warning.

## Next 30 days

1. **Note intelligence.** Use the Claude API to score CSM note quality on save, flag low-signal entries ("all good", "checked in"), and suggest follow-up questions based on account risk. Notes are the weakest input, and improving them compounds into every downstream score and alert.
2. **Expansion automation.** When seat utilization crosses 85% and a pricing-page visit is detected, create an expansion opportunity in the CRM and notify the AE. Thirteen accounts in the sample data are ready for that conversation now.
3. **Alert outcome tracking.** Close the loop on every alert: did the CSM act, did the account renew? Without outcome data the model never learns and stays a static snapshot.
