# Scripts

Helpers for the training repo. Run from anywhere — both scripts resolve paths relative to repo root.

## Setup

```bash
pip install pandas openpyxl
```

## `build_log.py` — regenerate the workbook

Builds `logs/training_log.xlsx` from scratch with all 7 tabs and ~3000 formulas.

```bash
python scripts/build_log.py
```

> **Warning:** overwrites the existing log. Don't run this once you've started logging real data unless you want to wipe and start over.

## `progression_check.py` — weekly progress report

Reads `logs/training_log.xlsx`, prints clean terminal report: weekly summary, PRs, e1RM trend per exercise, running trend, body metrics trend.

```bash
# full report
python scripts/progression_check.py

# focus on one week
python scripts/progression_check.py --week 4

# focus on one exercise
python scripts/progression_check.py --exercise "Bench Press"
```

Read-only. Safe to run anytime.

## When to run

- **Sundays:** weekly review with `progression_check.py`
- **End of block (Jun 22):** final report — feed the output into `docs/handoff.md` checklist
- **Never re-run `build_log.py` after Apr 28** unless wiping the log
