# Training — 11-Week Hypertrophy Block

> 4 days lifting + 2 days running. Bridge from Madrid HM recovery to **sub-2:50 30K** block (race Oct 11).

**Block dates:** April 27 – July 12, 2026 (11 weeks)
**Goal:** Maximum upper-body hypertrophy. +3–4 kg lean mass realistic, visible chest/back/shoulder/arm change.
**Approach:** Recomp at maintenance calories, RPE 8 from session 1 of accumulation, intensity techniques week 5+.
**Block 2:** July 13 – October 11, 2026 (13 weeks). **Sub-2:50 30K** target (5:40/km; floor sub-3:00). Race Sun Oct 11, 2026.

---

## Repo Map

```
training/
├── README.md                   ← you are here
├── PLAN.md                     ← 11-week phase structure + week-by-week
├── workouts/
│   ├── upper_A.md              ← Monday: chest/back priority
│   ├── lower_A.md              ← Tuesday: squat focus
│   ├── upper_B.md              ← Thursday: shoulders/arms priority
│   └── lower_B.md              ← Friday: hinge/posterior
├── docs/
│   ├── running.md              ← running plan (Wed + Sat)
│   ├── nutrition.md            ← recomp principles
│   └── handoff.md              ← STALE: written for a sub-2:00 HM block. Rewrite for sub-2:50 30K.
├── logs/
│   ├── training_log.xlsx       ← master logger (multi-tab)
│   └── session_log.md          ← daily narrative journal
└── scripts/
    ├── build_log.py            ← regenerate workbook
    ├── progression_check.py    ← analyze progression from log
    └── README.md               ← how to run scripts
```

---

## Weekly Schedule

| Day | Session | Duration |
|-----|---------|----------|
| Mon | Upper A (chest/back priority) | ~75 min |
| Tue | Lower A (squat focus) | ~75 min |
| Wed | Full rest | — |
| Thu | **Easy run** outdoor Z2 | 30–45 min |
| Fri | Upper B (shoulders/arms priority) | ~75 min |
| Sat | Lower B (hinge/posterior) | ~75 min |
| Sun | **Easy run** outdoor Z2 (slightly longer) | 35–50 min |

---

## Quick Start

1. **Read** `PLAN.md` — the 11-week phase structure
2. **Open** `logs/training_log.xlsx` — log every set/run/body metric here
3. **Each session day**, open the relevant workout file in `workouts/`
4. **Check Wed/Sat against** `docs/running.md` for that week's run prescription
5. **Weekly review:** Sundays, run `python scripts/progression_check.py`
6. **End of Week 11 (Sun Jul 12):** lift retest + body measurements + Block 2 handoff

---

## Core Rules

- **Track every set.** What gets logged gets progressed.
- **RPE 8 on working sets in accumulation+ phases.** Top sets close to limits on isolation. 1–2 RIR on compounds.
- **Aggressive double progression.** Top of rep range on all sets at prescribed RPE → +2.5 kg compound, +1–2 kg isolation NEXT session. Don't sit at the same weight for 2 weeks.
- **Sleep 7+ hrs.** Recomp + hypertrophy fails without it.
- **Don't chase the run.** These 2 runs are aerobic maintenance, not training. Easy means easy. Hold Z2.
- **When in doubt, push.** Adjust down only if fatigue markers actually fail (sleep, RHR, lift performance).

---

## Why this structure

- **4 lifts vs 2 runs:** running was suppressing hypertrophy recovery during the half-marathon block — flipped ratio for 11 weeks.
- **Upper twice, lower twice:** maximizes hypertrophy frequency for priority muscle groups while still hitting legs.
- **Runs Wed & Sat:** both fall the day after a lower session — legs not pre-fatigued for quality lift, runs aid recovery rather than blunt it.
- **Recomp not bulk:** finished race lean. Add muscle slowly without piling fat to cut later before race block.

🏁 Block 1 ends Sun July 12, 2026. Block 2 (**sub-2:50 30K**) starts Mon July 13. Race Sun October 11, 2026.
