# Training

A personal, fully-versioned endurance + calisthenics training log — and the coaching system built around it.

Every session, body metric, macro total and decision lives in this repo as plain Markdown, reviewed and prescribed session-by-session with [Claude Code](https://claude.com/claude-code). `CLAUDE.md` is the coach's standing brief; `logs/session_log.md` is the source of truth.

**Races**

| | Race | Result |
|---|---|---|
| 🏁 | [Rock 'n' Roll Madrid Half Marathon — 21K](https://rocknrollmadridrun.com/recorrido-21k/) · Apr 26, 2026 | **2:09:38** (6:09/km, 339 m climb, avg HR 180) — trained for in this repo |
| 🎯 | [Run in Budapest — 30K](https://marathon.runinbudapest.com/30-km/) · Oct 11, 2026 | **Target sub-3:10** (6:20/km) · stretch sub-3:00 — *in progress* |

---

## Current block — sub-3:10 30K

| | |
|---|---|
| **Race** | Budapest 30K, Sun Oct 11, 2026 |
| **Primary** | sub-3:10 — 6:20/km |
| **Stretch** | sub-3:00 — gated on an 8 km benchmark, Fri Sep 25 |
| **Week** | 4 runs / 2 calisthenics / 3 full rest, fitted around a university timetable |
| **Constraint** | Right-tibia stress-fracture history. **Distance is set by bone, not by breath.** |

**The core read:** a 25:58 5K predicts a 1:59 half; Madrid came in at 2:09:38. That ~10-minute spread isn't a speed problem — it's an endurance deficit from treadmill-only, low-volume training. Volume is the lever that closes it, so the plan optimises for *zero dark weeks* over hero sessions. Two goals have already died to month-long training gaps (a 50K ultra, then a sub-2:50 30K); consistency is the whole thesis.

**And the hard part:** peak long run is ~22 km against a 30 km race. That's +36% over the longest run ever done in the block — a finishing problem first, a pace problem second.

---

## Repo map

```
├── CLAUDE.md               ← the coaching brief: active goal, rules, gates, logging conventions
├── PLAN.md                 ← Block 1 phase structure + week-by-week (archived)
├── workouts/
│   ├── cali_fullbody.md    ← ACTIVE — full-body calisthenics, Mon + Fri AM
│   └── {upper,lower}_{A,B}.md  ← dormant barbell split (no longer have the gym)
├── docs/
│   ├── running.md          ← the running plan: weekly km, pacing, checkpoints
│   ├── nutrition.md        ← macro targets + cut/maintenance logic
│   ├── handoff.md          ← archived block-to-block handoff
│   └── block2_plan_prompt.md  ← the research prompt used to generate the current block
├── logs/
│   ├── session_log.md      ← SOURCE OF TRUTH — daily narrative journal
│   ├── sessions/           ← one filled log per session
│   ├── templates/          ← locked AM / PM / session / weekly formats
│   └── training_log.xlsx   ← generated workbook (sets, reps, load, HR, body)
└── scripts/
    ├── build_log.py        ← regenerate the workbook from the logs
    └── progression_check.py ← analyse progression trends
```

---

## How the system works

**Two-touch daily logging.** An AM block on waking (bodyweight fasted, sleep, resting HR vs baseline, motivation/energy/stress, per-muscle DOMS, joint flags) and a PM block before bed (session result, macros, bike km). ~4 minutes total. Templates are locked — no freeform entries, so the data stays comparable across months.

**A tracked-variable hierarchy.** Tier 1 (working sets per muscle, RPE, load/e1RM, protein, kcal) is logged every session; Tier 2 is daily; Tier 3 is weekly. What gets logged gets progressed.

**Explicit adjustment triggers**, written down *before* they're needed, so decisions aren't made on a bad morning: sleep <6 h + resting HR +10 → cut volume 20% · bodyweight loss >0.7 kg/wk for 2 weeks → eat more · HR drift +5 bpm at the same easy pace → drop a run · joint pain >3 days → swap the movement, never push through · fever → zero training, no exceptions.

**A weekly audit that hunts under-training.** Most training systems only catch over-reaching. Every Sunday this one also runs the reverse check — is the plan *lagging* an engine that's adapting faster than modelled? Madrid was under-run because a conservative read left fitness un-cashed; the audit exists so that doesn't repeat.

**Progression by the slowest increment that still moves.** All sets at the top of the rep range *and* every set at RPE ≤7 → then, and only then, the smallest available bump. Reps before load; bodyweight movements progress by leverage, not weight. Horizon is years, injury count is zero.

---

## Things this log got wrong (kept on purpose)

The failures are version-controlled alongside the wins, because they're the useful part:

- **A 50K ultra, abandoned** — 17 dark days left an 18 km ceiling 22 days out.
- **"Sub-2:20 30K", retired** — the number came from naive addition (21 km in 2:09 → "30 km in 2:20"), never a pace conversion. It implied a 1:36 half against a 2:09 PR.
- **Sub-2:50, retired** — killed by a 34-day gap, not by a bad session.
- **A job that was never modelled** — 44 km/week of walking as a waiter, logged as "steps" instead of as training load. The same mistake nearly repeated with a 96 km/week bike commute; it's now Tier 1 load.
- **Cadence, still unsolved** — a 166 metronome holds at 7:10/km and fails at 8:00/km. Turnover is pace-linked, not habitual.

---

## Reproducing the tooling

```bash
pip install openpyxl
python scripts/build_log.py         # rebuild logs/training_log.xlsx from the logs
python scripts/progression_check.py # progression analysis
```

---

*Personal training data — bodyweight, heart rate, sleep and illness notes — is published deliberately. Reuse the structure freely; the numbers are one person's and are not a prescription for anyone else.*
