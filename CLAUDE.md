# CLAUDE.md — Training Repo Context

Project memory for Claude Code. Read on every session.

---

## ⚠️ ACTIVE OVERRIDE — 2026-06-04: 50 KM ULTRA BUILD

**Goal pivoted.** User now training for a **50 km run/walk on Sat Aug 8, 2026** (~9-week build). This **supersedes Block 1 hypertrophy + the slow cut** for now. Until the user says otherwise:

- **Schedule:** 3 gym (Push / Legs / Pull, maintenance RPE 7–8, not to failure) + 3 run (Fri 5k social easy / Sat LONG / Sun medium back-to-back) + Thu rest. See table below + `docs/running.md` (50 km build) + `logs/session_log.md` (Build weeks).
- **Nutrition: CUT IS PAUSED.** Eat **~maintenance (~2700)**, carbs up on run/leg days, protein high. BW drifting UP is expected and correct — **do NOT flag BW gain or "deficit missed".** The SLOW CUT config below is dormant.
- **Lifting:** retain strength, don't chase PRs / aggressive bumps; running owns the legs. Trim Tue Legs in peak weeks (Sat ≥32 km).
- **Running priority:** the long run is the driver. Run/walk from start, easy Z2, in-run carbs 30–60 g/hr on runs >90 min. Pinpoint bone pain = STOP (stress-fracture risk). Never skip step-back weeks.
- Hypertrophy/cut sections below remain for reference + the eventual return to a strength block (or Block 2 sub-2:20 30K, race Oct 11). Resume only when the user says.

---

## What this repo is

Personal training log + plan for **Block 1: 11-Week Hypertrophy** (Apr 27 – Jul 12, 2026). Bridge from Madrid HM 2026-04-26 (2:09:38) to Block 2 sub-2:20 30K (Jul 13 – Oct 11, 2026; race Sun Oct 11).

User: solo lifter/runner, novice/early-intermediate strength tier, BW ~77 kg. **Goal (set 2026-05-17): slow cut to ~73 kg while getting stronger — deliberate deficit, high protein. See Block 1 targets + Nutrition baselines.**

---

## Repo layout

```
PLAN.md                     — 11-week phase structure, baselines, week-by-week
README.md                   — overview + repo map
workouts/{upper_A,lower_A,upper_B,lower_B}.md   — session prescriptions
docs/{running,nutrition,handoff}.md             — supporting plans
logs/training_log.xlsx      — master data log (sets, reps, weight, HR, body)
logs/session_log.md         — daily narrative journal, per-week schedule
scripts/{build_log,progression_check}.py        — workbook tooling
```

**Source of truth for current week's schedule and bumps:** `logs/session_log.md`. PLAN.md is the static plan; the session log carries swaps, bumps, and locked subs.

---

## Weekly schedule — ACTIVE (50 km build, from 2026-06-04)

| Day | Session |
|-----|---------|
| Mon | Gym — Push (upper, maintenance) |
| Tue | Gym — Legs (maintenance; trim in peak run wks) |
| Wed | Gym — Pull (upper, maintenance) |
| Thu | REST |
| Fri | Run — 5k social (EASY, fixed) |
| Sat | Run — LONG (the driver) |
| Sun | Run — medium-long (back-to-back) |

User may swap days week-to-week. Always read `logs/session_log.md` for the active week before prescribing today's session.

> Prior Block-1 hypertrophy schedule (dormant): Mon Upper A · Tue Lower A · Wed REST · Thu easy run · Fri Upper B · Sat Lower B · Sun easy run.

---

## How to prescribe a session

1. Read `logs/session_log.md` → find today's date → confirm planned session (incl. swaps).
2. Read the matching `workouts/<session>.md` for sets/reps/RPE/form cues.
3. Pull last-session loads + bump notes from `logs/session_log.md` ("Next session bumps:" lines).
4. Apply progression rule below. Output table: exercise, sets×reps, RPE, **starting load this session**, notes.

### Progression rule (aggressive double progression)

- Top of rep range hit across all sets at prescribed RPE → **+2.5 kg compound, +1–2 kg isolation NEXT session.** No waiting.
- First set RPE <7 at prescribed weight → **bump mid-session.** Don't finish at warm-up intensity.
- Reps stall 3 sessions same weight → vary rep range OR drop 5% and rebuild OR add intensity technique.

### Intensity techniques (Wk 5+)

Drop sets, myo-reps, lengthened partials. Isolation only Wks 5–7. Compounds Wk 8+. See PLAN.md §Intensity.

---

## Locked exercise subs (do not suggest swaps)

Upper A: **Pec Deck** (vs Cable Fly), **Face Pull** (vs Rear Delt Fly), **DB Curl** (vs Hammer).
Lower A: **Lying Leg Curl** (vs Seated).

Other movements: subs in workout file's Substitutions table are fair game.

---

## Logging conventions

**Two-touch daily logging:** AM (on wake) + PM (before bed). Standardized templates in `logs/templates/{session,daily,weekly}.md`. Don't freeform.

- **AM block (every day):** BW AM (kg), sleep h + quality, RHR + Δ vs baseline, motivation, energy, stress, DOMS map, joint flags. ~60 sec. **BW is logged AM only — never ask for or log BW in the PM block.**
- **PM block (every day):** workout result if lifting, day macros total. ~3 min.
- **Lifting day** → fill `templates/session.md` (locked prose-block format) → save filled file as `logs/sessions/YYYY-MM-DD_<session>.md` (e.g., `logs/sessions/2026-05-07_upper_B.md`).
- **Rest / run-only day** → `templates/daily.md` (AM block + PM block w/ run data + macros).
- **Sunday PM** → also fill `templates/weekly.md` (compliance scoreboard, volume by muscle, e1RM trend, body trend, decisions).
- **Data** (sets, reps, load, RPE, HR, BW) also → `logs/training_log.xlsx`.
- **Narrative + subjective** → `logs/session_log.md`. One bullet per day with summary + bumps + link to full filled log in `logs/sessions/`. Status: **DONE / PARTIAL / SKIPPED / MOVED / REST**.
- After every lifting session: explicit "Next session bumps:" line per lift in both the filled session file and the session_log.md bullet.

**BW measurement rule:** BW taken **AM only**, fasted, post-pee, same conditions daily. Logged in AM block exclusively. Use 7-day rolling avg for trend, ignore single-day spikes. Never request or record evening/PM BW.

### Variable hierarchy (track this priority order)

**Tier 1 — log every session (PM block on lifting days):**
1. Working sets/muscle/week (volume = primary hypertrophy driver)
2. RPE/RIR per set (proximity to failure)
3. Load + e1RM trend (mechanical tension)
4. Protein g/kg BW (target 1.6–2.2 → 125–170 g @ 77.7 kg)
5. Kcal vs target — **slow cut: deliberate deficit ~−400/d** (see Nutrition baselines)

**Tier 2a — AM daily:**
6. BW AM (fasted, post-pee, conditions held constant)
7. Sleep h + quality
8. Resting HR (Δ vs 7-day baseline)
9. Motivation / energy / stress 1–10
10. Per-muscle DOMS 0–3
11. Joint/tendon flags

**Tier 2b — PM daily:**
12. Day macros total
13. Workout result (lifting days)

**Tier 3 — weekly/biweekly:**
14. Pre-WO meal timing + composition
15. Creatine 5 g/d compliance
16. Hydration L, steps, caffeine, alcohol
17. Body circumferences (biweekly Mon)
18. Photos front/side/back (biweekly)
19. HR drift on easy runs

### Nutrition baselines — SLOW CUT ⏸️ PAUSED 2026-06-04 (see top override)

> **PAUSED for the 50 km build.** During the build eat ~maintenance (~2700), carbs up on run/leg days, protein high; **BW gain is expected — do NOT flag.** The cut config below resumes only when the user returns to a strength/cut block.

Goal as of 2026-05-17 (dormant): deliberate slow fat loss to ~73 kg while getting stronger. NOT recomp, NOT surplus. Deficit is intended — do not flag it as under-fueling.

| Macro | Target |
|-------|--------|
| Kcal | maintenance (~2700) **minus ~400 → ~2250–2400/d**; recalc maint per BW 7-day trend. Target loss **0.3–0.5 kg/wk** |
| Protein | **150–185 g (≥2.0 g/kg, push high — #1 muscle-retention lever in deficit)**, 4–5 meals, ≤5 h between feeds. 180+ g = good, never flag as "over". |
| Carbs | 3–5 g/kg lift days (230–390 g), 2–3 g/kg rest (155–230 g) |
| Fat | 0.8–1.0 g/kg (62–78 g) |
| Fiber | 25–35 g |
| Creatine | 5 g/d, every day |
| Pre-WO | carbs 1–2 g/kg + 30–40 g protein, 1–3 h prior |
| Post-WO | 30–40 g protein within 2 h |

### Adjustment triggers

- Sleep <6 h **and** RHR +10 → cut session volume 20%, RPE cap 7.
- **Cut pace (7-day avg):** loss 0.3–0.5 kg/wk = on target, hold. Loss >0.7 kg/wk for 2 wks → too fast, muscle risk → **+150–200 kcal/d**. BW flat ≥2 wks (stall) → **−150 kcal/d OR +1–2k steps**.
- **Strength = the cut guardrail:** e1RM regression OR top-set reps drop at same load 2 sessions running (not explained by sleep/illness) → cut too hard → diet break or +200 kcal/d. Strength holding/rising = cut is working, do not add food just because BW dropped.
- 3 sessions reps stalled same load → vary rep range OR drop 5% rebuild OR add intensity technique.
- Joint pain >3 days → swap variation, no push-through.
- HR drift +5 bpm at same easy pace vs 2 wks ago → drop a run.

---

## Working style with the user

- Terse. No fluff. Tables > prose for prescriptions.
- Always seed starting loads — never leave "TBD" on a session.
- Show RPE target + load target side-by-side.
- When user reports yesterday's session result → update `logs/session_log.md` with data + bumps + close that day's bullet.
- Do not invent baselines; pull from PLAN.md §Strength Baselines or last logged session.
- When user asks "give me today's session," output: filled `templates/session.md` §2 table (working-set prescription with seeded loads) + the §1 readiness checklist for them to fill on arrival. Don't dump full template unless asked.
- Treat user as athlete; you are the research team. Surface data trends (e1RM slope, BW 7-day avg, sleep avg, protein adherence %) every Sunday review without being asked.
- Flag anomalies proactively (RHR spike, missed protein 3+ days, **e1RM/strength regression**, **cut too fast >0.7 kg/wk**, cut stalled ≥2 wks). Steady BW loss 0.3–0.5 kg/wk is the GOAL — report as progress, never as a problem. A kcal deficit is intended; do not prescribe "eat more" unless a guardrail above trips.

---

## Block 1 targets

**Revised 2026-05-17 → SLOW CUT:** BW ~77 → **~73 kg** (lose fat, ~0.3–0.5 kg/wk) | **retain/build strength through the deficit** — bench/DL/weighted-PU still progressing (smaller jumps OK in deeper deficit, a stall ≠ failure) | preserve lean mass (high protein + hard lifting) | running fitness held 90–95% race-day. Cut likely extends past Block 1 (Jul 12) into Block 2 — do not force the full drop by block end.

Block ends Sun Jul 12, 2026 → measurements + lift retest + read `docs/handoff.md` for Block 2.

> **Superseded 2026-06-04** by the 50 km ultra build (see top override). Block-1 hypertrophy/cut targets are on hold; the 50 km run/walk (Sat Aug 8) is the active goal. Strength = retain, not progress, during the build.
