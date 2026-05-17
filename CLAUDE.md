# CLAUDE.md — Training Repo Context

Project memory for Claude Code. Read on every session.

---

## What this repo is

Personal training log + plan for **Block 1: 11-Week Hypertrophy** (Apr 27 – Jul 12, 2026). Bridge from Madrid HM 2026-04-26 (2:09:38) to Block 2 sub-2:20 30K (Jul 13 – Oct 11, 2026; race Sun Oct 11).

User: solo lifter/runner, novice/early-intermediate strength tier, BW ~77.7 kg.

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

## Weekly schedule (Wk 2+)

| Day | Session |
|-----|---------|
| Mon | Upper A |
| Tue | Lower A |
| Wed | REST |
| Thu | Easy run Z2 (30–45 min) |
| Fri | Upper B |
| Sat | Lower B |
| Sun | Easy run Z2 (35–50 min) |

User may swap days week-to-week. Always read `logs/session_log.md` for the active week before prescribing today's session.

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
5. Kcal vs maintenance (recomp = ±100)

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

### Nutrition baselines (BW 77.7 kg)

| Macro | Target |
|-------|--------|
| Kcal | maintenance ±100 (~2700 baseline; recalc per BW 7-day trend) |
| Protein | 125–170 g (1.6–2.2 g/kg), spread 4–5 meals, ≤5 h between feeds |
| Carbs | 3–5 g/kg lift days (230–390 g), 2–3 g/kg rest (155–230 g) |
| Fat | 0.8–1.0 g/kg (62–78 g) |
| Fiber | 25–35 g |
| Creatine | 5 g/d, every day |
| Pre-WO | carbs 1–2 g/kg + 30–40 g protein, 1–3 h prior |
| Post-WO | 30–40 g protein within 2 h |

### Adjustment triggers

- Sleep <6 h **and** RHR +10 → cut session volume 20%, RPE cap 7.
- BW drop >0.5 kg/wk for 2 wks → +200 kcal/d.
- BW gain >0.5 kg/wk for 2 wks → −200 kcal/d.
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
- Flag anomalies proactively (RHR spike, BW drift, missed protein 3+ days, e1RM regression).

---

## Block 1 targets

+3–4 kg lean mass | ±0.5 kg BW total | bench +10–15 kg | DL +15–20 kg | weighted PU +10–15 kg added load | running fitness held 90–95% race-day.

Block ends Sun Jul 12, 2026 → measurements + lift retest + read `docs/handoff.md` for Block 2.
