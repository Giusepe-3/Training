# CLAUDE.md — Training Repo Context

Project memory for Claude Code. Read on every session.

---

## ⚠️ ACTIVE OVERRIDE — 2026-07-17: BLOCK 2 — SUB-2:50 30K (race Sun Oct 11, 2026)

**Two prior goals are DEAD. Do not resurrect either without the user saying so.**
- ❌ **50 km ultra (Aug 8) — ABANDONED 2026-07-17.** User called it after a 17-day dark stretch (Jul 1–16, cycling only, zero running) left an 18k ceiling 22 days from race day. See `logs/session_log.md` § "50 KM ULTRA BUILD — ABANDONED".
- ❌ **"Sub-2:20 30K" — RETIRED as physically implausible, 2026-07-17.** 2:20 over 30K = **4:40/km**; the user's HM PR is **6:09/km**. It required a 1:36 half / 20:57 5K. The number was naive addition (21.1 km in 2:09 → "30 km in 2:20"), never a pace conversion. **If any doc still says "sub-2:20 30K", it is stale — fix it.**

**ACTIVE GOAL: 30K in sub-2:50 (5:40/km), Sun Oct 11, 2026.** Floor = sub-3:00. Set by user 2026-07-17.

- **Why it's real (the core read — keep this in mind for every prescription):** the user's 5K PR (25:58 @ 5:12/km, 2026-05-08) predicts a **1:59 HM**; he actually ran **2:09:38**. ~10 min slower than his own speed says. **He is not slow — he is under-endured** (treadmill-trained, low volume). Riegel anchors disagree for exactly this reason: 5K → 2:53 30K, HM → 3:08. **The spread IS the deficit, and volume is the lever that closes it.** Sub-2:50 ≈ a 1:57 HM by October = collecting endurance already paid for, not building new top-end speed.
- **Therefore: consistent weekly volume > any hero session.** This target dies from dark weeks, not from slow days. Jul 1–16 is the cautionary tale.
- **⚠️ THE JOB IS A TRAINING LOAD — added 2026-07-27. User is a WAITER: 12 000+ steps on a rest day, 17–30 k on a shift day.** W1 measured 78.5 k steps / 6 days ≈ **~44 km walked vs 16 km run** — the job is ~2.75× the running distance and ~3.3× the foot strikes, and it was completely unmodeled when the plan was built. **Bone: fine, arguably protective** (habitual ~1.2× BW loading; shin clean 8 straight checks). **Soft tissue + systemic recovery: this is the real cost** — calves/Achilles/feet never unload, 8 h standing = zero recovery hours, a "rest day" is still ~9 km on hard floors. Never read steps as a NEAT/kcal footnote again; **steps are Tier 1 LOAD** (see hierarchy below). Full analysis + rules → `docs/running.md` § JOB LOAD.
- **Nutrition: CUT RESUMES (goal-aligned now, not competing).** 77.6 → ~73 kg ≈ 5.9% BW ≈ **2–3% running economy ≈ 8–10 s/km ≈ 4–5 min over 30K** — roughly a third of the gap to 2:50, from the kitchen. Slow cut 0.3–0.5 kg/wk through the base weeks → **maintenance for the final 3–4 wks + taper** (no deficit while sharpening). ⏸️ **NOT started as of 2026-07-17** — re-entry week eats maintenance (~2700); start once the running rhythm holds. See Nutrition baselines below (now live again).
- **Lifting: maintenance, slow-steady.** Running owns the legs. Bump gate = ALL sets at top of rep range AND every set RPE ≤7 → smallest increment. Reps before load. See "Progression rule" below — the **slow-steady rule is ACTIVE; the aggressive double-progression text is DORMANT**, user reconfirmed 2026-07-17.
- **⚠️ TIBIA GATES EVERYTHING.** Right-tibia stress-fx watch (May–Jun) rehabbed clean, but it is the standing constraint. Pinpoint bone pain = STOP. **Engine ≠ chassis:** cycling held his aerobic system (~70–80% transfer) but zero impact tolerance — **runs will feel easy before the bone is ready. HR/RPE are not the limiter; impact tolerance is. Set distance by bone, not by breath.**
- **⚠️ Legs untrained since 2026-05-19** (never trained in the whole 50 km build). Cycling quads ≠ trained quads. Seed very light, keep off run-adjacent days, expect DOMS.
- **Open miss carried from the ultra build: cadence.** Parked 161–164 all build; hit 166 once (6/28) at fast pace only, then regressed to 159. Turnover is pace-linked, not habitual. Still unsolved.

---

## What this repo is

Personal training log. **Now: Block 2 — sub-2:50 30K (Jul 13 – Oct 11, 2026; race Sun Oct 11).** History: Block 1 11-Week Hypertrophy (Apr 27 – Jul 12) → 50 km ultra detour (Jun 4 – Jul 17, abandoned) → Block 2. Bridge from Madrid HM 2026-04-26 (2:09:38).

User: solo lifter/runner, novice/early-intermediate strength tier, BW ~77.6 kg (2026-07-17). **Goal: sub-2:50 30K Oct 11 + slow cut to ~73 kg (cut now SUPPORTS the race goal — see override above).**

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

## Weekly schedule — Block 2 ✅ LOCKED (revised 2026-07-27, shift-fitted)

**4 run / 2 lift / 3 rest.** Full plan + week-by-week km → `docs/running.md`.

**Shift schedule (supplied 2026-07-27):** Mon 17:30–23:00 (5.5 h) · Tue 11:00–17:00 (6 h) · **Wed 13:00–22:00 (9 h ← longest)** · Thu 17:30–23:00 (5.5 h) · **Fri/Sat/Sun OFF.** 26 h/wk, all Mon–Thu.

| Day | Shift | Session | Timing |
|-----|-------|---------|--------|
| Mon | 5.5 h eve | **REST — full** | post-long, on a work day |
| Tue | 6 h midday | Run easy Z2 **+ Gym Legs (light)** | **POST-shift ~18:00** — never before |
| Wed | **9 h longest** | **REST — full** | biggest step day of the week |
| Thu | 5.5 h eve | Run easy — **shortest run of the week** | **pre-shift ~13:00** (follows Wed's 9 h → easy+short by rule) |
| Fri | **OFF** | Run — **quality from W5** (easy W1–4) **+ Gym Upper** | free day, best-recovered. Checkpoint slot W4/W8 |
| Sat | **OFF** | **REST — full** | long-run primer |
| Sun | **OFF** | Run — **LONG** | free morning, unobstructed |

**Why revised (2026-07-27):** old template (Mon Legs · Tue/Wed run · Thu Upper · Fri REST · Sat run · Sun LONG) opened a 4-day unbroken load block and ran **Sat easy + Sun long back-to-back** — the goal-critical session always landed on tired legs, on top of a 20 k-step job. That pair failed Jul 25–26 (Sat RPE 7 on a 5 k → Sun long skipped → Mon 7/27 rest). New shape: same km, same session count, **3 full rest days (2 of them on work days), max 2 consecutive loaded days, long run off a rest day off a work-free Saturday.**

- **Rest lands Mon + Wed + Sat.** Wed is non-negotiable — it is the 9 h shift and the biggest step day.
- **Legs never precedes a shift.** Tue Legs runs *after* the 17:00 finish. **Tuesday is the one deliberate run-after-standing exception** (no usable pre-shift window, wide-open evening) and is therefore always the *easy* run, never quality.
- **Quality is Friday, not Thursday** — Thursday follows the 9 h Wednesday and would break the >22 k-steps rule. Fri is work-free, follows the shortest shift, and has Sat rest between it and the long.
- **Circadian, honestly:** Mon/Wed/Thu finish 22:00–23:00 → **a 1–2am bedtime those 3 nights is structural, not a discipline failure. Do not nag it.** The only night worth targeting is **Saturday** — no shift Sat, no shift Sun, and it's a rest day, sitting directly before the block's most important session.
- **Standing Calf Raise DROPPED W2–W4** (reassess W5) — 15+ km/day of loaded walking already trains soleus/gastroc to failure; it had stalled at 20 kg; and the 7/27 AM check showed every leg tissue clearing to ≤1 after two rest days **except calf, which held at 2.** **Tibialis Raise stays every Legs day, non-negotiable** (anterior tib = bone armor, walking does not train it).

> Dormant: 50 km build schedule (3 gym / 3 run, Sat LONG driver) — goal abandoned. Block-1 hypertrophy schedule (Mon Upper A · Tue Lower A · Wed REST · Thu run · Fri Upper B · Sat Lower B · Sun run) — block over.

Always read `logs/session_log.md` for the active week before prescribing today's session.

---

## How to prescribe a session

1. Read `logs/session_log.md` → find today's date → confirm planned session (incl. swaps).
2. Read the matching `workouts/<session>.md` for sets/reps/RPE/form cues.
3. Pull last-session loads + bump notes from `logs/session_log.md` ("Next session bumps:" lines).
4. Apply progression rule below. Output table: exercise, sets×reps, RPE, **starting load this session**, notes.

### Progression rule — SLOW STEADY ✅ ACTIVE (user directive 2026-06-10, reconfirmed 2026-07-17)

User, verbatim: *"I want a slow steady increase in the gym so it compounds over time and I have some incredible gains along the next few years with 0 injuries."* Multi-year horizon, zero injuries. **This governs. Do not prescribe aggressive bumps.**

- **Bump gate:** ALL sets at the TOP of the rep range **AND** every set RPE ≤7 → then the **smallest available increment** (+2.5 kg compound, +1 kg/DB isolation). A set creeping to RPE 8 = not yet; hold the load and clean it.
- **Reps before load.** One variable at a time. Rep-range progression preferred before load bumps on isolation.
- **Never re-open at a load that produced RPE 9+** (e.g. 35 kg pulldown, 6/10 — restart below it).
- **No aggressive mid-session bumps.** Autoregulate up only if RPE is clearly under target.
- **Off a layoff: seed conservative, autoregulate up. Loads landing under prescription = EXPECTED, not a flag.**
- Reps stall 3 sessions same weight → vary rep range OR drop 5% and rebuild.

> **DORMANT — aggressive double progression** (superseded 2026-06-10, do not use unless the user explicitly asks to push): top of rep range across all sets → +2.5 kg compound / +1–2 kg isolation next session, no waiting; first set RPE <7 → bump mid-session.

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
5. Kcal vs target — **slow cut: deliberate deficit ~−300/d, calorie-cycled** (see Nutrition baselines)
6. **Steps + shift hours — LOAD, not NEAT (promoted to Tier 1, 2026-07-27).** The waiter job is ~44 km/wk walked ≈ 2.75× the running. Gates: day >22 k steps → next run easy+short · two consecutive 25 k+ days → drop the week's shortest run, keep the long. Log every PM.

**Tier 2a — AM daily:**
7. BW AM (fasted, post-pee, conditions held constant)
8. Sleep h + quality
9. Resting HR (Δ vs 7-day baseline)
10. Motivation / energy / stress 1–10
11. Per-muscle DOMS 0–3
12. Joint/tendon flags

**Tier 2b — PM daily:**
13. Day macros total
14. Workout result (lifting days)

**Tier 3 — weekly/biweekly:**
15. Pre-WO meal timing + composition
16. Creatine 5 g/d compliance
17. Hydration L, caffeine, alcohol
18. Body circumferences (biweekly Mon)
19. Photos front/side/back (biweekly)
20. HR drift on easy runs

### Nutrition baselines — SLOW CUT ✅ LIVE AGAIN 2026-07-17 (⏸️ not started until running rhythm holds)

> **Un-paused 2026-07-17.** The 50 km build's "cut paused, BW gain expected" posture is DEAD. **The cut now serves the race goal** — 77.6 → ~73 kg ≈ 4–5 min over 30K (see top override). Deficit is intended; do not flag it as under-fueling.
> **⚠️ REVISED 2026-07-27 — later, smaller, cycled** (job load, see override): **starts W4 = Mon Aug 10, the down week** (was W3) · **deficit −300/d → ~2400**, not −400 · **calorie-cycled: maintenance ~2700 on shifts >18 k steps, deficit on low-step days** — same weekly deficit, far less recovery tax on the days that already cost the most. If the 7-day trend is flat 2 wks, then go −400.
> **Maintenance 2700 is EMPIRICAL, not estimated** — BW held 77.2–77.9 through W1 on 2 500–2 900 kcal *with* 17–23 k step shifts. The job is already inside the number; no upward correction owed.
> **Two timing caveats:** (1) **W1–W3 (Jul 20 – Aug 9) eat maintenance (~2700)** — no deficit while re-entering with detrained bone; cut opens W4. (2) **Final 3–4 wks + taper = maintenance** (W9 Mon Sep 14 → race), not deficit — no cutting while sharpening for Oct 11.

Goal (set 2026-05-17, revived 2026-07-17): deliberate slow fat loss to ~73 kg while retaining strength. NOT recomp, NOT surplus.

| Macro | Target |
|-------|--------|
| Kcal | maintenance (~2700) **minus ~300 → ~2400/d, calorie-cycled** (maint on >18 k-step shifts, deficit on low-step days); recalc maint per BW 7-day trend. Target loss **0.3–0.5 kg/wk** |
| Protein | **150–185 g (≥2.0 g/kg, push high — #1 muscle-retention lever in deficit)**, 4–5 meals, ≤5 h between feeds. 180+ g = good, never flag as "over". |
| Carbs | 3–5 g/kg lift days (230–390 g), 2–3 g/kg rest (155–230 g) |
| Fat | 0.8–1.0 g/kg (62–78 g) |
| Fiber | 25–35 g |
| Creatine | 5 g/d, every day |
| Pre-WO | carbs 1–2 g/kg + 30–40 g protein, 1–3 h prior |
| Post-WO | 30–40 g protein within 2 h |

### Adjustment triggers

- Sleep <6 h **and** RHR +10 → cut session volume 20%, RPE cap 7.
- **Step load (2026-07-27):** day >22 k steps → next day's run capped **easy + short** · two consecutive 25 k+ step days → **drop the week's shortest run, keep the long** · **run BEFORE the shift** wherever possible (running on legs that already stood 8 h is what 7/25 was).
- **"Legs don't feel good" 2 days running WITH shin clean = fatigue, not bone.** Take the rest, then resume at the **same** week's volume (repeat, don't compress). Do not escalate to an injury scare; do not claw back the missed km.
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
- **⚡ Every Sunday, run the PERFORMANCE EXTRACTION AUDIT (top section of `templates/weekly.md`) BEFORE anything else — mandatory, unprompted (user directive 2026-07-20).** It is the upside mirror of the red-flag block: the red flags catch over-reaching, the audit catches **under-dosing** — the plan lagging a fast-adapting engine (7/19: modeled ~7:00/km, ran 6:31; RHR held build-best through a layoff; Madrid's engine was bigger than a conservative read assumed → performance un-cashed). Compare model vs reality on paces, RPE, checkpoints; ratchet up when the engine is ahead; if the target looks soft 2 wks running, put the sub-2:45 stretch on the table and say so. **Guardrail: extraction ≠ risk — the tibia gate and "consistency > hero session" win every tie; the levers are cashing latent aerobic fitness and killing stale caution, never hero sessions on the bone.**

---

## Block 2 targets — ACTIVE (set 2026-07-17)

**Race: 30K, Sun Oct 11, 2026.**

| Target | Time | Pace | Notes |
|--------|------|------|-------|
| **Primary** | **sub-2:50** | **5:40/km** | ≈ a 1:57 HM. User's pick 2026-07-17. |
| Floor | sub-3:00 | 6:00/km | Still faster pace than Madrid (6:09/km) over +9 km. A strong day. |
| ~~Retired~~ | ~~sub-2:20~~ | ~~4:40/km~~ | Needed a 1:36 HM / 20:57 5K. Naive addition, not pace math. |

**Supporting targets:** BW 77.6 → **~73 kg** (0.3–0.5 kg/wk slow cut through base wks; maintenance final 3–4 wks + taper) — worth ~4–5 min over 30K, a third of the gap | **strength retained** at 2×/wk maintenance, slow-steady bump gate, no PR chasing | **cadence 165+** finally made habitual (unsolved since the ultra build) | **zero dark weeks** — volume consistency is the whole target.

**Key baselines (do not re-derive, do not invent):**
- HM PR: **2:09:38** Madrid 2026-04-26 (6:09/km, 339 m climb, avg HR 180, 78% Z5 = all-out). Flat-equivalent ≈ 2:05.
- 5K PR: **25:58** 2026-05-08 (5:12/km, all-out, max HR 204, RPE 9).
- Longest run: **18 km** 2026-06-22 (7:03/km @ avg HR 149, 95% Z2, RPE 4).
- Easy Z2 reference: **7:03/km @ HR 149** (18k, 6/22) · **7:22/km @ HR 150** (10k, 6/18).

> **Block 1 (11-Week Hypertrophy, Apr 27 – Jul 12) — CLOSED.** Never formally wrapped: the 50 km ultra detour (Jun 4) ate weeks 6–11, then the ultra was abandoned 2026-07-17. No end-of-block measurements or lift retest were taken. Hypertrophy sections below are reference for a future strength block. Old Block-1 target text, for the record: BW ~77 → ~73 kg, retain/build strength through the deficit, preserve lean mass, running held 90–95%.
