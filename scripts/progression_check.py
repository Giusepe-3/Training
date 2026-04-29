#!/usr/bin/env python3
"""
progression_check.py — analyze the training log and report progress.

Usage:
    python scripts/progression_check.py
    python scripts/progression_check.py --week 4         # focus on a single week
    python scripts/progression_check.py --exercise "Bench Press"  # focus on one lift

Reads:  logs/training_log.xlsx
Writes: nothing — prints a clean terminal report.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    print("This script requires pandas. Install with: pip install pandas openpyxl --break-system-packages")
    sys.exit(1)

REPO_ROOT = Path(__file__).resolve().parent.parent
LOG_PATH = REPO_ROOT / "logs" / "training_log.xlsx"
BLOCK_START = datetime(2026, 4, 27)
BLOCK_WEEKS = 11


def epley_1rm(weight: float, reps: float) -> float:
    """Epley estimated 1-rep max."""
    if pd.isna(weight) or pd.isna(reps) or reps <= 0:
        return float("nan")
    return round(weight * (1 + reps / 30), 1)


def load_log() -> dict[str, pd.DataFrame]:
    if not LOG_PATH.exists():
        print(f"Log file not found: {LOG_PATH}")
        sys.exit(1)
    sheets = pd.read_excel(LOG_PATH, sheet_name=["Lifting Log", "Running Log", "Body Metrics"])
    for name, df in sheets.items():
        df = df.dropna(how="all")
        if "Date" in df.columns:
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            df = df.dropna(subset=["Date"])
        # Recompute Week from Date directly — formulas may not be cached
        if "Date" in df.columns and len(df):
            df["Week"] = ((df["Date"] - BLOCK_START).dt.days // 7 + 1).clip(lower=1, upper=BLOCK_WEEKS).astype(int)
        # Recompute Volume for lifting log
        if name == "Lifting Log" and "Weight (kg)" in df.columns and "Reps" in df.columns:
            df["Volume (kg)"] = df["Weight (kg)"] * df["Reps"]
        # Recompute Pace for running log
        if name == "Running Log" and "Distance (km)" in df.columns and "Time (min)" in df.columns:
            df["Pace (min/km)"] = df["Time (min)"] / df["Distance (km)"]
        sheets[name] = df
    return sheets


def fmt_section(title: str) -> None:
    bar = "-" * len(title)
    print(f"\n{title}\n{bar}")


def report_overview(lifts: pd.DataFrame, runs: pd.DataFrame, body: pd.DataFrame) -> None:
    fmt_section("Block Overview")
    today = datetime.now().date()
    days_into = (today - BLOCK_START.date()).days
    if days_into < 0:
        print(f"Block has not started yet. Block starts {BLOCK_START.date()}.")
        return
    week = max(1, min(BLOCK_WEEKS, days_into // 7 + 1))
    print(f"Today: {today}")
    print(f"Block week: {week} of {BLOCK_WEEKS}")
    print(f"Days into block: {days_into}")
    print(f"Days remaining: {max(0, BLOCK_WEEKS * 7 - days_into)}")

    print(f"\nLifting sets logged: {len(lifts)}")
    print(f"Lifting sessions logged: {lifts.groupby('Date').size().shape[0] if len(lifts) else 0}")
    print(f"Runs logged: {len(runs)}")
    if len(runs):
        print(f"Total run distance: {runs['Distance (km)'].sum():.1f} km")
    print(f"Body weight entries: {len(body)}")


def report_weekly_summary(lifts: pd.DataFrame, runs: pd.DataFrame, body: pd.DataFrame) -> None:
    fmt_section("Weekly Summary")
    print(f"{'Wk':<4}{'Lifts':<8}{'Sets':<7}{'Volume(kg)':<13}{'AvgRPE':<9}{'Runs':<6}{'Run km':<9}{'Avg HR':<8}{'BW(kg)':<8}")
    for week in range(1, BLOCK_WEEKS + 1):
        wlifts = lifts[lifts["Week"] == week] if "Week" in lifts.columns else pd.DataFrame()
        wruns = runs[runs["Week"] == week] if "Week" in runs.columns else pd.DataFrame()
        wbody = body[body["Week"] == week] if "Week" in body.columns else pd.DataFrame()

        sessions = wlifts.groupby("Date").size().shape[0] if len(wlifts) else 0
        sets = len(wlifts)
        volume = wlifts["Volume (kg)"].sum() if "Volume (kg)" in wlifts.columns else 0
        avg_rpe = wlifts["RPE"].mean() if "RPE" in wlifts.columns and len(wlifts) else float("nan")
        n_runs = len(wruns)
        run_km = wruns["Distance (km)"].sum() if "Distance (km)" in wruns.columns else 0
        avg_hr = wruns["Avg HR"].mean() if "Avg HR" in wruns.columns and len(wruns) else float("nan")
        bw = wbody["Body Weight (kg)"].mean() if "Body Weight (kg)" in wbody.columns and len(wbody) else float("nan")

        rpe_str = f"{avg_rpe:.1f}" if not pd.isna(avg_rpe) else "-"
        hr_str = f"{avg_hr:.0f}" if not pd.isna(avg_hr) else "-"
        bw_str = f"{bw:.1f}" if not pd.isna(bw) else "-"
        print(f"{week:<4}{sessions:<8}{sets:<7}{volume:<13.0f}{rpe_str:<9}{n_runs:<6}{run_km:<9.1f}{hr_str:<8}{bw_str:<8}")


def report_personal_records(lifts: pd.DataFrame) -> None:
    fmt_section("Personal Records (top weight per exercise + e1RM)")
    if not len(lifts):
        print("No lifts logged yet.")
        return
    pr_rows = []
    for exercise, grp in lifts.groupby("Exercise"):
        max_w = grp["Weight (kg)"].max()
        best_set = grp[grp["Weight (kg)"] == max_w].iloc[-1]
        e1rm_all = grp.apply(lambda r: epley_1rm(r["Weight (kg)"], r["Reps"]), axis=1)
        best_e1rm = e1rm_all.max()
        pr_rows.append({
            "Exercise": exercise,
            "Best Weight": max_w,
            "Reps @ Best": int(best_set["Reps"]) if not pd.isna(best_set["Reps"]) else "-",
            "Best Date": best_set["Date"].date() if pd.notna(best_set["Date"]) else "-",
            "Best e1RM": best_e1rm,
        })
    pr_df = pd.DataFrame(pr_rows).sort_values("Best e1RM", ascending=False)
    for _, r in pr_df.iterrows():
        print(f"  {r['Exercise']:<28} {r['Best Weight']:>6.1f} kg × {str(r['Reps @ Best']):<3}  e1RM {r['Best e1RM']:>6.1f}  ({r['Best Date']})")


def report_progression(lifts: pd.DataFrame, exercise: str | None = None) -> None:
    fmt_section("Progression by Exercise (e1RM trend)")
    if not len(lifts):
        print("No lifts logged yet.")
        return
    df = lifts.copy()
    df["e1RM"] = df.apply(lambda r: epley_1rm(r["Weight (kg)"], r["Reps"]), axis=1)
    targets = [exercise] if exercise else df["Exercise"].dropna().unique()
    for ex in sorted(targets):
        sub = df[df["Exercise"] == ex].dropna(subset=["e1RM"])
        if not len(sub):
            continue
        weekly_best = sub.groupby("Week")["e1RM"].max()
        if len(weekly_best) < 2:
            change = ""
        else:
            first = weekly_best.iloc[0]
            last = weekly_best.iloc[-1]
            delta = last - first
            pct = (delta / first * 100) if first else 0
            change = f"  delta {delta:+.1f} kg ({pct:+.1f}%)"
        weeks_str = " -> ".join(f"W{int(w)}: {v:.1f}" for w, v in weekly_best.items())
        print(f"  {ex}")
        print(f"    {weeks_str}{change}")


def report_running_trend(runs: pd.DataFrame) -> None:
    fmt_section("Running Trend")
    if not len(runs):
        print("No runs logged yet.")
        return
    df = runs.dropna(subset=["Distance (km)", "Avg HR"]) if "Avg HR" in runs.columns else runs.copy()
    if len(df) < 2:
        print("Need at least 2 runs to assess trend.")
        return
    df = df.sort_values("Date").reset_index(drop=True)
    print(f"  Total runs: {len(df)}")
    print(f"  Total distance: {df['Distance (km)'].sum():.1f} km")
    if "Pace (min/km)" in df.columns:
        avg_pace = df["Pace (min/km)"].mean()
        print(f"  Avg pace: {avg_pace:.2f} min/km")
    if "Avg HR" in df.columns and df["Avg HR"].notna().any():
        first_hr = df["Avg HR"].dropna().iloc[0]
        last_hr = df["Avg HR"].dropna().iloc[-1]
        print(f"  HR trend: first run {first_hr:.0f} → last run {last_hr:.0f} ({last_hr - first_hr:+.0f})")
        if last_hr < first_hr:
            print("  ↳ HR dropping at similar effort — aerobic system improving.")


def report_body(body: pd.DataFrame) -> None:
    fmt_section("Body Metrics Trend")
    if not len(body):
        print("No body metrics logged yet.")
        return
    bw = body.dropna(subset=["Body Weight (kg)"]).sort_values("Date") if "Body Weight (kg)" in body.columns else pd.DataFrame()
    if len(bw) >= 2:
        first = bw["Body Weight (kg)"].iloc[0]
        last = bw["Body Weight (kg)"].iloc[-1]
        print(f"  Body weight: {first:.1f} kg → {last:.1f} kg ({last - first:+.1f} kg)")
    elif len(bw) == 1:
        print(f"  Body weight: {bw['Body Weight (kg)'].iloc[0]:.1f} kg (only 1 entry)")

    for metric in ["Waist (cm)", "Chest (cm)", "Arm (cm)", "Thigh (cm)"]:
        if metric not in body.columns:
            continue
        m = body.dropna(subset=[metric]).sort_values("Date")
        if len(m) >= 2:
            first = m[metric].iloc[0]
            last = m[metric].iloc[-1]
            print(f"  {metric}: {first:.1f} → {last:.1f} ({last - first:+.1f})")


def main() -> None:
    parser = argparse.ArgumentParser(description="Analyze training log progression.")
    parser.add_argument("--week", type=int, help="Focus on a single week")
    parser.add_argument("--exercise", type=str, help="Focus on a single exercise")
    args = parser.parse_args()

    sheets = load_log()
    lifts = sheets["Lifting Log"]
    runs = sheets["Running Log"]
    body = sheets["Body Metrics"]

    if args.week:
        lifts = lifts[lifts["Week"] == args.week]
        runs = runs[runs["Week"] == args.week]
        body = body[body["Week"] == args.week]
        print(f"\n[ Filtered to Week {args.week} ]")

    print("=" * 70)
    print("  Training Progression Report")
    print("=" * 70)

    report_overview(lifts, runs, body)
    if not args.week:
        report_weekly_summary(lifts, runs, body)
    report_personal_records(lifts)
    report_progression(lifts, exercise=args.exercise)
    report_running_trend(runs)
    report_body(body)
    print()


if __name__ == "__main__":
    main()
