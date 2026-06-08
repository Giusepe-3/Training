"""Generate training_log.xlsx — multi-tab logger for the 11-week block.

Pre-populates logged sessions (Wk 1+) and adds a Charts tab with progression
visualizations: e1RM trend per priority lift, weekly volume, weekly RPE, BW.

Run:  python3 scripts/build_log.py
"""
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.axis import DateAxis
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import Marker
from openpyxl.chart.trendline import Trendline
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

REPO_ROOT = Path(__file__).resolve().parent.parent
OUTPUT = REPO_ROOT / "logs" / "training_log.xlsx"

BLOCK_START = date(2026, 4, 27)

HEADER_FILL = PatternFill("solid", start_color="1F4E78")
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=12)
SECTION_FILL = PatternFill("solid", start_color="D9E1F2")
BODY_FONT = Font(name="Calibri", size=11)
ZEBRA_FILL = PatternFill("solid", start_color="F4F6FA")
THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
CENTER = Alignment(horizontal="center", vertical="center")
LEFT_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)


# ============== LOGGED SESSIONS DATA ==============
# Each set: (date, day_label, exercise, set_no, reps, weight_kg, rpe, notes)
# Sourced from logs/session_log.md bullets + logs/sessions/ filled logs.

LIFT_SETS = [
    # 2026-05-04 Mon Wk 2 Upper A
    (date(2026, 5, 4), "Upper A", "Bench Press",       1, 7, 32.5, 7, ""),
    (date(2026, 5, 4), "Upper A", "Bench Press",       2, 7, 32.5, 8, ""),
    (date(2026, 5, 4), "Upper A", "Bench Press",       3, 7, 32.5, 8, ""),
    (date(2026, 5, 4), "Upper A", "Bench Press",       4, 7, 32.5, 8, "bump 35 next"),
    (date(2026, 5, 4), "Upper A", "Barbell Row",       1, 7, 32.5, 8, ""),
    (date(2026, 5, 4), "Upper A", "Barbell Row",       2, 7, 32.5, 8, ""),
    (date(2026, 5, 4), "Upper A", "Barbell Row",       3, 7, 32.5, 8, ""),
    (date(2026, 5, 4), "Upper A", "Barbell Row",       4, 7, 32.5, 8, "bump 35 next"),
    (date(2026, 5, 4), "Upper A", "Incline DB Press",  1, 10, 12, 8, ""),
    (date(2026, 5, 4), "Upper A", "Incline DB Press",  2, 9,  12, 8, ""),
    (date(2026, 5, 4), "Upper A", "Incline DB Press",  3, 8,  12, 8, "hold 12, climb 10/10/10"),
    (date(2026, 5, 4), "Upper A", "Lat Pulldown",      1, 10, 35, 7, ""),
    (date(2026, 5, 4), "Upper A", "Lat Pulldown",      2, 10, 35, 8, ""),
    (date(2026, 5, 4), "Upper A", "Lat Pulldown",      3, 10, 35, 8, "bump 37.5 next"),
    (date(2026, 5, 4), "Upper A", "Pec Deck",          1, 12, 7.5, 9, ""),
    (date(2026, 5, 4), "Upper A", "Pec Deck",          2, 12, 7.5, 9, ""),
    (date(2026, 5, 4), "Upper A", "Pec Deck",          3, 12, 7.5, 9, "hold 7.5, climb 15/15/15"),
    (date(2026, 5, 4), "Upper A", "Face Pull",         1, 15, 10,  8, ""),
    (date(2026, 5, 4), "Upper A", "Face Pull",         2, 14, 10,  9, ""),
    (date(2026, 5, 4), "Upper A", "Face Pull",         3, 14, 10,  8, "bump 11-12 next"),
    (date(2026, 5, 4), "Upper A", "Triceps Pushdown",  1, 12, 10,  8, ""),
    (date(2026, 5, 4), "Upper A", "Triceps Pushdown",  2, 12, 10,  8, ""),
    (date(2026, 5, 4), "Upper A", "Triceps Pushdown",  3, 12, 10,  8, "bump 11-12 next"),
    (date(2026, 5, 4), "Upper A", "DB Curl",           1, 12, 6,   8, ""),
    (date(2026, 5, 4), "Upper A", "DB Curl",           2, 12, 6,   8, ""),
    (date(2026, 5, 4), "Upper A", "DB Curl",           3, 12, 6,   8, "bump 7-8 next"),

    # 2026-05-05 Tue Wk 2 Lower A
    (date(2026, 5, 5), "Lower A", "Back Squat",        1, 7, 35, 8, ""),
    (date(2026, 5, 5), "Lower A", "Back Squat",        2, 7, 35, 8, ""),
    (date(2026, 5, 5), "Lower A", "Back Squat",        3, 6, 35, 8, ""),
    (date(2026, 5, 5), "Lower A", "Back Squat",        4, 6, 35, 8, "hold 35, climb 7/7/7/7"),
    (date(2026, 5, 5), "Lower A", "Romanian Deadlift", 1, 10, 30, 8, ""),
    (date(2026, 5, 5), "Lower A", "Romanian Deadlift", 2, 10, 30, 8, ""),
    (date(2026, 5, 5), "Lower A", "Romanian Deadlift", 3, 8,  30, 8, "hold 30, climb 10/10/10"),
    (date(2026, 5, 5), "Lower A", "Leg Press",         1, 12, 20, 7, ""),
    (date(2026, 5, 5), "Lower A", "Leg Press",         2, 12, 20, 8, ""),
    (date(2026, 5, 5), "Lower A", "Leg Press",         3, 11, 20, 8, "bump 22.5-25 next"),
    (date(2026, 5, 5), "Lower A", "Lying Leg Curl",    1, 12, 15, 6, "warm-up, bumped"),
    (date(2026, 5, 5), "Lower A", "Lying Leg Curl",    2, 11, 20, 8, ""),
    (date(2026, 5, 5), "Lower A", "Lying Leg Curl",    3, 10, 20, 8, "hold 20, climb 12/12/12"),
    (date(2026, 5, 5), "Lower A", "Standing Calf Raise", 1, 15, 20, 6, "warm-up, bumped"),
    (date(2026, 5, 5), "Lower A", "Standing Calf Raise", 2, 14, 30, 9, ""),
    (date(2026, 5, 5), "Lower A", "Standing Calf Raise", 3, 13, 30, 8, ""),
    (date(2026, 5, 5), "Lower A", "Standing Calf Raise", 4, 12, 30, 9, "hold 30, climb 15/15/15"),
    (date(2026, 5, 5), "Lower A", "Hanging Knee Raise",  1, 10, 0,  7, "BW"),
    (date(2026, 5, 5), "Lower A", "Hanging Knee Raise",  2, 10, 0,  7, "BW"),
    (date(2026, 5, 5), "Lower A", "Hanging Knee Raise",  3, 10, 0,  7, "BW, add reps/weight next"),

    # 2026-05-07 Thu Wk 2 Upper B
    (date(2026, 5, 7), "Upper B", "Overhead Press",    1, 8, 22.5, 8, ""),
    (date(2026, 5, 7), "Upper B", "Overhead Press",    2, 8, 22.5, 8, ""),
    (date(2026, 5, 7), "Upper B", "Overhead Press",    3, 8, 22.5, 9, ""),
    (date(2026, 5, 7), "Upper B", "Overhead Press",    4, 8, 22.5, 9, "bump 25 next, +1 set Wk 3"),
    (date(2026, 5, 7), "Upper B", "Lat Pulldown",      1, 10, 40, 7, ""),
    (date(2026, 5, 7), "Upper B", "Lat Pulldown",      2, 10, 40, 7, ""),
    (date(2026, 5, 7), "Upper B", "Lat Pulldown",      3, 10, 40, 8, ""),
    (date(2026, 5, 7), "Upper B", "Lat Pulldown",      4, 10, 40, 9, "bump 42.5 next, +1 set Wk 3"),
    (date(2026, 5, 7), "Upper B", "Incline DB Press",  1, 10, 10,  8, ""),
    (date(2026, 5, 7), "Upper B", "Incline DB Press",  2, 10, 10,  8, ""),
    (date(2026, 5, 7), "Upper B", "Incline DB Press",  3, 10, 10,  8, "bump 12 next"),
    (date(2026, 5, 7), "Upper B", "Cable Row",         1, 12, 25, 5, "way under-loaded"),
    (date(2026, 5, 7), "Upper B", "Cable Row",         2, 12, 35, 7, ""),
    (date(2026, 5, 7), "Upper B", "Cable Row",         3, 12, 40, 10, "settle 32.5 next"),
    (date(2026, 5, 7), "Upper B", "DB Lateral Raise",  1, 13, 4,   8, ""),
    (date(2026, 5, 7), "Upper B", "DB Lateral Raise",  2, 14, 4,   8, ""),
    (date(2026, 5, 7), "Upper B", "DB Lateral Raise",  3, 13, 4,   9, ""),
    (date(2026, 5, 7), "Upper B", "DB Lateral Raise",  4, 14, 4,   10, "bump 5 next"),
    (date(2026, 5, 7), "Upper B", "Face Pull",         1, 20, 10,  7, ""),
    (date(2026, 5, 7), "Upper B", "Face Pull",         2, 20, 10,  8, ""),
    (date(2026, 5, 7), "Upper B", "Face Pull",         3, 20, 10,  9, "bump 12 next"),
    (date(2026, 5, 7), "Upper B", "DB Curl",           1, 10, 6,   7, ""),
    (date(2026, 5, 7), "Upper B", "DB Curl",           2, 10, 7,   8, ""),
    (date(2026, 5, 7), "Upper B", "DB Curl",           3, 10, 7,   9, "bump 8 next, climb 10/10/10"),
    (date(2026, 5, 7), "Upper B", "Skull Crusher",     1, 10, 13.6, 7, ""),
    (date(2026, 5, 7), "Upper B", "Skull Crusher",     2, 8,  13.6, 9, ""),
    (date(2026, 5, 7), "Upper B", "Skull Crusher",     3, 8,  13.6, 10, "hold 13.6, climb 10/10/10"),

    # 2026-05-10 Sun Wk 2 Lower B — first Lower B, data collection (RPE 6-7, 2 sets/lift)
    (date(2026, 5, 10), "Lower B", "Trap Bar Deadlift",     1, 6, 40, 6, "under-loaded"),
    (date(2026, 5, 10), "Lower B", "Trap Bar Deadlift",     2, 6, 40, 6, "bump 45 next"),
    (date(2026, 5, 10), "Lower B", "Bulgarian Split Squat", 1, 16, 0, 5, "BW; 8L+8R; pattern"),
    (date(2026, 5, 10), "Lower B", "Bulgarian Split Squat", 2, 24, 0, 6, "BW; 12L+12R; S2 overshot; 5/DB next"),
    (date(2026, 5, 10), "Lower B", "Hip Thrust",            1, 8, 10, 8, "logged 10 vs prescribed 40 — plate check"),
    (date(2026, 5, 10), "Lower B", "Hip Thrust",            2, 8, 10, 9, "bump 12.5 next (verify plates)"),
    (date(2026, 5, 10), "Lower B", "Walking Lunge",         1, 20, 10, 8, "5/DB=10kg total; 10L+10R"),
    (date(2026, 5, 10), "Lower B", "Walking Lunge",         2, 20, 10, 7, "bump 7.5/DB next"),
    (date(2026, 5, 10), "Lower B", "Seated Calf Raise",     1, 12, 9, 8, "soleus bias"),
    (date(2026, 5, 10), "Lower B", "Seated Calf Raise",     2, 12, 9, 9, "RPE overshoot; bump 11 next"),
    (date(2026, 5, 10), "Lower B", "Ab Roller",             1, 8, 0, 8, "BW knee rollout"),
    (date(2026, 5, 10), "Lower B", "Ab Roller",             2, 8, 0, 8, "BW; progress reps before load"),

    # 2026-05-11 Mon Wk 3 Upper A
    (date(2026, 5, 11), "Upper A", "Bench Press",      1, 8, 35, 8, ""),
    (date(2026, 5, 11), "Upper A", "Bench Press",      2, 8, 35, 8, ""),
    (date(2026, 5, 11), "Upper A", "Bench Press",      3, 8, 35, 9, ""),
    (date(2026, 5, 11), "Upper A", "Bench Press",      4, 8, 35, 9, ""),
    (date(2026, 5, 11), "Upper A", "Bench Press",      5, 8, 35, 9, "top×5, bump 37.5 next"),
    (date(2026, 5, 11), "Upper A", "Barbell Row",      1, 8, 35, 8, ""),
    (date(2026, 5, 11), "Upper A", "Barbell Row",      2, 8, 35, 8, ""),
    (date(2026, 5, 11), "Upper A", "Barbell Row",      3, 8, 35, 8, ""),
    (date(2026, 5, 11), "Upper A", "Barbell Row",      4, 8, 35, 8, ""),
    (date(2026, 5, 11), "Upper A", "Barbell Row",      5, 8, 35, 9, "top×5, bump 37.5 next"),
    (date(2026, 5, 11), "Upper A", "Incline DB Press", 1, 10, 12, 8, ""),
    (date(2026, 5, 11), "Upper A", "Incline DB Press", 2, 10, 12, 8, ""),
    (date(2026, 5, 11), "Upper A", "Incline DB Press", 3, 10, 12, 8, "climb hit, bump 14 next"),
    (date(2026, 5, 11), "Upper A", "Lat Pulldown",     1, 8, 37.5, 7, "S1 RPE7 → mid bump"),
    (date(2026, 5, 11), "Upper A", "Lat Pulldown",     2, 8, 40, 8, ""),
    (date(2026, 5, 11), "Upper A", "Lat Pulldown",     3, 8, 40, 8, ""),
    (date(2026, 5, 11), "Upper A", "Lat Pulldown",     4, 8, 40, 8, "bump 42.5 next"),
    (date(2026, 5, 11), "Upper A", "Pec Deck",         1, 12, 7.5, 9, ""),
    (date(2026, 5, 11), "Upper A", "Pec Deck",         2, 9, 7.5, 10, ""),
    (date(2026, 5, 11), "Upper A", "Pec Deck",         3, 10, 7.5, 10, "muscular failure; hold 7.5 climb"),
    (date(2026, 5, 11), "Upper A", "Face Pull",        1, 15, 12.5, 8, ""),
    (date(2026, 5, 11), "Upper A", "Face Pull",        2, 15, 12.5, 8, ""),
    (date(2026, 5, 11), "Upper A", "Face Pull",        3, 15, 12.5, 9, "top×3, bump 14 next"),
    (date(2026, 5, 11), "Upper A", "Triceps Pushdown", 1, 12, 12.5, 6, ""),
    (date(2026, 5, 11), "Upper A", "Triceps Pushdown", 2, 12, 15, 7, ""),
    (date(2026, 5, 11), "Upper A", "Triceps Pushdown", 3, 12, 17.5, 10, "messy ramp; start 15 next"),
    (date(2026, 5, 11), "Upper A", "DB Curl",          1, 12, 8, 8, ""),
    (date(2026, 5, 11), "Upper A", "DB Curl",          2, 12, 8, 9, ""),
    (date(2026, 5, 11), "Upper A", "DB Curl",          3, 11, 8, 10, "muscular failure; hold 8 climb"),

    # 2026-05-13 Wed Wk 3 Lower A (moved from Tue 5/12)
    (date(2026, 5, 13), "Lower A", "Back Squat",          1, 8, 35, 7, ""),
    (date(2026, 5, 13), "Lower A", "Back Squat",          2, 8, 35, 7.5, ""),
    (date(2026, 5, 13), "Lower A", "Back Squat",          3, 8, 35, 8, ""),
    (date(2026, 5, 13), "Lower A", "Back Squat",          4, 8, 35, 8, ""),
    (date(2026, 5, 13), "Lower A", "Back Squat",          5, 8, 35, 8, "top×5, bump 37.5 next"),
    (date(2026, 5, 13), "Lower A", "Romanian Deadlift",   1, 10, 30, 7, ""),
    (date(2026, 5, 13), "Lower A", "Romanian Deadlift",   2, 10, 30, 7, ""),
    (date(2026, 5, 13), "Lower A", "Romanian Deadlift",   3, 10, 32.5, 8, "S3 self-bump; 32.5 working next"),
    (date(2026, 5, 13), "Lower A", "Leg Press",           1, 12, 25, 7, ""),
    (date(2026, 5, 13), "Lower A", "Leg Press",           2, 12, 25, 7, ""),
    (date(2026, 5, 13), "Lower A", "Leg Press",           3, 12, 30, 7, "under-loaded; bump 30 next"),
    (date(2026, 5, 13), "Lower A", "Lying Leg Curl",      1, 12, 20, 7, ""),
    (date(2026, 5, 13), "Lower A", "Lying Leg Curl",      2, 12, 20, 7, ""),
    (date(2026, 5, 13), "Lower A", "Lying Leg Curl",      3, 12, 22.5, 9, "bump 22.5 next"),
    (date(2026, 5, 13), "Lower A", "Standing Calf Raise", 1, 15, 30, 8, ""),
    (date(2026, 5, 13), "Lower A", "Standing Calf Raise", 2, 15, 30, 9, ""),
    (date(2026, 5, 13), "Lower A", "Standing Calf Raise", 3, 15, 30, 10, ""),
    (date(2026, 5, 13), "Lower A", "Standing Calf Raise", 4, 12, 30, 10, "missed S4 reps; hold 30"),
    (date(2026, 5, 13), "Lower A", "Hanging Knee Raise",  1, 8, 0, 8, "BW; ab DOMS limiter"),
    (date(2026, 5, 13), "Lower A", "Hanging Knee Raise",  2, 8, 0, 9, "BW"),
    (date(2026, 5, 13), "Lower A", "Hanging Knee Raise",  3, 8, 0, 9, "BW; hold, target 10/10/10"),

    # 2026-05-16 Sat Wk 3 Upper B (moved from Fri 5/15 SKIPPED)
    (date(2026, 5, 16), "Upper B", "Overhead Press",   1, 8, 25, 9, ""),
    (date(2026, 5, 16), "Upper B", "Overhead Press",   2, 8, 25, 9, ""),
    (date(2026, 5, 16), "Upper B", "Overhead Press",   3, 8, 25, 10, ""),
    (date(2026, 5, 16), "Upper B", "Overhead Press",   4, 7, 25, 9, ""),
    (date(2026, 5, 16), "Upper B", "Overhead Press",   5, 7, 25, 10, "hold 25, bank 8×5"),
    (date(2026, 5, 16), "Upper B", "Lat Pulldown",     1, 8, 42.5, 7, ""),
    (date(2026, 5, 16), "Upper B", "Lat Pulldown",     2, 8, 42.5, 8, ""),
    (date(2026, 5, 16), "Upper B", "Lat Pulldown",     3, 8, 42.5, 8, ""),
    (date(2026, 5, 16), "Upper B", "Lat Pulldown",     4, 8, 42.5, 9, ""),
    (date(2026, 5, 16), "Upper B", "Lat Pulldown",     5, 7, 42.5, 10, "hold 42.5, close S5"),
    (date(2026, 5, 16), "Upper B", "Incline DB Press", 1, 10, 12, 7, ""),
    (date(2026, 5, 16), "Upper B", "Incline DB Press", 2, 10, 12, 7, ""),
    (date(2026, 5, 16), "Upper B", "Incline DB Press", 3, 10, 12, 9, "top×3, bump 14 next"),
    (date(2026, 5, 16), "Upper B", "Cable Row",        1, 12, 30, 7, "ran 30 vs prescribed 32.5"),
    (date(2026, 5, 16), "Upper B", "Cable Row",        2, 12, 30, 8, ""),
    (date(2026, 5, 16), "Upper B", "Cable Row",        3, 12, 30, 10, "bump 32.5 next"),
    (date(2026, 5, 16), "Upper B", "DB Lateral Raise", 1, 15, 5, 7, ""),
    (date(2026, 5, 16), "Upper B", "DB Lateral Raise", 2, 14, 5, 9, ""),
    (date(2026, 5, 16), "Upper B", "DB Lateral Raise", 3, 13, 5, 9, ""),
    (date(2026, 5, 16), "Upper B", "DB Lateral Raise", 4, 8, 5, 10, "S4 failure; hold 5"),
    (date(2026, 5, 16), "Upper B", "Face Pull",        1, 15, 12.5, 8, ""),
    (date(2026, 5, 16), "Upper B", "Face Pull",        2, 15, 12.5, 10, ""),
    (date(2026, 5, 16), "Upper B", "Face Pull",        3, 15, 12.5, 10, "hold 12.5, climb reps→20"),
    (date(2026, 5, 16), "Upper B", "DB Curl",          1, 10, 8, 9, ""),
    (date(2026, 5, 16), "Upper B", "DB Curl",          2, 10, 8, 8, ""),
    (date(2026, 5, 16), "Upper B", "DB Curl",          3, 10, 8, 8, "top×3, bump 10 next"),
    (date(2026, 5, 16), "Upper B", "Skull Crusher",    1, 10, 13.6, 7, ""),
    (date(2026, 5, 16), "Upper B", "Skull Crusher",    2, 10, 13.6, 8, ""),
    (date(2026, 5, 16), "Upper B", "Skull Crusher",    3, 10, 13.6, 9, "climb hit, bump 15.9 next"),

    # 2026-05-17 Sun Wk 3 Lower B (moved from Sat) — RPE 7 baseline, +5 Trap DL
    (date(2026, 5, 17), "Lower B", "Trap Bar Deadlift",     1, 6, 45, 8, "RIR2; under-loaded"),
    (date(2026, 5, 17), "Lower B", "Trap Bar Deadlift",     2, 6, 45, 7, "RIR3; bump 50 next; e1RM 54.0"),
    (date(2026, 5, 17), "Lower B", "Bulgarian Split Squat", 1, 20, 10, 8, "5/DB=10kg total; 10L+10R; RIR2"),
    (date(2026, 5, 17), "Lower B", "Bulgarian Split Squat", 2, 20, 10, 8, "10L+10R; top×all; bump 7.5/DB next"),
    (date(2026, 5, 17), "Lower B", "Hip Thrust",            1, 8, 12.5, 10, "MACHINE load scale; RIR0"),
    (date(2026, 5, 17), "Lower B", "Hip Thrust",            2, 8, 12.5, 10, "MACHINE; correctly loaded; hold 12.5"),
    (date(2026, 5, 17), "Lower B", "Walking Lunge",         1, 20, 14, 7, "7/DB=14kg total; 10L+10R"),
    (date(2026, 5, 17), "Lower B", "Walking Lunge",         2, 20, 14, 7, "top×all; bump 9/DB next"),
    (date(2026, 5, 17), "Lower B", "Seated Calf Raise",     1, 12, 30, 6, "soleus bias"),
    (date(2026, 5, 17), "Lower B", "Seated Calf Raise",     2, 12, 30, 6, "RPE6 under-loaded; bump 35 next"),
    (date(2026, 5, 17), "Lower B", "Ab Roller",             1, 10, 0, 10, "BW knee rollout"),
    (date(2026, 5, 17), "Lower B", "Ab Roller",             2, 10, 0, 10, "BW; hold, target 2×12"),

    # 2026-05-18 Mon Wk 4 Upper A — RPE 8 baseline; e1RM bench 44.3→47.5
    (date(2026, 5, 18), "Upper A", "Bench Press",      1, 8, 37.5, 8, ""),
    (date(2026, 5, 18), "Upper A", "Bench Press",      2, 8, 37.5, 8, ""),
    (date(2026, 5, 18), "Upper A", "Bench Press",      3, 8, 37.5, 9, ""),
    (date(2026, 5, 18), "Upper A", "Bench Press",      4, 8, 37.5, 10, "S4 failure RIR0"),
    (date(2026, 5, 18), "Upper A", "Bench Press",      5, 6, 37.5, 9, "fell to 6r; hold 37.5, clean 5×8 next; e1RM 47.5"),
    (date(2026, 5, 18), "Upper A", "Barbell Row",      1, 8, 37.5, 8, ""),
    (date(2026, 5, 18), "Upper A", "Barbell Row",      2, 8, 37.5, 8, ""),
    (date(2026, 5, 18), "Upper A", "Barbell Row",      3, 8, 37.5, 8, ""),
    (date(2026, 5, 18), "Upper A", "Barbell Row",      4, 8, 37.5, 9, ""),
    (date(2026, 5, 18), "Upper A", "Barbell Row",      5, 7, 37.5, 9, "S5 7r; hold 37.5, clean 5×8 next"),
    (date(2026, 5, 18), "Upper A", "Incline DB Press", 1, 10, 15, 8, "ran 15 vs start 14"),
    (date(2026, 5, 18), "Upper A", "Incline DB Press", 2, 10, 15, 8, ""),
    (date(2026, 5, 18), "Upper A", "Incline DB Press", 3, 10, 15, 9, "top×3, bump 16 next"),
    (date(2026, 5, 18), "Upper A", "Lat Pulldown",     1, 8, 42.5, 7, "RIR3 too light"),
    (date(2026, 5, 18), "Upper A", "Lat Pulldown",     2, 8, 42.5, 7, "RIR3"),
    (date(2026, 5, 18), "Upper A", "Lat Pulldown",     3, 8, 42.5, 7, "RIR3"),
    (date(2026, 5, 18), "Upper A", "Lat Pulldown",     4, 6, 45, 10, "self-bump S4; start 45 next"),
    (date(2026, 5, 18), "Upper A", "Pec Deck",         1, 15, 7.5, 8, ""),
    (date(2026, 5, 18), "Upper A", "Pec Deck",         2, 12, 7.5, 8, ""),
    (date(2026, 5, 18), "Upper A", "Pec Deck",         3, 12, 7.5, 10, "muscular failure; hold 7.5 chase 15×3"),
    (date(2026, 5, 18), "Upper A", "Face Pull",        1, 20, 15, 8, ""),
    (date(2026, 5, 18), "Upper A", "Face Pull",        2, 20, 15, 8, ""),
    (date(2026, 5, 18), "Upper A", "Face Pull",        3, 20, 15, 8, "top×3 RPE8, bump 16 next"),
    (date(2026, 5, 18), "Upper A", "Triceps Pushdown", 1, 12, 15, 8, ""),
    (date(2026, 5, 18), "Upper A", "Triceps Pushdown", 2, 12, 15, 8, ""),
    (date(2026, 5, 18), "Upper A", "Triceps Pushdown", 3, 12, 15, 8, "target clean×3, bump 17 next"),
    (date(2026, 5, 18), "Upper A", "DB Curl",          1, 12, 8, 7, ""),
    (date(2026, 5, 18), "Upper A", "DB Curl",          2, 12, 8, 8, ""),
    (date(2026, 5, 18), "Upper A", "DB Curl",          3, 10, 8, 8, "S3 10r; hold 8 chase 12×3"),
    # 2026-06-08 Mon Build W1 Push — first lift in 2+ wks (W0 dark); maintenance, autoregulate up
    (date(2026, 6, 8), "Push", "Standing Chest Press", 1, 15, 5,   5, "new locked sub vs bench (user pref + shoulder)"),
    (date(2026, 6, 8), "Push", "Standing Chest Press", 2, 15, 5,   6, ""),
    (date(2026, 6, 8), "Push", "Standing Chest Press", 3, 12, 5,   7, ""),
    (date(2026, 6, 8), "Push", "Standing Chest Press", 4, 10, 5,   6, "RPE5-6 trivial; bump 7.5-10 next"),
    (date(2026, 6, 8), "Push", "DB Shoulder Press",     1, 10, 3,  5, "shoulder gate CLEAN (weak, no pain)"),
    (date(2026, 6, 8), "Push", "DB Shoulder Press",     2, 10, 3,  5, ""),
    (date(2026, 6, 8), "Push", "DB Shoulder Press",     3, 10, 3,  5, "RPE5; bump 5/DB next, cap RPE7"),
    (date(2026, 6, 8), "Push", "Incline DB Press",      1, 10, 5,  7, ""),
    (date(2026, 6, 8), "Push", "Incline DB Press",      2, 10, 5,  7, ""),
    (date(2026, 6, 8), "Push", "Incline DB Press",      3, 10, 5,  7, "RPE7 top range; bump 6-7/DB next"),
    (date(2026, 6, 8), "Push", "Pec Deck",              1, 9,  2.5, 7, ""),
    (date(2026, 6, 8), "Push", "Pec Deck",              2, 6,  2.5, 7, ""),
    (date(2026, 6, 8), "Push", "Pec Deck",              3, 6,  2.5, 7, "reps fell 9->6->6; HOLD 2.5 rebuild 12+"),
    (date(2026, 6, 8), "Push", "Triceps Pushdown",      1, 15, 15, 8, ""),
    (date(2026, 6, 8), "Push", "Triceps Pushdown",      2, 12, 15, 8, ""),
    (date(2026, 6, 8), "Push", "Triceps Pushdown",      3, 12, 15, 9, "hit top @RPE8-9; bump 17.5 next"),
    (date(2026, 6, 8), "Push", "DB Lateral Raise",      1, 12, 2,  8, ""),
    (date(2026, 6, 8), "Push", "DB Lateral Raise",      2, 12, 2,  8, ""),
    (date(2026, 6, 8), "Push", "DB Lateral Raise",      3, 12, 2,  8, "bottom range; HOLD 2 chase 15s then 3/DB"),
]

# Body weight log: (date, BW kg, sleep h, notes)
BODY_LOG = [
    (date(2026, 4, 29), 77.7, 8.0, "Wk 1 Wed AM"),
    (date(2026, 5, 7),  77.3, 9.0, "Wk 2 Thu PM (post Upper B)"),
    (date(2026, 5, 8),  77.0, 8.0, "Wk 2 Fri PM (post 5K PR)"),
    (date(2026, 5, 9),  77.0, 8.0, "Wk 2 Sat PM (rest, alc 5u)"),
    (date(2026, 5, 10), 77.0, 7.0,  "Wk 2 Sun PM (post Lower B)"),
    (date(2026, 5, 11), 77.2, 7.83, "Wk 3 Mon PM (post Upper A)"),
    (date(2026, 5, 13), 77.2, 8.6,  "Wk 3 Wed PM (post Lower A)"),
    (date(2026, 5, 14), 77.3, 8.0,  "Wk 3 Thu PM (rest, protein 90g)"),
    # BW now AM-only from 2026-05-15 (user rule 2026-05-17): never log PM/evening BW going forward
    (date(2026, 5, 15), 77.7, 9.0,  "Wk 3 Fri AM (BW AM-only from here)"),
    (date(2026, 5, 16), 76.7, 7.5,  "Wk 3 Sat AM (-1kg vs Fri; slow cut, NOT illness)"),
    (date(2026, 5, 17), 76.2, 8.0,  "Wk 3 Sun AM (cut on-plan; AM slope toward ~73)"),
    (date(2026, 5, 18), 76.2, 8.0,  "Wk 4 Mon AM (flat vs Sun; post-Lower-B water noise, not stall)"),
    (date(2026, 6, 8),  77.1, 7.0,  "Build W1 Mon AM (cut PAUSED, maintenance — BW gain expected, NOT flagged)"),
]

# Running log: (date, dist_km, time_min, avg_hr, max_hr, rpe, type, notes)
RUN_LOG = [
    # 2026-05-08 Fri Wk 2 — planned Z2 25-30min, actual 5K PR @ Z5 65%
    (date(2026, 5, 8), 5.00, 25.97, 176, 204, 9, "Other", "5K PR 25:58 @5:12/km. Z5 65%, planned Z2 — full CNS dump"),
    # 2026-05-17 Sun Wk 3 — Sun Z2 easy (HR not recorded)
    (date(2026, 5, 17), 5.20, 35.0, 160, 0, 3, "Easy", "Sun Z2 chill 6:43/km; avg HR 160; max not recorded; aerobic base held"),
]


# ============== HELPERS ==============

def set_header_row(ws, row, headers, widths=None):
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=i, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = THIN_BORDER
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 26


def style_body_row(ws, row, n_cols, zebra=False, notes_col=None):
    """Apply body font, border, optional zebra fill, wrap on notes col."""
    fill = ZEBRA_FILL if zebra else None
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = BODY_FONT
        cell.border = THIN_BORDER
        if fill:
            cell.fill = fill
        if notes_col and c == notes_col:
            cell.alignment = LEFT_WRAP
        else:
            cell.alignment = CENTER


def add_dropdown(ws, col_letter, start_row, end_row, choices):
    dv = DataValidation(type="list", formula1=f'"{",".join(choices)}"', allow_blank=True)
    dv.add(f"{col_letter}{start_row}:{col_letter}{end_row}")
    ws.add_data_validation(dv)


def epley(weight, reps):
    if not weight or not reps:
        return None
    return round(weight * (1 + reps / 30), 1)


# ============== BUILD WORKBOOK ==============

wb = Workbook()


# -------- TAB: README --------
ws = wb.active
ws.title = "README"
ws.column_dimensions["A"].width = 110

readme = [
    ("Training Log — 11-Week Hypertrophy Block", "title"),
    ("Block: 2026-04-27 → 2026-07-12  •  User: Leonardo  •  Last rebuild driven by scripts/build_log.py", "section"),
    ("", ""),
    ("TABS", "section"),
    ("• Charts — KPI snapshot (live) + 6 progression charts. Open this first.", "body"),
    ("• Lifting Log — every working set. Edit here to add new sessions.", "body"),
    ("• Running Log — every run. Edit here to add runs.", "body"),
    ("• Body Metrics — daily BW + weekly tape measurements.", "body"),
    ("• Personal Records — auto best lifts (recomputes on open).", "body"),
    ("• Weekly Summary — auto weekly volume / RPE / runs / BW.", "body"),
    ("• Progression Data — weekly best e1RM per priority lift (chart source).", "body"),
    ("• Reference — HR / pace zones, schedule, RPE table.", "body"),
    ("", ""),
    ("HOW IT WORKS", "section"),
    ("• Sources of truth: logs/session_log.md (narrative) + logs/sessions/<date>_<sess>.md (filled session blocks).", "body"),
    ("• Pre-populated rows are baked into scripts/build_log.py — re-run only after adding new sessions to the script.", "body"),
    ("• Charts read from live formulas — they update automatically when you add data to the input tabs.", "body"),
    ("", ""),
    ("RPE LEGEND", "section"),
    ("RPE 6 = 4 RIR  •  7 = 3 RIR  •  8 = 2 RIR (most of block)  •  9 = 1 RIR (peak)  •  10 = failure (isolation only).", "body"),
]
for i, (line, kind) in enumerate(readme, 1):
    cell = ws.cell(row=i, column=1, value=line)
    if kind == "title":
        cell.font = Font(name="Calibri", bold=True, size=18, color="1F4E78")
        ws.row_dimensions[i].height = 28
    elif kind == "section":
        cell.font = Font(name="Calibri", bold=True, size=13, color="1F4E78")
    else:
        cell.font = Font(name="Calibri", size=11)


# -------- TAB: Lifting Log --------
ws = wb.create_sheet("Lifting Log")
headers = ["Date", "Week", "Day", "Exercise", "Set #", "Reps", "Weight (kg)", "RPE", "Volume (kg)", "e1RM (kg)", "Notes"]
widths = [12, 7, 10, 24, 7, 7, 12, 7, 13, 12, 42]
set_header_row(ws, 1, headers, widths)

# Pre-populate logged sets
for i, (d, day, ex, sn, reps, w, rpe, note) in enumerate(LIFT_SETS, start=2):
    ws.cell(row=i, column=1, value=d).number_format = "yyyy-mm-dd"
    ws.cell(row=i, column=3, value=day)
    ws.cell(row=i, column=4, value=ex)
    ws.cell(row=i, column=5, value=sn)
    ws.cell(row=i, column=6, value=reps)
    ws.cell(row=i, column=7, value=w).number_format = "0.0"
    ws.cell(row=i, column=8, value=rpe)
    ws.cell(row=i, column=11, value=note)

# Pre-fill formulas + style for many rows (logged + future capacity)
N_ROWS = 800
for r in range(2, 2 + N_ROWS):
    ws.cell(row=r, column=2, value=f'=IF(A{r}="","",IFERROR(INT((A{r}-DATE(2026,4,27))/7)+1,""))')
    ws.cell(row=r, column=9, value=f'=IF(OR(F{r}="",G{r}=""),"",F{r}*G{r})').number_format = "0"
    ws.cell(row=r, column=10, value=f'=IF(OR(F{r}="",G{r}=""),"",ROUND(G{r}*(1+F{r}/30),1))').number_format = "0.0"
    style_body_row(ws, r, 11, zebra=(r % 2 == 0), notes_col=11)
    ws.cell(row=r, column=1).number_format = "yyyy-mm-dd"
    ws.cell(row=r, column=4).alignment = Alignment(horizontal="left", vertical="center")
    ws.cell(row=r, column=7).number_format = "0.0"
    ws.row_dimensions[r].height = 18

add_dropdown(ws, "C", 2, 2 + N_ROWS - 1, ["Upper A", "Lower A", "Upper B", "Lower B"])
add_dropdown(ws, "H", 2, 2 + N_ROWS - 1, [str(x) for x in [5, 5.5, 6, 6.5, 7, 7.5, 8, 8.5, 9, 9.5, 10]])
ws.freeze_panes = "A2"


# -------- TAB: Running Log --------
ws = wb.create_sheet("Running Log")
headers = ["Date", "Week", "Distance (km)", "Time (min)", "Pace (min/km)", "Avg HR", "Max HR", "RPE", "Type", "Notes"]
widths = [12, 7, 14, 12, 14, 10, 10, 7, 16, 42]
set_header_row(ws, 1, headers, widths)

# Pre-populate runs
for i, run in enumerate(RUN_LOG, start=2):
    d, dist, t, ahr, mhr, rpe, typ, note = run
    ws.cell(row=i, column=1, value=d).number_format = "yyyy-mm-dd"
    ws.cell(row=i, column=3, value=dist).number_format = "0.00"
    ws.cell(row=i, column=4, value=t).number_format = "0.0"
    ws.cell(row=i, column=6, value=ahr)
    ws.cell(row=i, column=7, value=mhr)
    ws.cell(row=i, column=8, value=rpe)
    ws.cell(row=i, column=9, value=typ)
    ws.cell(row=i, column=10, value=note)

for r in range(2, 152):
    ws.cell(row=r, column=2, value=f'=IF(A{r}="","",IFERROR(INT((A{r}-DATE(2026,4,27))/7)+1,""))')
    ws.cell(row=r, column=5, value=f'=IF(OR(C{r}="",D{r}=""),"",TEXT(D{r}/C{r}/1440,"m:ss"))')
    style_body_row(ws, r, 10, zebra=(r % 2 == 0), notes_col=10)
    ws.cell(row=r, column=1).number_format = "yyyy-mm-dd"
    ws.cell(row=r, column=3).number_format = "0.00"
    ws.cell(row=r, column=4).number_format = "0.0"
    ws.row_dimensions[r].height = 18

add_dropdown(ws, "I", 2, 151, ["Easy", "Easy + pickups", "Long easy", "Tempo", "Other"])
add_dropdown(ws, "H", 2, 151, [str(x) for x in [3, 3.5, 4, 4.5, 5, 5.5, 6, 6.5, 7, 7.5, 8, 8.5, 9, 9.5, 10]])
ws.freeze_panes = "A2"


# -------- TAB: Body Metrics --------
ws = wb.create_sheet("Body Metrics")
headers = ["Date", "Week", "Body Weight (kg)", "Sleep (hrs)", "Waist (cm)", "Chest (cm)", "Arm (cm)", "Thigh (cm)", "Notes"]
widths = [12, 7, 17, 12, 12, 12, 11, 12, 42]
set_header_row(ws, 1, headers, widths)

for i, (d, bw, sleep, note) in enumerate(BODY_LOG, start=2):
    ws.cell(row=i, column=1, value=d).number_format = "yyyy-mm-dd"
    ws.cell(row=i, column=3, value=bw).number_format = "0.0"
    ws.cell(row=i, column=4, value=sleep)
    ws.cell(row=i, column=9, value=note)

for r in range(2, 200):
    ws.cell(row=r, column=2, value=f'=IF(A{r}="","",IFERROR(INT((A{r}-DATE(2026,4,27))/7)+1,""))')
    style_body_row(ws, r, 9, zebra=(r % 2 == 0), notes_col=9)
    ws.cell(row=r, column=1).number_format = "yyyy-mm-dd"
    for c in (3, 4, 5, 6, 7, 8):
        ws.cell(row=r, column=c).number_format = "0.0"
    ws.row_dimensions[r].height = 18
ws.freeze_panes = "A2"


# -------- TAB: Personal Records --------
ws = wb.create_sheet("Personal Records")
set_header_row(ws, 1, ["Exercise", "Best Weight (kg)", "Best Reps @ Weight", "Date Achieved", "Best e1RM (kg)", "Best Single-Set Volume"], [28, 16, 18, 14, 14, 22])

exercises = [
    "Bench Press", "Barbell Row", "Incline DB Press", "Lat Pulldown",
    "Pec Deck", "Triceps Pushdown", "DB Curl",
    "Back Squat", "Romanian Deadlift", "Leg Press", "Lying Leg Curl", "Standing Calf Raise", "Hanging Knee Raise",
    "Overhead Press", "Cable Row", "DB Lateral Raise", "Face Pull", "Skull Crusher",
    "Trap Bar Deadlift", "Bulgarian Split Squat", "Hip Thrust",
]
LIFT_TAB = "'Lifting Log'"
# SUMPRODUCT-MAX pattern (works without MAXIFS, which is missing in some Excel/LibreOffice versions)
for i, ex in enumerate(exercises, 2):
    has_data = f'COUNTIF({LIFT_TAB}!D:D,A{i})'
    ws.cell(row=i, column=1, value=ex)
    ws.cell(row=i, column=2, value=(
        f'=IF({has_data}=0,"",SUMPRODUCT(MAX(({LIFT_TAB}!D2:D801=A{i})*{LIFT_TAB}!G2:G801)))'
    )).number_format = "0.0"
    ws.cell(row=i, column=3, value=(
        f'=IF({has_data}=0,"",SUMPRODUCT(MAX(({LIFT_TAB}!D2:D801=A{i})*({LIFT_TAB}!G2:G801=B{i})*{LIFT_TAB}!F2:F801)))'
    ))
    ws.cell(row=i, column=4, value=(
        f'=IF({has_data}=0,"",SUMPRODUCT(MAX(({LIFT_TAB}!D2:D801=A{i})*({LIFT_TAB}!G2:G801=B{i})*{LIFT_TAB}!A2:A801)))'
    )).number_format = "yyyy-mm-dd"
    ws.cell(row=i, column=5, value=(
        f'=IF({has_data}=0,"",SUMPRODUCT(MAX(({LIFT_TAB}!D2:D801=A{i})*IFERROR({LIFT_TAB}!J2:J801*1,0))))'
    )).number_format = "0.0"
    ws.cell(row=i, column=6, value=(
        f'=IF({has_data}=0,"",SUMPRODUCT(MAX(({LIFT_TAB}!D2:D801=A{i})*IFERROR({LIFT_TAB}!I2:I801*1,0))))'
    )).number_format = "0"
    style_body_row(ws, i, 6, zebra=(i % 2 == 0))
    ws.cell(row=i, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[i].height = 20
ws.freeze_panes = "A2"


# -------- TAB: Weekly Summary --------
ws = wb.create_sheet("Weekly Summary")
headers = ["Week", "Date Range", "Lifting Sessions", "Total Sets", "Total Volume (kg)", "Avg RPE", "Runs", "Run km", "Run Time (min)", "Avg Run HR", "Avg BW (kg)"]
widths = [7, 22, 17, 12, 19, 10, 8, 10, 14, 12, 13]
set_header_row(ws, 1, headers, widths)

for week in range(1, 12):
    r = week + 1
    ws.cell(row=r, column=1, value=week)
    ws.cell(row=r, column=2, value=f'=TEXT(DATE(2026,4,27)+(A{r}-1)*7,"mmm d")&" – "&TEXT(DATE(2026,4,27)+A{r}*7-1,"mmm d")')
    # Sessions = distinct dates per week (each Set#=1 row contributes 1/N where N = # set#=1 rows on that date)
    ws.cell(row=r, column=3, value=(
        f'=SUMPRODUCT(IFERROR((({LIFT_TAB}!B2:B801=A{r})*({LIFT_TAB}!E2:E801=1))/'
        f'COUNTIFS({LIFT_TAB}!A:A,{LIFT_TAB}!A2:A801,{LIFT_TAB}!E:E,1),0))'
    ))
    ws.cell(row=r, column=4, value=f'=COUNTIFS({LIFT_TAB}!B:B,A{r})')
    ws.cell(row=r, column=5, value=f'=SUMIFS({LIFT_TAB}!I:I,{LIFT_TAB}!B:B,A{r})').number_format = "#,##0"
    ws.cell(row=r, column=6, value=f'=IFERROR(AVERAGEIFS({LIFT_TAB}!H:H,{LIFT_TAB}!B:B,A{r}),"")').number_format = "0.0"
    ws.cell(row=r, column=7, value=f"=COUNTIFS('Running Log'!B:B,A{r})")
    ws.cell(row=r, column=8, value=f"=SUMIFS('Running Log'!C:C,'Running Log'!B:B,A{r})").number_format = "0.00"
    ws.cell(row=r, column=9, value=f"=SUMIFS('Running Log'!D:D,'Running Log'!B:B,A{r})").number_format = "0"
    ws.cell(row=r, column=10, value=f"=IFERROR(AVERAGEIFS('Running Log'!F:F,'Running Log'!B:B,A{r}),\"\")").number_format = "0"
    ws.cell(row=r, column=11, value=f"=IFERROR(AVERAGEIFS('Body Metrics'!C:C,'Body Metrics'!B:B,A{r}),\"\")").number_format = "0.0"
    style_body_row(ws, r, 11, zebra=(r % 2 == 0))
    ws.row_dimensions[r].height = 20
ws.freeze_panes = "A2"


# -------- TAB: Progression Data (chart source) --------
# For each priority lift, weekly best e1RM (max e1RM in Lifting Log for that lift in that week).
ws = wb.create_sheet("Progression Data")
priority_lifts = [
    "Bench Press", "Overhead Press", "Back Squat", "Romanian Deadlift",
    "Lat Pulldown", "Barbell Row", "Cable Row", "Incline DB Press",
]
# Header: Week | Date Range | <lift1> | <lift2> | ...
header = ["Week", "Date Range"] + priority_lifts
widths = [8, 22] + [16] * len(priority_lifts)
set_header_row(ws, 1, header, widths)

for week in range(1, 12):
    r = week + 1
    ws.cell(row=r, column=1, value=week)
    ws.cell(row=r, column=2, value=f'=TEXT(DATE(2026,4,27)+(A{r}-1)*7,"mmm d")')
    for c, lift in enumerate(priority_lifts, start=3):
        # Max e1RM that week for this lift; blank if no sets logged (SUMPRODUCT-MAX, MAXIFS not portable)
        ws.cell(row=r, column=c, value=(
            f'=IF(COUNTIFS({LIFT_TAB}!B:B,A{r},{LIFT_TAB}!D:D,"{lift}")=0,"",'
            f'SUMPRODUCT(MAX(({LIFT_TAB}!B2:B801=A{r})*({LIFT_TAB}!D2:D801="{lift}")*'
            f'IFERROR({LIFT_TAB}!J2:J801*1,0))))'
        )).number_format = "0.0"
    style_body_row(ws, r, 2 + len(priority_lifts), zebra=(r % 2 == 0))
    ws.row_dimensions[r].height = 20
ws.freeze_panes = "C2"


# -------- TAB: Charts --------
ws_chart = wb.create_sheet("Charts")
ws_chart.column_dimensions["A"].width = 3
for col in ("B", "C", "D", "E", "F"):
    ws_chart.column_dimensions[col].width = 22

# Title
ws_chart.cell(row=1, column=2, value="Block 1 — Visual Progression").font = Font(name="Calibri", bold=True, size=18, color="1F4E78")
ws_chart.cell(row=2, column=2, value="Charts auto-update from Lifting Log / Running Log / Body Metrics. KPI cells live below.").font = Font(name="Calibri", italic=True, size=11, color="595959")
ws_chart.row_dimensions[1].height = 28

# KPI summary block (live formulas, not stale values)
kpi_row = 4
ws_chart.cell(row=kpi_row, column=2, value="KPI snapshot").font = Font(name="Calibri", bold=True, size=13, color="1F4E78")

def best_e1rm(lift_name):
    return (
        f'=IF(COUNTIF({LIFT_TAB}!D:D,"{lift_name}")=0,"",'
        f'SUMPRODUCT(MAX(({LIFT_TAB}!D2:D801="{lift_name}")*IFERROR({LIFT_TAB}!J2:J801*1,0))))'
    )

kpis = [
    ("Sessions logged",          f'=SUMPRODUCT(IFERROR(({LIFT_TAB}!E2:E801=1)/COUNTIFS({LIFT_TAB}!A:A,{LIFT_TAB}!A2:A801,{LIFT_TAB}!E:E,1),0))', "0"),
    ("Total working sets",       f"=COUNTA({LIFT_TAB}!E2:E801)",                              "0"),
    ("Total volume (kg)",        f"=SUM({LIFT_TAB}!I2:I801)",                                 "#,##0"),
    ("Avg RPE (all sets)",       f'=IFERROR(AVERAGE({LIFT_TAB}!H2:H801),"")',                 "0.00"),
    ("Bench best e1RM (kg)",     best_e1rm("Bench Press"),                                    "0.0"),
    ("Squat best e1RM (kg)",     best_e1rm("Back Squat"),                                     "0.0"),
    ("OHP best e1RM (kg)",       best_e1rm("Overhead Press"),                                 "0.0"),
    ("Latest BW (kg)",           f'=IFERROR(LOOKUP(2,1/(\'Body Metrics\'!C2:C200<>""),\'Body Metrics\'!C2:C200),"")', "0.0"),
    ("BW Δ vs Wk 1 (kg)",        f'=IFERROR(LOOKUP(2,1/(\'Body Metrics\'!C2:C200<>""),\'Body Metrics\'!C2:C200)-INDEX(\'Body Metrics\'!C:C,MATCH(TRUE,INDEX(\'Body Metrics\'!C2:C200<>"",0),0)+1),"")', "+0.00;-0.00;0.00"),
    ("Total run km",             f"=SUM('Running Log'!C2:C151)",                              "0.00"),
]
for i, (label, formula, fmt) in enumerate(kpis):
    r = kpi_row + 1 + i
    ws_chart.cell(row=r, column=2, value=label).font = BODY_FONT
    cell = ws_chart.cell(row=r, column=3, value=formula)
    cell.font = Font(name="Calibri", bold=True, size=11, color="1F4E78")
    cell.number_format = fmt
    cell.alignment = Alignment(horizontal="right")
    ws_chart.cell(row=r, column=2).fill = ZEBRA_FILL if i % 2 == 0 else PatternFill()
    ws_chart.cell(row=r, column=3).fill = ZEBRA_FILL if i % 2 == 0 else PatternFill()
    ws_chart.cell(row=r, column=2).border = THIN_BORDER
    ws_chart.cell(row=r, column=3).border = THIN_BORDER

# Charts start below KPI block
prog_tab = "'Progression Data'"
n_lifts = len(priority_lifts)

def styled_line(title, ymin=None, ymax=None):
    ch = LineChart()
    ch.title = title
    ch.style = 12
    ch.height = 11
    ch.width = 24
    ch.legend.position = "b"
    if ymin is not None:
        ch.y_axis.scaling.min = ymin
    if ymax is not None:
        ch.y_axis.scaling.max = ymax
    ch.y_axis.majorGridlines = None
    return ch

def styled_bar(title):
    ch = BarChart()
    ch.type = "col"
    ch.style = 11
    ch.title = title
    ch.height = 11
    ch.width = 24
    ch.legend = None
    ch.dataLabels = DataLabelList(showVal=True)
    return ch

# Chart 1: e1RM trend per priority lift
ch1 = styled_line("e1RM (kg) by Week — priority lifts")
ch1.y_axis.title = "e1RM (kg)"
ch1.x_axis.title = "Block week"
data = Reference(wb["Progression Data"], min_col=3, min_row=1, max_col=2 + n_lifts, max_row=12)
cats = Reference(wb["Progression Data"], min_col=1, min_row=2, max_row=12)
ch1.add_data(data, titles_from_data=True)
ch1.set_categories(cats)
for s in ch1.series:
    s.smooth = False
    s.marker = Marker(symbol="circle", size=6)
ws_chart.add_chart(ch1, "B18")

# Chart 2: Weekly Total Volume (Bar)
ch2 = styled_bar("Total Lifting Volume (kg) by Week")
ch2.y_axis.title = "Volume (kg)"
ch2.x_axis.title = "Block week"
data2 = Reference(wb["Weekly Summary"], min_col=5, min_row=1, max_col=5, max_row=12)
cats2 = Reference(wb["Weekly Summary"], min_col=1, min_row=2, max_row=12)
ch2.add_data(data2, titles_from_data=True)
ch2.set_categories(cats2)
ws_chart.add_chart(ch2, "B41")

# Chart 3: Weekly Avg RPE (Line)
ch3 = styled_line("Avg Session RPE by Week", ymin=5, ymax=10)
ch3.y_axis.title = "RPE"
ch3.x_axis.title = "Block week"
data3 = Reference(wb["Weekly Summary"], min_col=6, min_row=1, max_col=6, max_row=12)
cats3 = Reference(wb["Weekly Summary"], min_col=1, min_row=2, max_row=12)
ch3.add_data(data3, titles_from_data=True)
ch3.set_categories(cats3)
for s in ch3.series:
    s.marker = Marker(symbol="circle", size=6)
ws_chart.add_chart(ch3, "B64")

# Chart 4: Body Weight trend — DateAxis for proportional spacing
ch4 = styled_line("Body Weight (kg) over time")
ch4.y_axis.title = "BW (kg)"
ch4.x_axis = DateAxis(crossAx=100)
ch4.x_axis.number_format = "mmm d"
ch4.x_axis.majorTimeUnit = "days"
ch4.x_axis.title = "Date"
data4 = Reference(wb["Body Metrics"], min_col=3, min_row=1, max_col=3, max_row=200)
cats4 = Reference(wb["Body Metrics"], min_col=1, min_row=2, max_row=200)
ch4.add_data(data4, titles_from_data=True)
ch4.set_categories(cats4)
for s in ch4.series:
    s.marker = Marker(symbol="circle", size=6)
    s.trendline = Trendline(trendlineType="linear", dispEq=False, dispRSqr=False)
ws_chart.add_chart(ch4, "B87")

# Chart 5: Working sets per week (Bar)
ch5 = styled_bar("Working Sets by Week")
ch5.y_axis.title = "Sets"
ch5.x_axis.title = "Block week"
data5 = Reference(wb["Weekly Summary"], min_col=4, min_row=1, max_col=4, max_row=12)
cats5 = Reference(wb["Weekly Summary"], min_col=1, min_row=2, max_row=12)
ch5.add_data(data5, titles_from_data=True)
ch5.set_categories(cats5)
ws_chart.add_chart(ch5, "B110")

# Chart 6: Run distance per week (Bar)
ch6 = styled_bar("Run Distance (km) by Week")
ch6.y_axis.title = "km"
ch6.x_axis.title = "Block week"
data6 = Reference(wb["Weekly Summary"], min_col=8, min_row=1, max_col=8, max_row=12)
cats6 = Reference(wb["Weekly Summary"], min_col=1, min_row=2, max_row=12)
ch6.add_data(data6, titles_from_data=True)
ch6.set_categories(cats6)
ws_chart.add_chart(ch6, "B133")


# -------- TAB: Reference --------
ws = wb.create_sheet("Reference")
for i, w in enumerate([22, 22, 28, 22], 1):
    ws.column_dimensions[get_column_letter(i)].width = w

row = 1
ws.cell(row=row, column=1, value="Reference Data").font = Font(name="Calibri", bold=True, size=16, color="1F4E78")
ws.row_dimensions[row].height = 26
row += 2

def write_ref_table(ws, row, title, header, rows):
    ws.cell(row=row, column=1, value=title).font = Font(name="Calibri", bold=True, size=13, color="1F4E78")
    row += 1
    for c, v in enumerate(header, 1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = THIN_BORDER
        cell.alignment = CENTER
    ws.row_dimensions[row].height = 22
    row += 1
    for ri, r_data in enumerate(rows):
        for c, v in enumerate(r_data, 1):
            cell = ws.cell(row=row, column=c, value=v)
            cell.font = BODY_FONT
            cell.border = THIN_BORDER
            if ri % 2 == 0:
                cell.fill = ZEBRA_FILL
        ws.row_dimensions[row].height = 18
        row += 1
    return row + 1

row = write_ref_table(ws, row, "HR Zones", ["Zone", "BPM", "Use"], [
    ("Z1 Warm-up", "100–119", "Warm-up only"),
    ("Z2 Easy", "120–139", "Easy runs, recovery"),
    ("Z3 Aerobic", "140–159", "Steady running"),
    ("Z4 Threshold", "160–179", "Tempo/threshold reps"),
    ("Z5 Max", ">179", "Short hard intervals"),
])
row = write_ref_table(ws, row, "Pacing Zones", ["Zone", "Pace/km", "Speed (km/h)", "Use"], [
    ("Easy", "6:30–7:00", "8.6–9.2", "Wed + Sun runs"),
    ("Steady", "6:00–6:30", "9.2–10.0", "Slight progression"),
    ("Pickup", "5:30–5:50", "10.3–10.9", "Wks 6–7 only"),
])
row = write_ref_table(ws, row, "Weekly Schedule (Wk 2+)", ["Day", "Session"], [
    ("Mon", "Upper A"),
    ("Tue", "Lower A"),
    ("Wed", "REST"),
    ("Thu", "Easy run Z2"),
    ("Fri", "Upper B"),
    ("Sat", "Lower B"),
    ("Sun", "Easy run Z2"),
])
row = write_ref_table(ws, row, "RPE Reference", ["RPE", "Reps in Reserve", "Use"], [
    ("6", "4 left", "Warm-up"),
    ("7", "3 left", "Early sets, accumulation"),
    ("8", "2 left", "Top sets — most of block"),
    ("9", "1 left", "Peak weeks, isolation"),
    ("10", "Failure", "Isolation only"),
])


# -------- Reorder + save --------
order = ["README", "Charts", "Lifting Log", "Running Log", "Body Metrics",
         "Personal Records", "Weekly Summary", "Progression Data", "Reference"]
for idx, name in enumerate(order):
    if name in wb.sheetnames:
        wb.move_sheet(name, offset=idx - wb.sheetnames.index(name))

wb.save(OUTPUT)
print(f"Saved: {OUTPUT}")
print(f"  Lifting sets logged: {len(LIFT_SETS)}")
print(f"  Body weight entries: {len(BODY_LOG)}")
print(f"  Runs logged: {len(RUN_LOG)}")
