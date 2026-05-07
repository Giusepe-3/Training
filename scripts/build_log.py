"""Generate training_log.xlsx — multi-tab logger for the 11-week block.

Pre-populates logged sessions (Wk 1+) and adds a Charts tab with progression
visualizations: e1RM trend per priority lift, weekly volume, weekly RPE, BW.

Run:  python3 scripts/build_log.py
"""
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.trendline import Trendline
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

REPO_ROOT = Path(__file__).resolve().parent.parent
OUTPUT = REPO_ROOT / "logs" / "training_log.xlsx"

BLOCK_START = date(2026, 4, 27)

HEADER_FILL = PatternFill("solid", start_color="1F4E78")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=11)
SECTION_FILL = PatternFill("solid", start_color="D9E1F2")
BODY_FONT = Font(name="Arial", size=10)
THIN_BORDER = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)


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
]

# Body weight log: (date, BW kg, sleep h, notes)
BODY_LOG = [
    (date(2026, 4, 29), 77.7, 8.0, "Wk 1 Wed AM"),
    (date(2026, 5, 7),  77.3, 9.0, "Wk 2 Thu PM"),
]

# Running log: (date, dist_km, time_min, avg_hr, max_hr, rpe, type, notes)
RUN_LOG = []  # none logged yet


# ============== HELPERS ==============

def set_header_row(ws, row, headers, widths=None):
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=i, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 22


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
ws.column_dimensions["A"].width = 100

readme = [
    ("Training Log — 11-Week Hypertrophy Block", "title"),
    ("Block: 2026-04-27 → 2026-07-12", "section"),
    ("", ""),
    ("Tabs:", "section"),
    ("• Lifting Log — every working set", "body"),
    ("• Running Log — every run", "body"),
    ("• Body Metrics — daily BW + weekly measurements", "body"),
    ("• Personal Records — auto best lifts", "body"),
    ("• Weekly Summary — auto weekly volume/RPE/runs/BW", "body"),
    ("• Progression Data — auto weekly best e1RM per priority lift (chart source)", "body"),
    ("• Charts — visual progression: e1RM, weekly volume, RPE, BW", "body"),
    ("• Reference — HR/pace zones, schedule, RPE", "body"),
    ("", ""),
    ("Pre-populated from logs/session_log.md + logs/sessions/. Re-run scripts/build_log.py to refresh.", "body"),
    ("", ""),
    ("Data entry: edit Lifting Log / Running Log / Body Metrics. Auto tabs recompute on open.", "body"),
    ("", ""),
    ("RPE: 6=4 RIR, 7=3 RIR, 8=2 RIR, 9=1 RIR, 10=failure (isolation only).", "body"),
]
for i, (line, kind) in enumerate(readme, 1):
    cell = ws.cell(row=i, column=1, value=line)
    if kind == "title":
        cell.font = Font(name="Arial", bold=True, size=16, color="1F4E78")
    elif kind == "section":
        cell.font = Font(name="Arial", bold=True, size=12, color="1F4E78")
    else:
        cell.font = BODY_FONT


# -------- TAB: Lifting Log --------
ws = wb.create_sheet("Lifting Log")
headers = ["Date", "Week", "Day", "Exercise", "Set #", "Reps", "Weight (kg)", "RPE", "Volume (kg)", "e1RM (kg)", "Notes"]
widths = [12, 8, 10, 26, 8, 8, 12, 8, 13, 12, 32]
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
    for c in range(1, 12):
        cell = ws.cell(row=r, column=c)
        cell.font = BODY_FONT
        cell.border = THIN_BORDER
        if c == 1:
            cell.number_format = "yyyy-mm-dd"
        if c == 7:
            cell.number_format = "0.0"

add_dropdown(ws, "C", 2, 2 + N_ROWS - 1, ["Upper A", "Lower A", "Upper B", "Lower B"])
add_dropdown(ws, "H", 2, 2 + N_ROWS - 1, [str(x) for x in [5, 5.5, 6, 6.5, 7, 7.5, 8, 8.5, 9, 9.5, 10]])
ws.freeze_panes = "A2"


# -------- TAB: Running Log --------
ws = wb.create_sheet("Running Log")
headers = ["Date", "Week", "Distance (km)", "Time (min)", "Pace (min/km)", "Avg HR", "Max HR", "RPE", "Type", "Notes"]
widths = [12, 8, 14, 12, 14, 10, 10, 8, 14, 32]
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
    ws.cell(row=r, column=5, value=f'=IF(OR(C{r}="",D{r}=""),"",D{r}/C{r})').number_format = "0.00"
    for c in range(1, 11):
        cell = ws.cell(row=r, column=c)
        cell.font = BODY_FONT
        cell.border = THIN_BORDER
        if c == 1:
            cell.number_format = "yyyy-mm-dd"
        if c == 3:
            cell.number_format = "0.00"
        if c == 4:
            cell.number_format = "0.0"

add_dropdown(ws, "I", 2, 151, ["Easy", "Easy + pickups", "Long easy", "Tempo", "Other"])
add_dropdown(ws, "H", 2, 151, [str(x) for x in [3, 3.5, 4, 4.5, 5, 5.5, 6, 6.5, 7, 7.5, 8]])
ws.freeze_panes = "A2"


# -------- TAB: Body Metrics --------
ws = wb.create_sheet("Body Metrics")
headers = ["Date", "Week", "Body Weight (kg)", "Sleep (hrs)", "Waist (cm)", "Chest (cm)", "Arm (cm)", "Thigh (cm)", "Notes"]
widths = [12, 8, 17, 12, 12, 12, 11, 12, 32]
set_header_row(ws, 1, headers, widths)

for i, (d, bw, sleep, note) in enumerate(BODY_LOG, start=2):
    ws.cell(row=i, column=1, value=d).number_format = "yyyy-mm-dd"
    ws.cell(row=i, column=3, value=bw).number_format = "0.0"
    ws.cell(row=i, column=4, value=sleep)
    ws.cell(row=i, column=9, value=note)

for r in range(2, 200):
    ws.cell(row=r, column=2, value=f'=IF(A{r}="","",IFERROR(INT((A{r}-DATE(2026,4,27))/7)+1,""))')
    for c in range(1, 10):
        cell = ws.cell(row=r, column=c)
        cell.font = BODY_FONT
        cell.border = THIN_BORDER
        if c == 1:
            cell.number_format = "yyyy-mm-dd"
        if c in (3, 4, 5, 6, 7, 8):
            cell.number_format = "0.0"
ws.freeze_panes = "A2"


# -------- TAB: Personal Records --------
ws = wb.create_sheet("Personal Records")
for i, w in enumerate([28, 14, 12, 14, 14, 14], 1):
    ws.column_dimensions[get_column_letter(i)].width = w
set_header_row(ws, 1, ["Exercise", "Best Weight (kg)", "Best Reps @ Weight", "Date Achieved", "Best e1RM (kg)", "Best Single-Set Volume"])

exercises = [
    "Bench Press", "Barbell Row", "Incline DB Press", "Lat Pulldown",
    "Pec Deck", "Triceps Pushdown", "DB Curl",
    "Back Squat", "Romanian Deadlift", "Leg Press", "Lying Leg Curl", "Standing Calf Raise", "Hanging Knee Raise",
    "Overhead Press", "Cable Row", "DB Lateral Raise", "Face Pull", "Skull Crusher",
    "Trap Bar Deadlift", "Bulgarian Split Squat", "Hip Thrust",
]
LIFT_TAB = "'Lifting Log'"
for i, ex in enumerate(exercises, 2):
    ws.cell(row=i, column=1, value=ex).font = BODY_FONT
    ws.cell(row=i, column=2, value=f'=IFERROR(MAXIFS({LIFT_TAB}!G:G,{LIFT_TAB}!D:D,A{i}),"")').number_format = "0.0"
    ws.cell(row=i, column=3, value=f'=IFERROR(MAXIFS({LIFT_TAB}!F:F,{LIFT_TAB}!D:D,A{i},{LIFT_TAB}!G:G,B{i}),"")')
    ws.cell(row=i, column=4, value=f'=IFERROR(MAXIFS({LIFT_TAB}!A:A,{LIFT_TAB}!D:D,A{i},{LIFT_TAB}!G:G,B{i}),"")').number_format = "yyyy-mm-dd"
    ws.cell(row=i, column=5, value=f'=IFERROR(MAXIFS({LIFT_TAB}!J:J,{LIFT_TAB}!D:D,A{i}),"")').number_format = "0.0"
    ws.cell(row=i, column=6, value=f'=IFERROR(MAXIFS({LIFT_TAB}!I:I,{LIFT_TAB}!D:D,A{i}),"")').number_format = "0"
    for c in range(1, 7):
        ws.cell(row=i, column=c).border = THIN_BORDER
        ws.cell(row=i, column=c).font = BODY_FONT


# -------- TAB: Weekly Summary --------
ws = wb.create_sheet("Weekly Summary")
headers = ["Week", "Date Range", "Lifting Sessions", "Total Sets", "Total Volume (kg)", "Avg RPE", "Runs", "Run km", "Run Time (min)", "Avg Run HR", "Avg BW (kg)"]
widths = [8, 22, 17, 12, 19, 10, 8, 10, 14, 12, 13]
set_header_row(ws, 1, headers, widths)

for week in range(1, 12):
    r = week + 1
    ws.cell(row=r, column=1, value=week)
    ws.cell(row=r, column=2, value=f'=TEXT(DATE(2026,4,27)+(A{r}-1)*7,"mmm d")&" – "&TEXT(DATE(2026,4,27)+A{r}*7-1,"mmm d")')
    ws.cell(row=r, column=3, value=f'=COUNTIFS({LIFT_TAB}!B:B,A{r},{LIFT_TAB}!E:E,1)')
    ws.cell(row=r, column=4, value=f'=COUNTIFS({LIFT_TAB}!B:B,A{r})')
    ws.cell(row=r, column=5, value=f'=SUMIFS({LIFT_TAB}!I:I,{LIFT_TAB}!B:B,A{r})').number_format = "#,##0"
    ws.cell(row=r, column=6, value=f'=IFERROR(AVERAGEIFS({LIFT_TAB}!H:H,{LIFT_TAB}!B:B,A{r}),"")').number_format = "0.0"
    ws.cell(row=r, column=7, value=f"=COUNTIFS('Running Log'!B:B,A{r})")
    ws.cell(row=r, column=8, value=f"=SUMIFS('Running Log'!C:C,'Running Log'!B:B,A{r})").number_format = "0.00"
    ws.cell(row=r, column=9, value=f"=SUMIFS('Running Log'!D:D,'Running Log'!B:B,A{r})").number_format = "0"
    ws.cell(row=r, column=10, value=f"=IFERROR(AVERAGEIFS('Running Log'!F:F,'Running Log'!B:B,A{r}),\"\")").number_format = "0"
    ws.cell(row=r, column=11, value=f"=IFERROR(AVERAGEIFS('Body Metrics'!C:C,'Body Metrics'!B:B,A{r}),\"\")").number_format = "0.0"
    for c in range(1, 12):
        ws.cell(row=r, column=c).border = THIN_BORDER
        ws.cell(row=r, column=c).font = BODY_FONT
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
        # Max e1RM that week for this lift; blank if no sets logged
        ws.cell(row=r, column=c, value=(
            f'=IFERROR(IF(COUNTIFS({LIFT_TAB}!B:B,A{r},{LIFT_TAB}!D:D,"{lift}")=0,"",'
            f'MAXIFS({LIFT_TAB}!J:J,{LIFT_TAB}!B:B,A{r},{LIFT_TAB}!D:D,"{lift}")),"")'
        )).number_format = "0.0"
    for c in range(1, 3 + len(priority_lifts)):
        ws.cell(row=r, column=c).border = THIN_BORDER
        ws.cell(row=r, column=c).font = BODY_FONT


# -------- TAB: Charts --------
ws_chart = wb.create_sheet("Charts")
ws_chart.column_dimensions["A"].width = 4
ws_chart.cell(row=1, column=2, value="Visual Progression").font = Font(name="Arial", bold=True, size=16, color="1F4E78")

# Chart 1: e1RM trend per priority lift (Line chart, source = Progression Data)
prog_tab = "'Progression Data'"
n_lifts = len(priority_lifts)

ch1 = LineChart()
ch1.title = "e1RM (kg) by Week — priority lifts"
ch1.style = 12
ch1.y_axis.title = "e1RM (kg)"
ch1.x_axis.title = "Block week"
ch1.height = 12
ch1.width = 22
data = Reference(wb["Progression Data"], min_col=3, min_row=1, max_col=2 + n_lifts, max_row=12)
cats = Reference(wb["Progression Data"], min_col=1, min_row=2, max_row=12)
ch1.add_data(data, titles_from_data=True)
ch1.set_categories(cats)
for s in ch1.series:
    s.smooth = False
ws_chart.add_chart(ch1, "B3")

# Chart 2: Weekly Total Volume (Bar)
ch2 = BarChart()
ch2.type = "col"
ch2.style = 11
ch2.title = "Total Lifting Volume (kg) by Week"
ch2.y_axis.title = "Volume (kg)"
ch2.x_axis.title = "Block week"
ch2.height = 10
ch2.width = 22
data2 = Reference(wb["Weekly Summary"], min_col=5, min_row=1, max_col=5, max_row=12)
cats2 = Reference(wb["Weekly Summary"], min_col=1, min_row=2, max_row=12)
ch2.add_data(data2, titles_from_data=True)
ch2.set_categories(cats2)
ch2.dataLabels = DataLabelList(showVal=True)
ws_chart.add_chart(ch2, "B28")

# Chart 3: Weekly Avg RPE (Line)
ch3 = LineChart()
ch3.title = "Avg Session RPE by Week"
ch3.style = 13
ch3.y_axis.title = "RPE"
ch3.x_axis.title = "Block week"
ch3.height = 10
ch3.width = 22
data3 = Reference(wb["Weekly Summary"], min_col=6, min_row=1, max_col=6, max_row=12)
cats3 = Reference(wb["Weekly Summary"], min_col=1, min_row=2, max_row=12)
ch3.add_data(data3, titles_from_data=True)
ch3.set_categories(cats3)
ch3.y_axis.scaling.min = 5
ch3.y_axis.scaling.max = 10
ws_chart.add_chart(ch3, "B49")

# Chart 4: Body Weight trend (Line, raw entries — uses Body Metrics directly)
ch4 = LineChart()
ch4.title = "Body Weight (kg) over time"
ch4.style = 10
ch4.y_axis.title = "BW (kg)"
ch4.x_axis.title = "Date"
ch4.height = 10
ch4.width = 22
# 60 rows of capacity
data4 = Reference(wb["Body Metrics"], min_col=3, min_row=1, max_col=3, max_row=60)
cats4 = Reference(wb["Body Metrics"], min_col=1, min_row=2, max_row=60)
ch4.add_data(data4, titles_from_data=True)
ch4.set_categories(cats4)
# Add trendline
if ch4.series:
    ch4.series[0].trendline = Trendline(trendlineType="linear")
ws_chart.add_chart(ch4, "B70")

# Chart 5: Per-session set volume — Bench / OHP / Squat (Line, x = date)
# Build helper rows under Progression Data with per-session weighted-best-set.
# Use a simple approach: chart all sets directly from Lifting Log filtered manually.
# Skip per-session chart for now (would need extra plumbing). Weekly best e1RM covers it.

# Chart 6: Weekly Total Sets (volume by count)
ch6 = BarChart()
ch6.type = "col"
ch6.style = 14
ch6.title = "Working Sets by Week"
ch6.y_axis.title = "Total sets"
ch6.x_axis.title = "Block week"
ch6.height = 10
ch6.width = 22
data6 = Reference(wb["Weekly Summary"], min_col=4, min_row=1, max_col=4, max_row=12)
cats6 = Reference(wb["Weekly Summary"], min_col=1, min_row=2, max_row=12)
ch6.add_data(data6, titles_from_data=True)
ch6.set_categories(cats6)
ch6.dataLabels = DataLabelList(showVal=True)
ws_chart.add_chart(ch6, "B91")


# -------- TAB: Reference --------
ws = wb.create_sheet("Reference")
for i, w in enumerate([22, 22, 28, 22], 1):
    ws.column_dimensions[get_column_letter(i)].width = w

row = 1
ws.cell(row=row, column=1, value="Reference Data").font = Font(name="Arial", bold=True, size=14, color="1F4E78")
row += 2

def write_ref_table(ws, row, title, header, rows):
    ws.cell(row=row, column=1, value=title).font = Font(name="Arial", bold=True, size=12)
    row += 1
    for c, v in enumerate(header, 1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = THIN_BORDER
    row += 1
    for r_data in rows:
        for c, v in enumerate(r_data, 1):
            cell = ws.cell(row=row, column=c, value=v)
            cell.font = BODY_FONT
            cell.border = THIN_BORDER
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
