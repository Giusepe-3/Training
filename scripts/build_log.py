"""Generate training_log.xlsx — multi-tab logger for the 8-week block."""
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
from openpyxl.comments import Comment

REPO_ROOT = Path(__file__).resolve().parent.parent
OUTPUT = REPO_ROOT / "logs" / "training_log.xlsx"

HEADER_FILL = PatternFill("solid", start_color="1F4E78")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=11)
SECTION_FILL = PatternFill("solid", start_color="D9E1F2")
SECTION_FONT = Font(name="Arial", bold=True, size=11)
BODY_FONT = Font(name="Arial", size=10)
THIN_BORDER = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)


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
    from openpyxl.worksheet.datavalidation import DataValidation
    dv = DataValidation(type="list", formula1=f'"{",".join(choices)}"', allow_blank=True)
    dv.add(f"{col_letter}{start_row}:{col_letter}{end_row}")
    ws.add_data_validation(dv)


wb = Workbook()


# ============== TAB 1: README ==============
ws = wb.active
ws.title = "README"
ws.column_dimensions["A"].width = 100

readme_lines = [
    ("Training Log — 11-Week Hypertrophy Block", "title"),
    ("", "blank"),
    ("Block dates: April 27 – July 12, 2026", "section"),
    ("Goal: max upper-body hypertrophy + maintained aerobic base", "body"),
    ("Approach: 4 lifts + 2 runs per week, recomp at maintenance calories", "body"),
    ("", "blank"),
    ("Tabs in this workbook:", "section"),
    ("• Lifting Log — log every set: date, exercise, set #, reps, weight, RPE", "body"),
    ("• Running Log — log every run: date, distance, time, pace, HR, RPE", "body"),
    ("• Body Metrics — daily weight + weekly measurements", "body"),
    ("• Personal Records — auto-tracked best lifts (top weight per exercise)", "body"),
    ("• Weekly Summary — auto-aggregated weekly volume and averages", "body"),
    ("• Reference — pacing zones, HR zones, plan reminders", "body"),
    ("", "blank"),
    ("How to use:", "section"),
    ("1. Log every working set in the Lifting Log right after the set.", "body"),
    ("2. Log every run in the Running Log immediately after.", "body"),
    ("3. Weight in Body Metrics every morning fasted, measurements every Monday.", "body"),
    ("4. Personal Records, Weekly Summary update automatically — don't edit those tabs.", "body"),
    ("5. Sunday weekly review: scan Weekly Summary, scan Personal Records, plan next week.", "body"),
    ("", "blank"),
    ("Rules:", "section"),
    ("• Don't skip RPE — it is THE most important field for progression decisions.", "body"),
    ("• If you forget to log, write down what you remember. Imperfect data > no data.", "body"),
    ("• Edit anything in Lifting Log / Running Log / Body Metrics. Don't edit auto tabs.", "body"),
    ("", "blank"),
    ("RPE scale (Reps in Reserve):", "section"),
    ("• RPE 6 = 4 reps left in tank — use for warm-ups", "body"),
    ("• RPE 7 = 3 reps left — early sets, accumulation phase", "body"),
    ("• RPE 8 = 2 reps left — top sets, normal training", "body"),
    ("• RPE 9 = 1 rep left — peak sessions, isolation movements", "body"),
    ("• RPE 10 = failure — only on isolation, never on compounds", "body"),
]
for i, (line, kind) in enumerate(readme_lines, 1):
    cell = ws.cell(row=i, column=1, value=line)
    if kind == "title":
        cell.font = Font(name="Arial", bold=True, size=16, color="1F4E78")
    elif kind == "section":
        cell.font = Font(name="Arial", bold=True, size=12, color="1F4E78")
    else:
        cell.font = BODY_FONT


# ============== TAB 2: LIFTING LOG ==============
ws = wb.create_sheet("Lifting Log")
headers = ["Date", "Week", "Day", "Exercise", "Set #", "Reps", "Weight (kg)", "RPE", "Volume (kg)", "Notes"]
widths = [12, 8, 10, 28, 8, 8, 12, 8, 13, 35]
set_header_row(ws, 1, headers, widths)

# Pre-fill 600 rows with formulas where applicable
N_ROWS = 600
for r in range(2, 2 + N_ROWS):
    ws.cell(row=r, column=2, value=f'=IF(A{r}="","",IFERROR(INT((A{r}-DATE(2026,4,27))/7)+1,""))')
    ws.cell(row=r, column=9, value=f'=IF(OR(F{r}="",G{r}=""),"",F{r}*G{r})')
    for c in range(1, 11):
        cell = ws.cell(row=r, column=c)
        cell.font = BODY_FONT
        cell.border = THIN_BORDER
        if c == 1:
            cell.number_format = "yyyy-mm-dd"
        if c == 7:
            cell.number_format = "0.0"
        if c == 9:
            cell.number_format = "0"

# Day dropdown
add_dropdown(ws, "C", 2, 2 + N_ROWS - 1, ["Upper A", "Lower A", "Upper B", "Lower B"])
# RPE dropdown
add_dropdown(ws, "H", 2, 2 + N_ROWS - 1, [str(x) for x in [5, 5.5, 6, 6.5, 7, 7.5, 8, 8.5, 9, 9.5, 10]])

ws.freeze_panes = "A2"


# ============== TAB 3: RUNNING LOG ==============
ws = wb.create_sheet("Running Log")
headers = ["Date", "Week", "Distance (km)", "Time (min)", "Pace (min/km)", "Avg HR", "Max HR", "RPE", "Type", "Notes"]
widths = [12, 8, 14, 12, 14, 10, 10, 8, 14, 35]
set_header_row(ws, 1, headers, widths)

for r in range(2, 102):
    ws.cell(row=r, column=2, value=f'=IF(A{r}="","",IFERROR(INT((A{r}-DATE(2026,4,27))/7)+1,""))')
    ws.cell(row=r, column=5, value=f'=IF(OR(C{r}="",D{r}=""),"",D{r}/C{r})')
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
        if c == 5:
            cell.number_format = "0.00"

add_dropdown(ws, "I", 2, 101, ["Easy", "Easy + pickups", "Long easy", "Tempo", "Other"])
add_dropdown(ws, "H", 2, 101, [str(x) for x in [3, 3.5, 4, 4.5, 5, 5.5, 6, 6.5, 7, 7.5, 8]])

ws.freeze_panes = "A2"


# ============== TAB 4: BODY METRICS ==============
ws = wb.create_sheet("Body Metrics")
headers = ["Date", "Week", "Body Weight (kg)", "Sleep (hrs)", "Waist (cm)", "Chest (cm)", "Arm (cm)", "Thigh (cm)", "Notes"]
widths = [12, 8, 17, 12, 12, 12, 11, 12, 35]
set_header_row(ws, 1, headers, widths)

for r in range(2, 122):
    ws.cell(row=r, column=2, value=f'=IF(A{r}="","",IFERROR(INT((A{r}-DATE(2026,4,27))/7)+1,""))')
    for c in range(1, 10):
        cell = ws.cell(row=r, column=c)
        cell.font = BODY_FONT
        cell.border = THIN_BORDER
        if c == 1:
            cell.number_format = "yyyy-mm-dd"
        if c in (3, 5, 6, 7, 8):
            cell.number_format = "0.0"
        if c == 4:
            cell.number_format = "0.0"

ws.freeze_panes = "A2"


# ============== TAB 5: PERSONAL RECORDS ==============
ws = wb.create_sheet("Personal Records")
ws.column_dimensions["A"].width = 28
ws.column_dimensions["B"].width = 14
ws.column_dimensions["C"].width = 12
ws.column_dimensions["D"].width = 14
ws.column_dimensions["E"].width = 14
ws.column_dimensions["F"].width = 14

set_header_row(ws, 1, ["Exercise", "Best Weight (kg)", "Best Reps @ Weight", "Date Achieved", "Best e1RM (kg)", "Best Single-Set Volume"])

exercises = [
    "Bench Press", "Barbell Row", "Incline DB Press", "Lat Pulldown",
    "Cable Fly", "Triceps Pushdown", "DB Hammer Curl",
    "Back Squat", "Romanian Deadlift", "Leg Press", "Seated Leg Curl", "Standing Calf Raise",
    "Overhead Press", "Weighted Pull-up", "DB Shoulder Press", "Cable Row",
    "Cable Lateral Raise", "Face Pull", "Preacher Curl", "Skull Crusher",
    "Trap Bar Deadlift", "Bulgarian Split Squat", "Hip Thrust", "Walking Lunge", "Seated Calf Raise (Lower B)",
]

LIFT_TAB = "'Lifting Log'"

for i, ex in enumerate(exercises, 2):
    ws.cell(row=i, column=1, value=ex).font = BODY_FONT
    # Best weight ever lifted for this exercise
    ws.cell(row=i, column=2, value=f'=IFERROR(MAXIFS({LIFT_TAB}!G:G,{LIFT_TAB}!D:D,A{i}),"")')
    # Best reps achieved at best weight
    ws.cell(row=i, column=3, value=f'=IFERROR(MAXIFS({LIFT_TAB}!F:F,{LIFT_TAB}!D:D,A{i},{LIFT_TAB}!G:G,B{i}),"")')
    # Date of best
    ws.cell(row=i, column=4, value=f'=IFERROR(MAXIFS({LIFT_TAB}!A:A,{LIFT_TAB}!D:D,A{i},{LIFT_TAB}!G:G,B{i}),"")')
    # Estimated 1RM using Epley: weight * (1 + reps/30)
    ws.cell(row=i, column=5, value=f'=IF(OR(B{i}="",C{i}=""),"",ROUND(B{i}*(1+C{i}/30),1))')
    # Best single-set volume (weight × reps)
    ws.cell(row=i, column=6, value=f'=IFERROR(MAXIFS({LIFT_TAB}!I:I,{LIFT_TAB}!D:D,A{i}),"")')
    for c in range(1, 7):
        ws.cell(row=i, column=c).border = THIN_BORDER
        ws.cell(row=i, column=c).font = BODY_FONT
    ws.cell(row=i, column=2).number_format = "0.0"
    ws.cell(row=i, column=4).number_format = "yyyy-mm-dd"
    ws.cell(row=i, column=5).number_format = "0.0"
    ws.cell(row=i, column=6).number_format = "0"

# Note about how PRs work
note_row = len(exercises) + 4
ws.cell(row=note_row, column=1, value="How this works:").font = Font(name="Arial", bold=True, size=10)
ws.cell(row=note_row + 1, column=1, value="Best Weight = heaviest weight ever lifted for this exercise (any reps).")
ws.cell(row=note_row + 2, column=1, value="Best e1RM = estimated 1RM (Epley formula). Best single comparison metric for progress.")
ws.cell(row=note_row + 3, column=1, value="Best Single-Set Volume = highest weight × reps in any one set.")
ws.cell(row=note_row + 4, column=1, value="Add new exercises in column A as needed — formulas auto-extend down.")
for r in range(note_row, note_row + 5):
    ws.cell(row=r, column=1).font = Font(name="Arial", italic=True, size=10, color="595959")


# ============== TAB 6: WEEKLY SUMMARY ==============
ws = wb.create_sheet("Weekly Summary")
headers = [
    "Week", "Date Range", "Lifting Sessions", "Total Sets", "Total Lift Volume (kg)",
    "Avg Lift RPE", "Runs", "Run km", "Run Time (min)", "Avg Run HR", "Avg Body Weight (kg)"
]
widths = [8, 22, 17, 12, 21, 13, 8, 10, 14, 14, 21]
set_header_row(ws, 1, headers, widths)

# Build 11 weekly rows + a totals row
for week in range(1, 12):
    r = week + 1
    ws.cell(row=r, column=1, value=week)
    # Date range column — concatenate week start + week end
    ws.cell(row=r, column=2, value=f'=TEXT(DATE(2026,4,27)+(A{r}-1)*7,"mmm d")&" – "&TEXT(DATE(2026,4,27)+A{r}*7-1,"mmm d")')

    # Lift sessions = count distinct dates for that week (use COUNTA on unique-ish: count sets / sets per session is messy; use SUMPRODUCT trick)
    # Simpler: count unique dates where week matches. We'll use: number of distinct dates for that week in Lifting Log.
    # Approximation: count distinct dates via SUMPRODUCT(1/COUNTIFS) — but that's expensive. Use COUNTIFS on Set#=1 to count sessions reliably.
    ws.cell(row=r, column=3, value=f'=COUNTIFS({LIFT_TAB}!B:B,A{r},{LIFT_TAB}!E:E,1)')
    # Total sets that week
    ws.cell(row=r, column=4, value=f'=COUNTIFS({LIFT_TAB}!B:B,A{r})')
    # Total volume that week
    ws.cell(row=r, column=5, value=f'=SUMIFS({LIFT_TAB}!I:I,{LIFT_TAB}!B:B,A{r})')
    # Avg RPE that week
    ws.cell(row=r, column=6, value=f'=IFERROR(AVERAGEIFS({LIFT_TAB}!H:H,{LIFT_TAB}!B:B,A{r}),"")')
    # Run sessions
    ws.cell(row=r, column=7, value=f"=COUNTIFS('Running Log'!B:B,A{r})")
    # Run km
    ws.cell(row=r, column=8, value=f"=SUMIFS('Running Log'!C:C,'Running Log'!B:B,A{r})")
    # Run time
    ws.cell(row=r, column=9, value=f"=SUMIFS('Running Log'!D:D,'Running Log'!B:B,A{r})")
    # Avg run HR
    ws.cell(row=r, column=10, value=f"=IFERROR(AVERAGEIFS('Running Log'!F:F,'Running Log'!B:B,A{r}),\"\")")
    # Avg body weight (only count rows where weight was logged)
    ws.cell(row=r, column=11, value=f"=IFERROR(AVERAGEIFS('Body Metrics'!C:C,'Body Metrics'!B:B,A{r}),\"\")")

    for c in range(1, 12):
        ws.cell(row=r, column=c).border = THIN_BORDER
        ws.cell(row=r, column=c).font = BODY_FONT
    ws.cell(row=r, column=5).number_format = "#,##0"
    ws.cell(row=r, column=6).number_format = "0.0"
    ws.cell(row=r, column=8).number_format = "0.00"
    ws.cell(row=r, column=9).number_format = "0"
    ws.cell(row=r, column=10).number_format = "0"
    ws.cell(row=r, column=11).number_format = "0.0"

# Totals row
totals_row = 14
ws.cell(row=totals_row, column=1, value="TOTAL").font = Font(name="Arial", bold=True)
ws.cell(row=totals_row, column=2, value="11 weeks").font = Font(name="Arial", bold=True)
for col_idx in (3, 4, 5, 7, 8, 9):
    col_letter = get_column_letter(col_idx)
    ws.cell(row=totals_row, column=col_idx, value=f'=SUM({col_letter}2:{col_letter}12)')
    ws.cell(row=totals_row, column=col_idx).font = Font(name="Arial", bold=True)
ws.cell(row=totals_row, column=6, value=f'=IFERROR(AVERAGE(F2:F12),"")').font = Font(name="Arial", bold=True)
ws.cell(row=totals_row, column=10, value=f'=IFERROR(AVERAGE(J2:J12),"")').font = Font(name="Arial", bold=True)
ws.cell(row=totals_row, column=11, value=f'=IFERROR(AVERAGE(K2:K12),"")').font = Font(name="Arial", bold=True)

for c in range(1, 12):
    ws.cell(row=totals_row, column=c).fill = SECTION_FILL
    ws.cell(row=totals_row, column=c).border = THIN_BORDER

ws.cell(row=totals_row, column=5).number_format = "#,##0"
ws.cell(row=totals_row, column=6).number_format = "0.0"
ws.cell(row=totals_row, column=8).number_format = "0.00"
ws.cell(row=totals_row, column=9).number_format = "0"
ws.cell(row=totals_row, column=10).number_format = "0"
ws.cell(row=totals_row, column=11).number_format = "0.0"


# ============== TAB 7: REFERENCE ==============
ws = wb.create_sheet("Reference")
ws.column_dimensions["A"].width = 25
ws.column_dimensions["B"].width = 30
ws.column_dimensions["C"].width = 30
ws.column_dimensions["D"].width = 30

row = 1
ws.cell(row=row, column=1, value="Reference Data").font = Font(name="Arial", bold=True, size=14, color="1F4E78")
row += 2

ws.cell(row=row, column=1, value="HR Zones").font = Font(name="Arial", bold=True, size=12)
row += 1
hr_data = [
    ("Zone", "BPM", "Use", ""),
    ("Z1 — Warm-up", "100–119", "Warm-up only", ""),
    ("Z2 — Easy", "120–139", "Easy runs, recovery", ""),
    ("Z3 — Aerobic", "140–159", "Steady running", ""),
    ("Z4 — Threshold", "160–179", "Tempo/threshold reps", ""),
    ("Z5 — Maximum", ">179", "Short hard intervals", ""),
]
for r_data in hr_data:
    for c, v in enumerate(r_data, 1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.font = HEADER_FONT if r_data == hr_data[0] else BODY_FONT
        cell.fill = HEADER_FILL if r_data == hr_data[0] else PatternFill()
        cell.border = THIN_BORDER
    row += 1
row += 1

ws.cell(row=row, column=1, value="Pacing Zones (Outdoor)").font = Font(name="Arial", bold=True, size=12)
row += 1
pace_data = [
    ("Zone", "Pace/km", "Speed (km/h)", "Use"),
    ("🟢 Easy", "6:30–7:00", "8.6–9.2", "Wed + Sat runs"),
    ("🟡 Steady", "6:00–6:30", "9.2–10.0", "Slight progression"),
    ("🟠 Pickup", "5:30–5:50", "10.3–10.9", "Wks 6–7 only"),
]
for r_data in pace_data:
    for c, v in enumerate(r_data, 1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.font = HEADER_FONT if r_data == pace_data[0] else BODY_FONT
        cell.fill = HEADER_FILL if r_data == pace_data[0] else PatternFill()
        cell.border = THIN_BORDER
    row += 1
row += 1

ws.cell(row=row, column=1, value="Weekly Schedule").font = Font(name="Arial", bold=True, size=12)
row += 1
sched_data = [
    ("Day", "Session"),
    ("Monday", "Upper A"),
    ("Tuesday", "Lower A"),
    ("Wednesday", "Easy run"),
    ("Thursday", "Upper B"),
    ("Friday", "Lower B"),
    ("Saturday", "Easy run (slightly longer)"),
    ("Sunday", "Rest"),
]
for r_data in sched_data:
    for c, v in enumerate(r_data, 1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.font = HEADER_FONT if r_data == sched_data[0] else BODY_FONT
        cell.fill = HEADER_FILL if r_data == sched_data[0] else PatternFill()
        cell.border = THIN_BORDER
    row += 1
row += 1

ws.cell(row=row, column=1, value="RPE Reference").font = Font(name="Arial", bold=True, size=12)
row += 1
rpe_data = [
    ("RPE", "Reps in Reserve", "When to use"),
    ("6", "4 reps left", "Warm-up sets"),
    ("7", "3 reps left", "Early working sets, accumulation"),
    ("8", "2 reps left", "Top working sets — most of the block"),
    ("9", "1 rep left", "Peak weeks, isolation movements"),
    ("10", "Failure", "Isolation only — never on compounds"),
]
for r_data in rpe_data:
    for c, v in enumerate(r_data, 1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.font = HEADER_FONT if r_data == rpe_data[0] else BODY_FONT
        cell.fill = HEADER_FILL if r_data == rpe_data[0] else PatternFill()
        cell.border = THIN_BORDER
    row += 1


# ============== Reorder sheets so README is first ==============
wb.move_sheet("README", offset=-wb.sheetnames.index("README"))

wb.save(OUTPUT)
print(f"Saved: {OUTPUT}")
