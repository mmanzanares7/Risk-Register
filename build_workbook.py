"""Build risk_register.xlsx: a formula-driven risk register with a heat map and summary.

Every score, rating, flag, count and ranking in the workbook is an Excel formula.
Only the inputs (blue text) are typed values: risk details, 1-5 scores, rating
thresholds and the report as-of date.

Usage: python3 build_workbook.py ../risk_register.xlsx
"""
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUT = sys.argv[1] if len(sys.argv) > 1 else "risk_register.xlsx"

FONT = "Arial"
BLUE = "0000FF"
NAVY = "1F3864"
GREY = "F2F2F2"
FILL_LOW = PatternFill("solid", start_color="C6EFCE", end_color="C6EFCE")
FILL_MED = PatternFill("solid", start_color="FFEB9C", end_color="FFEB9C")
FILL_HIGH = PatternFill("solid", start_color="F8CBAD", end_color="F8CBAD")
FILL_CRIT = PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")
HEAD_FILL = PatternFill("solid", start_color=NAVY, end_color=NAVY)
HELP_FILL = PatternFill("solid", start_color=GREY, end_color=GREY)
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FIRST, LAST = 5, 54          # register data rows (room for 50 risks)
RNG = lambda col: f"Register!${col}${FIRST}:${col}${LAST}"


def f(bold=False, color="000000", size=10, italic=False):
    return Font(name=FONT, bold=bold, color=color, size=size, italic=italic)


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = f(True, NAVY, 14)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = f(italic=True, color="595959")


def header(ws, row, labels, start_col=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=lab)
        c.font = f(True, "FFFFFF")
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX


def band_formatting(ws, ref):
    """Colour a column of rating words."""
    for word, fill in (("Critical", FILL_CRIT), ("High", FILL_HIGH),
                       ("Medium", FILL_MED), ("Low", FILL_LOW)):
        ws.conditional_formatting.add(ref, CellIsRule(operator="equal", formula=[f'"{word}"'], fill=fill))


wb = Workbook()

# ---------------------------------------------------------------- Legend
lg = wb.active
lg.title = "Legend"
title(lg, "Risk Register: How to Use This Workbook",
      "Formula-driven register built from ITGC testing, a vulnerability assessment and an enterprise risk list")
rows = [
    ("Sheets", None),
    ("Register", "One row per risk. Edit the blue cells. Scores, ratings, flags and rank are formulas."),
    ("Heat Map", "5x5 likelihood and impact grids (inherent and residual) with live risk counts per cell."),
    ("Summary", "Counts by rating, status, source and category, plus the top 5 risks and a chart."),
    ("Scale", "Rating thresholds, report as-of date, and the 1-5 likelihood and impact definitions."),
    ("", None),
    ("Cell colour key", None),
    ("Blue text", "Input. Safe to change."),
    ("Black text", "Formula. Do not overwrite."),
    ("", None),
    ("To add a risk", "Copy the last filled row on Register down one row (formulas included) and edit the blue cells. "
                      "Formulas on Heat Map and Summary read rows 5 to 54 of Register."),
    ("To change the rating bands", "Edit the minimum scores on Scale (B5:B8). Ratings and heat map colours update."),
    ("To change the as-of date", "Edit Scale!B11. Days to Target and the OVERDUE flag update. "
                                 "Use =TODAY() there for a live view."),
    ("", None),
    ("Assumptions", None),
    ("Scoring", "Score = likelihood x impact, each 1 to 5, giving 1 to 25. Same bands as the risk-register-tracker repo."),
    ("Inherent vs residual", "Inherent = before the planned mitigation. Residual = after it is fully in place (a target, not a measurement)."),
    ("Sources", "R-01 to R-04 come from the itgc-controls-testing repo. R-05 to R-08 come from the vulnerability-assessment-report repo. "
                "R-09 to R-15 are the enterprise risks from the risk-register-tracker repo."),
    ("Rescaling", "The vulnerability assessment scored on a 1 to 3 scale. Those risks are rescored here on 1 to 5; "
                  "the original scores are recorded in each description. The new scores are my judgment, not measured data."),
    ("Data", "All data is synthetic and for portfolio demonstration. Dates and owners are illustrative."),
]
r = 4
for k, v in rows:
    if v is None and k:
        lg.cell(row=r, column=1, value=k).font = f(True, NAVY, 11)
    elif k or v:
        lg.cell(row=r, column=1, value=k).font = f(True)
        c = lg.cell(row=r, column=2, value=v)
        c.font = f()
        c.alignment = Alignment(wrap_text=True, vertical="top")
        lg.cell(row=r, column=1).alignment = Alignment(vertical="top")
    r += 1
lg["A11"].font = f(True, BLUE)   # "Blue text"
lg["A12"].font = f(True)
lg.column_dimensions["A"].width = 26
lg.column_dimensions["B"].width = 100
lg.sheet_view.showGridLines = False

# ---------------------------------------------------------------- Scale
sc = wb.create_sheet("Scale")
title(sc, "Scoring Scale and Settings")
header(sc, 4, ["Rating", "Min score", "Meaning"])
bands = [("Low", 1, "Accept or monitor"),
         ("Medium", 5, "Mitigate within the planning cycle"),
         ("High", 10, "Mitigate with an owner and a dated plan"),
         ("Critical", 15, "Escalate to management; mitigate first")]
for i, (name, mn, mean) in enumerate(bands):
    rr = 5 + i
    sc.cell(row=rr, column=1, value=name).font = f(True)
    c = sc.cell(row=rr, column=2, value=mn)
    c.font = f(color=BLUE)
    sc.cell(row=rr, column=3, value=mean).font = f()
    for col in (1, 2, 3):
        sc.cell(row=rr, column=col).border = BOX
sc["A5"].fill, sc["A6"].fill, sc["A7"].fill, sc["A8"].fill = FILL_LOW, FILL_MED, FILL_HIGH, FILL_CRIT
sc["A10"] = "Settings"
sc["A10"].font = f(True, NAVY, 11)
sc["A11"] = "Report as-of date"
sc["A11"].font = f(True)
sc["B11"] = date(2026, 10, 30)
sc["B11"].font = f(color=BLUE)
sc["B11"].number_format = "yyyy-mm-dd"
sc["C11"] = "Input. Fixed so results are reproducible; replace with =TODAY() for a live view."
sc["C11"].font = f(italic=True, color="595959")

header(sc, 14, ["Likelihood", "Level", "Definition"])
lik = [(1, "Rare", "Not expected in the next 12 months"),
       (2, "Unlikely", "Could occur, but no recent indicators"),
       (3, "Possible", "Has occurred in similar organizations or shows some indicators"),
       (4, "Likely", "Weakness is confirmed and exploitable with little effort"),
       (5, "Almost certain", "Exposed now or already occurring")]
imp = [(1, "Minimal", "Negligible cost or disruption"),
       (2, "Minor", "Limited disruption, absorbed in normal operations"),
       (3, "Moderate", "Noticeable disruption or loss of one function or dataset"),
       (4, "Major", "Financial reporting error, data breach or audit finding"),
       (5, "Severe", "Material misstatement, large breach or extended outage")]
for i, (n, name, d) in enumerate(lik):
    for j, v in enumerate((n, name, d)):
        c = sc.cell(row=15 + i, column=1 + j, value=v)
        c.font, c.border = f(), BOX
header(sc, 21, ["Impact", "Level", "Definition"])
for i, (n, name, d) in enumerate(imp):
    for j, v in enumerate((n, name, d)):
        c = sc.cell(row=22 + i, column=1 + j, value=v)
        c.font, c.border = f(), BOX
sc.column_dimensions["A"].width = 20
sc.column_dimensions["B"].width = 16
sc.column_dimensions["C"].width = 70
sc.sheet_view.showGridLines = False

# ---------------------------------------------------------------- Register
rg = wb.create_sheet("Register")
title(rg, "Risk Register", "Blue cells are inputs; black cells are formulas. Rows 5 to 54 feed the other sheets.")
cols = ["Risk ID", "Risk Title", "Category", "Source", "Description",
        "Likelihood (1-5)", "Impact (1-5)", "Inherent Score", "Inherent Rating",
        "Risk Owner", "Treatment", "Mitigation / Response",
        "Residual Likelihood", "Residual Impact", "Residual Score", "Residual Rating",
        "Risk Reduction", "Status", "Target Date", "Days to Target", "Overdue?", "Rank Key (helper)"]
header(rg, 4, cols)
rg.row_dimensions[4].height = 42

ITGC, VA, ENT = "ITGC Testing", "Vulnerability Assessment", "Enterprise Register"
risks = [
    ("R-01", "Terminated users retain active ERP access", "Access Control", ITGC,
     "Control 1 test: U1003 (terminated 2026-05-15) and U1010 (terminated 2026-06-02) were still Active in the finance ERP at the 2026-06-15 audit date.",
     4, 4, "IT Security", "Mitigate",
     "Disable both accounts now; trigger deprovisioning from HR termination events; reconcile HR terminations to active accounts monthly.",
     1, 4, "In Progress", date(2026, 10, 15)),
    ("R-02", "Periodic access reviews overdue", "Access Control", ITGC,
     "Control 2 test: U1010 (last review 2025-11-01) and U1016 (2025-09-01) are past the six-month review cycle.",
     3, 3, "Finance Controller", "Mitigate",
     "Complete the overdue reviews; calendar semi-annual recertification with manager sign-off tracked to completion.",
     1, 3, "In Progress", date(2026, 11, 30)),
    ("R-03", "Production changes without independent approval", "Change Management", ITGC,
     "Control 3 test: CH-2203 and CH-2207 had no approver; CH-2206 was approved (2026-05-18) after it was implemented (2026-05-14); CH-2208 was self-approved.",
     4, 4, "IT Infrastructure", "Mitigate",
     "Block deployment in the ticketing system until an independent approval is logged; prevent requester from approving own change; review the four changes after the fact.",
     2, 4, "Open", date(2026, 12, 15)),
    ("R-04", "AP and Purchasing roles held by one user", "Segregation of Duties", ITGC,
     "Control 4 test: U1016 holds AP Clerk and Purchasing Agent, so one person can create a purchase order and process its payment.",
     3, 4, "Finance Controller", "Mitigate",
     "Remove the Purchasing Agent role, or add supervisor review of all payments this user processes until it is removed.",
     1, 4, "Open", date(2026, 11, 15)),
    ("R-05", "Customer database reachable from the public internet", "Network Security", VA,
     "The customer database server has accepted connections from any public address since launch. Assessment scored data exfiltration 3x3 on a 1-3 scale; rescored 4x5 here.",
     4, 5, "IT Infrastructure", "Mitigate",
     "Restrict inbound access to an allow-list of office and VPN egress addresses; require MFA for administrative access.",
     1, 5, "Open", date(2026, 12, 31)),
    ("R-06", "Automated scanners find the exposed database", "Network Security", VA,
     "Public database ports are scanned continuously and can become a foothold for further attacks. Assessment scored 3x2; rescored 5x3 here.",
     5, 3, "IT Infrastructure", "Mitigate",
     "Same allow-list as R-05; alert on connection attempts from outside the allow-list.",
     1, 3, "Open", date(2026, 12, 31)),
    ("R-07", "Employee alters or deletes customer records", "Data Protection", VA,
     "A malicious or careless employee changes or deletes records, disrupting sales operations. Assessment scored 2x2; rescored 2x3 here.",
     2, 3, "IT Security", "Mitigate",
     "Role-based access with read-only query access for sales staff; audit logging on write operations.",
     1, 3, "Open", date(2027, 1, 31)),
    ("R-08", "Database connections use deprecated SSL", "Data Protection", VA,
     "The assessment recommended moving from SSL to TLS 1.2 or higher on all permitted connections because SSL has known weaknesses.",
     3, 3, "IT Infrastructure", "Mitigate",
     "Disable SSL, enforce TLS 1.2 or higher, and confirm with a configuration scan.",
     1, 3, "Open", date(2026, 12, 15)),
    ("R-09", "Unpatched ERP server vulnerability", "Vulnerability Management", ENT,
     "Critical CVE identified on the production ERP application server; no patch applied yet.",
     4, 5, "IT Infrastructure", "Mitigate",
     "Apply the vendor patch in the next maintenance window; add a compensating WAF rule in the interim.",
     2, 5, "In Progress", date(2026, 11, 20)),
    ("R-10", "No MFA on remote VPN access", "Access Control", ENT,
     "Remote access to the internal network does not require multi-factor authentication.",
     3, 5, "IT Security", "Mitigate",
     "Roll out MFA enforcement for all VPN users.",
     1, 5, "Open", date(2026, 12, 1)),
    ("R-11", "No tested ransomware response runbook", "Incident Response", ENT,
     "The organization has no documented, tested runbook for a ransomware scenario.",
     3, 5, "IT Security", "Mitigate",
     "Draft the runbook and run a tabletop exercise to test it.",
     2, 5, "Open", date(2027, 1, 15)),
    ("R-12", "Shared admin credentials on legacy file server", "Access Control", ENT,
     "A legacy file server uses one shared administrator account instead of individual accounts.",
     4, 3, "IT Infrastructure", "Mitigate",
     "Migrate to individual admin accounts with logging enabled.",
     1, 3, "Open", date(2026, 11, 30)),
    ("R-13", "Backup restoration untested for over a year", "Business Continuity", ENT,
     "The restoration procedure for the financial database had not been tested in over a year.",
     2, 5, "IT Infrastructure", "Mitigate",
     "Run and document a full restoration test.",
     1, 5, "Closed", date(2026, 10, 1)),
    ("R-14", "Payroll data in unencrypted spreadsheets", "Data Protection", ENT,
     "Payroll data was found in unencrypted spreadsheets on a shared drive.",
     3, 4, "Finance Controller", "Mitigate",
     "Move the data to an access-controlled system and encrypt existing files.",
     1, 4, "Open", date(2026, 12, 10)),
    ("R-15", "Critical vendor risk assessments out of date", "Third-Party Risk", ENT,
     "Several critical vendors have not been reassessed in over 18 months.",
     3, 3, "Procurement", "Mitigate",
     "Reassess all critical vendors now and move to a 12-month cycle.",
     2, 3, "Open", date(2027, 2, 1)),
]

for i, (rid, ttl, cat, src, desc, L, I, owner, treat, mit, rL, rI, status, tgt) in enumerate(risks):
    r = FIRST + i
    vals = {1: rid, 2: ttl, 3: cat, 4: src, 5: desc, 6: L, 7: I, 10: owner, 11: treat, 12: mit,
            13: rL, 14: rI, 18: status, 19: tgt}
    for col, v in vals.items():
        c = rg.cell(row=r, column=col, value=v)
        c.font = f(color=BLUE)
    rg.cell(row=r, column=8, value=f"=F{r}*G{r}")
    rg.cell(row=r, column=9, value=(f'=IF(H{r}>=Scale!$B$8,"Critical",IF(H{r}>=Scale!$B$7,"High",'
                                    f'IF(H{r}>=Scale!$B$6,"Medium","Low")))'))
    rg.cell(row=r, column=15, value=f"=M{r}*N{r}")
    rg.cell(row=r, column=16, value=(f'=IF(O{r}>=Scale!$B$8,"Critical",IF(O{r}>=Scale!$B$7,"High",'
                                     f'IF(O{r}>=Scale!$B$6,"Medium","Low")))'))
    rg.cell(row=r, column=17, value=f"=IF(H{r}=0,0,(H{r}-O{r})/H{r})")
    rg.cell(row=r, column=20, value=f'=IF(R{r}="Closed","",S{r}-Scale!$B$11)')
    rg.cell(row=r, column=21, value=f'=IF(AND(R{r}<>"Closed",S{r}<Scale!$B$11),"OVERDUE","")')
    rg.cell(row=r, column=22, value=f"=H{r}+(1000-ROW())/100000")
    for col in (8, 9, 15, 16, 17, 20, 21, 22):
        rg.cell(row=r, column=col).font = f(bold=col in (8, 15))
    rg.cell(row=r, column=17).number_format = "0%"
    rg.cell(row=r, column=19).number_format = "yyyy-mm-dd"
    rg.cell(row=r, column=22).number_format = "0.00000"
    rg.cell(row=r, column=22).font = f(color="7F7F7F")
    for col in range(1, 23):
        c = rg.cell(row=r, column=col)
        c.border = BOX
        c.alignment = Alignment(vertical="top", wrap_text=col in (2, 3, 4, 5, 10, 12),
                                horizontal="center" if col in (1, 6, 7, 8, 9, 13, 14, 15, 16, 17, 19, 20, 21) else None)
    rg.row_dimensions[r].height = 64

widths = [8, 30, 20, 22, 48, 11, 9, 10, 11, 18, 11, 46, 11, 10, 10, 11, 10, 12, 12, 10, 11, 11]
for i, w in enumerate(widths, 1):
    rg.column_dimensions[get_column_letter(i)].width = w
rg.freeze_panes = "C5"
rg.auto_filter.ref = f"A4:V{FIRST + len(risks) - 1}"

for ref in (f"I{FIRST}:I{LAST}", f"P{FIRST}:P{LAST}"):
    band_formatting(rg, ref)
rg.conditional_formatting.add(f"U{FIRST}:U{LAST}",
                              CellIsRule(operator="equal", formula=['"OVERDUE"'], fill=FILL_CRIT,
                                         font=Font(name=FONT, bold=True, color="9C0006")))

dv_score = DataValidation(type="whole", operator="between", formula1="1", formula2="5", allow_blank=True,
                          showErrorMessage=True, errorTitle="Score", error="Enter a whole number from 1 to 5.")
dv_status = DataValidation(type="list", formula1='"Open,In Progress,Closed"', allow_blank=True)
dv_treat = DataValidation(type="list", formula1='"Mitigate,Accept,Transfer,Avoid"', allow_blank=True)
for dv in (dv_score, dv_status, dv_treat):
    rg.add_data_validation(dv)
dv_score.add(f"F{FIRST}:G{LAST}")
dv_score.add(f"M{FIRST}:N{LAST}")
dv_status.add(f"R{FIRST}:R{LAST}")
dv_treat.add(f"K{FIRST}:K{LAST}")

# ---------------------------------------------------------------- Heat Map
hm = wb.create_sheet("Heat Map")
title(hm, "Risk Heat Maps", "Each cell counts the risks at that likelihood and impact. Colour = rating band for that cell.")


def grid(top, label, lcol, icol):
    hm.cell(row=top, column=1, value=label).font = f(True, NAVY, 11)
    hm.cell(row=top + 1, column=3, value="Likelihood").font = f(True)
    hm.merge_cells(start_row=top + 1, start_column=3, end_row=top + 1, end_column=7)
    hm.cell(row=top + 1, column=3).alignment = Alignment(horizontal="center")
    for j in range(5):
        c = hm.cell(row=top + 2, column=3 + j, value=j + 1)
        c.font, c.fill, c.border = f(True), HELP_FILL, BOX
        c.alignment = Alignment(horizontal="center")
    lab = hm.cell(row=top + 3, column=1, value="Impact")
    lab.font = f(True)
    for k in range(5):
        r = top + 3 + k
        c = hm.cell(row=r, column=2, value=5 - k)
        c.font, c.fill, c.border = f(True), HELP_FILL, BOX
        c.alignment = Alignment(horizontal="center", vertical="center")
        hm.row_dimensions[r].height = 40
        for j in range(5):
            colL = get_column_letter(3 + j)
            cell = hm.cell(row=r, column=3 + j,
                           value=f"=COUNTIFS({RNG(lcol)},{colL}${top + 2},{RNG(icol)},$B{r})")
            cell.font = f(True, size=12)
            cell.number_format = "0;-0;"
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BOX
    area = f"C{top + 3}:G{top + 7}"
    s = f"C${top + 2}*$B{top + 3}"
    for thr, fill in ((8, FILL_CRIT), (7, FILL_HIGH), (6, FILL_MED)):
        pass
    hm.conditional_formatting.add(area, FormulaRule(formula=[f"{s}>=Scale!$B$8"], fill=FILL_CRIT, stopIfTrue=True))
    hm.conditional_formatting.add(area, FormulaRule(formula=[f"{s}>=Scale!$B$7"], fill=FILL_HIGH, stopIfTrue=True))
    hm.conditional_formatting.add(area, FormulaRule(formula=[f"{s}>=Scale!$B$6"], fill=FILL_MED, stopIfTrue=True))
    hm.conditional_formatting.add(area, FormulaRule(formula=[f"{s}>=Scale!$B$5"], fill=FILL_LOW, stopIfTrue=True))
    hm.cell(row=top + 8, column=2, value="Total").font = f(True)
    t = hm.cell(row=top + 8, column=3, value=f"=SUM(C{top + 3}:G{top + 7})")
    t.font = f(True)


grid(4, "Inherent risk (before mitigation)", "F", "G")
grid(16, "Residual risk (after planned mitigation)", "M", "N")
hm.column_dimensions["A"].width = 10
hm.column_dimensions["B"].width = 10
for c in "CDEFG":
    hm.column_dimensions[c].width = 14
hm.sheet_view.showGridLines = False

# ---------------------------------------------------------------- Summary
sm = wb.create_sheet("Summary")
title(sm, "Risk Summary", "All figures are formulas over the Register sheet.")

sm["A4"] = "Risks by rating"
sm["A4"].font = f(True, NAVY, 11)
header(sm, 5, ["Rating", "Inherent", "Residual"])
for i, name in enumerate(["Critical", "High", "Medium", "Low"]):
    r = 6 + i
    sm.cell(row=r, column=1, value=name).font = f(True)
    sm.cell(row=r, column=2, value=f"=COUNTIF({RNG('I')},A{r})")
    sm.cell(row=r, column=3, value=f"=COUNTIF({RNG('P')},A{r})")
    for col in (1, 2, 3):
        sm.cell(row=r, column=col).border = BOX
        if col > 1:
            sm.cell(row=r, column=col).font = f()
            sm.cell(row=r, column=col).alignment = Alignment(horizontal="center")
band_formatting(sm, "A6:A9")
sm["A10"] = "Total"
sm["A10"].font = f(True)
sm["B10"] = "=SUM(B6:B9)"
sm["C10"] = "=SUM(C6:C9)"
for c in ("B10", "C10"):
    sm[c].font = f(True)
    sm[c].alignment = Alignment(horizontal="center")

sm["E4"] = "Key figures"
sm["E4"].font = f(True, NAVY, 11)
kf = [
    ("Average inherent score", f"=AVERAGE({RNG('H')})", "0.0"),
    ("Average residual score", f"=AVERAGE({RNG('O')})", "0.0"),
    ("Average risk reduction", f"=AVERAGE({RNG('Q')})", "0%"),
    ("Open or in-progress risks", f'=COUNTIF({RNG("R")},"Open")+COUNTIF({RNG("R")},"In Progress")', "0"),
    ("Overdue mitigations", f'=COUNTIF({RNG("U")},"OVERDUE")', "0"),
]
for i, (lab, fm, nf) in enumerate(kf):
    r = 5 + i
    sm.cell(row=r, column=5, value=lab).font = f(True)
    c = sm.cell(row=r, column=7, value=fm)
    c.font, c.number_format = f(), nf
    c.alignment = Alignment(horizontal="center")
    for col in (5, 6, 7):
        sm.cell(row=r, column=col).border = BOX

sm["A12"] = "Risks by status"
sm["A12"].font = f(True, NAVY, 11)
header(sm, 13, ["Status", "Count"])
for i, name in enumerate(["Open", "In Progress", "Closed"]):
    r = 14 + i
    sm.cell(row=r, column=1, value=name).font = f(True)
    c = sm.cell(row=r, column=2, value=f"=COUNTIF({RNG('R')},A{r})")
    c.font, c.alignment = f(), Alignment(horizontal="center")
    sm.cell(row=r, column=1).border = sm.cell(row=r, column=2).border = BOX

sm["E12"] = "Risks by source"
sm["E12"].font = f(True, NAVY, 11)
header(sm, 13, ["Source", "", "Count"], start_col=5)
for i, name in enumerate([ITGC, VA, ENT]):
    r = 14 + i
    sm.cell(row=r, column=5, value=name).font = f(True)
    c = sm.cell(row=r, column=7, value=f"=COUNTIF({RNG('D')},E{r})")
    c.font, c.alignment = f(), Alignment(horizontal="center")
    for col in (5, 6, 7):
        sm.cell(row=r, column=col).border = BOX

sm["A19"] = "Risks by category"
sm["A19"].font = f(True, NAVY, 11)
header(sm, 20, ["Category", "Count", "Total inherent score"])
cats = ["Access Control", "Change Management", "Segregation of Duties", "Network Security",
        "Data Protection", "Vulnerability Management", "Incident Response",
        "Business Continuity", "Third-Party Risk"]
for i, name in enumerate(cats):
    r = 21 + i
    sm.cell(row=r, column=1, value=name).font = f(True)
    sm.cell(row=r, column=2, value=f"=COUNTIF({RNG('C')},A{r})")
    sm.cell(row=r, column=3, value=f"=SUMIF({RNG('C')},A{r},{RNG('H')})")
    for col in (1, 2, 3):
        sm.cell(row=r, column=col).border = BOX
        if col > 1:
            sm.cell(row=r, column=col).font = f()
            sm.cell(row=r, column=col).alignment = Alignment(horizontal="center")
end = 21 + len(cats)
sm.cell(row=end, column=1, value="Total").font = f(True)
sm.cell(row=end, column=2, value=f"=SUM(B21:B{end - 1})").font = f(True)
sm.cell(row=end, column=3, value=f"=SUM(C21:C{end - 1})").font = f(True)
sm.cell(row=end, column=2).alignment = sm.cell(row=end, column=3).alignment = Alignment(horizontal="center")

top0 = end + 3
sm.cell(row=top0 - 1, column=1, value="Top 5 risks by inherent score").font = f(True, NAVY, 11)
header(sm, top0, ["Rank", "Risk ID", "Title", "Inherent", "Rating", "Residual", "Owner", "Status"])
for n in range(1, 6):
    r = top0 + n
    m = f"MATCH(LARGE({RNG('V')},{n}),{RNG('V')},0)"
    sm.cell(row=r, column=1, value=n)
    sm.cell(row=r, column=2, value=f"=INDEX({RNG('A')},{m})")
    sm.cell(row=r, column=3, value=f"=INDEX({RNG('B')},{m})")
    sm.cell(row=r, column=4, value=f"=INDEX({RNG('H')},{m})")
    sm.cell(row=r, column=5, value=f"=INDEX({RNG('I')},{m})")
    sm.cell(row=r, column=6, value=f"=INDEX({RNG('O')},{m})")
    sm.cell(row=r, column=7, value=f"=INDEX({RNG('J')},{m})")
    sm.cell(row=r, column=8, value=f"=INDEX({RNG('R')},{m})")
    for col in range(1, 9):
        c = sm.cell(row=r, column=col)
        c.font, c.border = f(), BOX
        c.alignment = Alignment(horizontal="center" if col != 3 else None, wrap_text=(col == 3))
band_formatting(sm, f"E{top0 + 1}:E{top0 + 5}")

for col, w in zip("ABCDEFGH", (24, 12, 40, 12, 24, 12, 20, 14)):
    sm.column_dimensions[col].width = w
sm.sheet_view.showGridLines = False

ch = BarChart()
ch.type = "col"
ch.grouping = "clustered"
ch.title = "Risks by rating: inherent vs residual"
ch.y_axis.title = "Number of risks"
ch.y_axis.majorGridlines = None
ch.add_data(Reference(sm, min_col=2, max_col=3, min_row=5, max_row=9), titles_from_data=True)
ch.set_categories(Reference(sm, min_col=1, min_row=6, max_row=9))
ch.series[0].graphicalProperties.solidFill = "C00000"
ch.series[1].graphicalProperties.solidFill = "2E75B6"
ch.height, ch.width = 7.5, 14
ch.x_axis.delete = False
ch.y_axis.delete = False
sm.add_chart(ch, "J4")

wb.move_sheet("Legend", offset=0)
wb._sheets = [wb["Legend"], wb["Register"], wb["Heat Map"], wb["Summary"], wb["Scale"]]
wb.active = 1
wb.save(OUT)
print("saved", OUT)
