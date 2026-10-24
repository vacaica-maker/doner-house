import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUT = "IFE_Matrix_Doner_Street.xlsx"

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
GROUP_FILL = PatternFill("solid", fgColor="D9E1F2")
TOTAL_FILL = PatternFill("solid", fgColor="F2F2F2")
WHITE_BOLD = Font(bold=True, color="FFFFFF")
BOLD = Font(bold=True)

STRENGTHS = [
    ("Fresh-to-order assembly line — fries fried per order, not pre-batched (≈10 min per order, Medium Combo, 3.6.1) — with ~168 units/day baseline capacity (3.4.2), enough to cover planned demand in every scenario", 0.13, 4),
    ("Store on the Calle 72 corridor next to the new Metro Line 1 and TransMilenio station, a node under redevelopment (Decreto 242/2026) where none of the three rivals is located (2.2.1, 2.5)", 0.10, 3),
    ("Dual quality-control checkpoints on the production line — an incoming Receiving & Quality Check on every delivery plus an outgoing per-order Quality Check before packaging — a documented consistency/food-safety control none of the three budget-comparable competitors shows evidence of (2.4.6, 3.6.1)", 0.08, 3),
    ("Lean processes around one core product allow fast onboarding, as associates need no degree and learn every station through on-the-job cross-training (3.3.2)", 0.07, 3),
    ("Standardized product with defined portions, costs and assembly steps makes the concept easy to replicate at further locations (3.4.2)", 0.07, 3),
    ("Structured, criteria-based hiring and onboarding for Store Managers — scored CV rubric, a practical P&L/leadership assessment, and a 30/60/90-day performance review — builds management quality in from day one (3.3.1, 3.3.2)", 0.06, 3),
]

WEAKNESSES = [
    ("Unknown brand with no reviews yet vs. Filiák's 4.8 rating (2.5.3)", 0.13, 1),
    ("No proprietary recipe, the product follows a standard döner recipe", 0.10, 2),
    ("No operating track record, all sales and marketing figures are forecasts (3.4.5)", 0.09, 2),
    ("No own delivery service, so delivery orders depend on aggregator apps charging 18% to 30% commission per order (2.4.1, 3.4.4)", 0.07, 2),
    ("Single Store Manager per store with no assistant, so days off or illness leave shift-level staff handling escalations meant for the store manager (3.3.2, 3.4.3)", 0.05, 2),
    ("100% equity-funded with no external line of credit, so the pessimistic scenario's cash shortfall has no financing backstop and would require an emergency owner capital injection (3.5.2)", 0.05, 2),
]

INTERPRETATION = (
    "With a total weighted score of 2.51, Döner Street's internal position is just above the average "
    "of 2.5. Its main strengths are fast service, a location on the redeveloped Calle 72 corridor, dual "
    "incoming/outgoing quality-control checkpoints, a lean and easily replicable concept, and a structured "
    "hiring process for Store Managers. They are offset by the typical weaknesses of a new business, above "
    "all an unknown brand, no proprietary recipe, and a concentrated marketing/finance risk profile -- a "
    "single Store Manager with no backup and no external financing cushion -- on top of the same "
    "aggregator-commission weakness that drives the delivery commission line in the 3.5.2 Estado de "
    "Resultados."
)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "IFE Matrix"

ws.column_dimensions["A"].width = 5.0
ws.column_dimensions["B"].width = 80.0
ws.column_dimensions["C"].width = 14.0
ws.column_dimensions["D"].width = 10.0
ws.column_dimensions["E"].width = 16.0

ws["A1"] = "3.9 Internal Factor Evaluation (IFE) Matrix · Döner Street"
ws["A1"].font = BOLD
ws.merge_cells("A1:E1")

ws["A2"] = ("Weights: 0.0 (not important) to 1.0 (all-important), must sum to 1.00. Ratings: 1 = major "
            "weakness, 2 = minor weakness, 3 = minor strength, 4 = major strength (David, Ch. 4). Blue cells "
            "are inputs.")
ws.merge_cells("A2:E2")

headers = ["#", "Key Internal Factors", "Weight", "Rating", "Weighted Score"]
for col, h in enumerate(headers, start=1):
    c = ws.cell(row=4, column=col, value=h)
    c.font = WHITE_BOLD
    c.fill = HEADER_FILL

row = 5
ws.cell(row=row, column=1, value="Strengths").font = BOLD
for col in range(1, 6):
    ws.cell(row=row, column=col).fill = GROUP_FILL
ws.merge_cells(f"A{row}:E{row}")
row += 1

strengths_start = row
for i, (text, weight, rating) in enumerate(STRENGTHS, start=1):
    ws.cell(row=row, column=1, value=i)
    ws.cell(row=row, column=2, value=text)
    wc = ws.cell(row=row, column=3, value=weight)
    wc.number_format = "0.00"
    ws.cell(row=row, column=4, value=rating)
    sc = ws.cell(row=row, column=5, value=f"=C{row}*D{row}")
    sc.number_format = "0.00"
    row += 1
strengths_end = row - 1

ws.cell(row=row, column=1, value="Weaknesses").font = BOLD
for col in range(1, 6):
    ws.cell(row=row, column=col).fill = GROUP_FILL
ws.merge_cells(f"A{row}:E{row}")
row += 1

weaknesses_start = row
for i, (text, weight, rating) in enumerate(WEAKNESSES, start=1):
    ws.cell(row=row, column=1, value=i)
    ws.cell(row=row, column=2, value=text)
    wc = ws.cell(row=row, column=3, value=weight)
    wc.number_format = "0.00"
    ws.cell(row=row, column=4, value=rating)
    sc = ws.cell(row=row, column=5, value=f"=C{row}*D{row}")
    sc.number_format = "0.00"
    row += 1
weaknesses_end = row - 1

total_row = row
ws.cell(row=total_row, column=2, value="Total").font = BOLD
for col in [1, 2, 3, 4, 5]:
    ws.cell(row=total_row, column=col).fill = TOTAL_FILL
tc = ws.cell(row=total_row, column=3, value=f"=SUM(C{strengths_start}:C{weaknesses_end})")
tc.font = BOLD
tc.number_format = "0.00"
ws.cell(row=total_row, column=4).font = BOLD
ec = ws.cell(row=total_row, column=5, value=f"=SUM(E{strengths_start}:E{weaknesses_end})")
ec.font = BOLD
ec.number_format = "0.00"

row = total_row + 1
ws.cell(row=row, column=2,
        value=f'=IF(ROUND(C{total_row},2)=1,"Check: weights sum to 1.00","Check: weights must sum to 1.00")')
row += 1
ws.cell(row=row, column=2,
        value=(f'=IF(E{total_row}>=2.5,"Score at or above the 2.5 average: internal position average to '
               f'strong","Score below the 2.5 average: internal position slightly weak")'))
row += 2

ws.cell(row=row, column=1, value="Interpretation").font = BOLD
ws.merge_cells(f"A{row}:E{row}")
row += 1
ws.cell(row=row, column=1, value=INTERPRETATION)
ws.merge_cells(f"A{row}:E{row}")
ws.cell(row=row, column=1).alignment = Alignment(wrap_text=True, vertical="top")
row += 1
ws.cell(row=row, column=1,
        value="Source: Fred R. David, Strategic Management, Ch. 4, IFE Matrix; section numbers refer to the published site")
ws.merge_cells(f"A{row}:E{row}")

wb.save(OUT)
print("Saved", OUT)
