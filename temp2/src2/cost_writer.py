import logging
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

log = logging.getLogger("remittance")
MONEY_FMT = '"$"#,##0.00'


def write_cost_sheet(path, sheet_name, cost_df, title, thru_date):
    if thru_date.upper() == "AUTO":
        thru_date = datetime.now().strftime("%m-%d-%Y")

    wb = load_workbook(path)
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    bold = Font(bold=True)

    ws["B1"] = title
    ws["B1"].font = bold
    ws["B2"] = f"Costs Thru {thru_date}"
    ws["B2"].font = bold

    ws["B5"] = "Total"
    ws["B5"].font = bold
    ws["D5"] = round(float(cost_df["AMOUNT"].sum()), 2)
    ws["D5"].number_format = MONEY_FMT
    ws["D5"].font = bold

    header_row = 7
    for col, name in enumerate(cost_df.columns, start=1):
        ws.cell(row=header_row, column=col, value=name).font = bold

    for r, row in enumerate(cost_df.itertuples(index=False), start=header_row + 1):
        for c, value in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=value)
            if c == 4:
                cell.number_format = MONEY_FMT

    for i, w in enumerate([20, 45, 12, 12, 32, 18], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(path)
    log.info("Sheet '%s' written with %d rows", sheet_name, len(cost_df))