import shutil
from datetime import date
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font

PAYMENT_SHEET = "Payments"          # must match the sheet name in your template
COST_SHEET = "Cost"                 # must match the sheet name in your template

HEADER_ROW = 7                      # used when the header row can't be found in the sheet
FIRST_HEADER = "DEFENDANT'S NAME"   # text used to locate the header row

PAYMENT_CURRENCY_COLS = ["AMOUNT PAID", "DIRECT PAYMENT", "CLIENT SHARE", "BW, LLP SHARE"]
COST_CURRENCY_COLS = ["AMOUNT"]
COST_DATE_COLS = ["Date"]

CURRENCY_FORMAT = '"$"#,##0.00'
DATE_FORMAT = "mm/dd/yyyy"


def thru_date():
    """Date shown in 'Payments Thru ...' / 'Costs Thru ...'."""
    return date.today().strftime("%m-%d-%Y")


def _pct_format(p):
    return "0%" if abs(p * 100 - round(p * 100)) < 1e-9 else "0.00%"


def prepare_output(template_file, output_file, log=None):
    """Copy the template to the output path (a fresh copy every run)."""
    template = Path(str(template_file).strip().strip('"'))
    if not template.exists():
        raise FileNotFoundError(f"Output template not found: {template}")
    output_file = Path(output_file)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(template, output_file)
    if log:
        log.info("Template copied to %s", output_file.resolve())
    return output_file


def _get_sheet(wb, name):
    for ws in wb.worksheets:
        if ws.title.strip().lower() == name.lower():
            return ws, False
    return wb.create_sheet(name), True


def _find_header_row(ws):
    for r in range(1, 21):
        for c in ws[r]:
            if c.value is not None and str(c.value).strip().upper() == FIRST_HEADER:
                return r
    return HEADER_ROW


def _write_table(ws, df, header_row):
    # 1) unmerge every merged range that reaches below the header row (e.g. E7:E8 'CLIENT / SHARE')
    for rng in list(ws.merged_cells.ranges):
        if rng.max_row > header_row:
            ws.unmerge_cells(str(rng))

    # 2) clear everything below the header row (this removes the old 2nd header line 'SHARE')
    if ws.max_row > header_row:
        ws.delete_rows(header_row + 1, ws.max_row - header_row)

    # 3) one-line header in row 7 (overwrites 'CLIENT' with 'CLIENT SHARE')
    for j, col in enumerate(df.columns, start=1):
        ws.cell(row=header_row, column=j, value=col)

    # 4) data starts in row 8
    clean = df.astype(object).where(df.notna(), None)
    for i, row in enumerate(clean.itertuples(index=False), start=header_row + 1):
        for j, value in enumerate(row, start=1):
            ws.cell(row=i, column=j, value=value)


def _autofit(ws, header_row):
    for col in ws.iter_cols(min_row=header_row):
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = width + 3


def _format_table(ws, header_row, currency_cols, date_cols):
    headers = {str(c.value).strip(): c.column for c in ws[header_row] if c.value is not None}
    for names, fmt in ((currency_cols, CURRENCY_FORMAT), (date_cols, DATE_FORMAT)):
        for name in names:
            col = headers.get(name)
            if col:
                for r in range(header_row + 1, ws.max_row + 1):
                    ws.cell(row=r, column=col).number_format = fmt


def _payment_block(ws, header_row, title, client_pct, bw_pct):
    first, last = header_row + 1, max(ws.max_row, header_row + 1)
    ws["A1"] = title
    ws["A2"] = f"Payments Thru {thru_date()}"
    ws["E4"], ws["F4"] = client_pct, bw_pct
    ws["E4"].number_format = _pct_format(client_pct)
    ws["F4"].number_format = _pct_format(bw_pct)
    ws["B5"] = "Total"
    for col in "CD":                                   # Amount Paid, Direct Payment: sum of the rows
        ws[f"{col}5"] = f"=SUM({col}{first}:{col}{last})"
    ws["E5"] = "=ROUND(C5*E4,2)"                       # client total = Amount Paid total x client %
    ws["F5"] = "=ROUND(C5*F4,2)"                       # BW total = Amount Paid total x BW %
    for col in "CDEF":
        ws[f"{col}5"].number_format = CURRENCY_FORMAT
    for ref in ("A1", "A2", "E4", "F4", "B5", "C5", "D5", "E5", "F5"):
        ws[ref].font = Font(bold=True)


def _cost_block(ws, header_row, title):
    first, last = header_row + 1, max(ws.max_row, header_row + 1)
    ws["B1"] = title
    ws["B2"] = f"Costs Thru {thru_date()}"
    ws["B5"] = "Total"
    ws["C5"] = f"=SUM(C{first}:C{last})"
    ws["C5"].number_format = CURRENCY_FORMAT
    for ref in ("B1", "B2", "B5", "C5"):
        ws[ref].font = Font(bold=True)


def _write_sheet(output_file, sheet_name, df, fill_block, currency_cols, date_cols):
    wb = load_workbook(output_file)
    ws, is_new = _get_sheet(wb, sheet_name)
    header_row = _find_header_row(ws)

    _write_table(ws, df, header_row)
    fill_block(ws, header_row)
    _format_table(ws, header_row, currency_cols, date_cols)
    if is_new:                                         # keep the template's own column widths
        _autofit(ws, header_row)

    wb.save(output_file)
    wb.close()


def write_payment_sheet(df, output_file, log, title, client_pct, bw_pct):
    _write_sheet(output_file, PAYMENT_SHEET, df,
                 lambda ws, h: _payment_block(ws, h, title, client_pct, bw_pct),
                 PAYMENT_CURRENCY_COLS, [])
    if log:
        log.info("Payment sheet: %d rows written", len(df))


def write_cost_sheet(df, output_file, log, title):
    _write_sheet(output_file, COST_SHEET, df,
                 lambda ws, h: _cost_block(ws, h, title),
                 COST_CURRENCY_COLS, COST_DATE_COLS)
    if log:
        log.info("Cost sheet: %d rows written", len(df))