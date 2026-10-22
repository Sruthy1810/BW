import re
import shutil
from copy import copy
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from statement_logo import ensure_statement_logo
from temp1.src1.t1_reader import excel_engine

PAYMENT_SHEET = "Payments"          # must match the sheet name in your template
COST_SHEET = "Cost"                 # must match the sheet name in your template

HEADER_ROW = 7                      # used when the header row can't be found in the sheet
FIRST_HEADER = "DEFENDANT'S NAME"   # text used to locate the header row

PAYMENT_CURRENCY_COLS = ["AMOUNT PAID", "DIRECT PAYMENT"]
COST_CURRENCY_COLS = ["AMOUNT"]
COST_DATE_COLS = ["Date"]

CURRENCY_FORMAT = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
SHARE_FORMAT = CURRENCY_FORMAT
DATE_FORMAT = "mm/dd/yyyy"

# ---- address on the statement sheets ----
ADDRESS_START_TEXT = "remittance statement"        # line that starts the address block (not included)
ADDRESS_END_TEXT = "collection receipts section"   # line that ends it (not included)
STATEMENT_SHEET_KEYS = ("statement", "summary")    # ASSUMPTION: sheets whose name contains one of these
ADDRESS_CELLS = ("A10", "A11", "A12")
THRU_CELL = "A16"


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


def _format_currency_cell(cell):
    cell.number_format = CURRENCY_FORMAT
    alignment = copy(cell.alignment)
    alignment.horizontal = "right"
    cell.alignment = alignment


def _format_table(ws, header_row, currency_cols, date_cols):
    headers = {str(c.value).strip(): c.column for c in ws[header_row] if c.value is not None}
    formats = (
        (currency_cols, CURRENCY_FORMAT),
        (("CLIENT SHARE", "BW, LLP SHARE"), SHARE_FORMAT),
        (date_cols, DATE_FORMAT),
    )
    for names, fmt in formats:
        for name in names:
            col = headers.get(name)
            if col:
                for r in range(header_row + 1, ws.max_row + 1):
                    cell = ws.cell(row=r, column=col)
                    if fmt in (CURRENCY_FORMAT, SHARE_FORMAT):
                        _format_currency_cell(cell)
                    else:
                        cell.number_format = fmt


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
    ws["E5"] = f"=SUM(E{first}:E{last})"
    ws["F5"] = f"=SUM(F{first}:F{last})"
    for col in "CD":
        _format_currency_cell(ws[f"{col}5"])
    for col in "EF":
        _format_currency_cell(ws[f"{col}5"])
    for ref in ("A1", "A2", "E4", "F4", "B5", "C5", "D5", "E5", "F5"):
        ws[ref].font = Font(bold=True)


def _cost_block(ws, header_row, title):
    first, last = header_row + 1, max(ws.max_row, header_row + 1)
    ws["B1"] = title
    ws["B2"] = f"Costs Thru {thru_date()}"
    ws["B5"] = "Total"
    ws["C5"] = f"=SUM(C{first}:C{last})"
    _format_currency_cell(ws["C5"])
    for ref in ("B1", "B2", "B5", "C5"):
        ws[ref].font = Font(bold=True)


def _write_sheet(output_file, sheet_name, df, fill_block, currency_cols, date_cols):
    wb = load_workbook(output_file)
    ws, is_new = _get_sheet(wb, sheet_name)
    header_row = _find_header_row(ws)

    _write_table(ws, df, header_row)
    fill_block(ws, header_row)
    _format_table(ws, header_row, currency_cols, date_cols)
    ensure_statement_logo(ws)
    if is_new:                                         # keep the template's own column widths
        _autofit(ws, header_row)

    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
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


def format_payment_share_precision(output_file):
    """Apply three-decimal display formatting to Temp1 share details and totals."""
    wb = load_workbook(output_file)
    ws = next(
        (sheet for sheet in wb.worksheets
         if sheet.title.strip().lower() == PAYMENT_SHEET.lower()),
        None,
    )
    if ws is None:
        wb.close()
        return

    header_row = _find_header_row(ws)
    for column in ("E", "F"):
        ws.column_dimensions[column].width = max(ws.column_dimensions[column].width or 0, 14)
        for row in range(1, ws.max_row + 1):
            cell = ws[f"{column}{row}"]
            alignment = copy(cell.alignment)
            alignment.shrink_to_fit = False
            cell.alignment = alignment
            if row == 5 or row > header_row:
                _format_currency_cell(cell)

    wb.save(output_file)
    wb.close()


def remove_empty_data_sheets(output_file, log=None):
    """Remove empty/all-zero data sheets without deleting Summary sheets."""
    wb = load_workbook(output_file)
    sheet_specs = {
        PAYMENT_SHEET.lower(): (None, (3, 4, 5, 6)),
        COST_SHEET.lower(): (None, (3,)),
        "intonly": (8, (3,)),
    }
    removed = []

    for ws in list(wb.worksheets):
        if "summary" in ws.title.lower():
            continue
        spec = sheet_specs.get(ws.title.strip().lower())
        if spec is None:
            continue

        data_start, amount_columns = spec
        if data_start is None:
            data_start = _find_header_row(ws) + 1
        data_rows = [
            row for row in ws.iter_rows(min_row=data_start, max_col=ws.max_column)
            if any(cell.value is not None and str(cell.value).strip() for cell in row)
        ]
        if data_rows:
            has_nonzero = False
            for row in data_rows:
                for column in amount_columns:
                    value = row[column - 1].value if column <= len(row) else None
                    if value is None or (isinstance(value, str) and not value.strip()):
                        continue
                    if isinstance(value, str) and value.startswith("="):
                        has_nonzero = True
                        break
                    try:
                        amount = Decimal(str(value).replace("$", "").replace(",", "").strip())
                    except InvalidOperation:
                        has_nonzero = True
                        break
                    if amount != 0:
                        has_nonzero = True
                        break
                if has_nonzero:
                    break
            if has_nonzero:
                continue

        removed.append(ws.title)
        wb.remove(ws)

    for sheet_name in removed:
        reference = re.compile(
            rf"(?<![A-Za-z0-9_])(?:'{re.escape(sheet_name)}'|{re.escape(sheet_name)})!",
            re.IGNORECASE,
        )
        for ws in wb.worksheets:
            if "summary" not in ws.title.lower():
                continue
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("=") and reference.search(cell.value):
                        cell.value = "=0"
                        if log:
                            log.info("Summary formula %s!%s set to zero after removing '%s'",
                                     ws.title, cell.coordinate, sheet_name)

    if removed:
        wb.save(output_file)
    wb.close()
    if log and removed:
        log.info("Removed empty/all-zero Temp1 sheets: %s", removed)


# =====================================================================
# Address on the statement sheets (rows 10-12) + "Statement Thru" (row 16)
# =====================================================================
def _text(value):
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _row_text(row):
    return " ".join(_text(v) for v in row.tolist() if pd.notna(v) and _text(v))


def read_address(input_file, log=None):
    """Lines between 'Remittance Statement ...' and 'Collection Receipts Section', as exactly 3 lines."""
    sheets = pd.read_excel(input_file, sheet_name=None, header=None, dtype=object,
                           engine=excel_engine(input_file))
    for sheet_name, df in sheets.items():
        start = None
        for pos in range(len(df)):
            text = _row_text(df.iloc[pos])
            if start is None:
                if ADDRESS_START_TEXT in text.lower():
                    start = pos
            elif ADDRESS_END_TEXT in text.lower():
                lines = [_row_text(df.iloc[i]) for i in range(start + 1, pos)]
                lines = [line for line in lines if line]
                if len(lines) > 3:                              # keep line 1 and 2, join the rest
                    lines = lines[:2] + [" ".join(lines[2:])]
                lines += [""] * (3 - len(lines))
                if log:
                    log.info("Address found on sheet '%s': %s", sheet_name, lines)
                return lines
    if log:
        log.warning("Address block not found between '%s' and '%s'",
                    ADDRESS_START_TEXT, ADDRESS_END_TEXT)
    return ["", "", ""]


def write_statement_address(output_file, input_file, log=None):
    """Read the address from the input file and fill every statement sheet. Call this LAST."""
    address_lines = read_address(input_file, log)
    thru = f"Statement Thru {datetime.now():%m-%d-%Y}"

    wb = load_workbook(output_file)
    done = []
    for ws in wb.worksheets:
        if not any(key in ws.title.lower() for key in STATEMENT_SHEET_KEYS):
            continue
        for cell_ref, text in zip(ADDRESS_CELLS, address_lines):
            ws[cell_ref] = text
        ws[THRU_CELL] = thru
        ensure_statement_logo(ws)                       # keep your existing logo helper on these sheets
        done.append(ws.title)

    wb.save(output_file)
    wb.close()
    if log:
        log.info("Statement address filled on: %s", done or "none found - check STATEMENT_SHEET_KEYS")