import math
import pandas as pd
import re

# 0-based column positions in the input sheet (A=0, B=1, ...). Change here if needed.
NAME_COL = 1         # B -> defendant name
REF_COL = 2          # C -> ref no.
DIRECT_COL = 4       # E -> direct payment / Ref. Acct #
PRINCIPAL_COL = 11    # F -> principal
INTEREST_COL = 12     # G -> interest
MIN_TOTAL_NUMBERS = 8 
SKIP_TEXT = "current balance"


def to_number(value):
    """'1,09,676.29' / '$31.66' / '(5.00)' -> float. Returns None if not a number."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace("$", "").replace(",", "").strip()
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]
    try:
        return float(text)
    except ValueError:
        return None


def clean_text(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    if isinstance(value, float) and value.is_integer():    # 83245.0 -> "83245"
        return str(int(value))
    return str(value).strip()

# ---------------------------------------------------------------------
# Business-defendant rules (client requirement)
# ---------------------------------------------------------------------
BUSINESS_SUFFIXES = {
    "llc", "llp", "lp", "pllc", "pc", "pa", "inc", "incorporated",
    "corp", "corporation", "co", "company", "ltd", "limited", "trust",
    "association", "assoc", "partnership", "holdings",
}   # periods are removed before comparing, so L.L.C. -> llc, P.C. -> pc, P.A. -> pa

DBA_PATTERN = re.compile(r"(?<![a-z0-9])d[./]?b[./]?a\.?(?![a-z0-9])")   # dba, d/b/a, d.b.a.
STARTS_PATTERN = re.compile(r"^(the|city of)(?![a-z0-9])")               # "The ...", "City of ..."


def _has_business_suffix(name):
    match = re.search(r"([a-z]+)\W*$", str(name).lower())
    return bool(match and match.group(1) in BUSINESS_SUFFIXES)


def is_business(name):
    """True if the defendant name looks like a business. Whole words only, not case-sensitive."""
    text = " ".join(str(name).lower().split())
    text = text.rstrip(" .,")                       # ignore trailing periods and commas
    if not text:
        return False

    if _has_business_suffix(text):
        return True
    if "&" in text or any(ch.isdigit() for ch in text):
        return True                                 # ampersand or a number anywhere
    if DBA_PATTERN.search(text):
        return True                                 # DBA / D/B/A anywhere
    if STARTS_PATTERN.match(text):
        return True                                 # begins with 'The' or 'City of'

    last_word = text.split(" ")[-1].replace(".", "").strip(",")
    return last_word in BUSINESS_SUFFIXES           # ends with a business suffix


# ---------------------------------------------------------------------
# Name formatting
# ---------------------------------------------------------------------



SUFFIXES = {"jr", "jr.", "sr", "sr.", "ii", "iii", "iv", "v"}


def format_name(name):
    """'Tabitha Wood' -> 'Wood, Tabitha'
    'Mary Ann Smith' -> 'Smith, Mary Ann'
    'John Smith Jr' -> 'Smith Jr, John'
    Business names, names already in 'Last, First' form and single words -> unchanged."""
    name = " ".join(str(name).split())
    if _has_business_suffix(name):
        return name
    if not name or "," in name:
        return name                      # empty or already in 'Last, First' form

    if is_business(name):
        return name                      # business: keep exactly as written

    parts = name.split(" ")
    if len(parts) < 2:
        return name                      # single word, nothing to swap

    suffix = ""
    if parts[-1].lower() in SUFFIXES and len(parts) > 2:
        suffix = " " + parts.pop()       # keep Jr / Sr / III with the last name

    last = parts[-1] + suffix
    first = " ".join(parts[:-1])
    return f"{last}, {first}"


def cell(row, idx):
    return row[idx] if idx in row.index else None


def has_text(row):
    """True if any cell contains letters (a name, a label...)."""
    return any(isinstance(v, str) and any(ch.isalpha() for ch in v)
               for v in row.tolist() if pd.notna(v))


def count_numbers(row):
    return sum(to_number(v) is not None for v in row.tolist())


def is_totals_row(row, text):
    """Grand totals row: 1,09,676.29  0.00  24,127.82 ... (10 numbers).
    Does not depend on the name/ref cells being empty."""
    n = count_numbers(row)
    if n < MIN_TOTAL_NUMBERS:
        return False
    if not has_text(row):          # only numbers -> totals row
        return True
    return "total" in text         # or a row labelled 'Total...'


def parse_rows(rows, log):
    records = []
    skipped_zero = 0
    sum_principal = 0.0
    sum_interest = 0.0
    found_totals = False

    for idx, row in rows.iterrows():
        text = " ".join(str(v) for v in row.tolist() if pd.notna(v)).lower()
        if not text.strip():
            continue                                    # blank row
        if SKIP_TEXT in text:
            continue                                    # 'Current Balance: $31.66'

        # ---- END POINT: stop at the totals row ----
        if is_totals_row(row, text):
            found_totals = True
            exp_p = to_number(cell(row, PRINCIPAL_COL)) or 0.0
            exp_i = to_number(cell(row, INTEREST_COL)) or 0.0
            log.info("Totals row reached (after-marker row %d): %s - stopping",
                     idx, text[:120])
            if round(exp_p - sum_principal, 2) != 0 or round(exp_i - sum_interest, 2) != 0:
                log.warning("Totals mismatch: report P=%.2f I=%.2f | collected P=%.2f I=%.2f",
                            exp_p, exp_i, sum_principal, sum_interest)
            break

        name = clean_text(cell(row, NAME_COL))
        ref = clean_text(cell(row, REF_COL))
        principal = to_number(cell(row, PRINCIPAL_COL))
        interest = to_number(cell(row, INTEREST_COL))

        if not name and not ref:
            continue                                    # stray / sub-header row
        if principal is None and interest is None:
            continue                                    # header row or name-only line

        principal = principal or 0.0
        interest = interest or 0.0
        sum_principal += principal
        sum_interest += interest

        if round(principal + interest, 2) == 0:
            skipped_zero += 1
            continue                                    # Principal + Interest = 0

        records.append({
            "name": format_name(name),
            "is_business": is_business(name),           # flag kept for later use
            "ref_no": ref,
            "principal": principal,
            "interest": interest,
            "direct_payment": to_number(cell(row, DIRECT_COL)) or 0.0,
        })

    if not found_totals:
        log.warning("Totals row was NOT found - the bot read to the end of the sheet")

    log.info("Valid rows: %d | skipped (principal + interest = 0): %d",
             len(records), skipped_zero)
    return pd.DataFrame(records, columns=["name", "is_business", "ref_no", "principal",
                                          "interest", "direct_payment"])

if __name__ == "__main__":
    for n in ["Coleman", "Incorporate", "Acme Inc.", "Smith, Inc", "Theodore Smith",
              "The Smith Family", "City of Atlanta", "John Doe d/b/a Cuts",
              "Ross, Stuart & Dawson", "Tabitha Wood"]:
        print(f"{n!r:35} business={is_business(n)}  ->  {format_name(n)}")