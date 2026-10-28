from decimal import ROUND_HALF_UP, Decimal

from temp3_Firm_Payments.src4.settings import COMMISSION_RATE, DECIMALS


def round3(value: Decimal) -> Decimal:
    return value.quantize(DECIMALS, rounding=ROUND_HALF_UP)


def calculate(payment: dict) -> dict:
    gross = payment["amount"] + payment["overpayment"]   # overpayment is negative in the report
    commission = round3(gross * COMMISSION_RATE)          # always 25% of gross, never from input
    net = round3(gross - commission)
    return {**payment, "gross": round3(gross), "commission": commission, "net": net}