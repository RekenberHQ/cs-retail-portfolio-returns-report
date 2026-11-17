"""Return rates are percentage points throughout (60 means 60%)."""
from datetime import date
import math

from dateutil.relativedelta import relativedelta

DEFAULT_SCOPE = "splátkový kalendář se zpeněžením majetkové podstaty"
THREE_YEAR_SCOPE = DEFAULT_SCOPE + " za 3 roky"


def parse_date(value):
    return date.fromisoformat(str(value)[:10]) if value else None


def rate(value):
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        raise ValueError("A return rate must be numeric, not boolean")
    result = float(value)
    if not math.isfinite(result) or result < 0:
        raise ValueError("A return rate must be finite and nonnegative")
    return result


def retirement_months(year):
    """Higher ordinary Czech age limit before 2025, in months.

    Only used for proposals in [2021-06-01, 2025-01-01).
    Source: https://www.cssz.gov.cz/starobni-duchod-podrobne
    Birth cohorts relevant to these proposals were unaffected by the 2025 reform.
    """
    return 720 if year < 1936 else min(780, 722 + 2 * (year - 1936))


def retired_at_proposal(debtors, proposal):
    """Best effort: use the higher ordinary age limit regardless of gender/children."""
    for debtor in debtors:
        born = parse_date(debtor.get("birth_date"))
        if born is None or born > proposal:
            continue
        if born + relativedelta(months=retirement_months(born.year)) <= proposal:
            return True
    return False


def select_expected(insolvency, discharge, debtors):
    """Return (value, review reasons), before manual overrides."""
    proposal = parse_date(insolvency.get("proposal_date"))
    if proposal is None:
        return None, ["Missing proposal_date; cannot select return-rate regime"]
    if proposal >= date(2025, 1, 1):
        return rate((discharge.get("approval") or {}).get("approved_expected_return_rate")), []
    estimates = {}
    for item in discharge.get("satisfaction_estimates") or []:
        scope, value = item.get("estimation_scope"), rate(item.get("rate"))
        if scope in estimates and estimates[scope] != value:
            return None, [f"Conflicting satisfaction estimates: {scope}"]
        estimates[scope] = value
    default, short = estimates.get(DEFAULT_SCOPE), estimates.get(THREE_YEAR_SCOPE)
    if proposal < date(2021, 6, 1):
        return default, []
    if short is not None and short >= 60:
        return short, []
    # Retirement cannot affect selection when no short estimate exists or both agree.
    if short is None or short == default:
        return default, []
    retired = retired_at_proposal(debtors, proposal)
    return (short if retired else default), []


def build_row(insolvency, discharge, debtors, current, correction):
    expected, reasons = select_expected(insolvency, discharge, debtors)
    note = str(correction.get("note") or "").strip()
    override = rate(correction.get("expected_return_rate"))
    if override is not None:
        expected, reasons = override, []
    approved = (discharge.get("approval") or {}).get("discharge_approved_date")
    approved = approved or discharge.get("discharge_approved_date")
    if approved and expected is None:
        reasons.append("Approved discharge missing required expected return rate")
    if not discharge:
        reasons.append("Missing discharge resource; approval status unknown")
    allowed = parse_date(discharge.get("discharge_allowed_date"))
    if allowed is None:
        reasons.append("Missing discharge_allowed_date")
    ref = insolvency["reference_number"]
    current_rate = rate(current)
    row = {
        "SPIS_ZNACKA": f"INS {ref['case_number']}/{ref['year']}",
        "DATUM_UPADKU": allowed,
        "STAV": insolvency["state"],
        "PREDP_USP": min(expected, 100) if expected is not None else None,
        "PREDP_USP_AKT": min(current_rate, 100) if current_rate is not None else None,
        "NOTE": note or None,
    }
    return row, reasons if not note else []
