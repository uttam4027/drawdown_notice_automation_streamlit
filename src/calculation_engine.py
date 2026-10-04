"""STEP 7: financial calculations, recomputed from Working inputs (not from cached / broken Excel formulas).

Formulas follow the workbook exactly:
    previous contribution = SUM(1st ... 14th Drawdown Amount Paid)         (Working col J)
    amount payable        = Amount Called + Pending previous - Excess      (Working col W)
    % of commitment       = (Amount payable + Previous contribution + Excess) / Commitment   (Working col AL)
"""
import logging

import config

log = logging.getLogger("drawdown")


def calculate_values(rec):
    previous = sum(rec["paid_amounts"].values())
    payable = rec["drawdown_called"] + rec["pending_previous"] - rec["excess_contribution"]
    commitment = rec["capital_commitment"]
    pct = (payable + previous + rec["excess_contribution"]) / commitment if commitment else None
    return {
        "capital_commitment": commitment,
        "capital_contributed": previous,
        "drawdown_called": rec["drawdown_called"],
        "pending_previous": rec["pending_previous"],
        "excess_contribution": rec["excess_contribution"],
        "adjustment": rec["adjustment"],
        "amount_payable": round(payable, 2),
        "remit_amount": round(payable, 2),
        "percentage": pct,
        "portfolio_investment": rec["portfolio_investment"],
        "management_fee": rec["management_fee"],
        "fund_expenses": rec["fund_expenses"],
    }


def validate_financial_calculations(rec, calc):
    """Compare recomputed values with the workbook's own numbers and sanity-check them.
    Returns a list of issue strings (empty list = PASS)."""
    issues = []
    tol, ptol = config.MONEY_TOLERANCE, config.PCT_TOLERANCE

    if not calc["capital_commitment"] or calc["capital_commitment"] <= 0:
        issues.append("Capital commitment is zero/blank")
    if calc["drawdown_called"] <= 0:
        issues.append("15th drawdown amount called is zero/blank")
    if calc["pending_previous"] < 0:
        issues.append("Pending previous contribution is negative")
    if calc["excess_contribution"] < 0:
        issues.append("Excess contribution is negative")
    if calc["amount_payable"] < 0:
        issues.append(f"Amount payable is negative ({calc['amount_payable']:,.2f})")

    # recomputed vs workbook values
    if abs(calc["capital_contributed"] - rec["capital_contributed_wb"]) > tol:
        issues.append(f"Previous contribution mismatch: calc {calc['capital_contributed']:,.2f} "
                      f"vs workbook {rec['capital_contributed_wb']:,.2f}")
    if abs(calc["amount_payable"] - rec["amount_payable_wb"]) > tol:
        issues.append(f"Amount payable mismatch: calc {calc['amount_payable']:,.2f} "
                      f"vs workbook {rec['amount_payable_wb']:,.2f}")
    if calc["percentage"] is not None and abs(calc["percentage"] - rec["percentage_wb"]) > ptol:
        issues.append(f"% of commitment mismatch: calc {calc['percentage']:.6f} vs workbook {rec['percentage_wb']:.6f}")
    if calc["percentage"] is not None and calc["percentage"] > 1:
        issues.append(f"% of commitment exceeds 100% ({calc['percentage']:.2%})")

    # portfolio + management fee + fund expenses should add up to the amount called
    split = rec["portfolio_investment"] + rec["management_fee"] + rec["fund_expenses"]
    if abs(split - rec["drawdown_called"]) > 1:      # allow rounding of a rupee
        issues.append(f"Portfolio+mgmt fee+fund expenses ({split:,.2f}) differ from amount called "
                      f"({rec['drawdown_called']:,.2f})")
    return issues
