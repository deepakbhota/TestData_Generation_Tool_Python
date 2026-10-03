"""Random but controlled payment data: totals per manufacturer split into rows."""
from __future__ import annotations

import random
from dataclasses import dataclass
from datetime import date
from decimal import ROUND_DOWN, ROUND_HALF_UP, Decimal

CENT = Decimal("0.01")
ZERO = Decimal("0.00")


@dataclass(frozen=True)
class Payment:
    manufacturer: str
    pay_date: date
    amount: Decimal  # always >= 0; direction is given by ``is_debit``
    is_debit: bool


def choose_target(exceed: bool, threshold: Decimal, rng: random.Random) -> Decimal:
    """Net payment total for one manufacturer: above or safely below the threshold."""
    lo, hi = (1.02, 1.5) if exceed else (0.05, 0.9)
    return (threshold * Decimal(str(round(rng.uniform(lo, hi), 4)))).quantize(CENT, ROUND_HALF_UP)


def choose_debit_total(credit_total: Decimal, rng: random.Random) -> Decimal:
    return (credit_total * Decimal(str(round(rng.uniform(0.03, 0.2), 4)))).quantize(CENT, ROUND_HALF_UP)


def split_amount(total: Decimal, parts: int, rng: random.Random) -> list[Decimal]:
    """Split ``total`` into ``parts`` positive 2dp amounts that sum exactly to it."""
    total = total.quantize(CENT)
    parts = max(1, min(parts, int(total / CENT))) if total > 0 else 1
    if parts == 1:
        return [total]
    weights = [rng.uniform(0.3, 1.7) for _ in range(parts)]
    scale = sum(weights)
    out = [(total * Decimal(str(w / scale))).quantize(CENT, ROUND_DOWN) for w in weights]
    out[-1] = total - sum(out[:-1])
    if min(out) <= 0:  # pathological tiny totals: fall back to an even split
        base = (total / parts).quantize(CENT, ROUND_DOWN)
        out = [base] * (parts - 1) + [total - base * (parts - 1)]
    return out


def make_payments(
    manufacturer: str,
    dates: list[date],
    rng: random.Random,
    *,
    credit_total: Decimal | None,
    debit_total: Decimal | None,
    rows_per_day: tuple[int, int] = (2, 5),
    zero_debits: bool = False,
) -> list[Payment]:
    """Credit rows (>=1 per covered day) and/or debit rows for one manufacturer."""
    lo, hi = rows_per_day
    out: list[Payment] = []
    if credit_total is not None:
        days = [d for d in dates for _ in range(rng.randint(lo, hi))]
        out += [Payment(manufacturer, d, a, False) for d, a in zip(days, split_amount(credit_total, len(days), rng))]
    if debit_total is not None:
        n = rng.randint(1, max(1, len(dates) * max(1, hi // 2)))
        days = [rng.choice(dates) for _ in range(n)]
        amounts = [ZERO] * n if zero_debits else split_amount(debit_total, n, rng)
        out += [Payment(manufacturer, d, a, True) for d, a in zip(days, amounts)]
    out.sort(key=lambda p: (p.pay_date, p.is_debit))
    return out
