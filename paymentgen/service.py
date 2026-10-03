"""Orchestration: turn the user's choices into one (V5) or two (V6 DD + DC) files."""
from __future__ import annotations

import dataclasses
import os
import random
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal

from .calendar_rules import covered_dates
from .generator import ZERO, Payment, choose_debit_total, choose_target, make_payments
from .layout import Layout, refresh_dates
from .render import render

ONE_MILLION = Decimal("1000000")


class RequestError(ValueError):
    pass


@dataclass
class Request:
    version: str  # "V5" (non-JLR) | "V6" (JLR)
    layouts: dict[str, Layout]  # {"V5": ...} or {"DD": ..., "DC": ...}
    file_date: date
    non_working_days: set[int]
    manufacturers: list[str]
    include_debits: bool  # V5: write debit rows; V6: DD carries real values (False -> all zero)
    exceed_threshold: bool
    threshold: Decimal = ONE_MILLION
    rows_per_day: tuple[int, int] = (2, 5)
    seed: int | None = None


@dataclass
class GeneratedFile:
    name: str
    data: bytes
    label: str
    summary: list[dict] = field(default_factory=list)  # per manufacturer: credits / debits / net / rows


def output_name(sample_name: str, file_date: date) -> str:
    stem, ext = os.path.splitext(os.path.basename(sample_name))
    new = refresh_dates(stem, file_date)
    if new == stem:
        new = f"{stem}_{file_date:%Y%m%d}"
    return new + ext


def _summary(payments: list[Payment]) -> list[dict]:
    out: dict[str, dict] = {}
    for p in payments:
        row = out.setdefault(p.manufacturer, {"Manufacturer": p.manufacturer, "Rows": 0, "Credits": ZERO, "Debits": ZERO})
        row["Rows"] += 1
        row["Debits" if p.is_debit else "Credits"] += p.amount
    for row in out.values():
        row["Net payment"] = row["Credits"] - row["Debits"]
    return list(out.values())


def generate(req: Request) -> list[GeneratedFile]:
    codes = [c.strip() for c in req.manufacturers if c and c.strip()]
    if not codes:
        raise RequestError("Enter at least one manufacturer code.")
    if len(set(c.upper() for c in codes)) != len(codes):
        raise RequestError("Manufacturer codes must be unique.")
    lo, hi = req.rows_per_day
    if lo < 1 or hi < lo:
        raise RequestError("Rows per day must be at least 1 and min <= max.")
    dates = covered_dates(req.file_date, req.non_working_days)
    rng = random.Random(req.seed)
    targets = {c: choose_target(req.exceed_threshold, req.threshold, rng) for c in codes}

    if req.version == "V5":
        layout = req.layouts["V5"]
        if req.include_debits and layout.debit_style == "unsigned":
            # the sample gives no way to tell debits apart: fall back to negative amounts
            layout = dataclasses.replace(layout, debit_style="signed")
        payments: list[Payment] = []
        for c in codes:
            debit = choose_debit_total(targets[c], rng) if req.include_debits else None
            # target is the NET payment: credits = net + debits
            credit = targets[c] + (debit or ZERO)
            payments += make_payments(c, dates, rng, credit_total=credit, debit_total=debit, rows_per_day=req.rows_per_day)
        data = render(layout, payments, req.file_date, req.seed)
        return [GeneratedFile(output_name(layout.source_name, req.file_date), data, "V5 payment file", _summary(payments))]

    if req.version == "V6":
        results = []
        dd_pay: list[Payment] = []
        dc_pay: list[Payment] = []
        for c in codes:
            dc_pay += make_payments(c, dates, rng, credit_total=targets[c], debit_total=None, rows_per_day=req.rows_per_day)
            dd_pay += make_payments(
                c, dates, rng, credit_total=None, debit_total=choose_debit_total(targets[c], rng),
                rows_per_day=req.rows_per_day, zero_debits=not req.include_debits,
            )
        for key, pays, label in (("DD", dd_pay, "V6 DD (debit) file"), ("DC", dc_pay, "V6 DC (credit) file")):
            layout = req.layouts[key]
            results.append(
                GeneratedFile(output_name(layout.source_name, req.file_date), render(layout, pays, req.file_date, req.seed), label, _summary(pays))
            )
        return results

    raise RequestError(f"Unknown version {req.version!r}.")
