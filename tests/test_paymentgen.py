from datetime import date
from decimal import Decimal

import pytest

from paymentgen.calendar_rules import CalendarError, covered_dates
from paymentgen.layout import apply_overrides, parse_sample
from paymentgen.service import Request, RequestError, generate
from tests import samples

WEEKEND = {5, 6}


def d(s):
    return date.fromisoformat(s)


# ---------------------------------------------------------------- calendar
@pytest.mark.parametrize(
    "file_date,non_working,expected",
    [
        ("2024-01-08", WEEKEND, ["2024-01-05", "2024-01-06", "2024-01-07"]),  # normal Monday
        ("2024-01-09", WEEKEND, ["2024-01-08"]),  # normal Tuesday
        ("2024-01-08", WEEKEND | {4}, ["2024-01-04", "2024-01-05", "2024-01-06", "2024-01-07"]),  # Fri holiday
        ("2024-01-09", WEEKEND | {0}, ["2024-01-05", "2024-01-06", "2024-01-07", "2024-01-08"]),  # Mon holiday
        ("2024-01-09", WEEKEND | {0, 4}, ["2024-01-04", "2024-01-05", "2024-01-06", "2024-01-07", "2024-01-08"]),
    ],
)
def test_covered_dates(file_date, non_working, expected):
    assert covered_dates(d(file_date), non_working) == [d(x) for x in expected]


def test_non_working_file_date_rejected():
    with pytest.raises(CalendarError):
        covered_dates(d("2024-01-06"), WEEKEND)


# ---------------------------------------------------------------- layout inference
def test_pipe_layout_detected():
    lay = parse_sample("pay_20240105.txt", samples.PIPE_V5.encode())
    assert lay.fmt == "delimited" and lay.delimiter == "|"
    assert len(lay.header) == 2 and lay.has_column_header and len(lay.trailer) == 1
    roles = {c.name: c.role for c in lay.columns}
    assert roles == {
        "MANUFACTURER": "manufacturer", "PAY_DATE": "payment_date", "AMOUNT": "amount",
        "DC_IND": "dc_indicator", "PAYMENT_REF": "reference",
    }
    assert lay.debit_style == "indicator"
    assert {(b.stat, b.index) for b in lay.bindings} == {("count", 1), ("net", 2)}


def test_fixed_layout_detected():
    lay = parse_sample("pay.dat", samples.FIXED_V5.encode())
    assert lay.fmt == "fixed" and len(lay.header) == 1 and len(lay.trailer) == 1
    assert [c.role for c in lay.columns] == ["manufacturer", "payment_date", "amount", "dc_indicator", "reference"]
    amt = lay.columns[2]
    assert amt.zero_pad == 12 and amt.implied_point and amt.decimals == 2


def test_excel_layout_detected():
    lay = parse_sample("pay.xlsx", samples.make_xlsx())
    assert lay.fmt == "excel" and lay.sheet_name == "Payments"
    assert [c.role for c in lay.columns][:3] == ["manufacturer", "payment_date", "amount"]


def test_jlr_csv_layouts():
    dd = parse_sample("DD.csv", samples.CSV_DD.encode())
    assert [c.quoted for c in dd.columns] == [True] * 4
    dc = parse_sample("DC.csv", samples.CSV_DC.encode())
    assert dc.newline == "\r\n" and dc.debit_style == "columns"


# ---------------------------------------------------------------- generation
def request(version, layouts, **kw):
    base = dict(
        version=version, layouts=layouts, file_date=d("2024-01-08"), non_working_days=WEEKEND,
        manufacturers=["AUD", "BMW"], include_debits=True, exceed_threshold=False, seed=7,
    )
    base.update(kw)
    return Request(**base)


def rows(data: bytes, delim="|"):
    return [l.split(delim) for l in data.decode().splitlines()]


def test_v5_multi_manufacturer_dates_debits_and_trailer():
    lay = parse_sample("pay_20240105.txt", samples.PIPE_V5.encode())
    [f] = generate(request("V5", {"V5": lay}))
    assert f.name == "pay_20240108.txt"
    lines = rows(f.data)
    assert lines[0] == ["HDR", "20240108", "PAYMENTS"]  # header date refreshed to the file date
    detail = lines[2:-1]
    assert {r[0] for r in detail} == {"AUD", "BMW"}
    assert {r[1] for r in detail} == {"20240105", "20240106", "20240107"}  # Fri/Sat/Sun
    assert any(r[3] == "DR" for r in detail) and any(r[3] == "CR" for r in detail)
    assert lines[-1][1] == str(len(detail))  # trailer record count refreshed
    credits = sum(Decimal(r[2]) for r in detail if r[3] == "CR")
    debits = sum(Decimal(r[2]) for r in detail if r[3] == "DR")
    assert Decimal(lines[-1][2]) == credits - debits  # trailer net total refreshed
    refs = [r[4] for r in detail]
    assert len(set(refs)) == len(refs) and all(r.startswith("PAY") and len(r) == 9 for r in refs)


def test_v5_without_debits():
    lay = parse_sample("pay.txt", samples.PIPE_V5.encode())
    [f] = generate(request("V5", {"V5": lay}, include_debits=False))
    assert all(r[3] == "CR" for r in rows(f.data)[2:-1])


@pytest.mark.parametrize("exceed", [True, False])
def test_million_threshold(exceed):
    lay = parse_sample("pay.txt", samples.PIPE_V5.encode())
    [f] = generate(request("V5", {"V5": lay}, exceed_threshold=exceed, manufacturers=["AUD", "BMW", "FORD"]))
    assert len(f.summary) == 3
    for row in f.summary:
        assert (row["Net payment"] > Decimal(1_000_000)) is exceed


def test_threshold_applies_with_non_working_weekdays_spanning_more_days():
    lay = parse_sample("pay.txt", samples.PIPE_V5.encode())
    [f] = generate(request("V5", {"V5": lay}, file_date=d("2024-01-09"), non_working_days=WEEKEND | {0, 4}))
    assert {r[1] for r in rows(f.data)[2:-1]} == {f"202401{n:02d}" for n in (4, 5, 6, 7, 8)}


def test_seed_is_reproducible():
    lay = parse_sample("pay.txt", samples.PIPE_V5.encode())
    a = generate(request("V5", {"V5": lay}))[0].data
    b = generate(request("V5", {"V5": lay}))[0].data
    assert a == b


def test_fixed_width_output_keeps_geometry():
    lay = parse_sample("pay.dat", samples.FIXED_V5.encode())
    [f] = generate(request("V5", {"V5": lay}))
    lines = f.data.decode().splitlines()
    assert lines[0] == "H 20240108 PAYFILE"
    detail = lines[1:-1]
    assert len({len(l) for l in detail}) == 1 and len(detail[0]) == len(samples.FIXED_V5.splitlines()[1])
    assert lines[-1] == f"T {len(detail):07d}"
    # round trip: the generated file is itself a valid sample with the same layout
    again = parse_sample("x.dat", f.data)
    assert [c.role for c in again.columns] == [c.role for c in lay.columns]


def test_excel_roundtrip():
    lay = parse_sample("pay.xlsx", samples.make_xlsx())
    [f] = generate(request("V5", {"V5": lay}))
    assert f.name.endswith(".xlsx")
    again = parse_sample("out.xlsx", f.data)
    assert again.sheet_name == "Payments" and [c.role for c in again.columns] == [c.role for c in lay.columns]
    assert again.columns[1].kind == "date" and again.columns[2].kind == "decimal"
    assert {str(v.date()) for v in again.column_values(1)} == {"2024-01-05", "2024-01-06", "2024-01-07"}
    assert again.trailer[0][1] == len(again.details)


def test_v6_dd_dc():
    dd = parse_sample("JLR_DD_20240105.csv", samples.CSV_DD.encode())
    dc = parse_sample("JLR_DC_20240105.csv", samples.CSV_DC.encode())
    # DD sample has no header row / roles from names -> tell the tool which column is which
    dd = apply_overrides(dd, [
        dict(name="Mfr", role="manufacturer"), dict(name="Date", role="payment_date"),
        dict(name="Seq", role="sequence"), dict(name="Debit", role="amount", decimals=2),
    ])
    req = request("V6", {"DD": dd, "DC": dc}, manufacturers=["JLR"], exceed_threshold=True)
    f_dd, f_dc = generate(req)
    assert f_dd.name == "JLR_DD_20240108.csv" and f_dc.name == "JLR_DC_20240108.csv"
    dd_rows = [l.split(",") for l in f_dd.data.decode().splitlines()]
    assert all(r[0] == '"JLR"' and r[1].startswith('"') for r in dd_rows)  # quoting preserved
    assert any(Decimal(r[3].strip('"')) != 0 for r in dd_rows)
    assert f_dc.data.decode().count("\r\n") == len(f_dc.data.decode().splitlines())  # CRLF preserved
    assert f_dc.summary[0]["Net payment"] > Decimal(1_000_000)
    # zero debit option
    f_dd0, _ = generate(request("V6", {"DD": dd, "DC": dc}, manufacturers=["JLR"], include_debits=False))
    assert all(Decimal(l.split(",")[3].strip('"')) == 0 for l in f_dd0.data.decode().splitlines())


def test_validation_errors():
    lay = parse_sample("pay.txt", samples.PIPE_V5.encode())
    with pytest.raises(RequestError):
        generate(request("V5", {"V5": lay}, manufacturers=["  "]))
    with pytest.raises(RequestError):
        generate(request("V5", {"V5": lay}, manufacturers=["AUD", "aud"]))
    with pytest.raises(CalendarError):
        generate(request("V5", {"V5": lay}, file_date=d("2024-01-06")))


def test_v6_auto_detection_and_positive_debit_amounts():
    dd = parse_sample("JLR_DD_20240105.csv", samples.CSV_DD.encode())
    assert [c.role for c in dd.columns] == ["manufacturer", "payment_date", "sequence", "amount"]
    dc = parse_sample("JLR_DC_20240105.csv", samples.CSV_DC.encode())
    f_dd, f_dc = generate(request("V6", {"DD": dd, "DC": dc}, manufacturers=["JLR"]))
    amounts = [Decimal(l.split(",")[3].strip('"')) for l in f_dd.data.decode().splitlines()]
    assert amounts and all(a > 0 for a in amounts)  # debit file keeps positive amounts
    seqs = [int(l.split(",")[2].strip('"')) for l in f_dd.data.decode().splitlines()]
    assert seqs == list(range(1, len(seqs) + 1))
    cols = [l.split(",") for l in f_dc.data.decode().splitlines()[1:]]
    assert all(Decimal(c[3]) == 0 for c in cols) and all(Decimal(c[2]) > 0 for c in cols)  # DC: credits only
    assert f_dc.data.decode().splitlines()[0] == "Manufacturer,Payment Date,Credit Amount,Debit Amount,Invoice"


def test_v5_unsigned_sample_falls_back_to_negative_debits():
    lay = parse_sample("pay.csv", b"AUD,20240104,10.50\nAUD,20240104,20.75\nBMW,20240104,5.00\n")
    assert lay.debit_style == "unsigned"
    [f] = generate(request("V5", {"V5": lay}, include_debits=True))
    assert any(l.split(",")[2].startswith("-") for l in f.data.decode().splitlines())
    [f] = generate(request("V5", {"V5": lay}, include_debits=False))
    assert not any(l.split(",")[2].startswith("-") for l in f.data.decode().splitlines())
