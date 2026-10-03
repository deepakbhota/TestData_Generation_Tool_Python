# TestData Generation Tool (Python)

Generates manufacturer **test payment files** in the same layout/format as a sample file you upload.

* **V5** - non-JLR: one file, one or many manufacturers, credits and (optionally) debits.
* **V6** - JLR: two files, **DD** (debit) and **DC** (credit).

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## What the UI offers

| Option | Behaviour |
|---|---|
| Sample file | Delimited text (csv/txt/dat, `, \| ; tab`), fixed-width text or Excel (`.xlsx`). Header / detail / trailer records, delimiter, quoting, line endings, encoding, number/date formats, sheet name and cell styles are detected and reproduced. Output has the same extension. |
| Layout check | Every column gets a *role* (manufacturer, payment_date, amount, credit_amount, debit_amount, dc_indicator, sequence, reference, other). Auto-detected; correct it in the "Detected layout" table (fixed-width: also start/width). |
| File date + non-working days | Tick the non-working weekdays (Sat/Sun by default). The file holds every day from the previous working day up to the day before the file date (Mon = Fri+Sat+Sun; Fri off -> Mon = Thu..Sun; Mon off -> Tue = Fri..Mon; Fri+Mon off -> Tue = Thu..Mon). |
| Single / multi manufacturer (V5) | One text box, or `＋ Add manufacturer` for more. |
| Debits | V5: include debit rows or not. JLR: DD file with debit values or all amounts zero. |
| Total exceeds 1m | Ticked: each manufacturer's **net** payment (credits - debits) is above the threshold (default 1,000,000, editable under *Advanced*). Unticked: safely below. |
| Advanced | rows per day per manufacturer, random seed for repeatable output. |

Header/trailer values are refreshed: dates become the file date, and record counts / totals that matched the sample's
own data (count, credit sum, debit sum, total, net) are recomputed.

## Assumptions to confirm

* "Total payment will increase 1m" is implemented as *net total above 1,000,000* (threshold configurable).
* How debits appear is learned from the sample (DR/CR indicator column, separate debit/credit columns, or negative
  amounts). If the sample shows none, V5 debits are written as negative amounts; JLR DD amounts stay positive.
* Header/trailer dates are set to the file generation date.
* Legacy `.xls` is not supported (save as `.xlsx`).

## Code layout

```
app.py                   Streamlit UI
paymentgen/calendar_rules.py   covered dates for a file date
paymentgen/layout.py     sample parsing, schema/role inference
paymentgen/generator.py  amount splitting and payment rows
paymentgen/render.py     writes text / fixed-width / Excel output
paymentgen/service.py    V5 / V6 orchestration
tests/                   pytest suite (python -m pytest)
```
