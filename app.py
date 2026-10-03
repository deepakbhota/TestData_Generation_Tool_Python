"""Streamlit UI for the manufacturer payment test-file generator.

Run with:  streamlit run app.py
"""
from __future__ import annotations

import hashlib
import io
import zipfile
from datetime import date, timedelta
from decimal import Decimal

import pandas as pd
import streamlit as st

from paymentgen.calendar_rules import DEFAULT_NON_WORKING, WEEKDAYS, CalendarError, covered_dates
from paymentgen.layout import ROLES, SampleError, apply_overrides, parse_sample
from paymentgen.service import ONE_MILLION, Request, RequestError, generate

st.set_page_config(page_title="Payment Test File Generator", page_icon="💷", layout="wide")


# ----------------------------------------------------------------------------- helpers
@st.cache_resource(show_spinner=False)
def _parse(name: str, digest: str, data: bytes):
    return parse_sample(name, data)


def layout_section(title: str, key: str, help_text: str, debits_needed: bool = False):
    """Upload a sample, show the detected layout and let the user correct the column mapping."""
    st.markdown(f"**{title}**")
    up = st.file_uploader(help_text, key=f"up_{key}", label_visibility="collapsed")
    if up is None:
        return None
    data = up.getvalue()
    try:
        base = _parse(up.name, hashlib.md5(data).hexdigest(), data)
    except SampleError as exc:
        st.error(str(exc))
        return None
    kind = {"delimited": f"delimited text (delimiter `{base.delimiter!r}`)", "fixed": "fixed-width text", "excel": "Excel workbook"}
    st.caption(
        f"Detected **{kind[base.fmt]}** - `{base.ext or 'no extension'}` - "
        f"{len(base.header)} header, {len(base.details)} detail, {len(base.trailer)} trailer record(s). "
        "Output is generated in this same format."
    )
    df = pd.DataFrame(
        {
            "name": [c.name for c in base.columns],
            "role": [c.role for c in base.columns],
            "kind": [c.kind for c in base.columns],
            "sample": [", ".join(c.pool[:3]) for c in base.columns],
            "start": [c.start for c in base.columns],
            "width": [c.width for c in base.columns],
            "decimals": [c.decimals for c in base.columns],
            "implied_point": [c.implied_point for c in base.columns],
        }
    )
    fixed = base.fmt == "fixed"
    with st.expander("Detected layout - check / correct the column mapping", expanded=False):
        st.caption(
            "Roles tell the generator what to put in each column. Change a role if the auto-detection is wrong. "
            "`implied_point` means amounts are stored without a decimal point (e.g. `000000012345` = 123.45)."
        )
        edited = st.data_editor(
            df,
            key=f"ed_{key}_{hashlib.md5(data).hexdigest()}",
            hide_index=True,
            use_container_width=True,
            num_rows="dynamic" if fixed else "fixed",
            disabled=["kind", "sample"] + ([] if fixed else ["start", "width"]),
            column_config={
                "role": st.column_config.SelectboxColumn("role", options=ROLES, required=True),
                "start": st.column_config.NumberColumn("start (0-based)", min_value=0, step=1),
                "width": st.column_config.NumberColumn("width", min_value=1, step=1),
                "decimals": st.column_config.NumberColumn("decimals", min_value=0, max_value=6, step=1),
            },
        )
        edits = edited.fillna({"role": "other", "start": 0, "width": 1, "decimals": 0, "implied_point": False}).to_dict("records")
    layout = apply_overrides(base, edits)
    for w in layout.warnings(debits_needed):
        st.warning(w)
    with st.expander("Sample file preview"):
        if layout.fmt == "excel":
            st.dataframe(pd.DataFrame(layout.header + layout.details + layout.trailer), use_container_width=True)
        else:
            st.code("\n".join(layout.header + layout.details + layout.trailer)[:4000], language=None)
    return layout


def manufacturer_inputs(multi: bool) -> list[str]:
    if not multi:
        return [st.text_input("Manufacturer code", key="mc_single", placeholder="e.g. AUD").strip()]
    n = st.session_state.setdefault("mc_count", 2)
    codes = []
    for i in range(n):
        c1, c2 = st.columns([6, 1])
        codes.append(c1.text_input(f"Manufacturer code {i + 1}", key=f"mc_{i}", placeholder="e.g. AUD").strip())
        if n > 1 and c2.button("✕", key=f"rm_{i}", help="Remove this code"):
            for j in range(i, n - 1):
                st.session_state[f"mc_{j}"] = st.session_state.get(f"mc_{j + 1}", "")
            st.session_state[f"mc_{n - 1}"] = ""
            st.session_state["mc_count"] = n - 1
            st.rerun()
    if st.button("＋ Add manufacturer"):
        st.session_state["mc_count"] = n + 1
        st.rerun()
    return codes


# ----------------------------------------------------------------------------- page
st.title("💷 Manufacturer Payment Test File Generator")
st.caption("Upload a sample file, choose the scenario, and generate a test file with the same layout and format.")

left, right = st.columns([3, 2], gap="large")

with left:
    st.subheader("1. File type & sample")
    version = st.radio(
        "Payment file version",
        ["V5 - non-JLR", "V6 - JLR (DD + DC)"],
        horizontal=True,
        help="V5: one file with the credits (and optionally debits) of one or many manufacturers. "
        "V6: two JLR files - DD (debit) and DC (credit).",
    )
    is_jlr = version.startswith("V6")
    layouts: dict = {}
    if is_jlr:
        layouts["DD"] = layout_section("Sample DD (debit) file", "dd", "Upload the sample JLR DD file")
        layouts["DC"] = layout_section("Sample DC (credit) file", "dc", "Upload the sample JLR DC file")
    else:
        layouts["V5"] = layout_section(
            "Sample V5 file", "v5", "Upload a sample non-JLR V5 file", debits_needed=st.session_state.get("inc_debits", True)
        )

with right:
    st.subheader("2. Dates")
    default_date = date.today()
    while default_date.weekday() in DEFAULT_NON_WORKING:  # first working day on/after today
        default_date += timedelta(days=1)
    file_date = st.date_input("File generation date", value=default_date)
    st.markdown("Non-working days")
    cols = st.columns(7)
    non_working = {
        i for i, name in enumerate(WEEKDAYS) if cols[i].checkbox(name[:3], value=i in DEFAULT_NON_WORKING, key=f"nw_{i}")
    }
    covered = None
    try:
        covered = covered_dates(file_date, non_working)
        st.info("This file will contain payments of: " + ", ".join(f"{d:%a %d %b}" for d in covered))
    except CalendarError as exc:
        st.error(str(exc))

    st.subheader("3. Manufacturers")
    if is_jlr:
        codes = [st.text_input("JLR manufacturer code", value="JLR", key="mc_jlr").strip()]
    else:
        multi = st.radio("Entries in the file", ["Single manufacturer", "Multiple manufacturers"], horizontal=True) == "Multiple manufacturers"
        codes = manufacturer_inputs(multi)

    st.subheader("4. Amounts")
    if is_jlr:
        include_debits = st.radio(
            "DD (debit) file values", ["Generate debit values", "Make all amounts zero"], horizontal=True
        ) == "Generate debit values"
    else:
        include_debits = st.checkbox("Include debit values", value=True, key="inc_debits")
    exceed = st.checkbox(
        "Manufacturer total payment exceeds 1m",
        value=False,
        help="Ticked: each manufacturer's net payment total is above the threshold. Unticked: safely below it.",
    )
    with st.expander("Advanced"):
        threshold = Decimal(str(st.number_input("Threshold (1m)", min_value=1000.0, value=float(ONE_MILLION), step=100000.0)))
        r1, r2 = st.columns(2)
        rows_min = int(r1.number_input("Min rows per day (per manufacturer)", min_value=1, value=2, step=1))
        rows_max = int(r2.number_input("Max rows per day (per manufacturer)", min_value=1, value=5, step=1))
        seed_txt = st.text_input("Random seed (optional, repeatable output)", value="")

st.divider()
required = list(layouts.values())
ready = all(l is not None for l in required) and covered is not None
if not ready:
    st.caption("Upload the sample file(s) and pick a valid file date to enable generation.")

if st.button("Generate test file(s)", type="primary", disabled=not ready):
    try:
        seed = int(seed_txt) if seed_txt.strip() else None
    except ValueError:
        st.error("Seed must be a whole number.")
        st.stop()
    req = Request(
        version="V6" if is_jlr else "V5",
        layouts=layouts,
        file_date=file_date,
        non_working_days=non_working,
        manufacturers=codes,
        include_debits=include_debits,
        exceed_threshold=exceed,
        threshold=threshold,
        rows_per_day=(rows_min, rows_max),
        seed=seed,
    )
    try:
        st.session_state["result"] = generate(req)
    except (RequestError, CalendarError) as exc:
        st.session_state.pop("result", None)
        st.error(str(exc))

for f in st.session_state.get("result", []):
    with st.container(border=True):
        st.markdown(f"#### {f.label}: `{f.name}`")
        summary = pd.DataFrame(f.summary)
        for c in ("Credits", "Debits", "Net payment"):
            summary[c] = summary[c].map(lambda v: f"{v:,.2f}")
        st.dataframe(summary, hide_index=True, use_container_width=True)
        st.download_button(f"Download {f.name}", f.data, file_name=f.name, key=f"dl_{f.name}")
        with st.expander("Preview"):
            if f.name.lower().endswith((".xlsx", ".xlsm")):
                st.dataframe(pd.read_excel(io.BytesIO(f.data), header=None).head(60), use_container_width=True)
            else:
                st.code("\n".join(f.data.decode("utf-8", "replace").splitlines()[:60]), language=None)

if len(st.session_state.get("result", [])) > 1:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for f in st.session_state["result"]:
            z.writestr(f.name, f.data)
    st.download_button("Download all as ZIP", buf.getvalue(), file_name="payment_files.zip", mime="application/zip")
