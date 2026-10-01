from __future__ import annotations

from collections import defaultdict
from datetime import date

import streamlit as st

APP_VERSION = "Online Deployment 3.1"

# Render something immediately so deployment/runtime problems are visible instead
# of presenting a blank page.
st.set_page_config(
    page_title="GSTR-3B Turnover Calculator",
    page_icon="📊",
    layout="wide",
)

st.title("GSTR-3B Automatic Turnover Calculator")
st.success(f"App loaded successfully — {APP_VERSION}")
st.caption(
    "Upload original GST portal GSTR-3B PDFs. The tool checks missing returns, "
    "calculates turnover and creates an Excel report."
)

try:
    from gstr3b_core import (
        FILING_CUTOFF_DAY,
        analyze_window,
        build_excel,
        build_multi_gstin_zip,
        fiscal_year_for_month,
        month_label,
        parse_gstr3b_pdf,
    )
except Exception as exc:
    st.error("The GST calculation engine could not load.")
    st.write("Open Streamlit → Manage app → Logs to see the deployment error.")
    st.exception(exc)
    st.stop()

st.info(
    "**TURNOVER RULE:** Table 3.1 **A + B + C only**. "
    "A = outward taxable supplies, B = zero-rated outward supplies, "
    "C = nil-rated/exempt outward supplies. "
    "D (Inward RCM) and E (Non-GST outward) are excluded."
)

with st.expander("How missing GSTR-3B months are decided", expanded=False):
    st.write(
        f"The filing cutoff used by this tool is the **{FILING_CUTOFF_DAY}th of the following month**. "
        "Before the 18th, the latest expected tax month is two months back. "
        "On or after the 18th, the immediately previous tax month becomes expected."
    )
    st.write(
        "Example: on 15-Sep-2026, July-2026 is the latest expected month, so a 12-month "
        "report needs Aug-2025 through Jul-2026."
    )

c1, c2 = st.columns(2)
with c1:
    period_label = st.radio(
        "Report period",
        ["Last 12 months", "Last 24 months"],
        index=0,
        horizontal=True,
    )
with c2:
    calculation_date = st.date_input(
        "Calculation date",
        value=date.today(),
        help=(
            "The selected date controls which GSTR-3B months are required under the "
            "18th-of-next-month rule."
        ),
    )

months = 12 if period_label == "Last 12 months" else 24

uploads = st.file_uploader(
    "Upload GSTR-3B PDF files",
    type=["pdf"],
    accept_multiple_files=True,
    help=(
        "Upload all available original GST portal GSTR-3B PDFs. Extra older/newer returns "
        "are allowed; the tool selects the required window automatically."
    ),
)

if not uploads:
    st.warning("Upload GSTR-3B PDF files to start.")
    st.stop()

parsed = []
errors = []
progress = st.progress(0)
status_box = st.empty()

for idx, uploaded in enumerate(uploads, start=1):
    status_box.write(f"Reading {idx} of {len(uploads)}: {uploaded.name}")
    try:
        parsed.append(parse_gstr3b_pdf(uploaded.getvalue(), uploaded.name))
    except Exception as exc:
        errors.append((uploaded.name, str(exc)))
    progress.progress(int(idx * 100 / len(uploads)))

status_box.empty()
progress.empty()

if errors:
    st.error(f"{len(errors)} file(s) could not be read.")
    with st.expander("Show PDF reading errors"):
        for name, message in errors:
            st.write(f"• **{name}** — {message}")

if not parsed:
    st.error("No valid GSTR-3B returns were detected.")
    st.stop()

by_gstin: dict[str, list] = defaultdict(list)
for item in parsed:
    by_gstin[item.gstin].append(item)

st.success(f"Successfully read {len(parsed)} return(s) for {len(by_gstin)} GSTIN(s).")

for gstin, items in by_gstin.items():
    analysis = analyze_window(items, calculation_date, months)
    trade_name = items[0].trade_name or items[0].legal_name

    with st.container(border=True):
        st.subheader(trade_name)
        st.caption(f"GSTIN: {gstin}")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric(
            "Required window",
            f"{month_label(analysis.start_month)} → {month_label(analysis.end_month)}",
        )
        m2.metric("Required months found", f"{analysis.found_month_count}/{months}")
        m3.metric("Missing months", len(analysis.missing_months))
        m4.metric("Turnover (A+B+C)", f"₹{analysis.turnover:,.2f}")

        if analysis.missing_months:
            st.error(
                "Missing GSTR-3B: "
                + ", ".join(month_label(x) for x in analysis.missing_months)
            )
        else:
            st.success(f"Complete {months}/{months} month coverage found.")

        if analysis.boundary_returns:
            st.warning(
                "A quarterly return overlaps the edge of the selected window. "
                "The quarterly value cannot be split accurately across the boundary."
            )

        if analysis.duplicates:
            st.warning(
                f"{len(analysis.duplicates)} duplicate period file(s) detected. "
                "The latest filed copy is used."
            )

        if analysis.outside_returns:
            st.caption(
                f"{len(analysis.outside_returns)} uploaded return(s) are outside this "
                f"{months}-month window and are not counted in the turnover."
            )

        with st.expander("Month-by-month completeness"):
            rows = []
            for month in analysis.required_months:
                covers = analysis.coverage[month]
                rows.append(
                    {
                        "Tax Month": month_label(month),
                        "FY": fiscal_year_for_month(month),
                        "Status": "FOUND" if covers else "MISSING",
                        "Covered By": (
                            "; ".join(f"{x.period} ({x.fiscal_year})" for x in covers)
                            if covers
                            else "-"
                        ),
                    }
                )
            st.dataframe(rows, use_container_width=True, hide_index=True)

st.divider()

if st.button("Generate Excel Report", type="primary", use_container_width=True):
    try:
        if len(by_gstin) == 1:
            items = next(iter(by_gstin.values()))
            data, filename, _ = build_excel(items, calculation_date, months)
            st.download_button(
                "Download Excel Report",
                data=data,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            data, filename = build_multi_gstin_zip(dict(by_gstin), calculation_date, months)
            st.download_button(
                "Download All Reports as ZIP",
                data=data,
                file_name=filename,
                mime="application/zip",
                use_container_width=True,
            )
    except Exception as exc:
        st.error("Excel report generation failed.")
        st.exception(exc)

st.caption(
    "Use original text-based GSTR-3B PDFs downloaded from the GST portal. "
    "Scanned/image-only PDFs are rejected instead of guessing figures."
)
