from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass, asdict
from datetime import date, datetime, time
from typing import Iterable, Optional

from pypdf import PdfReader
import xlsxwriter


FILING_CUTOFF_DAY = 18
MONEY_TOKEN_RE = re.compile(r"(?<![A-Za-z0-9])(?:-?\d[\d,]*(?:\.\d+)?|-)(?![A-Za-z0-9])")

MONTH_NO = {
    "april": 4, "apr": 4,
    "may": 5,
    "june": 6, "jun": 6,
    "july": 7, "jul": 7,
    "august": 8, "aug": 8,
    "september": 9, "sep": 9, "sept": 9,
    "october": 10, "oct": 10,
    "november": 11, "nov": 11,
    "december": 12, "dec": 12,
    "january": 1, "jan": 1,
    "february": 2, "feb": 2,
    "march": 3, "mar": 3,
}

QUARTERS = {
    "apr-jun": (4, 3), "april-june": (4, 3),
    "jul-sep": (7, 3), "july-september": (7, 3),
    "oct-dec": (10, 3), "october-december": (10, 3),
    "jan-mar": (1, 3), "january-march": (1, 3),
}


@dataclass
class GSTR3BReturn:
    source_file: str
    fiscal_year: str
    period: str
    gstin: str
    legal_name: str
    trade_name: str
    arn: str
    filing_date: Optional[datetime]
    a_taxable: float
    b_zero_rated: float
    c_nil_exempt: float
    d_inward_rcm: float
    e_non_gst_outward: float
    igst: float
    cgst: float
    sgst: float
    cess: float

    @property
    def turnover(self) -> float:
        # User-approved credit turnover rule: Table 3.1 A + B + C only.
        return self.a_taxable + self.b_zero_rated + self.c_nil_exempt

    def to_dict(self) -> dict:
        out = asdict(self)
        out["turnover"] = self.turnover
        return out


def _amount(token: str | None) -> float:
    if not token or token.strip() == "-":
        return 0.0
    try:
        return float(token.replace(",", ""))
    except Exception:
        return 0.0


def _first_group(text: str, pattern: str, default: str = "") -> str:
    match = re.search(pattern, text, flags=re.I | re.M)
    return match.group(1).strip() if match else default


def _row_numbers(table_text: str, start_pattern: str, end_pattern: str | None = None) -> list[float]:
    if end_pattern:
        pattern = rf"{start_pattern}(.*?)(?={end_pattern})"
    else:
        pattern = rf"{start_pattern}(.*)$"
    match = re.search(pattern, table_text, flags=re.I | re.S)
    if not match:
        return []
    tokens = MONEY_TOKEN_RE.findall(match.group(1))
    return [_amount(x) for x in tokens]


def parse_gstr3b_pdf(file_bytes: bytes, source_file: str = "uploaded.pdf") -> GSTR3BReturn:
    reader = PdfReader(io.BytesIO(file_bytes))
    text = "\n".join((page.extract_text() or "") for page in reader.pages)

    if len(text.strip()) < 150:
        raise ValueError("No usable text found. Please upload the original text-based GST portal GSTR-3B PDF.")

    fiscal_year = _first_group(text, r"^\s*Year\s+(\d{4}-\d{2})\s*$")
    period = _first_group(text, r"^\s*Period\s+([^\r\n]+?)\s*$")
    gstin = _first_group(text, r"GSTIN\s+of\s+the\s+supplier\s+([0-9A-Z]{15})")
    legal_name = _first_group(text, r"2\(a\)\.\s*Legal name of the registered person\s+([^\r\n]+)")
    trade_name = _first_group(text, r"2\(b\)\.\s*Trade name, if any\s+([^\r\n]+)")
    arn = _first_group(text, r"2\(c\)\.\s*ARN\s+([^\r\n]+)")
    filing_date_text = _first_group(text, r"2\(d\)\.\s*Date of ARN\s+(\d{1,2}/\d{1,2}/\d{4})")

    if not fiscal_year or not period or not gstin:
        raise ValueError("This does not look like a standard GST portal GSTR-3B PDF.")

    filing_date = None
    if filing_date_text:
        try:
            filing_date = datetime.strptime(filing_date_text, "%d/%m/%Y")
        except ValueError:
            pass

    table_match = re.search(
        r"3\.1\s+Details of Outward supplies and inward supplies liable to reverse charge"
        r"(.*?)(?=3\.1\.1\s+Details of Supplies|3\.2\s+Out of supplies)",
        text,
        flags=re.I | re.S,
    )
    if not table_match:
        raise ValueError("Table 3.1 could not be found in this GSTR-3B PDF.")

    table = table_match.group(1)
    a = _row_numbers(table, r"\(\s*a\s*\)\s*Outward taxable supplies", r"\(\s*b\s*\)")
    b = _row_numbers(table, r"\(\s*b\s*\)\s*Outward taxable supplies", r"\(\s*c\s*\)")
    c = _row_numbers(table, r"\(\s*c\s*\)\s*Other outward supplies", r"\(\s*d\s*\)")
    d = _row_numbers(table, r"\(\s*d\s*\)\s*Inward supplies", r"\(\s*e\s*\)")
    e = _row_numbers(table, r"\(\s*e\s*\)\s*Non-GST outward supplies")

    if not a:
        raise ValueError("Table 3.1 row (A) could not be read.")

    return GSTR3BReturn(
        source_file=source_file,
        fiscal_year=fiscal_year,
        period=period.strip(),
        gstin=gstin,
        legal_name=legal_name,
        trade_name=trade_name if trade_name and trade_name != "-" else legal_name,
        arn=arn,
        filing_date=filing_date,
        a_taxable=a[0] if len(a) > 0 else 0.0,
        b_zero_rated=b[0] if len(b) > 0 else 0.0,
        c_nil_exempt=c[0] if len(c) > 0 else 0.0,
        d_inward_rcm=d[0] if len(d) > 0 else 0.0,
        e_non_gst_outward=e[0] if len(e) > 0 else 0.0,
        igst=a[1] if len(a) > 1 else 0.0,
        cgst=a[2] if len(a) > 2 else 0.0,
        sgst=a[3] if len(a) > 3 else 0.0,
        cess=a[4] if len(a) > 4 else 0.0,
    )


def add_months(d: date, months: int) -> date:
    idx = d.year * 12 + d.month - 1 + months
    return date(idx // 12, idx % 12 + 1, 1)


def fiscal_year_for_month(d: date) -> str:
    start = d.year if d.month >= 4 else d.year - 1
    return f"{start}-{str(start + 1)[-2:]}"


def month_label(d: date) -> str:
    return d.strftime("%b-%y").upper()


def expected_filing_date(period_month: date, cutoff_day: int = FILING_CUTOFF_DAY) -> date:
    nxt = add_months(period_month, 1)
    return date(nxt.year, nxt.month, cutoff_day)


def latest_expected_tax_month(as_of_date: date, cutoff_day: int = FILING_CUTOFF_DAY) -> date:
    current = date(as_of_date.year, as_of_date.month, 1)
    return add_months(current, -1 if as_of_date.day >= cutoff_day else -2)


def _normalized_period(period: str) -> str:
    value = period.strip().lower().replace("–", "-").replace("—", "-")
    value = re.sub(r"\s+", "-", value)
    value = re.sub(r"-+", "-", value)
    return value


def return_covered_months(item: GSTR3BReturn) -> list[date]:
    start_year = int(item.fiscal_year[:4])
    p = _normalized_period(item.period)

    quarter = QUARTERS.get(p)
    if quarter:
        month, count = quarter
    else:
        p_word = item.period.strip().lower()
        month = MONTH_NO.get(p_word) or MONTH_NO.get(p_word[:3])
        count = 1
        if month is None:
            raise ValueError(f"Unrecognized GSTR-3B period: {item.period}")

    year = start_year if month >= 4 else start_year + 1
    first = date(year, month, 1)
    return [add_months(first, i) for i in range(count)]


def dedupe_returns(items: Iterable[GSTR3BReturn]) -> tuple[list[GSTR3BReturn], list[GSTR3BReturn]]:
    chosen: dict[tuple[date, ...], GSTR3BReturn] = {}
    duplicates: list[GSTR3BReturn] = []
    for item in items:
        key = tuple(return_covered_months(item))
        old = chosen.get(key)
        if old is None:
            chosen[key] = item
            continue
        old_date = old.filing_date or datetime.min
        new_date = item.filing_date or datetime.min
        if new_date >= old_date:
            duplicates.append(old)
            chosen[key] = item
        else:
            duplicates.append(item)
    clean = sorted(chosen.values(), key=lambda x: return_covered_months(x)[0])
    return clean, duplicates


@dataclass
class WindowAnalysis:
    months: int
    as_of_date: date
    latest_expected_month: date
    required_months: list[date]
    selected_returns: list[GSTR3BReturn]
    outside_returns: list[GSTR3BReturn]
    duplicates: list[GSTR3BReturn]
    coverage: dict[date, list[GSTR3BReturn]]
    missing_months: list[date]
    boundary_returns: list[GSTR3BReturn]

    @property
    def start_month(self) -> date:
        return self.required_months[0]

    @property
    def end_month(self) -> date:
        return self.required_months[-1]

    @property
    def found_month_count(self) -> int:
        return len(self.required_months) - len(self.missing_months)

    @property
    def is_complete(self) -> bool:
        return not self.missing_months and not self.boundary_returns

    @property
    def turnover(self) -> float:
        return sum(x.turnover for x in self.selected_returns)


def analyze_window(items: Iterable[GSTR3BReturn], as_of_date: date, months: int) -> WindowAnalysis:
    clean, duplicates = dedupe_returns(items)
    latest = latest_expected_tax_month(as_of_date)
    required = [add_months(latest, i) for i in range(-(months - 1), 1)]
    required_set = set(required)
    coverage = {m: [] for m in required}
    selected, outside, boundary = [], [], []

    for item in clean:
        covered = set(return_covered_months(item))
        inside = covered & required_set
        if inside:
            selected.append(item)
            for m in inside:
                coverage[m].append(item)
            if covered - required_set:
                boundary.append(item)
        else:
            outside.append(item)

    missing = [m for m in required if not coverage[m]]
    selected.sort(key=lambda x: return_covered_months(x)[0])
    outside.sort(key=lambda x: return_covered_months(x)[0])

    return WindowAnalysis(
        months=months,
        as_of_date=as_of_date,
        latest_expected_month=latest,
        required_months=required,
        selected_returns=selected,
        outside_returns=outside,
        duplicates=duplicates,
        coverage=coverage,
        missing_months=missing,
        boundary_returns=boundary,
    )


def _safe_filename(text: str) -> str:
    text = re.sub(r"[^A-Za-z0-9._ -]+", "", text or "GSTR3B")
    return re.sub(r"\s+", "_", text.strip())[:70] or "GSTR3B"


def build_excel(items: Iterable[GSTR3BReturn], as_of_date: date, months: int) -> tuple[bytes, str, WindowAnalysis]:
    source_items = list(items)
    if not source_items:
        raise ValueError("No GSTR-3B returns were supplied.")
    gstins = {x.gstin for x in source_items}
    if len(gstins) != 1:
        raise ValueError("One Excel report can contain only one GSTIN.")

    analysis = analyze_window(source_items, as_of_date, months)
    all_clean, _ = dedupe_returns(source_items)
    trade = next((x.trade_name for x in all_clean if x.trade_name), all_clean[0].legal_name)
    legal = next((x.legal_name for x in all_clean if x.legal_name), "")
    gstin = all_clean[0].gstin

    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"in_memory": True})

    light_blue = "#9DC3E6"
    dark_green = "#274E13"
    medium_green = "#385723"
    blue = "#2F75B5"
    pale_green = "#E2F0D9"
    pale_red = "#F4CCCC"
    pale_yellow = "#FFF2CC"
    gray = "#D9D9D9"
    white = "#FFFFFF"
    red = "#C00000"
    green = "#006100"

    title = wb.add_format({"bold": True, "font_size": 14, "align": "center", "valign": "vcenter", "bg_color": light_blue, "border": 1})
    subtitle = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "bg_color": "#DDEBF7", "border": 1})
    hdr = wb.add_format({"bold": True, "font_color": white, "bg_color": dark_green, "align": "center", "valign": "vcenter", "text_wrap": True, "border": 1})
    cell = wb.add_format({"border": 1, "align": "center", "valign": "vcenter"})
    money = wb.add_format({"border": 1, "align": "right", "valign": "vcenter", "num_format": "#,##0.00"})
    date_fmt = wb.add_format({"border": 1, "align": "center", "num_format": "dd-mmm-yyyy"})
    found = wb.add_format({"border": 1, "align": "center", "bold": True, "bg_color": pale_green, "font_color": green})
    missing = wb.add_format({"border": 1, "align": "center", "bold": True, "bg_color": pale_red, "font_color": red})
    kpi_label = wb.add_format({"bold": True, "font_color": white, "bg_color": dark_green, "align": "center", "valign": "vcenter", "border": 1})
    kpi_value = wb.add_format({"bold": True, "bg_color": pale_yellow, "align": "center", "valign": "vcenter", "border": 1, "num_format": "#,##0.00"})
    total = wb.add_format({"bold": True, "font_color": white, "bg_color": medium_green, "align": "right", "valign": "vcenter", "border": 1, "num_format": "#,##0.00"})
    total_text = wb.add_format({"bold": True, "font_color": white, "bg_color": medium_green, "align": "center", "valign": "vcenter", "border": 1})
    note = wb.add_format({"italic": True, "font_color": "#666666", "text_wrap": True})
    source = wb.add_format({"border": 1, "align": "left", "valign": "vcenter", "text_wrap": True})
    fy_fill = wb.add_format({"border": 1, "align": "center", "valign": "vcenter", "bg_color": pale_green})

    # Sheet 1: GST Calculation
    ws = wb.add_worksheet("GST Calculation")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 6)
    ws.set_column("B:B", 13)
    ws.set_column("C:C", 12)
    ws.set_column("D:D", 14)
    ws.set_column("E:H", 18)
    ws.set_column("I:I", 20)
    ws.set_column("J:J", 24)
    ws.set_column("K:K", 34)
    ws.merge_range("A1:K1", f"GSTR-3B TURNOVER REPORT - LAST {months} MONTHS", title)
    ws.merge_range("A2:K2", f"{trade} | GSTIN: {gstin}", subtitle)
    ws.merge_range("A3:K3", "TURNOVER RULE: Table 3.1 (A) + (B) + (C) only. D (Inward RCM) and E (Non-GST outward) are excluded.", note)

    ws.write("A5", "Calculation Date", kpi_label)
    ws.write_datetime("B5", datetime.combine(as_of_date, time()), date_fmt)
    ws.write("C5", "Required Window", kpi_label)
    ws.merge_range("D5:E5", f"{month_label(analysis.start_month)} to {month_label(analysis.end_month)}", cell)
    ws.write("F5", "Completeness", kpi_label)
    ws.write("G5", f"{analysis.found_month_count}/{months}", found if analysis.found_month_count == months else missing)
    ws.write("H5", "Turnover", kpi_label)
    ws.merge_range("I5:K5", analysis.turnover, kpi_value)

    headers = ["#", "Tax Month", "FY", "Expected By", "Status", "A - Taxable", "B - Zero Rated", "C - Nil/Exempt", "A+B+C Turnover", "Covered By", "Source File"]
    for col, h in enumerate(headers):
        ws.write(6, col, h, hdr)

    quarter_written: set[int] = set()
    for idx, m in enumerate(analysis.required_months, start=1):
        r = 6 + idx
        covers = analysis.coverage[m]
        ws.write_number(r, 0, idx, cell)
        ws.write(r, 1, month_label(m), cell)
        ws.write(r, 2, fiscal_year_for_month(m), fy_fill)
        ws.write_datetime(r, 3, datetime.combine(expected_filing_date(m), time()), date_fmt)
        if not covers:
            ws.write(r, 4, "MISSING", missing)
            for col in range(5, 9):
                ws.write_number(r, col, 0, money)
            ws.write(r, 9, "Required GSTR-3B not uploaded", missing)
            ws.write_blank(r, 10, None, missing)
            continue

        item = covers[0]
        covered_months = return_covered_months(item)
        is_quarterly = len(covered_months) > 1
        ws.write(r, 4, "FOUND" if not is_quarterly else "COVERED BY QUARTER", found)

        # A quarterly return's amount is written once only, on its first required covered month.
        write_amount = True
        if is_quarterly:
            marker = id(item)
            if marker in quarter_written:
                write_amount = False
            else:
                quarter_written.add(marker)

        vals = [item.a_taxable, item.b_zero_rated, item.c_nil_exempt, item.turnover] if write_amount else [0, 0, 0, 0]
        for col, val in enumerate(vals, start=5):
            ws.write_number(r, col, val, money)
        ws.write(r, 9, f"{item.period} ({item.fiscal_year})", cell)
        ws.write(r, 10, item.source_file, source)

    total_row = 7 + months
    ws.merge_range(total_row, 0, total_row, 7, f"TOTAL TURNOVER - LAST {months} MONTHS (A+B+C ONLY)", total_text)
    ws.write_number(total_row, 8, analysis.turnover, total)
    ws.merge_range(total_row, 9, total_row, 10, "COMPLETE" if analysis.is_complete else "CHECK MISSING / BOUNDARY RETURNS", total_text)

    if analysis.boundary_returns:
        ws.merge_range(total_row + 2, 0, total_row + 2, 10, "WARNING: A quarterly return overlaps the edge of the selected window, so exact trailing-period turnover cannot be split from that quarter.", note)
    if analysis.duplicates:
        ws.merge_range(total_row + 3, 0, total_row + 3, 10, f"Duplicate periods detected: {len(analysis.duplicates)}. The latest filed copy was used.", note)
    ws.freeze_panes(7, 0)

    # Sheet 2: Source Detail
    src = wb.add_worksheet("Source Detail")
    src.hide_gridlines(2)
    src.set_column("A:B", 13)
    src.set_column("C:C", 14)
    src.set_column("D:D", 20)
    src.set_column("E:J", 18)
    src.set_column("K:M", 16)
    src.set_column("N:N", 20)
    src.set_column("O:O", 34)
    src.merge_range("A1:O1", "GSTR-3B SOURCE DETAIL / AUDIT TRAIL", title)
    src.merge_range("A2:O2", f"{trade} | {legal} | GSTIN {gstin}", subtitle)
    src_headers = ["FY", "Period", "Filing Date", "ARN", "A - Taxable", "B - Zero Rated", "C - Nil/Exempt", "D - Inward RCM", "E - Non-GST", "Turnover A+B+C", "IGST", "CGST", "SGST", "Cess", "Source File"]
    for col, h in enumerate(src_headers):
        src.write(3, col, h, hdr)
    for i, item in enumerate(all_clean, start=4):
        src.write(i, 0, item.fiscal_year, cell)
        src.write(i, 1, item.period, cell)
        if item.filing_date:
            src.write_datetime(i, 2, item.filing_date, date_fmt)
        else:
            src.write_blank(i, 2, None, cell)
        src.write(i, 3, item.arn, cell)
        nums = [item.a_taxable, item.b_zero_rated, item.c_nil_exempt, item.d_inward_rcm, item.e_non_gst_outward, item.turnover, item.igst, item.cgst, item.sgst, item.cess]
        for j, value in enumerate(nums, start=4):
            src.write_number(i, j, value, money)
        src.write(i, 14, item.source_file, source)
    src.freeze_panes(4, 0)

    # Sheet 3: FY Summary
    fy = wb.add_worksheet("FY Summary")
    fy.hide_gridlines(2)
    fy.set_column("A:A", 16)
    fy.set_column("B:B", 18)
    fy.set_column("C:F", 20)
    fy.merge_range("A1:F1", "FINANCIAL YEAR SUMMARY - SELECTED WINDOW", title)
    fy_headers = ["Financial Year", "Returns Count", "A - Taxable", "B - Zero Rated", "C - Nil/Exempt", "Turnover A+B+C"]
    for col, h in enumerate(fy_headers):
        fy.write(2, col, h, hdr)
    fy_names = []
    for m in analysis.required_months:
        name = fiscal_year_for_month(m)
        if name not in fy_names:
            fy_names.append(name)
    for row_idx, name in enumerate(fy_names, start=3):
        selected = [x for x in analysis.selected_returns if x.fiscal_year == name]
        fy.write(row_idx, 0, name, fy_fill)
        fy.write_number(row_idx, 1, len(selected), cell)
        values = [sum(x.a_taxable for x in selected), sum(x.b_zero_rated for x in selected), sum(x.c_nil_exempt for x in selected), sum(x.turnover for x in selected)]
        for col_idx, val in enumerate(values, start=2):
            fy.write_number(row_idx, col_idx, val, money)
    fy_total_row = 3 + len(fy_names)
    fy.merge_range(fy_total_row, 0, fy_total_row, 4, "SELECTED WINDOW TOTAL", total_text)
    fy.write_number(fy_total_row, 5, analysis.turnover, total)

    wb.close()
    output.seek(0)
    filename = f"GSTR3B_{months}M_{_safe_filename(trade)}_{gstin}.xlsx"
    return output.getvalue(), filename, analysis


def build_multi_gstin_zip(grouped: dict[str, list[GSTR3BReturn]], as_of_date: date, months: int) -> tuple[bytes, str]:
    zbuf = io.BytesIO()
    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
        for _, items in grouped.items():
            data, filename, _ = build_excel(items, as_of_date, months)
            zf.writestr(filename, data)
    return zbuf.getvalue(), f"GSTR3B_{months}M_Reports.zip"
