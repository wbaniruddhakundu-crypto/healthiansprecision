"""Healthians day-wise inter-run precision workbook generator.

Reads text-based Siemens assay PDFs, groups reports by Day 1..Day 5,
then writes up to five numeric results per analyte/day into a supplied
Excel precision template without changing its date cells or merged layout.
"""

from __future__ import annotations

import re
from collections import defaultdict
from io import BytesIO
from typing import Dict, Iterable, List, Tuple

import fitz  # PyMuPDF
import openpyxl
import streamlit as st
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Healthians Inter-Run Precision", layout="wide")
st.title("Healthians – Inter-Run Precision Populator")
st.caption(
    "Upload your existing Excel precision template and all daily assay PDFs. "
    "The app keeps the dates and template layout, and fills up to five results per analyte per day."
)

# Include common units used in Siemens immunoassay and chemistry reports.
UNIT_PATTERN = (
    r"(?:ng/mL|ng/dL|pg/mL|pg/dL|ug/dL|µg/dL|μg/dL|mg/dL|g/dL|U/mL|IU/mL|"
    r"mIU/mL|uIU/mL|µIU/mL|μIU/mL|mIU/L|IU/L|pmol/L|nmol/L|mmol/L|"
    r"µmol/L|μmol/L|umol/L|ng/L|pg/L|%|mmol/mol|g/L|mg/L)"
)
RESULT_RE = re.compile(
    rf"(?<![A-Za-z0-9])(-?\d+(?:,\d{{3}})*(?:\.\d+)?)\s*({UNIT_PATTERN})(?![A-Za-z])",
    re.IGNORECASE,
)


def day_from_name(filename: str):
    """Return day number 1..5 from a filename, or None if not identifiable."""
    match = re.search(r"\bday\s*[-_ ]?([1-5])\b", filename, re.IGNORECASE)
    return int(match.group(1)) if match else None


def normalize_analyte_title(line: str) -> str:
    return re.sub(r"[^A-Za-z0-9_+.-]", "", line.strip())


def parse_pdf(pdf_bytes: bytes) -> Dict[str, List[Tuple[float, str]]]:
    """Extract result/unit pairs grouped under analyte headings.

    This parser expects selectable text in the PDF. Numeric RLU/MAU columns
    are ignored because only values accompanied by a recognized unit match.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    lines: List[str] = []
    try:
        for page in doc:
            lines.extend(page.get_text("text").splitlines())
    finally:
        doc.close()

    results: Dict[str, List[Tuple[float, str]]] = defaultdict(list)
    current_analyte = None

    for index, raw_line in enumerate(lines):
        line = raw_line.strip()
        if not line:
            continue

        lookahead = " ".join(part.strip().lower() for part in lines[index + 1:index + 10])
        title = normalize_analyte_title(line)
        is_clock_time = bool(re.fullmatch(r"\d{1,2}:\d{2}\s*(?:AM|PM)", line, re.IGNORECASE))
        is_date = bool(re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", line))
        if (
            title
            and not is_clock_time
            and not is_date
            and len(title) <= 24
            and "patient name" in lookahead
            and "sid" in lookahead
            and "result" in lookahead
            and title.upper() not in {"ASSAYREPORT", "ROUTINE", "TUMOR", "CONTROL"}
        ):
            current_analyte = title
            continue

        if current_analyte:
            for match in RESULT_RE.finditer(line):
                raw_value = match.group(1).replace(",", "")
                try:
                    value = float(raw_value)
                except ValueError:
                    continue
                unit = match.group(2)
                results[current_analyte].append((value, unit))

    return dict(results)


def unique_sheet_name(name: str, used: set) -> str:
    cleaned = re.sub(r"[\\/*?:\[\]]", "_", name)[:31] or "Analyte"
    candidate = cleaned
    suffix_number = 1
    while candidate in used:
        suffix = f"_{suffix_number}"
        candidate = cleaned[:31 - len(suffix)] + suffix
        suffix_number += 1
    used.add(candidate)
    return candidate


def write_analyte_sheet(ws, analyte: str, daily_values: Dict[int, List[Tuple[float, str]]]):
    """Populate only known non-merged destination cells; preserve template merges."""
    # A4:C4 is a merged label cell. The analyte name belongs in D4:F4,
    # whose writable anchor is D4. Never write into merged-cell placeholders.
    ws["D4"] = analyte
    ws["C6"] = analyte

    unit = ""
    for day in range(1, 6):
        if daily_values.get(day):
            unit = daily_values[day][0][1]
            break
    if unit:
        ws["D6"] = f"Unit: ({unit})"

    # Day 1..5 are columns B..F. Dates in row 10 are intentionally untouched.
    for day in range(1, 6):
        column = day + 1
        letter = get_column_letter(column)
        values = daily_values.get(day, [])[:5]
        for replicate in range(5):
            ws.cell(row=11 + replicate, column=column).value = (
                values[replicate][0] if replicate < len(values) else None
            )

        # Day-wise statistics for the five replicate cells.
        ws.cell(17, column).value = f"=COUNT({letter}11:{letter}15)"
        ws.cell(18, column).value = f'=IFERROR(AVERAGE({letter}11:{letter}15),"")'
        ws.cell(19, column).value = f'=IFERROR(STDEV({letter}11:{letter}15),"")'
        ws.cell(20, column).value = f'=IFERROR({letter}19/{letter}18,"")'

    # Overall inter-run summary across all five days.
    ws["E25"] = "=COUNT(B11:F15)"
    ws["E26"] = '=IFERROR(AVERAGE(B11:F15),"")'
    ws["E27"] = '=IFERROR(STDEV(B11:F15),"")'
    ws["E28"] = '=IFERROR(E27/E26,"")'


def build_workbook(template_bytes: bytes, uploaded_pdfs: Iterable[Tuple[str, bytes]]):
    workbook = openpyxl.load_workbook(BytesIO(template_bytes))
    template_sheet = workbook[workbook.sheetnames[0]]

    # Combine all PDFs belonging to the same day. This supports separate
    # routine, tumor-marker, vitamin-D, and other panel reports for one day.
    combined_by_day: Dict[int, Dict[str, List[Tuple[float, str]]]] = {
        day: defaultdict(list) for day in range(1, 6)
    }
    notes: List[str] = []
    detected_files: Dict[int, List[str]] = defaultdict(list)

    for filename, content in uploaded_pdfs:
        day = day_from_name(filename)
        if day is None:
            notes.append(f"Day number not found in filename; skipped: {filename}")
            continue
        parsed = parse_pdf(content)
        detected_files[day].append(filename)
        for analyte, values in parsed.items():
            combined_by_day[day][analyte].extend(values)

    analytes = sorted(
        {analyte for day_data in combined_by_day.values() for analyte in day_data},
        key=str.upper,
    )
    if not analytes:
        raise ValueError(
            "No analyte results were detected. Confirm that the PDFs contain selectable text; "
            "scanned PDFs need OCR and are not supported by this version."
        )

    used_names = set()
    for index, analyte in enumerate(analytes):
        if index == 0:
            ws = template_sheet
        else:
            ws = workbook.copy_worksheet(template_sheet)

        ws.title = unique_sheet_name(analyte, used_names)
        daily_values = {
            day: list(combined_by_day[day].get(analyte, []))
            for day in range(1, 6)
        }
        write_analyte_sheet(ws, analyte, daily_values)

    # Ask Excel/LibreOffice to recalculate formulas when the workbook opens.
    try:
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"
    except AttributeError:
        pass

    output = BytesIO()
    workbook.save(output)
    detected_days = sorted(day for day, files in detected_files.items() if files)
    file_summary = {day: files for day, files in sorted(detected_files.items())}
    return output.getvalue(), analytes, notes, detected_days, file_summary


with st.sidebar:
    st.header("Workflow")
    st.markdown(
        "1. Upload the original `.xlsx` template.\n"
        "2. Upload all Day 1–Day 5 PDF reports.\n"
        "3. Generate and download the completed workbook."
    )
    st.caption("Multiple PDFs for the same day are combined automatically.")

template_upload = st.file_uploader(
    "1. Upload Excel precision template (.xlsx)", type=["xlsx"]
)
pdf_uploads = st.file_uploader(
    "2. Upload all Day 1–Day 5 assay PDFs",
    type=["pdf"],
    accept_multiple_files=True,
    help="Filename must contain Day 1, Day 2, Day 3, Day 4, or Day 5.",
)

if st.button("Generate completed precision workbook", type="primary"):
    if template_upload is None:
        st.error("Please upload your Excel precision template.")
    elif not pdf_uploads:
        st.error("Please upload at least one daily assay PDF.")
    else:
        try:
            pdf_payload = [(file.name, file.getvalue()) for file in pdf_uploads]
            output_bytes, analytes, notes, days, file_summary = build_workbook(
                template_upload.getvalue(), pdf_payload
            )
            st.success(
                f"Workbook generated: {len(analytes)} analyte sheets. "
                f"Reports detected for Day(s): {', '.join(map(str, days)) or 'None'}."
            )
            with st.expander("Uploaded reports grouped by day", expanded=True):
                for day in range(1, 6):
                    names = file_summary.get(day, [])
                    st.write(f"**Day {day}:** " + (", ".join(names) if names else "No report uploaded"))

            st.write("**Analytes detected:** " + ", ".join(analytes))
            st.download_button(
                "Download completed Excel workbook",
                data=output_bytes,
                file_name="Healthians_Inter_Run_Precision_Daywise.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
            if notes:
                st.warning("\n".join(notes))
            st.info(
                "The first five result values per analyte/day are populated. Missing days or "
                "replicates remain blank. Existing date cells and merged template layout are "
                "preserved. RLU/MAU values are ignored. Please verify the workbook against "
                "the source PDFs before formal validation."
            )
        except Exception as error:
            st.error(f"Could not generate the workbook: {error}")
            st.exception(error)
