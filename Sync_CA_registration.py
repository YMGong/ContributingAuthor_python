"""
update_smu.py

Workflow
--------
1. Download the latest Microsoft Forms Excel workbook manually.
2. Run this script.
3. Select the downloaded Forms workbook.
4. Select the existing SMU workbook as the template.
5. Choose where to save the updated workbook.

Output filename defaults to:
    02. WGI CA list.xlsx

Required package:
    pip install openpyxl

Key logic
---------
- Copy all source submissions to the destination first.
- Family names are standardized (uppercase; Latin accents removed) before matching.
- Same person = same destination B (family name) + C (first names),
  compared case-insensitively and ignoring leading/trailing/repeated spaces.
- If source K is ever "No" for a person, destination I is 0 on all their rows.
- If source L is ever empty for a person, destination J is empty on all their rows.
- If source M is ever empty for a person, destination K is empty on all their rows.
- Red font in destination B is applied only after reconciliation, based on final I == 1.
- Dashboard_externalCAs includes authors whose final destination K is empty.
- Dashboard does not double-count the same person in the same chapter.
- If duplicate person+chapter submissions conflict, the latest submission is chosen
  using destination A as the ordering value.
- Gender and region pie charts omit zero-count categories, show count + percentage,
  use bottom horizontal legends, set legend/data-label text to 11 pt, and show exact labels such as 5 (25%).
"""

import re
import unicodedata
import os
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from copy import copy
from datetime import datetime, date
from tkinter import Tk, filedialog, messagebox

from openpyxl import load_workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import (
    Paragraph,
    ParagraphProperties,
    CharacterProperties,
    RegularTextRun,
    RichTextProperties,
)
from openpyxl.styles import Font


# ============================================================
# SETTINGS
# ============================================================

DESTINATION_DATA_SHEET = "Data"
DASHBOARD_SHEET = "Dashboard_externalCAs"
CHART_DATA_SHEET = "ChartData"

SOURCE_FIRST_DATA_ROW = 2
DESTINATION_FIRST_DATA_ROW = 2

COLUMN_MAPPING = {
    "A": "A",
    "H": "B",
    "I": "C",
    "J": "D",
    "O": "E",
    "Q": "F",
    "P": "G",
    "W": "H",
    "K": "I",
    "L": "J",
    "M": "K",
    "N": "L",
    "R": "M",
    "S": "N",
    "T": "O",
    "U": "P",
    "V": "Q",
}

REGIONS = [
    "Africa",
    "Asia",
    "Europe",
    "North America, Central America and the Caribbean",
    "South America",
    "South-West Pacific",
]


# ============================================================
# FILE SELECTION
# ============================================================

def select_files():
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    source_file = filedialog.askopenfilename(
        title="1. Select the downloaded Microsoft Forms workbook",
        filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")]
    )
    if not source_file:
        root.destroy()
        return None, None, None

    template_file = filedialog.askopenfilename(
        title="2. Select the SMU destination/template workbook",
        filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")]
    )
    if not template_file:
        root.destroy()
        return None, None, None

    output_file = filedialog.asksaveasfilename(
        title="3. Choose where to save the updated workbook",
        defaultextension=".xlsx",
        initialfile="02. WGI CA list.xlsx",
        filetypes=[("Excel workbook", "*.xlsx")]
    )

    root.destroy()

    if not output_file:
        return None, None, None

    return source_file, template_file, output_file


# ============================================================
# GENERAL HELPERS
# ============================================================

def normalize_text(value):
    if value is None:
        return ""
    return str(value).strip()


def normalize_family_name(value):
    """
    Standardize Latin-script family names before duplicate matching:
    - remove Latin diacritics/accents
    - collapse repeated whitespace
    - convert to uppercase

    Non-Latin scripts are preserved rather than transliterated.
    """
    text = " ".join(normalize_text(value).split())
    decomposed = unicodedata.normalize("NFD", text)
    text = "".join(
        ch for ch in decomposed
        if not (unicodedata.combining(ch) and ord(ch) <= 0x036F)
    )
    return unicodedata.normalize("NFC", text).upper()


def normalized_name_part(value):
    """Normalize a name component for duplicate matching."""
    return " ".join(normalize_text(value).casefold().split())


def person_key_from_destination(ws, row):
    family = normalized_name_part(ws[f"B{row}"].value)
    first = normalized_name_part(ws[f"C{row}"].value)

    # Avoid treating entirely blank names as one person.
    if not family and not first:
        return None

    return (family, first)


def is_empty(value):
    return value is None or str(value).strip() == ""


def latest_sort_key(value, fallback_row=0):
    """
    Produce a sortable key for destination A.

    Handles numbers, Excel/Python dates, numeric strings, common date strings,
    and arbitrary text. Row number is a deterministic tie-breaker.
    """
    if value is None:
        return (0, "", fallback_row)

    if isinstance(value, datetime):
        return (4, value.timestamp(), fallback_row)

    if isinstance(value, date):
        return (4, datetime.combine(value, datetime.min.time()).timestamp(), fallback_row)

    if isinstance(value, (int, float)):
        return (3, float(value), fallback_row)

    text = normalize_text(value)

    # Numeric-looking ID/submission number.
    try:
        return (3, float(text), fallback_row)
    except ValueError:
        pass

    # Common ISO-ish date/time strings.
    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%m/%d/%Y",
    ):
        try:
            dt = datetime.strptime(text, fmt)
            return (4, dt.timestamp(), fallback_row)
        except ValueError:
            pass

    # Lexical fallback.
    return (2, text.casefold(), fallback_row)


# ============================================================
# TRANSFORMATIONS
# ============================================================

def convert_gender(value):
    text = normalize_text(value)
    if text.lower() == "male":
        return "M"
    if text.lower() == "female":
        return "F"
    if text.upper() in ("M", "F"):
        return text.upper()
    return value


def convert_yes_no(value):
    if value == 1:
        return 1
    if value == 0:
        return 0

    text = normalize_text(value).lower()
    if text == "yes":
        return 1
    if text == "no":
        return 0
    return value


def extract_chapter(value):
    if value is None:
        return None

    if isinstance(value, (int, float)):
        number = int(value)
        if 1 <= number <= 10:
            return number

    text = normalize_text(value)
    match = re.search(
        r"\bchapter\s*[-:]?\s*(10|[1-9])\b",
        text,
        flags=re.IGNORECASE
    )

    if match:
        return int(match.group(1))

    return value


# ============================================================
# FORMATTING HELPERS
# ============================================================

def copy_cell_style(source_cell, destination_cell):
    if source_cell.has_style:
        destination_cell._style = copy(source_cell._style)

    destination_cell.font = copy(source_cell.font)
    destination_cell.fill = copy(source_cell.fill)
    destination_cell.border = copy(source_cell.border)
    destination_cell.alignment = copy(source_cell.alignment)
    destination_cell.protection = copy(source_cell.protection)
    destination_cell.number_format = source_cell.number_format


def copy_template_row_style(ws, template_row, target_row):
    for col in range(1, 18):  # A:Q
        copy_cell_style(
            ws.cell(row=template_row, column=col),
            ws.cell(row=target_row, column=col)
        )

    if ws.row_dimensions[template_row].height is not None:
        ws.row_dimensions[target_row].height = ws.row_dimensions[template_row].height


# ============================================================
# SOURCE / DESTINATION WORKBOOK HELPERS
# ============================================================

def find_source_sheet(workbook):
    candidates = []

    for ws in workbook.worksheets:
        if ws.max_row >= SOURCE_FIRST_DATA_ROW and ws.max_column >= 23:
            candidates.append(ws)

    if not candidates:
        raise ValueError(
            "Could not find a Forms response worksheet containing at least columns A through W."
        )

    return max(candidates, key=lambda sheet: sheet.max_row)


def find_last_source_row(ws):
    for row in range(ws.max_row, SOURCE_FIRST_DATA_ROW - 1, -1):
        if not is_empty(ws[f"A{row}"].value):
            return row

    return SOURCE_FIRST_DATA_ROW - 1


def get_destination_sheet(workbook):
    if DESTINATION_DATA_SHEET not in workbook.sheetnames:
        raise ValueError(
            f'The destination workbook does not contain a sheet named '
            f'"{DESTINATION_DATA_SHEET}". Current sheets: {", ".join(workbook.sheetnames)}'
        )

    return workbook[DESTINATION_DATA_SHEET]


def clear_old_destination_values(ws):
    """
    Clear A:Q from row 2 downward, retaining formatting.
    """
    for row in range(DESTINATION_FIRST_DATA_ROW, ws.max_row + 1):
        for col in range(1, 18):
            ws.cell(row=row, column=col).value = None


# ============================================================
# COPY ALL SUBMISSIONS
# ============================================================

def transfer_data(source_ws, destination_ws):
    source_last_row = find_last_source_row(source_ws)

    if source_last_row < SOURCE_FIRST_DATA_ROW:
        raise ValueError("No Forms responses were found.")

    number_of_records = source_last_row - SOURCE_FIRST_DATA_ROW + 1
    style_template_row = DESTINATION_FIRST_DATA_ROW

    # Store source-side information needed for later reconciliation.
    source_metadata = {}

    for index, source_row in enumerate(
        range(SOURCE_FIRST_DATA_ROW, source_last_row + 1)
    ):
        destination_row = DESTINATION_FIRST_DATA_ROW + index

        if destination_row > destination_ws.max_row:
            copy_template_row_style(
                destination_ws,
                style_template_row,
                destination_row
            )

        for source_col, destination_col in COLUMN_MAPPING.items():
            value = source_ws[f"{source_col}{source_row}"].value

            if source_col == "J":
                value = convert_gender(value)
            elif source_col == "W":
                value = extract_chapter(value)
            elif source_col == "K":
                value = convert_yes_no(value)

            destination_ws[f"{destination_col}{destination_row}"].value = value

        # Standardize family name BEFORE duplicate-person identification.
        destination_ws[f"B{destination_row}"].value = normalize_family_name(
            destination_ws[f"B{destination_row}"].value
        )

        source_metadata[destination_row] = {
            "source_k": source_ws[f"K{source_row}"].value,
            "source_l": source_ws[f"L{source_row}"].value,
            "source_m": source_ws[f"M{source_row}"].value,
        }

    return number_of_records, source_metadata


# ============================================================
# RECONCILE DUPLICATE PEOPLE
# ============================================================

def reconcile_people(destination_ws, source_metadata, record_count):
    """
    Reconcile all submissions belonging to the same person.

    Same person = destination B + C.

    Rule 1:
      If source K is ever No for the person, destination I = 0
      on every row belonging to that person.

    Rule 2:
      If source L is ever empty for the person, destination J is empty
      on every row belonging to that person.

    Rule 3:
      If source M is ever empty for the person, destination K is empty
      on every row belonging to that person.

    The red-font rule for destination B is then applied using final I.
    """
    first_row = DESTINATION_FIRST_DATA_ROW
    last_row = first_row + record_count - 1

    people = {}

    for row in range(first_row, last_row + 1):
        key = person_key_from_destination(destination_ws, row)
        if key is None:
            # Treat nameless records independently.
            key = ("__unnamed__", row)

        people.setdefault(key, []).append(row)

    for rows in people.values():
        any_no = False
        any_source_l_empty = False
        any_source_m_empty = False

        for row in rows:
            meta = source_metadata.get(row, {})

            source_k = normalize_text(meta.get("source_k")).casefold()
            if source_k == "no":
                any_no = True

            if is_empty(meta.get("source_l")):
                any_source_l_empty = True

            if is_empty(meta.get("source_m")):
                any_source_m_empty = True

        if any_no:
            for row in rows:
                destination_ws[f"I{row}"].value = 0

        if any_source_l_empty:
            for row in rows:
                destination_ws[f"J{row}"].value = None

        if any_source_m_empty:
            for row in rows:
                destination_ws[f"K{row}"].value = None

    # Apply B font after I has reached its final reconciled state.
    normal_font = copy(destination_ws[f"B{DESTINATION_FIRST_DATA_ROW}"].font)

    for row in range(first_row, last_row + 1):
        b_cell = destination_ws[f"B{row}"]
        row_font = copy(normal_font)

        if destination_ws[f"I{row}"].value == 1:
            row_font.color = "FF0000"

        b_cell.font = row_font

    return people


# ============================================================
# DASHBOARD STATISTICS
# ============================================================

def select_unique_dashboard_rows(destination_ws, record_count):
    """
    Dashboard population:
      - destination K must be empty
      - same person is counted only once per chapter
      - if a person has multiple submissions for the same chapter,
        choose the latest row using destination A
    """
    first_row = DESTINATION_FIRST_DATA_ROW
    last_row = first_row + record_count - 1

    selected = {}

    for row in range(first_row, last_row + 1):
        # External CA filter after reconciliation.
        if not is_empty(destination_ws[f"K{row}"].value):
            continue

        chapter = destination_ws[f"H{row}"].value
        try:
            chapter = int(chapter)
        except (TypeError, ValueError):
            continue

        if chapter not in range(1, 11):
            continue

        person_key = person_key_from_destination(destination_ws, row)
        if person_key is None:
            # Avoid merging unrelated nameless records.
            person_key = ("__unnamed__", row)

        unique_key = (person_key, chapter)
        candidate_key = latest_sort_key(destination_ws[f"A{row}"].value, row)

        if unique_key not in selected:
            selected[unique_key] = (row, candidate_key)
        else:
            existing_row, existing_key = selected[unique_key]
            if candidate_key > existing_key:
                selected[unique_key] = (row, candidate_key)

    return [row for row, _ in selected.values()]


def recreate_chart_data_sheet(workbook, destination_ws, record_count):
    if CHART_DATA_SHEET in workbook.sheetnames:
        del workbook[CHART_DATA_SHEET]

    ws = workbook.create_sheet(CHART_DATA_SHEET)

    selected_rows = select_unique_dashboard_rows(destination_ws, record_count)

    stats = {
        chapter: {
            "M": 0,
            "F": 0,
            "regions": {region: 0 for region in REGIONS},
            "people": 0,
        }
        for chapter in range(1, 11)
    }

    for row in selected_rows:
        chapter = int(destination_ws[f"H{row}"].value)
        stats[chapter]["people"] += 1

        gender = normalize_text(destination_ws[f"D{row}"].value).upper()
        if gender == "M":
            stats[chapter]["M"] += 1
        elif gender == "F":
            stats[chapter]["F"] += 1

        region = normalize_text(destination_ws[f"G{row}"].value)
        if region in stats[chapter]["regions"]:
            stats[chapter]["regions"][region] += 1

    # Variable-length blocks; zero-count categories are omitted completely.
    chart_ranges = {}

    for chapter in range(1, 11):
        base = 1 + (chapter - 1) * 12

        gender_items = [
            ("Male", stats[chapter]["M"]),
            ("Female", stats[chapter]["F"]),
        ]
        gender_items = [(name, count) for name, count in gender_items if count > 0]

        ws[f"A{base}"] = f"Chapter {chapter} Gender"
        ws[f"B{base}"] = "Count"
        for offset, (name, count) in enumerate(gender_items, start=1):
            ws[f"A{base + offset}"] = name
            ws[f"B{base + offset}"] = count

        region_items = [
            (region, stats[chapter]["regions"][region])
            for region in REGIONS
            if stats[chapter]["regions"][region] > 0
        ]

        ws[f"D{base}"] = f"Chapter {chapter} Region"
        ws[f"E{base}"] = "Count"
        for offset, (name, count) in enumerate(region_items, start=1):
            ws[f"D{base + offset}"] = name
            ws[f"E{base + offset}"] = count

        chart_ranges[chapter] = {
            "gender_start": base,
            "gender_count": len(gender_items),
            "region_start": base,
            "region_count": len(region_items),
        }

    ws.sheet_state = "hidden"
    return ws, stats, chart_ranges


# ============================================================
# CHART FORMATTING
# ============================================================

def make_rich_text(text, size_pt=11, bold=False):
    """
    Build chart rich text with a requested font size.
    DrawingML font size uses hundredths of a point.
    """
    props = CharacterProperties(sz=size_pt * 100, b=bold)
    run = RegularTextRun(rPr=props, t=text)

    paragraph = Paragraph(
        pPr=ParagraphProperties(defRPr=CharacterProperties(sz=size_pt * 100)),
        r=[run]
    )

    return RichText(
        bodyPr=RichTextProperties(),
        p=[paragraph]
    )


def set_chart_text_size(chart, size_pt=11):
    """Set the chart-wide default text size."""
    chart.txPr = RichText(
        bodyPr=RichTextProperties(),
        p=[
            Paragraph(
                pPr=ParagraphProperties(
                    defRPr=CharacterProperties(sz=size_pt * 100)
                ),
                endParaRPr=CharacterProperties(sz=size_pt * 100)
            )
        ]
    )


def _rich_text_props(size_pt=11):
    return RichText(
        bodyPr=RichTextProperties(),
        p=[
            Paragraph(
                pPr=ParagraphProperties(
                    defRPr=CharacterProperties(sz=size_pt * 100)
                ),
                endParaRPr=CharacterProperties(sz=size_pt * 100)
            )
        ]
    )


def configure_pie_chart(chart, title):
    chart.title = title
    chart.height = 7
    # Wider canvas gives the right-side legend its own space and visually
    # shifts the pie toward the left.
    chart.width = 20

    chart.legend.position = "b"
    chart.legend.overlay = False
    chart.legend.txPr = _rich_text_props(11)

    # Custom labels are injected into the chart XML after the workbook is saved
    # so that the exact format is: 5 (25%)
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = False
    chart.dataLabels.showPercent = False
    chart.dataLabels.showCatName = False
    chart.dataLabels.showSerName = False
    chart.dataLabels.showLegendKey = False
    chart.dataLabels.showLeaderLines = True  # for labels Excel moves outside
    chart.dataLabels.txPr = _rich_text_props(11)

    set_chart_text_size(chart, 11)


def recreate_dashboard(workbook, chart_data_ws, stats, chart_ranges, record_count):
    if DASHBOARD_SHEET in workbook.sheetnames:
        dashboard_index = workbook.sheetnames.index(DASHBOARD_SHEET)
        del workbook[DASHBOARD_SHEET]
        dashboard = workbook.create_sheet(DASHBOARD_SHEET, dashboard_index)
    else:
        dashboard = workbook.create_sheet(DASHBOARD_SHEET)

    dashboard["A1"] = "External CA Dashboard"
    dashboard["A1"].font = Font(bold=True, size=16)

    dashboard["A2"] = "Last updated:"
    dashboard["B2"] = datetime.now()
    dashboard["B2"].number_format = "yyyy-mm-dd hh:mm"

    dashboard["A3"] = "Total records in Data:"
    dashboard["B3"] = record_count

    dashboard["A4"] = "Dashboard filter:"
    dashboard["B4"] = "Authors that are not a AR7 WGI CLA or LA."

    dashboard["A5"] = "Dashboard counting:"
    dashboard["B5"] = "Each author is counted once per chapter; latest submission in column A is used."

    dashboard.column_dimensions["A"].width = 24
    dashboard.column_dimensions["B"].width = 78

    start_row = 8
    rows_per_chapter = 17

    for chapter in range(1, 11):
        block_row = start_row + (chapter - 1) * rows_per_chapter
        dashboard[f"A{block_row}"] = f"Chapter {chapter}"
        dashboard[f"A{block_row}"].font = Font(bold=True, size=14)

        ranges = chart_ranges[chapter]

        # Gender pie: only non-zero categories are included.
        if ranges["gender_count"] > 0:
            gender_chart = PieChart()
            configure_pie_chart(gender_chart, f"Chapter {chapter} – Gender")

            base = ranges["gender_start"]
            gender_data = Reference(
                chart_data_ws, min_col=2, min_row=base,
                max_row=base + ranges["gender_count"]
            )
            gender_labels = Reference(
                chart_data_ws, min_col=1, min_row=base + 1,
                max_row=base + ranges["gender_count"]
            )
            gender_chart.add_data(gender_data, titles_from_data=True)
            gender_chart.set_categories(gender_labels)
            dashboard.add_chart(gender_chart, f"A{block_row + 1}")

        # Region pie: only non-zero categories are included.
        if ranges["region_count"] > 0:
            region_chart = PieChart()
            configure_pie_chart(region_chart, f"Chapter {chapter} – Region")

            base = ranges["region_start"]
            region_data = Reference(
                chart_data_ws, min_col=5, min_row=base,
                max_row=base + ranges["region_count"]
            )
            region_labels = Reference(
                chart_data_ws, min_col=4, min_row=base + 1,
                max_row=base + ranges["region_count"]
            )
            region_chart.add_data(region_data, titles_from_data=True)
            region_chart.set_categories(region_labels)
            dashboard.add_chart(region_chart, f"K{block_row + 1}")

    return dashboard


# ============================================================
# EXACT PIE LABELS: "5 (25%)"
# ============================================================

CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

ET.register_namespace("c", CHART_NS)
ET.register_namespace("a", DRAWING_NS)


def _format_percentage(count, total):
    if total <= 0:
        return "0%"

    percentage = count * 100.0 / total

    # Whole percentages are cleaner as 25%; otherwise retain one decimal,
    # e.g. 33.3%.
    if abs(percentage - round(percentage)) < 1e-9:
        return f"{int(round(percentage))}%"

    return f"{percentage:.1f}%"


def _labels_for_chart(stats, chapter, chart_type):
    if chart_type == "Gender":
        counts = [
            stats[chapter]["M"],
            stats[chapter]["F"],
        ]
    else:
        counts = [
            stats[chapter]["regions"][region]
            for region in REGIONS
        ]

    # This must mirror ChartData: zero-count categories are omitted.
    counts = [count for count in counts if count > 0]
    total = sum(counts)

    return [
        f"{count} ({_format_percentage(count, total)})"
        for count in counts
    ]


def _chart_title(root):
    texts = root.findall(
        ".//c:title//a:t",
        {"c": CHART_NS, "a": DRAWING_NS}
    )
    return "".join(node.text or "" for node in texts)


def _set_title_upper_left(root):
    """Move the chart title to the upper-left of the chart area."""
    ns = {"c": CHART_NS}
    title = root.find(".//c:title", ns)
    if title is None:
        return

    # Remove an existing title layout, if any.
    existing = title.find("c:layout", ns)
    if existing is not None:
        title.remove(existing)

    layout = ET.Element(f"{{{CHART_NS}}}layout")
    manual = ET.SubElement(layout, f"{{{CHART_NS}}}manualLayout")

    x_mode = ET.SubElement(manual, f"{{{CHART_NS}}}xMode")
    x_mode.set("val", "edge")
    y_mode = ET.SubElement(manual, f"{{{CHART_NS}}}yMode")
    y_mode.set("val", "edge")

    x = ET.SubElement(manual, f"{{{CHART_NS}}}x")
    x.set("val", "0.02")
    y = ET.SubElement(manual, f"{{{CHART_NS}}}y")
    y.set("val", "0.02")

    # Layout belongs after title text but before overlay/formatting.
    insert_at = 1 if len(title) >= 1 else 0
    title.insert(insert_at, layout)


def _set_best_fit_data_labels(d_lbls):
    """
    Ask Excel to keep labels inside when they fit and move them outside
    when necessary. Outside labels may use leader lines.
    """
    # Remove any existing global label-position setting.
    for child in list(d_lbls):
        if child.tag == f"{{{CHART_NS}}}dLblPos":
            d_lbls.remove(child)

    pos = ET.Element(f"{{{CHART_NS}}}dLblPos")
    pos.set("val", "bestFit")

    # Place after point-specific dLbl elements.
    insert_at = 0
    for i, child in enumerate(list(d_lbls)):
        if child.tag == f"{{{CHART_NS}}}dLbl":
            insert_at = i + 1
    d_lbls.insert(insert_at, pos)


def _make_custom_data_label(index, label_text):
    """
    Create one OOXML data label with exact custom text and 11 pt font.
    """
    d_lbl = ET.Element(f"{{{CHART_NS}}}dLbl")

    idx = ET.SubElement(d_lbl, f"{{{CHART_NS}}}idx")
    idx.set("val", str(index))

    # Explicitly suppress the colored legend-key square for this data label.
    show_legend_key = ET.SubElement(d_lbl, f"{{{CHART_NS}}}showLegendKey")
    show_legend_key.set("val", "0")

    tx = ET.SubElement(d_lbl, f"{{{CHART_NS}}}tx")
    rich = ET.SubElement(tx, f"{{{CHART_NS}}}rich")

    body_pr = ET.SubElement(rich, f"{{{DRAWING_NS}}}bodyPr")
    body_pr.set("rot", "0")

    ET.SubElement(rich, f"{{{DRAWING_NS}}}lstStyle")

    paragraph = ET.SubElement(rich, f"{{{DRAWING_NS}}}p")
    run = ET.SubElement(paragraph, f"{{{DRAWING_NS}}}r")

    r_pr = ET.SubElement(run, f"{{{DRAWING_NS}}}rPr")
    r_pr.set("lang", "en-US")
    r_pr.set("sz", "1100")  # 11 pt

    text = ET.SubElement(run, f"{{{DRAWING_NS}}}t")
    text.text = label_text

    end_pr = ET.SubElement(paragraph, f"{{{DRAWING_NS}}}endParaRPr")
    end_pr.set("lang", "en-US")
    end_pr.set("sz", "1100")

    return d_lbl


def patch_exact_pie_labels(xlsx_path, stats):
    """
    openpyxl/Excel's automatic pie labels cannot reliably format value and
    percentage as exactly '5 (25%)'. This function patches the saved XLSX's
    chart XML so every non-zero pie slice gets that exact custom label.

    It finds charts by their titles:
        Chapter N – Gender
        Chapter N – Region
    """
    temp_fd, temp_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(temp_fd)

    try:
        with zipfile.ZipFile(xlsx_path, "r") as zin, \
             zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as zout:

            for item in zin.infolist():
                data = zin.read(item.filename)

                if (
                    item.filename.startswith("xl/charts/chart")
                    and item.filename.endswith(".xml")
                ):
                    root = ET.fromstring(data)
                    title = _chart_title(root)

                    match = re.search(
                        r"Chapter\s+(\d+)\s*[–-]\s*(Gender|Region)",
                        title,
                        flags=re.IGNORECASE
                    )

                    if match:
                        chapter = int(match.group(1))
                        chart_type = match.group(2).title()
                        labels = _labels_for_chart(stats, chapter, chart_type)

                        # Keep the title away from the pie/data labels.
                        _set_title_upper_left(root)

                        d_lbls = root.find(
                            ".//c:pieChart/c:dLbls",
                            {"c": CHART_NS}
                        )

                        if d_lbls is not None:
                            # Remove any existing point-specific labels.
                            for child in list(d_lbls):
                                if child.tag == f"{{{CHART_NS}}}dLbl":
                                    d_lbls.remove(child)

                            # Insert exact custom labels: e.g. 5 (25%).
                            insert_at = 0
                            for index, label_text in enumerate(labels):
                                d_lbls.insert(
                                    insert_at,
                                    _make_custom_data_label(index, label_text)
                                )
                                insert_at += 1

                            # Excel chooses inside/outside placement. If a label
                            # does not fit, it can be moved outside with a leader
                            # line. No legend-key/colored square is added.
                            _set_best_fit_data_labels(d_lbls)

                            data = ET.tostring(
                                root,
                                encoding="utf-8",
                                xml_declaration=True
                            )

                zout.writestr(item, data)

        os.replace(temp_path, xlsx_path)

    except Exception:
        if os.path.exists(temp_path):
            os.remove(temp_path)
        raise

from copy import copy
def set_workbook_font(workbook, font_name="Arial Narrow"):
    for ws in workbook.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    new_font = copy(cell.font)
                    new_font.name = font_name
                    cell.font = new_font
# ============================================================
# MAIN
# ============================================================

def main():
    source_file, template_file, output_file = select_files()

    if not source_file:
        return

    try:
        source_wb = load_workbook(source_file, data_only=False)
        destination_wb = load_workbook(template_file)

        source_ws = find_source_sheet(source_wb)
        destination_ws = get_destination_sheet(destination_wb)

        clear_old_destination_values(destination_ws)

        record_count, source_metadata = transfer_data(
            source_ws,
            destination_ws
        )

        # Reconcile duplicate people BEFORE dashboard calculations.
        reconcile_people(
            destination_ws,
            source_metadata,
            record_count
        )

        chart_data_ws, stats, chart_ranges = recreate_chart_data_sheet(
            destination_wb,
            destination_ws,
            record_count
        )

        recreate_dashboard(
            destination_wb,
            chart_data_ws,
            stats,
            chart_ranges,
            record_count
        )

        set_workbook_font(destination_wb, "Arial Narrow")
        destination_wb.save(output_file)

        # Patch chart XML so pie labels are exactly like: 5 (25%)
        patch_exact_pie_labels(output_file, stats)

        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)

        messagebox.showinfo(
            "Update complete",
            (
                "The workbook was created successfully.\n\n"
                f"Records copied: {record_count}\n"
                f"Dashboard: {DASHBOARD_SHEET}\n\n"
                f"Output file:\n{output_file}"
            )
        )

        root.destroy()

    except Exception as error:
        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)

        messagebox.showerror(
            "Update failed",
            (
                "The workbook could not be created.\n\n"
                f"{type(error).__name__}: {error}"
            )
        )

        root.destroy()
        raise


if __name__ == "__main__":
    main()
