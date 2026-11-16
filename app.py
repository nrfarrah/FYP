import re
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Construction Progress Organizer", page_icon="🏗️", layout="wide")

# These columns follow the user's sample progress-report layout.
COLUMNS = [
    "Date",
    "Tower",
    "Working hour",
    "weather",
    "rainfall (mm)",
    "Main power",
    "machinery",
    "today's activity",
    "plan/ tomorrow activity",
    "Soil nail compacted",
    "MACHINERIES/MATERIALS RECEIVED:",
    "Checklists",
]

DATE_PATTERNS = [
    re.compile(r"\b\d{1,2}[-/](?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\b", re.I),
    re.compile(r"\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b"),
    re.compile(r"\b\d{4}-\d{1,2}-\d{1,2}\b"),
]
TOWER_RE = re.compile(r"\b(?:MO\s*[-/]?\s*\d+|KEM\s*[-/]?\s*[A-Z0-9]+|KWT\s*[-/]?\s*[A-Z0-9]+|[A-Z]{2,6}[- ]\d{2,4})\b", re.I)

# Labels are intentionally editable so they can be adjusted to the contractor's wording.
FIELD_LABELS = {
    "Working hour": ["working hour", "working hours", "work hour", "work hours", "working time"],
    "weather": ["weather"],
    "rainfall (mm)": ["rainfall", "rain fall", "precipitation"],
    "Main power": ["main power", "manpower", "mainpower"],
    "machinery": ["machinery", "machine", "equipment"],
    "today's activity": ["today's activity", "todays activity", "today activity", "today's activities",
                         "today activities", "actual activity", "activities carried out", "work done"],
    "plan/ tomorrow activity": ["plan/ tomorrow activity", "plan tomorrow", "tomorrow activity",
                                "tomorrow's activity", "tomorrow activities", "planned activity",
                                "planned activities", "plan for tomorrow", "next day activity"],
    "Soil nail compacted": ["soil nail compacted", "soil nail", "soil nailing", "compacted"],
    "MACHINERIES/MATERIALS RECEIVED:": ["machineries/materials received", "machinery/materials received",
                                       "materials received", "machinery received", "materials delivered"],
    "Checklists": ["checklists", "checklist", "mp status", "mp"],
}

def find_date(text):
    for pattern in DATE_PATTERNS:
        m = pattern.search(text)
        if m:
            return m.group(0)
    return ""

def find_tower(text):
    m = TOWER_RE.search(text)
    return re.sub(r"\s+", "", m.group(0)).upper() if m else ""

def norm_label(text):
    return re.sub(r"[\s:_\-]+", " ", text.strip().lower()).strip()

# Build label matching from longest to shortest to avoid matching "machine"
# before "machineries/materials received".
LABEL_LOOKUP = []
for field, labels in FIELD_LABELS.items():
    for label in labels:
        LABEL_LOOKUP.append((norm_label(label), field))
LABEL_LOOKUP.sort(key=lambda x: len(x[0]), reverse=True)

def match_label(line):
    cleaned = line.strip()
    low = cleaned.lower()
    for label, field in LABEL_LOOKUP:
        # Accept "Label: value", "Label - value", and "Label value".
        pattern = r"^\s*" + re.escape(label).replace(r"\ ", r"\s+") + r"\s*(?::|=|-)?\s*(.*)$"
        m = re.match(pattern, low, re.I)
        if m:
            value = cleaned[len(cleaned) - len(m.group(1)):].strip() if m.group(1) else ""
            return field, value
    return None, ""

def split_inline_label(line):
    """Detect a field label appearing after other text on the same line."""
    low = line.lower()
    candidates = []
    for label, field in LABEL_LOOKUP:
        idx = low.find(label)
        if idx > 0:
            candidates.append((idx, label, field))
    if not candidates:
        return line, None, ""
    idx, label, field = sorted(candidates, key=lambda x: x[0])[0]
    before = line[:idx].strip(" ;,|-")
    after = line[idx + len(label):].strip(" :-=;,")
    return before, field, after

def empty_record(date="", tower=""):
    return {c: "" for c in COLUMNS} | {"Date": date, "Tower": tower}

def parse_message(message, source_number):
    """Extract one or more date/tower records from a single pasted message."""
    records = []
    current_date = find_date(message)
    current_tower = find_tower(message)
    current = empty_record(current_date, current_tower)
    current["_source"] = f"Message {source_number}"
    current_field = "today's activity"

    def add_value(field, value):
        value = value.strip(" \t:-|,;")
        if not value:
            return
        existing = current.get(field, "")
        current[field] = (existing + "\n" + value).strip() if existing else value

    def has_payload(row):
        return any(str(row.get(c, "")).strip() for c in COLUMNS[2:])

    def finish():
        nonlocal current
        if has_payload(current) or current["Date"] or current["Tower"]:
            records.append({c: current.get(c, "") for c in COLUMNS} | {"Source Message": current.get("_source", "")})
        current = None

    for raw in message.replace("•", "\n").splitlines():
        line = re.sub(r"^\s*(?:[-*]|\d+[.)])\s*", "", raw).strip()
        if not line:
            continue

        line_date = find_date(line)
        line_tower = find_tower(line)

        # A new date/tower header starts a new spreadsheet row when the current
        # row already contains information.
        if line_date or line_tower:
            if has_payload(current):
                finish()
                current = empty_record(line_date or current_date, line_tower or current_tower)
                current["_source"] = f"Message {source_number}"
                current_field = "today's activity"
            else:
                if line_date:
                    current["Date"] = line_date
                if line_tower:
                    current["Tower"] = line_tower

            line = DATE_PATTERNS[0].sub("", line)
            line = DATE_PATTERNS[1].sub("", line)
            line = DATE_PATTERNS[2].sub("", line)
            line = TOWER_RE.sub("", line).strip(" :-|,")
            if not line:
                continue

        field, value = match_label(line)
        if field:
            current_field = field
            if value:
                add_value(field, value)
            continue

        # Also handle messages where a label follows a sentence on the same line.
        before, inline_field, after = split_inline_label(line)
        if inline_field:
            if before:
                add_value(current_field, before)
            current_field = inline_field
            if after:
                add_value(inline_field, after)
            continue

        # Recognise common unlabelled information without guessing too much.
        lower = line.lower()
        if re.search(r"\b\d+(?:\.\d+)?\s*mm\b", lower) and not current["rainfall (mm)"]:
            add_value("rainfall (mm)", line)
        elif any(x in lower for x in ["sunny", "cloudy", "rainy", "heavy rain", "light rain"]):
            add_value("weather", line)
        elif re.search(r"\b\d{3,4}\s*(?:hrs|hours)\b", lower) or ("am" in lower and "pm" in lower and any(ch.isdigit() for ch in lower)):
            add_value("Working hour", line)
        elif any(x in lower for x in ["excavator", "drilling machine", "concrete mixer", "grout mixer",
                                      "lorry", "water pump", "air compressor", "compactor", "crane", "backhoe"]):
            add_value("machinery", line)
        elif any(x in lower for x in ["site engineer", "site supervisor", "operator", "general worker",
                                      "technician", "pax", "manpower"]):
            add_value("Main power", line)
        elif any(x in lower for x in ["complete", "completed", "mp (", "checklist", "pending"]):
            add_value("Checklists", line)
        else:
            # Unlabelled progress statements default to today's activity.
            add_value(current_field if current_field in ["today's activity", "plan/ tomorrow activity"] else "today's activity", line)

    finish()
    # Remove rows that contain no extracted content at all.
    return [r for r in records if any(str(r.get(c, "")).strip() for c in COLUMNS)]

def parse_all(messages):
    rows = []
    for i, msg in enumerate(messages, start=1):
        if msg.strip():
            rows.extend(parse_message(msg, i))
    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=COLUMNS + ["Source Message"])
    return df[COLUMNS + ["Source Message"]]

def make_formatted_excel(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Weekly Progress"
    ws.append(COLUMNS)

    # Add one row per extracted date/tower record, in the same column order as
    # the example supplied by the user.
    for _, row in df.iterrows():
        ws.append([row.get(col, "") for col in COLUMNS])

    header_fill = PatternFill("solid", fgColor="8DB4E2")
    header_font = Font(bold=True, color="000000", size=10)
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    widths = [13, 12, 19, 14, 15, 24, 23, 34, 34, 18, 27, 20]
    for col_idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[1].height = 34

    for row_idx in range(2, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 110
        for cell in ws[row_idx]:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # Include source message mapping on a separate sheet for traceability.
    source_ws = wb.create_sheet("Source Mapping")
    source_ws.append(["Excel row", "Source message"])
    for i, source in enumerate(df.get("Source Message", []), start=2):
        source_ws.append([i, source])
    for cell in source_ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
    source_ws.column_dimensions["A"].width = 14
    source_ws.column_dimensions["B"].width = 22

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


st.title("🏗️ Construction Progress Organizer")
st.write(
    "Paste each contractor update into a separate message box. The tool extracts the details "
    "and prepares an Excel report using the same column layout as your sample."
)
st.info(
    "The Excel columns follow your example: Date, Tower, Working hour, weather, rainfall, "
    "Main power, machinery, today's activity, plan/tomorrow activity, Soil nail compacted, "
    "Machineries/Materials received, and Checklists."
)

with st.expander("How to use", expanded=True):
    st.markdown(
        """
        1. Paste one contractor message into Message 1.
        2. Click **+ Add another message** for each additional update.
        3. Click **Organize messages**.
        4. Review and correct the table. Automatic extraction can make mistakes.
        5. Download the formatted Excel report.
        """
    )

if "messages" not in st.session_state:
    st.session_state.messages = [""]

def add_box():
    st.session_state.messages.append("")

def remove_box(i):
    if len(st.session_state.messages) > 1:
        st.session_state.messages.pop(i)

st.subheader("1. Paste messages")
for i in range(len(st.session_state.messages)):
    st.markdown(f"**Message {i + 1}**")
    st.session_state.messages[i] = st.text_area(
        f"Paste message {i + 1}",
        value=st.session_state.messages[i],
        key=f"msg_{i}",
        height=150,
        placeholder="Paste one complete WhatsApp progress update here..."
    )
    if len(st.session_state.messages) > 1:
        st.button(f"Remove Message {i + 1}", key=f"remove_{i}", on_click=remove_box, args=(i,))

st.button("+ Add another message", on_click=add_box)

if st.button("Organize messages", type="primary"):
    result = parse_all(st.session_state.messages)
    st.session_state.result_df = result
    if result.empty:
        st.warning("No usable details were detected. Please check the messages or fill in the missing details manually.")
    else:
        st.success(f"Created {len(result)} row(s). Review the extracted information below.")

if "result_df" in st.session_state and not st.session_state.result_df.empty:
    st.subheader("2. Review and edit the Excel table")
    st.caption("The table is editable. Correct any information that was placed in the wrong column.")
    editable = st.data_editor(
        st.session_state.result_df,
        use_container_width=True,
        num_rows="dynamic",
        key="edit_report",
        column_config={
            "today's activity": st.column_config.TextColumn(width="large"),
            "plan/ tomorrow activity": st.column_config.TextColumn(width="large"),
            "Main power": st.column_config.TextColumn(width="large"),
            "machinery": st.column_config.TextColumn(width="large"),
            "MACHINERIES/MATERIALS RECEIVED:": st.column_config.TextColumn(width="large"),
        }
    )

    st.subheader("3. Download formatted Excel")
    st.write("The downloaded workbook uses the same column arrangement and a blue header similar to your sample.")
    if st.button("Prepare Excel report", type="primary"):
        st.session_state.xlsx_bytes = make_formatted_excel(editable.fillna(""))
        st.success("Your Excel report is ready.")
    if "xlsx_bytes" in st.session_state:
        st.download_button(
            "⬇️ Download weekly_progress_report.xlsx",
            data=st.session_state.xlsx_bytes,
            file_name="weekly_progress_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.caption(
    "Prototype limitation: this version extracts information using labels and keywords. "
    "Contractor messages may use different wording, so always verify the generated report."
)
