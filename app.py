import re
from io import BytesIO
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Construction Progress Organizer", page_icon="🏗️", layout="wide")

COLUMNS = [
    "Date", "Tower / Location", "Today's Activities", "Tomorrow's Planned Activities",
    "Manpower", "Machinery", "Materials", "Weather", "Working Hours", "Other Notes",
    "Original Message"
]

DATE_PATTERNS = [
    (re.compile(r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b"), "%d/%m/%Y"),
    (re.compile(r"\b(\d{4}-\d{1,2}-\d{1,2})\b"), "%Y-%m-%d"),
]

TOWER_RE = re.compile(
    r"\b((?:MO|KEM|KWT|TOWER)[\s-]*[A-Z0-9]+(?:[\s/-]*[A-Z0-9]+)*)\b",
    re.IGNORECASE
)

CATEGORY_KEYWORDS = {
    "Manpower": ["manpower", "worker", "workers", "labour", "labor", "personnel", "staff"],
    "Machinery": ["machinery", "machine", "excavator", "lorry", "truck", "crane", "roller", "backhoe"],
    "Materials": ["material", "materials", "cement", "sand", "aggregate", "stone", "gabion", "geotextile", "steel", "concrete"],
    "Weather": ["weather", "rain", "raining", "sunny", "cloudy"],
    "Working Hours": ["working hours", "work hour", "started at", "finished at", "stop work", "resumed work"],
}

TOMORROW_MARKERS = [
    "tomorrow", "plan for tomorrow", "planned for tomorrow", "next day",
    "proposed activity", "planned activity", "will continue", "will proceed"
]
TODAY_MARKERS = ["today's activity", "today activity", "work done", "completed today", "actual activity"]


def normalize_date(value):
    value = value.strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return value


def extract_date(text):
    for pattern, _ in DATE_PATTERNS:
        match = pattern.search(text)
        if match:
            return normalize_date(match.group(1))
    return ""


def extract_tower(text):
    match = TOWER_RE.search(text)
    if not match:
        return ""
    value = re.sub(r"\s+", " ", match.group(1)).strip(" -/")
    return value.upper()


def classify_line(line):
    lower = line.lower()
    if any(marker in lower for marker in TOMORROW_MARKERS):
        return "tomorrow"
    if any(marker in lower for marker in TODAY_MARKERS):
        return "today"
    for field, words in CATEGORY_KEYWORDS.items():
        if any(word in lower for word in words):
            return field
    return "today"


def clean_line(line):
    return re.sub(r"^\s*(?:[-*•]|\d+[.)])\s*", "", line).strip()


def parse_messages(raw_text):
    """Heuristic parser: keeps date/tower context and groups extracted lines."""
    records = []
    current_date = ""
    current_tower = ""
    current = None

    def new_record(date, tower):
        return {
            "Date": date,
            "Tower / Location": tower,
            "Today's Activities": [],
            "Tomorrow's Planned Activities": [],
            "Manpower": [],
            "Machinery": [],
            "Materials": [],
            "Weather": [],
            "Working Hours": [],
            "Other Notes": [],
            "Original Message": ""
        }

    def flush():
        nonlocal current
        if current is None:
            return
        for key in COLUMNS[2:-1]:
            current[key] = "; ".join(current[key])
        records.append(current)
        current = None

    for original_line in raw_text.splitlines():
        line = clean_line(original_line)
        if not line:
            continue

        found_date = extract_date(line)
        found_tower = extract_tower(line)

        # A new date/tower header starts a new grouped record.
        is_header = bool(found_date or found_tower)
        if is_header:
            if current is not None and (found_date or found_tower):
                flush()
            if found_date:
                current_date = found_date
            if found_tower:
                current_tower = found_tower
            # Remove date and tower labels from the remaining text.
            content = line
            if found_date:
                content = content.replace(re.search(r"\b\d{1,4}[/-]\d{1,2}[/-]\d{2,4}\b|\b\d{4}-\d{1,2}-\d{1,2}\b", line).group(0), "")
            if found_tower:
                content = re.sub(re.escape(found_tower), "", content, flags=re.IGNORECASE)
            content = content.strip(" :-|,")
            if not content:
                continue
            line = content

        if current is None:
            current = new_record(current_date, current_tower)

        if not current["Original Message"]:
            current["Original Message"] = original_line.strip()
        else:
            current["Original Message"] += "\n" + original_line.strip()

        kind = classify_line(line)
        if kind == "tomorrow":
            current["Tomorrow's Planned Activities"].append(line)
        elif kind == "today":
            current["Today's Activities"].append(line)
        elif kind in current:
            current[kind].append(line)
        else:
            current["Other Notes"].append(line)

    flush()
    return pd.DataFrame(records, columns=COLUMNS)


def make_excel(progress_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        progress_df.to_excel(writer, index=False, sheet_name="Weekly Progress")
        instructions = pd.DataFrame({
            "Field": COLUMNS,
            "Description": [
                "Date of the reported work (YYYY-MM-DD where detected)",
                "Tower or work location detected from the message",
                "Work reported as completed / carried out",
                "Activities stated as planned for the next day",
                "Manpower or worker information",
                "Machinery or equipment information",
                "Materials mentioned",
                "Weather or rainfall information",
                "Working time information",
                "Other information requiring review",
                "Original text retained for checking"
            ]
        })
        instructions.to_excel(writer, index=False, sheet_name="Column Guide")
    output.seek(0)
    return output.getvalue()


st.title("🏗️ Construction Progress Organizer")
st.write(
    "Paste contractor progress messages to organize information by date and tower, "
    "review the extracted results, and export them to Excel."
)
st.info(
    "This is a rule-based prototype. It does not understand every WhatsApp writing style. "
    "Always check and correct the preview before exporting."
)

with st.expander("How to use"):
    st.markdown(
        """
        1. Paste the contractor's message in the text box.
        2. Select **Organize message**.
        3. Check the extracted date, tower and information in the editable table.
        4. Correct any mistakes manually.
        5. Optionally upload an existing workbook to add the new rows.
        6. Download the updated Excel file.

        **Tip:** The parser works best when dates and tower names are written clearly, for example:
        `17/11/2026 - Tower MO72` followed by bullet points.
        """
    )

sample = """17/11/2026 - Tower MO72
Today's activities:
- Continue construction of gabion wall
- Install geotextile at slope area
Manpower: 6 workers
Machinery: 1 excavator
Weather: Cloudy
Plan for tomorrow:
- Continue gabion wall construction

17/11/2026 - Tower MO89
- Carry out drainage works
- Deliver stone materials
Working hours: 8.00 am - 5.00 pm
"""
with st.expander("View sample message"):
    st.code(sample, language="text")

raw_text = st.text_area(
    "Paste contractor WhatsApp message here",
    height=250,
    placeholder="Paste the full progress message..."
)

if st.button("Organize message", type="primary", disabled=not raw_text.strip()):
    parsed = parse_messages(raw_text)
    if parsed.empty:
        st.warning("No information was detected. Try pasting a message with a date, tower name or activity details.")
    else:
        st.session_state["parsed_progress"] = parsed
        st.success(f"Organized {len(parsed)} record(s). Please review the table below.")

if "parsed_progress" in st.session_state:
    st.subheader("1. Review and edit extracted information")
    edited = st.data_editor(
        st.session_state["parsed_progress"],
        num_rows="dynamic",
        use_container_width=True,
        key="progress_editor"
    )
    st.session_state["edited_progress"] = edited

    st.subheader("2. Add to an existing Excel workbook (optional)")
    uploaded = st.file_uploader("Upload an existing .xlsx progress workbook", type=["xlsx"])
    existing_df = None
    if uploaded is not None:
        try:
            existing_df = pd.read_excel(uploaded, sheet_name="Weekly Progress")
            for column in COLUMNS:
                if column not in existing_df.columns:
                    existing_df[column] = ""
            existing_df = existing_df[COLUMNS]
            st.caption(f"Loaded {len(existing_df)} existing row(s) from the 'Weekly Progress' sheet.")
        except Exception as exc:
            st.error(
                "The workbook could not be read. It must be an .xlsx file with a sheet named "
                f"'Weekly Progress'. Details: {exc}"
            )

    if st.button("Prepare Excel download", type="primary"):
        final_df = edited.copy()
        if existing_df is not None:
            final_df = pd.concat([existing_df, final_df], ignore_index=True)
        final_df = final_df.fillna("")
        st.session_state["excel_bytes"] = make_excel(final_df)
        st.session_state["export_rows"] = len(final_df)
        st.success(f"Excel file prepared with {len(final_df)} total row(s).")

    if "excel_bytes" in st.session_state:
        st.download_button(
            "⬇️ Download construction_progress.xlsx",
            data=st.session_state["excel_bytes"],
            file_name="construction_progress.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.caption("The original message is included in the workbook so you can verify the extracted information.")

st.divider()
st.caption("Prototype for academic demonstration. Review all extracted data before using it in an official report.")
