import re
from io import BytesIO
from math import ceil
from datetime import datetime, timedelta

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
    "Remarks",
]

REF_YEAR = datetime.now().year  # used when a message has a date without a year

# Tower code prefixes (editable in the sidebar). Tower = prefix + number, e.g. MO72, OF30.
DEFAULT_TOWER_PREFIXES = ""  # blank = auto-detect any tower code

# Fields that hold a numbered list inside one Excel cell, and the numbering style.
LIST_FIELDS = {
    "Main power": ".",                          # 1. Site Engineer - 1 pax
    "machinery": ".",                           # 1. Excavator - 2
    "today's activity": ")",                    # 1) Slope clearing ...
    "plan/ tomorrow activity": ")",
    "MACHINERIES/MATERIALS RECEIVED:": ")",
    "Soil nail compacted": ")",
    "Remarks": ")",
}

# Labels are intentionally editable so they can be adjusted to the contractor's wording.
FIELD_LABELS = {
    "Working hour": ["working hour", "working hours", "work hour", "work hours", "working time"],
    "weather": ["weather"],
    "rainfall (mm)": ["rainfall", "rain fall", "precipitation"],
    "Main power": ["main power", "manpower", "mainpower"],
    "machinery": ["machinery", "machineries", "machine", "equipment"],
    "today's activity": ["today's activity", "todays activity", "today activity", "today's activities",
                         "today activities", "actual activity", "activities carried out", "work done"],
    "plan/ tomorrow activity": ["plan/ tomorrow activity", "plan tomorrow", "tomorrow activity",
                                "tomorrow's activity", "tomorrow activities", "planned activity",
                                "planned activities", "plan for tomorrow", "next day activity",
                                "plan activity", "tomorrow plan", "plan tomorrow activity",
                                "plan tomorrow activities", "plan for tomorrow activity",
                                "tomorrow plan activity", "plan"],
    "Soil nail compacted": ["soil nail compacted", "soil nails compacted", "soil nail completed",
                            "soil nails completed", "soil nail complete", "soil nailing completed",
                            "soil nail progress", "soil nail"],
    "MACHINERIES/MATERIALS RECEIVED:": ["machineries/materials received", "machinery/materials received",
                                       "machineries materials received", "machinery materials received",
                                       "material received", "materials received", "machinery received",
                                       "machineries received", "materials delivered", "material delivered"],
    "Checklists": ["checklists", "checklist", "mp status"],
    "Remarks": ["remarks", "remark", "notes", "note"],
}

# Lines containing these are report titles, not data. Add your own project title words here.
SKIP_PATTERNS = [
    re.compile(r"^[\W_]+$"),
    re.compile(r"daily\s+progress\s+report", re.I),
    re.compile(r"engineering\s+remedial\s+works", re.I),
]

# ---------------------------------------------------------------- dates
MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], start=1)}

DATE_TEXT_RE = re.compile(
    r"\b(\d{1,2})(?:st|nd|rd|th)?[\s\-/]*(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?"
    r"(?:[\s\-/,]+(\d{4}))?(?!\d)", re.I)
DATE_NUM_RE = re.compile(r"\b(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4}|\d{2})\b")
DATE_ISO_RE = re.compile(r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b")
WEEKDAY_RE = re.compile(r"\(?\b(?:monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b\)?", re.I)
NOISE_LINE_RE = re.compile(r"^(?:date|day)\s*[:\-]?\s*$", re.I)
BULLET_RE = re.compile(r"^\s*(?:[-–●▪·]+|\d+\s*[.)])\s*")


def make_date(d, m, y=None):
    try:
        y = int(y) if y else REF_YEAR
        if y < 100:
            y += 2000
        return datetime(y, m, d)
    except (ValueError, TypeError):
        return None


def find_date(text):
    """Return (datetime, (start, end)) of the first date in text, or (None, None)."""
    m = DATE_TEXT_RE.search(text)
    if m:
        dt = make_date(int(m.group(1)), MONTHS[m.group(2).lower()], m.group(3))
        if dt:
            return dt, m.span()
    m = DATE_NUM_RE.search(text)  # dd/mm/yyyy
    if m:
        dt = make_date(int(m.group(1)), int(m.group(2)), m.group(3))
        if dt:
            return dt, m.span()
    m = DATE_ISO_RE.search(text)
    if m:
        dt = make_date(int(m.group(3)), int(m.group(2)), m.group(1))
        if dt:
            return dt, m.span()
    return None, None


def fmt_date(dt):
    return dt.strftime("%d-%b")  # 25-Oct, same as the sample


def parse_date_cell(s):
    return find_date(str(s))[0]


# ---------------------------------------------------------------- towers
# Letter codes that look like a tower code but are not (SN1-5 = soil nail, MP = manpower ...).
NOT_TOWER = {"SN", "MP", "EHV", "FDT", "RC", "KV", "MM", "CM", "NO", "KG", "MT", "LV", "HV", "TW", "TC", "ST"}


def build_tower_re(prefixes):
    """Blank prefixes = auto-detect any code like MO72, OF30, KT105 (2-3 capital letters + number).
    With prefixes (e.g. "MO, OF") only those codes are accepted."""
    items = sorted({p.strip().upper() for p in prefixes.split(",") if p.strip()}, key=len, reverse=True)
    if items:
        return re.compile(r"\b(?:" + "|".join(re.escape(p) for p in items) + r")\s?[-/]?\s?\d{1,4}[A-Z]?\b")
    return re.compile(r"\b[A-Z]{2,3}-?\d{1,4}[A-Z]?\b")


def find_tower(text, tower_re):
    for m in tower_re.finditer(text):
        code = re.sub(r"[\s\-/]+", "", m.group(0))
        if re.match(r"[A-Z]+", code).group(0) not in NOT_TOWER:
            return code
    return ""


# "TOWER KEM-KWT 127", "Tower: MO72, ..." -> the name written after the word TOWER.
TOWER_PHRASE_RE = re.compile(r"\btower\b\s*[:\-–]?\s*([A-Za-z0-9][A-Za-z0-9\-/]*(?:[ ]+\d{1,4}[A-Za-z]?)?)", re.I)


def extract_tower(line, tower_re):
    """Return (tower name, line without the tower). Tower is "" when none is found."""
    for m in TOWER_PHRASE_RE.finditer(line):
        name = m.group(1).strip()
        # needs a letter, plus a digit or ALL-CAPS ("TOWER 127" and "tower base" are not names)
        if re.search(r"[A-Za-z]", name) and (re.search(r"\d", name) or name.isupper()):
            return re.sub(r"\s+", " ", name).upper(), line[:m.start()] + " " + line[m.end():]
    code = find_tower(line, tower_re)
    if code:
        return code, tower_re.sub(" ", line)
    return "", line


# ---------------------------------------------------------------- labels
def _label_regex(label):
    # words may be separated by spaces, "_" or "/" (so "PLAN/TOMORROW ACTIVITY" works too)
    body = r"[\s_/]+".join(re.escape(w) for w in re.split(r"[\s/]+", label) if w)
    # label must be followed by ":" / "=" / " - " / end of line, so words inside
    # normal sentences (e.g. "compressive", "Air Compressor") are never mistaken for labels.
    return re.compile(r"^\s*" + body + r"\s*(?:\(mm\))?\s*(?:[:=]+|[-–]\s+|[.\s]*$)\s*(.*)$", re.I)


LABEL_LOOKUP = sorted(
    [(label, field) for field, labels in FIELD_LABELS.items() for label in labels],
    key=lambda x: len(x[0]), reverse=True,
)
LABEL_LOOKUP = [(_label_regex(label), field) for label, field in LABEL_LOOKUP]


def match_label(line):
    for rx, field in LABEL_LOOKUP:
        m = rx.match(line)
        if m:
            return field, m.group(1).strip()
    return None, ""


# ---------------------------------------------------------------- guessing for unlabelled lines
COUNT_END_RE = re.compile(r"\d+\s*(?:pax|nos|unit|units|ton)?\s*$", re.I)
MANPOWER_RE = re.compile(r"\b(pax|operator|driver|engineer|supervisor|general worker|technician|foreman|surveyor)\b", re.I)
MACHINERY_RE = re.compile(r"\b(excavator|drilling machine|concrete mixer|grout mixer|lorry|water pump|"
                          r"air compressor|compactor|crane|backhoe|generator|breaker)\b", re.I)
RAIN_RE = re.compile(r"\b\d+(?:\.\d+)?\s*mm\b", re.I)
WEATHER_RE = re.compile(r"\b(sunny|cloudy|rainy|rain|fine|overcast|drizzle|showers?)\b", re.I)
HOURS_RE = re.compile(r"\b\d{3,4}\s*(?:hrs|hours)\b|\d{1,2}(?:[:.]\d{2})?\s*(?:am|pm)\s*[-–to]+\s*\d{1,2}(?:[:.]\d{2})?\s*(?:am|pm)", re.I)
CHECKLIST_RE = re.compile(r"\bMP\b|checklist", re.I)


def guess_field(line):
    if RAIN_RE.search(line):
        return "rainfall (mm)"
    if HOURS_RE.search(line):
        return "Working hour"
    if len(line.split()) <= 4 and WEATHER_RE.search(line):
        return "weather"
    if COUNT_END_RE.search(line):
        if MANPOWER_RE.search(line):
            return "Main power"
        if MACHINERY_RE.search(line):
            return "machinery"
    if CHECKLIST_RE.search(line):
        return "Checklists"
    return None


# ---------------------------------------------------------------- row building
def build_row(items, date, tower, source):
    row = {c: "" for c in COLUMNS}
    row["Date"], row["Tower"] = date, tower
    for field, vals in items.items():
        if not vals:
            continue
        if field in LIST_FIELDS:
            sep = LIST_FIELDS[field]
            row[field] = "\n".join(f"{i}{sep} {v}" for i, v in enumerate(vals, start=1))
        elif field == "Checklists":
            row[field] = "\n".join(vals)
        else:
            row[field] = " ".join(vals)
    # tidy up single-value fields to look like the sample
    if row["rainfall (mm)"]:
        m = re.search(r"\d+(?:\.\d+)?", row["rainfall (mm)"])
        if m:
            row["rainfall (mm)"] = f"{float(m.group()):.2f}mm"
    row["weather"] = row["weather"].upper()
    row["Source Message"] = source
    return row


def parse_message(message, source_number, tower_re):
    """Extract one or more date/tower records from a single pasted message."""
    message = (message.replace("\u200e", "").replace("’", "'").replace("‘", "'")
               .replace("*", "").replace("~", "").replace("•", "\n"))
    lines = message.splitlines()

    # Default date/tower = first ones found anywhere in the message.
    first_dt = next((find_date(l)[0] for l in lines if find_date(l)[0]), None)
    def header_tower(l):
        """Tower code from a header-like line only (a code inside a normal sentence, e.g.
        'SN1-5 base', is not the tower)."""
        t, rest = extract_tower(l, tower_re)
        if not t:
            return ""
        rest = WEEKDAY_RE.sub(" ", rest)
        if find_date(rest)[1]:
            s, e = find_date(rest)[1]
            rest = rest[:s] + " " + rest[e:]
        rest = re.sub(r"\b(?:date|tower|location|site|day)\b", " ", rest, flags=re.I)
        words = len(re.findall(r"[A-Za-z]+", rest))
        title = any(p.search(l) for p in SKIP_PATTERNS)
        return t if (words <= 1 or title) else ""

    first_tower = next((header_tower(BULLET_RE.sub("", l).strip().strip("_")) for l in lines
                        if header_tower(BULLET_RE.sub("", l).strip().strip("_"))), "")

    records = []
    state = {"date": fmt_date(first_dt) if first_dt else "", "tower": first_tower, "items": {}, "field": None}

    def has_payload():
        return any(state["items"].values())

    def flush():
        if has_payload():
            records.append(build_row(state["items"], state["date"], state["tower"], f"Message {source_number}"))
        state["items"] = {}
        state["field"] = None

    def add(field, value):
        value = value.strip(" \t:-–|,;")
        if value:
            state["items"].setdefault(field, []).append(value)

    def clean(raw):
        return BULLET_RE.sub("", raw).strip().strip("_").strip()

    # Index of the first real section heading. Free text before it (e.g. the project
    # title) is not data, so it is dropped.
    first_label_idx = next((i for i, l in enumerate(lines) if clean(l) and match_label(clean(l))[0]), None)

    for idx, raw in enumerate(lines):
        line = clean(raw)
        if not line:
            continue

        dt, span = find_date(line)

        # What is left of the line once date / tower / weekday words are removed?
        core = line
        if span:
            core = core[:span[0]] + " " + core[span[1]:]
        tower, core = extract_tower(core, tower_re)
        core = WEEKDAY_RE.sub(" ", core)
        core = re.sub(r"\b(?:date|tower|location|site|day)\b", " ", core, flags=re.I)
        core = core.strip(" :-–|,;&()._\t")
        n_words = len(re.findall(r"[A-Za-z]+", core))

        skip_title = any(p.search(line) for p in SKIP_PATTERNS)

        # Pure header line (date and/or tower, plus maybe a report title) -> may start a new row.
        if (tower or dt) and (n_words <= 1 or skip_title):
            new_date = fmt_date(dt) if dt else ""
            if has_payload() and ((new_date and new_date != state["date"]) or (tower and tower != state["tower"])):
                flush()
            if new_date:
                state["date"] = new_date
            if tower:
                state["tower"] = tower
            continue

        if skip_title or NOISE_LINE_RE.match(line):
            continue

        # Free text before the first heading = project title -> ignore it.
        if first_label_idx is not None and idx < first_label_idx and not guess_field(line):
            continue

        field, value = match_label(line)
        if field:
            state["field"] = field
            add(field, value)
            continue

        if state["field"]:
            add(state["field"], line)       # inside a labelled section
        else:
            add(guess_field(line) or "today's activity", line)

    flush()
    return records


def parse_all(messages, tower_re):
    rows = []
    for i, msg in enumerate(messages, start=1):
        if msg.strip():
            rows.extend(parse_message(msg, i, tower_re))
    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=COLUMNS + ["Source Message"])
    return df[COLUMNS + ["Source Message"]]


# ---------------------------------------------------------------- excel
def natural_key(s):
    return (s == "", [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", s)])


def build_layout(df, fill_gaps):
    """Group rows by tower (sorted by date). Optionally add empty rows for missing days."""
    recs = []
    for r in df.fillna("").to_dict("records"):
        r["_d"] = parse_date_cell(r.get("Date", ""))
        recs.append(r)
    dates = [r["_d"] for r in recs if r["_d"]]
    start, end = (min(dates), max(dates)) if dates else (None, None)

    blocks = []
    for tower in sorted({str(r.get("Tower", "")) for r in recs}, key=natural_key):
        mine = [r for r in recs if str(r.get("Tower", "")) == tower]
        if fill_gaps and tower and start:
            have = {r["_d"] for r in mine if r["_d"]}
            d = start
            while d <= end:
                if d not in have:
                    mine.append({**{c: "" for c in COLUMNS}, "Date": fmt_date(d), "Tower": tower,
                                 "_d": d, "Source Message": ""})
                d += timedelta(days=1)
        mine.sort(key=lambda r: (r["_d"] is None, r["_d"] or datetime.min))
        blocks.append(mine)
    return blocks


def est_lines(text, width):
    chars = max(int(width * 1.05), 5)
    return sum(max(1, ceil(len(part) / chars)) for part in str(text).split("\n"))


def make_formatted_excel(df, fill_gaps=True):
    wb = Workbook()
    ws = wb.active
    ws.title = "Weekly Progress"
    ws.append(COLUMNS)

    header_fill = PatternFill("solid", fgColor="8DB4E2")
    sep_fill = PatternFill("solid", fgColor="FCE4D6")
    header_font = Font(bold=True, color="000000", size=10)
    body_font = Font(size=10)
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    widths = [11, 10, 15, 11, 13, 26, 24, 38, 34, 24, 24, 18, 20]

    for col_idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[1].height = 34

    row_no = 1
    source_map = []
    for b_idx, block in enumerate(build_layout(df, fill_gaps)):
        if b_idx > 0:  # coloured separator row between towers, like the sample
            row_no += 1
            ws.append([""] * len(COLUMNS))
            for c in range(1, len(COLUMNS) + 1):
                ws.cell(row=row_no, column=c).fill = sep_fill
            ws.row_dimensions[row_no].height = 8
        for rec in block:
            row_no += 1
            ws.append([rec.get(c, "") for c in COLUMNS])
            lines = 1
            for c_idx, col in enumerate(COLUMNS, start=1):
                cell = ws.cell(row=row_no, column=c_idx)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border = border
                cell.font = body_font
                lines = max(lines, est_lines(rec.get(col, ""), widths[c_idx - 1]))
            ws.row_dimensions[row_no].height = min(max(15, 13.5 * lines), 400)
            source_map.append((row_no, rec.get("Source Message", "")))

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{row_no}"

    # Source message mapping on a separate sheet for traceability.
    source_ws = wb.create_sheet("Source Mapping")
    source_ws.append(["Excel row", "Source message"])
    for r, s in source_map:
        source_ws.append([r, s])
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


# ---------------------------------------------------------------- UI
st.title("🏗️ Construction Progress Organizer")
st.write(
    "Paste each contractor update into a separate message box. The tool extracts the details "
    "and prepares an Excel report using the same column layout as your sample."
)

with st.sidebar:
    st.header("Settings")
    tower_prefixes = st.text_input(
        "Tower code prefixes (optional)", value=DEFAULT_TOWER_PREFIXES,
        help="Leave blank to auto-detect any tower code (MO72, OF30, KT105 ...). "
             "Or list prefixes, comma separated, e.g. MO, OF, KT, to accept only those."
    )
    fill_gaps = st.checkbox(
        "Add empty rows for missing days", value=True,
        help="Fills every tower with all dates from the earliest to the latest date, like your sample."
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
    result = parse_all(st.session_state.messages, build_tower_re(tower_prefixes))
    st.session_state.result_df = result
    if result.empty:
        st.warning("No usable details were detected. Please check the messages or fill in the missing details manually.")
    else:
        st.success(f"Created {len(result)} row(s). Review the extracted information below.")

if "result_df" in st.session_state and not st.session_state.result_df.empty:
    st.subheader("2. Review and edit the Excel table")
    st.caption("The table is editable. Correct any information that was placed in the wrong column. "
               "Use the format 25-Oct for dates.")
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
    st.download_button(
        "⬇️ Download weekly_progress_report.xlsx",
        data=make_formatted_excel(editable.fillna(""), fill_gaps),
        file_name="weekly_progress_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

st.caption(
    "Prototype limitation: this version extracts information using labels and keywords. "
    "Contractor messages may use different wording, so always verify the generated report."
)
