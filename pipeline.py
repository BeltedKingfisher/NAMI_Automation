"""Core pipeline: uploaded survey files -> OCR markdown -> parsed rows -> .xlsx

No web framework in here on purpose. Streamlit, FastAPI, or a CLI can all call
process_uploads(). Swap the UI later without touching this file.
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

PAGES_PER_SURVEY = 3
MAX_PAGES_PER_BATCH = 9  # 3 surveys per call; Surya's limit is 10 pages


# --------------------------------------------------------------------------
# 1. OCR: plug your Surya / Datalab call in here
# --------------------------------------------------------------------------
def ocr_pages_to_markdown(page_images: list[bytes]) -> list[str]:
    """Take up to 10 page images (PNG/JPEG bytes), return one markdown string
    per page, in the same order. This is the ONLY function that touches Surya.
    """
    raise NotImplementedError("Call Surya here and return one markdown string per page.")


# --------------------------------------------------------------------------
# 2. Loading and grouping pages into surveys
# --------------------------------------------------------------------------
@dataclass
class Survey:
    source: str                      # filename (or "a.jpg..c.jpg" for grouped images)
    pages: list[bytes]               # image bytes, one per page
    notes: list[str] = field(default_factory=list)


def pdf_to_page_images(data: bytes, scale: float = 2.0) -> list[bytes]:
    import pypdfium2 as pdfium

    pdf = pdfium.PdfDocument(data)
    out = []
    for i in range(len(pdf)):
        buf = io.BytesIO()
        pdf[i].render(scale=scale).to_pil().save(buf, format="PNG")
        out.append(buf.getvalue())
    return out


def group_into_surveys(files: list[tuple[str, bytes]]) -> list[Survey]:
    """files = [(filename, bytes), ...]

    - each PDF is treated as one survey (expected: 3 pages)
    - loose images are sorted by filename and grouped 3 at a time
    """
    surveys: list[Survey] = []
    images: list[tuple[str, bytes]] = []

    for name, data in files:
        if name.lower().endswith(".pdf"):
            s = Survey(source=name, pages=pdf_to_page_images(data))
            if len(s.pages) != PAGES_PER_SURVEY:
                s.notes.append(f"Expected {PAGES_PER_SURVEY} pages, got {len(s.pages)}")
            surveys.append(s)
        else:
            images.append((name, data))

    images.sort(key=lambda x: x[0])
    for i in range(0, len(images), PAGES_PER_SURVEY):
        chunk = images[i : i + PAGES_PER_SURVEY]
        s = Survey(source=", ".join(n for n, _ in chunk), pages=[d for _, d in chunk])
        if len(chunk) != PAGES_PER_SURVEY:
            s.notes.append(f"Expected {PAGES_PER_SURVEY} pages, got {len(chunk)}")
        surveys.append(s)
    return surveys


def make_batches(surveys: list[Survey]) -> list[list[Survey]]:
    """Pack whole surveys into batches of <= MAX_PAGES_PER_BATCH pages,
    never splitting a survey across two OCR calls."""
    batches: list[list[Survey]] = []
    current: list[Survey] = []
    count = 0
    for s in surveys:
        if len(s.pages) > MAX_PAGES_PER_BATCH:
            s.pages = s.pages[:MAX_PAGES_PER_BATCH]
            s.notes.append(f"Too many pages; only first {MAX_PAGES_PER_BATCH} processed")
        if current and count + len(s.pages) > MAX_PAGES_PER_BATCH:
            batches.append(current)
            current, count = [], 0
        current.append(s)
        count += len(s.pages)
    if current:
        batches.append(current)
    return batches


# --------------------------------------------------------------------------
# 3. Parsing one page of Surya markdown into {field: value}
# --------------------------------------------------------------------------
CHECKED = "☒☑"


def parse_page(md: str, page_no: int) -> tuple[dict[str, str], dict[str, str]]:
    """Returns (values, flags). flags maps field -> reason it needs review."""
    values: dict[str, str] = {}
    flags: dict[str, str] = {}
    p = f"P{page_no}"

    # --- Likert table(s): rows with a label in col 0 and marks in other cols
    table_lines = [l for l in md.splitlines() if l.strip().startswith("|")]
    if table_lines:
        rows = [[c.strip() for c in l.strip().strip("|").split("|")] for l in table_lines]
        rows = [r for r in rows if not all(re.fullmatch(r":?-{2,}:?", c) or c == "" for c in r)]
        header, body = rows[0], rows[1:]
        for r in body:
            label = f"{p}: {r[0]}"
            marked = [header[i] for i in range(1, min(len(r), len(header))) if r[i]]
            if len(marked) == 1:
                values[label] = marked[0]
            else:
                values[label] = ", ".join(marked)
                flags[label] = "no mark found" if not marked else "multiple marks"

    # --- Numbered questions and consent, using blank-line-separated blocks
    text_no_table = "\n".join(l for l in md.splitlines() if not l.strip().startswith("|"))
    blocks = [b.strip() for b in re.split(r"\n\s*\n", text_no_table) if b.strip()]

    def is_question(b: str) -> bool:
        return re.match(r"^\d+\.\s", b) is not None

    def is_consent(b: str) -> bool:
        return b.lower().startswith("i am willing")

    i = 0
    while i < len(blocks):
        b = blocks[i]
        if is_question(b):
            num = re.match(r"^(\d+)\.", b).group(1)
            key = f"{p} Q{num}"
            nxt = blocks[i + 1] if i + 1 < len(blocks) else ""
            if nxt and not is_question(nxt) and not is_consent(nxt):
                values[key] = nxt
                i += 2
            else:
                values[key] = ""
                flags[key] = "blank answer"
                i += 1
        elif is_consent(b):
            key = f"{p} Consent"
            nxt = blocks[i + 1] if i + 1 < len(blocks) else ""
            yes = re.search(rf"[{CHECKED}]\s*Yes", nxt) is not None
            no = re.search(rf"[{CHECKED}]\s*No", nxt) is not None
            if yes != no:
                values[key] = "Yes" if yes else "No"
            else:
                values[key] = "Yes" if yes else ""
                flags[key] = "both boxes marked" if yes else "neither box marked"
            i += 2
        else:
            i += 1  # headings, stray text

    return values, flags


# --------------------------------------------------------------------------
# 4. Run OCR + parse for every survey
# --------------------------------------------------------------------------
@dataclass
class Row:
    source: str
    values: dict[str, str]
    flags: dict[str, str]


def process_uploads(files: list[tuple[str, bytes]], progress=None) -> list[Row]:
    surveys = group_into_surveys(files)
    batches = make_batches(surveys)
    rows: list[Row] = []

    for bi, batch in enumerate(batches):
        all_pages = [pg for s in batch for pg in s.pages]
        md_pages = ocr_pages_to_markdown(all_pages)

        cursor = 0
        for s in batch:
            values, flags = {}, {}
            for page_no in range(1, len(s.pages) + 1):
                v, f = parse_page(md_pages[cursor], page_no)
                values.update(v)
                flags.update(f)
                cursor += 1
            if s.notes:
                flags["Notes"] = "; ".join(s.notes)
            rows.append(Row(source=s.source, values=values, flags=flags))

        if progress:
            progress((bi + 1) / len(batches))
    return rows


# --------------------------------------------------------------------------
# 5. Excel output
# --------------------------------------------------------------------------
def rows_to_xlsx(rows: list[Row]) -> bytes:
    columns: list[str] = []
    for r in rows:
        for k in r.values:
            if k not in columns:
                columns.append(k)

    wb = Workbook()
    ws = wb.active
    ws.title = "Responses"
    headers = ["Source"] + columns + ["Review notes"]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        ws.cell(row=1, column=c).font = Font(name="Arial", bold=True)

    flag_fill = PatternFill("solid", start_color="FFFF00")
    for ri, r in enumerate(rows, start=2):
        ws.cell(row=ri, column=1, value=r.source)
        for ci, col in enumerate(columns, start=2):
            cell = ws.cell(row=ri, column=ci, value=r.values.get(col, ""))
            if col in r.flags:
                cell.fill = flag_fill
        notes = [f"{k}: {v}" for k, v in r.flags.items()]
        ws.cell(row=ri, column=len(headers), value="; ".join(notes))

    for ci, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(ci)].width = min(max(len(h) + 2, 14), 50)
    ws.freeze_panes = "B2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()