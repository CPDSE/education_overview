"""Regenerate index.html from Pharma_DS_Course_Overview.xlsx (sheet 'SDU').

Usage: python build.py
"""
import html
import re
from datetime import date
from pathlib import Path

import openpyxl

XLSX_PATH = Path(__file__).parent / "Pharma_DS_Course_Overview.xlsx"
HTML_PATH = Path(__file__).parent / "index.html"
SHEET_NAME = "SDU"
COLS_PER_ROW = 6
HEADLINE_PATTERN = re.compile(
    r"Bachelor of Pharmacy – Course Overview \(SDU, [^)]*\)"
)
FOOTER_PATTERN = re.compile(
    r"Last updated [^·]*· University of Southern Denmark"
)

CATEGORY_SLUGS = {
    "Quantitative & Analytical Methods": "cat-quant",
    "Chemistry": "cat-chem",
    "Biology & Biochemistry": "cat-bio",
    "Physiology & Pharmacology": "cat-physio",
    "Formulation & Manufacturing": "cat-form",
    "Pharmacy Practice & Society": "cat-practice",
}


def fmt_ects(value):
    if value is None:
        return ""
    if float(value) == int(value):
        return f"{int(value)} ECTS"
    return f"{value}".replace(".", ",") + " ECTS"


def fmt_ds_version(value):
    if value is None or value == "":
        return ""
    if isinstance(value, str):
        return f"DS {value}"
    if float(value) == int(value):
        return f"DS v{int(value)}"
    return f"DS v{value}"


def fmt_ds_tooltip(value):
    if value is None or value == "":
        return ""
    if isinstance(value, str):
        return "DS upgrade version not yet determined."
    version = float(value)
    if version == 1:
        return "No DS upgrade."
    if version == 2:
        return "Initial DS upgrade."
    if version == 3:
        return "Final version."
    if version > 2:
        return "Iterative refinement of the DS upgrade."
    return "DS upgrade version."


def load_courses():
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
    ws = wb[SHEET_NAME]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}

    courses = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        code = row[idx["Course Code"]]
        semester = row[idx["Semester"]]
        if code is None or semester is None:
            continue
        courses.append({
            "code": code,
            "title_da": row[idx["Danish Title"]] or "",
            "title_en": row[idx["English Title"]] or "",
            "ects": row[idx["ECTS"]],
            "semester": int(semester),
            "webpage": row[idx["Course Webpage"]] or "#",
            "category": str(row[idx["Course Category 1"]] or "").strip(),
            "ds_version": row[idx["DS Version"]],
            "ds_desc": row[idx["DS Elements"]] or "",
        })

    courses.sort(key=lambda c: c["semester"])
    by_semester = {}
    for c in courses:
        by_semester.setdefault(c["semester"], []).append(c)
    return by_semester


def render_box(course):
    code = html.escape(str(course["code"]))
    title_da = html.escape(course["title_da"])
    title_en = html.escape(course["title_en"])
    ects = fmt_ects(course["ects"])
    webpage = html.escape(str(course["webpage"]), quote=True)
    ds_desc = html.escape(course["ds_desc"])
    link_label = "Elective catalogue" if course["code"] == "—" else "Course page (SDU)"

    version_label = fmt_ds_version(course["ds_version"])
    version_tooltip = html.escape(fmt_ds_tooltip(course["ds_version"]), quote=True)
    version_tag = (
        f'<span class="tag tag-new" title="{version_tooltip}">{html.escape(version_label)}</span>'
        if version_label else ""
    )

    cat_slug = CATEGORY_SLUGS.get(course["category"], "")
    box_class = f"box {cat_slug}" if cat_slug else "box"

    return f"""    <div class="{box_class}">
      <div class="box-main">
        <div class="box-code">{code}</div>
        <div class="box-title-da">{title_da}</div>
        <div class="box-title-en">{title_en}</div>
        <div class="box-meta"><span class="tag tag-ects" title="Course size in ECTS credits">{ects}</span>{version_tag}</div>
        <a class="box-link" href="{webpage}" target="_blank">&#8599; {link_label}</a>
      </div>
      <div class="box-ds"><button class="ds-toggle" aria-expanded="false" onclick="toggleDS(this)">DS elements <span class="ds-chevron">&#9660;</span></button><div class="ds-body">{ds_desc}</div></div>
    </div>"""


def render_semester(sem_num, courses):
    season = "Efterår" if sem_num % 2 == 1 else "Forår"
    boxes = [render_box(c) for c in courses]
    while len(boxes) < COLS_PER_ROW:
        boxes.append('    <div class="box-empty"></div>')

    return f"""  <div class="sem-row">
    <div class="sem-label">
      <div class="sem-label-course"><div class="sem-tag">Semester</div><strong>{sem_num}</strong><span>({season})</span></div>
    </div>
{chr(10).join(boxes)}
  </div>"""


def render_grid(by_semester):
    rows = [render_semester(sem, by_semester.get(sem, [])) for sem in sorted(by_semester)]
    return "\n\n".join(rows)


def main():
    by_semester = load_courses()
    grid_html = render_grid(by_semester)

    text = HTML_PATH.read_text(encoding="utf-8")

    today = date.today()
    last_updated = f"{today.day} {today.strftime('%B %Y')}"
    headline = f"Bachelor of Pharmacy – Course Overview (SDU, last updated {last_updated})"
    text = HEADLINE_PATTERN.sub(headline, text)

    footer_text = f"Last updated {last_updated} · University of Southern Denmark"
    text = FOOTER_PATTERN.sub(footer_text, text)

    start_marker = '<div class="grid">'
    end_marker = '</div><!-- /grid -->'
    start = text.index(start_marker) + len(start_marker)
    end = text.index(end_marker)
    new_text = text[:start] + "\n\n" + grid_html + "\n\n" + text[end:]

    HTML_PATH.write_text(new_text, encoding="utf-8")
    print(f"Wrote {HTML_PATH} from {sum(len(v) for v in by_semester.values())} courses.")


if __name__ == "__main__":
    main()
