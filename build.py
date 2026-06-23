"""Regenerate index.html from Pharma_DS_Course_Overview.xlsx (sheet 'SDU').

Usage: python build.py
"""
import html
from pathlib import Path

import openpyxl

XLSX_PATH = Path(__file__).parent / "Pharma_DS_Course_Overview.xlsx"
HTML_PATH = Path(__file__).parent / "index.html"
SHEET_NAME = "SDU"
COLS_PER_ROW = 6

SEMESTER_TAGS = {1: "E24", 2: "F25", 3: "E25", 4: "F26", 5: "E26", 6: "F27"}


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
    version_tag = f'<span class="tag tag-new">{html.escape(version_label)}</span>' if version_label else ""

    return f"""    <div class="box">
      <div class="box-main">
        <div class="box-code">{code}</div>
        <div class="box-title-da">{title_da}</div>
        <div class="box-title-en">{title_en}</div>
        <div class="box-meta"><span class="tag tag-ects">{ects}</span>{version_tag}</div>
        <a class="box-link" href="{webpage}" target="_blank">&#8599; {link_label}</a>
      </div>
      <div class="box-ds"><button class="ds-toggle" aria-expanded="false" onclick="toggleDS(this)">DS elements <span class="ds-chevron">&#9660;</span></button><div class="ds-body">{ds_desc}</div></div>
    </div>"""


def render_semester(sem_num, courses):
    tag = SEMESTER_TAGS.get(sem_num, "")
    boxes = [render_box(c) for c in courses]
    while len(boxes) < COLS_PER_ROW:
        boxes.append('    <div class="box-empty"></div>')

    return f"""  <div class="sem-row">
    <div class="sem-label">
      <div class="sem-label-course"><div class="sem-tag">Sem.</div><strong>Sem. {sem_num}</strong><span>({tag})</span></div>
      <div class="sem-label-ds"><div class="sem-tag">DS elem.</div></div>
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
    start_marker = '<div class="grid">'
    end_marker = '</div><!-- /grid -->'
    start = text.index(start_marker) + len(start_marker)
    end = text.index(end_marker)
    new_text = text[:start] + "\n\n" + grid_html + "\n\n" + text[end:]

    HTML_PATH.write_text(new_text, encoding="utf-8")
    print(f"Wrote {HTML_PATH} from {sum(len(v) for v in by_semester.values())} courses.")


if __name__ == "__main__":
    main()
