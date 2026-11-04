"""Regenerate index.html from Pharma_DS_Course_Overview.xlsx (sheets 'SDU' and 'UCPH').

Usage: python build.py
"""
import html
import re
from datetime import date
from pathlib import Path

import openpyxl

XLSX_PATH = Path(__file__).parent / "Pharma_DS_Course_Overview.xlsx"
HTML_PATH = Path(__file__).parent / "index.html"

# Shared colour pool — categories are assigned slugs/colours in order, per university.
SLUG_POOL = ["cat-quant", "cat-chem", "cat-bio", "cat-physio", "cat-form", "cat-practice"]
COLOR_POOL = ["#DECBE4", "#B3CDE3", "#CCEBC5", "#FBB4AE", "#FED9A6", "#FFFFCC"]

# Data-science category badges use a separate namespace from course-category filters.
DS_CATEGORIES = {
    "Mathematics and statistics": ("ds-mathematics", "#7B61A8"),
    "Computing and programming": ("ds-computing", "#2878B5"),
    "Data governance": ("ds-governance", "#6C757D"),
    "Data handling": ("ds-handling", "#00A6A6"),
    "Exploration": ("ds-exploration", "#4E9F50"),
    "Analysis and modelling": ("ds-analysis", "#F28E2B"),
    "Machine learning and AI": ("ds-machine-learning", "#D9534F"),
    "Visualization and reporting": ("ds-visualization", "#B08D00"),
}

# Fixed order preserves the SDU category->colour assignment already established.
SDU_CATEGORY_ORDER = [
    "Quantitative & Analytical Methods",
    "Chemistry",
    "Biology & Biochemistry",
    "Physiology & Pharmacology",
    "Formulation & Manufacturing",
    "Pharmacy Practice & Society",
]

UNIVERSITIES = [
    {
        "key": "sdu",
        "sheet": "SDU",
        "category_order": SDU_CATEGORY_ORDER,
        "headline_pattern": re.compile(r"Bachelor of Pharmacy – Course Overview \(SDU, [^)]*\)"),
        "headline_fmt": "Bachelor of Pharmacy – Course Overview (SDU, last updated {date})",
        "footer_pattern": re.compile(r"Last updated [^·]*· University of Southern Denmark"),
        "footer_fmt": "Last updated {date} · University of Southern Denmark",
    },
    {
        "key": "ucph",
        "sheet": "UCPH",
        "category_order": None,  # auto-detected: order of first appearance in the sheet
        "headline_pattern": re.compile(r"University of Copenhagen – Course Overview \([^)]*\)"),
        "headline_fmt": "University of Copenhagen – Course Overview (last updated {date})",
        "footer_pattern": re.compile(r"Last updated [^·]*· University of Copenhagen"),
        "footer_fmt": "Last updated {date} · University of Copenhagen",
    },
]


def fmt_ects(value):
    if value is None:
        return ""
    if float(value) == int(value):
        return f"{int(value)} ECTS"
    return f"{value}".replace(".", ",") + " ECTS"


def is_ds_v2_plus(value):
    if value is None or value == "" or isinstance(value, str):
        return False
    return float(value) >= 2


def parse_categories(value, code):
    if value is None or value == "":
        return []
    categories = [category.strip() for category in str(value).split(";") if category.strip()]
    unknown = [category for category in categories if category not in DS_CATEGORIES]
    if unknown:
        raise ValueError(f"Unknown data-science category for {code}: {', '.join(unknown)}")
    return categories


def load_courses(sheet_name):
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    has_ds_category_mapping = (
        "DS Primary Categories" in idx and "DS Secondary Categories" in idx
    )

    courses = []
    category_order = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        code = row[idx["Course Code"]]
        semester = row[idx["Semester"]]
        title_da = row[idx["Danish Title"]]
        if semester is None or (code is None and not title_da):
            continue
        category = str(row[idx["Course Category 1"]] or "").strip()
        if category and category not in category_order:
            category_order.append(category)
        courses.append({
            "code": code,
            "title_da": row[idx["Danish Title"]] or "",
            "title_en": row[idx["English Title"]] or "",
            "ects": row[idx["ECTS"]],
            "semester": int(semester),
            "webpage": row[idx["Course Webpage"]] if row[idx["Course Webpage"]] not in (None, "-") else "#",
            "category": category,
            "ds_version": row[idx["DS Version"]],
            "ds_desc": row[idx["DS Elements"]] or "",
            "github": (row[idx["GitHub Repository"]] or "") if "GitHub Repository" in idx else "",
            "ds_primary": parse_categories(row[idx["DS Primary Categories"]], code)
            if has_ds_category_mapping else [],
        })

    courses.sort(key=lambda c: c["semester"])
    by_semester = {}
    for c in courses:
        by_semester.setdefault(c["semester"], []).append(c)
    return by_semester, category_order, has_ds_category_mapping


def build_category_map(category_order):
    order = category_order[: len(SLUG_POOL)]
    slug_map = {cat: SLUG_POOL[i] for i, cat in enumerate(order)}
    color_map = {cat: COLOR_POOL[i] for i, cat in enumerate(order)}
    return order, slug_map, color_map


URL_RE = re.compile(
    r"(?P<label>[^()<>]*?)\s*\((?P<label_url>https?://[^\s)]+)\)"
    r"|(?P<bare_url>https?://[^\s)]+)"
)


def linkify(escaped_text):
    """Link parenthesized URLs using their preceding text as the label."""
    def replace_url(match):
        label = match.group("label")
        url = match.group("label_url")
        if label is not None:
            return (
                f'<a href="{url}" target="_blank" rel="noopener">'
                f'{label.strip()}</a>'
            )
        url = match.group("bare_url")
        return f'<a href="{url}" target="_blank" rel="noopener">{url}</a>'

    return URL_RE.sub(replace_url, escaped_text)


def render_ds_badges(categories):
    if not categories:
        return ""
    category_order = {category: index for index, category in enumerate(DS_CATEGORIES)}
    badges = "".join(
        f'<span class="ds-badge {DS_CATEGORIES[category][0]}" '
        f'data-ds-cat="{DS_CATEGORIES[category][0]}">'
        f'{html.escape(category)}</span>'
        for category in sorted(categories, key=category_order.__getitem__)
    )
    return (
        f'<div class="ds-category-list" aria-label="Primary data-science categories">'
        f'{badges}</div>'
    )


def render_box(course, slug_map, show_ds_categories=False):
    code = html.escape(str(course["code"])) if course["code"] is not None else ""
    title_da = html.escape(course["title_da"])
    title_en = html.escape(course["title_en"])
    ects = fmt_ects(course["ects"])
    webpage = html.escape(str(course["webpage"]), quote=True)
    ds_desc = linkify(html.escape(course["ds_desc"]))
    has_primary_categories = bool(course["ds_primary"])
    ds_content = (
        render_ds_badges(course["ds_primary"])
        if show_ds_categories and has_primary_categories
        else ds_desc
    )
    has_ds_content = bool(
        course["ds_primary"] or course["ds_desc"] or is_ds_v2_plus(course["ds_version"])
    )

    cat_slug = slug_map.get(course["category"], "")
    ds_content_class = " box-has-ds box-ds-hi" if has_ds_content else ""
    box_class = f"box{ds_content_class} {cat_slug}" if cat_slug else f"box{ds_content_class}"
    bar_class = f"box-cat-bar {cat_slug}" if cat_slug else "box-cat-bar"
    ects_value = float(course["ects"]) if course["ects"] else 1
    code_link = (
        f'<a class="cat-code-link" href="{webpage}" target="_blank" rel="noopener" '
        f'aria-label="Open course page for {code}">{code}</a>'
        if webpage != "#" and code
        else code
    )
    github = str(course["github"]).strip()
    github_link = (
        f'<a class="repo-link" href="{html.escape(github, quote=True)}" target="_blank" '
        f'rel="noopener" title="Open public course materials on GitHub" '
        f'aria-label="Open public GitHub materials for {code}">GitHub</a>'
        if github.startswith("https://")
        else ""
    )
    ds_section = (
        ""
        if show_ds_categories
        else f'      <div class="box-ds"><div class="ds-body">{ds_content}</div></div>'
    )
    ds_main_content = ""
    if show_ds_categories and has_primary_categories:
        ds_main_content = f"        {ds_content}"
    elif show_ds_categories and ds_content:
        ds_main_content = f'        <div class="ds-summary">{ds_content}</div>'

    return f"""    <div class="{box_class}" style="--ects:{ects_value}">
      <div class="{bar_class}"><span class="cat-code">{code_link}</span><div class="box-meta">{github_link}<span class="tag tag-ects" title="Course size in ECTS credits">{ects}</span></div></div>
      <div class="box-main">
        <div class="box-title-da">{title_da}</div>
        <div class="box-title-en">{title_en}</div>
{ds_main_content}
      </div>
{ds_section}
    </div>"""


def render_semester(sem_num, courses, slug_map, show_ds_categories=False):
    season = "Fall" if sem_num % 2 == 1 else "Spring"
    boxes = [render_box(c, slug_map, show_ds_categories) for c in courses]
    if show_ds_categories:
        return f"""  <div class="semester-group">
    <div class="semester-heading"><span class="sem-tag">Semester</span><strong>{sem_num}</strong><span>({season})</span></div>
    <div class="sem-row">
{chr(10).join(boxes)}
    </div>
  </div>"""

    return f"""  <div class="sem-row">
    <div class="sem-label">
      <div class="sem-label-course"><div class="sem-tag">Semester</div><strong>{sem_num}</strong><span>({season})</span></div>
      <div class="sem-label-ds">Data Science Elements</div>
    </div>
{chr(10).join(boxes)}
  </div>"""


def render_grid(by_semester, slug_map, show_ds_categories=False):
    rows = [
        render_semester(sem, by_semester.get(sem, []), slug_map, show_ds_categories)
        for sem in sorted(by_semester)
    ]
    return "\n\n".join(rows)


def render_legend(category_order, slug_map, color_map):
    buttons = []
    for cat in category_order:
        slug = slug_map[cat]
        color = color_map[cat]
        name = html.escape(cat)
        buttons.append(
            f'  <button class="legend-btn" data-cat="{slug}" aria-pressed="false" style="background:{color}" '
            f'onclick="toggleCategoryFilter(this)">{name}</button>'
        )
    return "\n".join(buttons)


def render_ds_legend():
    items = []
    for category, (slug, color) in DS_CATEGORIES.items():
        items.append(
            f'<button type="button" class="ds-taxonomy-btn" data-ds-cat="{slug}" '
            f'aria-pressed="false" style="--ds-color:{color}" '
            f'onclick="toggleDSCategoryFilter(this)">{html.escape(category)}</button>'
        )
    return (
        '<div class="ds-taxonomy-legend" aria-label="Filter by data science category">'
        '<button class="ds-filter-btn" aria-pressed="false" onclick="toggleDSFilter(this)" '
        'title="Show only courses with data science elements; bold black outlines identify them">'
        'Data Science</button>'
        + "".join(items)
        + '</div>'
    )


def replace_between(text, start_marker, end_marker, new_content):
    start = text.index(start_marker) + len(start_marker)
    end = text.index(end_marker, start)
    return text[:start] + "\n" + new_content + "\n" + text[end:]


def main():
    text = HTML_PATH.read_text(encoding="utf-8")
    today = date.today()
    last_updated = f"{today.day} {today.strftime('%B %Y')}"

    total_courses = 0
    for uni in UNIVERSITIES:
        by_semester, detected_order, has_ds_category_mapping = load_courses(uni["sheet"])
        category_order = uni["category_order"] or detected_order
        category_order, slug_map, color_map = build_category_map(category_order)

        grid_html = render_grid(by_semester, slug_map, has_ds_category_mapping)
        legend_html = render_legend(category_order, slug_map, color_map)

        text = replace_between(text, f"<!-- GRID:{uni['key']} -->", f"<!-- /GRID:{uni['key']} -->", grid_html)
        text = replace_between(text, f"<!-- LEGEND:{uni['key']} -->", f"<!-- /LEGEND:{uni['key']} -->", legend_html)
        if has_ds_category_mapping:
            text = replace_between(
                text, f"<!-- DS-LEGEND:{uni['key']} -->",
                f"<!-- /DS-LEGEND:{uni['key']} -->", render_ds_legend()
            )

        headline = uni["headline_fmt"].format(date=last_updated)
        text = uni["headline_pattern"].sub(headline, text)

        footer_text = uni["footer_fmt"].format(date=last_updated)
        text = uni["footer_pattern"].sub(footer_text, text)

        total_courses += sum(len(v) for v in by_semester.values())

    HTML_PATH.write_text(text, encoding="utf-8")
    print(f"Wrote {HTML_PATH} from {total_courses} courses across {len(UNIVERSITIES)} universities.")


if __name__ == "__main__":
    main()
