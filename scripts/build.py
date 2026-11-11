"""Build the course and progression pages from source data in assets/data/.

Usage: python3 scripts/build.py
       python3 scripts/build.py --template templates/clean.html --output clean.html
"""
import html
import argparse
import csv
import json
import os
import re
from datetime import date
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "assets" / "data"
XLSX_PATH = DATA_DIR / "Pharma_DS_Course_Overview.xlsx"
HTML_PATH = ROOT / "index.html"
PROGRESSION_PATH = DATA_DIR / "Pharma_DS_Progression.xlsx"
EVIDENCE_PATH = DATA_DIR / "evidence.csv"

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

# (group, summary, examples) shown in the explanation panel and as tooltips.
DS_DESCRIPTIONS = {
    "Mathematics and statistics": (
        "Foundations",
        "The concepts behind the methods: functions and rates of change, basic calculus, "
        "linear algebra, distributions, uncertainty, hypothesis testing and regression.",
        "Linearization, a degradation rate as a derivative, AUC as an integral, a dataset as a "
        "matrix, standard errors and confidence intervals, t-tests and ANOVA.",
    ),
    "Computing and programming": (
        "Foundations",
        "Digital literacy and the thinking behind programming, so skills carry over between "
        "languages and tools. R is the main teaching language, used to build scripts and workflows "
        "others can rerun. Also covers troubleshooting code and critical use of AI tools.",
        "Organizing files and projects, writing and rerunning R scripts, reading error messages, "
        "finding help, using AI assistants critically when writing code.",
    ),
    "Data governance": (
        "Data workflow",
        "How data is organized, documented, stored and protected so it can be found, understood, "
        "trusted and reused, including the ethical and legal responsibilities of health data.",
        "Tidy data structure, file naming, metadata and the FAIR principles, keeping raw data "
        "unchanged, secure storage, GDPR and anonymization.",
    ),
    "Data handling": (
        "Data workflow",
        "Getting data into a usable shape: importing, cleaning, handling missing values and basic "
        "transformations, done reproducibly.",
        "Importing instrument output or spreadsheets, filtering and selecting, creating variables, "
        "converting units, joining tables, reshaping between wide and long format.",
    ),
    "Exploration": (
        "Data workflow",
        "Getting to know the data before drawing conclusions: quick summaries and plots to see "
        "patterns, catch problems and check whether the data suit the planned analysis.",
        "Histograms, boxplots and scatter plots, summary tables by group, spotting outliers, "
        "checking assumptions such as normality or linearity.",
    ),
    "Analysis and modelling": (
        "Data workflow",
        "Extracting results from data and judging how reliable they are, from statistical analysis "
        "to mechanistic models whose parameters have real meaning.",
        "t-tests and ANOVA, correlation, linear and nonlinear regression, confidence intervals on "
        "parameters, reaction and release kinetics, dose-response and pharmacokinetic models.",
    ),
    "Machine learning and AI": (
        "Data workflow",
        "Methods that learn patterns from data to reduce dimensions, find groups or predict "
        "outcomes, including tuning, validation and a conceptual understanding of modern AI.",
        "PCA, k-means and hierarchical clustering, random forests, training and test sets, "
        "cross-validation and overfitting, the basic ideas behind neural networks and LLMs.",
    ),
    "Visualization and reporting": (
        "Data workflow",
        "Communicating data and results through clear, well-designed figures and transparent "
        "reports that show both the findings and how they were reached.",
        "Choosing the right plot type, figure design, showing uncertainty, reproducible reports "
        "that combine text, code and results (e.g. Quarto), posters and presentations.",
    ),
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
    return f"{float(value):g} ECTS"


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


def parse_notes(value):
    """Parse category-prefixed notes, including indented Markdown-style content."""
    notes = []
    category = None
    content = []

    def save_note():
        if category is not None:
            notes.append((category, "\n".join(content).strip()))

    for line in str(value or "").splitlines():
        candidate, separator, note = line.partition(":")
        candidate = candidate.strip()
        if separator and candidate in DS_CATEGORIES:
            save_note()
            category = candidate
            content = [note.lstrip()]
        elif category is not None:
            content.append(line)
        elif line.strip():
            raise ValueError(f"Course detail content must start with a known category: {line}")
    save_note()
    return notes


def render_note_body(note):
    """Render plain text paragraphs, bold headings, and unordered lists."""
    blocks = []
    paragraph = []
    list_items = []

    def save_paragraph():
        if paragraph:
            blocks.append(f'<p>{" ".join(paragraph)}</p>')
            paragraph.clear()

    def save_list():
        if list_items:
            blocks.append(f'<ul>{"".join(list_items)}</ul>')
            list_items.clear()

    for line in note.splitlines():
        stripped = line.strip()
        if not stripped:
            save_paragraph()
            save_list()
        elif stripped.startswith("- "):
            save_paragraph()
            list_items.append(f"<li>{html.escape(stripped[2:])}</li>")
        elif stripped.startswith("**") and stripped.endswith("**"):
            save_paragraph()
            save_list()
            blocks.append(f"<h4>{html.escape(stripped[2:-2])}</h4>")
        else:
            save_list()
            paragraph.append(html.escape(stripped))
    save_paragraph()
    save_list()
    return "".join(blocks)


def load_details(wb, sheet_name):
    """Optional '<sheet> Details' sheet: one row per course code."""
    name = f"{sheet_name} Details"
    if name not in wb.sheetnames:
        return {}
    ws = wb[name]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    details = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        code = row[idx["Course Code"]]
        if not code:
            continue
        details[str(code)] = {
            "summary": str(row[idx["Summary"]] or "").strip(),
            "primary": parse_notes(row[idx["Primary Notes"]]),
            "secondary": parse_notes(row[idx["Secondary Notes"]]),
        }
    return details


def render_details_template(course, details):
    def note_items(notes, css_class):
        return "".join(
            f'<div class="detail-note {css_class}" style="--ds-color:{DS_CATEGORIES[cat][1]}">'
            f'<strong>{html.escape(cat)}</strong>{render_note_body(note)}</div>'
            for cat, note in notes
        )

    secondary = (
        f'<h3>Also present</h3>{note_items(details["secondary"], "detail-secondary")}'
        if details["secondary"]
        else ""
    )
    return (
        '<template class="course-details">'
        f'<h2>{html.escape(str(course["code"]))}: {html.escape(course["title_en"] or course["title_da"])}</h2>'
        f'<p class="detail-summary">{html.escape(details["summary"])}</p>'
        f'<h3>Main data science focus</h3>{note_items(details["primary"], "detail-primary")}'
        f'{secondary}</template>'
    )


def load_courses(sheet_name):
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
    ws = wb[sheet_name]
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}
    has_ds_category_mapping = (
        "DS Primary Categories" in idx and "DS Secondary Categories" in idx
    )

    details = load_details(wb, sheet_name)
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
            "ds_secondary": parse_categories(row[idx["DS Secondary Categories"]], code)
            if has_ds_category_mapping else [],
            "details": details.get(str(code)),
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
        f'data-ds-cat="{DS_CATEGORIES[category][0]}" '
        f'title="{html.escape(DS_DESCRIPTIONS[category][1], quote=True)}">'
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
    details = course.get("details")
    details_button = (
        f'<button type="button" class="details-btn" onclick="openCourseDetails(this)" '
        f'title="How data science is taught in this course" '
        f'aria-label="Data science details for {code}">i</button>'
        if details and show_ds_categories
        else ""
    )
    details_template = (
        "\n      " + render_details_template(course, details) if details_button else ""
    )
    box_class += " has-details" if details_button else ""
    ds_section = (
        ""
        if show_ds_categories
        else f'      <div class="box-ds"><div class="ds-body">{ds_content}</div></div>'
    )
    ds_main_content = ""
    if show_ds_categories and has_primary_categories:
        ds_main_content = (
            f'        <div class="card-foot">{ds_content}{details_button}</div>'
            if details_button
            else f"        {ds_content}"
        )
    elif show_ds_categories and ds_content:
        ds_main_content = f'        <div class="ds-summary">{ds_content}</div>'

    return f"""    <div class="{box_class}" style="--ects:{ects_value}">
      <div class="{bar_class}"><span class="cat-code">{code_link}</span><div class="box-meta">{github_link}<span class="tag tag-ects" title="Course size in ECTS credits">{ects}</span></div></div>
      <div class="box-main">
        <div class="box-title-en">{title_en}</div>
        <div class="box-title-da">{title_da}</div>
{ds_main_content}
      </div>
{ds_section}{details_template}
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
            f'title="{html.escape(DS_DESCRIPTIONS[category][1], quote=True)}" '
            f'onclick="toggleDSCategoryFilter(this)">{html.escape(category)}</button>'
        )
    groups = []
    for group in dict.fromkeys(d[0] for d in DS_DESCRIPTIONS.values()):
        entries = "".join(
            f'<div class="ds-info-item" style="--ds-color:{DS_CATEGORIES[category][1]}">'
            f'<strong>{html.escape(category)}</strong>'
            f'<p>{html.escape(summary)}</p>'
            f'<p class="ds-info-examples"><em>Examples:</em> {html.escape(examples)}</p></div>'
            for category, (g, summary, examples) in DS_DESCRIPTIONS.items()
            if g == group
        )
        groups.append(
            f'<div class="ds-info-group"><h3>{html.escape(group)}</h3>'
            f'<div class="ds-info-grid">{entries}</div></div>'
        )
    info = (
        '<details class="ds-info"><summary>What do the categories mean?</summary>'
        + "".join(groups)
        + '</details>'
    )
    return (
        '<div class="ds-taxonomy-heading">'
        '<h2>Data science categories</h2>'
        '<p>Coloured badges show data-science categories; course header colours show course categories.</p>'
        '</div>'
        '<div class="ds-all-filter">'
        '<button type="button" class="ds-filter-btn" aria-pressed="false" onclick="toggleDSFilter(this)" '
        'title="Show only courses with data science elements; bold black outlines identify them">'
        'Show DS courses</button>'
        '<span>Black outline = data science</span></div>'
        '<div class="ds-taxonomy-legend" aria-label="Filter by data science category">'
        + "".join(items)
        + '</div>'
        + info
    )


def replace_between(text, start_marker, end_marker, new_content):
    start = text.index(start_marker) + len(start_marker)
    end = text.index(end_marker, start)
    return text[:start] + "\n" + new_content + "\n" + text[end:]


def load_progression(path=PROGRESSION_PATH):
    """Read the editable proposal; fail rather than silently misrepresent a review."""
    wb = openpyxl.load_workbook(path, data_only=False)

    def records(sheet, fields):
        ws = wb[sheet]
        headers = [cell.value for cell in ws[4]]
        missing = set(fields) - set(headers)
        if missing:
            raise ValueError(f"{sheet}: missing headers: {', '.join(sorted(missing))}")
        result = []
        for row in ws.iter_rows(min_row=5):
            if not any(cell.value is not None for cell in row):
                continue
            if any(cell.data_type == "f" for cell in row):
                raise ValueError(f"{sheet}: source tables must contain plain values, not formulas")
            values = {}
            for header, key in fields.items():
                value = row[headers.index(header)].value
                if hasattr(value, "isoformat"):
                    value = value.date().isoformat() if hasattr(value, "date") else value.isoformat()
                values[key] = str(value if value is not None else "").strip()
            result.append(values)
        return result

    targets = records("Targets", {
        "Target ID": "id", "Semester": "semester", "Category": "category",
        "Emphasis": "emphasis", "Short title": "title",
        "Expected student action": "action", "Example task": "example",
        "Support available": "support", "Builds on": "builds_on",
        "Evidence to request": "evidence_requested", "Handover expectation": "handover",
        "Design note": "note",
    })
    review = records("Review", {
        "Target ID": "target_id", "Course(s)": "courses", "Match": "match",
        "Participation": "participation", "Observed student task": "observed_task",
        "Support in practice": "support", "Evidence reference": "evidence",
        "Teacher": "teacher", "Review date": "date", "Next change": "next_change",
        "Notes": "notes", "Priority": "priority",
    })
    links = records("Pilot links", {
        "Link ID": "id", "From target": "from", "To target": "to",
        "Evidence IDs": "evidence_ids", "Candidate change": "change",
        "Confirmation needed": "confirmation", "Decision": "decision",
        "Owner": "owner", "Review date": "date", "Notes": "notes",
    })
    sources = records("Sources", {
        "Source": "title", "Reference": "reference", "Checked on": "date",
        "Use and limitation": "limitation",
    })
    wb.close()
    target_ids = set()
    positions = set()
    emphases = {"Introduce", "Practise", "Extend", "Transfer", "Integrate", "No new target"}
    split_ids = lambda value: [part.strip() for part in value.split(";") if part.strip()]
    for target in targets:
        target_id = target["id"]
        if not target_id or target_id in target_ids:
            raise ValueError(f"Targets: duplicate or empty Target ID: {target_id}")
        target_ids.add(target_id)
        if target["semester"] not in {str(i) for i in range(1, 7)}:
            raise ValueError(f"{target_id}: semester must be an integer from 1 to 6")
        target["semester"] = int(target["semester"])
        if target["category"] not in DS_CATEGORIES or target["emphasis"] not in emphases:
            raise ValueError(f"{target_id}: unknown category or emphasis")
        position = (target["semester"], target["category"])
        if position in positions:
            raise ValueError(f"{target_id}: more than one target for {position}")
        positions.add(position)
        if any(not target[key] for key in ("title", "action", "example", "support", "evidence_requested", "handover", "note")):
            raise ValueError(f"{target_id}: incomplete proposed expectation")
        target["builds_on"] = split_ids(target["builds_on"])
    expected = {(s, c) for s in range(1, 7) for c in DS_CATEGORIES}
    if positions != expected:
        raise ValueError("Targets must cover all eight categories in each of the six semesters")
    by_id = {target["id"]: target for target in targets}
    for target in targets:
        for previous in target["builds_on"]:
            if previous not in by_id or by_id[previous]["semester"] > target["semester"]:
                raise ValueError(f"{target['id']}: unknown or future prerequisite {previous}")

    def visit(target_id, trail):
        if target_id in trail:
            raise ValueError(f"Targets: circular Builds on reference involving {target_id}")
        for previous in by_id[target_id]["builds_on"]:
            visit(previous, trail | {target_id})

    for target_id in by_id:
        visit(target_id, set())
    review_ids = set()
    for row in review:
        target_id = row["target_id"]
        if target_id not in target_ids or target_id in review_ids:
            raise ValueError(f"Review: unknown or duplicate target {target_id}")
        review_ids.add(target_id)
        if row["match"] not in {"Not reviewed", "Aligned", "Partly aligned", "Gap confirmed", "Not applicable"}:
            raise ValueError(f"{target_id}: unknown review status")
        if row["participation"] not in {"To confirm", "Required", "Optional"} or row["priority"] not in {"To discuss", "Candidate", "Selected", "Deferred"}:
            raise ValueError(f"{target_id}: unknown participation or priority")
        if row["match"] != "Not reviewed" and any(not row[key] for key in ("observed_task", "evidence", "teacher", "date")):
            raise ValueError(f"{target_id}: a reviewed judgement needs the observed task, evidence, teacher and date")
        if row["date"]:
            date.fromisoformat(row["date"])
    if review_ids != target_ids:
        raise ValueError("Review must contain exactly one record for each proposed target")
    link_ids = set()
    for link in links:
        if not link["id"] or link["id"] in link_ids:
            raise ValueError("Pilot links: duplicate or empty Link ID")
        link_ids.add(link["id"])
        if link["from"] not in target_ids or link["to"] not in target_ids:
            raise ValueError(f"{link['id']}: unknown handover target")
        if by_id[link["from"]]["semester"] >= by_id[link["to"]]["semester"]:
            raise ValueError(f"{link['id']}: a handover must lead to a later semester")
        if link["decision"] not in {"To discuss", "Selected", "Deferred", "Implemented"}:
            raise ValueError(f"{link['id']}: unknown decision")
        if not link["change"] or not link["confirmation"]:
            raise ValueError(f"{link['id']}: candidate change and confirmation question are required")
        if link["decision"] == "Implemented" and any(not link[key] for key in ("owner", "date", "notes")):
            raise ValueError(f"{link['id']}: an implemented change needs an owner, date and outcome notes")
        if link["date"]:
            date.fromisoformat(link["date"])
        link["evidence_ids"] = split_ids(link["evidence_ids"])
    return {"targets": targets, "review": review, "links": links, "sources": sources}


def build_clean(template, output):
    """Combine recorded course content, proposed progression and pilot evidence."""
    universities = {}
    for uni in UNIVERSITIES:
        semesters, _, mapped = load_courses(uni["sheet"])
        universities[uni["key"]] = {
            "mapped": mapped,
            "courses": [course for courses in semesters.values() for course in courses],
        }
    with EVIDENCE_PATH.open(encoding="utf-8-sig", newline="") as source:
        evidence = list(csv.DictReader(source))
    progression = load_progression(PROGRESSION_PATH)
    evidence_ids = {row["id"] for row in evidence}
    for link in progression["links"]:
        if not set(link["evidence_ids"]) <= evidence_ids:
            raise ValueError(f"{link['id']}: unknown exercise evidence reference")
    payload = {
        "universities": universities,
        "categories": DS_CATEGORIES,
        "descriptions": DS_DESCRIPTIONS,
        "evidence": evidence,
        "progression": progression,
        "updated": date.today().strftime("%d %B %Y").lstrip("0"),
    }
    # Embedded JSON must not be able to terminate its script element.
    data = json.dumps(payload, ensure_ascii=False).replace("<", "\\u003c").replace(
        ">", "\\u003e"
    ).replace("&", "\\u0026")
    text = template.read_text(encoding="utf-8")
    marker = "<!-- CLEAN:DATA -->"
    if text.count(marker) != 1:
        raise ValueError("Alternative template must contain exactly one CLEAN:DATA marker")
    text = text.replace(marker, f'<script id="course-data" type="application/json">{data}</script>')
    # Links and assets remain usable when output is written into a subdirectory.
    base = Path(os.path.relpath(ROOT, output.resolve().parent)).as_posix()
    if base != ".":
        for prefix in ('src="assets/', 'href="assets/', 'href="progression/', 'href="index.html'):
            attribute, value = prefix.split('="', 1)
            text = text.replace(prefix, f'{attribute}="{base}/{value}')
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(text, encoding="utf-8")
    print(f"Wrote {output} from the workbook and {len(evidence)} pilot evidence records.")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--template", type=Path, help="HTML template (defaults to index.html)")
    parser.add_argument("--output", type=Path, help="Destination HTML file")
    args = parser.parse_args()
    template = args.template or HTML_PATH
    text = template.read_text(encoding="utf-8")
    if "<!-- CLEAN:DATA -->" in text:
        if args.output is None:
            parser.error("the alternative template requires --output to preserve the template")
        if template.resolve() == args.output.resolve():
            parser.error("alternative template and output must be different files")
        build_clean(template, args.output)
        return
    output = args.output or (template if args.template else HTML_PATH)
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

    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(text, encoding="utf-8")
    print(f"Wrote {output} from {total_courses} courses across {len(UNIVERSITIES)} universities.")


if __name__ == "__main__":
    main()
