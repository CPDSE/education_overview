# Pharmacy Data Science — Course Overview and Progression

A static overview of data science in pharmacy education at SDU and UCPH, developed by the [Center for Pharmaceutical Data Science Education (CPDSE)](https://www.cpdse.dk).

**Published overview:** [cpdse.github.io/education_overview](https://cpdse.github.io/education_overview/)

## Pages

- **[index.html](index.html)**: the original course overview, with university views and category filters.
- **[clean.html](clean.html)**: the alternative overview with three views: Course overview, Proposed progression, and Exercise evidence. The SDU progression proposal covers six semesters and eight categories; the exercise pilot examines FA515 → FA503 → FA504.

In the original overview, course header colours represent course categories; badges represent data-science categories. The optional **Show course categories and filters** panel provides course-category filters and clears them when closed.

Progression expectations are a discussion draft awaiting teacher agreement. Recorded course categories describe content; they do not establish student proficiency or alignment with a proposed expectation. Unreviewed entries mean unknown, and optional exercises cannot establish a required prerequisite.

## Repository layout

```text
assets/
  data/
    Pharma_DS_Course_Overview.xlsx  # Course mapping for both universities
    Pharma_DS_Progression.xlsx      # Proposed SDU progression and teacher review
    evidence.csv                   # Provisional exercise evidence
    sources.json                   # Evidence provenance and fixed source revisions
  ...                              # Logos and supporting assets
scripts/
  build.py                         # Generates the HTML pages
templates/
  clean.html                       # Editable template for the alternative page
progression/                       # Pilot guide and teacher review documents
tests/                             # Build and data-validation checks
requirements.txt                   # Python build dependencies
index.html                         # Original page; generated sections updated in place
clean.html                         # Generated alternative page
```

## Setup and build

Use Python 3; the deployment workflow uses Python 3.12. From the repository root, create and activate a virtual environment:

```sh
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

On Windows, activate with `.venv\Scripts\activate` instead.

Build both pages with these two commands:

```sh
python scripts/build.py
python scripts/build.py --template templates/clean.html --output clean.html
```

| Output | Inputs |
|---|---|
| `index.html` | `assets/data/Pharma_DS_Course_Overview.xlsx` and the existing page structure |
| `clean.html` | Both Excel workbooks, `assets/data/evidence.csv`, and `templates/clean.html` |

Each invocation builds one page; there is no `--all` option. The original build updates its generated sections and displayed dates in place. The alternative build requires a separate output and preserves its template. Neither build edits the source data.

The script locates source data relative to the repository, independently of the working directory. Custom `--template` and `--output` paths are relative to the working directory unless absolute paths are supplied. Run `python scripts/build.py --help` for the available options.

## Preview and share

After building, serve the repository locally:

```sh
python -m http.server 8731
```

Open [the original overview](http://localhost:8731/index.html) or [the alternative overview](http://localhost:8731/clean.html). You can also open the generated HTML files directly. For an offline copy, include `assets/` and `progression/` alongside the HTML files so images, downloads and review links remain available.

## Update the content

- Edit course data in [Pharma_DS_Course_Overview.xlsx](assets/data/Pharma_DS_Course_Overview.xlsx), using the `SDU`, `UCPH` and optional course-detail sheets. Preserve headers and exact category names.
- Edit proposed expectations in [Pharma_DS_Progression.xlsx](assets/data/Pharma_DS_Progression.xlsx). **Targets** is the editable source; **Semester map** recalculates from it. **Review** records actual tasks, support, participation, evidence, teacher and date. **Pilot links** records candidate handovers; **Sources** explains their basis. Keep target IDs stable.
- Edit [evidence.csv](assets/data/evidence.csv) for exercise interpretations and [sources.json](assets/data/sources.json) for provenance. The JSON documents the evidence; it is not a build input. Follow the [pilot guide](progression/README.md) and [teacher review workflow](progression/workflow.md).
- Edit the alternative layout in `templates/clean.html`, preserving its `CLEAN:DATA` marker. Direct layout edits to generated `clean.html` are overwritten. For the original page, edit `index.html` outside the generated `GRID`, `LEGEND` and `DS-LEGEND` sections.

Save changes, then rebuild the relevant page. The alternative build validates the complete 6 × 8 proposal, target references, review records and exercise links. Its category comparison includes labelled primary and secondary records; course-overview filtering uses primary categories. FA516 remains outside the required-course progression proposal.

## Checks and deployment

Run the automated checks from the repository root:

```sh
python -m unittest discover -s tests -v
```

After visual changes, inspect both university views, filters, dialogs, narrow screens and landscape print preview.

The supplied [GitHub Actions workflow](.github/workflows/deploy.yml) installs dependencies, runs the checks, builds both pages and deploys the site to GitHub Pages on pushes to `main`. Both outputs stay at the repository root. There is no Node.js build step.
