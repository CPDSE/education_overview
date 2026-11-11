# Repository Guidelines

## Project Structure & Module Organization

This repository publishes a static pharmacy course overview for SDU and UCPH.

- `Pharma_DS_Course_Overview.xlsx`: course data in the `SDU` and `UCPH` sheets; optional `<sheet> Details` sheets provide course explanations.
- `build.py`: reads the workbook with `openpyxl` and updates course grids, legends, and dates in `index.html`.
- `index.html`: page structure, inline CSS and JavaScript, and generated content.
- `assets/`: CPDSE and university logos.
- `.github/workflows/deploy.yml`: builds and deploys to GitHub Pages on pushes to `main`, using Python 3.12.

## Build, Test, and Development Commands

Run commands from the repository root, preferably in an activated Python virtual environment:

- `python -m pip install openpyxl`: install the build dependency.
- `python build.py`: regenerate workbook-derived content in `index.html`; also updates displayed dates.
- `python -m http.server 8731`: serve the site at `http://localhost:8731` for browser review.

There is no Node.js build pipeline or configured formatter, linter, or automated test suite.

## Editing, Style & Naming Conventions

Edit course data in the workbook and regenerate the page. Edit layout, CSS, and browser behavior in `index.html`. Change rendering logic in `build.py`; direct edits inside generated sections will be overwritten. Preserve the `GRID`, `LEGEND`, and `DS-LEGEND` comment markers.

Use four-space Python indentation, `snake_case` functions, and uppercase constants. Match the existing two-space CSS/JavaScript indentation, `camelCase` JavaScript functions, and hyphenated CSS classes. Preserve workbook headers and exact category names from `DS_CATEGORIES`. Escape spreadsheet text when rendering HTML.

## Testing Guidelines

No coverage threshold or test naming convention is configured. After functional or data changes, run the build and inspect both university views. Check category filters, combined filters, course details, external links, narrow screens, and landscape print preview. Confirm course titles, semester placement, and ECTS against the workbook.

## Commit & Pull Request Guidelines

Existing commits use short, descriptive subjects such as “Include data science categories and GitHub links”; no formal prefix scheme is established. Keep changes focused. PR descriptions should explain the change and validation performed; include screenshots for visual changes and link relevant issues when available.

## Agent Instructions

Do not run Git commands automatically. The repository owner handles commits; use Git only when explicitly requested.
