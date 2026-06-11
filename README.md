# 6S Audit Form

A reusable **6S workplace-organization audit** toolkit (Sort · Set in Order · Shine · Standardize · Sustain · Safety) for warehouse/facility areas. Each location gets its own copy, fills out blank printable forms during a walk-through, and can add rooms as needed.

## What's here

```
apps-script/Code.gs              Google Sheets "6S Audit Tools" menu (add-room column + printable-PDF export)
template/build_template.py       openpyxl generator that builds the workbook from scratch
template/6S-Audit-Template.xlsx  The generated template (imports cleanly into Google Sheets)
docs/                            Notes & usage
```

## The template

A multi-tab workbook (`Start Here · 6S Overview · 6S Audit Sheet · Facility Summary`):

- **6S Audit Sheet** — 26 standard 6S items plus blank custom rows under each category, each scored **0–3** (`0 = Major Issue, 1 = Needs Work, 2 = Good, 3 = Excellent`).
- **Dynamic scoring** — Total / Max / Score % / # of Zeros / PASS-FAIL auto-size to however many items a room has, and stay blank on an unfilled form so blank PDFs print clean.
- **Pass rule** — score ≥ 67% of max **and** no zeros on any item. Any Safety item scored 0 is escalated to management same day.
- **Location** flows from the *Start Here* tab into every form.

Regenerate it with:

```bash
python -m pip install openpyxl
python template/build_template.py    # writes 6S Audit Template.xlsx
```

## Google Sheets menu (`apps-script/Code.gs`)

Install: **Extensions ▸ Apps Script** → paste `Code.gs` → Save → reload the sheet. Authorize on first run (Sheets + Drive + external request).

Adds a **6S Audit Tools** menu:

- **➕ Add room column** — prompts for a room name and inserts a 0–3 scoring column right of the questions (column C).
- **🖨️ Download printable PDF** — exports the active tab as a clean, blank, portrait, fit-to-width PDF (saved to Drive ▸ `6S Audit PDFs`, with a download link).

## Workflow

1. Set the location on *Start Here*.
2. Add rooms (a column per room, or duplicate the tab per room).
3. Export blank PDFs and audit by hand — or score digitally for auto totals.
4. Roll up results on *Facility Summary*.
