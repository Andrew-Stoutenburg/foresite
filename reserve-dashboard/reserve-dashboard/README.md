# Multi-Property Reserve Study Dashboard (POC)

A static, client-side dashboard for reviewing and adjusting reserve funding
plans across a portfolio of properties, built around the Reserve Advisors
funding model workbook (v7.0 template).

## Run locally

Any static file server works (SheetJS is bundled locally; there are no
network calls at runtime):

```bash
npx serve .          # or: python3 -m http.server 8000
```

Open the printed URL, then drag one or more funding model `.xlsx` workbooks
onto the upload zone. One file → Property View; several at once → Portfolio.

## Deploy to Vercel

```bash
npx vercel deploy    # static deployment, no build step, no backend
```

`vercel.json` pins it as a static site. Nothing else is required.

## What it does

**Property View** (`#/property/<id>`)
- Edit any project: move timing (next-replacement year), change unit cost,
  delete/restore, or phase it across multiple years (equal split, using the
  workbook's own phasing parameters: events-per-phase / spacing).
- Every edit recalculates the reserve balance, interest, and cash flow
  immediately using the workbook's exact formulas (see below).
- Funding plan: per-year editable contribution schedule, plus one-click
  scenarios — Study recommendation, Baseline (minimum funding; binary search
  on a uniform scale factor, ported from the validated `reserve-stress-test`
  algorithm), and Threshold (keep balance ≥ a dollar floor).
- Editable KPI cards: Starting reserves (balance, balance date, current-year
  contribution — the modal explains the resulting year-end balance assuming
  monthly contributions), Interest rate, and Inflation rate. A "Reset to
  study" button restores all of these to the imported values.
- Additional Assessments column in the funding ledger (manual, per year).
- Save funding plans to compare them as separate lines on the chart, with a
  clickable legend to toggle lines and a hover tooltip (contribution,
  year-end balance, and that year's projects).
- Filters: This year / Next year / Next 2 / Next 5 / Past due / All, plus
  keyword search; combinable.

**Portfolio View** (`#/portfolio`)
- Batch upload; cards per property with reserve level and low-point warning.
- Aggregate project table across all properties, filterable by window,
  keyword, cost range, and property.
- Coordination opportunities: same component family needed at ≥2 properties
  within a 3-year window.

## Fidelity to the workbook

The engine (`js/engine.js`) replicates the workbook's formulas, validated
cell-for-cell against the North Brook reference workbook
(`test/validate.js`: 383 checks, including all 1,085 expenditure year cells
and every year's interest and balance):

- Expenditure schedule: next-full-replacement + useful-life cycles, phased
  events via the D/E phase parameters, `ROUND(qty*cost*share*(1+i)^k, 1)`
  with the near-term / remaining-study inflation split.
- Cash flow: `interest = ROUND(rate*(begin + exp/2 + contrib/2))`, partial
  current year via remaining budgeted periods and months of interest.
- Hand-tuned cells: studies often contain manual per-cell edits (blanked
  occurrences, scaled costs like "65% in 2029 / 35% in 2036", static dollar
  entries). The parser detects these by comparing cached values against the
  formula result and preserves them as per-cell overrides ("tuned" badge).

Run the validation suite:

```bash
node test/validate.js "path/to/The North Brook ... Excel.xlsx"
```

## Architecture / future hooks

- `js/parser.js` — SheetJS workbook → clean property/component schema.
  Anchors on labels, flags ambiguity as warnings instead of guessing.
- `js/engine.js` — pure recalculation functions (no DOM/storage).
- `js/store.js` — persistence adapter. POC = localStorage; the async
  interface is designed to swap to IndexedDB (needed at ~100+ properties;
  localStorage caps ~5 MB) or a real DB/API without touching the views.
  **Future auth layer (CRM property permissions) hooks in here.**
- `js/app.js` — hash-routed views. **Future "download my changes" export
  (SheetJS write-back) belongs next to `refreshAfterEdit()`.**
- `js/charts.js` — dependency-free SVG chart.

Out of scope for this POC (per spec): authentication/CRM permissions,
server-side storage, Foresite-hosted Excel ingestion, board-member views,
funded-% / fully-funded-balance metrics.
