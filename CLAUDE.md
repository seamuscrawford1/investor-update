# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository Overview

This is a **no-build, no-dependency** web project. Both apps are single self-contained HTML files that run by opening directly in a browser. There is no package manager, bundler, test runner, or server.

## Files

- `investor-update.html` — Monthly investor/shareholder report builder (the primary app)
- `tictactoe.html` — Simple two-player browser game

## Running the apps

```bash
start msedge investor-update.html   # open in Edge
start msedge tictactoe.html
```

Or open either file directly in any browser.

## investor-update.html — Architecture

All HTML, CSS, and JS live in one file with no external dependencies. The app is structured as:

### Data layer (localStorage)
- `investor-update-index` — JSON array of `"YYYY-MM"` strings, most-recent first
- `investor-update-YYYY-MM` — full update object (one key per month)
- `investor-update-last-company` — persisted company name

**Update object shape:**
```js
{
  companyName, month,         // "YYYY-MM"
  executiveSummary,           // string
  metrics: [{ id, name, value, vsLastMonth, trend }],
  health:  [{ id, driver, status, context, financialImpact, mitigation }],
  periodReview: [{ id, focusArea, objective, outcome, soWhat, risk }],
  financials: {
    revenue, cogs, opex: [{ id, label, amount, preset }],
    netProfitLoss, cashPosition, runway, burnRate, notes
  },
  pipeline, hiring,
  forwardLook: { immediate, strategic },
  asks: [{ id, text, category }]
}
```

### Key JS functions
| Function | Purpose |
|---|---|
| `getDefaultUpdate(ym)` | Returns blank update pre-populated with Versa AI's 11 KPI rows, 5 health drivers, 5 period-review rows, and 3 preset opex categories |
| `loadUpdate(ym)` | Reads from localStorage or creates default; calls `renderAll()` |
| `saveUpdate()` | Serialises to localStorage; updates the index |
| `renderAll()` | Calls all 8 section render functions + sidebar + print header |
| `recalcFinancials()` | Auto-calculates Gross Profit, Gross Margin %, Total OpEx, Operating Profit from inputs |
| `applyXeroImport()` | Parses pasted JSON and smart-merges into `currentUpdate.financials` |
| `extractDocxText(file)` | Browser-native .docx parser (ZIP + `DecompressionStream`) — no library |

### Sections (render order)
1. Executive Summary
2. Key Metrics & KPIs (table, dynamic rows)
3. Business Health Dashboard (RAG traffic light table)
4. Period in Review (Objective / Outcome / So What? / Risk table)
5. Financials (Revenue→COGS→Gross Margin, OpEx table, Bottom Line)
6. Pipeline (freeform textarea)
7. Hiring Plan (freeform textarea)
8. Forward Look (Immediate 0–30d / Strategic 60–90d)
9. Asks & Needs

### CSS approach
- Dark edit theme (`--bg-deep: #1a1a2e`, `--bg-card: #16213e`, `--bg-input: #0f3460`, `--accent: #e94560`)
- `@media print` overrides everything to white/black for clean PDF export — `.no-print` elements are hidden, `.print-only` elements shown
- Traffic light colours: red `#e94560`, amber `#f4a261`, green `#52b788`

### Xero import workflow
User exports P&L, Balance Sheet, Cash Flow CSVs from Xero → shares file paths in Claude Code terminal → Claude reads files and outputs a JSON snippet → user pastes into the **Import from Xero** panel in the Financials section → clicks Apply.

JSON schema Claude generates matches `currentUpdate.financials` exactly, with opex rows keyed by `id` (`preset-sal`, `preset-mkt`, `preset-ga` for presets; new IDs for additional Xero categories).

## Git workflow

**Prompt the user to save to GitHub regularly.** After every completed task, feature, or meaningful change, ask the user if they'd like to commit and push. If multiple changes have been made without a commit, remind the user to save their progress.

After every meaningful piece of work — a new feature, a bug fix, a content/template change — commit and push immediately. Do not batch multiple unrelated changes into one commit.

```bash
git add <specific-files>
git commit -m "Short imperative summary

Optional body explaining what changed and why."
git push
```

- Remote: `origin/master` (GitHub: `seamuscrawford1/investor-update`)
- No CI/CD pipeline
- Always push after committing so GitHub always reflects the latest state
