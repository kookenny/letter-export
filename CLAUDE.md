# Letter Exporter

Extracts letter content from CaseWare Cloud SE author templates and generates formatted Word documents (.docx). Includes letter body content, guidance callouts, placeholder/dynamic text markers, and visibility settings.

## Quick Start (Local)

```bash
# Install dependencies
pip install -r requirements.txt
npm install

# Copy and fill in credentials
cp .env.example .env

# Web UI (recommended)
python web/app.py
# -> http://localhost:5001

# CLI - single letter
python tools/letter_extract.py --url "<caseware-url>"

# Discovery mode - inspect raw API data
python tools/letter_extract.py --url "<caseware-url>" --discover

# JSON only (skip Word generation)
python tools/letter_extract.py --url "<caseware-url>" --json-only
```

## Vercel Deployment

The app is also deployed to Vercel as a pure Node.js serverless function.

- `api/generate.js` — single serverless function (extraction + Word generation in Node.js)
- `public/` — static frontend served by Vercel
- `vercel.json` — 60s max duration for API calls

Environment variables (`CW_CA_CLIENT_ID`, `CW_CA_CLIENT_SECRET`, `CW_US_CLIENT_ID`, `CW_US_CLIENT_SECRET`) must be set in the Vercel dashboard.

The local development stack (Python + Node.js in `tools/` and `web/`) is preserved alongside the Vercel deployment.

## Authentication

Two methods, checked in order:

1. **OAuth** (preferred) - set per-region credentials in `.env`:
   - `CW_CA_CLIENT_ID` / `CW_CA_CLIENT_SECRET` for Canadian environments
   - `CW_US_CLIENT_ID` / `CW_US_CLIENT_SECRET` for US environments
2. **Cookie fallback** - set `CW_COOKIES` from browser DevTools

The region prefix is derived from the hostname (e.g. `ca.cwcloudpartner.com` -> `CA`).

## URL Format

Letters require a document fragment in the URL:
`https://{host}/{tenant}/e/eng/{engagementId}/index.jsp#/letter/{documentId}`

## Architecture

**Local development (two-language pipeline):**
1. **Python** (`tools/letter_extract.py`) - authenticates, fetches sections via CaseWare API, resolves visibility IDs, outputs structured JSON
2. **Node.js** (`tools/generate_docx.js`) - reads JSON, generates formatted Word document using docx-js

The Python script shells out to Node.js for the Word generation step.

**Vercel (single-language):**
- `api/generate.js` combines both extraction and Word generation in pure Node.js

## Output

Word document (landscape orientation) containing:
- **Letter body** - formatted text with bold, italic, lists, tables preserved
- **Placeholders** - shown as `(( placeholder ))` in blue (#0971AA)
- **Dynamic text** - shown as `[[ value ]]` in grey
- **Guidance callouts** - yellow-shaded boxes (matching the CaseWare UI)
- **Visibility (inline)** - grey italic annotations below each section with conditions
- **Visibility (summary table)** - table at end of document listing all visibility conditions

## Key API Details

- Letters are document type `letter` in CaseWare Cloud
- Letter sections are almost entirely `content` type (59 in the test template), plus `grouping` (root), `pagebreak`, and `settings`
- Guidance lives in `section.guidances.en`
- Visibility conditions use the same structure as note-visibility: `normallyVisible`, `allConditionsNeeded`, `conditions[]` with `response`, `condition_group`, `organization_type`, and `consolidation` types
- Visibility often lives on parent/ancestor sections and must be inherited

## Key Gotchas Discovered

- **Section ordering**: The `order` field uses fractional-index strings (`!~|`, `"`, `#`). In JavaScript, do NOT use `localeCompare` — it reorders special characters. Use simple `<`/`>` comparison.
- **Formula span whitespace**: Formula spans contain a single space `" "` as content. When they resolve to empty, this leaves double spaces. Always collapse whitespace after preprocessing.
- **List continuation**: CaseWare splits numbered lists across multiple `<ol>` elements. Use the `start` attribute to detect continuation (`start > 1` = continue, `start = 1` = new list).
- **Placeholder HTML nesting**: Placeholder spans contain nested `<span>` elements (display text + caret helper). Use a DOM parser (cheerio), not regex, to handle them correctly.

## Project Structure

```
api/generate.js              # Vercel: serverless function (extraction + docx)
public/                      # Vercel: static frontend
tools/letter_extract.py      # Local: Python API extraction -> JSON
tools/generate_docx.js       # Local: Node.js JSON -> Word (.docx)
web/app.py                   # Local: Flask backend (port 5001)
web/static/                  # Local: frontend (vanilla HTML/JS/CSS)
workflows/export_letter.md   # Full workflow SOP
vercel.json                  # Vercel config
.env.example                 # Credential template
```
