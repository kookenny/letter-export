# Letter Exporter

Extracts letter content from CaseWare Cloud SE author templates and generates formatted Word documents (.docx). Includes letter body content, guidance callouts, placeholder/dynamic text markers, and visibility settings.

## Quick Start

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

**Two-language pipeline:**
1. **Python** (`tools/letter_extract.py`) - authenticates, fetches sections via CaseWare API, resolves visibility IDs, outputs structured JSON
2. **Node.js** (`tools/generate_docx.js`) - reads JSON, generates formatted Word document using docx-js

The Python script shells out to Node.js for the Word generation step.

## Output

Word document containing:
- **Letter body** - formatted text with bold, italic, lists, tables preserved
- **Placeholders** - shown as `(( placeholder ))` in grey
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

## Project Structure

```
tools/letter_extract.py      # Python: API extraction -> JSON
tools/generate_docx.js       # Node.js: JSON -> Word (.docx)
web/app.py                   # Flask backend (port 5001)
web/static/                  # Frontend (vanilla HTML/JS/CSS)
workflows/export_letter.md   # Full workflow SOP
.env.example                 # Credential template
```
