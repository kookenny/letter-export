# Letter Exporter

Export letter content from CaseWare Cloud SE author templates to Word (.docx).

## Quick Start

```bash
# Install dependencies
pip install -r requirements.txt
npm install

# Copy and fill in credentials
cp .env.example .env
# Edit .env with your CaseWare Cloud OAuth credentials

# Start the web UI
python web/app.py
# Open http://localhost:5001
```

Paste a CaseWare letter URL (e.g. `https://ca.cwcloudpartner.com/.../index.jsp#/letter/<id>`) and click **Export Letter** to download the Word file.

## CLI Usage

```bash
# Export a letter to Word
python tools/letter_extract.py --url "https://ca.cwcloudpartner.com/ca-develop/e/eng/<engagementId>/index.jsp#/letter/<documentId>"

# Inspect raw API data (discovery mode)
python tools/letter_extract.py --url "<url>" --discover

# Export JSON only (skip Word generation)
python tools/letter_extract.py --url "<url>" --json-only
```

## Authentication

Set OAuth credentials in `.env` (per-region):
- `CW_CA_CLIENT_ID` / `CW_CA_CLIENT_SECRET` for Canadian environments
- `CW_US_CLIENT_ID` / `CW_US_CLIENT_SECRET` for US environments

Fallback: set `CW_COOKIES` from browser DevTools.

## Output

The exported Word document includes:
- Letter body content with formatting (bold, italic, lists, tables)
- Placeholders as `(( placeholder ))` and dynamic text as `[[ value ]]`
- Guidance callouts in yellow boxes
- Visibility conditions inline and as a summary table

## Project Structure

```
tools/letter_extract.py      # Python: API extraction -> JSON
tools/generate_docx.js       # Node.js: JSON -> Word (.docx)
web/app.py                   # Flask backend (port 5001)
web/static/                  # Frontend UI
.env.example                 # Credential template
```
