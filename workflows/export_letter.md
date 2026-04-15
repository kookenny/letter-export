# Workflow: Export Letter to Word

## Objective

Extract the full content of a CaseWare Cloud SE letter template and export it to a formatted Word document, including body content, guidance callouts, placeholder/formula markers, and visibility settings.

## Required Inputs

- **CaseWare URL** with a letter document fragment: `https://{host}/{tenant}/e/eng/{engagementId}/index.jsp#/letter/{documentId}`
- **OAuth credentials** in `.env` (or browser cookies as fallback)

## Steps

### 1. Parse the URL
Extract host, tenant, engagementId, and documentId from the pasted URL using regex patterns.

### 2. Authenticate
Try OAuth first (per-region `CW_{PREFIX}_CLIENT_ID/SECRET`), fall back to `CW_COOKIES`.

### 3. Fetch Data
- `document/get` — fetch all documents to find the letter name
- `section/get` — fetch all sections for the letter document (filtered by document ID)

### 4. Build Section Tree
- Organize sections into parent-child hierarchy using `parent` field
- Sort by `order` field (fractional-index, lexicographic sort)
- Skip `settings` and `toc` type sections
- Handle `pagebreak` type as Word page breaks

### 5. Resolve Visibility IDs
- Collect all procedureId, checklistId, responseId values from visibility conditions
- Fetch documents, checklists, and procedures to resolve IDs to human-readable names
- Parse visibility direction: normallyVisible=false -> "Show when", true -> "Hide when"

### 6. Preprocess HTML
- Resolve `<span placeholder="...">` to `(( text ))`
- Resolve `<span formula="id">` to `[[ calculated_value ]]` using section.attachables
- Preserve HTML structure for the docx converter

### 7. Extract Guidance
- Found in `section.guidances.en` field
- Preprocess the guidance HTML same as body content

### 8. Generate JSON
Output structured JSON with sections array, each containing:
- type, title, level, html_content, guidance_html, visibility

### 9. Generate Word Document (Node.js)
- Parse HTML using cheerio DOM parser
- Map elements to docx-js Paragraph/TextRun/Table objects
- Render guidance as yellow-shaded table cells
- Render visibility inline as grey italic annotations
- Append visibility summary table at end of document
- Use proper Word list styles for bullets/numbers (never unicode bullets)

## Expected Output

A `.docx` file with:
- Document title as Heading 1
- Section titles as Heading 2/3/4 based on depth
- Body content preserving bold, italic, underline, lists, tables
- Placeholders in grey `(( ))` markers
- Dynamic text in grey `[[ ]]` markers
- Yellow guidance callout boxes
- Grey italic visibility annotations
- Summary visibility table on final page

## Tools Used

- `tools/letter_extract.py` — Python extraction (API calls, data processing)
- `tools/generate_docx.js` — Node.js Word generation (docx-js)

## API Quirks Discovered

- Letter document type is `"letter"` (confirmed)
- Letter sections are almost all `content` type (59/62 in test template)
- One `grouping` section acts as the root container
- `pagebreak` is a section type (not a setting) — must be rendered as Word page break
- Guidance field is `guidances.en` (same as checklist procedures)
- Visibility conditions follow the same structure as note documents
- Visibility condition types include `response`, `condition_group`, `organization_type`, and `consolidation`
- HTML content includes embedded base64 images (firm logos) — represented as placeholders in Word
- Formula spans (`<span formula="..." class="dynamic-text formula glossary">`) contain just a single space `" "` as visible content — the resolved value comes from `section.attachables`
- Placeholder spans have nested inner `<span>` elements (display text + caret) — use a DOM parser, not regex
- Ordered lists are split across multiple `<ol>` elements separated by `<p>` tags; the `start` attribute signals continuation
- Common list style: `list-style-type:lower-alpha` with `start="1"`
- Section `order` field uses fractional-index strings — in JavaScript, sort with `<`/`>` NOT `localeCompare` (which is locale-aware and reorders special characters)
