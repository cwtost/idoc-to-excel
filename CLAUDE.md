# CLAUDE.md

IDoc-to-Excel converter application documentation for Claude Code.

## Project Overview

Flask-based SAP IDoc (txt/xml) parser. Users upload IDoc files in TXT or XML format and convert them to Excel documentation with segment breakdown, field descriptions, and data preview. Interface is in Polish.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run development server (http://localhost:5004)
python app.py

# Command-line usage
python idoc_parser.py <input.txt> [output.xlsx]
```

No test suite or linter is configured.

## Architecture

**Single-feature web app** — one backend file, one parser, one template.

- `app.py` — Flask app with two routes:
  - `GET /` — renders `index.html` upload interface
  - `POST /convert` — accepts multipart form with `idoc_file`; parses IDoc segments, generates Excel documentation, returns JSON with `{success, filename, xlsx_b64, summary[], rows[]}`

- `templates/index.html` — SPA frontend; drag-and-drop upload zone, tabs for summary vs. fields detail, calls `/convert` via `fetch`, renders two tables with segment data

- `idoc_parser.py` — Core parsing logic:
  - `parse_idoc(filepath)` — reads IDoc file, splits by segment terminator, extracts segment tags
  - `extract_fields_idoc(tag, raw_data)` — maps segment tag to field definitions, extracts values
  - `build_excel(segments, output_path)` — creates 2-sheet Excel workbook using openpyxl
  - `build_preview_data(segments)` — returns same data as Excel but as JSON for web preview

## Supported IDoc Types

- ORDERS (Purchase Orders)
- INVOIC (Invoices)
- DESADV (Despatch Advice)
- RECADV (Receiving Advice)

**Key segments** defined:
- UNB – Interchange header
- UNH – Message header
- BGM – Beginning of message
- DTM – Date/time
- NAD – Party (Name and Address)
- LIN – Line item
- QTY – Quantity
- MOA – Monetary amount
- UNT/UNZ – Message/interchange trailers

## File Format

IDoc flat files (`.txt`, `.edi`):
- Segment terminator: typically `~` or `'` (auto-detected)
- Data elements separated by `:` or `+` (composite)
- Character set: UTF-8 or Latin-1
- File size limit: 16 MB (app.py hardcoded)

## Excel Output

**Sheet 1:** "IDoc - Pola i wartości"
- Full segment breakdown: segment, occurrence count, field name, length, description, value

**Sheet 2:** "Podsumowanie"
- Summary table: segment, count, description, order in message

## Extending the Parser

1. **Add new segment definitions** to `SEGMENT_FIELDS` dict in `idoc_parser.py`
2. **Add Polish descriptions** to `SEG_DESC` dict
3. Test with sample IDoc file
