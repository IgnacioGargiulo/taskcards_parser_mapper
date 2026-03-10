# Task Card XML Parsers (IDMR -> TRAX API Payload)

Python utilities to parse IDMR Task Card XML and transform it into:

- Human-readable Excel exports (metadata/actionable/tables), and
- TRAX API payload XML format.

## What Is Included

- `taskcard_xml_parser.py`  
  Minimal Excel exporter that generates:
  - `TaskCard_Metadata`
  - `TaskCard_Actionable`
  - `TaskCard_Tables`

- `taskcard_xml_parser_full.py`  
  Extended parser version with additional sheets and options kept for backward compatibility.

- `taskcard_xml_to_api_payload.py`  
  Converter from source IDMR XML to TRAX-style payload XML.
  - Splits task content by section (`A.`, `B.`, `C.`, ...)
  - Creates one `<Task_Card_Item>` per section
  - Flattens hierarchy text (steps/substeps/warnings/notes)
  - Preserves table blocks as plain text with line breaks and tab-separated columns
  - Fills target-only fields with defaults when source mapping is unavailable

## Requirements

- Python 3.10+
- Dependencies from `requirements.txt`

Install:

```bash
pip install -r requirements.txt
```

## Usage

### 1) Export source XML to Excel (minimal)

```bash
python taskcard_xml_parser.py --xml 72071_TRAXExport1.xml --output 72071_TRAXExport1_extract.xlsx
```

### 2) Convert source XML to TRAX API payload XML

```bash
python taskcard_xml_to_api_payload.py --xml 72071_TRAXExport1.xml --output 72071_TRAXExport1_api_payload.xml
```

Optional defaults report:

```bash
python taskcard_xml_to_api_payload.py --xml 72071_TRAXExport1.xml --output 72071_TRAXExport1_api_payload.xml --report 72071_TRAXExport1_api_payload_report.json
```

## Input/Output Summary

- Input: IDMR Task Card XML (`Engineering_Task_Card/setTC/Task_Card_Element`)
- Output:
  - XLSX workbook (parser scripts)
  - XML payload conforming to expected TRAX structure (converter script)

## Notes

- Source XML may include `&nbsp;`; scripts normalize this before parsing.
- `Task_Card_Text` in the payload is emitted as text (not nested table XML), so table layout is represented with tabs/newlines.
- Some payload fields may be defaulted if they do not exist in source XML (see generated `--report` output).

## License

This project is licensed under the MIT License. See `LICENSE`.
