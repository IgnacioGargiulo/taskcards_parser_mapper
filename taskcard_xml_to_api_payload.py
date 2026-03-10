#!/usr/bin/env python3
"""
Convert TRAX task card XML into IDMR API payload XML format.

Key behavior:
- Preserves known top-level/task-item fields where available.
- Flattens hierarchical Task_Card_Text content into a deterministic multiline string.
- Includes table content in Task_Card_Text using tab-separated columns.
- Populates missing target-only fields with safe defaults.

Usage:
  python taskcard_xml_to_api_payload.py --xml 72071_TRAXExport1.xml
  python taskcard_xml_to_api_payload.py --xml 72071_TRAXExport1.xml --output payload.xml --report defaults.json
"""

import argparse
import json
import re
import xml.etree.ElementTree as ET
from collections import OrderedDict
from collections import defaultdict
from pathlib import Path
from xml.dom import minidom


CODE_RE = re.compile(r"\b(\d{2}-\d{2}-\d{2}-\d{3}-\d{3})\b")
TAG_TYPES = {"A", "W", "C", "N", "f"}


def normalize_ws(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def safe_xml_text(xml_path: Path) -> str:
    return xml_path.read_text(encoding="utf-8", errors="replace").replace("&nbsp;", " ")


def parse_xml(xml_path: Path) -> ET.Element:
    return ET.fromstring(safe_xml_text(xml_path))


def text_of(node: ET.Element | None) -> str:
    if node is None:
        return ""
    return normalize_ws("".join(node.itertext()))


def child_text(node: ET.Element | None, tag: str, default: str = "") -> str:
    if node is None:
        return default
    child = node.find(tag)
    value = text_of(child)
    return value if value else default


def bool_text(value: str) -> str:
    return "true" if str(value).strip().lower() == "true" else "false"


def extract_code(text: str) -> str:
    m = CODE_RE.search(text or "")
    return m.group(1) if m else ""


def split_aircraft_and_code(text: str) -> tuple[str, str]:
    clean = normalize_ws(text)
    if not clean:
        return "", ""
    if "|" in clean:
        left, right = [normalize_ws(x) for x in clean.split("|", 1)]
        code = extract_code(right) or extract_code(left)
        if code:
            left_no_code = normalize_ws(CODE_RE.sub(" ", left))
            aircraft = left_no_code if left_no_code and "task " not in left_no_code.lower() else ""
            return aircraft, code
        return left, ""
    code = extract_code(clean)
    if code:
        aircraft = normalize_ws(CODE_RE.sub(" ", clean))
        if "task " in aircraft.lower():
            aircraft = ""
        return aircraft, code
    return clean, ""


def parse_table_rows(table_node: ET.Element) -> list[list[str]]:
    rows: list[list[str]] = []
    for tr in table_node.findall(".//tr"):
        cells = [normalize_ws("".join(td.itertext())) for td in tr.findall("./th") + tr.findall("./td")]
        if any(cells):
            rows.append(cells)
    return rows


def flatten_task_text_by_section(task_text: ET.Element | None) -> list[dict]:
    if task_text is None:
        return []

    sections: "OrderedDict[str, dict]" = OrderedDict()
    preamble: list[str] = []

    def ensure_section(sec_no: str) -> dict:
        if sec_no not in sections:
            sections[sec_no] = {"section_no": sec_no, "section_title": "", "lines": []}
        return sections[sec_no]

    last_subtask_by_section: dict[str, str] = {}
    last_hierarchy_prefix_by_section: dict[str, str] = {}

    def add_content_line(
        text: str,
        current_section_no: str,
        step: str,
        substep: str,
        subtask: str,
        tag_type: str,
        force_prefix: bool = False,
        preserve_tabs: bool = False,
    ) -> None:
        text = (text or "").strip() if preserve_tabs else normalize_ws(text)
        if not text:
            return
        # Show parent step once, then substeps as (a)/(b)/... without repeating the step number.
        hierarchy_prefix = substep if substep else step

        prefix = []
        if hierarchy_prefix:
            last_prefix = last_hierarchy_prefix_by_section.get(current_section_no, "")
            if force_prefix or hierarchy_prefix != last_prefix:
                prefix.append(hierarchy_prefix)
                last_hierarchy_prefix_by_section[current_section_no] = hierarchy_prefix
        # Show subtask code only once when it changes inside a section.
        if subtask:
            last = last_subtask_by_section.get(current_section_no, "")
            if subtask != last:
                prefix.append(subtask)
                last_subtask_by_section[current_section_no] = subtask
        if tag_type == "W":
            prefix.append("WARNING:")
        elif tag_type == "C":
            prefix.append("CAUTION:")
        elif tag_type == "N":
            prefix.append("NOTE:")
        elif tag_type == "f":
            prefix.append("FLAG:")

        line = f"{' '.join(prefix)} {text}".strip() if prefix else text
        if current_section_no:
            ensure_section(current_section_no)["lines"].append(line)
        else:
            preamble.append(line)

    section_no = ""
    section_title = ""
    step_no = ""
    substep_no = ""
    order_no = ""
    current_t = ""
    current_aircraft = ""
    current_subtask = ""
    item_task_code = ""
    item_description = ""
    expecting_section_title = False

    for child in list(task_text):
        tag = child.tag

        if tag == "n":
            n_val = normalize_ws(child.text or "")
            if n_val:
                if re.match(r"^[A-Z]\.$", n_val):
                    section_no = n_val
                    section_title = ""
                    step_no = ""
                    substep_no = ""
                    expecting_section_title = True
                    ensure_section(section_no)
                elif re.match(r"^\(\d+\)$", n_val):
                    step_no = n_val
                    substep_no = ""
                elif re.match(r"^\([a-z]\)$", n_val):
                    substep_no = n_val
                elif re.match(r"^\d+$", n_val):
                    order_no = n_val
            continue

        if tag == "t":
            current_t = normalize_ws(child.text or "")
            tail = normalize_ws(child.tail or "")
            if not tail:
                continue

            # Task header line
            if tail.upper().startswith("TASK "):
                code = extract_code(tail)
                if code:
                    item_task_code = code
                add_content_line(f"TASK {tail}", section_no, "", "", "", "")
                continue

            if current_t == "A":
                aircraft, code = split_aircraft_and_code(tail)
                if aircraft:
                    current_aircraft = aircraft
                if code:
                    current_subtask = code
                # Do not emit effectivity/order lines into payload text.
                continue

            if current_t in TAG_TYPES:
                add_content_line(tail, section_no, step_no, substep_no, current_subtask, current_t, force_prefix=True)
                continue

            if expecting_section_title and not section_title:
                section_title = tail
                expecting_section_title = False
                sec = ensure_section(section_no)
                sec["section_title"] = section_title
                sec["lines"].insert(0, f"{section_no} {section_title}".strip())
                continue

            add_content_line(tail, section_no, step_no, substep_no, current_subtask, "")
            continue

        if tag == "p":
            text = normalize_ws("".join(child.itertext()))
            if not text:
                continue

            if expecting_section_title and not section_title:
                section_title = text
                expecting_section_title = False
                sec = ensure_section(section_no)
                sec["section_title"] = section_title
                sec["lines"].insert(0, f"{section_no} {section_title}".strip())
                continue

            if text.upper().startswith("TASK "):
                code = extract_code(text)
                if code:
                    item_task_code = code
                add_content_line(f"TASK {text}", section_no, "", "", "", "")
                continue

            if current_t == "A":
                aircraft, code = split_aircraft_and_code(text)
                if aircraft:
                    current_aircraft = aircraft
                if code:
                    current_subtask = code
                # Do not emit effectivity/order lines into payload text.
                continue

            if current_t in TAG_TYPES:
                add_content_line(text, section_no, step_no, substep_no, current_subtask, current_t, force_prefix=True)
                continue

            if item_task_code and not item_description and not section_no:
                item_description = text
                add_content_line(f"ITEM_DESCRIPTION: {item_description}", section_no, "", "", "", "")
                continue

            add_content_line(text, section_no, step_no, substep_no, current_subtask, "")
            continue

        if tag == "table":
            rows = parse_table_rows(child)
            if rows:
                # Plain table text: header + data rows, tab-separated.
                add_content_line(
                    "\t".join(rows[0]),
                    section_no,
                    "",
                    "",
                    "",
                    "",
                    preserve_tabs=True,
                )
                for row in rows[1:]:
                    add_content_line("\t".join(row), section_no, "", "", "", "", preserve_tabs=True)
            continue

        extra = normalize_ws("".join(child.itertext()))
        if extra:
            add_content_line(extra, section_no, step_no, substep_no, current_subtask, "")

    # If no sections were found, emit preamble as a single fallback section.
    if preamble and not sections:
        sections[""] = {"section_no": "", "section_title": "", "lines": preamble}

    out: list[dict] = []
    for sec in sections.values():
        text = "\n".join(sec["lines"]).strip()
        if text:
            out.append({"section_no": sec["section_no"], "section_title": sec["section_title"], "text": text})
    return out


def ensure_text(parent: ET.Element, tag: str, value: str) -> ET.Element:
    child = ET.SubElement(parent, tag)
    child.text = value
    return child


def clone_list(source_parent: ET.Element | None, source_item_tag: str, target_parent: ET.Element, target_item_tag: str, field_tags: list[str], defaults_log: dict[str, int]) -> None:
    if source_parent is None:
        return
    for src_item in source_parent.findall(source_item_tag):
        tgt_item = ET.SubElement(target_parent, target_item_tag)
        for tag in field_tags:
            value = child_text(src_item, tag, "")
            if value == "":
                defaults_log[f"{target_item_tag}.{tag}"] += 1
            ensure_text(tgt_item, tag, value)


def build_payload(source_root: ET.Element) -> tuple[ET.Element, dict[str, int]]:
    defaults_log: dict[str, int] = defaultdict(int)
    src_tc = source_root.find("./setTC/Task_Card_Element")
    if src_tc is None:
        raise RuntimeError("Task_Card_Element not found in source XML.")

    root = ET.Element("Engineering_Task_Card")
    set_tc = ET.SubElement(root, "setTC")
    tc = ET.SubElement(set_tc, "Task_Card_Element")

    # Core scalar fields
    scalar_map = [
        "Task_Card_Name",
        "Status",
        "Revision",
        "Revised_By",
        "Editor_Used",
        "TC_Description",
        "Type",
        "Category",
        "Chapter",
        "Section",
        "Paragraph",
        "Area",
        "Etops_Flag",
        "Corrosion_Flag",
        "SSID_Flag",
        "NDT_Flag",
        "ET_Flag",
        "MT_Flag",
        "PT_Flag",
        "UT_Flag",
    ]

    for tag in scalar_map:
        value = child_text(src_tc, tag, "")
        if value == "":
            defaults_log[f"Task_Card_Element.{tag}"] += 1
        if tag.endswith("_Flag"):
            value = bool_text(value)
        ensure_text(tc, tag, value)

    # Target-only scalars
    for tag, default in [
        ("Revised_Date", ""),
        ("Repair_Alteration_Identifier", ""),
        ("Repair_Alteration_Classification", ""),
        ("MPD_Reference", ""),
        ("RII_Flag", "false"),
    ]:
        ensure_text(tc, tag, default)
        defaults_log[f"Task_Card_Element.{tag}"] += 1

    # Task items
    tgt_items = ET.SubElement(tc, "Task_Card_Items")
    src_items = src_tc.find("./Task_Card_Items")
    emitted_item_no = 0
    if src_items is not None:
        for src_item in src_items.findall("./Task_Card_Item"):
            section_items = flatten_task_text_by_section(src_item.find("./Task_Card_Text"))
            # Fallback for unexpected structure: keep one payload item.
            if not section_items:
                section_items = [{"section_no": "", "section_title": "", "text": text_of(src_item.find("./Task_Card_Text"))}]

            for sec in section_items:
                emitted_item_no += 1
                tgt_item = ET.SubElement(tgt_items, "Task_Card_Item")
                ensure_text(tgt_item, "Item_Number", str(emitted_item_no))
                ensure_text(tgt_item, "Task_Card_Text", sec["text"])

                # Existing booleans from source
                for tag in [
                    "signoff_required_mechanic",
                    "signoff_required_inspector",
                    "signoff_required_duplicate_inspector",
                ]:
                    ensure_text(tgt_item, tag, bool_text(child_text(src_item, tag, "false")))

                # Target-only item attributes
                for tag, default in [
                    ("skill_Mechanic", ""),
                    ("man_Require_Mechanic", "0"),
                    ("man_Hours_Mechanic", "0.0"),
                    ("skill_Inspector", child_text(src_tc, "Skill_Inspector", "")),
                    ("man_Require_Inspector", "0"),
                    ("man_Hours_Inspector", "0.0"),
                    ("skill_Duplicate_Inspector", ""),
                ]:
                    ensure_text(tgt_item, tag, default)
                    if default in {"", "0", "0.0"}:
                        defaults_log[f"Task_Card_Item.{tag}"] += 1

    # Material requirements (target-only; include empty container by default)
    materials = ET.SubElement(tc, "Task_Card_Material_Requirements")
    src_mats = src_tc.find("./Task_Card_Material_Requirements")
    if src_mats is not None and src_mats.findall("./Task_Card_Material_Requirement"):
        clone_list(
            src_mats,
            "./Task_Card_Material_Requirement",
            materials,
            "Task_Card_Material_Requirement",
            ["PN", "Qty", "Requirement_Type", "Spare_Or_Tool"],
            defaults_log,
        )
    else:
        mat = ET.SubElement(materials, "Task_Card_Material_Requirement")
        for tag, default in [("PN", ""), ("Qty", "0"), ("Requirement_Type", ""), ("Spare_Or_Tool", "")]:
            ensure_text(mat, tag, default)
            defaults_log[f"Task_Card_Material_Requirement.{tag}"] += 1

    # Zones, Panels, Manual refs
    zones = ET.SubElement(tc, "Task_Card_Zones")
    clone_list(src_tc.find("./Task_Card_Zones"), "./Task_Card_Zone", zones, "Task_Card_Zone", ["Zone", "Aircraft_Type", "Aircraft_Series"], defaults_log)

    panels = ET.SubElement(tc, "Task_Card_Panels")
    clone_list(src_tc.find("./Task_Card_Panels"), "./Task_Card_Panel", panels, "Task_Card_Panel", ["Panel", "Aircraft_Type", "Aircraft_Series"], defaults_log)

    refs = ET.SubElement(tc, "Task_Card_Manual_References")
    clone_list(src_tc.find("./Task_Card_Manual_References"), "./Task_Card_Manual_Reference", refs, "Task_Card_Manual_Reference", ["Manual", "Reference", "Description"], defaults_log)

    # Linked files (target-only default)
    linked_files = ET.SubElement(tc, "Linked_Files")
    linked = ET.SubElement(linked_files, "Linked_File")
    ensure_text(linked, "File_Location", "")
    defaults_log["Linked_File.File_Location"] += 1

    # Survey questions - pass through if present, else default sample shape
    surveys = ET.SubElement(tc, "Survey_Questions")
    src_surveys = src_tc.find("./Survey_Questions")
    if src_surveys is not None and src_surveys.findall("./Survey_Question"):
        for src_q in src_surveys.findall("./Survey_Question"):
            surveys.append(src_q)
    else:
        sq = ET.SubElement(surveys, "Survey_Question")
        ensure_text(sq, "Question", "")
        ensure_text(sq, "Response_Mandatory", "false")
        defaults_log["Survey_Question.Question"] += 1

    # Effectiveness
    eff = ET.SubElement(tc, "Task_Card_Effectiveness")
    src_eff = src_tc.find("./Task_Card_Effectiveness")
    src_eff_items = src_eff.findall("./Task_Card_Effectivity") if src_eff is not None else []
    if src_eff_items:
        for src_eff_item in src_eff_items:
            tgt_eff = ET.SubElement(eff, "Task_Card_Effectivity")
            manual_codes = src_eff_item.findall("./Manual_Code")
            if manual_codes:
                for mc in manual_codes:
                    ensure_text(tgt_eff, "Manual_Code", normalize_ws(mc.text or ""))
            else:
                ensure_text(tgt_eff, "Manual_Code", "")
                defaults_log["Task_Card_Effectivity.Manual_Code"] += 1
    else:
        tgt_eff = ET.SubElement(eff, "Task_Card_Effectivity")
        ensure_text(tgt_eff, "Manual_Code", "")
        defaults_log["Task_Card_Effectivity.Manual_Code"] += 1

    return root, dict(defaults_log)


def write_pretty_xml(root: ET.Element, output_path: Path) -> None:
    raw = ET.tostring(root, encoding="utf-8")
    pretty = minidom.parseString(raw).toprettyxml(indent="  ", newl="\n")
    compact = "\n".join(line for line in pretty.splitlines() if line.strip()) + "\n"
    output_path.write_text(compact, encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert TRAX XML to API payload XML.")
    parser.add_argument("--xml", required=True, type=Path, help="Input TRAX XML path")
    parser.add_argument("--output", type=Path, help="Output payload XML path")
    parser.add_argument("--report", type=Path, help="Optional JSON report path for defaulted fields")
    args = parser.parse_args()

    output_path = args.output or args.xml.with_name(f"{args.xml.stem}_api_payload.xml")
    root = parse_xml(args.xml)
    payload_root, defaults_log = build_payload(root)
    write_pretty_xml(payload_root, output_path)

    if args.report:
        report = {
            "input_xml": str(args.xml),
            "output_xml": str(output_path),
            "defaulted_field_counts": defaults_log,
        }
        args.report.write_text(json.dumps(report, indent=2), encoding="utf-8")
        print(f"Wrote {output_path} | report: {args.report}")
    else:
        print(f"Wrote {output_path}")


if __name__ == "__main__":
    main()
