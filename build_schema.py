#!/usr/bin/env python3
"""
build_schema.py — reads structure.xlsx and emits form-schema.json.

Usage:
    python build_schema.py structure.xlsx [output.json]

The HTML/JS form engine fetches the output JSON at page load and renders
itself from it. This script never needs to change for ordinary revisions —
only structure.xlsx does. Edit the spreadsheet, re-run this, commit both.

All problems found are reported together (not just the first one), and
nothing is written unless the sheet is fully clean.
"""
import sys
import re
import json
from pathlib import Path

import openpyxl

SHOW_IF_EQ = re.compile(r"^\s*([A-Za-z0-9_]+)\s*(==|!=)\s*'([^']*)'\s*$")
SHOW_IF_SET = re.compile(r"^\s*([A-Za-z0-9_]+)\s+(all|not_all)\s*\[\s*([A-Za-z0-9_,\s]*)\s*\]\s*$")

VALID_TYPES = {"text", "email", "tel", "textarea", "date", "radio", "select", "checkbox_group", "info"}
VALID_REQUIRED = {"TRUE", "FALSE", "IF_VISIBLE"}
OPTION_TYPES = {"radio", "select", "checkbox_group"}


class SpecError(Exception):
    pass


def sheet_rows(ws):
    """Yield (row_number, {header: value}) for every non-empty data row."""
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    for i, row in enumerate(rows[1:], start=2):
        if all(v is None or str(v).strip() == "" for v in row):
            continue
        record = {headers[c]: row[c] for c in range(len(headers)) if headers[c]}
        yield i, record


def norm_bool(v, default="FALSE"):
    if v is None or str(v).strip() == "":
        return default
    return str(v).strip().upper()


def norm_list(v):
    """'all' | 'a,b,c' -> list, lowercased/stripped; 'all' passes through literally."""
    if v is None or str(v).strip() == "":
        return ["all"]
    v = str(v).strip()
    if v.lower() == "all":
        return ["all"]
    return [x.strip() for x in v.split(",") if x.strip()]


def parse_options(v):
    """'val:Label|val:Label' -> [{'value':..,'label':..}, ...]"""
    if v is None or str(v).strip() == "":
        return []
    out = []
    for part in str(v).split("|"):
        part = part.strip()
        if not part:
            continue
        if ":" not in part:
            raise SpecError(f"option '{part}' missing ':' between value and label")
        val, label = part.split(":", 1)
        out.append({"value": val.strip(), "label": label.strip()})
    return out


def build(xlsx_path: Path):
    errors = []
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    for required_sheet in ("Kinds", "Sections", "Fields"):
        if required_sheet not in wb.sheetnames:
            raise SpecError(f"missing required sheet '{required_sheet}'")

    # ---- Kinds ----
    kinds = []
    kind_ids = set()
    for rownum, r in sheet_rows(wb["Kinds"]):
        kid = str(r.get("kind_id", "")).strip()
        if not kid:
            errors.append(f"Kinds row {rownum}: empty kind_id")
            continue
        if kid in kind_ids:
            errors.append(f"Kinds row {rownum}: duplicate kind_id '{kid}'")
        kind_ids.add(kid)
        kinds.append({
            "id": kid,
            "display_name": str(r.get("display_name", "")).strip(),
            "description": str(r.get("description", "")).strip(),
        })

    def check_kind_refs(values, where):
        for v in values:
            if v != "all" and v not in kind_ids:
                errors.append(f"{where}: references unknown kind_id '{v}'")

    # ---- Sections ----
    sections = {}
    section_order = []
    for rownum, r in sheet_rows(wb["Sections"]):
        sid = str(r.get("section_id", "")).strip()
        if not sid:
            errors.append(f"Sections row {rownum}: empty section_id")
            continue
        if sid in sections:
            errors.append(f"Sections row {rownum}: duplicate section_id '{sid}'")
        applies = norm_list(r.get("applies_to_kinds"))
        check_kind_refs(applies, f"Sections row {rownum} ({sid})")
        try:
            order = float(r.get("order", 0) or 0)
        except (TypeError, ValueError):
            errors.append(f"Sections row {rownum} ({sid}): 'order' is not a number")
            order = 0
        sections[sid] = {
            "id": sid,
            "order": order,
            "title": str(r.get("title", "")).strip(),
            "applies_to_kinds": applies,
            "fields": [],
        }
        section_order.append(sid)

    # ---- Fields (pass 1: collect) ----
    fields = {}
    field_order = []
    field_rows_raw = list(sheet_rows(wb["Fields"]))
    for rownum, r in field_rows_raw:
        fid = str(r.get("field_id", "")).strip()
        if not fid:
            errors.append(f"Fields row {rownum}: empty field_id")
            continue
        if fid in fields:
            errors.append(f"Fields row {rownum}: duplicate field_id '{fid}'")
            continue

        ftype = str(r.get("type", "")).strip()
        if ftype not in VALID_TYPES:
            errors.append(f"Fields row {rownum} ({fid}): invalid type '{ftype}' "
                           f"(must be one of {sorted(VALID_TYPES)})")

        sid = str(r.get("section_id", "")).strip()
        if sid not in sections:
            errors.append(f"Fields row {rownum} ({fid}): unknown section_id '{sid}'")

        required = norm_bool(r.get("required"))
        if required not in VALID_REQUIRED:
            errors.append(f"Fields row {rownum} ({fid}): invalid required value '{required}' "
                           f"(must be TRUE, FALSE, or IF_VISIBLE)")

        applies = norm_list(r.get("applies_to_kinds"))
        check_kind_refs(applies, f"Fields row {rownum} ({fid})")

        try:
            options = parse_options(r.get("options"))
        except SpecError as e:
            errors.append(f"Fields row {rownum} ({fid}): {e}")
            options = []

        if ftype in OPTION_TYPES and not options:
            errors.append(f"Fields row {rownum} ({fid}): type '{ftype}' requires at least one option")
        if ftype not in OPTION_TYPES and options:
            errors.append(f"Fields row {rownum} ({fid}): type '{ftype}' should not have options")

        try:
            order = float(r.get("order", 0) or 0)
        except (TypeError, ValueError):
            errors.append(f"Fields row {rownum} ({fid}): 'order' is not a number")
            order = 0

        fields[fid] = {
            "id": fid,
            "section_id": sid,
            "order": order,
            "label": str(r.get("label", "")).strip(),
            "type": ftype,
            "options": options,
            "required": required,
            "applies_to_kinds": applies,
            "show_if_raw": (str(r.get("show_if")).strip() if r.get("show_if") else ""),
            "help_text": str(r.get("help_text", "")).strip() if r.get("help_text") else "",
            "_row": rownum,
        }
        field_order.append(fid)

    # ---- Fields (pass 2: show_if, now that every field_id is known) ----
    for fid, f in fields.items():
        raw = f.pop("show_if_raw")
        if not raw:
            f["show_if"] = None
            continue

        m_eq = SHOW_IF_EQ.match(raw)
        m_set = SHOW_IF_SET.match(raw)
        if m_eq:
            ref, op, val = m_eq.groups()
            if ref not in fields:
                errors.append(f"Fields row {f['_row']} ({fid}): show_if references unknown field '{ref}'")
            else:
                ref_field = fields[ref]
                if ref_field["type"] in OPTION_TYPES:
                    valid_vals = {o["value"] for o in ref_field["options"]}
                    if val not in valid_vals:
                        errors.append(
                            f"Fields row {f['_row']} ({fid}): show_if value '{val}' is not an option "
                            f"of '{ref}' (valid: {sorted(valid_vals)})"
                        )
            f["show_if"] = {"ref": ref, "op": op, "value": val}
        elif m_set:
            ref, op, raw_vals = m_set.groups()
            vals = [v.strip() for v in raw_vals.split(",") if v.strip()]
            if ref not in fields:
                errors.append(f"Fields row {f['_row']} ({fid}): show_if references unknown field '{ref}'")
            else:
                ref_field = fields[ref]
                if ref_field["type"] != "checkbox_group":
                    errors.append(
                        f"Fields row {f['_row']} ({fid}): show_if uses all/not_all on '{ref}', "
                        f"which is type '{ref_field['type']}', not checkbox_group"
                    )
                else:
                    valid_vals = {o["value"] for o in ref_field["options"]}
                    bad = [v for v in vals if v not in valid_vals]
                    if bad:
                        errors.append(
                            f"Fields row {f['_row']} ({fid}): show_if values {bad} not valid options "
                            f"of '{ref}' (valid: {sorted(valid_vals)})"
                        )
            f["show_if"] = {"ref": ref, "op": op, "values": vals}
        else:
            errors.append(f"Fields row {f['_row']} ({fid}): could not parse show_if '{raw}'")
            f["show_if"] = None

        if f["show_if"] and f["show_if"]["ref"] == fid:
            errors.append(f"Fields row {f['_row']} ({fid}): show_if cannot reference itself")

    # ---- PDF_Layout (optional overrides) ----
    pdf_overrides = {}
    if "PDF_Layout" in wb.sheetnames:
        for rownum, r in sheet_rows(wb["PDF_Layout"]):
            fid = str(r.get("field_id", "")).strip()
            if fid not in fields:
                errors.append(f"PDF_Layout row {rownum}: unknown field_id '{fid}'")
                continue
            pdf_overrides[fid] = {
                "page_break_before": norm_bool(r.get("page_break_before")) == "TRUE",
                "pdf_format": (str(r.get("pdf_format")).strip() if r.get("pdf_format") else None),
            }

    DEFAULT_PDF_FORMAT = {
        "textarea": "paragraph",
        "radio": "checklist",
        "checkbox_group": "checklist",
        "select": "table_row",
        "text": "table_row",
        "email": "table_row",
        "tel": "table_row",
        "date": "table_row",
        "info": "paragraph",
    }

    if errors:
        raise SpecError("\n".join(errors))

    # ---- Assemble ----
    for fid in field_order:
        f = fields[fid]
        f.pop("_row", None)
        override = pdf_overrides.get(fid, {})
        f["pdf"] = {
            "page_break_before": override.get("page_break_before", False),
            "pdf_format": override.get("pdf_format") or DEFAULT_PDF_FORMAT.get(f["type"], "table_row"),
        }
        sections[f["section_id"]]["fields"].append(f)

    for sid in sections:
        sections[sid]["fields"].sort(key=lambda f: f["order"])

    ordered_sections = [sections[sid] for sid in sorted(sections, key=lambda s: sections[s]["order"])]

    return {
        "version": 1,
        "kinds": kinds,
        "sections": ordered_sections,
    }


def main():
    if len(sys.argv) < 2:
        print("usage: python build_schema.py structure.xlsx [output.json]", file=sys.stderr)
        sys.exit(2)
    xlsx_path = Path(sys.argv[1])
    out_path = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("form-schema.json")

    try:
        schema = build(xlsx_path)
    except SpecError as e:
        print("Validation failed — nothing was written.\n", file=sys.stderr)
        print(str(e), file=sys.stderr)
        sys.exit(1)

    out_path.write_text(json.dumps(schema, indent=2))
    n_fields = sum(len(s["fields"]) for s in schema["sections"])
    print(f"OK: wrote {out_path}  "
          f"({len(schema['kinds'])} kinds, {len(schema['sections'])} sections, {n_fields} fields)")


if __name__ == "__main__":
    main()
