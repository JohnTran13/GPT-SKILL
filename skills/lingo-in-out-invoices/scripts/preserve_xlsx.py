"""Transplant explicitly planned cell contents from an authored XLSX into its base.

This is a fidelity packaging step after spreadsheet-tool authoring, not an XLSX
authoring engine. It leaves every other ZIP part byte-identical and refuses row,
sheet, merged-cell or shared/array-formula geometry changes. The plan is either
{"sheets": {"Sheet1": {"A1": {"value": 1}}}} or that inner sheets mapping.
Formula entries use {"formula": "=SUM(A1:A2)", "cache": 3}; cache is optional
when the authored workbook already contains the calculated value. Newly created
cells within existing rows require "style_from": "A2" in the entry.
"""
from __future__ import annotations

import argparse
import copy
from decimal import Decimal, InvalidOperation
import hashlib
import json
import math
from pathlib import Path
import re
import tempfile
from typing import Any
import xml.etree.ElementTree as ET
import zipfile

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
Q = lambda name: f"{{{NS}}}{name}"
ADDRESS = re.compile(r"([A-Z]{1,3})([1-9][0-9]*)\Z")
CONTENT_TAGS = {Q("f"), Q("v"), Q("is")}


def fail(message: str) -> None:
    raise ValueError(message)


def coordinate(address: str) -> tuple[int, int]:
    if not isinstance(address, str):
        fail(f"Cell address must be text: {address!r}")
    match = ADDRESS.fullmatch(address)
    if not match:
        fail(f"Invalid single-cell address: {address!r}")
    column = 0
    for letter in match[1]:
        column = column * 26 + ord(letter) - 64
    row = int(match[2])
    if column > 16384 or row > 1048576:
        fail(f"Address outside Excel bounds: {address}")
    return row, column


def in_range(address: str, reference: str) -> bool:
    parts = reference.replace("$", "").split(":")
    start = coordinate(parts[0]); end = coordinate(parts[-1])
    row, col = coordinate(address)
    return start[0] <= row <= end[0] and start[1] <= col <= end[1]


def xml(data: bytes) -> ET.Element:
    if b"<!DOCTYPE" in data.upper():
        fail("DOCTYPE declarations are unsupported")
    return ET.fromstring(data)


def same_value(left: Any, right: Any) -> bool:
    if isinstance(left, bool) or isinstance(right, bool):
        return type(left) is type(right) and left == right
    if isinstance(left, (int, float, Decimal)) and isinstance(right, (int, float, Decimal)):
        return Decimal(str(left)) == Decimal(str(right))
    return type(left) is type(right) and left == right


def valid_value(value: Any) -> bool:
    return value is None or isinstance(value, (str, bool, int)) or (
        isinstance(value, float) and math.isfinite(value)
    )


def plain_text(node: ET.Element | None) -> str:
    if node is None: return ""
    # Excel phonetic guides are annotations, not part of the displayed value.
    return "".join(
        (element.text or "") if element.tag == Q("t") else (element.findtext(Q("t")) or "")
        for element in node if element.tag in {Q("t"), Q("r")}
    )


class Book:
    def __init__(self, path: Path):
        self.path = path
        self.zip = zipfile.ZipFile(path)
        if self.zip.testzip() is not None:
            fail(f"Corrupt XLSX archive: {path}")
        names = self.zip.namelist()
        if len(names) != len(set(names)):
            fail(f"Duplicate ZIP entry names: {path}")
        if any(info.flag_bits & 1 for info in self.zip.infolist()):
            fail(f"Encrypted archive is unsupported: {path}")
        if "_xmlsignatures/origin.sigs" in names or any(n.startswith("_xmlsignatures/") for n in names):
            fail("Digitally signed workbook packages cannot be modified by this helper")
        try:
            workbook = xml(self.zip.read("xl/workbook.xml"))
            rels = xml(self.zip.read("xl/_rels/workbook.xml.rels"))
        except KeyError as exc:
            fail(f"Not a supported unencrypted XLSX package: {exc}")
        targets = {}
        for element in rels:
            if element.get("TargetMode") == "External":
                continue
            target = element.get("Target", "").replace("\\", "/")
            parts = []
            for part in (("xl/" + target) if not target.startswith("/") else target[1:]).split("/"):
                if part == "..":
                    if not parts: fail("Invalid workbook relationship path")
                    parts.pop()
                elif part and part != ".": parts.append(part)
            targets[element.get("Id")] = "/".join(parts)
        self.sheets = {}
        for sheet in workbook.findall(f"{Q('sheets')}/{Q('sheet')}"):
            target = targets.get(sheet.get(f"{{{REL}}}id"))
            if not target or target not in names:
                fail(f"Cannot resolve worksheet {sheet.get('name')!r}")
            self.sheets[sheet.get("name")] = target
        self.strings = []
        if "xl/sharedStrings.xml" in names:
            for item in xml(self.zip.read("xl/sharedStrings.xml")):
                self.strings.append(copy.deepcopy(item))

    def tree(self, sheet: str) -> ET.Element:
        if sheet not in self.sheets: fail(f"Missing sheet {sheet!r} in {self.path}")
        return xml(self.zip.read(self.sheets[sheet]))

    def value(self, cell: ET.Element | None) -> Any:
        if cell is None: return None
        typ = cell.get("t")
        value = cell.find(Q("v"))
        text = value.text if value is not None else None
        if typ == "e": fail(f"Authored formula error at {cell.get('r')}: {text}")
        if typ == "inlineStr": return plain_text(cell.find(Q("is")))
        if typ == "s":
            try: return plain_text(self.strings[int(text)])
            except (IndexError, TypeError, ValueError): fail("Invalid shared string index")
        if typ == "str": return text if text is not None else None
        if typ == "b": return text == "1" if text is not None else None
        if text is None: return None
        try: return Decimal(text)
        except InvalidOperation: fail(f"Unsupported cell value at {cell.get('r')}: {text!r}")

    def content(self, cell: ET.Element | None) -> tuple[str | None, list[ET.Element]]:
        if cell is None: return None, []
        typ = cell.get("t")
        children = [copy.deepcopy(child) for child in cell if child.tag in CONTENT_TAGS]
        if typ == "s":
            value = cell.findtext(Q("v"))
            shared = copy.deepcopy(self.strings[int(value)])
            shared.tag = Q("is")
            children = [child for child in children if child.tag != Q("v")]
            children.append(shared)
            typ = "inlineStr"
        elif typ == "str" and cell.find(Q("f")) is None:
            text = cell.findtext(Q("v"), default="")
            inline = ET.Element(Q("is")); ET.SubElement(inline, Q("t")).text = text
            children = [inline]; typ = "inlineStr"
        return typ, children

    def close(self): self.zip.close()


def read_plan(path: Path) -> dict:
    def unique_keys(pairs):
        result = {}
        for key, value in pairs:
            if key in result: fail(f"Duplicate JSON key in plan: {key!r}")
            result[key] = value
        return result
    data = json.loads(path.read_text(encoding="utf-8-sig"), object_pairs_hook=unique_keys)
    if not isinstance(data, dict): fail("Plan must be an object")
    if "sheets" in data:
        if set(data) != {"sheets"}: fail("Plan supports only the top-level sheets key")
        data = data["sheets"]
    if not isinstance(data, dict) or not any(data.values()): fail("Plan has no edits")
    for name, edits in data.items():
        if not isinstance(name, str) or not isinstance(edits, dict): fail("Invalid sheet edits mapping")
        for address, item in edits.items():
            coordinate(address)
            if not isinstance(item, dict) or not set(item) <= {"value", "formula", "cache", "style_from"}:
                fail(f"Invalid plan entry {name}!{address}")
            if ("value" in item) == ("formula" in item): fail(f"Use exactly one value or formula at {address}")
            if "value" in item and (not valid_value(item["value"]) or "cache" in item): fail(f"Invalid value entry {address}")
            if "formula" in item and (not isinstance(item["formula"], str) or not item["formula"].startswith("=") or len(item["formula"]) < 2): fail(f"Invalid formula at {address}")
            if "cache" in item and not valid_value(item["cache"]): fail(f"Invalid formula cache at {address}")
            if "style_from" in item: coordinate(item["style_from"])
    return data


def protect_formulas_and_merges(root: ET.Element, addresses: set[str]) -> None:
    for cell in root.iter(Q("c")):
        formula = cell.find(Q("f"))
        if formula is not None and formula.get("t") in {"shared", "array", "dataTable"}:
            reference = formula.get("ref")
            if cell.get("r") in addresses or (reference and any(in_range(a, reference) for a in addresses)):
                fail(f"Edit intersects unsupported {formula.get('t')} formula at {cell.get('r')}")
    for merged in root.iter(Q("mergeCell")):
        reference = merged.get("ref", "")
        for address in addresses:
            if in_range(address, reference) and address != reference.split(":")[0].replace("$", ""):
                fail(f"Edit is not the top-left cell of merged range {reference}")


def cached_content(children: list[ET.Element], value: Any) -> tuple[str | None, list[ET.Element]]:
    result = [c for c in children if c.tag not in {Q("v"), Q("is")}]
    if value is None: fail("Formula cache is missing; recalculate authored workbook before packaging")
    typ = "b" if isinstance(value, bool) else "str" if isinstance(value, str) else None
    node = ET.Element(Q("v")); node.text = "1" if value is True else "0" if value is False else str(value)
    result.append(node)
    return typ, result


def build(base: Book, authored: Book, plan: dict) -> dict[str, bytes]:
    updates = {}
    for name, edits in plan.items():
        if not edits: continue
        root = base.tree(name); other = authored.tree(name)
        data = root.find(Q("sheetData"))
        if data is None: fail(f"No sheetData in {name}")
        rows = {int(r.get("r")): r for r in data}
        cells = {c.get("r"): c for r in data for c in r if c.tag == Q("c")}
        source = {c.get("r"): c for c in other.iter(Q("c"))}
        dimension = root.find(Q("dimension"))
        protect_formulas_and_merges(root, set(edits))
        protect_formulas_and_merges(other, set(edits))
        for address, item in edits.items():
            row_number, _ = coordinate(address)
            if dimension is not None and not in_range(address, dimension.get("ref", "A1")):
                fail(f"Cell {name}!{address} expands the existing used range; extend natively before packaging")
            if row_number not in rows: fail(f"Missing existing row {name}!{row_number}; extend natively before packaging")
            incoming = source.get(address)
            formula = incoming.find(Q("f")) if incoming is not None else None
            incoming_value = authored.value(incoming)
            typ, children = authored.content(incoming)
            if "value" in item:
                if formula is not None or not same_value(incoming_value, item["value"]):
                    fail(f"Plan/authored value mismatch at {name}!{address}: {incoming_value!r} != {item['value']!r}")
            else:
                if formula is None or formula.text != item["formula"][1:]:
                    fail(f"Plan/authored formula mismatch at {name}!{address}")
                if "cache" in item:
                    if incoming_value is not None and not same_value(incoming_value, item["cache"]):
                        fail(f"Plan/authored formula cache mismatch at {name}!{address}")
                    typ, children = cached_content(children, item["cache"])
                elif incoming_value is None:
                    fail(f"No calculated formula cache at {name}!{address}")
            target = cells.get(address)
            if target is None:
                style_cell = cells.get(item.get("style_from"))
                if style_cell is None: fail(f"New cell {name}!{address} requires an existing style_from cell")
                target = ET.Element(Q("c"), {k: v for k, v in style_cell.attrib.items() if k not in {"r", "t", "cm", "vm"}})
                target.set("r", address); rows[row_number].append(target); cells[address] = target
            for child in list(target):
                if child.tag in CONTENT_TAGS: target.remove(child)
            target.attrib.pop("t", None)
            if typ is not None: target.set("t", typ)
            # Schema order places formula/value/inline text before cell extension data.
            for index, child in enumerate(children): target.insert(index, child)
            rows[row_number][:] = sorted(rows[row_number], key=lambda c: coordinate(c.get("r"))[1])
        serialized = ET.tostring(data, encoding="unicode")
        raw = base.zip.read(base.sheets[name]).decode("utf-8-sig")
        expression = r"<(?:[A-Za-z_][\w.-]*:)?sheetData\b[^>]*>.*?</(?:[A-Za-z_][\w.-]*:)?sheetData\s*>"
        updated, count = re.subn(expression, lambda match: serialized, raw, flags=re.S)
        if count != 1: fail(f"Cannot isolate sheetData in {name}")
        updates[base.sheets[name]] = updated.encode("utf-8")
    return updates


def verify(base: Book, final: Book, plan: dict, updates: dict) -> None:
    if base.zip.namelist() != final.zip.namelist(): fail("Package part inventory changed")
    for name in base.zip.namelist():
        if name not in updates and base.zip.read(name) != final.zip.read(name): fail(f"Untouched package part changed: {name}")
        if name.endswith(".xml") or name.endswith(".rels"): xml(final.zip.read(name))
    for name, edits in plan.items():
        if not edits: continue
        before = base.tree(name); after = final.tree(name)
        old_cells = {c.get("r"): c for c in before.iter(Q("c"))}
        new_cells = {c.get("r"): c for c in after.iter(Q("c"))}
        for address, cell in old_cells.items():
            changed = new_cells.get(address)
            if changed is None: fail(f"Existing cell disappeared: {name}!{address}")
            if address not in edits:
                if ET.tostring(cell) != ET.tostring(changed): fail(f"Untargeted cell changed: {name}!{address}")
            elif {k:v for k,v in cell.attrib.items() if k != "t"} != {k:v for k,v in changed.attrib.items() if k != "t"}:
                fail(f"Target cell formatting/metadata changed: {name}!{address}")
        for address, entry in edits.items():
            cell = new_cells[address]; value = final.value(cell)
            if "value" in entry and not same_value(value, entry["value"]): fail(f"Saved value mismatch: {name}!{address}")
            if "formula" in entry:
                if cell.findtext(Q("f")) != entry["formula"][1:]: fail(f"Saved formula mismatch: {name}!{address}")
                if "cache" in entry and not same_value(value, entry["cache"]): fail(f"Saved cache mismatch: {name}!{address}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    for argument in ("base", "authored", "plan", "output"): parser.add_argument(f"--{argument}", required=True, type=Path)
    parser.add_argument("--overwrite", action="store_true", help="Replace an existing result; input files are always protected")
    args = parser.parse_args()
    base_path, authored_path, output = args.base.resolve(), args.authored.resolve(), args.output.resolve()
    if output in {base_path, authored_path, args.plan.resolve()}: fail("Output must be distinct from all input files")
    if output.exists() and not args.overwrite: fail("Output exists; choose a new path or explicitly use --overwrite")
    if not output.parent.is_dir(): fail("Output parent directory must already exist")
    hashes = {p: hashlib.sha256(p.read_bytes()).hexdigest() for p in (base_path, authored_path, args.plan.resolve())}
    plan = read_plan(args.plan)
    base, authored = Book(base_path), Book(authored_path)
    temp_path = None
    try:
        updates = build(base, authored, plan)
        with tempfile.NamedTemporaryFile(prefix=".preserve-xlsx-", suffix=".xlsx", dir=output.parent, delete=False) as handle:
            temp_path = Path(handle.name)
        with zipfile.ZipFile(temp_path, "w") as dest:
            for info in base.zip.infolist(): dest.writestr(copy.copy(info), updates.get(info.filename, base.zip.read(info.filename)))
        final = Book(temp_path)
        try: verify(base, final, plan, updates)
        finally: final.close()
        for path, digest in hashes.items():
            if hashlib.sha256(path.read_bytes()).hexdigest() != digest: fail(f"Input changed during packaging: {path}")
        temp_path.replace(output); temp_path = None
        print(json.dumps({"output": str(output), "cells_updated": sum(len(x) for x in plan.values()), "sheets_updated": len(updates), "source_files_unchanged": True}, ensure_ascii=False))
    finally:
        base.close(); authored.close()
        if temp_path is not None: temp_path.unlink(missing_ok=True)


if __name__ == "__main__":
    try: main()
    except (ValueError, OSError, zipfile.BadZipFile, ET.ParseError) as exc:
        raise SystemExit(f"ERROR: {exc}")
