#!/usr/bin/env python3
"""Read an IN tree without modifying sources; emit normalized invoices and an audit.

Amounts are decimal strings, not binary floats. XML is authoritative when a PDF
duplicates it. Unsupported PDF layouts remain readable in documents[].text and
must be reviewed by the caller. ZIP members are read in memory, never extracted.
"""
from __future__ import annotations

import argparse
from collections import Counter
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
import hashlib
import io
import json
from pathlib import Path
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None


SCHEMA_VERSION = 1
MAX_ZIP_DEPTH = 8
MAX_MEMBER_BYTES = 64 * 1024 * 1024


def text(node, path):
    value = node.findtext(path)
    return value.strip() if value and value.strip() else None


def decimal_string(value):
    if value is None or str(value).strip() == "":
        return None
    try:
        number = Decimal(str(value).strip())
        if not number.is_finite():
            raise ValueError("Non-finite amount")
        return format(number, "f")
    except (InvalidOperation, ValueError):
        raise ValueError(f"Invalid decimal amount: {value!r}")


def norm_number(value):
    value = str(value).strip()
    return str(int(value)) if value.isdigit() else value


def invoice_key(record):
    issuer = record.get("seller_tax_id")
    series = record.get("series")
    number = record.get("number")
    if issuer and series and number is not None:
        return f"{issuer}|{series}|{norm_number(number)}"
    return None


def direction_for(relative):
    group = Path(relative.split("!", 1)[0]).parts[0].upper()
    if group.startswith("01_INCOME"):
        return "purchase"
    if group.startswith("02_OUT COME") or group.startswith("02_OUTCOME"):
        return "sale"
    return "unknown"


def parse_xml(data, source):
    root = ET.fromstring(data)
    for element in root.iter():
        element.tag = element.tag.split("}")[-1]
    payload = root if root.tag == "DLHDon" else root.find(".//DLHDon")
    if payload is None:
        return None
    common = payload.find("TTChung")
    body = payload.find("NDHDon")
    if common is None or body is None or not text(common, "SHDon"):
        return None
    model = text(common, "KHMSHDon") or ""
    serial = text(common, "KHHDon") or ""
    record = {
        "kind": "invoice", "direction": direction_for(source),
        "date": text(common, "NLap"), "series": model + serial,
        "number": text(common, "SHDon"),
        "seller_name": text(body, "NBan/Ten"),
        "seller_tax_id": text(body, "NBan/MST"),
        "buyer_name": text(body, "NMua/Ten"),
        "buyer_tax_id": text(body, "NMua/MST"),
        "currency": text(common, "DVTTe"),
        "base": decimal_string(text(body, "TToan/TgTCThue")),
        "vat": decimal_string(text(body, "TToan/TgTThue")),
        "payable": decimal_string(text(body, "TToan/TgTTTBSo")),
        "rate_groups": [], "items": [], "sources": [source],
        "preferred_source": source, "source_format": "xml", "warnings": [],
    }
    for group in body.findall("TToan/THTTLTSuat/LTSuat"):
        record["rate_groups"].append({
            "rate": text(group, "TSuat"),
            "base": decimal_string(text(group, "ThTien")),
            "vat": decimal_string(text(group, "TThue")),
        })
    for line in body.findall("DSHHDVu/HHDVu"):
        record["items"].append({
            "name": text(line, "THHDVu"), "nature": text(line, "TChat"),
            "unit": text(line, "DVTinh"),
            "quantity": decimal_string(text(line, "SLuong")),
            "unit_price": decimal_string(text(line, "DGia")),
            "amount": decimal_string(text(line, "ThTien")),
            "rate": text(line, "TSuat"),
        })
    if record["date"]:
        try:
            record["date"] = datetime.strptime(record["date"][:10], "%Y-%m-%d").date().isoformat()
        except ValueError:
            record["warnings"].append("Invalid invoice date; retain raw date for review")
    if not record["base"] and record["payable"]:
        record["warnings"].append("No separate before-tax base in XML; do not invent VAT")
    return record


def printed_money(value):
    """MISA Vietnamese print convention: dots group whole VND amounts."""
    if not re.fullmatch(r"\d{1,3}(?:\.\d{3})+|\d+", value):
        raise ValueError(f"Unsupported printed amount: {value}")
    return str(int(value.replace(".", "")))


def parse_misa_pdf(txt, source):
    """Parse only the recognized MISA display layout; no filename fallback."""
    date = re.search(r"Ngày lập\s*\(Date\s*\)\s*(\d{1,2})\s+(\d{1,2})\s+(\d{4})", txt)
    serial = re.search(r"HÓA ĐƠN GIÁ TRỊ GIA TĂNG\s*\(Series\s*\)\s*([1-6][A-Z0-9]+)\s*\n\s*(\d+)\(No\.?\)", txt)
    seller = re.search(r"([^\n]+)Đơn vị bán hàng\s*\n\s*([0-9-]{10,14})Mã số thuế", txt)
    buyer = re.search(r":\s*([^\n]+)Tên đơn vị\s*\(Company\s*'s name\s*\)", txt)
    buyer_tax = re.findall(r":\s*([0-9-]{10,14})Mã số thuế", txt)
    base = re.search(r"([\d.]+)Cộng tiền hàng\s*\(Subtotal", txt)
    payable = re.search(r"([\d.]+)Tổng cộng tiền thanh toán", txt)
    vat = re.search(r"([\d.]+)Tiền thuế GTGT", txt)
    rate = re.search(r"Thuế suất GTGT\s*\(VAT rate\):\s*(\d+(?:[.,]\d+)?)\s*%", txt)
    currency = re.search(r"Cộng tiền hàng\s*\(Subtotal\s*\)\s*([A-Z]{3})\s*:", txt)
    if not all([date, serial, seller, buyer, buyer_tax, base, payable, vat, rate]):
        return None
    day, month, year = map(int, date.groups())
    d = datetime(year, month, day).date().isoformat()
    record = {
        "kind": "invoice", "direction": direction_for(source), "date": d,
        "series": serial.group(1), "number": serial.group(2),
        "seller_name": seller.group(1).strip(), "seller_tax_id": seller.group(2),
        "buyer_name": buyer.group(1).strip(), "buyer_tax_id": buyer_tax[0],
        "currency": currency.group(1) if currency else None, "base": printed_money(base.group(1)),
        "vat": printed_money(vat.group(1)), "payable": printed_money(payable.group(1)),
        "rate_groups": [{"rate": rate.group(1).replace(",", ".") + "%", "base": None, "vat": None}],
        "items": [], "sources": [source], "preferred_source": source,
        "source_format": "pdf", "warnings": [],
    }
    # Keep exact full text in documents[] rather than guessing table line splits.
    record["warnings"].append("PDF line items require review of documents[].text")
    before, tax, total = (Decimal(record[k]) for k in ("base", "vat", "payable"))
    if before + tax != total:
        record["warnings"].append("Printed base + VAT does not equal payable")
    printed_rate = Decimal(rate.group(1).replace(",", ".")) / Decimal(100)
    if abs(before * printed_rate - tax) > Decimal("0.5"):
        record["warnings"].append("Printed rate conflicts with printed VAT amount")
    return record


def pdf_identity(txt):
    series = re.findall(r"(?<![A-Z0-9])([1-6][CK]\d{2}[A-Z0-9]{3})(?![A-Z0-9])", txt)
    numbers = []
    for pattern in [
        r"Số\s*(?:hóa\s*đơn)?\s*\((?:Invoice\s*)?No\.?\)\s*:\s*(\d+)",
        r"\n\s*(\d+)\(No\.?\)",
        r"[1-6][CK]\d{2}[A-Z0-9]{3}\s*\n\s*(\d+)\s*\n",
        r"\bSố\s*:\s*(\d+)",
        r"(?:^|\n)\s*(\d+)\s*\n\s*[1-6][CK]\d{2}[A-Z0-9]{3}\s*\n",
    ]:
        numbers.extend(norm_number(v) for v in re.findall(pattern, txt, re.I))
    return set(series), set(numbers)


def scan(input_dir, previous=None):
    root = input_dir.resolve()
    result = {
        "schema_version": SCHEMA_VERSION, "input_root": str(root),
        "scanned_at": datetime.now(timezone.utc).isoformat(),
        "files": {}, "invoices": [], "documents": [], "errors": [],
        "unparsed": [], "ignored": [], "delta": {},
    }
    candidates = []
    pdf_text_cache = {}

    def error(source, problem):
        result["errors"].append({"source": source, "error": str(problem)})

    def read_xml(data, source):
        if Path(source.split("!")[-1]).stem.lower().endswith("_view"):
            result["ignored"].append({"source": source, "reason": "XML display copy"})
            return
        try:
            record = parse_xml(data, source)
            if record:
                candidates.append(record)
            else:
                result["unparsed"].append({"source": source, "reason": "Not canonical invoice XML"})
        except Exception as exc:
            error(source, exc)

    def read_pdf(data, source):
        if PdfReader is None:
            error(source, "pypdf is unavailable; PDF not read")
            return
        digest = hashlib.sha256(data).hexdigest()
        if digest not in pdf_text_cache:
            reader = PdfReader(io.BytesIO(data))
            pdf_text_cache[digest] = (len(reader.pages), "\n".join(page.extract_text() or "" for page in reader.pages))
        pages, txt = pdf_text_cache[digest]
        doc = {"source": source, "pages": pages, "text": txt, "kind": "unparsed_pdf"}
        upper = txt.upper()
        if "PHIẾU XUẤT KHO KIÊM VẬN CHUYỂN NỘI BỘ" in upper:
            doc["kind"] = "delivery_slip"
        elif ("ỦY NHIỆM CHI" in upper or "UỶ NHIỆM CHI" in upper
              or "PHIẾU BÁO NỢ" in upper or "PHIẾU GIAO DỊCH GHI NỢ" in upper
              or ("CHI TIẾT GIAO DỊCH" in upper and "MÃ GIAO DỊCH" in upper)):
            doc["kind"] = "payment_slip"
        elif "DANH SÁCH VẬN ĐƠN BẢNG KÊ XUẤT HÓA ĐƠN" in upper:
            doc["kind"] = "supporting_statement"
        else:
            record = parse_misa_pdf(txt, source)
            if record:
                candidates.append(record)
                doc.update(kind="invoice_pdf", invoice_key=invoice_key(record))
        result["documents"].append(doc)

    def read_zip(data, source, depth=0):
        if depth > MAX_ZIP_DEPTH:
            error(source, "Nested ZIP depth limit exceeded")
            return
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                for member in archive.infolist():
                    if member.is_dir():
                        continue
                    suffix = Path(member.filename).suffix.lower()
                    if suffix not in {".xml", ".zip", ".pdf"}:
                        continue
                    nested = source + "!" + member.filename
                    if member.file_size > MAX_MEMBER_BYTES:
                        error(nested, "ZIP member exceeds in-memory size limit")
                        continue
                    try:
                        content = archive.read(member)
                        if suffix == ".zip":
                            read_zip(content, nested, depth + 1)
                        elif suffix == ".pdf":
                            read_pdf(content, nested)
                        else:
                            read_xml(content, nested)
                    except Exception as exc:
                        error(nested, exc)
        except Exception as exc:
            error(source, exc)

    for path in sorted(root.rglob("*")):
        if not path.is_file():
            continue
        relative = path.relative_to(root).as_posix()
        try:
            data = path.read_bytes()
            result["files"][relative] = {"sha256": hashlib.sha256(data).hexdigest(), "bytes": len(data)}
            suffix = path.suffix.lower()
            if suffix == ".xml":
                read_xml(data, relative)
            elif suffix == ".zip":
                read_zip(data, relative)
            elif suffix == ".pdf":
                read_pdf(data, relative)
            else:
                result["ignored"].append({"source": relative, "reason": f"Unsupported companion type {suffix or '(none)'}"})
        except Exception as exc:
            error(relative, exc)

    merged = {}
    # Prefer direct XML, then archive XML, then PDF. Do not silently collapse conflicts.
    candidates.sort(key=lambda r: (r["source_format"] != "xml", r["preferred_source"].count("!"), r["preferred_source"]))
    for record in candidates:
        key = invoice_key(record)
        if not key:
            error(record["preferred_source"], "Invoice lacks issuer tax ID, series or number; no stable key")
            result["unparsed"].append(record)
            continue
        record["key"] = key
        if key not in merged:
            merged[key] = record
            continue
        current = merged[key]
        current["sources"].extend(s for s in record["sources"] if s not in current["sources"])
        fields = ["date", "base", "vat", "payable", "currency", "buyer_tax_id"]
        conflicts = [f for f in fields if current.get(f) is not None and record.get(f) is not None and
                     (Decimal(current[f]) != Decimal(record[f]) if f in {"base", "vat", "payable"} else current[f] != record[f])]
        if conflicts:
            current["warnings"].append("Duplicate source differs in " + ", ".join(conflicts) + ": " + record["preferred_source"])
    # Associate other PDF renderings only when contents confirm tax ID, serial and
    # a label-derived invoice number. Filenames alone never establish identity.
    for doc in result["documents"]:
        if doc["kind"] != "unparsed_pdf":
            continue
        series, numbers = pdf_identity(doc["text"])
        compact = re.sub(r"\s+", "", doc["text"])
        matches = [r for r in merged.values() if r["series"] in series and norm_number(r["number"]) in numbers
                   and re.sub(r"\s+", "", r["seller_tax_id"]) in compact]
        if len(matches) == 1:
            record = matches[0]
            if doc["source"] not in record["sources"]:
                record["sources"].append(doc["source"])
            doc.update(kind="invoice_pdf", invoice_key=record["key"])
        else:
            reason = "No readable text; OCR required" if not doc["text"].strip() else "Unsupported PDF layout or unmatched invoice; review text"
            result["unparsed"].append({"source": doc["source"], "reason": reason})
    result["invoices"] = sorted(merged.values(), key=lambda r: (r.get("date") or "", r["key"]))
    old_files = previous.get("files", {}) if previous else {}
    old_keys = {r["key"] for r in previous.get("invoices", [])} if previous else set()
    old_records = {r["key"]: r for r in previous.get("invoices", [])} if previous else {}
    new_files = result["files"]
    result["delta"] = {
        "baseline_supplied": previous is not None,
        "files_added": sorted(set(new_files) - set(old_files)),
        "files_removed": sorted(set(old_files) - set(new_files)),
        "files_changed": sorted(k for k in set(new_files) & set(old_files) if new_files[k]["sha256"] != old_files[k]["sha256"]),
        "invoice_keys_added": sorted(set(merged) - old_keys),
        "invoice_keys_removed": sorted(old_keys - set(merged)),
        "invoice_keys_changed": sorted(k for k in set(merged) & old_keys
                                       if any(merged[k].get(f) != old_records[k].get(f)
                                              for f in ["date", "seller_name", "seller_tax_id", "buyer_name", "buyer_tax_id", "base", "vat", "payable", "currency", "rate_groups", "items"])),
    }
    result["summary"] = {
        "files": len(new_files), "invoices": len(merged),
        "directions": dict(Counter(r["direction"] for r in merged.values())),
        "documents": dict(Counter(d["kind"] for d in result["documents"])),
        "errors": len(result["errors"]), "unparsed": len(result["unparsed"]),
        "invoices_with_warnings": sum(bool(r["warnings"]) for r in merged.values()),
    }
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, type=Path, help="Existing IN directory; read only")
    parser.add_argument("--output", required=True, type=Path, help="Analysis JSON outside IN")
    parser.add_argument("--previous", type=Path, help="Previous successful scan JSON for change comparison")
    args = parser.parse_args()
    if not args.input.is_dir():
        parser.error("--input must be an existing directory")
    if args.output.resolve().is_relative_to(args.input.resolve()):
        parser.error("--output must be outside the read-only input directory")
    try:
        previous = json.loads(args.previous.read_text(encoding="utf-8")) if args.previous else None
        if previous and Path(previous["input_root"]).resolve() != args.input.resolve():
            parser.error("Previous scan belongs to a different input directory")
        result = scan(args.input, previous)
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    except (OSError, ValueError, KeyError) as exc:
        parser.exit(2, f"Scan failed: {exc}\n")
    print(json.dumps(result["summary"], ensure_ascii=False))
    if result["errors"]:
        print("Review errors[] before updating a workbook.", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
