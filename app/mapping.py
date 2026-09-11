from __future__ import annotations

import re
import zipfile
from pathlib import Path, PurePosixPath
from typing import Dict, Iterator, List, Optional, Sequence, Tuple
from xml.etree import ElementTree as ET

from .models import MappingIndex, MappingRecord, norm


SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
S = "{" + SHEET_NS + "}"
R = "{" + DOC_REL_NS + "}"

DEPT = {
    "DEPARTMENT",
    "DEPT",
    "CRID",
    "CHANGE RULE ID",
    "CHANGE RULE",
    "CHANGE RULE DEPARTMENT",
    "PROCESS DEPARTMENT",
}
PLANT = {"PLANT", "WERKS", "SITE", "PLANT CODE"}
WC = {
    "NEW WC",
    "NEW WORK CENTER",
    "NEW WORKCENTER",
    "WORK CENTER",
    "WORKCENTER",
    "WC",
    "ARBPL",
    "NEW_WC",
}
SOURCE = {
    "CHANGE RULE SOURCE FILES",
    "CHANGE RULE SOURCE FILE",
    "SOURCE FILE",
    "CRMS SOURCE",
}


class MappingWorkbookError(RuntimeError):
    pass


class DepartmentPlantListError(RuntimeError):
    pass


def nh(value: object) -> str:
    if value is None:
        return ""
    return re.sub(
        r"\s+",
        " ",
        str(value).replace("_", " ").replace("-", " ").strip().upper(),
    )


def colnum(reference: str) -> int:
    match = re.match(r"([A-Z]+)", (reference or "").upper())
    if not match:
        return 0
    number = 0
    for character in match.group(1):
        number = number * 26 + ord(character) - 64
    return number


def shared_strings(package: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in package.namelist():
        return []
    root = ET.fromstring(package.read("xl/sharedStrings.xml"))
    return [
        "".join(text.text or "" for text in item.iter(S + "t"))
        for item in root.findall(S + "si")
    ]


def sheets(package: zipfile.ZipFile) -> List[Tuple[str, str]]:
    workbook = ET.fromstring(package.read("xl/workbook.xml"))
    relationships = ET.fromstring(
        package.read("xl/_rels/workbook.xml.rels")
    )
    relationship_map = {
        relationship.attrib.get("Id", ""): relationship.attrib.get("Target", "")
        for relationship in relationships
    }

    output: List[Tuple[str, str]] = []
    node = workbook.find(S + "sheets")
    if node is None:
        return output

    for sheet in node.findall(S + "sheet"):
        target = relationship_map.get(sheet.attrib.get(R + "id", ""), "")
        if not target:
            continue
        if target.startswith("/"):
            path = target.lstrip("/")
        elif target.startswith("xl/"):
            path = target
        else:
            path = str(PurePosixPath("xl") / PurePosixPath(target))
        output.append((sheet.attrib.get("name", "Sheet"), path))
    return output


def cell_value(cell: ET.Element, strings: Sequence[str]):
    cell_type = cell.attrib.get("t", "")
    if cell_type == "inlineStr":
        return "".join(text.text or "" for text in cell.iter(S + "t"))

    value = cell.find(S + "v")
    if value is None:
        return None

    raw = value.text or ""
    if cell_type == "s":
        try:
            return strings[int(raw)]
        except (ValueError, IndexError):
            return raw
    if cell_type == "b":
        return raw == "1"
    return raw


def iter_rows(
    package: zipfile.ZipFile,
    path: str,
    strings: Sequence[str],
    max_rows: Optional[int] = None,
) -> Iterator[Tuple[int, List[object]]]:
    count = 0
    with package.open(path) as stream:
        for _event, element in ET.iterparse(stream, events=("end",)):
            if element.tag != S + "row":
                continue

            values: Dict[int, object] = {}
            for cell in element.findall(S + "c"):
                number = colnum(cell.attrib.get("r", ""))
                if number:
                    values[number] = cell_value(cell, strings)

            maximum = max(values) if values else 0
            yield int(element.attrib.get("r", count + 1)), [
                values.get(index) for index in range(1, maximum + 1)
            ]

            count += 1
            element.clear()
            if max_rows and count >= max_rows:
                break


def alias(values: Sequence[object], aliases: set[str]) -> Optional[int]:
    for index, value in enumerate(values):
        if nh(value) in aliases:
            return index
    return None


def inspect_mapping(path: str) -> Dict[str, object]:
    workbook_path = Path(path)
    if not workbook_path.exists():
        raise MappingWorkbookError(
            f"Mapping workbook not found: {workbook_path}"
        )

    try:
        with zipfile.ZipFile(workbook_path) as package:
            strings = shared_strings(package)
            candidates = []
            ordered = sorted(
                sheets(package),
                key=lambda item: (
                    0 if "SUMMARY" in item[0].upper() else 1,
                    item[0],
                ),
            )

            for sheet_name, sheet_path in ordered:
                for row_number, values in iter_rows(
                    package, sheet_path, strings, 50
                ):
                    department_index = alias(values, DEPT)
                    plant_index = alias(values, PLANT)
                    workcenter_index = alias(values, WC)
                    source_index = alias(values, SOURCE)

                    found = sum(
                        index is not None
                        for index in (
                            department_index,
                            plant_index,
                            workcenter_index,
                        )
                    )
                    if found < 2:
                        continue

                    score = (
                        found * 100
                        - row_number
                        + (30 if "SUMMARY" in sheet_name.upper() else 0)
                        + (50 if found == 3 else 0)
                    )
                    candidates.append(
                        {
                            "score": score,
                            "sheet_name": sheet_name,
                            "sheet_path": sheet_path,
                            "header_row": row_number,
                            "values": values,
                            "department_index": department_index,
                            "plant_index": plant_index,
                            "workcenter_index": workcenter_index,
                            "source_index": source_index,
                        }
                    )

            if not candidates:
                raise MappingWorkbookError(
                    "No worksheet with Department/CRID, Plant and Work Center "
                    "headers was found."
                )

            best = max(candidates, key=lambda item: item["score"])
            if None in (
                best["department_index"],
                best["plant_index"],
                best["workcenter_index"],
            ):
                raise MappingWorkbookError(
                    f"Worksheet '{best['sheet_name']}' is missing a required field."
                )
            return best

    except zipfile.BadZipFile as exc:
        raise MappingWorkbookError(f"Invalid XLSX: {exc}") from exc


def load_mapping(path: str, log=None) -> MappingIndex:
    info = inspect_mapping(path)
    workbook_path = Path(path)

    with zipfile.ZipFile(workbook_path) as package:
        strings = shared_strings(package)
        values = info["values"]
        department_index = int(info["department_index"])
        plant_index = int(info["plant_index"])
        workcenter_index = int(info["workcenter_index"])
        source_index = info["source_index"]

        mapping = MappingIndex(
            source_path=str(workbook_path.resolve()),
            sheet_name=str(info["sheet_name"]),
            header_row=int(info["header_row"]),
            department_column=nh(values[department_index]),
            plant_column=nh(values[plant_index]),
            workcenter_column=nh(values[workcenter_index]),
        )

        seen = set()
        for row_number, row in iter_rows(
            package, str(info["sheet_path"]), strings
        ):
            if row_number <= mapping.header_row:
                continue

            department = norm(
                row[department_index]
                if department_index < len(row)
                else None
            )
            plant = norm(
                row[plant_index] if plant_index < len(row) else None
            )
            workcenter = norm(
                row[workcenter_index]
                if workcenter_index < len(row)
                else None
            )
            source = (
                str(row[source_index] or "").strip()
                if source_index is not None and source_index < len(row)
                else ""
            )

            if not department and not plant and not workcenter:
                continue
            if not department or not plant:
                mapping.skipped_rows += 1
                continue

            key = (department, plant, workcenter, source)
            if key in seen:
                continue

            seen.add(key)
            mapping.records.append(
                MappingRecord(
                    department=department,
                    plant=plant,
                    workcenter=workcenter,
                    source_file=source,
                    source_sheet=mapping.sheet_name,
                    source_row=row_number,
                )
            )

        mapping.rebuild()
        if log:
            log(
                f"Excel mapping loaded: {len(mapping.records):,} rows, "
                f"{len(mapping.plants)} plants, "
                f"{len(mapping.departments)} departments/CRIDs."
            )
        return mapping


def _split_department_plant_line(line: str) -> Optional[Tuple[str, str]]:
    stripped = line.strip().lstrip("\ufeff")
    if not stripped or stripped.startswith("#"):
        return None

    parts: List[str]
    if "|" in stripped:
        parts = stripped.split("|", 1)
    elif "\t" in stripped:
        parts = stripped.split("\t", 1)
    elif "," in stripped:
        parts = stripped.split(",", 1)
    elif ";" in stripped:
        parts = stripped.split(";", 1)
    else:
        parts = re.split(r"\s{2,}", stripped, maxsplit=1)
        if len(parts) == 1:
            simple = stripped.split()
            if len(simple) == 2:
                parts = simple

    if len(parts) != 2:
        raise ValueError(
            "Expected two fields separated by |, tab, comma, semicolon, "
            "or two or more spaces."
        )

    return parts[0].strip(), parts[1].strip()


def parse_department_plant_text(
    text: str,
    source_name: str = "Pasted Department-Plant list",
) -> Tuple[MappingIndex, List[str]]:
    mapping = MappingIndex(
        source_path=source_name,
        sheet_name="Department-Plant List",
        header_row=1,
        department_column="DEPARTMENT",
        plant_column="PLANT",
        workcenter_column="",
    )
    warnings: List[str] = []
    seen = set()

    for line_number, line in enumerate(text.splitlines(), start=1):
        try:
            parsed = _split_department_plant_line(line)
        except ValueError as exc:
            warnings.append(f"Line {line_number}: {exc}")
            continue

        if parsed is None:
            continue

        department_raw, plant_raw = parsed
        department_header = nh(department_raw)
        plant_header = nh(plant_raw)
        if (
            department_header in DEPT
            and plant_header in PLANT
        ):
            continue

        department = norm(department_raw)
        plant = norm(plant_raw)

        if not department or not plant:
            warnings.append(
                f"Line {line_number}: Department and Plant are both required."
            )
            continue

        key = (department, plant)
        if key in seen:
            continue

        seen.add(key)
        mapping.records.append(
            MappingRecord(
                department=department,
                plant=plant,
                workcenter="",
                source_file=source_name,
                source_sheet="Department-Plant List",
                source_row=line_number,
            )
        )

    if not mapping.records:
        detail = f" {warnings[0]}" if warnings else ""
        raise DepartmentPlantListError(
            "No valid Department | Plant rows were found." + detail
        )

    mapping.skipped_rows = len(warnings)
    mapping.rebuild()
    return mapping, warnings


def load_department_plant_file(path: str) -> Tuple[MappingIndex, List[str]]:
    file_path = Path(path)
    if not file_path.exists():
        raise DepartmentPlantListError(
            f"Department-Plant file not found: {file_path}"
        )

    text = file_path.read_text(encoding="utf-8-sig", errors="replace")
    return parse_department_plant_text(
        text, source_name=f"Imported list: {file_path.name}"
    )


def merge_mappings(*mappings: Optional[MappingIndex]) -> Optional[MappingIndex]:
    available = [mapping for mapping in mappings if mapping is not None]
    if not available:
        return None

    merged = MappingIndex(
        source_path=" + ".join(
            mapping.source_path for mapping in available if mapping.source_path
        ),
        sheet_name="Combined Sources",
        department_column="DEPARTMENT / CRID",
        plant_column="PLANT",
        workcenter_column="WORKCENTER (when available)",
    )

    seen = set()
    for mapping in available:
        merged.skipped_rows += mapping.skipped_rows
        for record in mapping.records:
            key = (
                norm(record.department),
                norm(record.plant),
                norm(record.workcenter),
                record.source_file,
                record.source_sheet,
            )
            if key in seen:
                continue
            seen.add(key)
            merged.records.append(record)

    merged.rebuild()
    return merged
