from __future__ import annotations

import math
import re
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from typing import Callable, Dict, List, Optional, Sequence, Set, Tuple
from xml.etree import ElementTree as ET

from .models import (
    ConnectionRecord,
    FileRecord,
    MappingIndex,
    PageRecord,
    ShapeRecord,
    department_slug,
    norm,
)

VNS = "http://schemas.microsoft.com/office/visio/2012/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
V = "{" + VNS + "}"
R = "{" + RNS + "}"

GENERIC = [
    re.compile(r"(?<![A-Z0-9])([A-Z]{1,6}\d{3})(?![A-Z0-9])"),
    re.compile(r"(?<![A-Z0-9])([A-Z]{2}\d{2}[A-Z]{2,8})(?![A-Z0-9])"),
    re.compile(r"(?<![A-Z0-9])(P\d{6})(?![A-Z0-9])"),
]

# Affine matrix: x' = a*x + c*y + e, y' = b*x + d*y + f
Matrix = Tuple[float, float, float, float, float, float]
IDENTITY: Matrix = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)


class VsdxReadError(RuntimeError):
    pass


_DEPARTMENT_TOKEN_ALIASES = {
    "assy": "assembly",
    "canti": "cantilever",
    "fab": "fabrication",
    "mac": "machining",
    "rub": "rubber",
}
_DEPARTMENT_COMPOUNDS = {
    "cutrub": ("cut", "rubber"),
    "mat4sub": ("material", "subcon"),
    "mat4subcon": ("material", "subcon"),
    "rubfab": ("rubber", "fabrication"),
}
_DEPARTMENT_NOISE = {"and", "build", "for", "the"}


def _department_terms(value: object) -> Tuple[str, ...]:
    terms: List[str] = []
    for token in department_slug(value).split("_"):
        expanded = _DEPARTMENT_COMPOUNDS.get(token, (token,))
        for item in expanded:
            item = _DEPARTMENT_TOKEN_ALIASES.get(item, item)
            if item and item not in _DEPARTMENT_NOISE:
                terms.append(item)
    return tuple(sorted(set(terms)))


def page_department_candidate(
    page_name: str,
    known_departments: Sequence[str] = (),
) -> str:
    """Create a neutral CRID suggestion from the VSM tab name only."""
    lowered = str(page_name or "").lower()
    if "combined into" in lowered or re.search(r"\bold\b", lowered):
        return ""
    slug = department_slug(page_name)
    if not slug or slug in {"template", "page", "sheet", "drawing", "untitled"}:
        return ""
    return slug


def flt(value, default=None):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def cells(element: ET.Element) -> Dict[str, str]:
    return {
        cell.attrib.get("N", ""): cell.attrib.get("V", "")
        for cell in element.findall(V + "Cell")
        if cell.attrib.get("N")
    }


def clean_text(text: str) -> str:
    text = (text or "").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s*\n+", "\n", text)
    return text.strip()


def shape_text(shape: ET.Element) -> str:
    text_node = shape.find(V + "Text")
    return clean_text("".join(text_node.itertext())) if text_node is not None else ""


def props(shape: ET.Element) -> Dict[str, str]:
    output: Dict[str, str] = {}
    for section in shape.findall(V + "Section"):
        if section.attrib.get("N") != "Property":
            continue
        for row in section.findall(V + "Row"):
            cell_map = cells(row)
            label = (
                cell_map.get("Label")
                or row.attrib.get("N")
                or row.attrib.get("IX", "")
            )
            if label:
                output[str(label)] = str(cell_map.get("Value", ""))
    return output


def kind(shape: ET.Element, cell_map: Dict[str, str]) -> str:
    name = (shape.attrib.get("Name", "") + " " + shape.attrib.get("NameU", "")).lower()
    if all(value in cell_map for value in ("BeginX", "BeginY", "EndX", "EndY")):
        return "connector"
    if "decision" in name or "diamond" in name:
        return "decision"
    if "start/end" in name or "terminator" in name:
        return "start_end"
    if "document" in name:
        return "document"
    if shape.find(V + "Shapes") is not None:
        return "group"
    return "process"


def detect_wc(text: str, known: Sequence[str]) -> List[str]:
    normalized = norm(text)
    if not normalized:
        return []

    found: Set[str] = set()
    for workcenter in known:
        workcenter = norm(workcenter)
        if not workcenter:
            continue
        if len(workcenter) <= 2:
            if normalized == workcenter:
                found.add(workcenter)
        elif re.search(
            rf"(?<![A-Z0-9]){re.escape(workcenter)}(?![A-Z0-9])", normalized
        ):
            found.add(workcenter)

    for pattern in GENERIC:
        found.update(norm(match) for match in pattern.findall(normalized))
    return sorted(found)


def page_dim(page: ET.Element, name: str, default: float) -> float:
    page_sheet = page.find(V + "PageSheet")
    if page_sheet is None:
        return default
    return flt(cells(page_sheet).get(name), default) or default


def quick_scan(path: str) -> List[PageRecord]:
    file_path = Path(path)
    if not file_path.exists():
        raise VsdxReadError(f"VSDX not found: {file_path}")

    try:
        with zipfile.ZipFile(file_path) as package:
            pages_root = ET.fromstring(package.read("visio/pages/pages.xml"))
            rels_root = ET.fromstring(
                package.read("visio/pages/_rels/pages.xml.rels")
            )
            rel_map = {
                relationship.attrib["Id"]: relationship.attrib["Target"]
                for relationship in rels_root
                if relationship.attrib.get("Id")
                and relationship.attrib.get("Target")
            }

            output: List[PageRecord] = []
            for page in pages_root.findall(V + "Page"):
                relationship = page.find(V + "Rel")
                if relationship is None:
                    continue
                target = rel_map.get(relationship.attrib.get(R + "id", ""), "")
                if not target:
                    continue
                output.append(
                    PageRecord(
                        page_id=page.attrib.get("ID", ""),
                        name=page.attrib.get(
                            "Name", page.attrib.get("NameU", "Untitled")
                        ),
                        name_u=page.attrib.get(
                            "NameU", page.attrib.get("Name", "Untitled")
                        ),
                        xml_path=f"visio/pages/{target}",
                        width=page_dim(page, "PageWidth", 11.0),
                        height=page_dim(page, "PageHeight", 8.5),
                    )
                )
            return output
    except Exception as exc:
        raise VsdxReadError(f"Unable to read VSDX: {exc}") from exc


def compose(first: Matrix, second: Matrix) -> Matrix:
    """Return first(second(point))."""
    a1, b1, c1, d1, e1, f1 = first
    a2, b2, c2, d2, e2, f2 = second
    return (
        a1 * a2 + c1 * b2,
        b1 * a2 + d1 * b2,
        a1 * c2 + c1 * d2,
        b1 * c2 + d1 * d2,
        a1 * e2 + c1 * f2 + e1,
        b1 * e2 + d1 * f2 + f1,
    )


def apply(matrix: Matrix, x: float, y: float) -> Tuple[float, float]:
    a, b, c, d, e, f = matrix
    return a * x + c * y + e, b * x + d * y + f


def shape_transform(cell_map: Dict[str, str]) -> Matrix:
    width = abs(flt(cell_map.get("Width"), 0.0) or 0.0)
    height = abs(flt(cell_map.get("Height"), 0.0) or 0.0)
    pin_x = flt(cell_map.get("PinX"), 0.0) or 0.0
    pin_y = flt(cell_map.get("PinY"), 0.0) or 0.0
    loc_pin_x = flt(cell_map.get("LocPinX"), width / 2.0)
    loc_pin_y = flt(cell_map.get("LocPinY"), height / 2.0)
    angle = flt(cell_map.get("Angle"), 0.0) or 0.0
    flip_x = -1.0 if str(cell_map.get("FlipX", "0")) in {"1", "TRUE"} else 1.0
    flip_y = -1.0 if str(cell_map.get("FlipY", "0")) in {"1", "TRUE"} else 1.0

    cosine = math.cos(angle)
    sine = math.sin(angle)
    translate_to_origin: Matrix = (1, 0, 0, 1, -loc_pin_x, -loc_pin_y)
    scale: Matrix = (flip_x, 0, 0, flip_y, 0, 0)
    rotate: Matrix = (cosine, sine, -sine, cosine, 0, 0)
    translate_to_pin: Matrix = (1, 0, 0, 1, pin_x, pin_y)
    return compose(
        translate_to_pin,
        compose(rotate, compose(scale, translate_to_origin)),
    )


def matrix_angle(matrix: Matrix) -> float:
    return math.atan2(matrix[1], matrix[0])


def parse_shapes(
    node: Optional[ET.Element],
    known: Sequence[str],
    parent_id: str = "",
    parent_transform: Matrix = IDENTITY,
    depth: int = 0,
) -> List[ShapeRecord]:
    if node is None:
        return []

    output: List[ShapeRecord] = []
    for shape in node.findall(V + "Shape"):
        cell_map = cells(shape)
        child_node = shape.find(V + "Shapes")
        has_children = child_node is not None and bool(child_node.findall(V + "Shape"))
        shape_kind = kind(shape, cell_map)

        width = abs(flt(cell_map.get("Width"), 0.0) or 0.0)
        height = abs(flt(cell_map.get("Height"), 0.0) or 0.0)
        pin_local_x = flt(cell_map.get("PinX"))
        pin_local_y = flt(cell_map.get("PinY"))
        pin_absolute = (
            apply(parent_transform, pin_local_x, pin_local_y)
            if pin_local_x is not None and pin_local_y is not None
            else (None, None)
        )

        own_transform = compose(parent_transform, shape_transform(cell_map))
        corners = (
            [
                apply(own_transform, 0.0, 0.0),
                apply(own_transform, width, 0.0),
                apply(own_transform, width, height),
                apply(own_transform, 0.0, height),
            ]
            if width > 0 and height > 0
            else []
        )

        begin_x = flt(cell_map.get("BeginX"))
        begin_y = flt(cell_map.get("BeginY"))
        end_x = flt(cell_map.get("EndX"))
        end_y = flt(cell_map.get("EndY"))
        begin_absolute = (
            apply(parent_transform, begin_x, begin_y)
            if begin_x is not None and begin_y is not None
            else (None, None)
        )
        end_absolute = (
            apply(parent_transform, end_x, end_y)
            if end_x is not None and end_y is not None
            else (None, None)
        )

        record = ShapeRecord(
            shape_id=shape.attrib.get("ID", ""),
            name=shape.attrib.get("Name", ""),
            name_u=shape.attrib.get("NameU", ""),
            kind=shape_kind,
            shape_type=shape.attrib.get("Type", ""),
            master_id=shape.attrib.get("Master", ""),
            parent_id=parent_id,
            depth=depth,
            has_children=has_children,
            text=shape_text(shape),
            pin_x=pin_absolute[0],
            pin_y=pin_absolute[1],
            width=width or None,
            height=height or None,
            begin_x=begin_absolute[0],
            begin_y=begin_absolute[1],
            end_x=end_absolute[0],
            end_y=end_absolute[1],
            angle=matrix_angle(own_transform),
            corners=corners,
            properties=props(shape),
        )
        record.workcenters = detect_wc(record.text, known)
        output.append(record)

        if child_node is not None:
            output.extend(
                parse_shapes(
                    child_node,
                    known,
                    parent_id=record.shape_id,
                    parent_transform=own_transform,
                    depth=depth + 1,
                )
            )

    return output


def parse_connections(root: ET.Element) -> List[ConnectionRecord]:
    node = root.find(V + "Connects")
    if node is None:
        return []

    grouped = defaultdict(dict)
    for connection in node.findall(V + "Connect"):
        connector_id = connection.attrib.get("FromSheet", "")
        from_cell = connection.attrib.get("FromCell", "")
        target = connection.attrib.get("ToSheet", "")
        target_cell = connection.attrib.get("ToCell", "")
        if not connector_id or not target:
            continue
        if from_cell.startswith("Begin"):
            grouped[connector_id]["begin"] = (target, target_cell)
        elif from_cell.startswith("End"):
            grouped[connector_id]["end"] = (target, target_cell)

    output = []
    for connector_id, ends in grouped.items():
        begin = ends.get("begin", ("", ""))
        end = ends.get("end", ("", ""))
        output.append(
            ConnectionRecord(
                connector_id,
                begin[0],
                end[0],
                begin[1],
                end[1],
            )
        )
    return output


def apply_mapping(
    page: PageRecord,
    plant: str,
    mapping: Optional[MappingIndex],
    known_workcenters: Sequence[str] = (),
) -> None:
    known = sorted(
        {
            norm(value)
            for value in [
                *(mapping.workcenters_for_plant(plant) if mapping else []),
                *known_workcenters,
            ]
            if norm(value)
        }
    )
    for shape in page.shapes:
        shape.workcenters = detect_wc(shape.text, known)

    workcenter_counts = Counter()
    matched: Set[str] = set()
    unmatched: Set[str] = set()
    known_set = set(known)

    for shape in page.shapes:
        for workcenter in shape.workcenters:
            workcenter_counts[workcenter] += 1
            if workcenter in known_set:
                matched.add(workcenter)
            else:
                unmatched.add(workcenter)

    name_candidate = page_department_candidate(page.name)

    page.workcenter_counts = dict(sorted(workcenter_counts.items()))
    page.candidate_counts = {name_candidate: 1} if name_candidate else {}
    page.candidate_wcs = {name_candidate: set(workcenter_counts)} if name_candidate else {}
    page.matched_wcs = sorted(matched)
    page.unmatched_wcs = sorted(unmatched)
    page.suggested_department = name_candidate
    page.suggestion_source = "Tab name" if name_candidate else ""
    page.confidence = 1.0 if name_candidate else 0.0
    if page.assignment_mode == "auto":
        # A suggestion is evidence, not approval. The user confirms the CRID.
        page.assigned_department = ""

    page.loaded_plant = norm(plant)


def load_page(
    path: str,
    page: PageRecord,
    plant: str,
    mapping: Optional[MappingIndex] = None,
    force: bool = False,
    known_workcenters: Sequence[str] = (),
) -> PageRecord:
    if page.loaded and not force:
        apply_mapping(page, plant, mapping, known_workcenters)
        return page

    page.loading = True
    page.load_error = ""
    try:
        with zipfile.ZipFile(path) as package:
            root = ET.fromstring(package.read(page.xml_path))
        known = sorted(
            {
                *(mapping.workcenters_for_plant(plant) if mapping else []),
                *(norm(value) for value in known_workcenters if norm(value)),
            }
        )
        page.shapes = parse_shapes(root.find(V + "Shapes"), known)
        page.connections = parse_connections(root)
        page.loaded = True
        apply_mapping(page, plant, mapping, known_workcenters)
        return page
    except Exception as exc:
        page.loaded = False
        page.load_error = str(exc)
        raise VsdxReadError(f"Failed reading tab '{page.name}': {exc}") from exc
    finally:
        page.loading = False


def analyze(
    path: str,
    plant: str,
    mapping: Optional[MappingIndex] = None,
    log: Optional[Callable[[str], None]] = None,
    page_callback=None,
    pages: Optional[List[PageRecord]] = None,
    known_workcenters: Sequence[str] = (),
) -> FileRecord:
    file_path = Path(path)
    pages = pages if pages is not None else quick_scan(path)
    record = FileRecord(str(file_path.resolve()), norm(plant), pages, "Processing", "")

    try:
        total = len(pages)
        for index, page in enumerate(pages, start=1):
            if log:
                source = "cached" if page.loaded else "VSDX XML"
                log(
                    f"[{file_path.name}] Reading tab {index}/{total}: "
                    f"{page.name} ({source})"
                )
            load_page(path, page, record.plant, mapping, known_workcenters=known_workcenters)
            if page_callback:
                page_callback(index, total, page)
        record.status = "Analyzed"
        return record
    except Exception as exc:
        record.status = "Error"
        record.error = str(exc)
        raise VsdxReadError(f"Failed reading {file_path.name}: {exc}") from exc
