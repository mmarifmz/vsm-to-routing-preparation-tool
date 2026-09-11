from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple


Point = Tuple[float, float]


def norm(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip().upper()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def placeholder_wc(value: str) -> bool:
    return norm(value) in {
        "",
        "BLANK",
        "BLANKS",
        "N/A",
        "NA",
        "NONE",
        "NULL",
        "-",
        "TBC",
        "TBD",
    }


@dataclass(frozen=True)
class MappingRecord:
    department: str
    plant: str
    workcenter: str = ""
    source_file: str = ""
    source_sheet: str = ""
    source_row: int = 0


@dataclass
class MappingIndex:
    source_path: str = ""
    sheet_name: str = ""
    header_row: int = 0
    department_column: str = ""
    plant_column: str = ""
    workcenter_column: str = ""
    records: List[MappingRecord] = field(default_factory=list)
    skipped_rows: int = 0

    _lookup: Dict[Tuple[str, str], Set[str]] = field(default_factory=dict, init=False)
    _dept_by_plant: Dict[str, Set[str]] = field(default_factory=dict, init=False)
    _wc_by_plant: Dict[str, Set[str]] = field(default_factory=dict, init=False)
    _sources_by_dept_plant: Dict[Tuple[str, str], Set[str]] = field(
        default_factory=dict, init=False
    )

    def rebuild(self) -> None:
        self._lookup.clear()
        self._dept_by_plant.clear()
        self._wc_by_plant.clear()
        self._sources_by_dept_plant.clear()

        for row in self.records:
            plant = norm(row.plant)
            workcenter = norm(row.workcenter)
            department = norm(row.department)
            if not plant or not department:
                continue

            self._dept_by_plant.setdefault(plant, set()).add(department)
            source = row.source_sheet.strip() or row.source_file.strip() or "Mapping source"
            self._sources_by_dept_plant.setdefault((plant, department), set()).add(source)

            if workcenter and not placeholder_wc(workcenter):
                self._lookup.setdefault((plant, workcenter), set()).add(department)
                self._wc_by_plant.setdefault(plant, set()).add(workcenter)

    @property
    def plants(self) -> List[str]:
        return sorted(self._dept_by_plant)

    @property
    def departments(self) -> List[str]:
        return sorted(
            {
                department
                for departments in self._dept_by_plant.values()
                for department in departments
            }
        )

    @property
    def department_plant_pair_count(self) -> int:
        return sum(len(values) for values in self._dept_by_plant.values())

    @property
    def workcenter_record_count(self) -> int:
        return sum(
            1
            for row in self.records
            if row.workcenter and not placeholder_wc(row.workcenter)
        )

    def departments_for_plant(self, plant: str) -> List[str]:
        return sorted(self._dept_by_plant.get(norm(plant), set()))

    def workcenters_for_plant(self, plant: str) -> List[str]:
        return sorted(self._wc_by_plant.get(norm(plant), set()))

    def departments_for_wc(self, plant: str, workcenter: str) -> List[str]:
        return sorted(self._lookup.get((norm(plant), norm(workcenter)), set()))

    def sources_for_department(self, plant: str, department: str) -> List[str]:
        return sorted(
            self._sources_by_dept_plant.get((norm(plant), norm(department)), set())
        )


@dataclass
class ShapeRecord:
    shape_id: str
    name: str = ""
    name_u: str = ""
    kind: str = "process"
    shape_type: str = ""
    master_id: str = ""
    parent_id: str = ""
    depth: int = 0
    has_children: bool = False
    text: str = ""
    pin_x: Optional[float] = None
    pin_y: Optional[float] = None
    width: Optional[float] = None
    height: Optional[float] = None
    begin_x: Optional[float] = None
    begin_y: Optional[float] = None
    end_x: Optional[float] = None
    end_y: Optional[float] = None
    angle: Optional[float] = None
    corners: List[Point] = field(default_factory=list)
    workcenters: List[str] = field(default_factory=list)
    properties: Dict[str, str] = field(default_factory=dict)


@dataclass
class ConnectionRecord:
    connector_id: str
    source_shape_id: str = ""
    target_shape_id: str = ""
    source_cell: str = ""
    target_cell: str = ""




@dataclass
class ChangeRuleDraftRecord:
    crid: str = ""
    plant: str = ""
    alternate: str = ""
    sequence: int = 0
    subseq: int = 1
    old_wc: str = ""
    new_wc: str = ""
    op_descriptions: str = ""
    rule_type: str = "ALL"
    column: str = ""
    status: str = "Ready"
    notes: str = ""
    source_shape_id: str = ""


@dataclass
class PageRecord:
    page_id: str
    name: str
    name_u: str
    xml_path: str
    width: float = 11.0
    height: float = 8.5
    shapes: List[ShapeRecord] = field(default_factory=list)
    connections: List[ConnectionRecord] = field(default_factory=list)
    workcenter_counts: Dict[str, int] = field(default_factory=dict)
    candidate_counts: Dict[str, int] = field(default_factory=dict)
    candidate_wcs: Dict[str, Set[str]] = field(default_factory=dict)
    suggested_department: str = ""
    assigned_department: str = ""
    assignment_mode: str = "auto"  # auto, manual, or cleared
    confidence: float = 0.0
    matched_wcs: List[str] = field(default_factory=list)
    unmatched_wcs: List[str] = field(default_factory=list)
    loaded: bool = False
    loading: bool = False
    load_error: str = ""
    loaded_plant: str = ""
    change_rule_draft: List[ChangeRuleDraftRecord] = field(default_factory=list)
    change_rule_generated_at: str = ""
    change_rule_rerun_at: str = ""
    change_rule_source: str = ""
    change_rule_generation_count: int = 0

    @property
    def text_shape_count(self) -> int:
        return sum(bool(shape.text.strip()) for shape in self.shapes)

    @property
    def connector_count(self) -> int:
        return sum(shape.kind == "connector" for shape in self.shapes)

    @property
    def drawable_shape_count(self) -> int:
        return sum(
            shape.kind != "connector" and not shape.has_children
            for shape in self.shapes
        )


@dataclass
class FileRecord:
    path: str
    plant: str = ""
    pages: List[PageRecord] = field(default_factory=list)
    status: str = "Ready"
    error: str = ""

    @property
    def file_name(self) -> str:
        return Path(self.path).name
