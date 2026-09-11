from __future__ import annotations

import hashlib
import json
import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from .models import ChangeRuleDraftRecord, ConnectionRecord, FileRecord, PageRecord, ShapeRecord


PACKAGE_SCHEMA = 2
PACKAGE_EXTENSION = ".vpa"


class AnalysisPackageError(RuntimeError):
    pass


@dataclass
class ImportedAnalysisPackage:
    files: List[FileRecord]
    workbook_mapping_path: str = ""
    manual_department_plant_text: str = ""
    routing_template_path: str = ""
    package_path: str = ""
    created_at: str = ""
    app_version: str = ""


def _safe_member_name(value: str, fallback: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._ ()\[\]-]+", "_", Path(value).name).strip(" .")
    return cleaned or fallback


def _shape_to_dict(shape: ShapeRecord) -> Dict[str, Any]:
    return {
        "shape_id": shape.shape_id,
        "name": shape.name,
        "name_u": shape.name_u,
        "kind": shape.kind,
        "shape_type": shape.shape_type,
        "master_id": shape.master_id,
        "parent_id": shape.parent_id,
        "depth": shape.depth,
        "has_children": shape.has_children,
        "text": shape.text,
        "pin_x": shape.pin_x,
        "pin_y": shape.pin_y,
        "width": shape.width,
        "height": shape.height,
        "begin_x": shape.begin_x,
        "begin_y": shape.begin_y,
        "end_x": shape.end_x,
        "end_y": shape.end_y,
        "angle": shape.angle,
        "corners": [list(point) for point in shape.corners],
        "workcenters": list(shape.workcenters),
        "properties": dict(shape.properties),
    }


def _shape_from_dict(data: Dict[str, Any]) -> ShapeRecord:
    return ShapeRecord(
        shape_id=str(data.get("shape_id", "")),
        name=str(data.get("name", "")),
        name_u=str(data.get("name_u", "")),
        kind=str(data.get("kind", "process")),
        shape_type=str(data.get("shape_type", "")),
        master_id=str(data.get("master_id", "")),
        parent_id=str(data.get("parent_id", "")),
        depth=int(data.get("depth", 0) or 0),
        has_children=bool(data.get("has_children", False)),
        text=str(data.get("text", "")),
        pin_x=data.get("pin_x"),
        pin_y=data.get("pin_y"),
        width=data.get("width"),
        height=data.get("height"),
        begin_x=data.get("begin_x"),
        begin_y=data.get("begin_y"),
        end_x=data.get("end_x"),
        end_y=data.get("end_y"),
        angle=data.get("angle"),
        corners=[tuple(point) for point in data.get("corners", []) if len(point) == 2],
        workcenters=[str(value) for value in data.get("workcenters", [])],
        properties={str(key): str(value) for key, value in data.get("properties", {}).items()},
    )


def _connection_to_dict(connection: ConnectionRecord) -> Dict[str, Any]:
    return {
        "connector_id": connection.connector_id,
        "source_shape_id": connection.source_shape_id,
        "target_shape_id": connection.target_shape_id,
        "source_cell": connection.source_cell,
        "target_cell": connection.target_cell,
    }


def _connection_from_dict(data: Dict[str, Any]) -> ConnectionRecord:
    return ConnectionRecord(
        connector_id=str(data.get("connector_id", "")),
        source_shape_id=str(data.get("source_shape_id", "")),
        target_shape_id=str(data.get("target_shape_id", "")),
        source_cell=str(data.get("source_cell", "")),
        target_cell=str(data.get("target_cell", "")),
    )


def _page_to_dict(page: PageRecord) -> Dict[str, Any]:
    return {
        "page_id": page.page_id,
        "name": page.name,
        "name_u": page.name_u,
        "xml_path": page.xml_path,
        "width": page.width,
        "height": page.height,
        "shapes": [_shape_to_dict(shape) for shape in page.shapes],
        "connections": [_connection_to_dict(item) for item in page.connections],
        "workcenter_counts": dict(page.workcenter_counts),
        "candidate_counts": dict(page.candidate_counts),
        "candidate_wcs": {
            department: sorted(workcenters)
            for department, workcenters in page.candidate_wcs.items()
        },
        "suggested_department": page.suggested_department,
        "assigned_department": page.assigned_department,
        "assignment_mode": page.assignment_mode,
        "confidence": page.confidence,
        "matched_wcs": list(page.matched_wcs),
        "unmatched_wcs": list(page.unmatched_wcs),
        "loaded": page.loaded,
        "load_error": page.load_error,
        "loaded_plant": page.loaded_plant,
        "change_rule_draft": [row.__dict__.copy() for row in page.change_rule_draft],
        "change_rule_generated_at": page.change_rule_generated_at,
        "change_rule_rerun_at": page.change_rule_rerun_at,
        "change_rule_source": page.change_rule_source,
        "change_rule_generation_count": page.change_rule_generation_count,
    }


def _page_from_dict(data: Dict[str, Any]) -> PageRecord:
    page = PageRecord(
        page_id=str(data.get("page_id", "")),
        name=str(data.get("name", "Untitled")),
        name_u=str(data.get("name_u", data.get("name", "Untitled"))),
        xml_path=str(data.get("xml_path", "")),
        width=float(data.get("width", 11.0) or 11.0),
        height=float(data.get("height", 8.5) or 8.5),
        shapes=[_shape_from_dict(item) for item in data.get("shapes", [])],
        connections=[_connection_from_dict(item) for item in data.get("connections", [])],
        workcenter_counts={
            str(key): int(value) for key, value in data.get("workcenter_counts", {}).items()
        },
        candidate_counts={
            str(key): int(value) for key, value in data.get("candidate_counts", {}).items()
        },
        candidate_wcs={
            str(key): set(str(value) for value in values)
            for key, values in data.get("candidate_wcs", {}).items()
        },
        suggested_department=str(data.get("suggested_department", "")),
        assigned_department=str(data.get("assigned_department", "")),
        assignment_mode=str(data.get("assignment_mode", "auto") or "auto"),
        confidence=float(data.get("confidence", 0.0) or 0.0),
        matched_wcs=[str(value) for value in data.get("matched_wcs", [])],
        unmatched_wcs=[str(value) for value in data.get("unmatched_wcs", [])],
        loaded=bool(data.get("loaded", False)),
        loading=False,
        load_error=str(data.get("load_error", "")),
        loaded_plant=str(data.get("loaded_plant", "")),
        change_rule_draft=[ChangeRuleDraftRecord(**item) for item in data.get("change_rule_draft", [])],
        change_rule_generated_at=str(data.get("change_rule_generated_at", "")),
        change_rule_rerun_at=str(data.get("change_rule_rerun_at", "")),
        change_rule_source=str(data.get("change_rule_source", "")),
        change_rule_generation_count=int(data.get("change_rule_generation_count", 0) or 0),
    )
    if page.shapes:
        page.loaded = True
    return page


def save_analysis_package(
    output_path: str,
    files: Sequence[FileRecord],
    workbook_mapping_path: str = "",
    manual_department_plant_text: str = "",
    routing_template_path: str = "",
    logs: Optional[Sequence[str]] = None,
    app_version: str = "1.3.8",
) -> str:
    if not files:
        raise AnalysisPackageError("There are no VSDX files to save in the analysis package.")

    output = Path(output_path)
    if output.suffix.lower() != PACKAGE_EXTENSION:
        output = output.with_suffix(PACKAGE_EXTENSION)
    output.parent.mkdir(parents=True, exist_ok=True)

    manifest: Dict[str, Any] = {
        "schema_version": PACKAGE_SCHEMA,
        "app_version": app_version,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "manual_department_plant_member": "",
        "workbook_mapping_member": "",
        "routing_template_member": "",
        "files": [],
        "logs": list(logs or []),
    }

    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as package:
        used_names = set()
        for index, file_record in enumerate(files, start=1):
            source = Path(file_record.path)
            if not source.exists():
                raise AnalysisPackageError(
                    f"Cannot save package because the selected VSDX file is missing:\n{source}"
                )

            base = _safe_member_name(source.name, f"file_{index:03d}.vsdx")
            member = f"vsdx/{index:03d}_{base}"
            while member.lower() in used_names:
                member = f"vsdx/{index:03d}_{len(used_names):03d}_{base}"
            used_names.add(member.lower())
            package.write(source, member)

            manifest["files"].append(
                {
                    "member": member,
                    "original_path": str(source),
                    "original_name": source.name,
                    "plant": file_record.plant,
                    "status": file_record.status,
                    "error": file_record.error,
                    "pages": [_page_to_dict(page) for page in file_record.pages],
                }
            )

        mapping_source = Path(workbook_mapping_path) if workbook_mapping_path else None
        if mapping_source and mapping_source.exists() and mapping_source.is_file():
            mapping_member = f"mapping/{_safe_member_name(mapping_source.name, 'workcentre_mapping.xlsx')}"
            package.write(mapping_source, mapping_member)
            manifest["workbook_mapping_member"] = mapping_member

        if manual_department_plant_text.strip():
            manual_member = "mapping/department_plant_list.txt"
            package.writestr(manual_member, manual_department_plant_text.encode("utf-8"))
            manifest["manual_department_plant_member"] = manual_member

        routing_source = Path(routing_template_path) if routing_template_path else None
        if routing_source and routing_source.exists() and routing_source.is_file():
            routing_member = f"routing/{_safe_member_name(routing_source.name, 'routing_template.xlsx')}"
            package.write(routing_source, routing_member)
            manifest["routing_template_member"] = routing_member

        package.writestr(
            "manifest.json",
            json.dumps(manifest, ensure_ascii=False, indent=2).encode("utf-8"),
        )

    return str(output)


def _package_cache_directory(package_path: Path) -> Path:
    stat = package_path.stat()
    digest = hashlib.sha1(
        f"{package_path.resolve()}|{stat.st_size}|{stat.st_mtime_ns}".encode("utf-8")
    ).hexdigest()[:16]
    root = Path(tempfile.gettempdir()) / "VSDX_Process_Metadata_Analyzer" / "packages" / digest
    root.mkdir(parents=True, exist_ok=True)
    return root


def import_analysis_package(package_path: str) -> ImportedAnalysisPackage:
    source = Path(package_path)
    if not source.exists():
        raise AnalysisPackageError(f"Analysis package not found:\n{source}")

    try:
        with zipfile.ZipFile(source) as package:
            if "manifest.json" not in package.namelist():
                raise AnalysisPackageError("The selected file is not a valid analysis package.")
            manifest = json.loads(package.read("manifest.json").decode("utf-8"))

            schema = int(manifest.get("schema_version", 0) or 0)
            if schema > PACKAGE_SCHEMA:
                raise AnalysisPackageError(
                    f"This package uses schema {schema}, which is newer than this app supports."
                )

            cache = _package_cache_directory(source)
            files: List[FileRecord] = []
            for index, file_data in enumerate(manifest.get("files", []), start=1):
                member = str(file_data.get("member", ""))
                if not member or member not in package.namelist():
                    raise AnalysisPackageError(
                        f"Embedded VSDX file #{index} is missing from the package."
                    )
                original_name = _safe_member_name(
                    str(file_data.get("original_name", Path(member).name)),
                    f"file_{index:03d}.vsdx",
                )
                destination = cache / "vsdx" / f"{index:03d}" / original_name
                destination.parent.mkdir(parents=True, exist_ok=True)
                with package.open(member) as source_stream, destination.open("wb") as target_stream:
                    shutil.copyfileobj(source_stream, target_stream, length=1024 * 1024)

                files.append(
                    FileRecord(
                        path=str(destination),
                        plant=str(file_data.get("plant", "")),
                        pages=[_page_from_dict(item) for item in file_data.get("pages", [])],
                        status=str(file_data.get("status", "Imported")),
                        error=str(file_data.get("error", "")),
                    )
                )

            workbook_mapping_path = ""
            mapping_member = str(manifest.get("workbook_mapping_member", ""))
            if mapping_member and mapping_member in package.namelist():
                destination = cache / "mapping" / Path(mapping_member).name
                destination.parent.mkdir(parents=True, exist_ok=True)
                with package.open(mapping_member) as source_stream, destination.open("wb") as target_stream:
                    shutil.copyfileobj(source_stream, target_stream, length=1024 * 1024)
                workbook_mapping_path = str(destination)

            manual_text = ""
            manual_member = str(manifest.get("manual_department_plant_member", ""))
            if manual_member and manual_member in package.namelist():
                manual_text = package.read(manual_member).decode("utf-8-sig", errors="replace")

            routing_template_path = ""
            routing_member = str(manifest.get("routing_template_member", ""))
            if routing_member and routing_member in package.namelist():
                destination = cache / "routing" / Path(routing_member).name
                destination.parent.mkdir(parents=True, exist_ok=True)
                with package.open(routing_member) as source_stream, destination.open("wb") as target_stream:
                    shutil.copyfileobj(source_stream, target_stream, length=1024 * 1024)
                routing_template_path = str(destination)

            return ImportedAnalysisPackage(
                files=files,
                workbook_mapping_path=workbook_mapping_path,
                manual_department_plant_text=manual_text,
                routing_template_path=routing_template_path,
                package_path=str(source.resolve()),
                created_at=str(manifest.get("created_at", "")),
                app_version=str(manifest.get("app_version", "")),
            )
    except zipfile.BadZipFile as exc:
        raise AnalysisPackageError(f"Invalid analysis package: {exc}") from exc
    except json.JSONDecodeError as exc:
        raise AnalysisPackageError(f"Invalid package manifest: {exc}") from exc
