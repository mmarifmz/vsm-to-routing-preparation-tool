from __future__ import annotations

import argparse
import json
from pathlib import Path

from app.exporter import export
from app.gui import launch
from app.mapping import (
    load_department_plant_file,
    load_mapping,
    merge_mappings,
)
from app.vsdx import analyze


def smoke(args):
    logs = []
    workbook_mapping = (
        load_mapping(args.mapping, logs.append)
        if args.mapping
        else None
    )
    manual_mapping = (
        load_department_plant_file(args.department_list)[0]
        if args.department_list
        else None
    )
    mapping = merge_mappings(
        workbook_mapping, manual_mapping
    )

    result = analyze(
        args.vsdx,
        args.plant,
        mapping,
        logs.append,
    )
    output = args.output or str(
        Path(args.vsdx).with_name(
            Path(args.vsdx).stem + "_metadata.xlsx"
        )
    )
    export(output, [result], mapping, logs)

    print(
        json.dumps(
            {
                "file": result.file_name,
                "plant": result.plant,
                "pages": len(result.pages),
                "shapes": sum(
                    len(page.shapes)
                    for page in result.pages
                ),
                "connections": sum(
                    len(page.connections)
                    for page in result.pages
                ),
                "workcenters": sorted(
                    {
                        workcenter
                        for page in result.pages
                        for workcenter
                        in page.workcenter_counts
                    }
                ),
                "department_options": (
                    mapping.departments_for_plant(
                        args.plant
                    )
                    if mapping
                    else []
                ),
                "output": output,
                "tabs": [
                    {
                        "name": page.name,
                        "shapes": len(page.shapes),
                        "workcenters": list(
                            page.workcenter_counts
                        ),
                        "suggested_department": (
                            page.suggested_department
                        ),
                        "confidence": round(
                            page.confidence, 4
                        ),
                    }
                    for page in result.pages
                ],
            },
            indent=2,
        )
    )
    return 0


def main():
    parser = argparse.ArgumentParser(
        description=(
            "VSM to ROUTING Preparation Tool v.1.3.8"
        )
    )
    parser.add_argument("--smoke-test", action="store_true")
    parser.add_argument("--vsdx")
    parser.add_argument("--mapping")
    parser.add_argument(
        "--department-list",
        help="Optional Department | Plant TXT/CSV file.",
    )
    parser.add_argument("--plant", default="")
    parser.add_argument("--output")
    args = parser.parse_args()

    if args.smoke_test:
        if not args.vsdx or not args.plant:
            parser.error(
                "--smoke-test requires --vsdx and --plant"
            )
        return smoke(args)

    launch()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
