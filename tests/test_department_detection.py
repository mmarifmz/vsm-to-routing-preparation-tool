from __future__ import annotations

import unittest

from app.models import MappingIndex, MappingRecord, PageRecord, ShapeRecord, department_slug
from app.vsdx import apply_mapping, page_department_candidate


class DepartmentDetectionTests(unittest.TestCase):
    def test_department_slug_examples(self):
        self.assertEqual(department_slug("Bearing Assembly"), "bearing_assembly")
        self.assertEqual(department_slug("Fab+Lining"), "fab_lining")
        self.assertEqual(
            department_slug("Bearing Assembly - complete 2026-05-05"),
            "bearing_assembly",
        )
        self.assertEqual(
            department_slug("Fab+Rubber 2026-04-21 - Complete"),
            "fab_rubber",
        )
        self.assertEqual(department_slug("Valves - Working 11/09/2026"), "valves")
        self.assertEqual(department_slug("Lining in progress"), "lining")
        self.assertEqual(department_slug("Calendar - May 19, 2026"), "calendar")

    def test_current_page_filter(self):
        self.assertEqual(page_department_candidate("Template"), "")
        self.assertEqual(page_department_candidate("Horizontal Pumps - Old"), "")
        self.assertEqual(
            page_department_candidate("Warman (COMBINED INTO HORIZ PUMPS)"),
            "",
        )
        self.assertEqual(
            page_department_candidate("Horizontal Pumps - 05/12/2026"),
            "horizontal_pumps",
        )

    def test_page_name_breaks_single_department_mapping_collapse(self):
        mapping = MappingIndex(
            records=[MappingRecord("MAC_MOD_MISC", "CA06", "WC100")]
        )
        mapping.rebuild()
        page = PageRecord(
            "1",
            "Bearing Assembly - complete",
            "Bearing Assembly - complete",
            "page1.xml",
            shapes=[ShapeRecord("1", text="WC100")],
        )

        apply_mapping(page, "CA06", mapping)

        self.assertEqual(page.suggested_department, "bearing_assembly")
        self.assertEqual(page.assigned_department, "")
        self.assertEqual(page.suggestion_source, "Tab name")

    def test_existing_business_code_is_preserved(self):
        mapping = MappingIndex(
            records=[MappingRecord("MAC_MOD_MISC", "CA06", "WC100")]
        )
        mapping.rebuild()
        page = PageRecord(
            "1",
            "Machining/Mod/Misc - complete",
            "Machining/Mod/Misc - complete",
            "page1.xml",
            shapes=[ShapeRecord("1", text="WC100")],
        )

        apply_mapping(page, "CA06", mapping)

        self.assertEqual(page.suggested_department, "machining_mod_misc")
        self.assertEqual(page.assigned_department, "")
        self.assertEqual(page.suggestion_source, "Tab name")

    def test_descriptive_page_name_matches_short_business_code(self):
        self.assertEqual(
            page_department_candidate("Hose Build - complete", ["HOSE"]),
            "hose_build",
        )

    def test_workcenter_mapping_does_not_assign_crid(self):
        mapping = MappingIndex(
            records=[MappingRecord("MAC_MOD_MISC", "CA06", "WC100")]
        )
        mapping.rebuild()
        page = PageRecord(
            "1",
            "Valves - Complete",
            "Valves - Complete",
            "page1.xml",
            shapes=[ShapeRecord("1", text="WC100")],
        )

        apply_mapping(page, "CA06", mapping)

        self.assertEqual(page.suggested_department, "valves")
        self.assertEqual(page.assigned_department, "")
        self.assertEqual(page.workcenter_counts, {"WC100": 1})

    def test_confirmed_reference_vocabulary_discovers_workcenter(self):
        page = PageRecord(
            "1",
            "Valves - Complete",
            "Valves - Complete",
            "page1.xml",
            shapes=[ShapeRecord("1", text="WPAK702")],
        )

        apply_mapping(page, "CA06", None, ["WPAK702"])

        self.assertEqual(page.workcenter_counts, {"WPAK702": 1})
        self.assertEqual(page.suggested_department, "valves")
        self.assertEqual(page.assigned_department, "")


if __name__ == "__main__":
    unittest.main()
