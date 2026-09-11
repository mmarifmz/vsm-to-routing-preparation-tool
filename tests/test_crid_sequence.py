from __future__ import annotations

import unittest

from app.models import ConnectionRecord, PageRecord, ShapeRecord
from app.routing import ChangeRuleRow, generate_change_rule_draft, validate_change_rule_rows


def shape(shape_id: str, x: float, y: float, workcenter: str) -> ShapeRecord:
    return ShapeRecord(
        shape_id=shape_id,
        text=workcenter,
        pin_x=x,
        pin_y=y,
        workcenters=[workcenter],
    )


def process(shape_id: str, x: float, y: float, text: str) -> ShapeRecord:
    return ShapeRecord(
        shape_id=shape_id,
        text=text,
        pin_x=x,
        pin_y=y,
        width=1.0,
        height=0.75,
    )


class CridSequenceTests(unittest.TestCase):
    def test_aligned_operations_share_sequence_and_increment_subsequence(self):
        page = PageRecord(
            "1",
            "Department - complete",
            "Department - complete",
            "page.xml",
            shapes=[
                shape("1", 10.0, 8.0, "WC001"),
                shape("2", 20.0, 8.0, "WC002"),
                shape("3", 20.0, 6.0, "WC003"),
                shape("4", 30.0, 8.0, "WC004"),
            ],
        )

        rows = generate_change_rule_draft(page, "CA06", "DEPT_A")

        self.assertEqual(
            [(row.sequence, row.subseq) for row in rows],
            [(1, 1), (2, 1), (2, 2), (3, 1)],
        )
        self.assertEqual([row.new_wc for row in rows], ["WC001", "WC002", "WC003", "WC004"])
        self.assertEqual([row.column for row in rows], ["SEQ1", "SEQ2", "SEQ2", "SEQ3"])

    def test_sequence_restarts_for_another_crid(self):
        page = PageRecord(
            "1",
            "Department - complete",
            "Department - complete",
            "page.xml",
            shapes=[shape("1", 10.0, 8.0, "WC001")],
        )

        first = generate_change_rule_draft(page, "CA06", "DEPT_A")
        second = generate_change_rule_draft(page, "CA06", "DEPT_B")

        self.assertEqual(first[0].sequence, 1)
        self.assertEqual(second[0].sequence, 1)

    def test_connector_flow_orders_a_snaking_department_process(self):
        stages = [
            process("p1", 2.0, 6.0, "Stage"),
            process("p2", 7.0, 6.0, "Assembly"),
            process("p3", 6.0, 3.0, "Test"),
            process("p4", 2.0, 3.0, "Tag"),
            process("p5", 3.0, 1.0, "Inspect"),
            process("p6", 5.0, 1.0, "Pack"),
        ]
        labels = [
            shape("w1", 2.0, 5.52, "WC001"),
            shape("w2", 7.0, 5.52, "WC002"),
            shape("w3", 6.0, 2.52, "WC003"),
            shape("w4", 2.0, 2.52, "WC004"),
            shape("w5", 3.0, 0.52, "WC005"),
            shape("w6", 5.0, 0.52, "WC006"),
        ]
        loose_connector = ShapeRecord(
            shape_id="c4",
            kind="connector",
            begin_x=2.5,
            begin_y=3.0,
            end_x=2.5,
            end_y=1.0,
        )
        page = PageRecord(
            "1",
            "Valves - complete",
            "Valves - complete",
            "page.xml",
            shapes=stages + labels + [loose_connector],
            connections=[
                ConnectionRecord("c1", "p1", "p2"),
                ConnectionRecord("c2", "p2", "p3"),
                ConnectionRecord("c3", "p3", "p4"),
                ConnectionRecord("c4", "p4", ""),
                ConnectionRecord("c5", "p5", "p6"),
            ],
        )

        rows = generate_change_rule_draft(page, "CA06", "VALVES")

        self.assertEqual([row.new_wc for row in rows], [f"WC00{i}" for i in range(1, 7)])
        self.assertEqual([(row.sequence, row.subseq) for row in rows], [(i, 1) for i in range(1, 7)])

    def test_same_workcenter_at_two_process_positions_is_not_discarded(self):
        page = PageRecord(
            "1",
            "Department - complete",
            "Department - complete",
            "page.xml",
            shapes=[
                shape("1", 10.0, 8.0, "WC001"),
                shape("2", 20.0, 8.0, "WC001"),
            ],
        )

        rows = generate_change_rule_draft(page, "CA06", "DEPT_A")

        self.assertEqual([(row.sequence, row.new_wc) for row in rows], [(1, "WC001"), (2, "WC001")])

    def test_duplicate_validation_is_scoped_to_crid(self):
        rows = [
            ChangeRuleRow("DEPT_A", "CA06", sequence=1, subseq=1, old_wc="OLD1", new_wc="NEW1", op_descriptions="One", column="SEQ1"),
            ChangeRuleRow("DEPT_B", "CA06", sequence=1, subseq=1, old_wc="OLD2", new_wc="NEW2", op_descriptions="Two", column="SEQ1"),
        ]

        validated = validate_change_rule_rows(rows)

        self.assertFalse(any("Duplicate SEQUENCE/SUBSEQ" in row.notes for row in validated))


if __name__ == "__main__":
    unittest.main()
