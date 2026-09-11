# Changelog

## 1.3.8

- Simplified the Step 2 panel titles to **Active Plant mapping** and **VSM files and plant assignment**, and expanded the active mapping table to show more records without scrolling.
- Added completed/dated VSM page-name inference when a Plant mapping is incomplete, with stable snake_case keys such as `bearing_assembly` and `fab_lining`.
- Added alias reconciliation so descriptive page names retain an existing business CRID when one is available, plus visible suggestion provenance in Step 2.
- Excluded pages marked `Old` or `Combined into` from page-name inference and added automated regression tests for CA06-style detection.
- Added a gated **Step 0 — Data sources** for the raw P41 Routing extraction, PTS03 WorkCenter mapping, and MRP/Production Supervisor reference.
- Added workbook-tab discovery and an explicit **Confirm & Load** action so no raw sheet is consumed immediately after file browsing.
- Locked VSM file and folder upload until all three source contracts pass validation.
- Replaced the prebuilt Variant As-Is dependency with direct `MATERIAL_IN_SCOPE`, `MAPL`, and `PLPO` processing.
- Added MRP and Production Supervisor validation and retained missing/ambiguous records as visible review items.
- Added the CA02 Calendar first-round draft assumption and exposed its review note in the Material Selection table.
- Changed repeated OLD WC branches with multiple NEW WC targets from `ALL` to `LIST` Change Rule type.
- Moved large P41 loading to a background worker so the desktop interface remains responsive.

## 1.3.5

- Reworked the **VSM tabs / pages** action row so **Selected Tab** and **Regenerate All** retain their full labels in narrow three-pane layouts.
- Increased usable height for **2. VSDX files and plant assignment** by keeping the input-source card compact and reserving a seven-row file-table area.
- Renamed the app to **VSM to ROUTING Preparation Tool v.1.3.5** and added the footer credit: **Software by COA Data Team Copyright 2026**.
- Corrected the rounded-button renderer so it draws a single exterior border only; internal construction lines no longer appear within buttons.
- Replaced the desktop application's standard rectangular buttons with a custom rounded-button renderer: 13-pixel curved corners, soft hover states, clear blue primary actions, and readable pale-blue disabled actions.
- Disabled **Run metadata analysis** until at least one VSDX is loaded and every loaded VSDX has a Plant assignment, preventing an empty or invalid full regeneration run.

## 1.3.4

- Switched the desktop widgets to Tk's controllable `clam` theme because the native Windows theme ignored disabled-button colour maps. Disabled actions now render with a visible pale-blue fill and slate text; enabled primary actions render solid blue with white text.
- Refreshed the desktop interface with the selected Option 3 blue-and-indigo workspace: light blue canvas, white cards, stronger typography, higher-contrast tables, and raised primary/secondary action buttons.
- Replaced **Analyze VSM** in **VSM tabs / pages** with two clear actions: **Selected Tab** refreshes only the highlighted VSM, while **Regenerate All** runs the established analysis across every loaded VSDX tab.
- Preserved the existing routing-generation and metadata-analysis logic; this release changes the interaction model and visual system, not routing calculations.

## 1.3.3

- Replaced the long Plant dropdown with direct mapping-driven Plant buttons (for example CA02, CA06, CA15), plus **Other…** for an uncommon code.
- Renamed **Detected tabs / pages** to **VSM tabs / pages** and added **Analyze VSM** for the highlighted tab.
- Step 2 now shows a filterable active-Plant mapping list (CRID, workcenter, source) in place of the mapping-file editor; the editor remains in Step 1 and Advanced workspace.
- Moved routing exports to their own always-visible action row; **Export Change Rule** no longer clips in narrow panes.
- Kept generation and export controls out of the Assign context stage so that stage focuses only on Plant, tab and CRID confirmation.
- Replaced the all-at-once primary layout with a six-stage manufacturing workflow wizard: Load inputs, Assign context, Inspect process, Generate draft, Review readiness, and Export package.
- Each stage shows only its relevant panels and is gated by the previous production condition.
- Added Previous/Next stage controls and an Advanced workspace escape hatch for the full expert layout.
- Preserved all v1.3.2 UI hardening and all routing-generation business logic.

## 1.3.2

- Hardened selected-file and selected-tab context: the header now continuously shows Plant, VSDX, tab, CRID, draft rows and readiness.
- Strengthened Treeview selection contrast so selected VSDX and draft rows remain legible when focus changes.
- Added explicit unsaved/saved Plant state and disabled Plant Apply until a selected-file value changes.
- Added detected-tab status, routing status/search filters, sortable draft columns, alternating draft rows and readiness-row filtering.
- Replaced the three rerun buttons with a policy-aware **Rerun** menu.
- Added preview zoom presets, display toggles, keyboard shortcuts and persisted UI geometry/workspace/zoom preferences.
- Preserved routing-generation calculations and all v1.3.1 workflows.

## 1.3.1

- Added **Generate Selected CRID**.
- Added **Generate All CRIDs for Plant**.
- Added **Generate All Loaded Plants**.
- Added Department/CRID filter with **ALL DEPARTMENTS** view.
- Added **Rerun Selected CRID**, **Rerun Selected Plant** and **Rerun Everything**.
- Added rerun policies: preview changes, replace existing draft, merge/preserve manual edits.
- Added before/after rerun comparison window.
- Added Department/CRID readiness dashboard.
- Added Plant-level and all-Plants consolidated routing-generation exports.
- Added rerun history and generation timestamps.
- Updated `.vpa` schema to preserve rerun metadata.
- Sequence numbering remains independent per CRID and starts at `SEQ1`.
- Preserved all v1.3 and v1.2.1 analyzer/routing preparation capabilities.

## 1.3.0

- Renamed product to VSDX Routing-Generation Preparation Tool.
- Added Change Rule Draft workspace.
- Added 10-field routing-generation draft schema.
- Added PTS03 OLD WC lookup.
- Added Visio process sequencing and branch/subsequence preparation.
- Added editable draft grid and readiness validation.
- Added clean Change Rule export and template-population export.
- Added draft persistence in `.vpa` analysis packages.
- Bundled routing-generation reference/template workbook.
- Preserved all v1.2.1 analyzer capabilities.
