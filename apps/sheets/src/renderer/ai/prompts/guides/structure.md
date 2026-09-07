# Structural changes guide (structure)

## Operation definitions

- `{op:"insert_rows", sheetId, row:2, count:1}` — row is 1-based; new rows are inserted **before** that row.
- `{op:"delete_rows", sheetId, row, count}` — deletes count rows starting at row.
- `{op:"insert_cols", sheetId, column:"C", count}` — new columns are inserted before that column; column uses the column letter.
- `{op:"delete_cols", sheetId, column, count}`
- `{op:"add_sheet", name}` — creates a blank worksheet (no sheetId needed).
- `{op:"delete_sheet", sheetId}` — deletes a worksheet; the last remaining sheet cannot be deleted. Formulas on other sheets referencing the deleted sheet become #REF! (listed in the preview warnings); for real xlsx files there is an additional fail-closed guard on save (sheets referenced by formulas/charts/defined names refuse deletion).
- `{op:"duplicate_sheet", sheetId, name?}` — copies a whole worksheet (contents and formats); name is auto-generated when omitted. Unavailable in large-file streaming mode; sheets containing pivot tables cannot be duplicated.
- `{op:"set_sheet_hidden", sheetId, hidden:true|false}` — hides/shows a worksheet; at least one sheet must stay visible.
- `{op:"move_sheet", sheetId, position:1}` — moves to the 1-based tab position.

## Multi-sheet workbook organization

When a task legitimately needs multiple sheets (user asked for it, or the data model clearly requires it), plan the sheet set BEFORE creating any, and order tabs so the workbook reads like a report:

1. **Ringkasan/Summary/Dashboard first** (tab 1) — the executive view: key figures, conclusions, navigation notes. Keep it small and formula-linked to the detail sheets (never hard-code numbers that exist elsewhere).
2. **Input/Data sheets next** — raw or source data, one domain per sheet, named clearly (e.g. "Data Penjualan", "Data Karyawan").
3. **Analysis/calculation sheets after the data** — pivot summaries, per-unit computations, working papers that feed the summary.
4. **Reference/lookup sheets last** (master lists, assumptions, parameters).

Rules:
- Create sheets in reading order (Summary → Data → Analysis → Reference) so no move_sheet cleanup is needed; if order ends up wrong, fix with `{op:"move_sheet"}` at the end.
- Every non-obvious sheet benefits from a one-line title in A1 (bold, larger font) stating what it contains.
- Cross-sheet formulas are fine and preferred over duplicating data; reference the Data sheet, never copy its values into the Analysis sheet.
- Don't over-split: a small workbook (≤ ~50 rows total) usually fits better in one sheet with clearly separated sections than in 4+ tabs.

## Sheet protection

`{op:"protect_sheet", sheetId, protected:true|false}` — **layout-class**, can share a batch with content/format operations (exempt from the batching discipline below). Written into the file on save (passwords not supported); password-protected sheets cannot be unprotected. The editor itself does not enforce the lock.

## Mandatory batching discipline

Structural operations move cell addresses and **cannot appear in the same batch as content/format/sort-layout operations**. The correct rhythm:

1. Submit the structural changes alone first (multiple structural ops may share a batch, applied in order, effective on submit)
2. Call get_workbook_context / read_range again for the shifted layout
3. Then submit the content/format changes

## Formula references

- On row/column insertion/deletion, all A1 references in the workbook's formulas are rewritten automatically for the shift (including absolute $ references and cross-sheet prefixed references).
- Formulas whose reference target is deleted entirely become **#REF!** — the preview warnings list the affected cells precisely. When this happens, ask the user to confirm, or proactively suggest a fix.
- Range references (like SUM(B2:B10)) shrink automatically on partial deletion — you don't need to repair formulas by hand.

## Common mistakes

- ❌ insert_rows followed by writing into the "new rows" in the same batch — rejected by the mixed-batch rule, and the new rows' addresses don't exist before apply.
- ❌ Continuing to write with old addresses right after a structural change applies — re-read the context first.
- ❌ Using delete_rows to clear data — when you only want to clear contents, use clear_range (structure unchanged, formula references unaffected).
