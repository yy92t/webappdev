# Google Apps Script: Flexible Pivot Table Builder

This script creates pivot tables in Google Sheets using the Advanced Sheets API. It accepts flexible column references (header names, letters, or numeric indices) and can be pointed at any tabular data range.

## Prerequisites

1. Open Extensions > Apps Script on your Google Sheet.
2. Add this project as a bound script.
3. Enable Advanced Google Services:
   - In Apps Script: Services (left sidebar) > + > Google Sheets API > Add.
   - Ensure the Google Sheets API is also enabled in your Cloud project if prompted.
4. Paste the `PivotTableBuilder.gs` content into the editor and save.

## Quick Start

1. Prepare a sheet named `Data` with a header row (e.g., columns like `ID, Region, Status, Amount`).
2. Create a blank sheet (or let the script create it) named `Pivot`.
3. Run `createExamplePivot()` from the script editor or use the custom menu "Pivot Tools > Create Example Pivot".

The example:
- Rows: Region
- Columns: Status
- Values: Sum of Amount, Count of ID
- Output to sheet "Pivot" at A1.

## Build Your Own Pivot

Use `createPivotTable(config)` with a configuration like:

```javascript
createPivotTable({
  sourceSheet: 'RawData',
  sourceRangeA1: 'A1:Z', // must include header row
  destinationSheet: 'MyPivot',
  anchorCellA1: 'B2', // optional, default A1

  rows: [
    { column: 'Product', showTotals: true, sortOrder: 'ASC' }
  ],
  columns: [
    { column: 'Quarter', showTotals: true, sortOrder: 'DESC' }
  ],
  values: [
    { column: 'Revenue', summarizeFunction: 'SUM', name: 'Revenue (Sum)' },
    { column: 'OrderID', summarizeFunction: 'COUNT', name: 'Orders' }
  ]
});
```

### Column References

For `rows`, `columns`, and `values`, the `column` field accepts:
- Header name: `'Revenue'`
- Column letter: `'C'` or `'AA'`
- 1-based column index: `3`

The reference must point within `sourceRangeA1`.

### Supported summarizeFunction values

- SUM, COUNTA, COUNT, MAX, MIN, AVERAGE, MEDIAN, PRODUCT, STDEV, STDEVP, VAR, VARP, CUSTOM

If you use `CUSTOM`, provide a `formula` in the value spec.

### Grouping and Sorting

- `sortOrder`: `'ASC'` or `'DESC'`
- `showTotals`: defaults to `true`
- You can pass a `groupRule` on a row/column group if you need date or histogram bucketing (advanced Sheets API object).

## Notes

- This script uses the Sheets Advanced Service via `Sheets.Spreadsheets.batchUpdate` with `updateCells` and a `pivotTable` cell at the anchor position.
- Filters and more advanced pivot configurations can be added by extending the request object.
- The pivot table is created at the specified anchor cell on the destination sheet. You can create multiple pivots by selecting different anchors.

## Troubleshooting

- ReferenceError: "Sheets is not defined"
  - Ensure you enabled the Advanced Google Sheets API service in the Apps Script editor (Services > + > Google Sheets API).

- Column reference outside range
  - The column reference must be within the `sourceRangeA1`. Expand the range or adjust the column reference.

- Headers not found
  - Make sure your `sourceRangeA1` includes the header row and that header names match exactly (case-sensitive).
