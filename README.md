# Excel Department & Employment Status Splitter

This application lets you select an Excel workbook, choose a worksheet, select one or more grouping columns, review detected groups, and generate a new workbook with one worksheet per unique grouping combination.

## Features

- Browse for Excel workbooks (.xlsx, .xlsm, .xls) or CSV files (.csv)
- Multi-sheet workbooks open with `All sheets` selected; choose an individual worksheet when needed
- Combine all worksheets across multiple uploaded workbooks; CSV files are treated as a single worksheet
- Select one grouping column by default and add more grouping columns as needed
- Preview all detected groups for the selected column combination
- Review possible spelling and capitalization matches in selected grouping columns and choose whether to merge them
- Review possible duplicate records by selected matching columns, with their sources and original values
- Search for and manually merge any two grouping values, with an optional preferred name
- Optionally write chosen merge spellings into exported grouping cells
- Preview selected output worksheets by name and inspect their table rows before generating the workbook
- Sort detected groups alphabetically or by record count
- Browse detected groups in pages, keeping selections when changing pages or sort order
- Uncheck groups you do not want to generate
- Merge matching groups across worksheets, or turn off `Combine duplicate groups` to preview and export each source separately
- Sanitize sheet names for Excel restrictions
- Preserve source rows and headers in each generated sheet
- Format generated sheets like the FULL DATAS attendance tables: dark blue headers, white bold text, Arial 11, alternating blue rows, fitted column widths, and frozen headers and leading columns
- Generate the workbook first, then choose the output location in a Save As dialog

## Install dependencies

```bash
python -m pip install -r requirements.txt
```

## Run the application

```bash
python main.py
C:/Python314/python.exe main.py
```

## Packaging into a Windows .exe

Install PyInstaller:

```bash
python -m pip install pyinstaller
```

Create the executable:

```bash
python -m PyInstaller --clean main.spec
```

The executable will be generated as `dist/Excel Arranger.exe`, with the app icon embedded and included for its window.

## Grouping logic

The app reads each row in the selected worksheet and creates a grouping key using the selected columns:

```text
(value_from_column_1, value_from_column_2, ...)
```

If a selected grouping value is empty, it is replaced with:

- `UNKNOWN GROUP`

Each unique combination becomes a separate worksheet with the original headers and matching records. Leaving only the default grouping column selected creates one sheet per value in that column.

CSV files are read as single-sheet tabular data and are exported to an Excel workbook using the same grouping process. Excel and CSV files can be uploaded together; Excel worksheets and CSV files are all combined when `All sheets` is selected.

CSV values are imported as text so identifiers with leading zeros and values such as `NA` are preserved. Excel cells retain their stored value types. Duplicate column headers are rejected, and text beginning with `=` is written as text.

Worksheets containing formulas are rejected because splitting can change cell references or remove referenced sheets. To split their current results, make a copy and replace formulas with values first. The app checks all worksheets when loading a workbook in the interface.

Selecting `All sheets` combines rows from every worksheet in every selected Excel workbook before grouping. This is the default for workbooks containing multiple worksheets and for multiple uploaded files. Different worksheets and files may have different columns or column order. The output uses the first-seen union of headers, and missing values are left blank. Empty worksheets and blank rows are skipped.

Click **Review duplicate records** to compare records in the selected worksheets. Choose one or more matching columns and click **Find matches**; columns shared by all selected sources are selected initially. Text matching ignores capitalization and extra whitespace, and records with a blank or missing matching value are skipped. Results show each matching record, source, and position in the loaded table (which may differ from the original file line when blank lines are skipped); select a result to inspect its values. This review does not remove or change records and is independent of **Combine duplicate groups**.

With `Combine duplicate groups` enabled, matching groups from different worksheets or files are merged. Disable it to keep each source worksheet's groups separate. The preview includes the source workbook and worksheet names, and selected preview groups are used for export. Loaded worksheets are cached for both preview and export; reload a file to pick up changes made after it was opened.

Detected groups can be sorted alphabetically, by record count ascending, or by record count descending. The selected order is also used for the generated worksheets.

Generated `.xlsx` worksheets include filterable Excel tables using `TableStyleMedium2`, 31-point header rows, and 23-point data rows. Known attendance columns use widths matched by header name, so reordered columns keep the appropriate width. Other columns fit their contents up to a readable maximum. Legacy `.xls` exports use equivalent cell colors, fonts, row heights, and frozen panes.

## Review spelling suggestions

After selecting the grouping columns (for example, add `ID Number`), click **Review spelling**. The app lists possible matches for all selected columns, with the record count for each spelling. Choose either spelling for a combined group or keep each pair separate. Use **Undo last choice** to reverse decisions, then apply your choices or cancel the review. You can reopen the review to undo applied choices while the grouping and group selection are unchanged. Suggestions are based on similarity, so review each one before accepting it; only values already present in the selected data are suggested. Numeric and alphanumeric identifiers are excluded.

Approved choices apply to the group preview and generated worksheet names for the current worksheet selection. By default, exported rows retain their original cell values. Select **Correct merged values in output** to write the chosen spelling into the corresponding grouping cells in the generated workbook. Other cells remain unchanged. Applied spelling choices are saved locally for the same grouping column and reused in future workbooks, even when the preferred spelling is absent. Click **Saved spellings** to review or remove them; removal affects future workbooks. Undoing an applied choice and applying the review also removes that saved choice. Manual merges remain limited to the current worksheet selection.

To merge values that are not suggested, click **Manual merge**, choose the grouping column, and search for and select one value in each list. Enter a preferred name if needed, or leave it blank to use the first value, then click **Merge**. To use an existing value as the preferred name, select it as one of the two values. These merges also honor **Correct merged values in output**; the input files are not modified.
