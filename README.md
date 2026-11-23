# Excel Department & Employment Status Splitter

This application lets you select an Excel workbook, choose a worksheet, select one or more grouping columns, review detected groups, and generate a new workbook with one worksheet per unique grouping combination.

## Features

- Browse for Excel workbooks (.xlsx, .xlsm, .xls) or CSV files (.csv)
- Multi-sheet workbooks open with `All sheets` selected; choose an individual worksheet when needed
- Combine all worksheets across multiple uploaded workbooks; CSV files are treated as a single worksheet
- Select one grouping column by default and add more grouping columns as needed
- Preview all detected groups for the selected column combination
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

With `Combine duplicate groups` enabled, matching groups from different worksheets or files are merged. Disable it to keep each source worksheet's groups separate. The preview includes the source workbook and worksheet names, and selected preview groups are used for export. Loaded worksheets are cached for both preview and export; reload a file to pick up changes made after it was opened.

Detected groups can be sorted alphabetically, by record count ascending, or by record count descending. The selected order is also used for the generated worksheets.

Generated `.xlsx` worksheets include filterable Excel tables using `TableStyleMedium2`, 31-point header rows, and 23-point data rows. Known attendance columns use widths matched by header name, so reordered columns keep the appropriate width. Other columns fit their contents up to a readable maximum. Legacy `.xls` exports use equivalent cell colors, fonts, row heights, and frozen panes.
