# Excel Department & Employment Status Splitter

This application lets you select an Excel workbook, choose a worksheet, select one or more grouping columns, review detected groups, and generate a new workbook with one worksheet per unique grouping combination.

## Features

- Browse for Excel workbooks (.xlsx, .xlsm, .xls) or CSV files (.csv)
- View available worksheets or choose `All sheets`; CSV files are treated as a single worksheet
- Select one grouping column by default and add more grouping columns as needed
- Preview all detected groups for the selected column combination
- Sort detected groups alphabetically or by record count
- Uncheck groups you do not want to generate
- Sanitize sheet names for Excel restrictions
- Preserve source rows and headers in each generated sheet
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
pyinstaller --onefile --windowed main.py
```

The executable will be generated in the `dist` folder.

## Grouping logic

The app reads each row in the selected worksheet and creates a grouping key using the selected columns:

```text
(value_from_column_1, value_from_column_2, ...)
```

If a selected grouping value is empty, it is replaced with:

- `UNKNOWN GROUP`

Each unique combination becomes a separate worksheet with the original headers and matching records. Leaving only the default grouping column selected creates one sheet per value in that column.

CSV files are read as single-sheet tabular data and are exported to an Excel workbook using the same grouping process. Excel and CSV files can be uploaded together; Excel worksheets and CSV files are all combined when `All sheets` is selected.

Selecting `All sheets` combines rows from every worksheet in every selected Excel workbook before grouping. Different files may have different columns or column order. The output uses the first-seen union of headers, and missing values are left blank. Empty worksheets and files are skipped.

Detected groups can be sorted alphabetically, by record count ascending, or by record count descending. The selected order is also used for the generated worksheets.
