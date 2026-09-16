# Excel Department & Employment Status Splitter

This application lets you select an Excel workbook, choose a worksheet, map the Department and ID Number / Employment Status columns, review detected groups, and generate a new workbook with one worksheet per Department + ID Number combination.

## Features

- Browse for an Excel workbook (.xlsx, .xlsm, .xls)
- View available worksheets
- Detect and select the Department and ID Number columns dynamically
- Preview all detected Department + ID Number groups
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

The app reads each row in the selected worksheet and creates a grouping key using:

```text
(department_value, id_number_value)
```

If a Department or ID Number value is empty, it is replaced with:

- `UNKNOWN DEPARTMENT`
- `UNKNOWN STATUS`

Each unique combination becomes a separate worksheet with the original headers and matching records.
