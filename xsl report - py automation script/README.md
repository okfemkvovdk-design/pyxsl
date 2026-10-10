# Excel Data Control

This script cleans an Excel workbook, reports duplicate rows, and can answer a
few quick questions about a worksheet. It is intended for `.xlsx` and `.xlsm`
files.

## Install

From this folder, install the required package:

```powershell
python -m pip install -r requirements.txt
```

## Clean a workbook

Run the script with the workbook path:

```powershell
python excel_data_control.py "input.xlsx"
```

The script saves two files beside the original:

- `input_cleaned.xlsx` has fully blank rows removed.
- `input_duplicates.xlsx` lists duplicate rows. Duplicate rows stay in the
  cleaned workbook.

For a macro-enabled `.xlsm` workbook, the cleaned output is also `.xlsm`.

To save the results somewhere else:

```powershell
python excel_data_control.py "input.xlsx" `
  --cleaned-output "output\cleaned.xlsx" `
  --duplicates-output "output\duplicates.xlsx"
```

Duplicate matching compares the complete row values within each worksheet.
The report gives the worksheet, row number, first matching row, group size, and
row values. Duplicate row numbers refer to the cleaned workbook.

If blank rows are removed from a workbook that contains formulas or tables, the
script displays a warning. Review formulas and table ranges in the cleaned
file: some workbook references are not automatically adjusted when rows are
deleted.

## Ask a question about the data

Add `--ask` to ask questions after the files are created:

```powershell
python excel_data_control.py "input.xlsx" --ask
```

Choose a worksheet, then enter a question. The first row is treated as the
column headings. Examples:

- `average Sales`
- `sum Sales`
- `min Sales`
- `max Sales`
- `count Status`
- `count Status=Paid`
- `rows`

Average, sum, minimum, and maximum use numeric cells only. `count Status`
counts non-empty cells in the column; `count Status=Paid` counts exact matches
without regard to capitalization. `rows` counts non-empty data rows, excluding
the heading row. Question mode reads the original workbook so it can use saved
formula results when they are available. The script does not recalculate
formulas. Value matching ignores capitalization and surrounding spaces. Enter
`quit` to leave question mode.
