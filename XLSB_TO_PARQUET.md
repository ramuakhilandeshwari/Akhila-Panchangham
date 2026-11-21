# XLSB to Parquet: extraction gist

## In brief

[`convert_xlsb_to_parquet.py`](./convert_xlsb_to_parquet.py) reads an Excel
`.xlsb` workbook and writes **one Parquet file per worksheet**, plus a
`manifest.json` that maps each file back to its worksheet name and order.
Cells are exported by position (`column_1`, `column_2`, and so on), and each
record includes its original worksheet row number (`source_row`).

This keeps worksheet boundaries and cell values available for downstream
analysis. It does **not** turn Excel formulas or cross-sheet references into
working links in Parquet. Keep the original workbook if those Excel behaviors
are important.

## Install and run

From PowerShell in this project folder:

```powershell
python -m pip install -r .\requirements.txt
python .\convert_xlsb_to_parquet.py ".\DRUK VIDESHI PANCHANGAMU JSRK LATEST 03.04.2026.xlsb"
```

By default, the output folder is created beside the workbook and named
`<workbook-name>_parquet`. To choose another output folder:

```powershell
python .\convert_xlsb_to_parquet.py ".\workbook.xlsb" --output-dir ".\parquet-output"
```

The script requires Python, `pyxlsb` to read XLSB files, and `pyarrow` to write
Parquet. The versions are listed in [`requirements.txt`](./requirements.txt).

## What the extraction does

1. Opens the workbook and enumerates its worksheets in workbook order.
2. Reads each worksheet once to count rows, determine its positional column
   count, and infer a Parquet type for each column.
3. Reads each worksheet again and writes its values in batches of up to 10,000
   rows. This avoids holding the entire workbook's data in memory.
4. Writes a manifest containing the original worksheet names, output filenames,
   row counts, column positions, and inferred types.

Parquet columns are named by position rather than by guessing which row contains
headers. A source worksheet row such as:

| A | B | C |
|---|---|---|
| Date | Event | Note |
| 2026-01-01 | Example | Holiday |

is stored conceptually as:

| source_row | column_1 | column_2 | column_3 |
|---:|---|---|---|
| 1 | Date | Event | Note |
| 2 | 2026-01-01 | Example | Holiday |

Blank cells become nulls. Type inference promotes mixed integer and decimal
values to floating point; other incompatible mixtures are represented as text.
The exact output types for each worksheet are recorded in the manifest.

## Output layout

```text
<workbook-name>_parquet/
    sheet_001.parquet
    sheet_002.parquet
    ...
    manifest.json
```

The numbered filenames are stable within a particular workbook order. Use the
manifest, rather than assuming a filename means a particular subject, to find
the original worksheet name. Keep the Parquet files **and** the manifest
together when copying or publishing the export.

## How to handle sheets that refer to one another

Each worksheet stays in its own Parquet file, and the manifest preserves the
mapping from the file to the original sheet. The exporter does not merge sheets
or infer join keys. If two sheets are related in Excel, use the relevant
exported columns to join them in your analysis, and document the key columns and
their meanings in your data workflow.

Parquet contains values and a tabular schema; it is not an Excel workbook.
Formulas, cross-sheet formula references, named ranges, formatting, macros, and
recalculation behavior are not preserved. Formula cells are exported only as
the values exposed by `pyxlsb`; those values are not live formulas.

## Storing the original in GitLab

For preserving the original workbook and its Excel relationships, use Git LFS
instead of replacing the workbook with Parquet. This project has an LFS
attribute for `.xlsb` files in [`.gitattributes`](./.gitattributes). Git LFS
must be enabled for the GitLab project, and its storage and transfer quotas
still apply.

The current [`.gitignore`](./.gitignore) ignores the supplied workbook by name.
If you intend to add that original file to Git LFS, force-add it:

```powershell
git lfs install
git add .gitattributes
git add -f "DRUK VIDESHI PANCHANGAMU JSRK LATEST 03.04.2026.xlsb"
```

Commit and push as usual after checking that the GitLab project supports LFS.
If you only need the extracted data in Git, add the Parquet output folder and
manifest instead; exports can still be large, so consider LFS for those files
too.

## Troubleshooting

If Python reports `PermissionError` while opening the workbook, check that the
file is accessible to the Windows account running Python, close applications
that may be locking it, and try a readable local copy. The converter needs read
access to the source file and write access to the output directory.
