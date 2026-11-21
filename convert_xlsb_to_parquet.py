#!/usr/bin/env python3
"""Export an XLSB workbook as one positional Parquet table per worksheet."""

from __future__ import annotations

import argparse
import json
import math
import numbers
import os
from datetime import date, datetime, time
from pathlib import Path
from typing import Any, Iterator

import pyarrow as pa
import pyarrow.parquet as pq
from pyxlsb import open_workbook


BATCH_SIZE = 10_000


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Export every XLSB worksheet to its own Parquet file. "
            "Cells are retained as positional columns (column_1, column_2, ...)."
        )
    )
    parser.add_argument("workbook", type=Path, help="Path to the .xlsb workbook")
    parser.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        help="Output directory (default: <workbook name>_parquet beside the workbook)",
    )
    return parser.parse_args()


def cell_kind(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, datetime):
        return "timestamp"
    if isinstance(value, date):
        return "timestamp"
    if isinstance(value, numbers.Integral):
        return "int"
    if isinstance(value, numbers.Real):
        return "float"
    if isinstance(value, (str, bytes)):
        return "string"
    return "string"


def promote_kind(current: str | None, incoming: str | None) -> str | None:
    if incoming is None:
        return current
    if current is None or current == incoming:
        return incoming
    if {current, incoming} <= {"int", "float"}:
        return "float"
    return "string"


def arrow_type(kind: str | None) -> pa.DataType:
    return {
        "bool": pa.bool_(),
        "int": pa.int64(),
        "float": pa.float64(),
        "timestamp": pa.timestamp("us"),
        "string": pa.string(),
        None: pa.string(),
    }[kind]


def rows_for_sheet(workbook: Any, sheet_name: str) -> Iterator[tuple[int, dict[int, Any]]]:
    with workbook.get_sheet(sheet_name) as sheet:
        fallback_row_number = 0
        for row in sheet.rows(sparse=False):
            fallback_row_number += 1
            values: dict[int, Any] = {}
            source_row: int | None = None
            for cell in row:
                if cell is None:
                    continue
                column_number = int(cell.c)
                values[column_number] = cell.v
                if source_row is None:
                    source_row = int(cell.r) + 1
            yield source_row or fallback_row_number, values


def infer_sheet(
    workbook: Any, sheet_name: str
) -> tuple[int, int, list[str | None]]:
    row_count = 0
    column_count = 0
    kinds: list[str | None] = []

    for _, values in rows_for_sheet(workbook, sheet_name):
        row_count += 1
        if values:
            column_count = max(column_count, max(values) + 1)
            if len(kinds) < column_count:
                kinds.extend([None] * (column_count - len(kinds)))
            for column_number, value in values.items():
                kinds[column_number] = promote_kind(
                    kinds[column_number], cell_kind(value)
                )

    return row_count, column_count, kinds


def to_arrow_value(value: Any, kind: str | None) -> Any:
    if value is None:
        return None
    if kind == "string":
        if isinstance(value, bytes):
            return value.decode("utf-8", errors="replace")
        return value.isoformat() if isinstance(value, (date, time)) else str(value)
    if kind == "float":
        result = float(value)
        return result if math.isfinite(result) else None
    if kind == "timestamp" and isinstance(value, date) and not isinstance(value, datetime):
        return datetime.combine(value, time.min)
    return value


def write_sheet(
    workbook: Any,
    sheet_name: str,
    output_path: Path,
    row_count: int,
    column_count: int,
    kinds: list[str | None],
) -> None:
    fields = [pa.field("source_row", pa.int64())]
    fields.extend(
        pa.field(f"column_{index + 1}", arrow_type(kinds[index]))
        for index in range(column_count)
    )
    schema = pa.schema(fields)
    temporary_path = output_path.with_name(output_path.name + ".tmp")

    try:
        with pq.ParquetWriter(temporary_path, schema) as writer:
            batch: list[dict[str, Any]] = []
            written_rows = 0
            for source_row, values in rows_for_sheet(workbook, sheet_name):
                record: dict[str, Any] = {"source_row": source_row}
                for column_number in range(column_count):
                    record[f"column_{column_number + 1}"] = to_arrow_value(
                        values.get(column_number), kinds[column_number]
                    )
                batch.append(record)
                written_rows += 1
                if len(batch) >= BATCH_SIZE:
                    writer.write_table(pa.Table.from_pylist(batch, schema=schema))
                    batch.clear()
            if batch:
                writer.write_table(pa.Table.from_pylist(batch, schema=schema))
        if written_rows != row_count:
            raise RuntimeError(
                f"Worksheet {sheet_name!r} changed while exporting "
                f"({row_count} rows scanned, {written_rows} rows written)."
            )
        os.replace(temporary_path, output_path)
    except Exception:
        if temporary_path.exists():
            temporary_path.unlink()
        raise


def main() -> None:
    args = parse_args()
    workbook_path = args.workbook.expanduser().resolve()
    if not workbook_path.is_file():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")
    if workbook_path.suffix.lower() != ".xlsb":
        raise ValueError(f"Expected an .xlsb workbook: {workbook_path}")

    output_dir = args.output_dir
    if output_dir is None:
        output_dir = workbook_path.with_name(workbook_path.stem + "_parquet")
    output_dir = output_dir.expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    manifest: dict[str, Any] = {
        "format_version": 1,
        "source_workbook": workbook_path.name,
        "representation": (
            "One Parquet file per worksheet; all populated cell values are "
            "stored in positional columns, with source_row preserving row positions."
        ),
        "preserves_formula_definitions_or_recalculation": False,
        "notes": [
            "Formula cells are exported as the cached values exposed by pyxlsb.",
            "Sheet formulas, cross-sheet links, named ranges, formatting, and macros "
            "are not preserved in Parquet.",
            "Use the original XLSB when Excel behavior or formula dependencies are needed.",
        ],
        "sheets": [],
    }

    with open_workbook(str(workbook_path)) as workbook:
        for index, sheet_name in enumerate(workbook.sheets, start=1):
            row_count, column_count, kinds = infer_sheet(workbook, sheet_name)
            parquet_name = f"sheet_{index:03d}.parquet"
            parquet_path = output_dir / parquet_name
            write_sheet(
                workbook,
                sheet_name,
                parquet_path,
                row_count,
                column_count,
                kinds,
            )
            manifest["sheets"].append(
                {
                    "order": index,
                    "name": sheet_name,
                    "file": parquet_name,
                    "rows": row_count,
                    "columns": [
                        {
                            "position": column_index + 1,
                            "name": f"column_{column_index + 1}",
                            "inferred_type": kinds[column_index] or "null",
                        }
                        for column_index in range(column_count)
                    ],
                }
            )
            print(
                f"Exported {sheet_name!r}: {row_count} rows, "
                f"{column_count} columns -> {parquet_name}"
            )

    manifest_path = output_dir / "manifest.json"
    temporary_manifest = output_dir / "manifest.json.tmp"
    temporary_manifest.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    os.replace(temporary_manifest, manifest_path)
    print(f"Manifest written to: {manifest_path}")


if __name__ == "__main__":
    main()
