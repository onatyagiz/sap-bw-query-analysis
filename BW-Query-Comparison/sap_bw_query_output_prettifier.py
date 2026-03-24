"""
Creates a replica of a hierarchical SAP BW query Excel export.

The original files use a sparse column layout where a value in column C (for
example "Filter") implies that all following rows belong to that section until
a new value appears in column C or a higher-level column.

This script fills those values downward so every data row shows its full
hierarchical context. All other columns (Description, Value, Nodekey, etc.)
are preserved as-is.
"""

import openpyxl
from pathlib import Path


# Columns A–N (1-based in openpyxl) hold the hierarchical object labels.
HIER_COLS = list(range(1, 15))   # 1 = A, 2 = B, … 14 = N

# Rows 1–3 are header / title rows — skip fill-down for these.
DATA_START_ROW = 4


def fill_down_hierarchy(ws):
    """Fill hierarchical columns A–N downward in-place.

    Rule: when a new value appears in column X, all columns to the right
    (X+1 … N) are reset so they no longer carry a stale ancestor label.
    """
    current = {}   # col (1-based) -> current fill value

    for row_idx in range(DATA_START_ROW, ws.max_row + 1):

        # Find the leftmost hierarchical column that has a new value this row.
        new_col = None
        for col in HIER_COLS:
            v = ws.cell(row=row_idx, column=col).value
            if v is not None and str(v).strip():
                new_col = col
                break

        if new_col is not None:
            # A new ancestor appeared: reset all deeper columns.
            for col in HIER_COLS:
                if col > new_col:
                    current[col] = None
            current[new_col] = ws.cell(row=row_idx, column=new_col).value

        # Write the running values into every hierarchical cell on this row.
        for col in HIER_COLS:
            v = current.get(col)
            if v is not None:
                ws.cell(row=row_idx, column=col).value = v


def create_replica(input_path, output_path):
    print(f"Reading  : {input_path}")
    wb = openpyxl.load_workbook(input_path)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        fill_down_hierarchy(ws)
        print(f"  Sheet '{sheet_name}': filled {ws.max_row} rows × "
              f"{len(HIER_COLS)} hierarchical columns")

    wb.save(output_path)
    print(f"Saved    : {output_path}\n")


if __name__ == "__main__":
    base = Path(__file__).parent.parent / "Query Comparison"
    for src in sorted(base.glob("G_*.xlsx")):
        dst = src.parent / f"{src.stem}_replica.xlsx"
        create_replica(str(src), str(dst))
