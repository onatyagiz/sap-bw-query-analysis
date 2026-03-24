# SAP BW Query Comparison

A set of Python scripts for analysing and comparing two SAP BW query definition Excel exports.

## Folder Structure

```
Excel-Analysis/
├── BW-Query-Comparison/          ← scripts (this folder)
│   ├── sap_bw_master_analysis.py
│   ├── sap_bw_query_output_prettifier.py
│   ├── sap_bw_compare_queries.py
│   └── sap_bw_analyze_key_figures.py
└── Query Comparison/             ← input & output files
    ├── G_<QUERY_A>.xlsx
    ├── G_<QUERY_B>.xlsx
    └── bw_query_analysis_report.xlsx   ← combined output
```

## Requirements

```bash
pip install openpyxl
```

Python 3.8 or later.

## Usage

### Master Script (recommended)

Runs all analyses and produces a single combined Excel report.

```bash
cd Excel-Analysis
python3 BW-Query-Comparison/sap_bw_master_analysis.py
```

You will be prompted for the two `.xlsx` file paths. The output file `bw_query_analysis_report.xlsx` is saved in the same folder as the inputs.

### Individual Scripts

Each script can also be run standalone. By default they look for two `G_*.xlsx` files in the `Query Comparison/` folder next to the `BW-Query-Comparison/` folder.

| Script | Purpose | Outputs |
|--------|---------|---------|
| `sap_bw_query_output_prettifier.py` | Fills sparse hierarchy columns A–N downward so every row shows its full ancestor context | `*_replica.xlsx` per input file |
| `sap_bw_compare_queries.py` | Element-by-element comparison of two queries | `comparison_report.txt`, `comparison_report.xlsx` |
| `sap_bw_analyze_key_figures.py` | Analysis of Restricted KFs, Calculated KFs, and Formulas in the Rows/Columns section | `kf_analysis_report.txt`, `kf_analysis_report.xlsx` |

```bash
cd Excel-Analysis
python3 BW-Query-Comparison/sap_bw_query_output_prettifier.py
python3 BW-Query-Comparison/sap_bw_compare_queries.py
python3 BW-Query-Comparison/sap_bw_analyze_key_figures.py
```

## Combined Report Sheets

The master script produces a single workbook with eight sheets:

| Sheet | Description |
|-------|-------------|
| **Query A (Filled)** | Source data for File A with hierarchy columns filled down |
| **Query B (Filled)** | Source data for File B with hierarchy columns filled down |
| **Comparison** | All query elements colour-coded: orange = only in A, green = only in B, white = common |
| **Comparison Summary** | Element counts broken down by section |
| **Restricted KFs** | Restricted Key Figures compared by structural content (Key Figure + Characteristic restrictions) |
| **Calculated KFs** | Calculated Key Figures matched by description, formula text compared side-by-side |
| **Formulas** | Standalone Row_XXXXX formulas matched by description, formula text compared |
| **KF Summary** | Counts for RKFs, CKFs, and Formulas |

## How the Analysis Works

### Hierarchy Fill (`sap_bw_query_output_prettifier`)

SAP BW exports use a sparse column layout: a value in column C (e.g. `Filter`) applies to all following rows until a new value appears in column C or a shallower column. The prettifier fills these values downward so every row has its full hierarchical context visible.

### Element Comparison (`sap_bw_compare_queries`)

Parses the Nodekey / Parent Nodekey tree (columns S, T, U) to reconstruct the query hierarchy. Each element is identified by `(section, definition type, technical value)`. Comparison is order-independent — only existence matters. Unnamed structural containers (Filter, Rows/Columns, etc.) are treated as transparent groupings.

### Key Figure Analysis (`sap_bw_analyze_key_figures`)

Focuses on the **Rows/Columns** section of the query and extracts three object types:

- **Restricted Key Figures (RKF)** — identified by Key Figure + Characteristic restrictions (ranges/variables). Two RKFs with different technical names but identical restrictions are flagged as structurally equivalent.
- **Calculated Key Figures (CKF)** — matched across files by fuzzy description name; formula expressions compared side-by-side.
- **Standalone Formulas (Row_XXXXX)** — matched by fuzzy description name; formula text compared side-by-side.

## Input File Format

Standard SAP BW query definition Excel export (`G_*.xlsx`). The scripts expect:

| Column | Content |
|--------|---------|
| A – N  | Hierarchical object labels (sparse, filled by prettifier) |
| O (15) | Description |
| Q (17) | Technical value / formula expression |
| S (19) | Nodekey (unique node ID) |
| T (20) | Parent Nodekey |
| U (21) | Definition type (Query, Filter, Restricted KF, Calculated KF, Formula, …) |

Rows 1–3 are treated as headers; data starts at row 4.
