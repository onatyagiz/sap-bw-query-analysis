"""
Compare two hierarchical SAP BW query Excel exports.

Hierarchy is parsed using the Nodekey / Parent-Nodekey columns.
Each node is placed in its section (top-2-level ancestor context).
Comparison is by existence only — order is irrelevant.
Identity = (section, definition_type, technical_value_or_description).
"""

import openpyxl
from pathlib import Path

# Unnamed structural containers — skipped as identity nodes but their
# children are still collected under the same section.
UNNAMED_CONTAINERS = {
    'Filter', 'Characteristic Restrictions', 'Default Values',
    'Query Variables', 'Rows/Columns', 'Rows', 'Columns',
    'Free Characteristics', 'Structure',
}


# ── Parsing ────────────────────────────────────────────────────────────────

def has_nodekey_columns(ws):
    """True for exports with Nodekey / Parent / Definition in columns S–U."""
    if ws.max_column < 21:
        return False
    return any(ws.cell(row=r, column=19).value == 'Nodekey' for r in range(1, 4))


def parse_indented(ws):
    """Build the node dict from an export without Nodekey columns.

    Newer exports drop columns S–U and move the technical value to column P.
    The tree is implied by indentation instead: a row's definition is the
    label in its leftmost non-empty column A–N, and its parent is the nearest
    earlier row labelled further to the left. Must run on the raw export —
    a filled-down sheet has lost that indentation.
    """
    nodes = {}
    stack = []   # (depth, nodekey) of the current ancestor chain

    for row_idx, row in enumerate(ws.iter_rows(min_row=4, values_only=True), 4):
        depth = defn = None
        for col, v in enumerate(row[:14]):
            if v is not None and str(v).strip():
                depth, defn = col, str(v).strip()
                break
        if defn is None:
            continue

        while stack and stack[-1][0] >= depth:
            stack.pop()

        desc  = row[14] if len(row) > 14 else None   # Column O – Description
        value = row[15] if len(row) > 15 else None   # Column P – Technical value
        nk = str(row_idx)

        nodes[nk] = {
            'parent':      stack[-1][1] if stack else '',
            'definition':  defn,
            'description': str(desc).strip()  if desc  is not None else '',
            'value':       str(value).strip() if value is not None else '',
            'children':    [],
        }
        stack.append((depth, nk))

    for nk, node in nodes.items():
        p = node['parent']
        if p and p in nodes:
            nodes[p]['children'].append(nk)

    return nodes


def parse_excel(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    if not has_nodekey_columns(ws):
        return parse_indented(ws)
    nodes = {}

    for row in ws.iter_rows(values_only=True):
        nodekey = row[18]   # Column S
        parent  = row[19]   # Column T
        defn    = row[20]   # Column U – Definition/type
        desc    = row[14]   # Column O – Description
        value   = row[16]   # Column Q – Technical value

        if nodekey is None:
            continue

        nk   = str(nodekey).strip()
        par  = str(parent).strip()  if parent is not None else ''
        defn = str(defn).strip()    if defn   is not None else ''
        desc = str(desc).strip()    if desc   is not None else ''
        val  = str(value).strip()   if value  is not None else ''

        if not nk or not defn:
            continue

        nodes[nk] = {
            'parent':     par,
            'definition': defn,
            'description': desc,
            'value':      val,
            'children':   [],
        }

    for nk, node in nodes.items():
        p = node['parent']
        if p and p in nodes:
            nodes[p]['children'].append(nk)

    return nodes


def find_root(nodes):
    for nk, node in nodes.items():
        if not node['parent'] or node['parent'] not in nodes:
            return nk
    return None


# ── Section annotation ─────────────────────────────────────────────────────

def annotate_sections(nodes, root_key):
    """Add a 'section' string to every node.

    Section = "<depth-1 definition>" or "<depth-1> > <depth-2>",
    measured from the Query root. The root itself gets section=''.
    """
    def walk(nk, d1, d2):
        node = nodes[nk]
        defn = node['definition']

        if d1 is None:
            # This is the root (Query)
            node['section'] = ''
            for child in node['children']:
                walk(child, defn, None)
        elif d2 is None:
            # Depth-1 child of root
            node['section'] = d1
            for child in node['children']:
                walk(child, d1, defn)
        else:
            # Depth ≥ 2: section stays at d1 > d2
            node['section'] = f"{d1} > {d2}"
            for child in node['children']:
                walk(child, d1, d2)

    walk(root_key, None, None)


# ── Flat element collection ────────────────────────────────────────────────

def collect_elements(nodes, root_key):
    """Return dict: (section, definition, identity_value) -> node.

    Unnamed structural containers contribute no identity entry of their own,
    but their children are collected under the same section.
    A named Selection (has description) IS included as an identity.
    """
    annotate_sections(nodes, root_key)
    result = {}

    def is_unnamed_container(node):
        return (node['definition'] in UNNAMED_CONTAINERS
                and not node['value'] and not node['description'])

    def walk(nk):
        if nk == root_key:
            for child in nodes[root_key]['children']:
                walk(child)
            return

        node = nodes[nk]

        if not is_unnamed_container(node):
            section = node['section']
            defn    = node['definition']
            val     = node['value']
            desc    = node['description']
            identity_val = val if val else desc
            key = (section, defn, identity_val)
            if key not in result:
                result[key] = node

        for child in node['children']:
            walk(child)

    walk(root_key)
    return result


# ── Text report ────────────────────────────────────────────────────────────

def write_text_report(stream, file1, file2, set1, set2):
    import sys
    old = sys.stdout
    sys.stdout = stream

    ids1      = set(set1.keys())
    ids2      = set(set2.keys())
    only_a    = sorted(ids1 - ids2)
    only_b    = sorted(ids2 - ids1)
    common    = ids1 & ids2

    print("=" * 80)
    print("QUERY COMPARISON REPORT")
    print("=" * 80)
    print(f"  File A : {Path(file1).name}")
    print(f"  File B : {Path(file2).name}")
    print()
    print(f"  Elements in A  : {len(ids1)}")
    print(f"  Elements in B  : {len(ids2)}")
    print(f"  Common         : {len(common)}")
    print(f"  Only in A      : {len(only_a)}")
    print(f"  Only in B      : {len(only_b)}")
    print()

    def print_group(items, node_map, title):
        if not items:
            return
        print("─" * 80)
        print(f"{title}  ({len(items)} items)")
        print("─" * 80)
        # Group by section, then by definition
        by_section = {}
        for key in items:
            sec = key[0] or '(Query)'
            by_section.setdefault(sec, []).append(key)
        for sec in sorted(by_section):
            print(f"\n  ▸ {sec}")
            by_defn = {}
            for key in by_section[sec]:
                by_defn.setdefault(key[1], []).append(key)
            for defn in sorted(by_defn):
                print(f"    [ {defn} ]")
                for key in sorted(by_defn[defn]):
                    n = node_map[key]
                    v = n['value']
                    d = n['description']
                    if v and d and v != d:
                        print(f"      {v}  —  {d}")
                    elif v:
                        print(f"      {v}")
                    else:
                        print(f"      {d}")
        print()

    print_group(only_a, set1, "ELEMENTS ONLY IN FILE A")
    print_group(only_b, set2, "ELEMENTS ONLY IN FILE B")

    if not only_a and not only_b:
        print("No differences found.")

    print("=" * 80)
    sys.stdout = old


# ── Excel report ───────────────────────────────────────────────────────────

def write_excel_report(file1, file2, set1, set2, out_path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    ids1   = set(set1.keys())
    ids2   = set(set2.keys())
    all_ids = sorted(ids1 | ids2)   # sorted by (section, defn, value)

    wb = Workbook()
    ws = wb.active
    ws.title = "Comparison"

    # Colours
    C_HEADER   = "1F4E79"
    C_SECTION  = "D6E4F0"   # blue-grey  – section group header
    C_SUBSEC   = "EBF3FB"   # lighter    – sub-section row
    C_ONLY_A   = "FCE4D6"   # orange     – only in A
    C_ONLY_B   = "E2EFDA"   # green      – only in B
    C_COMMON   = "FFFFFF"   # white      – in both

    thin   = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    name_a = Path(file1).stem
    name_b = Path(file2).stem

    # Header
    headers = ["Section", "Element Type", "Technical Value", "Description",
               f"In A\n({name_a[:24]})", f"In B\n({name_b[:24]})", "Status"]
    ws.append(headers)
    for col in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col)
        c.font      = Font(bold=True, color="FFFFFF", size=11)
        c.fill      = PatternFill("solid", fgColor=C_HEADER)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = border
    ws.row_dimensions[1].height = 36

    row_num    = 2
    prev_sec   = None

    for key in all_ids:
        section, defn, ident_val = key
        in_a = key in ids1
        in_b = key in ids2

        node   = set1.get(key) or set2.get(key)
        val    = node['value']
        desc   = node['description']
        status = "Common" if (in_a and in_b) else ("Only in A" if in_a else "Only in B")
        fill_c = C_COMMON if (in_a and in_b) else (C_ONLY_A if in_a else C_ONLY_B)

        # Section divider row when section changes
        if section != prev_sec:
            sec_label = section or "(Query)"
            ws.append([sec_label, "", "", "", "", "", ""])
            for col in range(1, 8):
                c = ws.cell(row=row_num, column=col)
                c.fill      = PatternFill("solid", fgColor=C_SECTION)
                c.font      = Font(bold=True, size=10)
                c.alignment = Alignment(horizontal="left", vertical="center")
                c.border    = border
            # Merge section label across all columns for readability
            ws.merge_cells(start_row=row_num, start_column=1,
                           end_row=row_num,   end_column=7)
            ws.row_dimensions[row_num].height = 18
            row_num  += 1
            prev_sec  = section

        ws.append(["", defn, val, desc,
                   "✓" if in_a else "✗",
                   "✓" if in_b else "✗",
                   status])

        fill = PatternFill("solid", fgColor=fill_c)
        for col in range(1, 8):
            c = ws.cell(row=row_num, column=col)
            c.fill      = fill
            c.border    = border
            c.alignment = Alignment(
                horizontal="center" if col in (5, 6, 7) else "left",
                vertical="center"
            )
        row_num += 1

    # Column widths
    for i, w in enumerate([28, 22, 32, 44, 9, 9, 12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:G{row_num - 1}"

    # Summary sheet
    ws2 = wb.create_sheet("Summary")
    ws2.column_dimensions["A"].width = 36
    ws2.column_dimensions["B"].width = 16

    rows = [
        ("File A", Path(file1).name),
        ("File B", Path(file2).name),
        ("", ""),
        ("Elements in A",   len(ids1)),
        ("Elements in B",   len(ids2)),
        ("Common",          len(ids1 & ids2)),
        ("Only in A",       len(ids1 - ids2)),
        ("Only in B",       len(ids2 - ids1)),
        ("", ""),
    ]
    # Per-section breakdown
    sections = sorted({k[0] for k in ids1 | ids2})
    for sec in sections:
        sec_a = {k for k in ids1 if k[0] == sec}
        sec_b = {k for k in ids2 if k[0] == sec}
        sec_label = sec or "(Query)"
        rows.append((f"  {sec_label} — only in A", len(sec_a - ids2)))
        rows.append((f"  {sec_label} — only in B", len(ids2 & {k for k in ids2 if k[0] == sec} - ids1)))

    for r, (label, val) in enumerate(rows, 1):
        ws2.cell(row=r, column=1, value=label).font = Font(
            bold=bool(label and not label.startswith("  ")))
        ws2.cell(row=r, column=2, value=val)

    wb.save(out_path)


# ── Entry point ────────────────────────────────────────────────────────────

def compare(file1, file2):
    nodes1 = parse_excel(file1)
    nodes2 = parse_excel(file2)

    root1  = find_root(nodes1)
    root2  = find_root(nodes2)

    if root1 is None or root2 is None:
        print("ERROR: Could not find root node.")
        return

    set1 = collect_elements(nodes1, root1)
    set2 = collect_elements(nodes2, root2)

    import sys
    write_text_report(sys.stdout, file1, file2, set1, set2)

    txt_path  = Path(file1).parent / "comparison_report.txt"
    xlsx_path = Path(file1).parent / "comparison_report.xlsx"

    with open(txt_path, "w", encoding="utf-8") as f:
        write_text_report(f, file1, file2, set1, set2)

    write_excel_report(file1, file2, set1, set2, xlsx_path)

    print(f"\nText report : {txt_path}")
    print(f"Excel report: {xlsx_path}")


if __name__ == "__main__":
    base  = Path(__file__).parent.parent / "Query Comparison"
    files = sorted(base.glob("G_*.xlsx"))
    if len(files) != 2:
        print(f"Expected exactly 2 .xlsx files, found: {files}")
    else:
        compare(str(files[0]), str(files[1]))
