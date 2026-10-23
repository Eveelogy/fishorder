import json
import os
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

EXCEL_FILE = "Book 2.xlsx"
BACKUP_FILE = "Book 2 (Original).xlsx"
HTML_FILE = "Fish_Order_Catalog.html"


def col_letter(n):
    """1-based column index to Excel column letter (1 -> A, 27 -> AA)."""
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


def read_workbook_data(xlsx_path):
    ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(xlsx_path, "r") as z:
        ss = []
        if "xl/sharedStrings.xml" in z.namelist():
            ss_xml = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in ss_xml.findall("main:si", ns):
                ss.append("".join(t.text or "" for t in si.findall(".//main:t", ns)))

        wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
        rels_xml = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rels = {r.attrib["Id"]: r.attrib["Target"] for r in rels_xml}

        sheets = []
        for sheet in wb_xml.findall(".//main:sheet", ns):
            name = sheet.attrib["name"]
            rid = sheet.attrib[
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            ]
            target = "xl/" + rels[rid]
            ws_xml = ET.fromstring(z.read(target))
            rows = []
            for row in ws_xml.findall(".//main:row", ns):
                vals = {}
                for c in row.findall("main:c", ns):
                    col = re.sub(r"\d+", "", c.attrib.get("r", ""))
                    v_el = c.find("main:v", ns)
                    v = v_el.text if v_el is not None else None
                    if c.attrib.get("t") == "s" and v is not None:
                        v = ss[int(v)]
                    vals[col] = (v or "").strip()
                rows.append(vals)
            sheets.append((name, rows))
    return sheets


def build_excel_xml(sheets_spec):
    """
    sheets_spec: list of dicts:
      {
        'name': str,
        'headers': [str, ...],
        'rows': [[val, ...], ...],
        'col_widths': [float, ...],
      }
    Returns a dict of archive_path -> bytes for a complete, styled OpenXML .xlsx file
    with frozen header rows and AutoFilter enabled on every sheet.
    """
    files = {}

    # 1. [Content_Types].xml
    overrides = [
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
    ]
    for i in range(1, len(sheets_spec) + 1):
        overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )

    files["[Content_Types].xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        + "".join(overrides)
        + "</Types>"
    ).encode("utf-8")

    # 2. _rels/.rels
    files["_rels/.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    ).encode("utf-8")

    # 3. xl/workbook.xml and xl/_rels/workbook.xml.rels
    sheet_els = []
    rel_els = []
    defined_names = []
    for i, sp in enumerate(sheets_spec, 1):
        safe_name = escape(sp["name"], {'"': "&quot;"})
        sheet_els.append(f'<sheet name="{safe_name}" sheetId="{i}" r:id="rId{i}"/>')
        rel_els.append(
            f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
        )
        last_col = col_letter(len(sp["headers"]))
        last_row = len(sp["rows"]) + 1
        quoted_sheet = "'" + sp["name"].replace("'", "''") + "'"
        defined_names.append(
            f'<definedName name="_xlnm._FilterDatabase" localSheetId="{i-1}" hidden="1">{escape(quoted_sheet)}!$A$1:${last_col}${last_row}</definedName>'
        )

    rel_els.append(
        f'<Relationship Id="rId{len(sheets_spec)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    )
    rel_els.append(
        f'<Relationship Id="rId{len(sheets_spec)+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
    )

    files["xl/workbook.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="14000" activeTab="0"/></bookViews>'
        "<sheets>" + "".join(sheet_els) + "</sheets>"
        "<definedNames>" + "".join(defined_names) + "</definedNames>"
        "</workbook>"
    ).encode("utf-8")

    files["xl/_rels/workbook.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(rel_els)
        + "</Relationships>"
    ).encode("utf-8")

    # 4. xl/styles.xml
    # Style 0: default
    # Style 1: Header (bold white text on dark navy #0F2C59, centered vertically, thin border)
    # Style 2: Normal data cell (left-aligned, thin light border)
    # Style 3: Centered data cell (for Assortment Code, Vendor Count, Yes flags, SKU/Vendor#)
    # Style 4: Highlighted 'Yes' vendor badge cell (soft green fill #DCFCE7, dark green bold text)
    # Style 5: Unlisted / Warning cell (soft amber fill #FEF3C7, dark amber text)
    files["xl/styles.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="5">'
        '<font><sz val="10"/><color rgb="FF1E293B"/><name val="Segoe UI"/><family val="2"/></font>'
        '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/><family val="2"/></font>'
        '<font><b/><sz val="10"/><color rgb="FF166534"/><name val="Segoe UI"/><family val="2"/></font>'
        '<font><sz val="10"/><color rgb="FF92400E"/><name val="Segoe UI"/><family val="2"/></font>'
        '<font><b/><sz val="10"/><color rgb="FF0F172A"/><name val="Segoe UI"/><family val="2"/></font>'
        "</fonts>"
        '<fills count="5">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF0F2C59"/><bgColor indexed="64"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFDCFCE7"/><bgColor indexed="64"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFEF3C7"/><bgColor indexed="64"/></patternFill></fill>'
        "</fills>"
        '<borders count="2">'
        "<border><left/><right/><top/><bottom/><diagonal/></border>"
        '<border>'
        '<left style="thin"><color rgb="FFE2E8F0"/></left>'
        '<right style="thin"><color rgb="FFE2E8F0"/></right>'
        '<top style="thin"><color rgb="FFE2E8F0"/></top>'
        '<bottom style="thin"><color rgb="FFE2E8F0"/></bottom>'
        "<diagonal/>"
        "</border>"
        "</borders>"
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="6">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>'
        '<xf numFmtId="0" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
        '<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
        '<xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
        "</cellXfs>"
        "</styleSheet>"
    ).encode("utf-8")

    # 5. Build worksheets and sharedStrings
    string_map = {}
    shared_strings = []

    def get_ss_idx(s):
        if s not in string_map:
            string_map[s] = len(shared_strings)
            shared_strings.append(s)
        return string_map[s]

    for i, sp in enumerate(sheets_spec, 1):
        headers = sp["headers"]
        rows = sp["rows"]
        col_widths = sp["col_widths"]
        center_cols = sp.get("center_cols", set())
        vendor_flag_cols = sp.get("vendor_flag_cols", set())
        num_cols = len(headers)
        last_col = col_letter(num_cols)
        total_rows = len(rows) + 1

        cols_xml = ["<cols>"]
        for c_idx, w in enumerate(col_widths, 1):
            cols_xml.append(
                f'<col min="{c_idx}" max="{c_idx}" width="{w}" customWidth="1"/>'
            )
        cols_xml.append("</cols>")

        rows_xml = []
        # Header row (row 1)
        h_cells = []
        for c_idx, h_text in enumerate(headers, 1):
            ref = f"{col_letter(c_idx)}1"
            ss_i = get_ss_idx(str(h_text))
            h_cells.append(f'<c r="{ref}" s="1" t="s"><v>{ss_i}</v></c>')
        rows_xml.append(
            f'<row r="1" spans="1:{num_cols}" ht="26" customHeight="1">'
            + "".join(h_cells)
            + "</row>"
        )

        # Data rows
        for r_idx, row_vals in enumerate(rows, 2):
            r_cells = []
            for c_idx, val in enumerate(row_vals, 1):
                ref = f"{col_letter(c_idx)}{r_idx}"
                val_str = "" if val is None else str(val)
                # Determine style
                style_id = 2
                if c_idx in vendor_flag_cols and val_str == "Yes":
                    style_id = 4
                elif val_str == "Unlisted":
                    style_id = 5
                elif c_idx in center_cols:
                    style_id = 3

                if isinstance(val, int):
                    r_cells.append(f'<c r="{ref}" s="{style_id}"><v>{val}</v></c>')
                elif val_str == "":
                    r_cells.append(f'<c r="{ref}" s="{style_id}"/>')
                else:
                    ss_i = get_ss_idx(val_str)
                    r_cells.append(
                        f'<c r="{ref}" s="{style_id}" t="s"><v>{ss_i}</v></c>'
                    )
            rows_xml.append(
                f'<row r="{r_idx}" spans="1:{num_cols}" ht="20" customHeight="1">'
                + "".join(r_cells)
                + "</row>"
            )

        ws_content = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<dimension ref="A1:{last_col}{total_rows}"/>'
            "<sheetViews>"
            '<sheetView workbookViewId="0">'
            '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
            '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
            "</sheetView>"
            "</sheetViews>"
            '<sheetFormatPr defaultRowHeight="20"/>'
            + "".join(cols_xml)
            + "<sheetData>"
            + "".join(rows_xml)
            + "</sheetData>"
            + f'<autoFilter ref="A1:{last_col}{total_rows}"/>'
            + "</worksheet>"
        )
        files[f"xl/worksheets/sheet{i}.xml"] = ws_content.encode("utf-8")

    # 6. xl/sharedStrings.xml
    ss_items = [
        f"<si><t>{escape(s)}</t></si>" for s in shared_strings
    ]
    files["xl/sharedStrings.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(shared_strings)}" uniqueCount="{len(shared_strings)}">'
        + "".join(ss_items)
        + "</sst>"
    ).encode("utf-8")

    return files


def classify_water_type(cat_num, desc=""):
    c = (cat_num or "").strip()
    if c.startswith("85"):
        return "Saltwater"
    if c.startswith("84"):
        return "Freshwater"
    if c.startswith("78") or c.startswith("87"):
        return "Plants / Terrarium"
    # Fallback for unlisted Order sheet items (all 124 Order items are freshwater fish)
    return "Freshwater"


def main():
    if not os.path.exists(BACKUP_FILE):
        shutil.copy2(EXCEL_FILE, BACKUP_FILE)
        print(f"Backed up original file to {BACKUP_FILE}")

    raw_sheets = read_workbook_data(BACKUP_FILE)

    vendor_sheets = []
    vendor_num_to_name = {}
    vendor_num_to_sheet = {}
    sku_master = {}

    for name, rows in raw_sheets:
        if name == "Order":
            order_raw_rows = rows
            continue
        vendor_sheets.append((name, rows))
        if " " in name:
            v_name, v_num = name.rsplit(" ", 1)
            vendor_num_to_name[v_num.strip()] = v_name.strip()
            vendor_num_to_sheet[v_num.strip()] = name.strip()

        for r in rows:
            sku = (r.get("A") or "").strip()
            if not sku or not sku.isdigit():
                continue
            desc = (r.get("B") or "").strip()
            cat_num = (r.get("C") or "").strip()
            cat_desc = (r.get("D") or "").strip()
            assort = (r.get("E") or "").strip()

            if sku not in sku_master:
                sku_master[sku] = {
                    "sku": sku,
                    "desc": desc,
                    "cat_num": cat_num,
                    "cat_desc": cat_desc,
                    "water_type": classify_water_type(cat_num, desc),
                    "assort": assort,
                    "vendors": [],
                }
            if name not in sku_master[sku]["vendors"]:
                sku_master[sku]["vendors"].append(name)

    vendor_sheet_names = [name for name, _ in vendor_sheets]

    master_items = sorted(
        sku_master.values(),
        key=lambda x: (x["assort"] or "Z", x["desc"], int(x["sku"]) if x["sku"].isdigit() else 0),
    )

    # Build Cleaned Order rows (all 124 valid fish rows)
    cleaned_order = []
    for r in order_raw_rows:
        sku = (r.get("A") or "").strip()
        if not sku or not sku.isdigit():
            continue
        desc = (r.get("B") or "").strip()
        v_num = (r.get("F") or "").strip()
        v_name = vendor_num_to_name.get(v_num, "Unlisted")
        v_sheet = vendor_num_to_sheet.get(v_num, "Unlisted")
        master_entry = sku_master.get(sku)
        assort = master_entry["assort"] if master_entry else "Unlisted"
        cat_desc = master_entry["cat_desc"] if master_entry else "Unlisted"
        water_type = (
            master_entry["water_type"]
            if master_entry
            else classify_water_type("", desc)
        )
        all_vendors = (
            ", ".join(master_entry["vendors"]) if master_entry else "Unlisted"
        )
        cleaned_order.append(
            {
                "sku": sku,
                "desc": desc,
                "v_num": v_num,
                "v_name": v_name,
                "v_sheet": v_sheet,
                "water_type": water_type,
                "assort": assort,
                "cat_desc": cat_desc,
                "all_vendors": all_vendors,
            }
        )

    print(f"Total Master SKUs: {len(master_items)}")
    print(f"Total Cleaned Order rows: {len(cleaned_order)}")

    # Prepare Excel sheets specification
    # Sheet 1: Order (Cleaned Order sheet)
    order_headers = [
        "SKU",
        "Description",
        "Vendor #",
        "Vendor Name",
        "Water Type",
        "Assortment Code",
        "Category Description",
        "All Available Vendors",
    ]
    order_excel_rows = [
        [
            item["sku"],
            item["desc"],
            item["v_num"],
            item["v_name"],
            item["water_type"],
            item["assort"],
            item["cat_desc"],
            item["all_vendors"],
        ]
        for item in cleaned_order
    ]

    # Sheet 2: Master SKU List
    master_headers = [
        "SKU #",
        "Description",
        "Water Type",
        "Category #",
        "Category Description",
        "Assortment Code",
        "Vendor Count",
        "Available Vendors",
    ] + vendor_sheet_names

    master_excel_rows = []
    for item in master_items:
        v_set = set(item["vendors"])
        row = [
            item["sku"],
            item["desc"],
            item["water_type"],
            item["cat_num"],
            item["cat_desc"],
            item["assort"],
            len(item["vendors"]),
            ", ".join(item["vendors"]),
        ] + ["Yes" if vn in v_set else "" for vn in vendor_sheet_names]
        master_excel_rows.append(row)

    sheets_spec = [
        {
            "name": "Order",
            "headers": order_headers,
            "rows": order_excel_rows,
            "col_widths": [14, 34, 14, 18, 18, 18, 28, 42],
            "center_cols": {1, 3, 5, 6},
            "vendor_flag_cols": set(),
        },
        {
            "name": "Master SKU List",
            "headers": master_headers,
            "rows": master_excel_rows,
            "col_widths": [14, 34, 18, 14, 28, 18, 15, 42]
            + [18] * len(vendor_sheet_names),
            "center_cols": {1, 3, 4, 6, 7}
            | set(range(9, 9 + len(vendor_sheet_names))),
            "vendor_flag_cols": set(range(9, 9 + len(vendor_sheet_names))),
        },
    ]

    # Also include the 13 vendor sheets (deduplicated & with clean headers + AutoFilter)
    vendor_headers = [
        "SKU #",
        "Description",
        "Category #",
        "Category Description",
        "Assortment Code",
    ]
    for v_name, v_rows in vendor_sheets:
        seen = set()
        dedup_rows = []
        for r in v_rows:
            sku = (r.get("A") or "").strip()
            if not sku or not sku.isdigit() or sku in seen:
                continue
            seen.add(sku)
            dedup_rows.append(
                [
                    sku,
                    (r.get("B") or "").strip(),
                    (r.get("C") or "").strip(),
                    (r.get("D") or "").strip(),
                    (r.get("E") or "").strip(),
                ]
            )
        sheets_spec.append(
            {
                "name": v_name,
                "headers": vendor_headers,
                "rows": dedup_rows,
                "col_widths": [14, 34, 14, 28, 18],
                "center_cols": {1, 3, 5},
                "vendor_flag_cols": set(),
            }
        )

    xlsx_files = build_excel_xml(sheets_spec)
    with zipfile.ZipFile(EXCEL_FILE, "w", zipfile.ZIP_DEFLATED) as zf:
        for path, data in xlsx_files.items():
            zf.writestr(path, data)
    print(f"Successfully updated {EXCEL_FILE}")

    # Build standalone interactive HTML file
    build_html_viewer(master_items, cleaned_order, vendor_sheet_names)
    print(f"Successfully created {HTML_FILE}")


def build_html_viewer(master_items, cleaned_order, vendor_sheet_names):
    assort_codes = sorted({x["assort"] for x in master_items if x["assort"]})
    data_json = json.dumps(
        {
            "master": master_items,
            "order": cleaned_order,
            "vendors": vendor_sheet_names,
            "assortCodes": assort_codes,
        }
    )

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Fish Order &amp; Master SKU Catalog</title>
<style>
  :root {{
    --bg: #f8fafc;
    --card: #ffffff;
    --navy: #0f2c59;
    --navy-light: #1e3a8a;
    --border: #e2e8f0;
    --text: #0f172a;
    --muted: #64748b;
    --accent: #0284c7;
    --green-bg: #dcfce7;
    --green-text: #166534;
    --amber-bg: #fef3c7;
    --amber-text: #92400e;
  }}
  * {{ box-sizing: border-box; }}
  body {{
    margin: 0;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    background: var(--bg);
    color: var(--text);
  }}
  header {{
    background: linear-gradient(135deg, #0f2c59 0%, #1e3a8a 100%);
    color: #fff;
    padding: 20px 28px;
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    justify-content: space-between;
    gap: 16px;
    box-shadow: 0 2px 8px rgba(15, 44, 89, 0.15);
  }}
  header h1 {{
    margin: 0;
    font-size: 1.35rem;
    font-weight: 700;
    letter-spacing: -0.01em;
  }}
  header p {{
    margin: 4px 0 0;
    font-size: 0.875rem;
    color: #cbd5e1;
  }}
  .tabs {{
    display: flex;
    gap: 8px;
    background: rgba(255, 255, 255, 0.12);
    padding: 5px;
    border-radius: 10px;
  }}
  .tab-btn {{
    border: none;
    background: transparent;
    color: #e2e8f0;
    padding: 9px 18px;
    border-radius: 7px;
    font-size: 0.9rem;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.15s;
  }}
  .tab-btn.active {{
    background: #fff;
    color: var(--navy);
    box-shadow: 0 1px 4px rgba(0, 0, 0, 0.12);
  }}
  .container {{
    max-width: 1600px;
    margin: 0 auto;
    padding: 20px 28px;
  }}
  .controls {{
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 16px;
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 14px;
    align-items: end;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
  }}
  .control-group {{
    display: flex;
    flex-direction: column;
    gap: 6px;
  }}
  .control-group label {{
    font-size: 0.75rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.04em;
    color: var(--muted);
  }}
  input[type="search"], select {{
    width: 100%;
    padding: 9px 12px;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    font-size: 0.9rem;
    color: var(--text);
    background: #fff;
  }}
  input[type="search"]:focus, select:focus {{
    outline: 2px solid var(--accent);
    border-color: transparent;
  }}
  .btn-row {{
    display: flex;
    gap: 8px;
    align-items: center;
  }}
  .btn {{
    padding: 9px 14px;
    border-radius: 8px;
    border: 1px solid #cbd5e1;
    background: #fff;
    color: var(--text);
    font-size: 0.85rem;
    font-weight: 600;
    cursor: pointer;
    white-space: nowrap;
  }}
  .btn:hover {{ background: #f1f5f9; }}
  .btn-primary {{
    background: var(--navy);
    color: #fff;
    border-color: var(--navy);
  }}
  .btn-primary:hover {{ background: var(--navy-light); }}
  .stats-bar {{
    display: flex;
    flex-wrap: wrap;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 12px;
    font-size: 0.875rem;
    color: var(--muted);
    gap: 10px;
  }}
  .table-wrap {{
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 12px;
    overflow: auto;
    max-height: calc(100vh - 270px);
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
  }}
  table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 0.875rem;
  }}
  thead th {{
    position: sticky;
    top: 0;
    z-index: 2;
    background: var(--navy);
    color: #fff;
    text-align: left;
    padding: 11px 14px;
    font-weight: 600;
    white-space: nowrap;
    cursor: pointer;
    user-select: none;
    border-bottom: 2px solid #0a1f40;
  }}
  thead th:hover {{
    background: var(--navy-light);
  }}
  thead th .sort-ind {{
    margin-left: 6px;
    font-size: 0.75rem;
    opacity: 0.8;
  }}
  tbody tr {{
    border-bottom: 1px solid var(--border);
  }}
  tbody tr:hover {{
    background: #f1f5f9;
  }}
  td {{
    padding: 9px 14px;
    vertical-align: middle;
  }}
  .mono {{
    font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
    font-weight: 600;
  }}
  .badge {{
    display: inline-block;
    padding: 2px 8px;
    border-radius: 999px;
    font-size: 0.78rem;
    font-weight: 700;
  }}
  .badge-assort {{
    background: #e0f2fe;
    color: #0369a1;
  }}
  .badge-vendor {{
    background: #f1f5f9;
    color: #334155;
    border: 1px solid #cbd5e1;
    margin: 2px 4px 2px 0;
    cursor: pointer;
  }}
  .badge-vendor:hover {{
    background: #e2e8f0;
  }}
  .badge-yes {{
    background: var(--green-bg);
    color: var(--green-text);
  }}
  .badge-unlisted {{
    background: var(--amber-bg);
    color: var(--amber-text);
  }}
  .assort-select {{
    appearance: none;
    -webkit-appearance: none;
    padding: 3px 22px 3px 10px;
    border-radius: 999px;
    font-size: 0.78rem;
    font-weight: 700;
    border: 1px solid transparent;
    cursor: pointer;
    text-align: center;
    background-image: url("data:image/svg+xml;charset=UTF-8,%3csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2.5' stroke-linecap='round' stroke-linejoin='round'%3e%3cpolyline points='6 9 12 15 18 9'%3e%3c/polyline%3e%3c/svg%3e");
    background-repeat: no-repeat;
    background-position: right 6px center;
    background-size: 11px;
    width: auto;
    min-width: 68px;
  }}
  .assort-select.is-listed {{
    background-color: #e0f2fe;
    color: #0369a1;
    border-color: #bae6fd;
  }}
  .assort-select.is-unlisted {{
    background-color: var(--amber-bg);
    color: var(--amber-text);
    border-color: #fde68a;
  }}
  .assort-select.water-fresh {{
    background-color: #dcfce7;
    color: #166534;
    border-color: #bbf7d0;
  }}
  .assort-select.water-salt {{
    background-color: #e0e7ff;
    color: #1e40af;
    border-color: #c7d2fe;
  }}
  .assort-select.water-plants {{
    background-color: #fef9c3;
    color: #854d0e;
    border-color: #fde047;
  }}
  .assort-select.is-edited {{
    box-shadow: 0 0 0 2px #0284c7;
  }}
  .editable-cell {{
    cursor: pointer;
    border-radius: 6px;
    padding: 4px 6px;
    margin: -4px -6px;
    transition: background 0.12s;
    display: inline-block;
    min-width: 60px;
  }}
  .editable-cell:hover {{
    background: #e2e8f0;
  }}
  .editable-cell.is-edited {{
    background: #eff6ff;
    border-bottom: 1px dashed #0284c7;
  }}
  .edit-banner {{
    display: inline-flex;
    align-items: center;
    gap: 8px;
    font-size: 0.82rem;
    color: #0369a1;
    background: #e0f2fe;
    padding: 5px 12px;
    border-radius: 8px;
    font-weight: 600;
  }}
  .assort-pills {{
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
    align-items: center;
  }}
  .assort-pill {{
    display: inline-flex;
    align-items: center;
    gap: 5px;
    padding: 6px 11px;
    border-radius: 8px;
    border: 1px solid #cbd5e1;
    background: #fff;
    color: var(--text);
    font-size: 0.82rem;
    font-weight: 600;
    cursor: pointer;
    user-select: none;
    transition: all 0.15s;
  }}
  .assort-pill:hover {{
    border-color: var(--accent);
    background: #f0f9ff;
  }}
  .assort-pill.active {{
    background: var(--navy);
    color: #fff;
    border-color: var(--navy);
  }}
  .assort-pill .pill-count {{
    font-size: 0.72rem;
    opacity: 0.75;
    font-weight: 500;
  }}
  .center {{ text-align: center; }}
</style>
</head>
<body>
<header>
  <div>
    <h1>Fish Order &amp; Vendor Master SKU Catalog</h1>
    <p>Searchable, sortable &amp; editable across all 13 vendor sheets and the cleaned Order sheet</p>
  </div>
  <div class="tabs">
    <button class="tab-btn active" id="tabMaster" onclick="switchTab('master')">Master SKU List (1,307 SKUs)</button>
    <button class="tab-btn" id="tabOrder" onclick="switchTab('order')">Cleaned Order Sheet (124 Items)</button>
  </div>
</header>

<div class="container">
  <div class="controls">
    <div class="control-group" style="grid-column: span 2;">
      <label for="searchInput">Search SKU, Description, Category, or Vendor</label>
      <input type="search" id="searchInput" placeholder="Type SKU #, fish name (e.g. Clown, Koi, Betta), category, or vendor..." oninput="render()">
    </div>
    <div class="control-group">
      <label for="vendorFilter">Filter by Vendor</label>
      <select id="vendorFilter" onchange="onVendorFilterChange()">
        <option value="">All Vendors</option>
      </select>
    </div>
    <div class="control-group">
      <label for="sortBySelect">Quick Sort</label>
      <select id="sortBySelect" onchange="handleQuickSort()">
        <option value="assort:asc">Assortment Code (A → Z)</option>
        <option value="water_type:asc">Water Type (Fresh → Salt)</option>
        <option value="vendors:asc">Vendor Name (A → Z)</option>
        <option value="vendor_count:desc">Most Vendors First</option>
        <option value="desc:asc">Description (A → Z)</option>
        <option value="sku:asc">SKU # (Low → High)</option>
        <option value="cat_desc:asc">Category (A → Z)</option>
      </select>
    </div>
    <div class="control-group" style="grid-column: span 2;">
      <label>Water Type Filter (Click Freshwater, Saltwater, or Plants / Terrarium)</label>
      <div class="assort-pills" id="waterTypePills"></div>
    </div>
    <div class="control-group" style="grid-column: span 2;">
      <label>Display Assortment Codes (Click one or multiple, e.g. G or A + B)</label>
      <div class="assort-pills" id="assortPills"></div>
    </div>
    <div class="btn-row" style="grid-column: 1 / -1; justify-content: space-between; flex-wrap: wrap; border-top: 1px solid var(--border); padding-top: 10px; margin-top: 2px;">
      <div style="display:flex; gap:8px; align-items:center; flex-wrap:wrap;">
        <span style="font-size:0.82rem; color:var(--muted);">✏️ <strong>Inline Editing:</strong> Use the <strong>Water Type</strong> or <strong>Assortment Code</strong> dropdown in any row, or click any <strong>Description</strong>, <strong>Vendor Name</strong>, or <strong>Category</strong> cell to edit it. Edits auto-save in your browser.</span>
        <span id="editStatusBadge" class="edit-banner" style="display:none;"></span>
      </div>
      <div style="display:flex; gap:8px; align-items:center; flex-wrap:wrap;">
        <button class="btn" id="toggleMatrixBtn" onclick="toggleVendorColumns()">Show Individual Vendor Columns</button>
        <button class="btn" onclick="resetFilters()">Reset Filters</button>
        <button class="btn" id="clearEditsBtn" style="display:none; color:#b91c1c; border-color:#fecaca;" onclick="clearAllEdits()">Revert All Edits</button>
        <button class="btn" onclick="saveStandaloneHTML()" title="Download an updated HTML file with your edits permanently baked in">Save Updated HTML</button>
        <button class="btn btn-primary" onclick="exportCSV()">Export CSV</button>
      </div>
    </div>
  </div>

  <div class="stats-bar">
    <div id="rowCountText">Showing 0 rows</div>
    <div>Tip: Click any column header to sort, or click a vendor badge in the table to filter.</div>
  </div>

  <div class="table-wrap">
    <table>
      <thead id="tableHead"></thead>
      <tbody id="tableBody"></tbody>
    </table>
  </div>
</div>

<script id="catalogDataScript">
const INITIAL_DATA = {data_json};
</script>
<script>
const STORAGE_KEY = 'fish_catalog_custom_edits_v1';
const WATER_TYPES = ['Freshwater', 'Saltwater', 'Plants / Terrarium'];
let DATA = JSON.parse(JSON.stringify(INITIAL_DATA));
let customEdits = {{ skuEdits: {{}}, vendorNameByNum: {{}} }};

let currentTab = 'master';
let sortField = 'assort';
let sortDir = 1;
let showMatrixCols = false;
let selectedAssorts = new Set();
let selectedWaterTypes = new Set();

function loadSavedEdits() {{
  try {{
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {{
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === 'object') {{
        customEdits.skuEdits = parsed.skuEdits || {{}};
        customEdits.vendorNameByNum = parsed.vendorNameByNum || {{}};
      }}
    }}
  }} catch (e) {{
    console.warn('Could not read localStorage edits:', e);
  }}
  applyEditsToData();
}}

function saveEditsToStorage() {{
  try {{
    localStorage.setItem(STORAGE_KEY, JSON.stringify(customEdits));
  }} catch (e) {{
    console.warn('Could not save edits to localStorage:', e);
  }}
  updateEditStatusUI();
}}

function applyEditsToData() {{
  DATA = JSON.parse(JSON.stringify(INITIAL_DATA));

  // Apply SKU-level edits (assort, water_type, desc, cat_desc) to both master and order lists
  DATA.master.forEach(r => {{
    const ed = customEdits.skuEdits[r.sku];
    if (ed) {{
      if (ed.assort !== undefined) r.assort = ed.assort;
      if (ed.water_type !== undefined) r.water_type = ed.water_type;
      if (ed.desc !== undefined) r.desc = ed.desc;
      if (ed.cat_desc !== undefined) r.cat_desc = ed.cat_desc;
      r._edited = true;
    }}
  }});

  DATA.order.forEach(r => {{
    const ed = customEdits.skuEdits[r.sku];
    if (ed) {{
      if (ed.assort !== undefined) r.assort = ed.assort;
      if (ed.water_type !== undefined) r.water_type = ed.water_type;
      if (ed.desc !== undefined) r.desc = ed.desc;
      if (ed.cat_desc !== undefined) r.cat_desc = ed.cat_desc;
      r._edited = true;
    }}
    if (customEdits.vendorNameByNum[r.v_num] !== undefined) {{
      r.v_name = customEdits.vendorNameByNum[r.v_num];
      r._vEdited = true;
    }}
  }});

  // Rebuild assortCodes list if any new assortment code was added
  const codeSet = new Set(INITIAL_DATA.assortCodes);
  DATA.master.forEach(r => {{ if (r.assort && r.assort !== 'Unlisted') codeSet.add(r.assort); }});
  DATA.order.forEach(r => {{ if (r.assort && r.assort !== 'Unlisted') codeSet.add(r.assort); }});
  DATA.assortCodes = [...codeSet].sort();

  updateEditStatusUI();
}}

function updateEditStatusUI() {{
  const skuCount = Object.keys(customEdits.skuEdits).length;
  const vCount = Object.keys(customEdits.vendorNameByNum).length;
  const total = skuCount + vCount;
  const badge = document.getElementById('editStatusBadge');
  const clearBtn = document.getElementById('clearEditsBtn');
  if (total > 0) {{
    badge.style.display = 'inline-flex';
    badge.textContent = `${{total}} custom edit${{total === 1 ? '' : 's'}} saved`;
    clearBtn.style.display = 'inline-block';
  }} else {{
    badge.style.display = 'none';
    clearBtn.style.display = 'none';
  }}
}}

function updateSkuAssort(sku, value) {{
  if (value === '__custom__') {{
    const custom = prompt('Enter Assortment Code (e.g. A, B, C, E, G, S):', '');
    if (!custom || !custom.trim()) {{
      render();
      return;
    }}
    value = custom.trim().toUpperCase();
  }}
  if (!customEdits.skuEdits[sku]) customEdits.skuEdits[sku] = {{}};
  customEdits.skuEdits[sku].assort = value;
  saveEditsToStorage();
  applyEditsToData();
  renderWaterTypePills();
  renderAssortPills();
  render();
}}

function updateSkuWaterType(sku, value) {{
  if (!customEdits.skuEdits[sku]) customEdits.skuEdits[sku] = {{}};
  customEdits.skuEdits[sku].water_type = value;
  saveEditsToStorage();
  applyEditsToData();
  renderWaterTypePills();
  renderAssortPills();
  render();
}}

function editSkuField(sku, field, label, currentVal) {{
  const next = prompt(`Edit ${{label}} for SKU ${{sku}}:`, currentVal);
  if (next === null) return;
  const trimmed = next.trim() || 'Unlisted';
  if (!customEdits.skuEdits[sku]) customEdits.skuEdits[sku] = {{}};
  customEdits.skuEdits[sku][field] = trimmed;
  saveEditsToStorage();
  applyEditsToData();
  render();
}}

function editVendorName(vNum, currentName) {{
  const next = prompt(`Enter Vendor Name for Vendor # ${{vNum}} (updates all rows with Vendor # ${{vNum}}):`, currentName === 'Unlisted' ? '' : currentName);
  if (next === null) return;
  const trimmed = next.trim() || 'Unlisted';
  customEdits.vendorNameByNum[vNum] = trimmed;
  saveEditsToStorage();
  applyEditsToData();
  initFilters();
  render();
}}

function clearAllEdits() {{
  if (!confirm('Revert all custom edits back to the original spreadsheet values?')) return;
  customEdits = {{ skuEdits: {{}}, vendorNameByNum: {{}} }};
  localStorage.removeItem(STORAGE_KEY);
  applyEditsToData();
  initFilters();
  render();
}}

function saveStandaloneHTML() {{
  // Bake current DATA into INITIAL_DATA so the downloaded HTML file permanently contains all edits
  const bakedData = JSON.parse(JSON.stringify(DATA));
  bakedData.master.forEach(r => {{ delete r._edited; }});
  bakedData.order.forEach(r => {{ delete r._edited; delete r._vEdited; }});
  const docClone = document.documentElement.outerHTML.replace(
    /<script id="catalogDataScript">[\\s\\S]*?<\\/script>/,
    `<script id="catalogDataScript">\\nconst INITIAL_DATA = ${{JSON.stringify(bakedData)}};\\n<\\/script>`
  );
  const blob = new Blob(['<!DOCTYPE html>\\n' + docClone], {{ type: 'text/html;charset=utf-8;' }});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'Fish_Order_Catalog.html';
  a.click();
  URL.revokeObjectURL(url);
}}

function renderWaterTypePills() {{
  const container = document.getElementById('waterTypePills');
  const vf = document.getElementById('vendorFilter').value;
  let sourceRows = currentTab === 'master' ? [...DATA.master] : [...DATA.order];
  if (vf) {{
    sourceRows = currentTab === 'master'
      ? sourceRows.filter(r => r.vendors.includes(vf))
      : sourceRows.filter(r => (r.v_name + ' (' + r.v_num + ')') === vf);
  }}

  const counts = {{}};
  sourceRows.forEach(r => {{
    counts[r.water_type] = (counts[r.water_type] || 0) + 1;
  }});

  const allActive = selectedWaterTypes.size === 0;
  let html = `<button type="button" class="assort-pill ${{allActive ? 'active' : ''}}" onclick="clearWaterTypeSelection()">All <span class="pill-count">(${{sourceRows.length}})</span></button>`;

  WATER_TYPES.forEach(wt => {{
    const isActive = selectedWaterTypes.has(wt);
    const cnt = counts[wt] || 0;
    html += `<button type="button" class="assort-pill ${{isActive ? 'active' : ''}}" onclick="toggleWaterType('${{wt}}')">${{wt}} <span class="pill-count">(${{cnt}})</span></button>`;
  }});

  container.innerHTML = html;
}}

function toggleWaterType(wt) {{
  if (selectedWaterTypes.has(wt)) {{
    selectedWaterTypes.delete(wt);
  }} else {{
    selectedWaterTypes.add(wt);
  }}
  renderWaterTypePills();
  renderAssortPills();
  render();
}}

function clearWaterTypeSelection() {{
  selectedWaterTypes.clear();
  renderWaterTypePills();
  renderAssortPills();
  render();
}}

function renderAssortPills() {{
  const container = document.getElementById('assortPills');
  const vf = document.getElementById('vendorFilter').value;
  let sourceRows = currentTab === 'master' ? [...DATA.master] : [...DATA.order];
  if (vf) {{
    sourceRows = currentTab === 'master'
      ? sourceRows.filter(r => r.vendors.includes(vf))
      : sourceRows.filter(r => (r.v_name + ' (' + r.v_num + ')') === vf);
  }}
  if (selectedWaterTypes.size > 0) {{
    sourceRows = sourceRows.filter(r => selectedWaterTypes.has(r.water_type));
  }}

  const counts = {{}};
  sourceRows.forEach(r => {{
    counts[r.assort] = (counts[r.assort] || 0) + 1;
  }});

  const codes = [...DATA.assortCodes];
  if (currentTab === 'order' && !codes.includes('Unlisted')) {{
    codes.push('Unlisted');
  }}

  const allActive = selectedAssorts.size === 0;
  let html = `<button type="button" class="assort-pill ${{allActive ? 'active' : ''}}" onclick="clearAssortSelection()">All <span class="pill-count">(${{sourceRows.length}})</span></button>`;

  codes.forEach(c => {{
    const isActive = selectedAssorts.has(c);
    const cnt = counts[c] || 0;
    html += `<button type="button" class="assort-pill ${{isActive ? 'active' : ''}}" onclick="toggleAssortCode('${{c}}')">${{c}} <span class="pill-count">(${{cnt}})</span></button>`;
  }});

  container.innerHTML = html;
}}

function toggleAssortCode(code) {{
  if (selectedAssorts.has(code)) {{
    selectedAssorts.delete(code);
  }} else {{
    selectedAssorts.add(code);
  }}
  renderAssortPills();
  render();
}}

function clearAssortSelection() {{
  selectedAssorts.clear();
  renderAssortPills();
  render();
}}

function onVendorFilterChange() {{
  renderWaterTypePills();
  renderAssortPills();
  render();
}}

function initFilters() {{
  const vf = document.getElementById('vendorFilter');
  const prevVendor = vf.value;
  vf.innerHTML = '<option value="">All Vendors</option>';
  if (currentTab === 'master') {{
    DATA.vendors.forEach(v => {{
      const opt = document.createElement('option');
      opt.value = v;
      opt.textContent = v;
      vf.appendChild(opt);
    }});
  }} else {{
    const orderVendors = [...new Set(DATA.order.map(r => r.v_name + ' (' + r.v_num + ')'))];
    orderVendors.forEach(v => {{
      const opt = document.createElement('option');
      opt.value = v;
      opt.textContent = v;
      vf.appendChild(opt);
    }});
  }}
  if ([...vf.options].some(o => o.value === prevVendor)) vf.value = prevVendor;

  if (currentTab === 'master' && selectedAssorts.has('Unlisted')) {{
    selectedAssorts.delete('Unlisted');
  }}
  renderWaterTypePills();
  renderAssortPills();
}}

function switchTab(tab) {{
  currentTab = tab;
  document.getElementById('tabMaster').classList.toggle('active', tab === 'master');
  document.getElementById('tabOrder').classList.toggle('active', tab === 'order');
  document.getElementById('toggleMatrixBtn').style.display = tab === 'master' ? 'inline-block' : 'none';
  initFilters();
  render();
}}

function toggleVendorColumns() {{
  showMatrixCols = !showMatrixCols;
  document.getElementById('toggleMatrixBtn').textContent = showMatrixCols
    ? 'Hide Individual Vendor Columns'
    : 'Show Individual Vendor Columns';
  render();
}}

function handleQuickSort() {{
  const val = document.getElementById('sortBySelect').value;
  const [field, dir] = val.split(':');
  sortField = field;
  sortDir = dir === 'asc' ? 1 : -1;
  render();
}}

function setSort(field) {{
  if (sortField === field) {{
    sortDir = -sortDir;
  }} else {{
    sortField = field;
    sortDir = 1;
  }}
  render();
}}

function filterByVendorBadge(v) {{
  const vf = document.getElementById('vendorFilter');
  vf.value = v;
  onVendorFilterChange();
}}

function resetFilters() {{
  document.getElementById('searchInput').value = '';
  document.getElementById('vendorFilter').value = '';
  selectedWaterTypes.clear();
  selectedAssorts.clear();
  renderWaterTypePills();
  renderAssortPills();
  render();
}}

function getFilteredData() {{
  const q = document.getElementById('searchInput').value.trim().toLowerCase();
  const vf = document.getElementById('vendorFilter').value;

  let rows = currentTab === 'master' ? [...DATA.master] : [...DATA.order];

  if (q) {{
    rows = rows.filter(r => {{
      if (currentTab === 'master') {{
        return r.sku.toLowerCase().includes(q) ||
               r.desc.toLowerCase().includes(q) ||
               r.cat_num.toLowerCase().includes(q) ||
               r.cat_desc.toLowerCase().includes(q) ||
               r.water_type.toLowerCase().includes(q) ||
               r.vendors.some(v => v.toLowerCase().includes(q));
      }} else {{
        return r.sku.toLowerCase().includes(q) ||
               r.desc.toLowerCase().includes(q) ||
               r.v_num.toLowerCase().includes(q) ||
               r.v_name.toLowerCase().includes(q) ||
               r.water_type.toLowerCase().includes(q) ||
               r.cat_desc.toLowerCase().includes(q) ||
               r.all_vendors.toLowerCase().includes(q);
      }}
    }});
  }}

  if (vf) {{
    if (currentTab === 'master') {{
      rows = rows.filter(r => r.vendors.includes(vf));
    }} else {{
      rows = rows.filter(r => (r.v_name + ' (' + r.v_num + ')') === vf);
    }}
  }}

  if (selectedWaterTypes.size > 0) {{
    rows = rows.filter(r => selectedWaterTypes.has(r.water_type));
  }}

  if (selectedAssorts.size > 0) {{
    rows = rows.filter(r => selectedAssorts.has(r.assort));
  }}

  rows.sort((a, b) => {{
    let va, vb;
    if (sortField === 'sku' || sortField === 'cat_num' || sortField === 'v_num') {{
      va = parseInt(a[sortField] || '0', 10);
      vb = parseInt(b[sortField] || '0', 10);
    }} else if (sortField === 'vendor_count') {{
      va = a.vendors ? a.vendors.length : 0;
      vb = b.vendors ? b.vendors.length : 0;
    }} else if (sortField === 'vendors') {{
      va = currentTab === 'master' ? a.vendors.join(', ') : a.v_name;
      vb = currentTab === 'master' ? b.vendors.join(', ') : b.v_name;
    }} else if (sortField.startsWith('vcol:')) {{
      const vName = sortField.slice(5);
      va = a.vendors.includes(vName) ? 0 : 1;
      vb = b.vendors.includes(vName) ? 0 : 1;
    }} else {{
      va = (a[sortField] || '').toString().toLowerCase();
      vb = (b[sortField] || '').toString().toLowerCase();
    }}
    if (va < vb) return -1 * sortDir;
    if (va > vb) return 1 * sortDir;
    return 0;
  }});

  return rows;
}}

function escapeAttr(str) {{
  return String(str ?? '').replace(/&/g, '&amp;').replace(/'/g, '&#39;').replace(/"/g, '&quot;');
}}

function buildAssortSelectHTML(sku, currentAssort, isEdited) {{
  const codes = [...new Set([...DATA.assortCodes, 'Unlisted'])];
  if (currentAssort && !codes.includes(currentAssort)) codes.unshift(currentAssort);
  const cls = (currentAssort === 'Unlisted' ? 'is-unlisted' : 'is-listed') + (isEdited ? ' is-edited' : '');
  const opts = codes.map(c => `<option value="${{escapeAttr(c)}}" ${{c === currentAssort ? 'selected' : ''}}>${{c}}</option>`).join('') +
               `<option value="__custom__">+ Custom Code...</option>`;
  return `<select class="assort-select ${{cls}}" onchange="updateSkuAssort('${{sku}}', this.value)" title="Click to change Assortment Code for SKU ${{sku}}">${{opts}}</select>`;
}}

function buildWaterTypeSelectHTML(sku, currentWt, isEdited) {{
  let wtClass = 'water-fresh';
  if (currentWt === 'Saltwater') wtClass = 'water-salt';
  else if (currentWt === 'Plants / Terrarium') wtClass = 'water-plants';
  const cls = wtClass + (isEdited ? ' is-edited' : '');
  const opts = WATER_TYPES.map(wt => `<option value="${{escapeAttr(wt)}}" ${{wt === currentWt ? 'selected' : ''}}>${{wt}}</option>`).join('');
  return `<select class="assort-select ${{cls}}" onchange="updateSkuWaterType('${{sku}}', this.value)" title="Click to change Water Type for SKU ${{sku}}">${{opts}}</select>`;
}}

function render() {{
  const rows = getFilteredData();
  const thead = document.getElementById('tableHead');
  const tbody = document.getElementById('tableBody');

  const arrow = (f) => sortField === f ? `<span class="sort-ind">${{sortDir === 1 ? '▲' : '▼'}}</span>` : '';

  if (currentTab === 'master') {{
    let cols = [
      {{ key: 'sku', label: 'SKU #' }},
      {{ key: 'desc', label: 'Description' }},
      {{ key: 'water_type', label: 'Water Type' }},
      {{ key: 'cat_num', label: 'Category #' }},
      {{ key: 'cat_desc', label: 'Category Description' }},
      {{ key: 'assort', label: 'Assortment Code' }},
      {{ key: 'vendor_count', label: 'Vendor Count' }},
      {{ key: 'vendors', label: 'Available Vendors' }}
    ];
    if (showMatrixCols) {{
      DATA.vendors.forEach(v => cols.push({{ key: 'vcol:' + v, label: v }}));
    }}
    thead.innerHTML = '<tr>' + cols.map(c => `<th onclick="setSort('${{c.key}}')">${{c.label}}${{arrow(c.key)}}</th>`).join('') + '</tr>';

    tbody.innerHTML = rows.map(r => {{
      const ed = customEdits.skuEdits[r.sku] || {{}};
      const badges = r.vendors.map(v => `<span class="badge badge-vendor" onclick="filterByVendorBadge('${{v}}')" title="Click to filter by ${{v}}">${{v}}</span>`).join('');
      let matrixCells = '';
      if (showMatrixCols) {{
        matrixCells = DATA.vendors.map(v => `<td class="center">${{r.vendors.includes(v) ? '<span class="badge badge-yes">Yes</span>' : '-'}}</td>`).join('');
      }}
      return `<tr>
        <td class="mono">${{r.sku}}</td>
        <td><span class="editable-cell ${{ed.desc !== undefined ? 'is-edited' : ''}}" onclick="editSkuField('${{r.sku}}', 'desc', 'Description', '${{escapeAttr(r.desc)}}')" title="Click to edit Description"><strong>${{r.desc}}</strong></span></td>
        <td class="center">${{buildWaterTypeSelectHTML(r.sku, r.water_type, ed.water_type !== undefined)}}</td>
        <td class="mono center">${{r.cat_num}}</td>
        <td><span class="editable-cell ${{ed.cat_desc !== undefined ? 'is-edited' : ''}}" onclick="editSkuField('${{r.sku}}', 'cat_desc', 'Category Description', '${{escapeAttr(r.cat_desc)}}')" title="Click to edit Category Description">${{r.cat_desc}}</span></td>
        <td class="center">${{buildAssortSelectHTML(r.sku, r.assort, ed.assort !== undefined)}}</td>
        <td class="center mono">${{r.vendors.length}}</td>
        <td>${{badges}}</td>
        ${{matrixCells}}
      </tr>`;
    }}).join('');
  }} else {{
    const cols = [
      {{ key: 'sku', label: 'SKU' }},
      {{ key: 'desc', label: 'Description' }},
      {{ key: 'v_num', label: 'Vendor #' }},
      {{ key: 'v_name', label: 'Vendor Name' }},
      {{ key: 'water_type', label: 'Water Type' }},
      {{ key: 'assort', label: 'Assortment Code' }},
      {{ key: 'cat_desc', label: 'Category Description' }},
      {{ key: 'all_vendors', label: 'All Available Vendors (Master List)' }}
    ];
    thead.innerHTML = '<tr>' + cols.map(c => `<th onclick="setSort('${{c.key}}')">${{c.label}}${{arrow(c.key)}}</th>`).join('') + '</tr>';

    tbody.innerHTML = rows.map(r => {{
      const ed = customEdits.skuEdits[r.sku] || {{}};
      const vEdited = customEdits.vendorNameByNum[r.v_num] !== undefined;
      const vBadge = r.v_name === 'Unlisted'
        ? '<span class="badge badge-unlisted">Unlisted</span>'
        : `<strong>${{r.v_name}}</strong>`;
      const catDisplay = r.cat_desc === 'Unlisted'
        ? '<span class="badge badge-unlisted">Unlisted</span>'
        : r.cat_desc;
      return `<tr>
        <td class="mono">${{r.sku}}</td>
        <td><span class="editable-cell ${{ed.desc !== undefined ? 'is-edited' : ''}}" onclick="editSkuField('${{r.sku}}', 'desc', 'Description', '${{escapeAttr(r.desc)}}')" title="Click to edit Description"><strong>${{r.desc}}</strong></span></td>
        <td class="mono center">${{r.v_num}}</td>
        <td><span class="editable-cell ${{vEdited ? 'is-edited' : ''}}" onclick="editVendorName('${{r.v_num}}', '${{escapeAttr(r.v_name)}}')" title="Click to set Vendor Name for Vendor # ${{r.v_num}}">${{vBadge}}</span></td>
        <td class="center">${{buildWaterTypeSelectHTML(r.sku, r.water_type, ed.water_type !== undefined)}}</td>
        <td class="center">${{buildAssortSelectHTML(r.sku, r.assort, ed.assort !== undefined)}}</td>
        <td><span class="editable-cell ${{ed.cat_desc !== undefined ? 'is-edited' : ''}}" onclick="editSkuField('${{r.sku}}', 'cat_desc', 'Category Description', '${{escapeAttr(r.cat_desc)}}')" title="Click to edit Category Description">${{catDisplay}}</span></td>
        <td>${{r.all_vendors === 'Unlisted' ? '<span class="badge badge-unlisted">Unlisted</span>' : r.all_vendors}}</td>
      </tr>`;
    }}).join('');
  }}

  const filterNotes = [];
  if (selectedWaterTypes.size > 0) filterNotes.push(`Water: ${{[...selectedWaterTypes].join(', ')}}`);
  if (selectedAssorts.size > 0) filterNotes.push(`Assortment: ${{[...selectedAssorts].sort().join(', ')}}`);
  const activeText = filterNotes.length > 0 ? ` (${{filterNotes.join(' | ')}})` : '';
  document.getElementById('rowCountText').textContent =
    `Showing ${{rows.length.toLocaleString()}} of ${{(currentTab === 'master' ? DATA.master.length : DATA.order.length).toLocaleString()}} rows${{activeText}}`;
}}

function exportCSV() {{
  const rows = getFilteredData();
  let csvRows = [];
  const esc = (v) => '"' + String(v ?? '').replace(/"/g, '""') + '"';

  if (currentTab === 'master') {{
    const headers = ['SKU #', 'Description', 'Water Type', 'Category #', 'Category Description', 'Assortment Code', 'Vendor Count', 'Available Vendors', ...DATA.vendors];
    csvRows.push(headers.map(esc).join(','));
    rows.forEach(r => {{
      const line = [
        r.sku, r.desc, r.water_type, r.cat_num, r.cat_desc, r.assort, r.vendors.length, r.vendors.join(', '),
        ...DATA.vendors.map(v => r.vendors.includes(v) ? 'Yes' : '')
      ];
      csvRows.push(line.map(esc).join(','));
    }});
  }} else {{
    const headers = ['SKU', 'Description', 'Vendor #', 'Vendor Name', 'Water Type', 'Assortment Code', 'Category Description', 'All Available Vendors'];
    csvRows.push(headers.map(esc).join(','));
    rows.forEach(r => {{
      const line = [r.sku, r.desc, r.v_num, r.v_name, r.water_type, r.assort, r.cat_desc, r.all_vendors];
      csvRows.push(line.map(esc).join(','));
    }});
  }}

  const blob = new Blob([csvRows.join('\\n')], {{ type: 'text/csv;charset=utf-8;' }});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = currentTab === 'master' ? 'Master_SKU_List.csv' : 'Cleaned_Order_Sheet.csv';
  a.click();
  URL.revokeObjectURL(url);
}}

loadSavedEdits();
initFilters();
render();
</script>
</body>
</html>
"""
    with open(HTML_FILE, "w", encoding="utf-8") as f:
        f.write(html)


if __name__ == "__main__":
    main()
