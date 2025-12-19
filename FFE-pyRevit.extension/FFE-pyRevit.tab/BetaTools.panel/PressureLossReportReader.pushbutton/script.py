# -*- coding: utf-8 -*-
"""
Extract ONLY these 3 report tables from the Pressure Loss Report HTML:

Table1 = 'Detail Information of Straight Segment by Sections'
Table2 = 'Fitting and Accessory Loss Coefficient Summary by Sections'
Table3 = 'Total Pressure Loss Calculations by Sections'

This report stores the title INSIDE the outer table (as a <th colspan="3"> row).
So we identify the correct outer table by searching its own cells for the title.

Outputs (written next to the selected .html file):
- <name>__straight_segments.csv
- <name>__fittings_accessories.csv
- <name>__total_pressure_loss.csv

Also prints debug counts to pyRevit output so you can confirm matches.
"""

from __future__ import print_function

import os
import re
import csv
import codecs
from HTMLParser import HTMLParser  # IronPython 2.7 compatible

from pyrevit import script
from pyrevit.script import output

output_window = output.get_output()

# ----------------------------
# Utilities
# ----------------------------

_ws_re = re.compile(r"\s+", re.UNICODE)

def clean_text(s):
    if s is None:
        return ""
    s = s.replace("\xa0", " ")  # &nbsp;
    s = _ws_re.sub(" ", s)
    return s.strip()

def safe_filename(s):
    return re.sub(r"[^\w\-\.]+", "_", s)

def write_csv(path, headers, rows):
    with codecs.open(path, "w", "utf-8") as f:
        w = csv.writer(f)
        w.writerow([h for h in headers])
        for r in rows:
            w.writerow(["" if v is None else v for v in r])

def dicts_to_rows(dicts, preferred_order=None):
    if not dicts:
        return [], []
    keys = []
    seen = set()

    if preferred_order:
        for k in preferred_order:
            if any(k in d for d in dicts):
                keys.append(k)
                seen.add(k)

    for d in dicts:
        for k in d.keys():
            if k not in seen:
                keys.append(k)
                seen.add(k)

    rows = []
    for d in dicts:
        rows.append([d.get(k, u"") for k in keys])
    return keys, rows

# ----------------------------
# Table model + parser
# ----------------------------

class TableNode(object):
    def __init__(self, depth):
        self.depth = depth
        self.rows = []
        self.current_row = None
        self._cell_text = []

    def start_row(self):
        self.current_row = []
        self.rows.append(self.current_row)

    def start_cell(self):
        self._cell_text = []

    def add_text(self, t):
        if t:
            self._cell_text.append(t)

    def end_cell(self):
        txt = clean_text("".join(self._cell_text))
        if self.current_row is None:
            self.start_row()
        self.current_row.append(txt)

    def is_empty(self):
        for r in self.rows:
            for c in r:
                if clean_text(c):
                    return False
        return True

    def contains_text(self, needle):
        needle_l = clean_text(needle).lower()
        for r in self.rows:
            for c in r:
                if clean_text(c).lower() == needle_l:
                    return True
        return False

class ReportHTMLParser(HTMLParser):
    def __init__(self):
        HTMLParser.__init__(self)
        self.tables = []
        self._table_stack = []
        self._in_cell = False

    def handle_starttag(self, tag, attrs):  
        t = tag.lower()
        if t == "table":
            node = TableNode(depth=len(self._table_stack))
            self.tables.append(node)
            self._table_stack.append(node)

        if t == "tr" and self._table_stack:
            self._table_stack[-1].start_row()

        if t in ("td", "th") and self._table_stack:
            self._in_cell = True
            self._table_stack[-1].start_cell()

    def handle_endtag(self, tag):
        t = tag.lower()

        if t in ("td", "th") and self._table_stack and self._in_cell:
            self._table_stack[-1].end_cell()
            self._in_cell = False

        if t == "table" and self._table_stack:
            self._table_stack.pop()

    def handle_data(self, data):
        if not data:
            return
        if self._table_stack and self._in_cell:
            self._table_stack[-1].add_text(data)

    def handle_entityref(self, name):
        # m = {"nbsp": u"\xa0", "amp": u"&", "lt": u"<", "gt": u">", "quot": u"\"", "apos": u"'"}
        m = {"nbsp": "\xa0", "amp": "&", "lt": "<", "gt": ">", "quot": "\"", "apos": "'"}
        # ch = m.get(name, u"")
        ch = m.get(name, "")
        if ch:
            self.handle_data(ch)

    def handle_charref(self, name):
        try:
            if name.lower().startswith("x"):
                ch = unichr(int(name[1:], 16))
            else:
                ch = unichr(int(name))
            self.handle_data(ch)
        except:
            pass

# ----------------------------
# Extraction logic
# ----------------------------

TABLE1 = "Total Pressure Loss Calculations by Sections"
TABLE2 = "Detail Information of Straight Segment by Sections"
TABLE3 = "Fitting and Accessory Loss Coefficient Summary by Sections"

def find_outer_table_by_title_cell(tables, title):
    """Return (index, table) where any cell exactly equals title."""
    for i, t in enumerate(tables):
        if t.contains_text(title):
            return i, t
    return None, None

def iter_section_rows(outer_table):
    """
    Yield (section_number, section_total_loss) for each section row in the OUTER table.
    Outer row pattern: [Section] [embedded table in 2nd col] [Total loss in 3rd col]
    """
    for r in outer_table.rows:
        if not r:
            continue
        sec = clean_text(r[0]) if len(r) >= 1 else u""
        if sec.isdigit():
            sec_total = clean_text(r[2]) if len(r) > 2 else u""
            yield sec, sec_total

def collect_embedded_tables_after(tables, outer_index, outer_depth):
    """
    Embedded tables are parsed as their own TableNodes at depth = outer_depth + 1,
    and appear after the outer table start in document order.
    We collect depth+1 tables until the next depth==outer_depth table appears.
    """
    # child_depth = outer_depth + 1
    child_depth = outer_depth
    embedded = []
    if outer_index is None:
        return embedded

    for t in tables[outer_index + 1:]:
        if t.depth == child_depth and not t.is_empty():
            embedded.append(t)
            continue
        if t.depth == outer_depth:
            break
    return embedded

def stitch_outer_to_embedded(outer, embedded_tables, section_total_col_name):
    """
    Pair each section row with one embedded table (by order).
    Return list[dict] with 'Section' + section_total_col_name + embedded table columns.
    """
    section_rows = list(iter_section_rows(outer))
    pair_count = min(len(section_rows), len(embedded_tables))

    records = []
    for i in range(pair_count):
        sec, sec_total = section_rows[i]
        emb = embedded_tables[i]
        if not emb.rows:
            continue

        header = emb.rows[0]
        if not any(clean_text(h) for h in header):
            max_len = max(len(r) for r in emb.rows[1:]) if len(emb.rows) > 1 else 0
            header = [u"Col{}".format(j+1) for j in range(max_len)]

        for data_row in emb.rows[1:]:
            rec = {u"Section": sec, section_total_col_name: sec_total}
            for cidx, cname in enumerate(header):
                key = clean_text(cname) if clean_text(cname) else u"Col{}".format(cidx+1)
                rec[key] = clean_text(data_row[cidx]) if cidx < len(data_row) else u""
            records.append(rec)

    return records, len(section_rows), len(embedded_tables), pair_count

# ----------------------------
# Main
# ----------------------------

logger = script.get_logger()
# output = script.get_output()

try:
    from System.Windows.Forms import OpenFileDialog, DialogResult
except Exception:
    OpenFileDialog = None

def pick_html_file():
    if OpenFileDialog is None:
        return None
    dlg = OpenFileDialog()
    dlg.Filter = "HTML Files (*.html;*.htm)|*.html;*.htm|All Files (*.*)|*.*"
    dlg.Title = "Select Pressure Loss Report HTML"
    if dlg.ShowDialog() == DialogResult.OK:
        return dlg.FileName
    return None

html_path = pick_html_file()
if not html_path:
    script.exit()

with codecs.open(html_path, "r", "utf-8") as f:
    html_text = f.read()

parser = ReportHTMLParser()
parser.feed(html_text)
parser.close()

tables = [t for t in parser.tables if not t.is_empty()]
# print("2")#### <- TESTING 
# print(tables[2].rows)#### <- TESTING 


if not tables:
    logger.error("No tables detected in the selected HTML.")
    script.exit()


# START WORKING AREA >>>
######################################################
######################################################
######################################################
TotalLossbySection = tables[2].rows

# Get Critical Paths 
CriticalPaths = TotalLossbySection[-1][0].split(' ')[3].split('-')

output_window.print_md("**CP:** {}".format(CriticalPaths))


Duct_SectionsSummary = []
Fittings_SectionsSummary = []
Duct_HTML_Index = None
Fitting_HTML_Index = None
n=0
# Extract Duct and Fitting Tables
for t in tables:
    table_rows = t.rows 
    for row in table_rows:
        if "Detail Information of Straight Segment by Sections" in row:
            Duct_Sections = table_rows[1:]       # Collect Duct Section and Total Pressure Loss Table
            Duct_HTML_Index = n
        if "Fitting and Accessory Loss Coefficient Summary by Sections" in row:
            Fitting_Sections = table_rows[1:]    # Collect Fitting Section and Total Pressure Loss Table
            Fitting_HTML_Index = n
    
        if Duct_HTML_Index is not None and Fitting_HTML_Index is None:
            Duct_SectionsSummary.append(row)
        if Fitting_HTML_Index is not None:
            Fittings_SectionsSummary.append(row)

    n = n + 1


# Separate Duct and Fitting Section Tables
d_section_tablenode = tables[Duct_HTML_Index+1:Fitting_HTML_Index]
f_section_tablenode = tables[Fitting_HTML_Index+1:]


DuctReport = []
for tn, s in zip(d_section_tablenode, Duct_Sections):
    t_rows = tn.rows
    if t_rows not in Duct_Sections:
        for t in t_rows:
            t.insert(0, s[0])  # Insert Section Number at position 0
            DuctReport.append(t)

FittingsReport = []
for tn, s in zip(f_section_tablenode, Fitting_Sections):
    t_rows = tn.rows
    if t_rows not in Fitting_Sections:
        for t in t_rows:
            t.insert(0, s[0])  # Insert Section Number at position 0
            FittingsReport.append(t)



# Create dict
dict_DuctReport     = {"Header": DuctReport[0], "Data": DuctReport[1:]}
dict_FittingsReport  = {"Header": FittingsReport[0], "Data": FittingsReport[1:]}


"""
# Use this for inserting Flow values into Fittings Report
dict_DuctReport["Header"].insert(0, "Section")
for data_row in dict_DuctReport["Data"]:
    section_value = ""
    element_id = data_row[dict_DuctReport["Header"].index("Element ID")]
    for d_section in Duct_Sections:
        if element_id in d_section[1]:
            section_value = d_section[0]
            break
    data_row.insert(0, section_value)
"""

# Insert Category Column
dict_DuctReport["Header"].insert(1, "Category")
for data_row in dict_DuctReport["Data"]:
    data_row.insert(1, "Duct")

dict_FittingsReport["Header"].insert(1, "Category")
for data_row in dict_FittingsReport["Data"]:
    data_row.insert(1, "Fittings")



#____________________________________________________________________ OUTPUT TABLES
output_window.print_table(table_data=dict_DuctReport["Data"], columns=dict_DuctReport["Header"], title="Duct Report")
output_window.print_table(table_data=dict_FittingsReport["Data"], columns=dict_FittingsReport["Header"], title="Fittings Report")


ColumnIndex = ["System Name", "Category", "Element ID", "Type Mark", "ASHRAE Table", "Comments", "Section", "Size", "Flow", "Length", "Velocity", "Friction", "Pressure Loss"]

######################################################
######################################################
######################################################
# <<< END WORKING AREA


base_dir = os.path.dirname(html_path)
base_name = safe_filename(os.path.splitext(os.path.basename(html_path))[0])

def extract_one(title, out_suffix, section_total_col_name, preferred_order):
    idx, outer = find_outer_table_by_title_cell(tables, title)
    # output_window.print_md("**idx:** {}".format(idx))
    # output_window.print_md("**outer:** {}".format(outer))

    if outer is None:
        return None, 0, 0, 0, 0

    embedded = collect_embedded_tables_after(tables, idx, outer.depth)
    records, sec_cnt, emb_cnt, paired = stitch_outer_to_embedded(outer, embedded, section_total_col_name)
    # print(records)#### <- TESTING 
    headers, rows = dicts_to_rows(records, preferred_order=preferred_order)

    output_window.print_md("**Title:** {}".format(title))#### <- TESTING
    output_window.print_md("**Headers:** {}".format(headers))#### <- TESTING
    output_window.print_md("**Rows:** {}".format(rows))#### <- TESTING

    out_path = os.path.join(base_dir, base_name + out_suffix)
    # if headers and rows:
        # write_csv(out_path, headers, rows)
        # print(rows)#### <- TESTING
    # else:
    #     # still write headers if we have them; otherwise write a small diagnostic file
    #     if headers:
    #         write_csv(out_path, headers, [])
    #     else:
    #         write_csv(out_path, ["ERROR"], ["No data rows captured"])

    return out_path, sec_cnt, emb_cnt, paired, len(records)

"""
# Table1
t1_path, t1_sec, t1_emb, t1_paired, t1_rows = extract_one(
    TABLE1,
    "__total_pressure_loss.csv",
    section_total_col_name="Section Pressure Loss",
    preferred_order=["Section", "Element", "Flow", "Size", "Velocity", "Length", "Friction",
                     "Total Pressure Loss", "Section Pressure Loss"]
    # preferred_order=[u"Section", u"Total Pressure Loss", u"Element ID", u"Type Mark", u"Comments", u"Size", u"Flow",
    #                  u"Length", u"Velocity", u"Friction", u"System Name", u"Pressure Loss"]
)

# Table2
t2_path, t2_sec, t2_emb, t2_paired, t2_rows = extract_one(
    TABLE2,
    "__straight_segments.csv",
    section_total_col_name="Total Pressure Loss",
    preferred_order=["Section", "Element ID", "Type Mark", "Comments", "Size", "Flow", "Length",
                     "Velocity", "Friction", "System Name", "Pressure Loss", "Total Pressure Loss"]
    # preferred_order=[u"Section", u"Total Pressure Loss", u"Element ID", u"Type Mark", u"Comments", u"ASHRAE Table",
    #                  u"Size", u"System Name", u"Pressure Loss"]
)

# Table3
t3_path, t3_sec, t3_emb, t3_paired, t3_rows = extract_one(
    TABLE3,
    "__fittings_accessories.csv",
    section_total_col_name="Total Pressure Loss",
    preferred_order=["Section", "Element ID", "Type Mark", "Comments", "ASHRAE Table", "Size",
                     "System Name", "Pressure Loss", "Total Pressure Loss"]
    # preferred_order=[u"Section", u"Section Pressure Loss", u"Element", u"Flow", u"Size", u"Velocity",
    #                  u"Length", u"Friction", u"Total Pressure Loss"]
)
"""
# output_window.print_md("### Pressure Loss Report — Extraction Results")
# output_window.print_md("* Source: `{}`".format(html_path))
# output_window.print_md("* Non-empty tables parsed: **{}**".format(len(tables)))

def report(name, path, sec_cnt, emb_cnt, paired, rows):
    if not path:
        output_window.print_md("* **{}**: Not found (title row not detected)".format(name))
        return
    output_window.print_md("* **{}** → `{}`".format(name, path))
    output_window.print_md("  - Section rows found: **{}**".format(sec_cnt))
    output_window.print_md("  - Embedded tables found: **{}**".format(emb_cnt))
    output_window.print_md("  - Section↔Embedded pairs used: **{}**".format(paired))
    output_window.print_md("  - Output rows written: **{}**".format(rows))

# report("Table1 (Straight Segments)", t1_path, t1_sec, t1_emb, t1_paired, t1_rows)
# report("Table2 (Fittings/Accessories)", t2_path, t2_sec, t2_emb, t2_paired, t2_rows)
# report("Table3 (Total Pressure Loss)", t3_path, t3_sec, t3_emb, t3_paired, t3_rows)
