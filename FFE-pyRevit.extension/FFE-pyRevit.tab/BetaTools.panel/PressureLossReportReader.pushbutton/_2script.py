# -*- coding: utf-8 -*-
"""
Read tables from a "Duct Pressure Loss Report" HTML export and write CSVs.

Tested against the HTML structure shown in the uploaded sample
(including nested <Table class="report-embedded-table"> tables).  :contentReference[oaicite:1]{index=1}

Outputs (written next to the selected .html file):
- <name>__tables_raw.csv                 (every table, flattened)
- <name>__project_info.csv               (2-col key/value if found)
- <name>__system_info.csv                (2-col key/value if found)
- <name>__tpl_sections.csv               (Total Pressure Loss by Sections, normalized)
- <name>__straight_segments.csv          (Straight Segment detail, normalized)
- <name>__fittings_accessories.csv       (Fitting/Accessory K summary, normalized)
"""

from __future__ import print_function

import os
import re
import csv
import codecs

from HTMLParser import HTMLParser  # IronPython 2.7 compatible

from pyrevit import script

# ----------------------------
# Helpers
# ----------------------------

_ws_re = re.compile(r"\s+", re.UNICODE)

def clean_text(s):
    if s is None:
        return u""
    # HTMLParser already resolves many entities via handle_entityref/charref below
    s = s.replace(u"\xa0", u" ")  # &nbsp;
    s = _ws_re.sub(u" ", s)
    return s.strip()

def safe_filename(s):
    return re.sub(r"[^\w\-\.]+", "_", s)

def write_csv(path, headers, rows):
    # IronPython: use codecs for UTF-8
    with codecs.open(path, "w", "utf-8") as f:
        w = csv.writer(f)
        if headers:
            w.writerow([h for h in headers])
        for r in rows:
            w.writerow([u"" if v is None else v for v in r])

# ----------------------------
# Minimal DOM for tables
# ----------------------------

class TableNode(object):
    def __init__(self, depth):
        self.depth = depth
        self.rows = []          # list[list[text]]
        self.current_row = None
        self.current_cell = None
        self._cell_text = []
        self.caption_hint = u"" # inferred from surrounding headings

    def start_row(self):
        self.current_row = []
        self.rows.append(self.current_row)

    def start_cell(self):
        self._cell_text = []

    def add_text(self, t):
        if t:
            self._cell_text.append(t)

    def end_cell(self):
        txt = clean_text(u"".join(self._cell_text))
        if self.current_row is None:
            self.start_row()
        self.current_row.append(txt)

    def is_empty(self):
        if not self.rows:
            return True
        for r in self.rows:
            for c in r:
                if clean_text(c):
                    return False
        return True

# ----------------------------
# HTML parser
# ----------------------------

class PressureLossHTMLParser(HTMLParser):
    def __init__(self):
        HTMLParser.__init__(self)
        self.tables = []            # all TableNode in document order
        self._table_stack = []      # stack[TableNode]
        self._in_th_or_td = False
        self._pending_heading = u"" # last seen <a class="report-title/subtitle"> or <th colspan=...>

        # track <a class="report-title"> / <a class="report-subtitle">
        self._in_a = False
        self._a_class = None
        self._a_text = []

        # track heading th text for next table
        self._in_th = False
        self._th_text = []

    def handle_starttag(self, tag, attrs):
        t = tag.lower()
        attrs_d = dict((k.lower(), v) for (k, v) in attrs)

        if t == "a":
            self._in_a = True
            self._a_class = (attrs_d.get("class") or "").lower()
            self._a_text = []

        if t == "th":
            self._in_th = True
            self._th_text = []

        if t == "table":
            node = TableNode(depth=len(self._table_stack))
            node.caption_hint = clean_text(self._pending_heading)
            self.tables.append(node)
            self._table_stack.append(node)
            # once used, clear pending heading so it doesn't smear across unrelated tables
            self._pending_heading = u""

        if t == "tr" and self._table_stack:
            self._table_stack[-1].start_row()

        if t in ("td", "th") and self._table_stack:
            self._in_th_or_td = True
            self._table_stack[-1].start_cell()

    def handle_endtag(self, tag):
        t = tag.lower()

        if t == "a" and self._in_a:
            a_text = clean_text(u"".join(self._a_text))
            # Use title/subtitle text as a hint for the next table
            if a_text:
                if "report-title" in (self._a_class or ""):
                    self._pending_heading = a_text
                elif "report-subtitle" in (self._a_class or ""):
                    self._pending_heading = a_text
            self._in_a = False
            self._a_class = None
            self._a_text = []

        if t == "th" and self._in_th:
            th_text = clean_text(u"".join(self._th_text))
            # Many of these reports put a section heading in a <th colspan="3"> row
            # We treat that as the heading for the next table if it looks like a title
            if th_text and len(th_text) >= 6:
                # Only set if not inside a table cell capture
                # (but this still works fine as a "next table" hint)
                self._pending_heading = th_text
            self._in_th = False
            self._th_text = []

        if t in ("td", "th") and self._table_stack and self._in_th_or_td:
            self._table_stack[-1].end_cell()
            self._in_th_or_td = False

        if t == "table" and self._table_stack:
            self._table_stack.pop()

    def handle_data(self, data):
        if not data:
            return
        txt = data

        if self._in_a:
            self._a_text.append(txt)

        if self._in_th:
            self._th_text.append(txt)

        if self._table_stack and self._in_th_or_td:
            self._table_stack[-1].add_text(txt)

    def handle_entityref(self, name):
        # basic entity resolution
        m = {
            "nbsp": u"\xa0",
            "amp": u"&",
            "lt": u"<",
            "gt": u">",
            "quot": u"\"",
            "apos": u"'",
        }
        ch = m.get(name, u"")
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
# Normalizers for the known report tables
# ----------------------------

def normalize_kv_table(table):
    """If it's a 2-col key/value table, return dict; else None."""
    rows = table.rows or []
    kv = {}
    for r in rows:
        if len(r) >= 2:
            k = clean_text(r[0])
            v = clean_text(r[1])
            if k:
                kv[k] = v
    return kv if kv else None

def normalize_sections_with_embedded(table, expected_header):
    """
    Handles tables shaped like:
    [Section] [embedded-table of details] [Section Pressure Loss]
    Returns list of dict rows (one per embedded row).
    """
    rows = table.rows or []
    if not rows:
        return []

    # detect header row that contains "Section" and "Section Pressure Loss"
    # then embedded header row is usually inside the embedded table, which we can't see directly.
    # However: this parser flattens text per cell; embedded tables are parsed as separate TableNodes.
    # So we normalize using the OUTER table rows only (Section + Section Pressure Loss),
    # and then associate embedded tables by proximity separately in a simpler heuristic.

    # We will do a proximity association:
    # - The first row after the title block is the OUTER header row.
    # - Each subsequent row with a numeric Section in col0 corresponds to one embedded table that
    #   appears immediately after it in document order at the next depth level.
    return []  # normalization is done by a higher-level stitcher

def stitch_section_tables(all_tables):
    """
    Stitch the 3 major section tables by using:
    - outer title (caption_hint) OR first row/heading detection
    - embedded tables that follow each section row (depth+1)
    This is tuned for the uploaded report structure. :contentReference[oaicite:2]{index=2}
    """
    # Filter out empties
    tables = [t for t in all_tables if not t.is_empty()]

    # Identify the three outer tables by their caption hint
    def find_by_caption_contains(needle):
        for i, t in enumerate(tables):
            cap = (t.caption_hint or u"").lower()
            if needle.lower() in cap:
                return i, t
        return None, None

    idx_tpl, tpl_outer = find_by_caption_contains("total pressure loss calculations by sections")
    idx_ss, ss_outer = find_by_caption_contains("detail information of straight segment by sections")
    idx_fit, fit_outer = find_by_caption_contains("fitting and accessory loss coefficient summary by sections")

    # Also detect project/system info tables (caption hints are often empty; use first cell keys)
    project_table = None
    system_table = None
    for t in tables[:6]:  # early in doc
        if len(t.rows) >= 5 and len(t.rows[0]) == 2:
            k0 = (t.rows[0][0] or u"").lower()
            if "project name" in k0 and project_table is None:
                project_table = t
        if t.caption_hint and "system information" in t.caption_hint.lower():
            system_table = t

    # Build normalized records by walking outer table rows and grabbing the immediately following
    # embedded-table nodes at depth+1. In this export, embedded tables have class "report-embedded-table"
    # but we don't keep attributes; we rely on depth and proximity.
    def normalize_outer_with_embedded(outer_idx, outer_table, embedded_columns, total_col_name):
        if outer_table is None:
            return []

        out = []
        # In our parsed representation, the embedded table content does NOT live in the same cell;
        # it's represented as separate TableNodes that appear in document order after the outer row.
        # We assume one embedded table per section row.
        expected_depth = outer_table.depth + 1

        # The outer table rows: header row then multiple section rows. Section number is first col.
        # The third col is section total pressure loss (string).
        # Example: ["3", "", "0.02 in-wg"]  (middle col may be empty after flatten)
        section_rows = []
        for r in outer_table.rows:
            if not r:
                continue
            sec = clean_text(r[0]) if len(r) >= 1 else u""
            if sec.isdigit():
                section_rows.append(r)

        # Now find candidate embedded tables that occur after this outer table in list order
        # and have depth == expected_depth.
        embedded_candidates = []
        if outer_idx is not None:
            for t in tables[outer_idx+1:]:
                if t.depth == expected_depth:
                    embedded_candidates.append(t)
                # stop when we hit another outer table at same depth (heuristic)
                if t.depth == outer_table.depth and t is not outer_table:
                    break

        # Pair them up in order: len(section_rows) should match len(embedded_candidates) for this export
        pair_count = min(len(section_rows), len(embedded_candidates))
        for i in range(pair_count):
            sec_row = section_rows[i]
            sec_num = clean_text(sec_row[0]) if len(sec_row) > 0 else u""
            sec_total = clean_text(sec_row[2]) if len(sec_row) > 2 else u""
            emb = embedded_candidates[i]

            # first row of emb is the embedded header
            emb_rows = emb.rows or []
            if not emb_rows:
                continue
            emb_header = emb_rows[0]
            # if embedded header doesn't look like header, fall back to passed columns
            if len(emb_header) >= len(embedded_columns) and any(h.lower() in (emb_header[0] or u"").lower() for h in ("element", "element id")):
                cols = emb_header
            else:
                cols = embedded_columns

            for dr in emb_rows[1:]:
                rec = {"Section": sec_num, total_col_name: sec_total}
                for cidx, cname in enumerate(cols):
                    if cidx < len(dr):
                        rec[clean_text(cname)] = clean_text(dr[cidx])
                    else:
                        rec[clean_text(cname)] = u""
                out.append(rec)
        return out

    # Normalize each of the three tables
    tpl_records = normalize_outer_with_embedded(
        idx_tpl, tpl_outer,
        embedded_columns=["Element", "Flow", "Size", "Velocity", "Length", "Friction", "Total Pressure Loss"],
        total_col_name="Section Pressure Loss"
    )

    ss_records = normalize_outer_with_embedded(
        idx_ss, ss_outer,
        embedded_columns=["Element ID", "Type Mark", "Comments", "Size", "Flow", "Length", "Velocity", "Friction", "System Name", "Pressure Loss"],
        total_col_name="Total Pressure Loss"
    )

    fit_records = normalize_outer_with_embedded(
        idx_fit, fit_outer,
        embedded_columns=["Element ID", "Type Mark", "Comments", "ASHRAE Table", "Size", "System Name", "Pressure Loss"],
        total_col_name="Total Pressure Loss"
    )

    return project_table, system_table, tpl_records, ss_records, fit_records

# ----------------------------
# Main
# ----------------------------

logger = script.get_logger()
output = script.get_output()

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

if not os.path.exists(html_path):
    logger.error("File not found: {}".format(html_path))
    script.exit()

with codecs.open(html_path, "r", "utf-8") as f:
    html_text = f.read()

parser = PressureLossHTMLParser()
parser.feed(html_text)
parser.close()

# Filter to non-empty tables
tables = [t for t in parser.tables if not t.is_empty()]
if not tables:
    logger.error("No tables detected in the selected HTML.")
    script.exit()

base_dir = os.path.dirname(html_path)
base_name = os.path.splitext(os.path.basename(html_path))[0]
base_name = safe_filename(base_name)

# 1) Raw dump: each row is [table_index, depth, caption_hint, row_index, col_index, value]
raw_rows = []
for ti, t in enumerate(tables):
    cap = clean_text(t.caption_hint)
    for ri, r in enumerate(t.rows):
        for ci, v in enumerate(r):
            raw_rows.append([str(ti), str(t.depth), cap, str(ri), str(ci), clean_text(v)])

raw_path = os.path.join(base_dir, base_name + "__tables_raw.csv")
write_csv(raw_path, ["table_index", "depth", "caption_hint", "row_index", "col_index", "value"], raw_rows)

# 2) Normalized outputs
project_table, system_table, tpl_records, ss_records, fit_records = stitch_section_tables(tables)

if project_table:
    kv = normalize_kv_table(project_table)
    proj_path = os.path.join(base_dir, base_name + "__project_info.csv")
    write_csv(proj_path, ["Key", "Value"], [[k, kv.get(k, u"")] for k in sorted(kv.keys())])
else:
    proj_path = None

if system_table:
    kv = normalize_kv_table(system_table)
    sys_path = os.path.join(base_dir, base_name + "__system_info.csv")
    write_csv(sys_path, ["Key", "Value"], [[k, kv.get(k, u"")] for k in sorted(kv.keys())])
else:
    sys_path = None

def write_records(path, records):
    if not records:
        return
    # union headers in stable order
    headers = []
    seen = set()
    for rec in records:
        for k in rec.keys():
            if k not in seen:
                headers.append(k)
                seen.add(k)
    rows = []
    for rec in records:
        rows.append([rec.get(h, u"") for h in headers])
    write_csv(path, headers, rows)

tpl_path = os.path.join(base_dir, base_name + "__tpl_sections.csv")
ss_path = os.path.join(base_dir, base_name + "__straight_segments.csv")
fit_path = os.path.join(base_dir, base_name + "__fittings_accessories.csv")

write_records(tpl_path, tpl_records)
write_records(ss_path, ss_records)
write_records(fit_path, fit_records)

# Report
output.print_md("### Parsed tables from HTML")
output.print_md("* Source: `{}`".format(html_path))
output.print_md("* Tables detected: **{}**".format(len(tables)))
output.print_md("* Raw dump: `{}`".format(raw_path))

if proj_path:
    output.print_md("* Project info: `{}`".format(proj_path))
if sys_path:
    output.print_md("* System info: `{}`".format(sys_path))

output.print_md("* Total Pressure Loss by Sections (normalized): `{}` (rows: {})".format(tpl_path, len(tpl_records)))
output.print_md("* Straight Segments (normalized): `{}` (rows: {})".format(ss_path, len(ss_records)))
output.print_md("* Fittings/Accessories (normalized): `{}` (rows: {})".format(fit_path, len(fit_records)))
