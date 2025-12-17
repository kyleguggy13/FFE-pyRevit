# -*- coding: utf-8 -*-
"""
Pressure Loss Report (HTML) -> CSV
IronPython 2.7 / pyRevit compatible (NO pandas, NO numpy)

Re-implements the logic from:
- HTMLparse.py :contentReference[oaicite:2]{index=2}
- Reader_Pressure Loss Report.py :contentReference[oaicite:3]{index=3}

Targets:
- Detail Information of Straight Segment by Sections
- Fitting and Accessory Loss Coefficient Summary by Sections
- Total Pressure Loss Calculations by Sections

Output:
- <htmlname>__ductdata_all.csv
"""

from __future__ import print_function

import os
import re
import csv
import codecs
from HTMLParser import HTMLParser

from pyrevit import script

# -------------------------------------------------------------------
# Configuration (table title strings)
# -------------------------------------------------------------------

MATCH_DUCT = u"Detail Information of Straight Segment by Sections"
MATCH_FIT  = u"Fitting and Accessory Loss Coefficient Summary by Sections"
MATCH_CP   = u"Total Pressure Loss Calculations by Sections"

# Desired output column order (mirrors your Pandas ColumnIndex) :contentReference[oaicite:4]{index=4}
COLUMN_ORDER = [
    u"System Name", u"Category", u"Element ID", u"Type Mark", u"ASHRAE Table",
    u"Critical Path", u"Section", u"Size", u"Flow", u"Length",
    u"Velocity", u"Friction", u"Pressure Loss"
]

# -------------------------------------------------------------------
# Helpers
# -------------------------------------------------------------------

_ws = re.compile(r"\s+", re.UNICODE)

def clean_text(s):
    if s is None:
        return u""
    s = s.replace(u"\xa0", u" ")  # nbsp
    s = _ws.sub(u" ", s)
    return s.strip()

def to_float(s):
    s = clean_text(s)
    if not s:
        return None
    try:
        return float(s)
    except:
        return None

def to_int(s):
    s = clean_text(s)
    if not s:
        return None
    try:
        return int(float(s))
    except:
        return None

def strip_suffix(s, suffix):
    s = clean_text(s)
    if s.endswith(suffix):
        return clean_text(s[:-len(suffix)])
    return s

def is_nullish(s):
    return clean_text(s) == u""

def safe_filename(s):
    return re.sub(r"[^\w\-\.]+", "_", s)

def write_csv(path, headers, row_dicts):
    # Write dict rows in header order; blank for missing
    with codecs.open(path, "w", "utf-8") as f:
        w = csv.writer(f)
        w.writerow([h for h in headers])
        for d in row_dicts:
            w.writerow([d.get(h, u"") for h in headers])

# -------------------------------------------------------------------
# HTML table extraction (no external libs)
# -------------------------------------------------------------------

class _Table(object):
    def __init__(self):
        self.rows = []      # list[list[str]]
        self._row = None
        self._cell_buf = []
        self._in_cell = False

    def start_row(self):
        self._row = []
        self.rows.append(self._row)

    def start_cell(self):
        self._cell_buf = []
        self._in_cell = True

    def add_cell_text(self, txt):
        if self._in_cell:
            self._cell_buf.append(txt)

    def end_cell(self):
        val = clean_text(u"".join(self._cell_buf))
        if self._row is None:
            self.start_row()
        self._row.append(val)
        self._in_cell = False

    def all_text(self):
        # Flatten for matching
        parts = []
        for r in self.rows:
            for c in r:
                if c:
                    parts.append(c)
        return u" ".join(parts)

    def is_empty(self):
        for r in self.rows:
            for c in r:
                if clean_text(c):
                    return False
        return True

class TableHTMLParser(HTMLParser):
    def __init__(self):
        HTMLParser.__init__(self)
        self.tables = []
        self._stack = []
        self._cur = None

    def handle_starttag(self, tag, attrs):
        t = tag.lower()
        if t == "table":
            self._cur = _Table()
            self.tables.append(self._cur)
            self._stack.append(self._cur)

        if t == "tr" and self._stack:
            self._stack[-1].start_row()

        if t in ("td", "th") and self._stack:
            self._stack[-1].start_cell()

    def handle_data(self, data):
        if self._stack:
            self._stack[-1].add_cell_text(data)

    def handle_entityref(self, name):
        m = {"nbsp": u"\xa0", "amp": u"&", "lt": u"<", "gt": u">", "quot": u"\"", "apos": u"'"}
        ch = m.get(name, u"")
        if ch and self._stack:
            self._stack[-1].add_cell_text(ch)

    def handle_charref(self, name):
        try:
            if name.lower().startswith("x"):
                ch = unichr(int(name[1:], 16))
            else:
                ch = unichr(int(name))
            if self._stack:
                self._stack[-1].add_cell_text(ch)
        except:
            pass

    def handle_endtag(self, tag):
        t = tag.lower()
        if t in ("td", "th") and self._stack:
            self._stack[-1].end_cell()

        if t == "table" and self._stack:
            self._stack.pop()

# -------------------------------------------------------------------
# Locate and normalize the three target tables (multi-system support)
# -------------------------------------------------------------------

def find_tables_by_match(all_tables, match_text):
    """Return list of tables whose flattened text contains match_text (case-insensitive)."""
    out = []
    needle = clean_text(match_text).lower()
    for tb in all_tables:
        if tb.is_empty():
            continue
        if needle in tb.all_text().lower():
            out.append(tb)
    return out

def find_header_row_index(rows, required_headers):
    """
    Find the first row that contains ALL required headers (case-insensitive).
    Returns index or None.
    """
    req = [h.lower() for h in required_headers]
    for i, r in enumerate(rows):
        row_l = [clean_text(c).lower() for c in r]
        ok = True
        for h in req:
            if h not in row_l:
                ok = False
                break
        if ok:
            return i
    return None

def normalize_table_to_dict_rows(tb, required_headers):
    """
    Convert a parsed table to (headers, dict_rows) by discovering the header row.
    Pads rows to header length.
    """
    rows = tb.rows
    hidx = find_header_row_index(rows, required_headers)
    if hidx is None:
        return [], []

    headers = [clean_text(c) for c in rows[hidx]]
    data_rows = rows[hidx + 1:]

    dict_rows = []
    hlen = len(headers)
    for r in data_rows:
        # pad
        rr = list(r) + [u""] * max(0, hlen - len(r))
        rr = rr[:hlen]
        # skip completely empty
        if all(is_nullish(x) for x in rr):
            continue
        d = {}
        for j in range(hlen):
            d[headers[j]] = rr[j]
        dict_rows.append(d)

    return headers, dict_rows

# -------------------------------------------------------------------
# Re-implementation of HTMLparse() without Pandas :contentReference[oaicite:5]{index=5}
# -------------------------------------------------------------------

def get_critical_path_list(cp_table_dict_rows):
    """
    Your pandas version:
        CriticalPath = HTMLtable.iloc[-1,0].split(' ')[3].split('-') :contentReference[oaicite:6]{index=6}
    Here: search any cell containing 'Critical Path' and parse the same tokenization.
    """
    # fallback: use last non-empty cell in last non-empty row
    best = u""
    for d in reversed(cp_table_dict_rows):
        for k, v in d.items():
            if clean_text(v):
                best = v
                break
        if best:
            break

    # Prefer explicit "Critical Path" cell
    for d in cp_table_dict_rows:
        for v in d.values():
            vv = clean_text(v)
            if "critical path" in vv.lower():
                best = vv

    if not best:
        return []

    parts = best.split()
    if len(parts) < 4:
        return []
    seg = parts[3]  # same assumption as your original code
    return [p for p in seg.split("-") if p]

def build_section_map(table_dict_rows, element_id_key=u"Element ID"):
    """
    Recreates:
      df_DuctSections = dfTable_Duct.loc[dfTable_Duct.iloc[:,-1].isnull()]
      df_DuctReport   = dfTable_Duct.loc[dfTable_Duct.iloc[:,-1].notnull()]

    Using dict rows:
    - "Section rows": last column is blank (nullish) -> contain Section + Details text
    - "Report rows": last column not blank -> actual element rows

    Then maps Element ID -> Section by checking if the Element ID appears inside Details. :contentReference[oaicite:7]{index=7}
    """
    if not table_dict_rows:
        return {}, [], []

    # determine "last column" name from first row's keys order
    # (dicts are unordered in IronPython; derive from original headers elsewhere)
    # Instead: detect last column by finding a key that looks like "Mark" or any sparse tail column.
    # Practical approach: define last column as the last non-empty key in the row's keys sorted by insertion is not possible.
    # So we detect "section rows" by: element_id_key empty -> section row; element_id_key non-empty -> report row
    # BUT your file: section rows often have Section in first column and a big Details string in second, while Element ID is blank.
    # We'll use: if Element ID is nullish => section row; else report row.
    section_rows = []
    report_rows = []

    for r in table_dict_rows:
        eid = clean_text(r.get(element_id_key, u""))
        if is_nullish(eid):
            section_rows.append(r)
        else:
            report_rows.append(r)

    # Build Section -> Details blob list (some sections may have multiple lines)
    # Use 'Section' and the next column as "Details" (best-effort)
    # Many exports include the Section number under header 'Section'
    # and the details in a column header like 'Details' or blank.
    # We'll heuristically pick the longest text field in the row (besides Section) as Details.
    section_to_details = {}
    for r in section_rows:
        sec = clean_text(r.get(u"Section", u""))
        if not sec:
            # Sometimes Section may be under another header; try first numeric-looking value
            for v in r.values():
                vv = clean_text(v)
                if vv.isdigit():
                    sec = vv
                    break
        if not sec:
            continue

        # pick details as the longest non-empty value that's not the section itself
        details = u""
        for v in r.values():
            vv = clean_text(v)
            if vv and vv != sec and len(vv) > len(details):
                details = vv

        if sec not in section_to_details:
            section_to_details[sec] = []
        if details:
            section_to_details[sec].append(details)

    # Map Element ID -> Section by searching details blobs
    eid_to_section = {}
    for sec, blobs in section_to_details.items():
        joined = u" ".join(blobs)
        # We'll fill eid_to_section later by checking each report row's Element ID
        # (more efficient than trying to extract IDs from details)
        section_to_details[sec] = joined

    for rr in report_rows:
        eid = clean_text(rr.get(element_id_key, u""))
        if not eid:
            continue
        # find first section whose details contains this eid
        for sec, blob in section_to_details.items():
            if eid in blob:
                eid_to_section[eid] = sec
                break

    return eid_to_section, section_rows, report_rows

def coerce_units_on_rows(report_rows, is_duct):
    """
    Recreates your numeric conversions (strip units and cast). :contentReference[oaicite:8]{index=8}
    Returns new list of dict rows with cleaned numbers as strings (CSV-friendly).
    """
    out = []
    for r in report_rows:
        rr = dict(r)

        # Velocity: "#### FPM"
        if u"Velocity" in rr:
            v = strip_suffix(rr.get(u"Velocity"), u" FPM")
            iv = to_int(v)
            rr[u"Velocity"] = u"" if iv is None else unicode(iv)

        # Friction: "x.xx in-wg/100ft"
        if u"Friction" in rr:
            f = strip_suffix(rr.get(u"Friction"), u" in-wg/100ft")
            fv = to_float(f)
            rr[u"Friction"] = u"" if fv is None else unicode(fv)

        # Flow: "### CFM"
        if u"Flow" in rr:
            fl = strip_suffix(rr.get(u"Flow"), u" CFM")
            flv = to_int(fl)
            rr[u"Flow"] = u"" if flv is None else unicode(flv)

        # Pressure Loss: "x.xxx in-wg"
        if u"Pressure Loss" in rr:
            pl = strip_suffix(rr.get(u"Pressure Loss"), u" in-wg")
            plv = to_float(pl)
            rr[u"Pressure Loss"] = u"" if plv is None else unicode(plv)

        out.append(rr)
    return out

def merge_duct_and_fittings(duct_rows, fit_rows, critical_sections):
    """
    Recreates:
      - add Category
      - filter CP subsets
      - concat
      - for each section, set Flow = max(Flow) across the section (applied to fittings too)
      - rename Comments -> Critical Path :contentReference[oaicite:9]{index=9}
    """
    all_rows = []

    for r in duct_rows:
        rr = dict(r)
        rr[u"Category"] = u"Duct"
        all_rows.append(rr)

    for r in fit_rows:
        rr = dict(r)
        rr[u"Category"] = u"Fitting"
        all_rows.append(rr)

    # Rename Comments -> Critical Path (keep original if already exists)
    for rr in all_rows:
        if u"Critical Path" not in rr:
            rr[u"Critical Path"] = rr.get(u"Comments", u"")
        # optionally remove Comments to avoid duplication
        if u"Comments" in rr:
            del rr[u"Comments"]

    # Compute max Flow per Section (string ints)
    sec_to_max_flow = {}
    for rr in all_rows:
        sec = clean_text(rr.get(u"Section", u""))
        flow = to_int(rr.get(u"Flow", u""))
        if not sec:
            continue
        if flow is None:
            continue
        if sec not in sec_to_max_flow or flow > sec_to_max_flow[sec]:
            sec_to_max_flow[sec] = flow

    # Apply max Flow to all rows in section (including fittings)
    for rr in all_rows:
        sec = clean_text(rr.get(u"Section", u""))
        if sec in sec_to_max_flow:
            rr[u"Flow"] = unicode(sec_to_max_flow[sec])

    # Optionally build CP list (not written separately here; easy to add)
    # cp_rows = [r for r in all_rows if clean_text(r.get(u"Section", u"")) in critical_sections]

    return all_rows

def reindex_and_project_columns(rows):
    """Ensure every output row has the same columns in COLUMN_ORDER."""
    out = []
    for r in rows:
        rr = {}
        for k in COLUMN_ORDER:
            rr[k] = clean_text(r.get(k, u""))
        out.append(rr)
    return out

# -------------------------------------------------------------------
# Main: pick file, parse, extract per-system tables, transform, write CSV
# -------------------------------------------------------------------

logger = script.get_logger()
output = script.get_output()

# file picker
try:
    from System.Windows.Forms import OpenFileDialog, DialogResult
except Exception:
    OpenFileDialog = None

def pick_html():
    if OpenFileDialog is None:
        return None
    dlg = OpenFileDialog()
    dlg.Filter = "HTML Files (*.html;*.htm)|*.html;*.htm|All Files (*.*)|*.*"
    dlg.Title = "Select Pressure Loss Report HTML"
    if dlg.ShowDialog() == DialogResult.OK:
        return dlg.FileName
    return None

html_path = pick_html()
if not html_path:
    script.exit()

with codecs.open(html_path, "r", "utf-8") as f:
    html_text = f.read()

p = TableHTMLParser()
p.feed(html_text)
p.close()

tables = [t for t in p.tables if not t.is_empty()]

# Find all occurrences of each table type (system-by-system)
duct_tables = find_tables_by_match(tables, MATCH_DUCT)
fit_tables  = find_tables_by_match(tables, MATCH_FIT)
cp_tables   = find_tables_by_match(tables, MATCH_CP)

# Normalize each table to dict rows by finding the real header row
# Duct/Fit must contain 'Element ID' to identify header row
# CP must contain 'Section' and 'Section Pressure Loss' usually
SYSTEM_ROWS_ALL = []

sys_count = min(len(duct_tables), len(fit_tables), len(cp_tables))
if sys_count == 0:
    logger.error("Could not find matching tables for one or more targets. "
                 "Found: Duct=%s, Fittings=%s, CP=%s" % (len(duct_tables), len(fit_tables), len(cp_tables)))
    script.exit()

for i in range(sys_count):
    duct_tb = duct_tables[i]
    fit_tb  = fit_tables[i]
    cp_tb   = cp_tables[i]

    duct_headers, duct_dict_rows = normalize_table_to_dict_rows(duct_tb, required_headers=[u"Element ID", u"Pressure Loss"])
    fit_headers,  fit_dict_rows  = normalize_table_to_dict_rows(fit_tb,  required_headers=[u"Element ID", u"Pressure Loss"])
    cp_headers,   cp_dict_rows   = normalize_table_to_dict_rows(cp_tb,   required_headers=[u"Section"])

    if not duct_dict_rows or not fit_dict_rows or not cp_dict_rows:
        # Surface debug info
        output.print_md("**System %s parse issue**" % (i+1))
        output.print_md("- Duct rows: %s" % len(duct_dict_rows))
        output.print_md("- Fit rows: %s" % len(fit_dict_rows))
        output.print_md("- CP rows: %s" % len(cp_dict_rows))
        continue

    # Critical path sections list (strings)
    critical_sections = get_critical_path_list(cp_dict_rows)

    # Build section mapping and split into section/report rows
    duct_eid_to_sec, duct_section_rows, duct_report_rows = build_section_map(duct_dict_rows, element_id_key=u"Element ID")
    fit_eid_to_sec,  fit_section_rows,  fit_report_rows  = build_section_map(fit_dict_rows,  element_id_key=u"Element ID")

    # Insert Section into report rows (based on Element ID -> Section map)
    for rr in duct_report_rows:
        eid = clean_text(rr.get(u"Element ID", u""))
        rr[u"Section"] = clean_text(duct_eid_to_sec.get(eid, u""))

    for rr in fit_report_rows:
        eid = clean_text(rr.get(u"Element ID", u""))
        rr[u"Section"] = clean_text(fit_eid_to_sec.get(eid, u""))

    # Unit coercion (numbers become normalized strings)
    duct_report_rows = coerce_units_on_rows(duct_report_rows, is_duct=True)
    fit_report_rows  = coerce_units_on_rows(fit_report_rows,  is_duct=False)

    # Merge duct + fittings; Flow max per section; rename Comments->Critical Path
    merged = merge_duct_and_fittings(duct_report_rows, fit_report_rows, critical_sections)

    # Ensure System Name exists (present in your tables; keep if missing)
    # If missing, set blank (already handled by projection). You can derive from earlier tables if needed.

    merged = reindex_and_project_columns(merged)
    SYSTEM_ROWS_ALL.extend(merged)

# Write combined CSV (equivalent to df_DuctData in your reader script) :contentReference[oaicite:10]{index=10}
base_dir = os.path.dirname(html_path)
base_name = safe_filename(os.path.splitext(os.path.basename(html_path))[0])
out_path = os.path.join(base_dir, base_name + "__ductdata_all.csv")

write_csv(out_path, COLUMN_ORDER, SYSTEM_ROWS_ALL)

output.print_md("### Pressure Loss Report Export (no Pandas)")
output.print_md("* Source: `%s`" % html_path)
output.print_md("* Systems processed: **%s**" % sys_count)
output.print_md("* Output rows: **%s**" % len(SYSTEM_ROWS_ALL))
output.print_md("* Wrote: `%s`" % out_path)
