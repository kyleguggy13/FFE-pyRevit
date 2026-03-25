# -*- coding: utf-8 -*-
__title__     = "Lineage Viewer"
__version__   = 'Version = v0.1'
__doc__       = """Version = v0.1
Date    = 03.25.2026
______________________________________________________________
Description:
-> 
______________________________________________________________
How-to:
-> 
______________________________________________________________
Last update:
- [03.25.2026] - v0.1 BETA RELEASE
______________________________________________________________
Author: Kyle Guggenheim"""


#____________________________________________________________________ IMPORTS (SYSTEM)
from System import String
from collections import defaultdict
import time


#____________________________________________________________________ IMPORTS (AUTODESK)
import sys
import clr
clr.AddReference("System")
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
from Autodesk.Revit.DB import ElementCategoryFilter, ElementId, FamilyInstance


#____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import revit, DB, UI, script
from pyrevit.script import output
from pyrevit import forms

#____________________________________________________________________ VARIABLES
app         = __revit__.Application
uidoc       = __revit__.ActiveUIDocument
doc         = __revit__.ActiveUIDocument.Document   #type: Document
selection   = uidoc.Selection                       #type: Selection

log_status = ""
action = "Lineage Viewer"

output_window = output.get_output()
"""Output window for displaying results."""



#____________________________________________________________________ FUNCTIONS

"""
pyRevit | Lineage Viewer

-

"""

# -*- coding: utf-8 -*-
"""
__title__ = "Lineage Viewer"
__doc__   = "Collect linked items in the active Revit model and generate a relationship graphic."
"""

import os
import re
import io
import time
import math
import tempfile
import traceback
import subprocess

from collections import defaultdict

from pyrevit import revit, script, forms
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
    RevitLinkType,
    ImportInstance,
    ImageInstance,
    ImageType,
    PointCloudType,
    ExternalFileUtils,
    ModelPathUtils
)

output = script.get_output()
doc = revit.doc


# =============================================================================
# helpers
# =============================================================================

def html_escape(text):
    if text is None:
        return ""
    s = str(text)
    s = s.replace("&", "&amp;")
    s = s.replace("<", "&lt;")
    s = s.replace(">", "&gt;")
    s = s.replace('"', "&quot;")
    return s


def clean_filename(name):
    if not name:
        return "Untitled"
    return re.sub(r'[\\/*?:"<>|]+', "_", name)


def basename_or_value(path_value):
    if not path_value:
        return ""
    try:
        return os.path.basename(path_value)
    except:
        return str(path_value)


def shorten_text(text, max_len):
    if not text:
        return ""
    s = str(text)
    if len(s) <= max_len:
        return s
    return s[:max_len - 3] + "..."


def open_file(filepath):
    try:
        os.startfile(filepath)
        return True
    except:
        pass

    try:
        subprocess.Popen(['cmd', '/c', 'start', '', filepath], shell=True)
        return True
    except:
        pass

    return False


def try_get_external_path(owner_doc, element_id):
    """Try to resolve an external file path from an element/type id."""
    try:
        ext_ref = ExternalFileUtils.GetExternalFileReference(owner_doc, element_id)
        if ext_ref:
            mp = ext_ref.GetAbsolutePath()
            if mp:
                return ModelPathUtils.ConvertModelPathToUserVisiblePath(mp)
    except:
        pass
    return ""


def safe_get_link_status(link_type):
    try:
        status = RevitLinkType.GetLinkedFileStatus(link_type)
        if status:
            return str(status)
    except:
        pass
    return "Unknown"


def get_import_name(imp, owner_doc):
    try:
        etype = owner_doc.GetElement(imp.GetTypeId())
        if etype:
            return etype.Name
    except:
        pass
    try:
        return imp.Name
    except:
        return "CAD Item"


def get_image_path(img_type):
    try:
        return img_type.Path
    except:
        pass
    return ""


def get_pointcloud_path(pc_type):
    # API availability varies by version
    for attr in ["Path", "GetPath"]:
        try:
            value = getattr(pc_type, attr)
            if callable(value):
                p = value()
            else:
                p = value
            if p:
                return str(p)
        except:
            pass
    return ""


def doc_key(xdoc):
    try:
        if xdoc.PathName:
            return xdoc.PathName
    except:
        pass

    try:
        return "DOCID:{}".format(xdoc.GetHashCode())
    except:
        return "DOC:{}".format(id(xdoc))


# =============================================================================
# graph data
# =============================================================================

class GraphNode(object):
    def __init__(self, node_id, label, sublabel, kind, depth, parent_id=None, status=""):
        self.id = node_id
        self.label = label
        self.sublabel = sublabel
        self.kind = kind
        self.depth = depth
        self.parent_id = parent_id
        self.status = status


nodes = {}
edges = []
node_counter = [0]
max_depth = [0]
visited_doc_keys = set()


def next_id(prefix):
    node_counter[0] += 1
    return "{}_{}".format(prefix, node_counter[0])


def add_node(label, sublabel, kind, depth, parent_id=None, status=""):
    nid = next_id(kind)
    nodes[nid] = GraphNode(nid, label, sublabel, kind, depth, parent_id, status)
    if parent_id:
        edges.append((nid, parent_id))
    if depth > max_depth[0]:
        max_depth[0] = depth
    return nid


# =============================================================================
# collection
# =============================================================================

def collect_imports(owner_doc, parent_node_id, depth):
    try:
        imports = FilteredElementCollector(owner_doc).OfClass(ImportInstance).WhereElementIsNotElementType().ToElements()
    except:
        imports = []

    for imp in imports:
        if imp.IsLinked:
            kind = "cadlink"
        else:
            kind = "cadimport"
        
        try:
            # name = get_import_name(imp, owner_doc)
            name = imp.Parameter[DB.BuiltInParameter.IMPORT_SYMBOL_NAME].AsString()

            is_link = False
            try:
                is_link = imp.IsLinked
            except:
                pass

            path_value = try_get_external_path(owner_doc, imp.GetTypeId())
            sub = basename_or_value(path_value) if path_value else "No path available"

            # kind = "cadlink" if is_link else "cadimport"
            # status = "Linked" if is_link else "Imported"

            add_node(
                label=name,
                sublabel=sub,
                kind=kind,
                depth=depth,
                parent_id=parent_node_id
            )
        except:
            pass


def collect_images_and_pdfs(owner_doc, parent_node_id, depth):
    """Collect placed images/PDFs first, then unplaced image types as fallback."""

    placed_type_ids = set()

    # -------------------------------------------------------------------------
    # placed image/pdf instances
    # -------------------------------------------------------------------------
    try:
        image_instances = (
            FilteredElementCollector(owner_doc)
            .OfClass(ImageInstance)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        print("Found {} image instances.".format(len(image_instances)))     # <-- TESTING
    except:
        image_instances = []

    for img_inst in image_instances:
        print("Processing image instance with id {} (Name: {})).".format(img_inst.Id.ToString(), img_inst.Name))     # <-- TESTING
        try:
            # img_type = owner_doc.GetElement(img_inst.GetTypeId())
            img_inst_id = img_inst.Id.ToString()
            
            if not img_inst_id:
                print("no img_inst_id")     # <-- TESTING
                continue
            # if not img_type:
                # continue

            # placed_type_ids.add(img_type.Id.ToString())
            placed_type_ids.add(img_inst.Id.ToString())

            # try:
            #     source_str = str(img_type.Source)
            # except:
            #     source_str = "Image"

            # try:
            #     status_str = str(img_type.Status)
            # except:
            #     status_str = ""

            # try:
            #     page_number = img_type.PageNumber
            # except:
            #     page_number = 1

            # try:
            #     path_value = img_type.Path
            # except:
            #     path_value = ""

            # try:
            #     owner_view = owner_doc.GetElement(img_inst.OwnerViewId)
            #     owner_view_name = owner_view.Name if owner_view else "Unknown View"
            # except:
            #     owner_view_name = "Unknown View"


            source_str = "Image"
            status_str = ""
            page_number = 1
            path_value = ""
            owner_view_name = "Unknown View"

            ext = os.path.splitext(path_value)[1].lower() if path_value else ""
            is_pdf = ext == ".pdf" or page_number > 1

            if is_pdf:
                # label = img_type.Name or "PDF"
                label = img_inst.Name or "PDF"
                sub = "{} | Page {} | View: {}".format(
                    basename_or_value(path_value) if path_value else "PDF",
                    page_number,
                    owner_view_name
                )
                kind = "pdf"
            else:
                # label = img_type.Name or "Image"
                label = img_inst.Name or "Image"
                sub = "{} | View: {}".format(
                    basename_or_value(path_value) if path_value else "Image",
                    owner_view_name
                )
                kind = "image"

            add_node(
                label=label,
                sublabel=sub,
                kind=kind,
                depth=depth,
                parent_id=parent_node_id
            )
            print("add_node: {}, {}, {}, {}, {}, {}".format(
                label, sub, kind, depth, parent_node_id, status_str if status_str else source_str
            ))     # <-- TESTING
        except:
            pass

    # -------------------------------------------------------------------------
    # fallback image/pdf types not currently placed
    # -------------------------------------------------------------------------
    # try:
    #     image_types = FilteredElementCollector(owner_doc).OfClass(ImageType).ToElements()
    # except:
    #     image_types = []

    # for img_type in image_types:
    #     try:
    #         # Skip types that already have placed instances collected above
    #         try:
    #             if img_type.Id.ToString() in placed_type_ids:
    #                 continue
    #         except:
    #             pass

    #         try:
    #             path_value = img_type.Path
    #         except:
    #             path_value = ""

    #         try:
    #             source_str = str(img_type.Source)
    #         except:
    #             source_str = "Image"

    #         try:
    #             status_str = str(img_type.Status)
    #         except:
    #             status_str = ""

    #         try:
    #             page_number = img_type.PageNumber
    #         except:
    #             page_number = 1

    #         ext = os.path.splitext(path_value)[1].lower() if path_value else ""
    #         is_pdf = ext == ".pdf" or page_number > 1

    #         if is_pdf:
    #             label = img_type.Name or "PDF"
    #             sub = "{} | Page {} | Unplaced".format(
    #                 basename_or_value(path_value) if path_value else "PDF",
    #                 page_number
    #             )
    #             kind = "pdf"
    #         else:
    #             label = img_type.Name or "Image"
    #             sub = "{} | Unplaced".format(
    #                 basename_or_value(path_value) if path_value else "Image"
    #             )
    #             kind = "image"

    #         add_node(
    #             label=label,
    #             sublabel=sub,
    #             kind=kind,
    #             depth=depth,
    #             parent_id=parent_node_id,
    #             status=status_str if status_str else source_str
    #         )
    #     except:
    #         pass


def collect_pointclouds(owner_doc, parent_node_id, depth):
    try:
        pointcloud_types = FilteredElementCollector(owner_doc).OfClass(PointCloudType).ToElements()
    except:
        pointcloud_types = []

    for pc_type in pointcloud_types:
        try:
            path_value = get_pointcloud_path(pc_type)
            sub = basename_or_value(path_value) if path_value else "Point cloud"

            add_node(
                label=pc_type.Name,
                sublabel=sub,
                kind="pointcloud",
                depth=depth,
                parent_id=parent_node_id
            )
        except:
            pass


def collect_revit_links(owner_doc, parent_node_id, depth):
    try:
        link_instances = FilteredElementCollector(owner_doc).OfClass(RevitLinkInstance).ToElements()
    except:
        link_instances = []

    for inst in link_instances:
        try:
            link_type = owner_doc.GetElement(inst.GetTypeId())
            link_doc = inst.GetLinkDocument()

            instance_name = ""
            try:
                instance_name = inst.Name
            except:
                pass

            if not instance_name and link_type:
                instance_name = link_type.Name
            if not instance_name:
                instance_name = "Revit Link"

            path_value = ""
            if link_type:
                path_value = try_get_external_path(owner_doc, link_type.Id)

            sub = basename_or_value(path_value) if path_value else "No path available"
            status = safe_get_link_status(link_type) if link_type else "Unknown"

            node_id = add_node(
                label=instance_name,
                sublabel=sub,
                kind="revitlink",
                depth=depth,
                parent_id=parent_node_id
            )

            # Recurse into loaded linked docs only
            if link_doc:
                k = doc_key(link_doc)
                if k not in visited_doc_keys:
                    visited_doc_keys.add(k)
                    collect_doc_contents(link_doc, node_id, depth + 1)

        except:
            pass


def collect_doc_contents(owner_doc, parent_node_id, depth):
    collect_revit_links(owner_doc, parent_node_id, depth)
    collect_imports(owner_doc, parent_node_id, depth)
    collect_images_and_pdfs(owner_doc, parent_node_id, depth)
    collect_pointclouds(owner_doc, parent_node_id, depth)


# =============================================================================
# styling / layout
# =============================================================================

def node_style(kind):
    styles = {
        "host": {
            "icon": "R",
            "border": "#3B8D82",
            "icon_fill": "#EAF7F4",
            "icon_stroke": "#3B8D82",
            "tag": "Active Model"
        },
        "revitlink": {
            "icon": "R",
            "border": "#7E9CB2",
            "icon_fill": "#F3F7FA",
            "icon_stroke": "#7E9CB2",
            "tag": "Revit Link"
        },
        "cadlink": {
            "icon": "C",
            "border": "#9B8DB9",
            "icon_fill": "#F5F2FB",
            "icon_stroke": "#9B8DB9",
            "tag": "CAD Link"
        },
        "cadimport": {
            "icon": "C",
            "border": "#B69561",
            "icon_fill": "#FBF6EC",
            "icon_stroke": "#B69561",
            "tag": "CAD Import"
        },
        "image": {
            "icon": "I",
            "border": "#C68957",
            "icon_fill": "#FCF2EA",
            "icon_stroke": "#C68957",
            "tag": "Image"
        },
        "pointcloud": {
            "icon": "P",
            "border": "#6E9FB2",
            "icon_fill": "#EEF7FA",
            "icon_stroke": "#6E9FB2",
            "tag": "Point Cloud"
        },
        "pdf": {
            "icon": "P",
            "border": "#D08A57",
            "icon_fill": "#FCF3EC",
            "icon_stroke": "#D08A57",
            "tag": "PDF"
        },
    }
    return styles.get(kind, styles["revitlink"])


def compute_layout():
    columns = defaultdict(list)

    for n in nodes.values():
        columns[n.depth].append(n)

    for depth in columns:
        columns[depth] = sorted(columns[depth], key=lambda x: (x.kind, x.label.lower()))

    node_w = 370
    node_h = 96
    x_gap = 130
    y_gap = 24
    margin_x = 50
    margin_y = 40

    positions = {}
    max_column_height = 0

    for depth in range(0, max_depth[0] + 1):
        col_nodes = columns.get(depth, [])
        col_height = len(col_nodes) * node_h + max(0, len(col_nodes) - 1) * y_gap
        if col_height > max_column_height:
            max_column_height = col_height

    total_h = max(max_column_height + margin_y * 2, 350)
    total_w = margin_x * 2 + (max_depth[0] + 1) * node_w + max_depth[0] * x_gap

    for depth in range(0, max_depth[0] + 1):
        col_nodes = columns.get(depth, [])
        x = margin_x + depth * (node_w + x_gap)

        col_height = len(col_nodes) * node_h + max(0, len(col_nodes) - 1) * y_gap
        start_y = margin_y + max(0, (max_column_height - col_height) / 2.0)

        for i, n in enumerate(col_nodes):
            y = start_y + i * (node_h + y_gap)
            positions[n.id] = (x, y)

    return positions, total_w, total_h, node_w, node_h


def build_svg():
    positions, total_w, total_h, node_w, node_h = compute_layout()

    svg = []
    svg.append('<svg xmlns="http://www.w3.org/2000/svg" width="{0}" height="{1}" viewBox="0 0 {0} {1}">'.format(int(total_w), int(total_h)))
    svg.append("""
    <defs>
      <filter id="shadow" x="-20%" y="-20%" width="160%" height="160%">
        <feDropShadow dx="0" dy="2" stdDeviation="2.4" flood-color="#000000" flood-opacity="0.12"/>
      </filter>
    </defs>
    """)

    # connections
    for child_id, parent_id in edges:
        if child_id not in positions or parent_id not in positions:
            continue

        child_node = nodes[child_id]
        parent_node = nodes[parent_id]

        cx, cy = positions[child_id]
        px, py = positions[parent_id]

        child_left_x = cx
        child_right_x = cx + node_w
        child_mid_y = cy + (node_h / 2.0)

        parent_left_x = px
        parent_right_x = px + node_w
        parent_mid_y = py + (node_h / 2.0)

        if parent_node.depth < child_node.depth:
            # parent -> child
            x1 = parent_right_x
            y1 = parent_mid_y
            x2 = child_left_x
            y2 = child_mid_y
        elif parent_node.depth > child_node.depth:
            # reverse fallback
            x1 = parent_left_x
            y1 = parent_mid_y
            x2 = child_right_x
            y2 = child_mid_y
        else:
            # same column fallback
            if px <= cx:
                x1 = parent_right_x
                y1 = parent_mid_y
                x2 = child_left_x
                y2 = child_mid_y
            else:
                x1 = parent_left_x
                y1 = parent_mid_y
                x2 = child_right_x
                y2 = child_mid_y

        dx = max(50, abs(x2 - x1) / 2.0)

        if x1 <= x2:
            c1x = x1 + dx
            c2x = x2 - dx
        else:
            c1x = x1 - dx
            c2x = x2 + dx

        path_d = "M {0} {1} C {2} {1}, {3} {4}, {5} {4}".format(
            x1, y1, c1x, c2x, y2, x2
        )

        svg.append('<path d="{0}" fill="none" stroke="#C4CAD1" stroke-width="2.2"/>'.format(path_d))

    # nodes
    for n in nodes.values():
        if n.id not in positions:
            continue

        x, y = positions[n.id]
        st = node_style(n.kind)

        # node container with shadow
        svg.append('<g filter="url(#shadow)">')

        # node background
        svg.append('<rect x="{0}" y="{1}" rx="10" ry="10" width="{2}" height="{3}" fill="#FFFFFF" stroke="{4}" stroke-width="1.8"/>'.format(
            x, y, node_w, node_h, st["border"]
        ))
        svg.append('</g>')

        # icon background
        svg.append('<rect x="{0}" y="{1}" rx="6" ry="6" width="26" height="26" fill="{2}" stroke="{3}" stroke-width="1.2"/>'.format(
            x + 14, y + 14, st["icon_fill"], st["icon_stroke"]
        ))

        # icon label
        svg.append('<text x="{0}" y="{1}" font-family="Segoe UI, Arial" font-size="14" font-weight="700" fill="{2}">{3}</text>'.format(
            x + 21, y + 32, st["icon_stroke"], html_escape(st["icon"])
        ))

        # main label
        svg.append('<text x="{0}" y="{1}" font-family="Segoe UI, Arial" font-size="14" font-weight="700" fill="#2F3438">{2}</text>'.format(
            x + 52, y + 28, html_escape(shorten_text(n.label, 42))
        ))

        # sub label
        svg.append('<text x="{0}" y="{1}" font-family="Segoe UI, Arial" font-size="12" fill="#676E75">{2}</text>'.format(
            x + 52, y + 50, html_escape(st["tag"])
        ))

        # additional sub label
        svg.append('<text x="{0}" y="{1}" font-family="Segoe UI, Arial" font-size="12" fill="#80878E">{2}</text>'.format(
            x + 52, y + 71, html_escape(shorten_text(n.sublabel, 48))
        ))

        # status label
        if n.status:
            svg.append('<text x="{0}" y="{1}" font-family="Segoe UI, Arial" font-size="12" fill="#7B8288">{2}</text>'.format(
                x + 52, y + 88, html_escape(n.status)
            ))

    svg.append("</svg>")
    return "".join(svg), int(total_w), int(total_h)


def build_html():
    host_name = doc.Title
    host_path = doc.PathName if doc.PathName else "Unsaved model"

    svg_markup, width_px, height_px = build_svg()

    summary_counts = defaultdict(int)
    for n in nodes.values():
        summary_counts[n.kind] += 1

    html = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Revit Linked Items Relationship Graph</title>
<style>
    body {{
        margin: 0;
        background: #F3F4F6;
        font-family: Segoe UI, Arial, sans-serif;
        color: #333333;
    }}
    .header {{
        background: #FFFFFF;
        border-bottom: 1px solid #D7DCE1;
        padding: 18px 24px 14px 24px;
    }}
    .title {{
        font-size: 24px;
        font-weight: 700;
        margin-bottom: 6px;
    }}
    .meta {{
        font-size: 13px;
        color: #6A7178;
        line-height: 1.55;
    }}
    .wrap {{
        padding: 16px 20px 22px 20px;
    }}
    .summary {{
        margin-bottom: 14px;
        padding: 10px 12px;
        background: #FFFFFF;
        border: 1px solid #D7DCE1;
        border-radius: 8px;
        font-size: 13px;
        color: #5F666D;
    }}
    .canvas {{
        background: #F3F4F6;
        overflow: auto;
        border-radius: 10px;
    }}
</style>
</head>
<body>
    <div class="header">
        <div class="title">Revit Linked Items Relationship Graph</div>
        <div class="meta"><strong>Model:</strong> {host_name}</div>
        <div class="meta"><strong>Path:</strong> {host_path}</div>
        <div class="meta"><strong>Generated:</strong> {generated}</div>
        <div class="meta"><strong>Note:</strong> Nested contents are only shown for loaded Revit links.</div>
    </div>

    <div class="wrap">
        <div class="summary">
            <strong>Counts</strong><br>
            Revit Links: {revit_count} &nbsp; | &nbsp;
            CAD Links: {cadlink_count} &nbsp; | &nbsp;
            CAD Imports: {cadimport_count} &nbsp; | &nbsp;
            Images: {image_count} &nbsp; | &nbsp;
            PDFs: {pdf_count} &nbsp; | &nbsp;
            Point Clouds: {pc_count}
        </div>

        <div class="canvas">
            {svg}
        </div>
    </div>
</body>
</html>
""".format(
        host_name=html_escape(host_name),
        host_path=html_escape(host_path),
        generated=html_escape(time.strftime("%Y-%m-%d %I:%M:%S %p")),
        revit_count=summary_counts.get("revitlink", 0),
        cadlink_count=summary_counts.get("cadlink", 0),
        cadimport_count=summary_counts.get("cadimport", 0),
        image_count=summary_counts.get("image", 0),
        pdf_count=summary_counts.get("pdf", 0),
        pc_count=summary_counts.get("pointcloud", 0),
        svg=svg_markup
    )

    return html


# =============================================================================
# main
# =============================================================================

try:
    host_node_id = add_node(
        label=doc.Title,
        sublabel=basename_or_value(doc.PathName) if doc.PathName else "Unsaved model",
        kind="host",
        depth=0,
        parent_id=None,
    )
    # host_node_id = add_node(
    #     label="label",
    #     sublabel="sublabel",
    #     kind="kind",
    #     depth=0,
    #     parent_id="parent_id"
    # )

    visited_doc_keys.add(doc_key(doc))
    collect_doc_contents(doc, host_node_id, 1)

    if len(nodes) == 1:
        forms.alert("No linked items were found in the active model.", exitscript=True)

    html = build_html()

    out_dir = tempfile.gettempdir()
    out_name = "Revit_Linked_Items_Graph_{0}.html".format(clean_filename(doc.Title))
    out_path = os.path.join(out_dir, out_name)

    with io.open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    output.print_md("### Linked Items Graph Created")
    output.print_md("**File:** `{}`".format(out_path))

    if open_file(out_path):
        output.print_md("Opened report in default browser.")
    else:
        output.print_md("Could not auto-open report. Open it manually from the path above.")

except Exception as ex:
    output.print_md("### Script Failed")
    output.print_md("**Error:** `{}`".format(str(ex)))
    output.print_md("```")
    output.print_md(traceback.format_exc())
    output.print_md("```")

