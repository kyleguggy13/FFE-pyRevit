# -*- coding: utf-8 -*-
"""CPython helper for PDF2Revit.

This file is intentionally independent from Revit and pyRevit. The pyRevit
button calls it through a local CPython venv with PyMuPDF installed.
"""

# from __future__ import absolute_import

import json
import math
import os
import sys
import traceback

try:
    import fitz
except Exception:
    fitz = None


ANGLE_TOLERANCE_RAD = math.radians(8.0)
MIN_SEGMENT_PDF = 1.0
MIN_WALL_LENGTH_FT = 2.0
MAX_DETECTED_OPENINGS = 120


def read_json(path):
    with open(path, "r", encoding="utf-8") as file_obj:
        return json.load(file_obj)


def write_json(path, payload):
    parent = os.path.dirname(path)
    if parent and not os.path.exists(parent):
        os.makedirs(parent)
    with open(path, "w", encoding="utf-8") as file_obj:
        json.dump(payload, file_obj, ensure_ascii=True, indent=2)


def require_fitz():
    if fitz is None:
        raise RuntimeError("PyMuPDF is not available. Install the PyMuPDF package in the PDF2Revit venv.")


def point_xy(point):
    try:
        return float(point.x), float(point.y)
    except AttributeError:
        return float(point[0]), float(point[1])


def distance(a, b):
    return math.hypot(float(b[0]) - float(a[0]), float(b[1]) - float(a[1]))


def dot(a, b):
    return float(a[0]) * float(b[0]) + float(a[1]) * float(b[1])


def sub(a, b):
    return float(a[0]) - float(b[0]), float(a[1]) - float(b[1])


def add(a, b):
    return float(a[0]) + float(b[0]), float(a[1]) + float(b[1])


def mul(a, scale):
    return float(a[0]) * scale, float(a[1]) * scale


def normalize_angle(angle):
    angle = angle % math.pi
    if angle < 0:
        angle += math.pi
    if abs(angle - math.pi) < 0.000001:
        angle = 0.0
    return angle


def angle_diff(a, b):
    diff = abs(normalize_angle(a) - normalize_angle(b))
    return min(diff, math.pi - diff)


def segment_angle(segment):
    a = segment["a"]
    b = segment["b"]
    return normalize_angle(math.atan2(b[1] - a[1], b[0] - a[0]))


def line_basis(angle):
    unit = (math.cos(angle), math.sin(angle))
    normal = (-math.sin(angle), math.cos(angle))
    return unit, normal


def segment_axis_data(segment):
    angle = segment_angle(segment)
    unit, normal = line_basis(angle)
    t0 = dot(segment["a"], unit)
    t1 = dot(segment["b"], unit)
    offset = dot(segment["a"], normal)
    return {
        "angle": angle,
        "unit": unit,
        "normal": normal,
        "t0": min(t0, t1),
        "t1": max(t0, t1),
        "offset": offset,
    }


def point_from_axis(unit, normal, t_value, offset):
    return add(mul(unit, t_value), mul(normal, offset))


def make_segment(a, b, source):
    length = distance(a, b)
    if length < MIN_SEGMENT_PDF:
        return None
    return {
        "a": (float(a[0]), float(a[1])),
        "b": (float(b[0]), float(b[1])),
        "length": length,
        "source": source,
    }


def add_segment(segments, a, b, source):
    segment = make_segment(a, b, source)
    if segment:
        segments.append(segment)


def rect_points(rect):
    x0 = float(rect.x0)
    y0 = float(rect.y0)
    x1 = float(rect.x1)
    y1 = float(rect.y1)
    return [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]


def cubic_point(p0, p1, p2, p3, t_value):
    inv = 1.0 - t_value
    x = (
        inv * inv * inv * p0[0]
        + 3.0 * inv * inv * t_value * p1[0]
        + 3.0 * inv * t_value * t_value * p2[0]
        + t_value * t_value * t_value * p3[0]
    )
    y = (
        inv * inv * inv * p0[1]
        + 3.0 * inv * inv * t_value * p1[1]
        + 3.0 * inv * t_value * t_value * p2[1]
        + t_value * t_value * t_value * p3[1]
    )
    return x, y


def flatten_cubic(points, steps=12):
    flattened = []
    for index in range(steps + 1):
        flattened.append(cubic_point(points[0], points[1], points[2], points[3], float(index) / float(steps)))
    return flattened


def collect_segments(page):
    segments = []
    curves = []
    drawings = page.get_drawings()

    for drawing in drawings:
        current = None
        for item in drawing.get("items", []):
            command = item[0]

            if command == "l" and len(item) >= 3:
                a = point_xy(item[1])
                b = point_xy(item[2])
                add_segment(segments, a, b, "line")
                current = b

            elif command == "re" and len(item) >= 2:
                points = rect_points(item[1])
                for index in range(len(points)):
                    add_segment(segments, points[index], points[(index + 1) % len(points)], "rect")
                current = points[-1]

            elif command == "c" and len(item) >= 5:
                p0 = point_xy(item[1])
                p1 = point_xy(item[2])
                p2 = point_xy(item[3])
                p3 = point_xy(item[4])
                curve_points = flatten_cubic([p0, p1, p2, p3])
                for index in range(len(curve_points) - 1):
                    add_segment(segments, curve_points[index], curve_points[index + 1], "curve")
                curves.append({
                    "points": curve_points,
                    "source": "curve",
                })
                current = p3

            elif command == "qu" and len(item) >= 2:
                try:
                    quad = item[1]
                    points = [point_xy(quad.ul), point_xy(quad.ur), point_xy(quad.lr), point_xy(quad.ll)]
                except Exception:
                    points = []
                for index in range(len(points)):
                    add_segment(segments, points[index], points[(index + 1) % len(points)], "quad")
                if points:
                    current = points[-1]

            elif command == "m" and len(item) >= 2:
                current = point_xy(item[1])

            elif command == "h" and current:
                current = current

    return segments, curves


def merge_collinear_segments(segments, offset_tolerance, gap_tolerance):
    groups = {}
    merged = []

    for segment in segments:
        data = segment_axis_data(segment)
        angle_key = int(round(data["angle"] / ANGLE_TOLERANCE_RAD))
        offset_key = int(round(data["offset"] / max(0.01, offset_tolerance)))
        key = (angle_key, offset_key)
        item = {
            "angle": data["angle"],
            "unit": data["unit"],
            "normal": data["normal"],
            "offset": data["offset"],
            "t0": data["t0"],
            "t1": data["t1"],
            "source": segment.get("source") or "line",
        }
        groups.setdefault(key, []).append(item)

    for items in groups.values():
        items.sort(key=lambda item: item["t0"])
        active = None

        for item in items:
            if active is None:
                active = dict(item)
                continue

            if item["t0"] <= active["t1"] + gap_tolerance:
                active["t1"] = max(active["t1"], item["t1"])
                active["offset"] = (active["offset"] + item["offset"]) / 2.0
            else:
                a = point_from_axis(active["unit"], active["normal"], active["t0"], active["offset"])
                b = point_from_axis(active["unit"], active["normal"], active["t1"], active["offset"])
                segment = make_segment(a, b, active.get("source") or "merged")
                if segment:
                    merged.append(segment)
                active = dict(item)

        if active:
            a = point_from_axis(active["unit"], active["normal"], active["t0"], active["offset"])
            b = point_from_axis(active["unit"], active["normal"], active["t1"], active["offset"])
            segment = make_segment(a, b, active.get("source") or "merged")
            if segment:
                merged.append(segment)

    return merged


def calibration_to_scale(calibration):
    point_a = calibration.get("pointA") or {}
    point_b = calibration.get("pointB") or {}
    pdf_a = (float(point_a.get("x")), float(point_a.get("y")))
    pdf_b = (float(point_b.get("x")), float(point_b.get("y")))
    pdf_distance = distance(pdf_a, pdf_b)
    known_distance = float(calibration.get("distance") or 0.0)
    unit = (calibration.get("unit") or "ft").lower()

    if pdf_distance <= 0:
        raise ValueError("Calibration points cannot be the same point.")
    if known_distance <= 0:
        raise ValueError("Known calibration distance must be greater than zero.")

    if unit == "in":
        feet = known_distance / 12.0
    elif unit == "m":
        feet = known_distance * 3.280839895
    elif unit == "mm":
        feet = known_distance * 0.003280839895
    else:
        feet = known_distance

    return {
        "origin": pdf_a,
        "feet_per_pdf_point": feet / pdf_distance,
        "known_distance_feet": feet,
        "pdf_distance": pdf_distance,
    }


def make_transform(calibration):
    scale_info = calibration_to_scale(calibration)
    origin = scale_info["origin"]
    scale = scale_info["feet_per_pdf_point"]

    def pdf_to_revit(point):
        return [
            (float(point[0]) - origin[0]) * scale,
            (origin[1] - float(point[1])) * scale,
        ]

    return scale_info, pdf_to_revit


def build_wall_candidates(segments, pdf_to_revit, wall_width_feet, scale_info):
    scale = scale_info["feet_per_pdf_point"]
    wall_width_pdf = max(0.50, float(wall_width_feet or 0.5) / scale)
    line_segments = [segment for segment in segments if segment.get("source") != "curve"]
    line_segments = [
        segment for segment in line_segments
        if segment["length"] * scale >= max(MIN_WALL_LENGTH_FT, float(wall_width_feet or 0.5) * 3.0)
    ]

    if len(line_segments) > 1400:
        line_segments = sorted(line_segments, key=lambda item: item["length"], reverse=True)[:1400]

    merged_lines = merge_collinear_segments(
        line_segments,
        max(1.5, wall_width_pdf * 0.20),
        max(2.0, wall_width_pdf * 1.50)
    )
    center_segments = []
    min_sep = max(0.5, wall_width_pdf * 0.25)
    max_sep = max(1.5, wall_width_pdf * 1.85)
    min_overlap_pdf = MIN_WALL_LENGTH_FT / scale

    axis_data = [(segment, segment_axis_data(segment)) for segment in merged_lines]
    for index, item_a in enumerate(axis_data):
        segment_a, data_a = item_a
        for segment_b, data_b in axis_data[index + 1:]:
            if angle_diff(data_a["angle"], data_b["angle"]) > ANGLE_TOLERANCE_RAD:
                continue

            separation = abs(data_a["offset"] - data_b["offset"])
            if separation < min_sep or separation > max_sep:
                continue

            overlap_start = max(data_a["t0"], data_b["t0"])
            overlap_end = min(data_a["t1"], data_b["t1"])
            if overlap_end - overlap_start < min_overlap_pdf:
                continue

            center_offset = (data_a["offset"] + data_b["offset"]) / 2.0
            start = point_from_axis(data_a["unit"], data_a["normal"], overlap_start, center_offset)
            end = point_from_axis(data_a["unit"], data_a["normal"], overlap_end, center_offset)
            center_segments.append({
                "a": start,
                "b": end,
                "length": distance(start, end),
                "source": "wall-center",
            })

    merged_centers = merge_collinear_segments(
        center_segments,
        max(1.0, wall_width_pdf * 0.65),
        max(2.0, wall_width_pdf * 2.50)
    )

    walls = []
    seen = set()
    for segment in sorted(merged_centers, key=lambda item: item["length"], reverse=True):
        a = segment["a"]
        b = segment["b"]
        revit_a = pdf_to_revit(a)
        revit_b = pdf_to_revit(b)
        length_feet = distance(revit_a, revit_b)
        if length_feet < MIN_WALL_LENGTH_FT:
            continue

        key_points = sorted([
            (round(revit_a[0], 2), round(revit_a[1], 2)),
            (round(revit_b[0], 2), round(revit_b[1], 2)),
        ])
        key = tuple(key_points)
        if key in seen:
            continue
        seen.add(key)

        walls.append({
            "id": "wall_{0:03d}".format(len(walls) + 1),
            "points": [revit_a, revit_b],
            "pdf_points": [[a[0], a[1]], [b[0], b[1]]],
            "length_feet": length_feet,
        })

    return walls


def polygon_area(points):
    if len(points) < 3:
        return 0.0
    total = 0.0
    for index, point in enumerate(points):
        next_point = points[(index + 1) % len(points)]
        total += point[0] * next_point[1] - next_point[0] * point[1]
    return abs(total) / 2.0


def convex_hull(point_items):
    unique = {}
    for item in point_items:
        key = (round(item["xy"][0], 4), round(item["xy"][1], 4))
        unique[key] = item

    points = sorted(unique.values(), key=lambda item: (item["xy"][0], item["xy"][1]))
    if len(points) <= 1:
        return points

    def cross(origin, a, b):
        return (
            (a["xy"][0] - origin["xy"][0]) * (b["xy"][1] - origin["xy"][1])
            - (a["xy"][1] - origin["xy"][1]) * (b["xy"][0] - origin["xy"][0])
        )

    lower = []
    for point in points:
        while len(lower) >= 2 and cross(lower[-2], lower[-1], point) <= 0:
            lower.pop()
        lower.append(point)

    upper = []
    for point in reversed(points):
        while len(upper) >= 2 and cross(upper[-2], upper[-1], point) <= 0:
            upper.pop()
        upper.append(point)

    return lower[:-1] + upper[:-1]


def build_floors(walls):
    point_items = []
    for wall in walls:
        for index, point in enumerate(wall.get("points") or []):
            pdf_point = wall.get("pdf_points")[index]
            point_items.append({
                "xy": point,
                "pdf": pdf_point,
            })

    hull = convex_hull(point_items)
    if len(hull) < 3:
        return [], ["No clean floor loop was detected from the wall geometry."]

    loop = [item["xy"] for item in hull]
    area = polygon_area(loop)
    if area < 10.0:
        return [], ["No clean floor loop was detected from the wall geometry."]

    return [{
        "id": "floor_001",
        "loop": loop,
        "pdf_loop": [item["pdf"] for item in hull],
        "area_square_feet": area,
    }], ["Floor loop uses the exterior convex hull of detected walls."]


def point_segment_projection(point, a, b):
    ap = sub(point, a)
    ab = sub(b, a)
    denom = dot(ab, ab)
    if denom <= 0:
        return 0.0, a, distance(point, a)
    t_value = max(0.0, min(1.0, dot(ap, ab) / denom))
    closest = add(a, mul(ab, t_value))
    return t_value, closest, distance(point, closest)


def nearest_wall(point, walls):
    best = None
    for wall in walls:
        points = wall.get("points") or []
        if len(points) != 2:
            continue
        t_value, closest, point_distance = point_segment_projection(point, points[0], points[1])
        if best is None or point_distance < best["distance"]:
            best = {
                "wall": wall,
                "t": t_value,
                "closest": closest,
                "distance": point_distance,
            }
    return best


def curve_stats(curve, pdf_to_revit):
    pdf_points = curve.get("points") or []
    points = [pdf_to_revit(point) for point in pdf_points]
    if len(points) < 2:
        return None

    xs = [point[0] for point in points]
    ys = [point[1] for point in points]
    width = max(xs) - min(xs)
    height = max(ys) - min(ys)
    chord = distance(points[0], points[-1])
    arc_length = sum(distance(points[index], points[index + 1]) for index in range(len(points) - 1))
    center = [(min(xs) + max(xs)) / 2.0, (min(ys) + max(ys)) / 2.0]
    pdf_xs = [point[0] for point in pdf_points]
    pdf_ys = [point[1] for point in pdf_points]
    pdf_center = [(min(pdf_xs) + max(pdf_xs)) / 2.0, (min(pdf_ys) + max(pdf_ys)) / 2.0]

    return {
        "width": width,
        "height": height,
        "chord": chord,
        "arc_length": arc_length,
        "center": center,
        "pdf_center": pdf_center,
    }


def is_duplicate_point(items, point, threshold):
    for item in items:
        if distance(item.get("point"), point) < threshold:
            return True
    return False


def build_doors(curves, walls, pdf_to_revit, wall_width_feet):
    doors = []
    max_host_distance = max(1.5, float(wall_width_feet or 0.5) * 4.0)

    for curve in curves:
        stats = curve_stats(curve, pdf_to_revit)
        if not stats:
            continue

        span = max(stats["width"], stats["height"])
        depth = min(stats["width"], stats["height"])
        if span < 1.5 or span > 5.5:
            continue
        if depth < 0.45:
            continue
        if stats["arc_length"] <= stats["chord"] * 1.12:
            continue

        host = nearest_wall(stats["center"], walls)
        if not host or host["distance"] > max_host_distance:
            continue
        if host["t"] <= 0.02 or host["t"] >= 0.98:
            continue
        if is_duplicate_point(doors, host["closest"], 2.0):
            continue

        doors.append({
            "id": "door_{0:03d}".format(len(doors) + 1),
            "point": host["closest"],
            "pdf_point": stats["pdf_center"],
            "host_wall_id": host["wall"].get("id"),
            "width_feet": span,
        })

        if len(doors) >= MAX_DETECTED_OPENINGS:
            break

    return doors


def segment_angle_feet(point_a, point_b):
    return normalize_angle(math.atan2(point_b[1] - point_a[1], point_b[0] - point_a[0]))


def build_windows(segments, walls, pdf_to_revit, wall_width_feet):
    windows = []
    max_host_distance = max(0.75, float(wall_width_feet or 0.5) * 2.5)

    for segment in segments:
        if segment.get("source") == "curve":
            continue

        a = pdf_to_revit(segment["a"])
        b = pdf_to_revit(segment["b"])
        length_feet = distance(a, b)
        if length_feet < 1.0 or length_feet > 8.0:
            continue

        center = [(a[0] + b[0]) / 2.0, (a[1] + b[1]) / 2.0]
        candidate_angle = segment_angle_feet(a, b)
        host = nearest_wall(center, walls)
        if not host or host["distance"] > max_host_distance:
            continue
        if host["t"] <= 0.03 or host["t"] >= 0.97:
            continue

        wall_points = host["wall"].get("points") or []
        wall_angle = segment_angle_feet(wall_points[0], wall_points[1])
        wall_length = host["wall"].get("length_feet") or distance(wall_points[0], wall_points[1])

        if angle_diff(candidate_angle, wall_angle) > ANGLE_TOLERANCE_RAD:
            continue
        if length_feet > wall_length * 0.40:
            continue

        duplicate = False
        for item in windows:
            if item.get("host_wall_id") == host["wall"].get("id") and distance(item.get("point"), host["closest"]) < 2.0:
                duplicate = True
                break
        if duplicate:
            continue

        windows.append({
            "id": "window_{0:03d}".format(len(windows) + 1),
            "point": host["closest"],
            "pdf_point": [(segment["a"][0] + segment["b"][0]) / 2.0, (segment["a"][1] + segment["b"][1]) / 2.0],
            "host_wall_id": host["wall"].get("id"),
            "width_feet": length_feet,
        })

        if len(windows) >= MAX_DETECTED_OPENINGS:
            break

    return windows


def open_document(pdf_path):
    require_fitz()
    if not os.path.exists(pdf_path):
        raise IOError("PDF was not found: {0}".format(pdf_path))
    return fitz.open(pdf_path)


def get_page(document, page_index):
    page_index = int(page_index or 0)
    if page_index < 0 or page_index >= document.page_count:
        raise IndexError("PDF page index is out of range.")
    return document.load_page(page_index)


def page_payload(page):
    rect = page.rect
    return {
        "width": float(rect.width),
        "height": float(rect.height),
        "rotation": int(page.rotation or 0),
    }


def operation_info(request):
    document = open_document(request.get("pdf_path"))
    try:
        return {
            "ok": True,
            "operation": "info",
            "page_count": int(document.page_count),
        }
    finally:
        document.close()


def operation_preview(request):
    document = open_document(request.get("pdf_path"))
    try:
        page = get_page(document, request.get("page_index"))
        output_dir = request.get("output_dir") or os.getcwd()
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        zoom = 2.0
        matrix = fitz.Matrix(zoom, zoom)
        pixmap = page.get_pixmap(matrix=matrix, alpha=False)
        image_path = os.path.join(output_dir, "preview-page-{0}.png".format(int(request.get("page_index") or 0) + 1))
        pixmap.save(image_path)

        return {
            "ok": True,
            "operation": "preview",
            "preview_image_path": image_path,
            "page": page_payload(page),
        }
    finally:
        document.close()


def operation_analyze(request):
    document = open_document(request.get("pdf_path"))
    try:
        page = get_page(document, request.get("page_index"))
        segments, curves = collect_segments(page)
        warnings = []

        if not segments:
            return {
                "ok": True,
                "operation": "analyze",
                "page": page_payload(page),
                "elements": {
                    "walls": [],
                    "floors": [],
                    "doors": [],
                    "windows": [],
                },
                "warnings": ["Unsupported or raster-only PDF: no vector geometry was found."],
                "summary": {
                    "rawSegmentCount": 0,
                    "curveCount": 0,
                },
            }

        settings = request.get("settings") or {}
        wall_width_feet = float(settings.get("wallWidthFeet") or 0.5)
        scale_info, pdf_to_revit = make_transform(request.get("calibration") or {})

        walls = build_wall_candidates(segments, pdf_to_revit, wall_width_feet, scale_info)
        if not walls:
            warnings.append("No walls were detected from close parallel vector lines.")

        floors, floor_warnings = build_floors(walls)
        warnings.extend(floor_warnings)

        doors = build_doors(curves, walls, pdf_to_revit, wall_width_feet)
        windows = build_windows(segments, walls, pdf_to_revit, wall_width_feet)

        if not doors:
            warnings.append("No door swing arcs were detected.")
        if not windows:
            warnings.append("No window line groups were detected.")

        return {
            "ok": True,
            "operation": "analyze",
            "page": page_payload(page),
            "calibration": scale_info,
            "elements": {
                "walls": walls,
                "floors": floors,
                "doors": doors,
                "windows": windows,
            },
            "warnings": warnings,
            "summary": {
                "rawSegmentCount": len(segments),
                "curveCount": len(curves),
                "wallCount": len(walls),
                "floorCount": len(floors),
                "doorCount": len(doors),
                "windowCount": len(windows),
            },
        }
    finally:
        document.close()


def dispatch(request):
    operation = request.get("operation")
    if operation == "info":
        return operation_info(request)
    if operation == "preview":
        return operation_preview(request)
    if operation == "analyze":
        return operation_analyze(request)
    raise ValueError("Unknown helper operation: {0}".format(operation))


def main(argv):
    if len(argv) != 3:
        sys.stderr.write("Usage: pdf2revit_helper.py request.json response.json\n")
        return 2

    request_path = argv[1]
    response_path = argv[2]

    try:
        request = read_json(request_path)
        response = dispatch(request)
        if "ok" not in response:
            response["ok"] = True
        write_json(response_path, response)
        return 0
    except Exception as exc:
        write_json(response_path, {
            "ok": False,
            "error": str(exc),
            "traceback": traceback.format_exc(),
        })
        return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
