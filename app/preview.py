from __future__ import annotations

import hashlib
import math
import platform
import tempfile
from pathlib import Path
from typing import List, Sequence, Tuple

from PIL import Image, ImageDraw, ImageFont

from .models import PageRecord, ShapeRecord


def font(size: int, bold: bool = False):
    candidates = (
        ["arialbd.ttf", "segoeuib.ttf"] if bold else ["arial.ttf", "segoeui.ttf"]
    )
    for candidate in candidates:
        try:
            return ImageFont.truetype(candidate, max(8, size))
        except Exception:
            continue
    return ImageFont.load_default()


def truncate(text: str, limit: int = 80) -> str:
    text = " ".join((text or "").split())
    return text if len(text) <= limit else text[: limit - 1] + "…"


def centroid(points: Sequence[Tuple[float, float]]) -> Tuple[float, float]:
    if not points:
        return 0.0, 0.0
    return (
        sum(point[0] for point in points) / len(points),
        sum(point[1] for point in points) / len(points),
    )


def wrap_text(draw: ImageDraw.ImageDraw, text: str, text_font, max_width: float, max_lines: int) -> str:
    words = " ".join((text or "").split()).split()
    if not words or max_width <= 5 or max_lines <= 0:
        return ""

    lines: List[str] = []
    current = ""
    for word in words:
        candidate = f"{current} {word}".strip()
        box = draw.textbbox((0, 0), candidate, font=text_font)
        if box[2] - box[0] <= max_width or not current:
            current = candidate
        else:
            lines.append(current)
            current = word
            if len(lines) >= max_lines:
                break

    if current and len(lines) < max_lines:
        lines.append(current)

    used_words = " ".join(lines).split()
    if len(used_words) < len(words) and lines:
        last = lines[-1]
        while last:
            candidate = last.rstrip("…") + "…"
            box = draw.textbbox((0, 0), candidate, font=text_font)
            if box[2] - box[0] <= max_width:
                lines[-1] = candidate
                break
            last = last[:-1]

    return "\n".join(lines)


def shape_style(shape: ShapeRecord):
    name = f"{shape.name} {shape.name_u}".lower()
    text = (shape.text or "").lower()

    fill = "#F8FAFC"
    outline = "#64748B"
    text_color = "#0F172A"

    if shape.workcenters:
        fill, outline = "#DCFCE7", "#16A34A"
    elif shape.kind == "decision":
        fill, outline = "#FFF7ED", "#EA580C"
    elif shape.kind == "start_end":
        fill, outline = "#EFF6FF", "#2563EB"
    elif shape.kind == "document":
        fill, outline = "#F5F3FF", "#7C3AED"
    elif "note" in name or "annotation" in name:
        fill, outline = "#FEFCE8", "#CA8A04"
    elif any(role in text for role in ("millwright", "warehouse", "quality", "planner")):
        fill, outline = "#F0FDF4", "#22C55E"

    return fill, outline, text_color


def render(page: PageRecord, width: int = 1200, height: int = 720, zoom: float = 1.0):
    width = max(500, int(width * zoom))
    height = max(320, int(height * zoom))
    image = Image.new("RGB", (width, height), "#E2E8F0")
    draw = ImageDraw.Draw(image)

    margin = 30
    page_width = max(page.width, 1.0)
    page_height = max(page.height, 1.0)
    scale = min(
        (width - 2 * margin) / page_width,
        (height - 2 * margin) / page_height,
    )
    offset_x = (width - page_width * scale) / 2
    offset_y = (height - page_height * scale) / 2

    def point(x: float, y: float) -> Tuple[float, float]:
        return offset_x + x * scale, offset_y + (page_height - y) * scale

    draw.rectangle(
        [
            offset_x,
            offset_y,
            offset_x + page_width * scale,
            offset_y + page_height * scale,
        ],
        fill="white",
        outline="#CBD5E1",
        width=1,
    )

    if not page.loaded:
        message = "Loading selected Visio tab…" if page.loading else "Select or reload this tab to build the visual preview."
        message_font = font(15, bold=True)
        box = draw.textbbox((0, 0), message, font=message_font)
        draw.text(
            ((width - (box[2] - box[0])) / 2, (height - (box[3] - box[1])) / 2),
            message,
            fill="#475569",
            font=message_font,
        )
    else:
        # Connector geometry first.
        for shape in page.shapes:
            if shape.kind != "connector":
                continue
            if None in (shape.begin_x, shape.begin_y, shape.end_x, shape.end_y):
                continue
            start = point(shape.begin_x, shape.begin_y)
            end = point(shape.end_x, shape.end_y)
            line_width = max(1, min(4, int(scale * 0.018)))
            draw.line([start, end], fill="#94A3B8", width=line_width)

            angle = math.atan2(end[1] - start[1], end[0] - start[0])
            arrow_size = max(4.0, min(10.0, scale * 0.07))
            left = (
                end[0] - arrow_size * math.cos(angle - 0.45),
                end[1] - arrow_size * math.sin(angle - 0.45),
            )
            right = (
                end[0] - arrow_size * math.cos(angle + 0.45),
                end[1] - arrow_size * math.sin(angle + 0.45),
            )
            draw.polygon([end, left, right], fill="#94A3B8")

        # Draw leaf shapes. Group shells are skipped so their children remain visible.
        drawable = [
            shape
            for shape in page.shapes
            if shape.kind != "connector"
            and not shape.has_children
            and shape.corners
        ]
        drawable.sort(key=lambda shape: (shape.depth, shape.shape_id))

        for shape in drawable:
            polygon = [point(x, y) for x, y in shape.corners]
            fill, outline, text_color = shape_style(shape)
            min_x = min(point_[0] for point_ in polygon)
            max_x = max(point_[0] for point_ in polygon)
            min_y = min(point_[1] for point_ in polygon)
            max_y = max(point_[1] for point_ in polygon)
            pixel_width = max_x - min_x
            pixel_height = max_y - min_y

            if pixel_width < 1 or pixel_height < 1:
                continue

            if shape.kind == "decision":
                midpoints = [
                    ((polygon[0][0] + polygon[1][0]) / 2, (polygon[0][1] + polygon[1][1]) / 2),
                    ((polygon[1][0] + polygon[2][0]) / 2, (polygon[1][1] + polygon[2][1]) / 2),
                    ((polygon[2][0] + polygon[3][0]) / 2, (polygon[2][1] + polygon[3][1]) / 2),
                    ((polygon[3][0] + polygon[0][0]) / 2, (polygon[3][1] + polygon[0][1]) / 2),
                ]
                draw.polygon(midpoints, fill=fill, outline=outline)
            elif shape.kind == "start_end" and abs(shape.angle or 0) < 0.02:
                draw.rounded_rectangle(
                    [min_x, min_y, max_x, max_y],
                    radius=max(3, int(min(pixel_width, pixel_height) * 0.35)),
                    fill=fill,
                    outline=outline,
                    width=1,
                )
            else:
                draw.polygon(polygon, fill=fill, outline=outline)

            if shape.text and pixel_width > 18 and pixel_height > 8:
                text_size = max(
                    7,
                    min(
                        13,
                        int(min(pixel_height / 4.0, pixel_width / 13.0)),
                    ),
                )
                text_font = font(text_size)
                max_lines = max(1, int((pixel_height - 4) / max(text_size + 2, 1)))
                wrapped = wrap_text(
                    draw,
                    shape.text,
                    text_font,
                    max(pixel_width - 6, 4),
                    min(max_lines, 5),
                )
                if wrapped:
                    box = draw.multiline_textbbox(
                        (0, 0), wrapped, font=text_font, spacing=1, align="center"
                    )
                    text_width = box[2] - box[0]
                    text_height = box[3] - box[1]
                    center_x, center_y = centroid(polygon)
                    draw.multiline_text(
                        (center_x - text_width / 2, center_y - text_height / 2),
                        wrapped,
                        fill=text_color,
                        font=text_font,
                        spacing=1,
                        align="center",
                    )

    title = f"{page.name} | Visual preview"
    badge_width = min(width - 32, 700)
    draw.rounded_rectangle([16, 12, badge_width, 44], radius=8, fill="#0F172A")
    draw.text((28, 21), truncate(title, 95), fill="white", font=font(11, bold=True))

    if page.loaded:
        summary = (
            f"{page.drawable_shape_count:,} visible shapes  •  "
            f"{len(page.connections):,} connections  •  "
            f"{len(page.workcenter_counts):,} workcenters"
        )
        summary_font = font(9)
        summary_box = draw.textbbox((0, 0), summary, font=summary_font)
        summary_width = summary_box[2] - summary_box[0] + 20
        draw.rounded_rectangle(
            [width - summary_width - 16, 12, width - 16, 40],
            radius=7,
            fill="#F8FAFC",
            outline="#CBD5E1",
        )
        draw.text(
            (width - summary_width - 6, 20),
            summary,
            fill="#334155",
            font=summary_font,
        )

    return image


class VisioRenderer:
    @staticmethod
    def supported() -> bool:
        if platform.system() != "Windows":
            return False
        try:
            import win32com.client  # noqa: F401

            return True
        except Exception:
            return False

    @staticmethod
    def cache(file_path: str, page: PageRecord) -> Path:
        root = Path(tempfile.gettempdir()) / "vsdx_metadata_previews_v12"
        root.mkdir(parents=True, exist_ok=True)
        stat = Path(file_path).stat()
        key = (
            f"{Path(file_path).resolve()}|{stat.st_mtime_ns}|"
            f"{page.page_id}|{page.name}"
        )
        return root / (hashlib.sha1(key.encode()).hexdigest() + ".png")

    @classmethod
    def export(cls, file_path: str, page: PageRecord) -> Path:
        if not cls.supported():
            raise RuntimeError("Microsoft Visio COM preview is unavailable.")

        output = cls.cache(file_path, page)
        if output.exists() and output.stat().st_size:
            return output

        import win32com.client

        application = document = None
        try:
            application = win32com.client.DispatchEx("Visio.Application")
            application.Visible = False
            document = application.Documents.OpenEx(str(Path(file_path).resolve()), 66)
            selected = None
            for index in range(1, document.Pages.Count + 1):
                candidate = document.Pages.Item(index)
                if str(candidate.Name) == page.name or str(candidate.NameU) == page.name_u:
                    selected = candidate
                    break
            if selected is None:
                raise RuntimeError(f"Page not found: {page.name}")
            selected.Export(str(output))
            return output
        finally:
            if document:
                try:
                    document.Close()
                except Exception:
                    pass
            if application:
                try:
                    application.Quit()
                except Exception:
                    pass
