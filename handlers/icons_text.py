"""
Icons & Text Handler — 3×2 grid of icon+label+description cells.

Template layout 27 "Icons & Text":
    idx  0 — Title (full width)
    idx 53 — Subtitle (full width, h=0.55in)

    Row 1 (top):
        idx 41 — Icon 1 (1.02in × 1.02in, left col)
        idx 19 — Label 1 (3.70in × 0.26in)
        idx 20 — Description 1 (3.70in × 0.71in)
        idx 42 — Icon 2 (1.02in × 1.02in, centre col)
        idx 24 — Label 2 (3.70in × 0.26in)
        idx 48 — Description 2 (3.70in × 0.71in)
        idx 44 — Icon 3 (1.02in × 1.02in, right col)
        idx 43 — Label 3 (3.70in × 0.26in)
        idx 49 — Description 3 (3.70in × 0.71in)

    Row 2 (bottom):
        idx 45 — Icon 4 (1.02in × 1.02in, left col)
        idx 30 — Label 4 (3.70in × 0.26in)
        idx 50 — Description 4 (3.70in × 0.71in)
        idx 46 — Icon 5 (1.02in × 1.02in, centre col)
        idx 33 — Label 5 (3.70in × 0.26in)
        idx 51 — Description 5 (3.70in × 0.71in)
        idx 47 — Icon 6 (1.02in × 1.02in, right col)
        idx 36 — Label 6 (3.70in × 0.26in)
        idx 52 — Description 6 (3.70in × 0.71in)

    idx 39 — Footer
    idx 40 — Slide number

Detection: slide with 6 short items (each <80 chars), typically
with consistent structure. Think "six features", "six values",
"six key points" etc. Icon placeholders are left empty.

Content to look for:
    - Feature overview (6 product/service features)
    - Company values (6 core values)
    - Team capabilities or service offerings
    - Course module overviews

CRITICAL RULE: Never set font properties on placeholders.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_shapes_with_text


class IconsTextHandler(SlideHandler):

    name = "Icons & Text"
    description = "3×2 grid of icon+label+description cells"
    layout_name = "Icons & Text"
    layout_index = 27

    PH_TITLE = 0
    PH_SUBTITLE = 53

    # (icon_idx, label_idx, description_idx) for each cell
    CELLS = [
        (41, 19, 20),  # Row 1, Col 1
        (42, 24, 48),  # Row 1, Col 2
        (44, 43, 49),  # Row 1, Col 3
        (45, 30, 50),  # Row 2, Col 1
        (46, 33, 51),  # Row 2, Col 2
        (47, 36, 52),  # Row 2, Col 3
    ]

    PH_FOOTER = 39
    PH_SLIDE_NUM = 40

    PLACEHOLDER_NOISE = [
        "[click", "[add", "[your", "[insert", "[subtitle",
        "click to add", "click icon", "click to edit",
        "\u00a9 the university", "this content is protected",
        "cricos code",
    ]

    FOOTER_PATTERNS = [
        r"(?i)cricos", r"(?i)hbis\s+innovation",
        r"(?i)executive\s+education",
        r"(?i)presentation\s+title",
    ]

    SLIDE_WIDTH = 12192000

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect slides with 6 short items in a grid.

        Content patterns:
        - Source layout name contains "icons" and "text"
        - 6+ short text items (each <80 chars)
        - Items distributed across columns and rows
        """
        if slide_index == 0:
            return 0.0

        layout_name = slide.slide_layout.name.lower()
        if "icon" in layout_name and "text" in layout_name:
            # But not "icons + two contents" which is a different layout
            if "two content" not in layout_name:
                return 0.75

        shapes = extract_shapes_with_text(slide)
        meaningful = self._get_meaningful_shapes(shapes)

        if len(meaningful) < 7:  # title + 6 items
            return 0.0

        title_shape = self._find_title(meaningful)
        body_shapes = [s for s in meaningful if s is not title_shape]

        # Filter to short items (could be labels or descriptions)
        short_items = [s for s in body_shapes if len(s["text"]) < 80]

        if len(short_items) < 6:
            return 0.0

        # Check distribution across columns
        third_w = self.SLIDE_WIDTH / 3
        left = [s for s in short_items if (s["left"] or 0) + (s["width"] or 0) / 2 < third_w]
        centre = [s for s in short_items if third_w <= (s["left"] or 0) + (s["width"] or 0) / 2 < 2 * third_w]
        right = [s for s in short_items if (s["left"] or 0) + (s["width"] or 0) / 2 >= 2 * third_w]

        if left and centre and right:
            # Check if items are also in two rows
            mid_y = sum(s["top"] or 0 for s in short_items) / len(short_items)
            top_row = [s for s in short_items if (s["top"] or 0) < mid_y]
            bottom_row = [s for s in short_items if (s["top"] or 0) >= mid_y]

            if len(top_row) >= 3 and len(bottom_row) >= 3:
                return 0.55
            elif len(short_items) >= 6:
                return 0.45

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, subtitle, and up to 6 items. Each item has a
        label and description. Items are extracted by position, sorted
        row-by-row, left-to-right.
        """
        shapes = extract_shapes_with_text(slide)

        result = {
            "title": "",
            "subtitle": "",
            "items": [],  # List of {"label": str, "description": str}
            "footer": "",
        }

        if not shapes:
            return result

        meaningful = self._get_meaningful_shapes(shapes)
        if not meaningful:
            return result

        title_shape = None
        body_shapes = []
        footer_text = ""

        for s in meaningful:
            if s["is_placeholder"]:
                idx = s["shape_idx"]
                if idx in (0, 3, 15):
                    title_shape = s
                elif idx in (17, 39):
                    if not footer_text:
                        footer_text = s["text"]
                elif idx in (18, 40):
                    pass
                else:
                    body_shapes.append(s)
            elif self._is_footer_text(s["text"]):
                if not footer_text:
                    footer_text = s["text"]
            else:
                body_shapes.append(s)

        result["footer"] = footer_text

        # Fallback title
        if not title_shape and body_shapes:
            with_font = [s for s in body_shapes if s["font_size"]]
            if with_font:
                candidate = max(with_font, key=lambda s: s["font_size"])
                if len(candidate["text"]) < 120:
                    title_shape = candidate
                    body_shapes = [s for s in body_shapes if s is not title_shape]

        if title_shape:
            result["title"] = title_shape["text"]

        # Subtitle
        body_shapes.sort(key=lambda s: (s["top"] or 0, s["left"] or 0))
        if body_shapes and title_shape:
            candidate = body_shapes[0]
            if len(candidate["text"]) < 100 and len(body_shapes) > 1:
                near_top = (candidate["top"] or 0) < (title_shape["top"] or 0) + (title_shape["height"] or 0) * 2
                if near_top:
                    result["subtitle"] = candidate["text"]
                    body_shapes = body_shapes[1:]

        # Group remaining shapes into label/description pairs
        # Sort by row then column
        body_shapes.sort(key=lambda s: (s["top"] or 0, s["left"] or 0))

        # Try to pair: short text above longer text at similar x position
        used = set()
        pairs = []

        for i, s in enumerate(body_shapes):
            if i in used:
                continue
            if len(s["text"]) >= 80:
                # This is a description without a label
                pairs.append({"label": "", "description": s["text"]})
                used.add(i)
                continue

            # Look for a nearby description below this label
            label_cx = (s["left"] or 0) + (s["width"] or 0) / 2
            label_bottom = (s["top"] or 0) + (s["height"] or 0)

            best_desc = None
            best_idx = None
            for j, d in enumerate(body_shapes):
                if j in used or j == i:
                    continue
                desc_cx = (d["left"] or 0) + (d["width"] or 0) / 2
                desc_top = d["top"] or 0
                # Same column (within 2 inches) and below
                if abs(label_cx - desc_cx) < 2 * 914400 and desc_top > (s["top"] or 0):
                    if best_desc is None or desc_top < (best_desc["top"] or 0):
                        best_desc = d
                        best_idx = j

            if best_desc:
                pairs.append({"label": s["text"], "description": best_desc["text"]})
                used.add(i)
                used.add(best_idx)
            else:
                # Label without description
                pairs.append({"label": s["text"], "description": ""})
                used.add(i)

        result["items"] = pairs[:6]

        return result

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill title, subtitle, and up to 6 label+description cells.
        Icon placeholders are left empty. NEVER set font properties.
        """
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        if content.get("subtitle") and self.PH_SUBTITLE in placeholders:
            placeholders[self.PH_SUBTITLE].text = content["subtitle"]

        items = content.get("items", [])
        for i, (icon_idx, label_idx, desc_idx) in enumerate(self.CELLS):
            if i >= len(items):
                break
            item = items[i]
            if item.get("label") and label_idx in placeholders:
                placeholders[label_idx].text = item["label"]
            if item.get("description") and desc_idx in placeholders:
                placeholders[desc_idx].text = item["description"]

    # --- Helpers ---

    def _get_meaningful_shapes(self, shapes: list) -> list:
        meaningful = []
        for s in shapes:
            text = s["text"].strip()
            if not text:
                continue
            if any(p in text.lower() for p in self.PLACEHOLDER_NOISE):
                continue
            if len(text) <= 1:
                continue
            if re.match(r"^\d{1,3}$", text) and len(text) <= 3:
                continue
            meaningful.append(s)
        return meaningful

    def _find_title(self, shapes: list):
        for s in shapes:
            if s["is_placeholder"] and s["shape_idx"] in (0, 3, 15):
                return s
        sorted_by_top = sorted(shapes, key=lambda s: s["top"] or 0)
        for s in sorted_by_top:
            if len(s["text"]) < 120:
                if s.get("font_size") and s["font_size"] >= 20:
                    return s
                elif not s.get("font_size") and len(s["text"]) < 80:
                    return s
        return None

    def _is_footer_text(self, text: str) -> bool:
        return any(re.search(p, text) for p in self.FOOTER_PATTERNS)

    def get_placeholder_map(self) -> dict:
        return {
            "title": self.PH_TITLE,
            "subtitle": self.PH_SUBTITLE,
            "cells": self.CELLS,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
