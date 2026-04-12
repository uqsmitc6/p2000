"""
Two Content Horizontal Handler — Slides with stacked top/bottom content rows.

Template layout 14 "Two Content Layout Horizontal":
    idx  0 — Title
    idx 31 — Subtitle (optional)
    idx 10 — Top content row   (full width, top=2.46in, h=2.08in)
    idx 34 — Bottom content row (full width, top=4.90in, h=2.08in)
    idx 19 — Footer (programme name)
    idx 20 — Slide number

Detection: slide has two distinct content groups stacked vertically
(NOT side-by-side — that's Two Content). May also catch slides where
the source layout name indicates horizontal/stacked two-row arrangement.

Distinguished from Two Content by vertical rather than horizontal
arrangement of body content areas.

CRITICAL RULE: Never set font properties on placeholders.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_shapes_with_text


class TwoContentHorizontalHandler(SlideHandler):

    name = "Two Content Horizontal"
    description = "Content slide with two stacked (top/bottom) content rows"
    layout_name = "Two Content Layout Horizontal"
    layout_index = 14

    PH_TITLE = 0
    PH_SUBTITLE = 31
    PH_TOP = 10
    PH_BOTTOM = 34
    PH_FOOTER = 19
    PH_SLIDE_NUM = 20

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

    SLIDE_HEIGHT = 6858000  # Standard 7.5in EMU

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect slides with stacked (top/bottom) content layout.

        Heuristics:
        - Source layout name contains "horizontal" or "stacked"
        - Two body-sized shapes at same horizontal position but
          different vertical positions, both near full width
        """
        if slide_index == 0:
            return 0.0

        # Check source layout name
        layout_name = slide.slide_layout.name.lower()
        if "horizontal" in layout_name and ("two" in layout_name or "content" in layout_name):
            return 0.75

        shapes = extract_shapes_with_text(slide)
        if not shapes:
            return 0.0

        meaningful = self._get_meaningful_shapes(shapes)
        if len(meaningful) < 3:  # title + 2 body areas
            return 0.0

        # Find title
        title_shape = self._find_title(meaningful)
        body_candidates = [s for s in meaningful if s is not title_shape]

        if len(body_candidates) < 2:
            return 0.0

        # Look for two body shapes stacked vertically (similar left, different top)
        slide_width = 12192000
        mid_height = self.SLIDE_HEIGHT / 2

        for i, a in enumerate(body_candidates):
            for b in body_candidates[i + 1:]:
                if a["left"] is None or b["left"] is None:
                    continue
                if a["top"] is None or b["top"] is None:
                    continue

                # Similar horizontal position (within 20% of slide width)
                left_diff = abs(a["left"] - b["left"])
                if left_diff > slide_width * 0.2:
                    continue

                # Different vertical positions (separated by at least 1.5in)
                top_diff = abs(a["top"] - b["top"])
                if top_diff < 1371600:  # 1.5in in EMU
                    continue

                # Both should be near full width (>60% of slide)
                a_wide = (a["width"] or 0) > slide_width * 0.6
                b_wide = (b["width"] or 0) > slide_width * 0.6
                if not (a_wide and b_wide):
                    continue

                # Both should have substantial text
                if len(a["text"]) < 20 or len(b["text"]) < 20:
                    continue

                return 0.55

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, subtitle, and two rows of stacked content.
        Top/bottom assignment is based on vertical position.
        """
        shapes = extract_shapes_with_text(slide)

        result = {
            "title": "",
            "subtitle": "",
            "top_content": "",
            "bottom_content": "",
            "footer": "",
        }

        if not shapes:
            return result

        filtered = self._get_meaningful_shapes(shapes)
        if not filtered:
            return result

        # Separate by role
        title_shape = None
        body_shapes = []
        footer_text = ""

        for s in filtered:
            if s["is_placeholder"]:
                idx = s["shape_idx"]
                if idx in (0, 3, 15):
                    title_shape = s
                elif idx in (17, 19):
                    if not footer_text:
                        footer_text = s["text"]
                elif idx in (18, 20):
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

        # Subtitle detection
        body_shapes.sort(key=lambda s: (s["top"] or 0, s["left"] or 0))
        if body_shapes and title_shape:
            candidate = body_shapes[0]
            is_short = len(candidate["text"]) < 100
            is_near_top = candidate["top"] < (title_shape["top"] + title_shape["height"] * 2)
            has_more = len(body_shapes) > 1
            if is_short and is_near_top and has_more:
                result["subtitle"] = candidate["text"]
                body_shapes = body_shapes[1:]

        # Split into top/bottom rows based on vertical midpoint
        if body_shapes:
            # Find the vertical midpoint of all body content
            all_tops = [s["top"] for s in body_shapes if s["top"] is not None]
            if all_tops:
                min_top = min(all_tops)
                max_top = max(all_tops)
                mid_top = (min_top + max_top) / 2

                top_shapes = [s for s in body_shapes if (s["top"] or 0) <= mid_top]
                bottom_shapes = [s for s in body_shapes if (s["top"] or 0) > mid_top]

                # Sort each row by position
                top_shapes.sort(key=lambda s: (s["top"] or 0, s["left"] or 0))
                bottom_shapes.sort(key=lambda s: (s["top"] or 0, s["left"] or 0))

                result["top_content"] = "\n".join(s["text"] for s in top_shapes)
                result["bottom_content"] = "\n".join(s["text"] for s in bottom_shapes)

                # If everything ended up in one row, split evenly
                if result["top_content"] and not result["bottom_content"]:
                    all_text = [s["text"] for s in top_shapes]
                    mid = len(all_text) // 2
                    if mid > 0:
                        result["top_content"] = "\n".join(all_text[:mid])
                        result["bottom_content"] = "\n".join(all_text[mid:])

        return result

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """Fill Two Content Horizontal placeholders. NEVER set font properties."""
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        if content.get("subtitle") and self.PH_SUBTITLE in placeholders:
            placeholders[self.PH_SUBTITLE].text = content["subtitle"]

        if content.get("top_content") and self.PH_TOP in placeholders:
            placeholders[self.PH_TOP].text = content["top_content"]

        if content.get("bottom_content") and self.PH_BOTTOM in placeholders:
            placeholders[self.PH_BOTTOM].text = content["bottom_content"]

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
            "top": self.PH_TOP,
            "bottom": self.PH_BOTTOM,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
