"""
Multi-layout 2 Handler — Two stacked content blocks left, two labelled
content blocks right.

Template layout 37 "Multi-layout 2":
    idx  0 — Title (top, w=9.76in)
    idx 31 — Subtitle (full width, w=9.76in, h=0.55in)
    idx 10 — Left top content (OBJECT, 5.90in × 2.15in)
    idx 36 — Left bottom content (OBJECT, 5.90in × 2.15in)
    idx 35 — Right top label (BODY, 5.91in × 0.35in)
    idx 19 — Right top content (OBJECT, 5.90in × 1.71in)
    idx 26 — Right bottom label (BODY, 5.91in × 0.35in)
    idx 20 — Right bottom content (OBJECT, 5.90in × 1.85in)
    idx 33 — Footer
    idx 34 — Slide number

Detection: slide with 4 content areas split into 2 stacked left and
2 labelled stacked right. Very similar to Multi-layout 1 but with
labels on the right column.

Content to look for:
    - Comparison slides: two items on left vs two items on right
    - Dashboard with labelled data panels
    - Before/after with paired descriptions
    - Case studies with categorised findings

CRITICAL RULE: Never set font properties on placeholders.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_shapes_with_text


class MultiLayout2Handler(SlideHandler):

    name = "Multi-layout 2"
    description = "Two stacked content blocks left, two labelled content blocks right"
    layout_name = "Multi-layout 2"
    layout_index = 37

    PH_TITLE = 0
    PH_SUBTITLE = 31
    PH_LEFT_TOP = 10
    PH_LEFT_BOTTOM = 36
    PH_RIGHT_TOP_LABEL = 35
    PH_RIGHT_TOP = 19
    PH_RIGHT_BOTTOM_LABEL = 26
    PH_RIGHT_BOTTOM = 20
    PH_FOOTER = 33
    PH_SLIDE_NUM = 34

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
    SLIDE_HEIGHT = 6858000
    SLIDE_MID_X = SLIDE_WIDTH // 2
    SLIDE_MID_Y = SLIDE_HEIGHT // 2

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect slides matching multi-layout 2 pattern.

        Content patterns:
        - Source layout name contains "multi-layout 2" or "multi layout 2"
        - 4+ text areas in a 2×2 grid (similar to Multi-layout 1)
        - Right side has short labels above longer content
        """
        if slide_index == 0:
            return 0.0

        layout_name = slide.slide_layout.name.lower()
        if "multi-layout 2" in layout_name or "multi layout 2" in layout_name:
            return 0.75

        # Content-based heuristic same as Multi-layout 1 but at lower confidence
        # since Multi-layout 1 already covers the 2×2 grid pattern
        shapes = extract_shapes_with_text(slide)
        meaningful = self._get_meaningful_shapes(shapes)

        if len(meaningful) < 5:
            return 0.0

        title_shape = self._find_title(meaningful)
        body_shapes = [s for s in meaningful if s is not title_shape]

        if len(body_shapes) < 4:
            return 0.0

        # Check for labelled right column pattern:
        # short texts (<60 chars) directly above longer texts on the right
        right_shapes = [
            s for s in body_shapes
            if (s["left"] or 0) + (s["width"] or 0) / 2 >= self.SLIDE_MID_X
        ]
        right_shapes.sort(key=lambda s: (s["top"] or 0))

        label_content_pairs = 0
        for i in range(len(right_shapes) - 1):
            curr = right_shapes[i]
            nxt = right_shapes[i + 1]
            if len(curr["text"]) < 60 and len(nxt["text"]) >= 30:
                label_content_pairs += 1

        left_shapes = [
            s for s in body_shapes
            if (s["left"] or 0) + (s["width"] or 0) / 2 < self.SLIDE_MID_X
        ]

        if label_content_pairs >= 2 and len(left_shapes) >= 2:
            return 0.52

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, subtitle, two left content blocks, and two
        labelled right content blocks.
        """
        shapes = extract_shapes_with_text(slide)

        result = {
            "title": "",
            "subtitle": "",
            "left_top": "",
            "left_bottom": "",
            "right_top_label": "",
            "right_top": "",
            "right_bottom_label": "",
            "right_bottom": "",
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
                elif idx in (17, 33):
                    if not footer_text:
                        footer_text = s["text"]
                elif idx in (18, 34):
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

        # Split left/right
        left_shapes = sorted(
            [s for s in body_shapes
             if (s["left"] or 0) + (s["width"] or 0) / 2 < self.SLIDE_MID_X],
            key=lambda s: (s["top"] or 0),
        )
        right_shapes = sorted(
            [s for s in body_shapes
             if (s["left"] or 0) + (s["width"] or 0) / 2 >= self.SLIDE_MID_X],
            key=lambda s: (s["top"] or 0),
        )

        # Left: split into top/bottom halves
        if left_shapes:
            mid = len(left_shapes) // 2 or 1
            result["left_top"] = "\n".join(s["text"] for s in left_shapes[:mid])
            result["left_bottom"] = "\n".join(s["text"] for s in left_shapes[mid:])

        # Right: try label/content pairs
        if right_shapes:
            # First pair
            if len(right_shapes) >= 2:
                if len(right_shapes[0]["text"]) < 60:
                    result["right_top_label"] = right_shapes[0]["text"]
                    idx = 1
                else:
                    idx = 0

                # Collect content until next short label or end
                content_parts = []
                label2_idx = None
                for i in range(idx, len(right_shapes)):
                    if i > idx and len(right_shapes[i]["text"]) < 60 and label2_idx is None:
                        label2_idx = i
                        break
                    content_parts.append(right_shapes[i]["text"])
                result["right_top"] = "\n".join(content_parts)

                # Second pair
                if label2_idx is not None and label2_idx < len(right_shapes):
                    result["right_bottom_label"] = right_shapes[label2_idx]["text"]
                    remaining = [s["text"] for s in right_shapes[label2_idx + 1:]]
                    result["right_bottom"] = "\n".join(remaining)
                elif not content_parts:
                    # No label pattern — just split in half
                    mid = len(right_shapes) // 2 or 1
                    result["right_top"] = "\n".join(s["text"] for s in right_shapes[:mid])
                    result["right_bottom"] = "\n".join(s["text"] for s in right_shapes[mid:])
            else:
                result["right_top"] = right_shapes[0]["text"]

        return result

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """Fill multi-layout 2 placeholders. NEVER set font properties."""
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        if content.get("subtitle") and self.PH_SUBTITLE in placeholders:
            placeholders[self.PH_SUBTITLE].text = content["subtitle"]

        if content.get("left_top") and self.PH_LEFT_TOP in placeholders:
            placeholders[self.PH_LEFT_TOP].text = content["left_top"]

        if content.get("left_bottom") and self.PH_LEFT_BOTTOM in placeholders:
            placeholders[self.PH_LEFT_BOTTOM].text = content["left_bottom"]

        if content.get("right_top_label") and self.PH_RIGHT_TOP_LABEL in placeholders:
            placeholders[self.PH_RIGHT_TOP_LABEL].text = content["right_top_label"]

        if content.get("right_top") and self.PH_RIGHT_TOP in placeholders:
            placeholders[self.PH_RIGHT_TOP].text = content["right_top"]

        if content.get("right_bottom_label") and self.PH_RIGHT_BOTTOM_LABEL in placeholders:
            placeholders[self.PH_RIGHT_BOTTOM_LABEL].text = content["right_bottom_label"]

        if content.get("right_bottom") and self.PH_RIGHT_BOTTOM in placeholders:
            placeholders[self.PH_RIGHT_BOTTOM].text = content["right_bottom"]

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
            "left_top": self.PH_LEFT_TOP,
            "left_bottom": self.PH_LEFT_BOTTOM,
            "right_top_label": self.PH_RIGHT_TOP_LABEL,
            "right_top": self.PH_RIGHT_TOP,
            "right_bottom_label": self.PH_RIGHT_BOTTOM_LABEL,
            "right_bottom": self.PH_RIGHT_BOTTOM,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
