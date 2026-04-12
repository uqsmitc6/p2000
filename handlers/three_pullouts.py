"""
Three Pullouts Handler — Three side-by-side callout/highlight areas.

Template layout 35 "Three Pullouts":
    idx  0 — Title
    idx 34 — Subtitle
    idx 19 — Left pullout (left=0.54in, w=3.92in, h=4.02in)
    idx 20 — Centre pullout (left=4.71in, w=3.92in, h=4.02in)
    idx 21 — Right pullout (left=8.87in, w=3.92in, h=4.02in)
    idx 23 — Footer
    idx 24 — Slide number

Detection: slide has three distinct highlighted/callout content areas.
Similar to Three Content but styled as callouts/cards rather than
plain columns. Distinguished by shorter, more structured text in
each area (key stats, pillars, feature highlights).

Content to look for:
    - Three key points, stats, or features presented equally
    - Pillar slides (3 organisational pillars)
    - Comparison cards
    - "Three things" format common in business presentations

CRITICAL RULE: Never set font properties on placeholders.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_shapes_with_text


class ThreePulloutsHandler(SlideHandler):

    name = "Three Pullouts"
    description = "Three side-by-side callout/highlight areas"
    layout_name = "Three Pullouts"
    layout_index = 35

    PH_TITLE = 0
    PH_SUBTITLE = 34
    PH_LEFT = 19
    PH_CENTRE = 20
    PH_RIGHT = 21
    PH_FOOTER = 23
    PH_SLIDE_NUM = 24

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
        Detect slides with three callout/highlight areas.

        Content patterns:
        - Source layout name contains "pullout" or "callout" or "highlight"
        - Three short text areas (each <200 chars) spanning left/centre/right
        - Similar structure/length in each area (balanced presentation)
        """
        if slide_index == 0:
            return 0.0

        layout_name = slide.slide_layout.name.lower()
        if any(kw in layout_name for kw in ("pullout", "callout", "highlight", "three pill")):
            return 0.75

        shapes = extract_shapes_with_text(slide)
        meaningful = self._get_meaningful_shapes(shapes)

        if len(meaningful) < 4:  # title + 3 areas
            return 0.0

        # Find title
        title_shape = self._find_title(meaningful)
        body_shapes = [s for s in meaningful if s is not title_shape]

        if len(body_shapes) < 3:
            return 0.0

        # Split into columns
        third_w = self.SLIDE_WIDTH / 3
        left = [s for s in body_shapes if (s["left"] or 0) + (s["width"] or 0) / 2 < third_w]
        centre = [s for s in body_shapes if third_w <= (s["left"] or 0) + (s["width"] or 0) / 2 < 2 * third_w]
        right = [s for s in body_shapes if (s["left"] or 0) + (s["width"] or 0) / 2 >= 2 * third_w]

        if not (left and centre and right):
            return 0.0

        # Pullouts are typically SHORT and balanced — each area <200 chars
        left_len = sum(len(s["text"]) for s in left)
        centre_len = sum(len(s["text"]) for s in centre)
        right_len = sum(len(s["text"]) for s in right)

        all_short = left_len < 200 and centre_len < 200 and right_len < 200

        # Check balance — pullouts should be roughly similar length
        lengths = [left_len, centre_len, right_len]
        max_len = max(lengths)
        min_len = min(lengths)
        balanced = min_len > max_len * 0.3 if max_len > 0 else False

        if all_short and balanced and min_len >= 10:
            return 0.55
        elif all_short and min_len >= 10:
            return 0.45

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, subtitle, and three pullout areas.
        """
        shapes = extract_shapes_with_text(slide)

        result = {
            "title": "",
            "subtitle": "",
            "left": "",
            "centre": "",
            "right": "",
            "footer": "",
        }

        if not shapes:
            return result

        meaningful = self._get_meaningful_shapes(shapes)
        if not meaningful:
            return result

        # Separate by role
        title_shape = None
        body_shapes = []
        footer_text = ""

        for s in meaningful:
            if s["is_placeholder"]:
                idx = s["shape_idx"]
                if idx in (0, 3, 15):
                    title_shape = s
                elif idx in (17, 23):
                    if not footer_text:
                        footer_text = s["text"]
                elif idx in (18, 24):
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
            is_short = len(candidate["text"]) < 100
            is_near_top = candidate["top"] < (title_shape["top"] + title_shape["height"] * 2)
            has_more = len(body_shapes) > 1
            if is_short and is_near_top and has_more:
                result["subtitle"] = candidate["text"]
                body_shapes = body_shapes[1:]

        # Split into three columns
        third_w = self.SLIDE_WIDTH / 3
        left_texts = []
        centre_texts = []
        right_texts = []

        for s in body_shapes:
            shape_centre = (s["left"] or 0) + (s["width"] or 0) / 2
            if shape_centre < third_w:
                left_texts.append(s)
            elif shape_centre < 2 * third_w:
                centre_texts.append(s)
            else:
                right_texts.append(s)

        for group in [left_texts, centre_texts, right_texts]:
            group.sort(key=lambda s: (s["top"] or 0))

        result["left"] = "\n".join(s["text"] for s in left_texts)
        result["centre"] = "\n".join(s["text"] for s in centre_texts)
        result["right"] = "\n".join(s["text"] for s in right_texts)

        # Redistribute if imbalanced
        cols = [result["left"], result["centre"], result["right"]]
        non_empty = [c for c in cols if c]
        if len(non_empty) < 3 and body_shapes:
            all_texts = [s["text"] for s in body_shapes]
            n = len(all_texts)
            third = max(1, n // 3)
            result["left"] = "\n".join(all_texts[:third])
            result["centre"] = "\n".join(all_texts[third:2*third])
            result["right"] = "\n".join(all_texts[2*third:])

        return result

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """Fill Three Pullouts placeholders. NEVER set font properties."""
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        if content.get("subtitle") and self.PH_SUBTITLE in placeholders:
            placeholders[self.PH_SUBTITLE].text = content["subtitle"]

        if content.get("left") and self.PH_LEFT in placeholders:
            placeholders[self.PH_LEFT].text = content["left"]

        if content.get("centre") and self.PH_CENTRE in placeholders:
            placeholders[self.PH_CENTRE].text = content["centre"]

        if content.get("right") and self.PH_RIGHT in placeholders:
            placeholders[self.PH_RIGHT].text = content["right"]

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
            "left": self.PH_LEFT,
            "centre": self.PH_CENTRE,
            "right": self.PH_RIGHT,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
