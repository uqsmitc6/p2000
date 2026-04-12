"""
Title, Subtitle, 2 Graphs Handler — Side-by-side graph/chart comparison.

Template layout 13 "Title, Subtitle, 2 Graphs":
    idx  0 — Title
    idx 31 — Subtitle
    idx 32 — Left graph label (tiny, 5.82in × 0.31in)
    idx 10 — Left graph/content area (5.81in × 4.02in)
    idx 34 — Right graph label (tiny, 5.82in × 0.31in)
    idx 33 — Right graph/content area (5.81in × 4.02in)
    idx 21 — Footer
    idx 22 — Slide number

Detection: slide has two charts, graphs, or images positioned side by
side with labels. Distinguished from Two Content by the presence of
visual chart/image content rather than text-heavy columns.

Content to look for:
    - Side-by-side charts or graphs (before/after, comparison)
    - Two data visualisations with labels
    - Paired images with descriptive headers

CRITICAL RULE: Never set font properties on placeholders.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_shapes_with_text, extract_images


class TwoGraphsHandler(SlideHandler):

    name = "Title, Subtitle, 2 Graphs"
    description = "Side-by-side graph/chart areas with title and labels"
    layout_name = "Title, Subtitle, 2 Graphs"
    layout_index = 13

    PH_TITLE = 0
    PH_SUBTITLE = 31
    PH_LEFT_LABEL = 32
    PH_LEFT_GRAPH = 10
    PH_RIGHT_LABEL = 34
    PH_RIGHT_GRAPH = 33
    PH_FOOTER = 21
    PH_SLIDE_NUM = 22

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
    SLIDE_MID = SLIDE_WIDTH // 2

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect slides with two side-by-side graphs/charts.

        Content patterns:
        - Source layout name contains "graph" (but not "block")
        - Two images positioned side by side (each taking ~half width)
        - Multiple group shapes side by side (embedded charts)
        """
        if slide_index == 0:
            return 0.0

        layout_name = slide.slide_layout.name.lower()
        if "2 graph" in layout_name or "two graph" in layout_name:
            return 0.75

        images = extract_images(slide)
        shapes = extract_shapes_with_text(slide)

        if not images:
            # Check for group shapes (embedded charts often become groups)
            group_count = sum(1 for s in slide.shapes if s.shape_type == 6)
            if group_count >= 2:
                meaningful = self._get_meaningful_shapes(shapes)
                title = self._find_title(meaningful) if meaningful else None
                if title:
                    return 0.50
            return 0.0

        # Look for two images side by side
        real_images = [
            img for img in images
            if img.get("width") and img.get("height")
            and img["width"] > self.SLIDE_WIDTH * 0.2
            and img["height"] > 1000000  # >1 inch tall
        ]

        if len(real_images) >= 2:
            # Check if they're positioned side by side
            sorted_by_left = sorted(real_images, key=lambda i: i.get("left") or 0)
            left_img = sorted_by_left[0]
            right_img = sorted_by_left[-1]

            left_centre = (left_img["left"] or 0) + (left_img["width"] or 0) / 2
            right_centre = (right_img["left"] or 0) + (right_img["width"] or 0) / 2

            if left_centre < self.SLIDE_MID and right_centre > self.SLIDE_MID:
                return 0.55

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, subtitle, graph labels, and any text content
        for the graph areas. Images/charts are preserved by
        _preserve_visual_shapes().
        """
        shapes = extract_shapes_with_text(slide)

        result = {
            "title": "",
            "subtitle": "",
            "left_label": "",
            "left_content": "",
            "right_label": "",
            "right_content": "",
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
                elif idx in (17, 21):
                    if not footer_text:
                        footer_text = s["text"]
                elif idx in (18, 22):
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

        # Split remaining into left/right
        left_texts = []
        right_texts = []

        for s in body_shapes:
            shape_centre = (s["left"] or 0) + (s["width"] or 0) / 2
            if shape_centre < self.SLIDE_MID:
                left_texts.append(s)
            else:
                right_texts.append(s)

        # First short text in each column = label, rest = content
        for texts, label_key, content_key in [
            (left_texts, "left_label", "left_content"),
            (right_texts, "right_label", "right_content"),
        ]:
            texts.sort(key=lambda s: (s["top"] or 0))
            if texts:
                if len(texts[0]["text"]) < 60:
                    result[label_key] = texts[0]["text"]
                    result[content_key] = "\n".join(s["text"] for s in texts[1:])
                else:
                    result[content_key] = "\n".join(s["text"] for s in texts)

        return result

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """Fill 2 Graphs placeholders. NEVER set font properties."""
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        if content.get("subtitle") and self.PH_SUBTITLE in placeholders:
            placeholders[self.PH_SUBTITLE].text = content["subtitle"]

        if content.get("left_label") and self.PH_LEFT_LABEL in placeholders:
            placeholders[self.PH_LEFT_LABEL].text = content["left_label"]

        if content.get("left_content") and self.PH_LEFT_GRAPH in placeholders:
            placeholders[self.PH_LEFT_GRAPH].text = content["left_content"]

        if content.get("right_label") and self.PH_RIGHT_LABEL in placeholders:
            placeholders[self.PH_RIGHT_LABEL].text = content["right_label"]

        if content.get("right_content") and self.PH_RIGHT_GRAPH in placeholders:
            placeholders[self.PH_RIGHT_GRAPH].text = content["right_content"]

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
                return s
        return None

    def _is_footer_text(self, text: str) -> bool:
        return any(re.search(p, text) for p in self.FOOTER_PATTERNS)

    def get_placeholder_map(self) -> dict:
        return {
            "title": self.PH_TITLE,
            "subtitle": self.PH_SUBTITLE,
            "left_label": self.PH_LEFT_LABEL,
            "left_graph": self.PH_LEFT_GRAPH,
            "right_label": self.PH_RIGHT_LABEL,
            "right_graph": self.PH_RIGHT_GRAPH,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
