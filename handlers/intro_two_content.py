"""
Intro + Two Content Handler — Title and intro text on left, two columns on right.

Template layout 11 "Intro + two content layout":
    idx  0 — Title (left column, top, w=3.92in, h=1.50in)
    idx 31 — Intro/body text (left column, below title, w=4.10in, h=2.99in)
    idx 37 — Centre content column (left=4.79in, w=3.81in, h=4.96in)
    idx 36 — Right content column (left=8.89in, w=3.92in, h=4.96in)
    idx 17 — Footer
    idx 18 — Slide number

Detection: slide has a title + introductory paragraph on the left with
two content columns on the right. Distinguished from Two Content by
the presence of the intro text area alongside the columns.

CRITICAL RULE: Never set font properties on placeholders.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_shapes_with_text


class IntroTwoContentHandler(SlideHandler):

    name = "Intro + Two Content"
    description = "Title and introductory text with two content columns"
    layout_name = "Intro + two content layout"
    layout_index = 11

    PH_TITLE = 0
    PH_INTRO = 31       # Intro/body text (left column)
    PH_COL_CENTRE = 37  # Centre column
    PH_COL_RIGHT = 36   # Right column
    PH_FOOTER = 17
    PH_SLIDE_NUM = 18

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
        Detect slides with intro text + two content columns.
        """
        if slide_index == 0:
            return 0.0

        layout_name = slide.slide_layout.name.lower()
        if "intro" in layout_name and ("two" in layout_name or "content" in layout_name):
            return 0.75

        # Heuristic: title + intro paragraph on left + two body areas centre/right
        shapes = extract_shapes_with_text(slide)
        meaningful = self._get_meaningful_shapes(shapes)

        if len(meaningful) < 3:
            return 0.0

        third_w = self.SLIDE_WIDTH / 3

        # Count substantial text in each third
        left_shapes = [s for s in meaningful if (s["left"] or 0) + (s["width"] or 0) / 2 < third_w]
        centre_shapes = [s for s in meaningful if third_w <= (s["left"] or 0) + (s["width"] or 0) / 2 < 2 * third_w]
        right_shapes = [s for s in meaningful if (s["left"] or 0) + (s["width"] or 0) / 2 >= 2 * third_w]

        # Need content in all three areas (left=intro, centre+right=columns)
        left_text = sum(len(s["text"]) for s in left_shapes)
        centre_text = sum(len(s["text"]) for s in centre_shapes)
        right_text = sum(len(s["text"]) for s in right_shapes)

        if left_text >= 50 and centre_text >= 20 and right_text >= 20:
            return 0.52

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, intro text, and two content columns.
        """
        shapes = extract_shapes_with_text(slide)

        result = {
            "title": "",
            "intro": "",
            "col_centre": "",
            "col_right": "",
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
                elif idx in (17,):
                    if not footer_text:
                        footer_text = s["text"]
                elif idx in (18,):
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

        # Split remaining body into three columns
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

        # Sort each by vertical position
        left_texts.sort(key=lambda s: (s["top"] or 0))
        centre_texts.sort(key=lambda s: (s["top"] or 0))
        right_texts.sort(key=lambda s: (s["top"] or 0))

        result["intro"] = "\n".join(s["text"] for s in left_texts)
        result["col_centre"] = "\n".join(s["text"] for s in centre_texts)
        result["col_right"] = "\n".join(s["text"] for s in right_texts)

        return result

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """Fill Intro + Two Content placeholders. NEVER set font properties."""
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        if content.get("intro") and self.PH_INTRO in placeholders:
            placeholders[self.PH_INTRO].text = content["intro"]

        if content.get("col_centre") and self.PH_COL_CENTRE in placeholders:
            placeholders[self.PH_COL_CENTRE].text = content["col_centre"]

        if content.get("col_right") and self.PH_COL_RIGHT in placeholders:
            placeholders[self.PH_COL_RIGHT].text = content["col_right"]

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

    def _is_footer_text(self, text: str) -> bool:
        return any(re.search(p, text) for p in self.FOOTER_PATTERNS)

    def get_placeholder_map(self) -> dict:
        return {
            "title": self.PH_TITLE,
            "intro": self.PH_INTRO,
            "col_centre": self.PH_COL_CENTRE,
            "col_right": self.PH_COL_RIGHT,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
