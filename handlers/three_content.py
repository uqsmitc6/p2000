"""
Three Content Handler — Slides with three side-by-side content columns.

Template layout 8 "Three content layout":
    idx  0 — Title
    idx 31 — Subtitle (optional)
    idx 10 — Left content column   (left=0.52in, w=3.92in)
    idx 35 — Centre content column (left=4.79in, w=3.81in)
    idx 34 — Right content column  (left=8.89in, w=3.92in)
    idx 17 — Footer (programme name)
    idx 18 — Slide number

Detection: slide has three distinct text groups arranged side by side,
or the source layout name indicates three columns. Extends the Two
Content detection logic to look for a third column.

CRITICAL RULE: Never set font properties on placeholders.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_shapes_with_text


class ThreeContentHandler(SlideHandler):

    name = "Three Content"
    description = "Content slide with three side-by-side content columns"
    layout_name = "Three content layout"
    layout_index = 8

    PH_TITLE = 0
    PH_SUBTITLE = 31
    PH_LEFT = 10
    PH_CENTRE = 35
    PH_RIGHT = 34
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

    SLIDE_WIDTH = 12192000  # Standard 16:9 EMU

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect slides with three-column content.

        Heuristics:
        - Source layout name contains "three" and "content"/"column"
        - Three body-sized shapes at similar vertical positions
          with centres in left/centre/right thirds of the slide
        """
        if slide_index == 0:
            return 0.0

        # Check source layout name
        layout_name = slide.slide_layout.name.lower()
        if "three" in layout_name and ("content" in layout_name or "column" in layout_name):
            return 0.75

        shapes = extract_shapes_with_text(slide)
        if not shapes:
            return 0.0

        # Filter noise
        meaningful = self._get_meaningful_shapes(shapes)

        if len(meaningful) < 4:  # Need title + 3 body areas
            return 0.0

        # Find title
        title_shape = self._find_title(meaningful)
        body_candidates = [s for s in meaningful if s is not title_shape]

        if len(body_candidates) < 3:
            return 0.0

        # Check if at least three body shapes span left/centre/right thirds
        third_w = self.SLIDE_WIDTH / 3
        left_found = False
        centre_found = False
        right_found = False

        for s in body_candidates:
            if not s["left"] or not s["width"]:
                continue
            shape_centre = s["left"] + s["width"] / 2
            if shape_centre < third_w:
                left_found = True
            elif shape_centre < 2 * third_w:
                centre_found = True
            else:
                right_found = True

        if left_found and centre_found and right_found:
            # Check that the three column shapes have substantial text
            col_shapes = self._split_into_columns(body_candidates)
            min_len = min(
                sum(len(s["text"]) for s in col)
                for col in col_shapes if col
            ) if all(col_shapes) else 0
            if min_len >= 15:
                return 0.58
            return 0.40

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, subtitle, and three columns of content.
        """
        shapes = extract_shapes_with_text(slide)

        result = {
            "title": "",
            "subtitle": "",
            "left_content": "",
            "centre_content": "",
            "right_content": "",
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

        # Fallback title detection
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

        # Split into three columns
        if body_shapes:
            left, centre, right = self._split_into_columns(body_shapes)

            result["left_content"] = "\n".join(s["text"] for s in left)
            result["centre_content"] = "\n".join(s["text"] for s in centre)
            result["right_content"] = "\n".join(s["text"] for s in right)

            # If all content ended up in fewer than 3 columns, redistribute
            cols = [result["left_content"], result["centre_content"], result["right_content"]]
            non_empty = [c for c in cols if c]
            if len(non_empty) < 3 and body_shapes:
                # Fall back to even distribution by shape count
                all_texts = [s["text"] for s in body_shapes]
                n = len(all_texts)
                third = max(1, n // 3)
                result["left_content"] = "\n".join(all_texts[:third])
                result["centre_content"] = "\n".join(all_texts[third:2*third])
                result["right_content"] = "\n".join(all_texts[2*third:])

        return result

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """Fill Three Content placeholders. NEVER set font properties."""
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        if content.get("subtitle") and self.PH_SUBTITLE in placeholders:
            placeholders[self.PH_SUBTITLE].text = content["subtitle"]

        if content.get("left_content") and self.PH_LEFT in placeholders:
            placeholders[self.PH_LEFT].text = content["left_content"]

        if content.get("centre_content") and self.PH_CENTRE in placeholders:
            placeholders[self.PH_CENTRE].text = content["centre_content"]

        if content.get("right_content") and self.PH_RIGHT in placeholders:
            placeholders[self.PH_RIGHT].text = content["right_content"]

    # --- Helpers ---

    def _get_meaningful_shapes(self, shapes: list) -> list:
        """Filter out noise, footer text, and trivial shapes."""
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
        """Find the title shape by placeholder idx or font size."""
        for s in shapes:
            if s["is_placeholder"] and s["shape_idx"] in (0, 3, 15):
                return s
        # Fallback: topmost short text with large font
        sorted_by_top = sorted(shapes, key=lambda s: s["top"] or 0)
        for s in sorted_by_top:
            if len(s["text"]) < 120:
                if s.get("font_size") and s["font_size"] >= 20:
                    return s
                elif not s.get("font_size") and len(s["text"]) < 80:
                    return s
        return None

    def _split_into_columns(self, shapes: list) -> tuple:
        """Split shapes into left, centre, right based on horizontal position."""
        third_w = self.SLIDE_WIDTH / 3

        left = []
        centre = []
        right = []

        for s in shapes:
            shape_centre = (s["left"] or 0) + (s["width"] or 0) / 2
            if shape_centre < third_w:
                left.append(s)
            elif shape_centre < 2 * third_w:
                centre.append(s)
            else:
                right.append(s)

        # Sort each column by vertical position
        left.sort(key=lambda s: (s["top"] or 0, s["left"] or 0))
        centre.sort(key=lambda s: (s["top"] or 0, s["left"] or 0))
        right.sort(key=lambda s: (s["top"] or 0, s["left"] or 0))

        return left, centre, right

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
