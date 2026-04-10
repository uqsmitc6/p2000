"""
Split Content Handler — Asymmetric two-column layouts (1/3 + 2/3).

Template layouts:
    Layout 15 "One Third Two Third Title and Content":
        idx  0 — Title (full width)
        idx 31 — Subtitle (full width)
        idx 10 — Left content (narrow, ~3.9")
        idx 35 — Right content (wide, ~8.0")
        idx 17 — Footer
        idx 18 — Slide number

    Layout 16 "Two Third One Third Title and Content":
        idx  0 — Title (full width)
        idx 31 — Subtitle (full width)
        idx 10 — Left content (wide, ~8.1")
        idx 34 — Right content (narrow, ~3.9")
        idx 17 — Footer
        idx 18 — Slide number

Detection: similar to Two Content but with unequal column widths.
The handler automatically selects layout 15 or 16 based on which
column has more content (heavier content goes in the 2/3 column).

When source content is ambiguous, defaults to layout 15 (narrow left
intro/key point, wide right for detailed content) — the most common
pattern in academic presentations.

CRITICAL RULE: Never set font properties on placeholders.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_text_elements, extract_images, extract_shapes_with_text


class SplitContentHandler(SlideHandler):

    name = "Split Content"
    description = "Asymmetric two-column layout (1/3 + 2/3 split)"
    layout_name = "One Third Two Third Title and Content"
    layout_index = 15  # Default: narrow left, wide right

    # Layout 15 placeholders (1/3 left, 2/3 right)
    PH_TITLE = 0
    PH_SUBTITLE = 31
    PH_LEFT_15 = 10   # Narrow column (1/3)
    PH_RIGHT_15 = 35   # Wide column (2/3)

    # Layout 16 placeholders (2/3 left, 1/3 right)
    PH_LEFT_16 = 10   # Wide column (2/3)
    PH_RIGHT_16 = 34   # Narrow column (1/3)

    PH_FOOTER = 17
    PH_SLIDE_NUM = 18

    PLACEHOLDER_NOISE = [
        "[click", "[add", "[your", "[insert", "[subtitle",
        "click to add", "click icon", "click to edit",
        "\u00a9 the university", "this content is protected",
        "cricos code",
    ]

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect slides with asymmetric two-column content.

        Signals:
        - Source layout name mentions "third" or "split"
        - Two content areas with significantly different widths
        - One area is substantially larger/more content than the other
        """
        if slide_index == 0:
            return 0.0

        # Check source layout name
        layout_name = slide.slide_layout.name.lower()
        if ("third" in layout_name or "split" in layout_name) and "content" in layout_name:
            return 0.75

        shapes = extract_shapes_with_text(slide)
        if not shapes:
            return 0.0

        # Filter noise
        meaningful = [
            s for s in shapes
            if s["text"].strip()
            and not any(p in s["text"].lower() for p in self.PLACEHOLDER_NOISE)
            and len(s["text"].strip()) > 5
        ]

        if len(meaningful) < 3:
            return 0.0

        # Find body shapes (exclude title and title-like shapes)
        # First, identify the likely title
        title_shape = None
        for s in meaningful:
            if s["is_placeholder"] and s["shape_idx"] in (0, 3, 15):
                title_shape = s
                break
        if not title_shape:
            sorted_by_top = sorted(meaningful, key=lambda s: s["top"] or 0)
            for s in sorted_by_top:
                # Title is typically short and at the top
                if len(s["text"]) < 120:
                    # Check font size if available, otherwise use position + length
                    if s.get("font_size") and s["font_size"] >= 20:
                        title_shape = s
                        break
                    elif not s.get("font_size") and len(s["text"]) < 80:
                        title_shape = s
                        break

        body_candidates = [s for s in meaningful if s is not title_shape]

        if len(body_candidates) < 2:
            return 0.0

        # Check for side-by-side shapes with unequal widths
        slide_width = 12192000
        half_width = slide_width / 2

        for i, a in enumerate(body_candidates):
            for b in body_candidates[i + 1:]:
                if a["left"] is None or b["left"] is None:
                    continue
                if a["top"] is None or b["top"] is None:
                    continue

                # Similar vertical position
                top_diff = abs(a["top"] - b["top"])
                if top_diff > 1000000:
                    continue

                # Different horizontal positions
                left_diff = abs(a["left"] - b["left"])
                if left_diff < half_width * 0.3:
                    continue

                # Check for unequal widths (at least 40% difference)
                if a["width"] and b["width"]:
                    width_ratio = min(a["width"], b["width"]) / max(a["width"], b["width"])
                    if width_ratio < 0.6:  # One column is less than 60% the width of the other
                        return 0.58  # Beat Two Content's 0.55

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, subtitle, and two columns of content.
        Also determines which layout to use based on content volume.

        The '_use_layout_16' flag in the result tells fill_slide
        to use layout 16 (2/3 left, 1/3 right) instead of the default
        layout 15 (1/3 left, 2/3 right).
        """
        shapes = extract_shapes_with_text(slide)

        result = {
            "title": "",
            "subtitle": "",
            "narrow_content": "",   # Goes in the 1/3 column
            "wide_content": "",     # Goes in the 2/3 column
            "footer": "",
            "_use_layout_16": False,  # If True, wide content on LEFT
        }

        if not shapes:
            return result

        # Filter noise
        filtered = []
        for s in shapes:
            text = s["text"].strip()
            if not text:
                continue
            if any(p in text.lower() for p in self.PLACEHOLDER_NOISE):
                continue
            if re.match(r"^\d{1,3}$", text) and len(text) <= 3:
                continue
            filtered.append(s)

        if not filtered:
            return result

        # Separate roles
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
        body_shapes.sort(key=lambda s: (s["top"], s["left"]))
        if body_shapes and title_shape:
            candidate = body_shapes[0]
            is_short = len(candidate["text"]) < 100
            is_near_top = candidate["top"] < (title_shape["top"] + title_shape["height"] * 2)
            has_more = len(body_shapes) > 1
            if is_short and is_near_top and has_more:
                result["subtitle"] = candidate["text"]
                body_shapes = body_shapes[1:]

        # Split into left/right columns
        if body_shapes:
            slide_midpoint = 6096000

            left_shapes = []
            right_shapes = []

            for s in body_shapes:
                shape_centre = (s["left"] or 0) + (s["width"] or 0) / 2
                if shape_centre < slide_midpoint:
                    left_shapes.append(s)
                else:
                    right_shapes.append(s)

            left_shapes.sort(key=lambda s: (s["top"], s["left"]))
            right_shapes.sort(key=lambda s: (s["top"], s["left"]))

            left_text = "\n".join(s["text"] for s in left_shapes)
            right_text = "\n".join(s["text"] for s in right_shapes)

            # Determine which side has more content → that goes in 2/3
            left_len = len(left_text)
            right_len = len(right_text)

            if left_len > right_len * 1.5:
                # Left has substantially more → use layout 16 (2/3 left)
                result["wide_content"] = left_text
                result["narrow_content"] = right_text
                result["_use_layout_16"] = True
            else:
                # Right has more or roughly equal → layout 15 (2/3 right)
                result["narrow_content"] = left_text
                result["wide_content"] = right_text
                result["_use_layout_16"] = False

            # Handle case where all content is on one side
            if not result["narrow_content"] and result["wide_content"]:
                # Split the content: first item → narrow, rest → wide
                lines = result["wide_content"].split("\n")
                if len(lines) > 1:
                    result["narrow_content"] = lines[0]
                    result["wide_content"] = "\n".join(lines[1:])

        return result

    def _is_footer_text(self, text: str) -> bool:
        footer_patterns = [
            r"(?i)cricos", r"(?i)hbis\s+innovation",
            r"(?i)executive\s+education",
            r"(?i)presentation\s+title",
        ]
        return any(re.search(p, text) for p in footer_patterns)

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill split content placeholders.

        Uses layout 15 (1/3 left, 2/3 right) by default.
        If content indicates layout 16, the caller should have created
        the slide with layout_index=16 instead.

        Since we can't change the layout after slide creation, we fill
        whichever placeholders exist on the slide.
        """
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        # Title
        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        # Subtitle
        if content.get("subtitle") and self.PH_SUBTITLE in placeholders:
            placeholders[self.PH_SUBTITLE].text = content["subtitle"]

        # Content columns — fill whichever placeholders exist
        # Layout 15: idx 10 = narrow (1/3), idx 35 = wide (2/3)
        # Layout 16: idx 10 = wide (2/3), idx 34 = narrow (1/3)
        use_16 = content.get("_use_layout_16", False)

        if use_16:
            # Layout 16: wide on left (idx 10), narrow on right (idx 34)
            if content.get("wide_content") and 10 in placeholders:
                placeholders[10].text = content["wide_content"]
            if content.get("narrow_content") and 34 in placeholders:
                placeholders[34].text = content["narrow_content"]
        else:
            # Layout 15: narrow on left (idx 10), wide on right (idx 35)
            if content.get("narrow_content") and 10 in placeholders:
                placeholders[10].text = content["narrow_content"]
            if content.get("wide_content") and 35 in placeholders:
                placeholders[35].text = content["wide_content"]

    def get_placeholder_map(self) -> dict:
        """Return placeholder map for layout 15 (default)."""
        return {
            "title": self.PH_TITLE,
            "subtitle": self.PH_SUBTITLE,
            "left": self.PH_LEFT_15,
            "right": self.PH_RIGHT_15,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }

    def get_layout_index(self, content: dict) -> int:
        """
        Return the appropriate layout index based on content analysis.
        Called by the converter to determine which template layout to use.
        """
        if content.get("_use_layout_16", False):
            return 16
        return 15
