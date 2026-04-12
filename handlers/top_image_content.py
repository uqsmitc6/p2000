"""
Top Image + Content Handler — Image with content columns above.

Template layouts:
    idx  9 — Top image + two content layout
    idx 10 — Top image + three content layout

Layout 9 "Top image + two content":
    idx  0 — Title (left, top=1.39in, w=3.81in)
    idx 37 — Centre content column (top=1.39in, left=4.79in, w=3.81in)
    idx 36 — Right content column (top=1.39in, left=8.89in, w=3.92in)
    idx 38 — Picture placeholder (bottom half, full-width, h=3.67in)

Layout 10 "Top image + three content":
    idx 35 — Left content column (top=1.39in, left=0.52in, w=3.92in)
    idx 37 — Centre content column (top=1.39in, left=4.79in, w=3.81in)
    idx 36 — Right content column (top=1.39in, left=8.89in, w=3.92in)
    idx 38 — Picture placeholder (bottom half, full-width, h=3.67in)

Note: Layout 9 has a title in the left column position; Layout 10
replaces the title with a third content column.

Detection: primarily by source layout name. Heuristic: image + multiple
text columns where the image occupies the bottom portion of the slide.

CRITICAL RULE: Never set font properties on placeholders.
"""

import io
import re
from handlers.base import SlideHandler
from utils.extractor import extract_shapes_with_text, extract_images


# Layout configs: (layout_index, has_title, content_ph_indices)
LAYOUT_TWO_CONTENT = (9, True, [37, 36])      # Title + 2 columns + image
LAYOUT_THREE_CONTENT = (10, False, [35, 37, 36])  # 3 columns + image (no title)


class TopImageContentHandler(SlideHandler):

    name = "Top Image + Content"
    description = "Image at bottom with content columns above"
    layout_name = "Top image + two content layout"
    layout_index = 9  # Default, overridden by get_layout_index()

    PH_TITLE = 0       # Only in layout 9
    PH_PICTURE = 38    # Same in both layouts

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
        Detect slides with image + content columns layout.
        """
        if slide_index == 0:
            return 0.0

        layout_name = slide.slide_layout.name.lower()
        if "top image" in layout_name and "content" in layout_name:
            return 0.75

        # Heuristic: large image in bottom half + multiple text areas in top half
        images = extract_images(slide)
        shapes = extract_shapes_with_text(slide)
        meaningful = self._get_meaningful_shapes(shapes)

        if not images or len(meaningful) < 2:
            return 0.0

        slide_height = 6858000
        mid_height = slide_height / 2

        # Look for a large image in the bottom half
        bottom_images = [
            img for img in images
            if img.get("top") and img["top"] > mid_height * 0.7
            and img.get("width") and img["width"] > self.SLIDE_WIDTH * 0.5
        ]

        if not bottom_images:
            return 0.0

        # Count text elements in the top half
        top_texts = [
            s for s in meaningful
            if (s["top"] or 0) < mid_height
        ]

        if len(top_texts) >= 2:
            return 0.52

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, content columns, and image.
        Determines two-column vs three-column variant.
        """
        shapes = extract_shapes_with_text(slide)
        images = extract_images(slide)

        result = {
            "title": "",
            "columns": [],      # List of column text strings
            "image_blob": None,
            "image_content_type": "image/png",
            "_variant": LAYOUT_TWO_CONTENT,
        }

        # Extract primary image
        images_with_blobs = [img for img in images if img.get("blob")]
        if images_with_blobs:
            primary = max(
                images_with_blobs,
                key=lambda i: (i.get("width") or 0) * (i.get("height") or 0),
            )
            result["image_blob"] = primary["blob"]
            result["image_content_type"] = primary.get("content_type") or "image/png"

        if not shapes:
            return result

        meaningful = self._get_meaningful_shapes(shapes)
        if not meaningful:
            return result

        # Find title
        title_shape = None
        body_shapes = []

        for s in meaningful:
            if s["is_placeholder"] and s["shape_idx"] in (0, 3, 15):
                title_shape = s
            elif self._is_footer_text(s["text"]):
                pass
            else:
                body_shapes.append(s)

        # Fallback title
        if not title_shape and body_shapes:
            sorted_by_top = sorted(body_shapes, key=lambda s: (s["top"] or 0, s["left"] or 0))
            for s in sorted_by_top:
                if len(s["text"]) < 80:
                    title_shape = s
                    body_shapes = [bs for bs in body_shapes if bs is not s]
                    break

        if title_shape:
            result["title"] = title_shape["text"]

        # Split body into columns based on horizontal position
        third_w = self.SLIDE_WIDTH / 3

        left_texts = []
        centre_texts = []
        right_texts = []

        for s in body_shapes:
            shape_centre = (s["left"] or 0) + (s["width"] or 0) / 2
            if shape_centre < third_w:
                left_texts.append(s["text"])
            elif shape_centre < 2 * third_w:
                centre_texts.append(s["text"])
            else:
                right_texts.append(s["text"])

        # Determine variant
        non_empty_cols = sum(1 for col in [left_texts, centre_texts, right_texts] if col)

        if non_empty_cols >= 3 or (not title_shape and non_empty_cols >= 2):
            result["_variant"] = LAYOUT_THREE_CONTENT
            result["columns"] = [
                "\n".join(left_texts),
                "\n".join(centre_texts),
                "\n".join(right_texts),
            ]
        else:
            result["_variant"] = LAYOUT_TWO_CONTENT
            # Two-column: centre and right
            result["columns"] = [
                "\n".join(centre_texts),
                "\n".join(right_texts),
            ]
            # If content only in left+one other, redistribute
            if left_texts and not result["title"]:
                result["title"] = "\n".join(left_texts)

        return result

    # --- Dynamic layout ---

    def get_layout_index(self, content: dict) -> int:
        variant = content.get("_variant", LAYOUT_TWO_CONTENT)
        return variant[0]

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill content columns and insert image.
        NEVER set font properties.
        """
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}
        variant = content.get("_variant", LAYOUT_TWO_CONTENT)
        content_indices = variant[2]
        has_title = variant[1]

        # Title (only layout 9)
        if has_title and content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        # Content columns
        columns = content.get("columns", [])
        for i, ph_idx in enumerate(content_indices):
            if i < len(columns) and columns[i] and ph_idx in placeholders:
                placeholders[ph_idx].text = columns[i]

        # Picture
        if content.get("image_blob") and self.PH_PICTURE in placeholders:
            image_stream = io.BytesIO(content["image_blob"])
            placeholders[self.PH_PICTURE].insert_picture(image_stream)

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
            "picture": self.PH_PICTURE,
            "columns": "varies by variant",
        }
