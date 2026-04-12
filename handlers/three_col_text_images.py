"""
Three Column Text & Images Handler — Three text columns above three images.

Template layout 26 "Three Column Text & Images":
    idx  0 — Title (full width)
    idx 53 — Subtitle (full width, h=0.55in)
    idx 10 — Left text (3.92in × 1.44in)
    idx 35 — Centre text (3.81in × 1.44in)
    idx 34 — Right text (3.92in × 1.44in)
    idx 14 — Left picture (3.92in × 2.75in)
    idx 16 — Centre picture (3.81in × 2.75in)
    idx 18 — Right picture (3.93in × 2.75in)
    idx 23 — Footer
    idx 24 — Slide number

Detection: slide with three text areas and three images arranged in
matching columns. Distinguished from Three Content by having images
below each text column. Distinguished from Image Collage by having
structured text for each image.

Content to look for:
    - Product/service comparison with photos
    - Team member profiles (name + photo × 3)
    - Feature showcase with screenshots
    - Location/campus comparison slides

CRITICAL RULE: Never set font properties on placeholders.
"""

import io
import re
from handlers.base import SlideHandler
from utils.extractor import extract_shapes_with_text, extract_images


class ThreeColTextImagesHandler(SlideHandler):

    name = "Three Column Text & Images"
    description = "Three text columns with matching images below"
    layout_name = "Three Column Text & Images"
    layout_index = 26

    PH_TITLE = 0
    PH_SUBTITLE = 53
    PH_LEFT_TEXT = 10
    PH_CENTRE_TEXT = 35
    PH_RIGHT_TEXT = 34
    PH_LEFT_PIC = 14
    PH_CENTRE_PIC = 16
    PH_RIGHT_PIC = 18
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
    SLIDE_HEIGHT = 6858000

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect slides with three text+image column pairs.

        Content patterns:
        - Source layout name contains "three column" and "image"
        - 3 images + 3 text areas, each pair in a vertical column
        - Images positioned below text areas
        """
        if slide_index == 0:
            return 0.0

        layout_name = slide.slide_layout.name.lower()
        if "three column" in layout_name and "image" in layout_name:
            return 0.75

        images = extract_images(slide)
        shapes = extract_shapes_with_text(slide)

        if not images or len(images) < 3:
            return 0.0

        # Filter to real content images
        real_images = [
            img for img in images
            if img.get("width") and img.get("height")
            and img["width"] > self.SLIDE_WIDTH * 0.15
            and img["height"] > self.SLIDE_HEIGHT * 0.15
            and not (img["width"] > self.SLIDE_WIDTH * 0.8
                     and img["height"] > self.SLIDE_HEIGHT * 0.8)
        ]

        if len(real_images) < 3:
            return 0.0

        meaningful = self._get_meaningful_shapes(shapes)

        # Need title + at least 3 text areas
        title = self._find_title(meaningful)
        body = [s for s in meaningful if s is not title]

        if len(body) < 3:
            return 0.0

        # Check that images are in the bottom half and text in the top half
        mid_y = self.SLIDE_HEIGHT // 2
        images_below = sum(
            1 for img in real_images
            if (img.get("top") or 0) + (img.get("height") or 0) / 2 > mid_y
        )
        text_above = sum(
            1 for s in body
            if (s["top"] or 0) + (s["height"] or 0) / 2 < mid_y
        )

        if images_below >= 3 and text_above >= 3:
            return 0.55

        # More relaxed: 3 images + 3 text areas regardless of position
        if len(real_images) >= 3 and len(body) >= 3:
            total_text = sum(len(s["text"]) for s in body)
            if total_text >= 50:
                return 0.42

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, subtitle, three text columns, and three images.
        Images sorted by horizontal position (left to right).
        """
        shapes = extract_shapes_with_text(slide)
        images = extract_images(slide)

        result = {
            "title": "",
            "subtitle": "",
            "left_text": "",
            "centre_text": "",
            "right_text": "",
            "footer": "",
            "images": [],
        }

        # Extract images sorted by horizontal position (left to right)
        images_with_blobs = [img for img in images if img.get("blob")]
        images_with_blobs.sort(key=lambda i: i.get("left") or 0)
        for img in images_with_blobs[:3]:
            result["images"].append({
                "blob": img["blob"],
                "width": img.get("width"),
                "height": img.get("height"),
                "content_type": img.get("content_type") or "image/png",
            })

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
            if len(candidate["text"]) < 100 and len(body_shapes) > 1:
                near_top = (candidate["top"] or 0) < (title_shape["top"] or 0) + (title_shape["height"] or 0) * 2
                if near_top:
                    result["subtitle"] = candidate["text"]
                    body_shapes = body_shapes[1:]

        # Split into three columns
        third_w = self.SLIDE_WIDTH / 3
        left, centre, right = [], [], []

        for s in body_shapes:
            cx = (s["left"] or 0) + (s["width"] or 0) / 2
            if cx < third_w:
                left.append(s)
            elif cx < 2 * third_w:
                centre.append(s)
            else:
                right.append(s)

        for group in [left, centre, right]:
            group.sort(key=lambda s: (s["top"] or 0))

        result["left_text"] = "\n".join(s["text"] for s in left)
        result["centre_text"] = "\n".join(s["text"] for s in centre)
        result["right_text"] = "\n".join(s["text"] for s in right)

        return result

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill three text columns and three picture placeholders.
        NEVER set font properties.
        """
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        if content.get("subtitle") and self.PH_SUBTITLE in placeholders:
            placeholders[self.PH_SUBTITLE].text = content["subtitle"]

        if content.get("left_text") and self.PH_LEFT_TEXT in placeholders:
            placeholders[self.PH_LEFT_TEXT].text = content["left_text"]

        if content.get("centre_text") and self.PH_CENTRE_TEXT in placeholders:
            placeholders[self.PH_CENTRE_TEXT].text = content["centre_text"]

        if content.get("right_text") and self.PH_RIGHT_TEXT in placeholders:
            placeholders[self.PH_RIGHT_TEXT].text = content["right_text"]

        # Insert images into picture placeholders (left, centre, right)
        images = content.get("images", [])
        pic_phs = [self.PH_LEFT_PIC, self.PH_CENTRE_PIC, self.PH_RIGHT_PIC]
        for i, ph_idx in enumerate(pic_phs):
            if i >= len(images):
                break
            if ph_idx in placeholders:
                image_stream = io.BytesIO(images[i]["blob"])
                placeholders[ph_idx].insert_picture(image_stream)

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
            "left_text": self.PH_LEFT_TEXT,
            "centre_text": self.PH_CENTRE_TEXT,
            "right_text": self.PH_RIGHT_TEXT,
            "left_pic": self.PH_LEFT_PIC,
            "centre_pic": self.PH_CENTRE_PIC,
            "right_pic": self.PH_RIGHT_PIC,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
