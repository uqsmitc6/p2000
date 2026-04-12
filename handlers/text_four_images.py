"""
Text with 4 Images Handler — Text content left with 2×2 image grid right.

Template layout 25 "Text with 4 Images":
    idx  0 — Title (left, w=3.74in, h=0.51in)
    idx 32 — Subtitle (left, w=3.74in, h=0.55in)
    idx 10 — Body content (left, w=3.74in, h=4.51in)
    idx 12 — Picture top-left (4.25in × 3.75in)
    idx 16 — Picture bottom-left (4.25in × 3.75in)
    idx 17 — Picture top-right (4.29in × 3.75in)
    idx 18 — Picture bottom-right (4.29in × 3.75in)
    idx 20 — Footer

Detection: slide has text content on one side with multiple (2-4)
images arranged in a grid. Distinguished from Image Collage by having
substantial body text alongside the images. Distinguished from
Text with Image by having 2+ images.

Content to look for:
    - Product feature slides with screenshots
    - Photo evidence with descriptive text
    - Case study slides with multiple supporting images
    - Before/after comparisons with explanation

CRITICAL RULE: Never set font properties on placeholders.
"""

import io
import re
from handlers.base import SlideHandler
from utils.extractor import extract_shapes_with_text, extract_images


class TextFourImagesHandler(SlideHandler):

    name = "Text with 4 Images"
    description = "Text content left with 2×2 image grid right"
    layout_name = "Text with 4 Images"
    layout_index = 25

    PH_TITLE = 0
    PH_SUBTITLE = 32
    PH_BODY = 10
    # Picture placeholders: top-left, bottom-left, top-right, bottom-right
    PH_PICTURES = [12, 16, 17, 18]
    PH_FOOTER = 20

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
    SLIDE_MID = SLIDE_WIDTH // 2

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect slides with text + multiple images (2-4).

        Content patterns:
        - Source layout name contains "4 image" or "four image"
        - 2+ real content images on one side of the slide
        - Substantial body text (100+ chars) on the opposite side
        """
        if slide_index == 0:
            return 0.0

        layout_name = slide.slide_layout.name.lower()
        if "4 image" in layout_name or "four image" in layout_name:
            return 0.75

        images = extract_images(slide)
        if not images:
            return 0.0

        # Filter to real content images (not tiny logos or full-bleed backgrounds)
        real_images = [
            img for img in images
            if img.get("width") and img.get("height")
            and img["width"] > self.SLIDE_WIDTH * 0.1
            and img["height"] > self.SLIDE_HEIGHT * 0.1
            and not (img["width"] > self.SLIDE_WIDTH * 0.8
                     and img["height"] > self.SLIDE_HEIGHT * 0.8)
        ]

        if len(real_images) < 2:
            return 0.0

        # Check for substantial text
        shapes = extract_shapes_with_text(slide)
        meaningful = self._get_meaningful_shapes(shapes)
        total_text = sum(len(s["text"]) for s in meaningful)

        if total_text < 100:
            return 0.0  # Too little text — more likely Image Collage

        # Check that images cluster on one side and text on the other
        image_centres = [
            (img.get("left") or 0) + (img.get("width") or 0) / 2
            for img in real_images
        ]
        images_right = sum(1 for c in image_centres if c > self.SLIDE_MID)
        images_left = len(image_centres) - images_right

        # Text shapes — find side with most text
        text_left = sum(
            len(s["text"]) for s in meaningful
            if (s["left"] or 0) + (s["width"] or 0) / 2 < self.SLIDE_MID
        )
        text_right = sum(
            len(s["text"]) for s in meaningful
            if (s["left"] or 0) + (s["width"] or 0) / 2 >= self.SLIDE_MID
        )

        # Images on right, text on left (or vice versa)
        if (images_right >= 2 and text_left >= 100) or \
           (images_left >= 2 and text_right >= 100):
            if len(real_images) >= 4:
                return 0.58
            elif len(real_images) >= 2:
                return 0.50
        elif len(real_images) >= 2 and total_text >= 100:
            # Images and text mixed but enough of both
            return 0.42

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, subtitle, body text, and up to 4 images
        (sorted by area, largest first).
        """
        shapes = extract_shapes_with_text(slide)
        images = extract_images(slide)

        result = {
            "title": "",
            "subtitle": "",
            "body": "",
            "footer": "",
            "images": [],
        }

        # Extract images sorted by area (largest first)
        images_with_blobs = [img for img in images if img.get("blob")]
        images_with_blobs.sort(
            key=lambda i: (i.get("width") or 0) * (i.get("height") or 0),
            reverse=True,
        )
        for img in images_with_blobs[:4]:
            result["images"].append({
                "blob": img["blob"],
                "width": img.get("width"),
                "height": img.get("height"),
                "content_type": img.get("content_type") or "image/png",
            })

        if not shapes:
            return result

        meaningful = self._get_meaningful_shapes(shapes)

        # Find title, subtitle, body, footer
        title_shape = None
        body_parts = []
        footer_text = ""

        for s in meaningful:
            if s["is_placeholder"]:
                idx = s["shape_idx"]
                if idx in (0, 3, 15):
                    title_shape = s
                elif idx in (17, 20):
                    if not footer_text:
                        footer_text = s["text"]
                elif idx in (18,):
                    pass  # Slide number
                else:
                    body_parts.append(s)
            elif self._is_footer_text(s["text"]):
                if not footer_text:
                    footer_text = s["text"]
            else:
                body_parts.append(s)

        result["footer"] = footer_text

        # Fallback title
        if not title_shape and body_parts:
            with_font = [s for s in body_parts if s["font_size"]]
            if with_font:
                candidate = max(with_font, key=lambda s: s["font_size"])
                if len(candidate["text"]) < 120:
                    title_shape = candidate
                    body_parts = [s for s in body_parts if s is not title_shape]
            elif body_parts:
                candidate = body_parts[0]
                if len(candidate["text"]) < 80:
                    title_shape = candidate
                    body_parts = body_parts[1:]

        if title_shape:
            result["title"] = title_shape["text"]

        # Subtitle — short text near top
        body_parts.sort(key=lambda s: (s["top"] or 0, s["left"] or 0))
        if body_parts and title_shape:
            candidate = body_parts[0]
            if len(candidate["text"]) < 100 and len(body_parts) > 1:
                result["subtitle"] = candidate["text"]
                body_parts = body_parts[1:]

        # Remaining text becomes body
        result["body"] = "\n".join(s["text"] for s in body_parts)

        return result

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill title, subtitle, body, and insert images into picture
        placeholders. NEVER set font properties.
        """
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        if content.get("subtitle") and self.PH_SUBTITLE in placeholders:
            placeholders[self.PH_SUBTITLE].text = content["subtitle"]

        if content.get("body") and self.PH_BODY in placeholders:
            placeholders[self.PH_BODY].text = content["body"]

        # Insert images into picture placeholders
        images = content.get("images", [])
        for i, ph_idx in enumerate(self.PH_PICTURES):
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

    def _is_footer_text(self, text: str) -> bool:
        return any(re.search(p, text) for p in self.FOOTER_PATTERNS)

    def get_placeholder_map(self) -> dict:
        return {
            "title": self.PH_TITLE,
            "subtitle": self.PH_SUBTITLE,
            "body": self.PH_BODY,
            "pictures": self.PH_PICTURES,
            "footer": self.PH_FOOTER,
        }
