"""
Image Collage Handler — Multiple images in a collage/mosaic arrangement.

Template layout 41 "Image collage":
    idx  0 — Title (left, w=3.23in, h=1.89in)
    idx 38 — Caption/description text (left bottom, w=3.07in, h=0.54in)
    idx 39 — Picture 1 (small, 2.52in × 1.81in, top-centre)
    idx 14 — Picture 2 (tall, 2.52in × 3.58in, mid-centre)
    idx 21 — Picture 3 (tall, 2.52in × 3.58in, mid-right)
    idx 40 — Picture 4 (small, 2.52in × 1.81in, bottom-right)
    idx 41 — Picture 5 (large, 3.24in × 5.47in, far-right)
    idx 10 — Footer
    idx 11 — Slide number

Detection: slide has 3+ images without substantial body text. The
images are the star, not the text. Distinguished from Text with Image
by having multiple images and minimal text.

Content to look for:
    - Photo gallery slides
    - Project showcase with multiple images
    - Visual portfolio or mood board slides

CRITICAL RULE: Never set font properties on placeholders.
"""

import io
import re
from handlers.base import SlideHandler
from utils.extractor import extract_shapes_with_text, extract_images


class ImageCollageHandler(SlideHandler):

    name = "Image Collage"
    description = "Multiple images in a collage arrangement with title"
    layout_name = "Image collage"
    layout_index = 41

    PH_TITLE = 0
    PH_CAPTION = 38
    # Picture placeholders in order (largest to smallest for priority fill)
    PH_PICTURES = [41, 14, 21, 39, 40]
    PH_FOOTER = 10
    PH_SLIDE_NUM = 11

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
        Detect slides with multiple images (3+) and minimal text.

        Content patterns:
        - Source layout name contains "collage" or "gallery"
        - 3+ real content images (not logos/backgrounds)
        - Minimal body text (< 100 chars)
        """
        if slide_index == 0:
            return 0.0

        layout_name = slide.slide_layout.name.lower()
        if "collage" in layout_name or "gallery" in layout_name:
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
            and not (img["width"] > self.SLIDE_WIDTH * 0.8 and img["height"] > self.SLIDE_HEIGHT * 0.8)
        ]

        if len(real_images) < 3:
            return 0.0

        # Check text volume — collages have minimal text
        shapes = extract_shapes_with_text(slide)
        meaningful = self._get_meaningful_shapes(shapes)
        total_text = sum(len(s["text"]) for s in meaningful)

        if total_text > 200:
            return 0.0  # Too much text — this is a content slide

        if len(real_images) >= 4:
            return 0.60
        elif len(real_images) >= 3:
            return 0.52

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, caption, and up to 5 images (sorted by area, largest first).
        """
        shapes = extract_shapes_with_text(slide)
        images = extract_images(slide)

        result = {
            "title": "",
            "caption": "",
            "footer": "",
            "images": [],  # List of {"blob": bytes, "width": emu, "height": emu}
        }

        # Extract images sorted by area (largest first)
        images_with_blobs = [img for img in images if img.get("blob")]
        images_with_blobs.sort(
            key=lambda i: (i.get("width") or 0) * (i.get("height") or 0),
            reverse=True,
        )
        # Take up to 5 images
        for img in images_with_blobs[:5]:
            result["images"].append({
                "blob": img["blob"],
                "width": img.get("width"),
                "height": img.get("height"),
                "content_type": img.get("content_type") or "image/png",
            })

        if not shapes:
            return result

        meaningful = self._get_meaningful_shapes(shapes)

        # Find title and caption
        title_shape = None
        caption_text = ""

        for s in meaningful:
            if s["is_placeholder"] and s["shape_idx"] in (0, 3, 15):
                title_shape = s
            elif s["is_placeholder"] and s["shape_idx"] in (10,):
                pass  # Footer in this layout
            elif self._is_footer_text(s["text"]):
                pass
            elif not title_shape and len(s["text"]) < 80:
                title_shape = s
            else:
                if not caption_text:
                    caption_text = s["text"]

        if title_shape:
            result["title"] = title_shape["text"]
        result["caption"] = caption_text

        # Footer
        for s in meaningful:
            if s["is_placeholder"] and s["shape_idx"] in (10, 17):
                result["footer"] = s["text"]
                break

        return result

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill title, caption, and insert images into picture placeholders.
        Images are matched to placeholders in size order (largest image
        to largest placeholder).
        NEVER set font properties.
        """
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        if content.get("caption") and self.PH_CAPTION in placeholders:
            placeholders[self.PH_CAPTION].text = content["caption"]

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
            "caption": self.PH_CAPTION,
            "pictures": self.PH_PICTURES,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
