"""
Text with Image Handler — Content slides with text and an image side by side.

Template layouts (image on right):
    idx 19 — Text with Image Two Thirds (image 2/3, text 1/3)
    idx 20 — Text with Image Half (50/50 split)
    idx 21 — Text with Image One Third (image 1/3, text 2/3)

Dynamic layout selection based on text-to-image ratio:
    - Lots of text + small image → layout 21 (image 1/3)
    - Moderate text + moderate image → layout 20 (half)
    - Minimal text + large image → layout 19 (image 2/3)

Placeholder index mapping varies by layout:
    Layout 19: title=0, subtitle=31, content=10, picture=34, footer=33, slidenum=24
    Layout 20: title=0, subtitle=31, content=10, picture=32, footer=33, slidenum=24
    Layout 21: title=0, subtitle=31, content=10, picture=34, footer=33, slidenum=24

Detection: slide has at least one embedded image AND meaningful body text.
Slides with an image but very little text should go to Title Only instead.

CRITICAL RULE: Never set font properties on placeholders.
"""

import io
import re
from handlers.base import SlideHandler
from utils.extractor import extract_text_elements, extract_images, extract_shapes_with_text


# Layout configs: (layout_index, picture_placeholder_idx)
LAYOUT_TWO_THIRDS = (19, 34)   # Image takes 2/3
LAYOUT_HALF = (20, 32)         # 50/50
LAYOUT_ONE_THIRD = (21, 34)    # Image takes 1/3

# Thresholds for text length to decide layout
# These are total body text character counts (excluding title/subtitle)
TEXT_HEAVY_THRESHOLD = 200     # Above this → image 1/3 (text needs space)
TEXT_LIGHT_THRESHOLD = 60      # Below this → image 2/3 (image is the star)


class TextImageHandler(SlideHandler):

    name = "Text with Image"
    description = "Content slide with text and an image side by side"
    layout_name = "Text with Image Half"
    layout_index = 20  # Default, overridden by get_layout_index()

    # Common placeholders (same across all three variants)
    PH_TITLE = 0
    PH_SUBTITLE = 31
    PH_CONTENT = 10
    PH_FOOTER = 33
    PH_SLIDE_NUM = 24

    # Noise to filter
    PLACEHOLDER_NOISE = [
        "[click", "[add", "[your", "[insert", "[subtitle",
        "click to add", "click icon", "click to edit",
        "\u00a9 the university", "this content is protected",
        "cricos code",
    ]

    # Minimum body text to qualify as Text with Image (vs Title Only)
    MIN_BODY_TEXT = 30

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect slides that have both meaningful body text AND at least one image.

        Key distinction from Title Only:
        - Text with Image: image + substantial body text (>30 chars)
        - Title Only: image + title only, minimal/no body text
        """
        if slide_index == 0:
            return 0.0

        images = extract_images(slide)
        if not images:
            return 0.0

        real_images = [img for img in images if img.get("blob")]
        if not real_images:
            return 0.0

        # Filter out images that are NOT content images.
        # Slide dimensions: 12192000 x 6858000 EMU.
        SLIDE_W = 12192000
        SLIDE_H = 6858000

        content_images = []
        for img in real_images:
            w = img.get("width") or 0
            h = img.get("height") or 0
            w_pct = w / SLIDE_W if SLIDE_W else 0
            h_pct = h / SLIDE_H if SLIDE_H else 0

            # Skip very large images (background/decorative)
            if w_pct > 0.7 and h_pct > 0.5:
                continue

            # Skip very small images (logos, icons, branding strips)
            # A content image should be at least 15% of slide width
            # AND at least 15% of slide height
            if w_pct < 0.15 or h_pct < 0.15:
                continue

            content_images.append(img)

        if not content_images:
            return 0.0  # Only logos/decorative images → not Text with Image

        texts = extract_text_elements(slide)
        if not texts:
            return 0.0

        # Filter noise and captions
        meaningful = [
            t for t in texts
            if not any(p in t["text"].lower() for p in self.PLACEHOLDER_NOISE)
            and len(t["text"].strip()) > 1
            and not self._is_image_caption(t["text"])
        ]

        if not meaningful:
            return 0.0

        # Separate title-like text from body text
        # Title: shortest text at top, or largest font
        sorted_texts = sorted(meaningful, key=lambda t: t.get("top", 0) or 0)
        title_candidates = [t for t in sorted_texts if len(t["text"]) < 100]
        body_candidates = [t for t in meaningful if t not in title_candidates[:1]]

        # Calculate body text volume (excluding likely title)
        body_text = sum(len(t["text"]) for t in body_candidates)

        if body_text >= self.MIN_BODY_TEXT and len(real_images) >= 1:
            return 0.65
        elif body_text > 10:
            return 0.45  # Low body text — might be Title Only instead
        else:
            return 0.0  # No meaningful body text → Title Only territory

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, subtitle, body text, and primary image.
        Also determines the best layout variant based on text volume.
        """
        shapes = extract_shapes_with_text(slide)
        images = extract_images(slide)

        result = {
            "title": "",
            "subtitle": "",
            "content": "",
            "footer": "",
            "image_blob": None,
            "image_content_type": None,
            "_layout_variant": LAYOUT_HALF,  # Default: 50/50
        }

        # --- Extract primary image (largest by area) ---
        real_images = [img for img in images if img.get("blob")]
        if real_images:
            primary = max(real_images, key=lambda i: (i["width"] or 0) * (i["height"] or 0))
            result["image_blob"] = primary["blob"]
            result["image_content_type"] = primary.get("content_type", "image/png")

            # Store image dimensions for layout selection
            result["_image_width"] = primary.get("width", 0)
            result["_image_height"] = primary.get("height", 0)

        # --- Extract text content ---
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
            if self._is_image_caption(text):
                continue
            filtered.append(s)

        if not filtered:
            return result

        # Separate by role using placeholder indices
        title_shape = None
        body_shapes = []
        footer_text = ""

        for s in filtered:
            if s["is_placeholder"]:
                idx = s["shape_idx"]
                if idx in (0, 3, 15):
                    title_shape = s
                elif idx in (1, 10):
                    body_shapes.append(s)
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

            if not title_shape:
                sorted_by_top = sorted(body_shapes, key=lambda s: s["top"])
                for s in sorted_by_top:
                    if len(s["text"]) < 120:
                        title_shape = s
                        body_shapes = [bs for bs in body_shapes if bs is not s]
                        break

        if title_shape:
            result["title"] = title_shape["text"]

        # Subtitle detection
        body_shapes.sort(key=lambda s: (s["top"], s["left"]))
        if body_shapes and title_shape:
            candidate = body_shapes[0]
            is_short = len(candidate["text"]) < 100
            is_near_top = candidate["top"] < (title_shape["top"] + title_shape["height"] * 2)
            has_more_content = len(body_shapes) > 1
            if is_short and is_near_top and has_more_content:
                result["subtitle"] = candidate["text"]
                body_shapes = body_shapes[1:]

        # Body content
        if body_shapes:
            result["content"] = "\n".join(s["text"] for s in body_shapes)

        # --- Dynamic layout selection ---
        body_length = len(result["content"])
        image_area = result.get("_image_width", 0) * result.get("_image_height", 0)
        slide_area = 12192000 * 6858000  # Standard 16:9

        # Image fraction of slide area
        image_fraction = image_area / slide_area if slide_area > 0 else 0

        if body_length >= TEXT_HEAVY_THRESHOLD:
            # Lots of text → image gets 1/3
            result["_layout_variant"] = LAYOUT_ONE_THIRD
        elif body_length <= TEXT_LIGHT_THRESHOLD or image_fraction > 0.35:
            # Light text or large image → image gets 2/3
            result["_layout_variant"] = LAYOUT_TWO_THIRDS
        else:
            # Moderate → 50/50
            result["_layout_variant"] = LAYOUT_HALF

        return result

    def _is_image_caption(self, text: str) -> bool:
        """Check if text is an image license/caption that should be filtered."""
        caption_patterns = [
            r"(?i)image\s+licensed\s+through",
            r"(?i)adobe\s+stock[:\s]+\d",
            r"(?i)shutterstock[:\s]+\d",
            r"(?i)getty\s+images",
            r"(?i)\u00a9\s*\d{4}",
            r"(?i)source:\s*(http|www)",
            r"(?i)this\s+photo\s+by",
            r"(?i)licensed\s+under\s+cc",
            r"(?i)creative\s+commons",
        ]
        return any(re.search(p, text) for p in caption_patterns)

    def _is_footer_text(self, text: str) -> bool:
        """Check if text looks like a footer/programme name."""
        footer_patterns = [
            r"(?i)cricos", r"(?i)hbis\s+innovation",
            r"(?i)executive\s+education",
            r"(?i)presentation\s+title",
        ]
        return any(re.search(p, text) for p in footer_patterns)

    # --- Dynamic layout ---

    def get_layout_index(self, content: dict) -> int:
        """Return the appropriate layout index based on content analysis."""
        variant = content.get("_layout_variant", LAYOUT_HALF)
        return variant[0]  # First element is layout_index

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill text placeholders and insert image into picture placeholder.
        The picture placeholder index varies by layout variant.
        NEVER set font properties.
        """
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        # Title
        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        # Subtitle
        if content.get("subtitle") and self.PH_SUBTITLE in placeholders:
            placeholders[self.PH_SUBTITLE].text = content["subtitle"]

        # Content body
        if content.get("content") and self.PH_CONTENT in placeholders:
            placeholders[self.PH_CONTENT].text = content["content"]

        # Picture — the placeholder index depends on which layout was used
        if content.get("image_blob"):
            variant = content.get("_layout_variant", LAYOUT_HALF)
            picture_idx = variant[1]  # Second element is picture PH index
            if picture_idx in placeholders:
                image_stream = io.BytesIO(content["image_blob"])
                placeholders[picture_idx].insert_picture(image_stream)

    def get_placeholder_map(self) -> dict:
        return {
            "title": self.PH_TITLE,
            "subtitle": self.PH_SUBTITLE,
            "content": self.PH_CONTENT,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
