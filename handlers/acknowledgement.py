"""
Acknowledgement of Country Handler — Always inserted as slide 2.

Uses layout 20 "Text with Image Half":
    idx  0 — Title ("Acknowledgement of Country")
    idx 10 — Content (AoC text)
    idx 31 — Subtitle (not used)
    idx 32 — Picture (Brisbane River artwork)
    idx 33 — Footer
    idx 24 — Slide number

This handler is SPECIAL — it does not detect content from the source deck.
Instead, it is always auto-inserted at position 2 by the converter.
If an existing AoC slide is detected in the source, it is consumed (not
duplicated) but the output always uses the standard branded version.

Detection: looks for "acknowledgement of country", "traditional owners",
"traditional custodians", "elders past" in any slide.

CRITICAL RULE: Never set font properties on placeholders.
"""

import io
import os
import re
from handlers.base import SlideHandler
from utils.extractor import extract_text_elements


# Path to the Brisbane River artwork (bundled with the app)
AOC_IMAGE_PATH = os.path.join(
    os.path.dirname(os.path.dirname(__file__)),
    "templates",
    "aoc_brisbane_river.jpg",
)

# Standard AoC text
AOC_TITLE = "Acknowledgement of Country"

AOC_BODY = (
    "The University of Queensland (UQ) acknowledges the Traditional Owners "
    "and their custodianship of the lands on which we meet.\n\n"
    "We pay our respects to their Ancestors and their descendants, who "
    "continue cultural and spiritual connections to Country.\n\n"
    "We recognise their valuable contributions to Australian and global society.\n\n"
    "Sovereignty was never ceded, and this always was and always will be "
    "Aboriginal land."
)

AOC_CAPTION = (
    "The Brisbane River pattern from A Guidance Through Time "
    "by Casey Coolwell and Kyra Mancktelow."
)


class AcknowledgementHandler(SlideHandler):

    name = "Acknowledgement of Country"
    description = "Standard UQ Acknowledgement of Country slide"
    layout_name = "Text with Image Half"
    layout_index = 20

    PH_TITLE = 0
    PH_CONTENT = 10
    PH_SUBTITLE = 31
    PH_PICTURE = 32
    PH_FOOTER = 33
    PH_SLIDE_NUM = 24

    AOC_PATTERNS = [
        r"(?i)acknowledgement\s+of\s+country",
        r"(?i)traditional\s+(owners|custodians)",
        r"(?i)elders\s+past",
        r"(?i)custodianship\s+of\s+the\s+lands",
    ]

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect existing AoC slides in the source deck.
        Returns high confidence if AoC language found.
        """
        texts = extract_text_elements(slide)
        all_text = " ".join(t["text"] for t in texts)

        for pattern in self.AOC_PATTERNS:
            if re.search(pattern, all_text):
                return 0.95

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        AoC content is always the standard text — we don't extract from source.
        """
        return self._get_standard_content()

    @staticmethod
    def _get_standard_content() -> dict:
        """Return the standard AoC slide content."""
        return {
            "title": AOC_TITLE,
            "content": AOC_BODY,
            "footer": "",
            "image_path": AOC_IMAGE_PATH,
        }

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill AoC placeholders with standard content and Brisbane River artwork.
        NEVER set font properties.
        """
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        # Title
        if self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content.get("title", AOC_TITLE)

        # Body text
        if self.PH_CONTENT in placeholders:
            placeholders[self.PH_CONTENT].text = content.get("content", AOC_BODY)

        # Brisbane River artwork
        image_path = content.get("image_path", AOC_IMAGE_PATH)
        if self.PH_PICTURE in placeholders and image_path and os.path.exists(image_path):
            with open(image_path, "rb") as f:
                image_stream = io.BytesIO(f.read())
            placeholders[self.PH_PICTURE].insert_picture(image_stream)

    def get_placeholder_map(self) -> dict:
        return {
            "title": self.PH_TITLE,
            "content": self.PH_CONTENT,
            "subtitle": self.PH_SUBTITLE,
            "picture": self.PH_PICTURE,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
