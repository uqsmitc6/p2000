"""
Section Divider Handler — White background, section number top-left,
title below, optional description on right, footer + slide number.

Template placeholders (layout index 5, "Section Divider"):
    idx 11 — Section number/label (e.g. "01", "Session 1:", "Block 1")
    idx  0 — Section title (the topic name)
    idx 13 — Description text (right side, optional)
    idx 10 — Footer (programme name)
    idx 14 — Slide number

Detection heuristics:
    - Layout name contains "divider" or "section" → high confidence
    - Slide with very few text elements (≤3), short text, no bullets → medium
    - A number/label pattern + short title → medium-high
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_text_elements, extract_images, extract_shapes_with_text


class SectionDividerHandler(SlideHandler):

    name = "Section Divider"
    description = "Section break slide with number, title, and optional description"
    layout_name = "Section Divider"
    layout_index = 5

    # Placeholder indices in the template
    PH_SECTION_NUM = 11
    PH_TITLE = 0
    PH_DESCRIPTION = 13
    PH_FOOTER = 10
    PH_SLIDE_NUM = 14

    # --- Detection ---

    # Patterns that indicate a section number/label.
    # IMPORTANT: bare numbers like "13" are usually slide numbers, not
    # section numbers. Only zero-padded ("01") or labelled ("Session 1")
    # numbers count as section indicators.
    SECTION_NUM_PATTERNS = [
        r"^0\d{1,2}$",                         # "01", "02" — zero-padded only
        r"(?i)^session\s*\d",                   # "Session 1:", "Session 2"
        r"(?i)^block\s*\d",                     # "Block 1", "Block 2"
        r"(?i)^module\s*\d",                    # "Module 1"
        r"(?i)^part\s*\d",                      # "Part 1"
        r"(?i)^week\s*\d",                      # "Week 1"
        r"(?i)^day\s*\d",                       # "Day 1"
        r"(?i)^unit\s*\d",                      # "Unit 1"
        r"(?i)^topic\s*\d",                     # "Topic 1"
        r"(?i)^section\s*\d",                   # "Section 1"
    ]

    # Layout names that strongly indicate a section divider
    DIVIDER_LAYOUT_NAMES = [
        "section divider", "section header", "divider",
    ]

    # Placeholder/template noise to filter
    PLACEHOLDER_NOISE = [
        "[divider", "[section", "[click", "[add", "[your", "[insert",
        "click to add", "click icon", "\u00a9 the university",
        "lorem ipsum", "divider slide title goes here",
    ]

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect whether this slide is a section divider.

        IMPORTANT: Many source templates use layout names like "Section Header"
        for general content slides. We must check content density — if the slide
        has lots of text, bullet points, or citations, it's NOT a divider even
        if the layout name says so.
        """
        # Never classify first slide as divider (that's the cover)
        if slide_index == 0:
            return 0.0

        texts = extract_text_elements(slide)
        images = extract_images(slide)

        # Filter noise
        meaningful = [
            t for t in texts
            if not any(p in t["text"].lower() for p in self.PLACEHOLDER_NOISE)
            and len(t["text"].strip()) > 0
        ]

        # Filter out slide numbers and footer-like text (very small or at bottom)
        content_texts = [
            t for t in meaningful
            if (t["font_size"] is None or t["font_size"] >= 10)
        ]

        # --- Content density checks ---
        # Count total text length and bullet indicators
        total_text_len = sum(len(t["text"]) for t in content_texts)
        has_bullets = any(
            any(c in t["text"] for c in ("•", "–", "▪", "►", "◆", "■"))
            or t["text"].strip().startswith(("-", "*"))
            or re.match(r"^\d+[\.\)]\s", t["text"].strip())
            for t in content_texts
        )
        # Count how many text elements look like citation/reference patterns
        citation_count = sum(
            1 for t in content_texts
            if re.search(r"\(\d{4}\)", t["text"]) or re.search(r",\s*\d{4}", t["text"])
        )

        # A slide is "content-heavy" if it has lots of text, bullets, or citations
        is_content_heavy = (
            total_text_len > 200
            or len(content_texts) > 4
            or has_bullets
            or citation_count >= 2
        )

        # Heuristic: section number pattern present (zero-padded or labelled)
        has_section_num = any(
            any(re.search(p, t["text"].strip()) for p in self.SECTION_NUM_PATTERNS)
            for t in content_texts
        )

        # Very few meaningful text elements (dividers are sparse)
        is_sparse = len(content_texts) <= 3

        # No images (dividers are text-only in the template)
        no_images = len(images) == 0

        layout_name = slide.slide_layout.name.lower()
        layout_matches = any(d in layout_name for d in self.DIVIDER_LAYOUT_NAMES)

        # Layout name matches, but check content density first
        if layout_matches:
            if is_content_heavy:
                # Layout says divider but content says otherwise — return low
                # score so it falls through to Title and Content or References
                return 0.30
            elif is_sparse and no_images:
                return 0.95
            else:
                # Some content but not overwhelming — moderate confidence
                return 0.60

        # Strong signal: section number + sparse
        if has_section_num and is_sparse and no_images:
            return 0.8

        # Weak signal: sparse + very short text + no images.
        if is_sparse and no_images and content_texts:
            max_text_len = max(len(t["text"]) for t in content_texts)
            if max_text_len < 50 and len(content_texts) <= 2:
                return 0.3

        return 0.1

    # --- Extraction & Classification ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract and classify content from an input section divider slide.

        Uses shape-level grouping to properly separate:
        - Section number (e.g. "01", "Session 1:")
        - Title (largest/topmost shape)
        - Description (remaining text)
        - Footer (programme name)
        """
        shapes = extract_shapes_with_text(slide)

        result = {
            "section_num": "",
            "title": "",
            "description": "",
            "footer": "",
        }

        if not shapes:
            return result

        # Filter noise
        filtered = []
        for s in shapes:
            text = s["text"].strip()
            if not text or len(text) <= 1:
                continue
            if any(p in text.lower() for p in self.PLACEHOLDER_NOISE):
                continue
            filtered.append(s)

        section_nums = []
        titles = []
        footer_candidates = []

        for s in filtered:
            text = s["text"].strip()

            # Handle multi-line shapes: check each line
            lines = [l.strip() for l in text.split("\n") if l.strip()]

            # Skip shapes that are just bare slide numbers
            if len(lines) == 1 and re.match(r"^\d{1,3}$", lines[0]) and not re.match(r"^0\d", lines[0]):
                continue

            # Check if any line is a section number/label
            is_section_num = False
            for line in lines:
                if any(re.search(p, line) for p in self.SECTION_NUM_PATTERNS):
                    is_section_num = True
                    section_nums.append({"text": line, **s})
                    break

            if is_section_num:
                continue

            # Check if it looks like a footer
            if self._is_footer_text(text):
                footer_candidates.append(s)
                continue

            # Remaining: title or description
            titles.append(s)

        # Title: shape with largest font, or topmost shape
        titles.sort(key=lambda s: (s["font_size"] or 0), reverse=True)

        if titles:
            result["title"] = titles[0]["text"]
        if len(titles) >= 2:
            # Join remaining as description
            desc_shapes = sorted(titles[1:], key=lambda s: (s["top"], s["left"]))
            result["description"] = "\n".join(s["text"] for s in desc_shapes)

        # Section number
        if section_nums:
            result["section_num"] = section_nums[0]["text"]

        # Footer
        if footer_candidates:
            result["footer"] = footer_candidates[0]["text"]

        return result

    def _is_footer_text(self, text: str) -> bool:
        """Check if text looks like a footer/programme name."""
        footer_patterns = [
            r"(?i)cricos",
            r"(?i)^uq\s",
            r"(?i)university\s+of\s+queensland",
            r"(?i)executive\s+education",
            r"(?i)business\s+school",
            r"(?i)hbis\s+innovation",
            r"(?i)^presentation\s+title",
        ]
        # Only match if the text is also short (footers are typically
        # brief programme names, not slide content)
        if len(text) > 60:
            return False
        return any(re.search(p, text) for p in footer_patterns)

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill Section Divider placeholders with extracted content.
        ONLY sets .text — all formatting inherits from the template.
        """
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        # Section number/label
        if content.get("section_num") and self.PH_SECTION_NUM in placeholders:
            placeholders[self.PH_SECTION_NUM].text = content["section_num"]

        # Title
        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        # Description (right side)
        if content.get("description") and self.PH_DESCRIPTION in placeholders:
            placeholders[self.PH_DESCRIPTION].text = content["description"]

        # Footer (programme name)
        if content.get("footer") and self.PH_FOOTER in placeholders:
            placeholders[self.PH_FOOTER].text = content["footer"]

    def get_placeholder_map(self) -> dict:
        return {
            "section_num": self.PH_SECTION_NUM,
            "title": self.PH_TITLE,
            "description": self.PH_DESCRIPTION,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
