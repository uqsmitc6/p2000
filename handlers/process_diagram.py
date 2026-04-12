"""
Process Diagram Handler — Process flow slides with title and visual diagram.

Template layout 45 "Process Diagram":
    idx  0 — Title (top, w=9.76in)
    idx 31 — Text Title 1 (off-screen at left=19.11in — template guide, not usable)
    idx 32 — Text Description 1 (off-screen at left=19.11in — template guide, not usable)
    idx 17 — Footer
    idx 18 — Slide number

The process diagram visual content lives in shapes (arrows, boxes,
connectors) that are NOT in placeholders. This handler creates the
branded shell with the title; the actual diagram is preserved by
_preserve_visual_shapes() in converter.py which deep-copies group
shapes and images from the source slide.

Detection: slide has a title + multiple shapes arranged in a
flow/process pattern (horizontal or vertical sequence), or the
source layout name contains "process" or "diagram" or "flow".

Content to look for:
    - Slides with step-by-step flows (arrows between boxes)
    - SmartArt converted to groups
    - Process charts, workflow diagrams, decision trees
    - Typically has a title but minimal body text

CRITICAL RULE: Never set font properties on placeholders.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_shapes_with_text, extract_images


class ProcessDiagramHandler(SlideHandler):

    name = "Process Diagram"
    description = "Process flow or diagram slide with title"
    layout_name = "Process Diagram"
    layout_index = 45

    PH_TITLE = 0
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

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect process diagram / flow chart slides.

        Content patterns:
        - Source layout name contains process/diagram/flow/workflow
        - Multiple group shapes (SmartArt, grouped shapes)
        - Title + group shapes with little body text
        - Many small shapes arranged in a line (process steps)
        """
        if slide_index == 0:
            return 0.0

        layout_name = slide.slide_layout.name.lower()
        if any(kw in layout_name for kw in ("process", "diagram", "flow", "workflow")):
            return 0.75

        # Heuristic: look for group shapes and/or many small shapes
        # (indicative of a diagram)
        shapes = extract_shapes_with_text(slide)
        meaningful = self._get_meaningful_shapes(shapes)

        # Count shape types
        group_count = 0
        total_shape_count = 0

        for shape in slide.shapes:
            total_shape_count += 1
            if shape.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP
                group_count += 1

        # Process diagrams typically have group shapes (from SmartArt)
        # or many individual shapes (manually built diagrams)
        has_title = any(
            s["is_placeholder"] and s["shape_idx"] in (0, 3, 15)
            for s in meaningful
        ) or any(len(s["text"]) < 80 for s in meaningful)

        # Body text total (excluding title-like short text)
        body_text = sum(
            len(s["text"]) for s in meaningful
            if len(s["text"]) >= 80
        )

        if group_count >= 2 and has_title and body_text < 100:
            # Multiple groups + title + little body text = likely diagram
            return 0.58

        if total_shape_count >= 8 and group_count >= 1 and has_title and body_text < 50:
            # Many shapes including a group = process diagram
            return 0.52

        # Many shapes with group text (text extracted from groups)
        group_text_shapes = [s for s in shapes if s.get("is_group_text")]
        if len(group_text_shapes) >= 3 and has_title and body_text < 100:
            return 0.55

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title only. Diagram shapes are preserved by
        _preserve_visual_shapes() in converter.py.
        """
        shapes = extract_shapes_with_text(slide)

        result = {
            "title": "",
            "footer": "",
        }

        if not shapes:
            return result

        meaningful = self._get_meaningful_shapes(shapes)

        # Find title
        for s in meaningful:
            if s["is_placeholder"] and s["shape_idx"] in (0, 3, 15):
                result["title"] = s["text"]
                break

        # Fallback: topmost short text
        if not result["title"] and meaningful:
            sorted_by_top = sorted(meaningful, key=lambda s: (s["top"] or 0))
            for s in sorted_by_top:
                if len(s["text"]) < 100 and not s.get("is_group_text"):
                    result["title"] = s["text"]
                    break

        # Footer
        for s in meaningful:
            if s["is_placeholder"] and s["shape_idx"] in (17,):
                result["footer"] = s["text"]
                break
            elif self._is_footer_text(s["text"]):
                result["footer"] = s["text"]
                break

        return result

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill only the title. Diagram content is preserved by
        _preserve_visual_shapes() from converter.py.
        NEVER set font properties.
        """
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

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
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
