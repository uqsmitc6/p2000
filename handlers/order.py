"""
Order Handler — Four numbered steps in a horizontal sequence.

Template layout 28 "Order":
    idx  0 — Title (full width)
    idx 42 — Subtitle (full width, h=0.55in)

    Step 1: idx 29 (number circle, 1.73in × 1.73in)
            idx 19 (label, 2.57in × 0.26in)
            idx 20 (description, 2.57in × 1.66in)

    Step 2: idx 30 (number circle)
            idx 31 (label)
            idx 32 (description)

    Step 3: idx 33 (number circle)
            idx 34 (label)
            idx 35 (description)

    Step 4: idx 36 (number circle)
            idx 37 (label)
            idx 38 (description)

    idx 40 — Footer
    idx 41 — Slide number

The number circle placeholders (29, 30, 33, 36) are large squares
intended to hold step numbers (1, 2, 3, 4). Labels are short titles
for each step. Descriptions provide detail.

Detection: slide presenting a numbered sequence or ordered process
(3-4 steps). Distinguished from Process Diagram by having text-based
steps rather than visual flowcharts.

Content to look for:
    - Step-by-step instructions or procedures
    - Numbered processes (onboarding, project phases)
    - Sequential workflows described in text
    - Timeline with ordered milestones

CRITICAL RULE: Never set font properties on placeholders.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_shapes_with_text


class OrderHandler(SlideHandler):

    name = "Order"
    description = "Four numbered steps in a horizontal sequence"
    layout_name = "Order"
    layout_index = 28

    PH_TITLE = 0
    PH_SUBTITLE = 42

    # (number_idx, label_idx, description_idx) for each step
    STEPS = [
        (29, 19, 20),  # Step 1
        (30, 31, 32),  # Step 2
        (33, 34, 35),  # Step 3
        (36, 37, 38),  # Step 4
    ]

    PH_FOOTER = 40
    PH_SLIDE_NUM = 41

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

    # Patterns that suggest ordered/sequential content
    ORDER_KEYWORDS = [
        r"(?i)\bstep\s+\d", r"(?i)\bphase\s+\d",
        r"(?i)\bstage\s+\d", r"(?i)^\d+\.\s",
        r"(?i)\bfirst\b.*\bsecond\b", r"(?i)\bthen\b",
        r"(?i)\bnext\b", r"(?i)\bfinally\b",
    ]

    SLIDE_WIDTH = 12192000

    # --- Detection ---

    def detect(self, slide, slide_index: int) -> float:
        """
        Detect slides with numbered/ordered steps.

        Content patterns:
        - Source layout name contains "order" or "step" or "sequence"
        - 3-4 short text groups with numbering patterns
        - Text distributed across 3-4 horizontal positions
        """
        if slide_index == 0:
            return 0.0

        layout_name = slide.slide_layout.name.lower()
        if layout_name == "order" or "order" in layout_name:
            return 0.75

        shapes = extract_shapes_with_text(slide)
        meaningful = self._get_meaningful_shapes(shapes)

        if len(meaningful) < 4:
            return 0.0

        title_shape = self._find_title(meaningful)
        body_shapes = [s for s in meaningful if s is not title_shape]

        if len(body_shapes) < 3:
            return 0.0

        # Check for ordering keywords in the text
        all_text = " ".join(s["text"] for s in body_shapes)
        order_signals = sum(
            1 for p in self.ORDER_KEYWORDS if re.search(p, all_text)
        )

        # Check distribution across columns (3-4 groups)
        quarter_w = self.SLIDE_WIDTH / 4
        cols = [0, 0, 0, 0]
        for s in body_shapes:
            cx = (s["left"] or 0) + (s["width"] or 0) / 2
            col = min(3, int(cx / quarter_w))
            cols[col] += 1

        filled_cols = sum(1 for c in cols if c > 0)

        if filled_cols >= 3 and order_signals >= 2:
            return 0.55
        elif filled_cols >= 4 and len(body_shapes) >= 8:
            # 4 groups with multiple items each
            return 0.48

        return 0.0

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, subtitle, and up to 4 steps. Each step has a
        number, label, and description. Steps are extracted by horizontal
        position (left to right).
        """
        shapes = extract_shapes_with_text(slide)

        result = {
            "title": "",
            "subtitle": "",
            "steps": [],  # List of {"number": str, "label": str, "description": str}
            "footer": "",
        }

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
                elif idx in (17, 40):
                    if not footer_text:
                        footer_text = s["text"]
                elif idx in (18, 41):
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

        # Group into columns (up to 4)
        quarter_w = self.SLIDE_WIDTH / 4
        columns = [[] for _ in range(4)]

        for s in body_shapes:
            cx = (s["left"] or 0) + (s["width"] or 0) / 2
            col = min(3, int(cx / quarter_w))
            columns[col].append(s)

        for col in columns:
            col.sort(key=lambda s: (s["top"] or 0))

        # Build steps from each non-empty column
        for i, col_shapes in enumerate(columns):
            if not col_shapes:
                continue

            step = {"number": str(len(result["steps"]) + 1), "label": "", "description": ""}

            if len(col_shapes) == 1:
                step["label"] = col_shapes[0]["text"]
            elif len(col_shapes) == 2:
                step["label"] = col_shapes[0]["text"]
                step["description"] = col_shapes[1]["text"]
            else:
                # First is number/icon, second is label, rest is description
                first_text = col_shapes[0]["text"]
                if re.match(r"^\d{1,2}$", first_text.strip()):
                    step["number"] = first_text.strip()
                    step["label"] = col_shapes[1]["text"] if len(col_shapes) > 1 else ""
                    step["description"] = "\n".join(
                        s["text"] for s in col_shapes[2:]
                    )
                else:
                    step["label"] = first_text
                    step["description"] = "\n".join(
                        s["text"] for s in col_shapes[1:]
                    )

            result["steps"].append(step)

            if len(result["steps"]) >= 4:
                break

        return result

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill title, subtitle, and up to 4 step cells (number + label +
        description). NEVER set font properties.
        """
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        if content.get("subtitle") and self.PH_SUBTITLE in placeholders:
            placeholders[self.PH_SUBTITLE].text = content["subtitle"]

        steps = content.get("steps", [])
        for i, (num_idx, label_idx, desc_idx) in enumerate(self.STEPS):
            if i >= len(steps):
                break
            step = steps[i]
            if step.get("number") and num_idx in placeholders:
                placeholders[num_idx].text = step["number"]
            if step.get("label") and label_idx in placeholders:
                placeholders[label_idx].text = step["label"]
            if step.get("description") and desc_idx in placeholders:
                placeholders[desc_idx].text = step["description"]

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
            # Don't filter single digits here — they could be step numbers
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
            "steps": self.STEPS,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
