"""
Title and Table Handler — Slides containing tabular data.

Template layout 40 "Title and Table":
    idx  0 — Title
    idx 31 — Subtitle (optional)
    idx 19 — Table placeholder
    idx 17 — Footer (programme name)
    idx 18 — Slide number

Detection: slide has a table shape (shape.has_table == True).

Table handling:
    - Extracts row/column data from source table
    - Creates a new table in the template's table placeholder
    - Populates cells with extracted text
    - Does NOT copy cell formatting (merges, colours etc.) — relies on
      theme/master styles from the UQ template

CRITICAL RULE: Never set font properties on placeholders or table cells.
"""

import re
from handlers.base import SlideHandler
from utils.extractor import extract_text_elements, extract_shapes_with_text


class TitleTableHandler(SlideHandler):

    name = "Title and Table"
    description = "Content slide with a title and a data table"
    layout_name = "Title and Table"
    layout_index = 40

    # Placeholder indices for layout 40
    PH_TITLE = 0
    PH_SUBTITLE = 31
    PH_TABLE = 19
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
        Detect slides that contain a table.
        High confidence because tables are unambiguous.
        """
        if slide_index == 0:
            return 0.0

        has_table = False
        for shape in slide.shapes:
            if shape.has_table:
                has_table = True
                break

        if not has_table:
            return 0.0

        # Table present → high confidence
        return 0.75

    # --- Extraction ---

    def extract_content(self, slide, slide_index: int) -> dict:
        """
        Extract title, subtitle, and table data from a slide.

        Table data is stored as a list of rows, where each row is a
        list of cell text strings. First row is assumed to be the header.
        """
        shapes = extract_shapes_with_text(slide)

        result = {
            "title": "",
            "subtitle": "",
            "table_data": [],      # list of lists: [[cell, cell, ...], ...]
            "table_rows": 0,
            "table_cols": 0,
            "footer": "",
        }

        # --- Extract table data ---
        for shape in slide.shapes:
            if shape.has_table:
                table = shape.table
                result["table_rows"] = len(table.rows)
                result["table_cols"] = len(table.columns)

                for row in table.rows:
                    row_data = []
                    for cell in row.cells:
                        # Get cell text, preserving paragraph breaks
                        cell_text = "\n".join(
                            para.text.strip()
                            for para in cell.text_frame.paragraphs
                            if para.text.strip()
                        )
                        row_data.append(cell_text)
                    result["table_data"].append(row_data)

                break  # Use first table found

        # --- Extract text content (title, subtitle, footer) ---
        if not shapes:
            return result

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
                    pass  # slide number
                else:
                    # Other text placeholders (not the table itself)
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

        # Subtitle — first remaining body shape if short
        if body_shapes and title_shape:
            body_shapes.sort(key=lambda s: (s["top"], s["left"]))
            candidate = body_shapes[0]
            if len(candidate["text"]) < 100:
                result["subtitle"] = candidate["text"]

        return result

    def _is_footer_text(self, text: str) -> bool:
        """Check if text looks like a footer/programme name."""
        footer_patterns = [
            r"(?i)cricos", r"(?i)hbis\s+innovation",
            r"(?i)executive\s+education",
            r"(?i)presentation\s+title",
        ]
        return any(re.search(p, text) for p in footer_patterns)

    # --- Output ---

    def fill_slide(self, slide, content: dict) -> None:
        """
        Fill title/subtitle placeholders and create table from extracted data.
        NEVER set font properties — all formatting inherits from theme.
        """
        placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

        # Title
        if content.get("title") and self.PH_TITLE in placeholders:
            placeholders[self.PH_TITLE].text = content["title"]

        # Subtitle
        if content.get("subtitle") and self.PH_SUBTITLE in placeholders:
            placeholders[self.PH_SUBTITLE].text = content["subtitle"]

        # Table
        table_data = content.get("table_data", [])
        rows = content.get("table_rows", 0)
        cols = content.get("table_cols", 0)

        if table_data and rows > 0 and cols > 0 and self.PH_TABLE in placeholders:
            table_ph = placeholders[self.PH_TABLE]

            # insert_table returns a GraphicFrame; .table gives the Table object
            graphic_frame = table_ph.insert_table(rows, cols)
            table = graphic_frame.table

            for r_idx, row_data in enumerate(table_data):
                for c_idx, cell_text in enumerate(row_data):
                    if c_idx < cols and r_idx < rows:
                        table.cell(r_idx, c_idx).text = cell_text

    def get_placeholder_map(self) -> dict:
        return {
            "title": self.PH_TITLE,
            "subtitle": self.PH_SUBTITLE,
            "table": self.PH_TABLE,
            "footer": self.PH_FOOTER,
            "slide_num": self.PH_SLIDE_NUM,
        }
