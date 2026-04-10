"""
Content extraction utilities for input slides.
Handles the messy reality of academic PowerPoint files — text in random
shapes, mixed font sizes in single placeholders, soft returns, etc.
"""

from lxml import etree

NS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}


def extract_text_elements(slide) -> list[dict]:
    """
    Extract all text elements from a slide, splitting by font size
    changes and line breaks within paragraphs.

    Returns a list of dicts with keys:
        text, font_size, font_name, bold, left, top, width, height
    """
    elements = []

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        for para in shape.text_frame.paragraphs:
            run_groups = _split_paragraph_into_groups(para)

            for group in run_groups:
                elements.append({
                    **group,
                    "left": shape.left,
                    "top": shape.top,
                    "width": shape.width,
                    "height": shape.height,
                })

            # Paragraph with no runs but has text
            if not para.runs and para.text.strip():
                elements.append({
                    "text": para.text.strip(),
                    "font_size": None,
                    "font_name": None,
                    "bold": False,
                    "left": shape.left,
                    "top": shape.top,
                    "width": shape.width,
                    "height": shape.height,
                })

    return elements


def _split_paragraph_into_groups(para) -> list[dict]:
    """
    Split a paragraph into groups based on <a:br/> breaks and font
    size changes. Each group becomes a separate content element.
    """
    groups = []
    current_text = ""
    current_size = None
    current_name = None
    current_bold = False

    para_elem = para._p

    for child in para_elem:
        tag = child.tag.split('}')[1] if '}' in child.tag else child.tag

        if tag == 'br':
            # Line break — flush current group
            if current_text.strip():
                groups.append({
                    "text": current_text.strip(),
                    "font_size": current_size,
                    "font_name": current_name,
                    "bold": current_bold,
                })
                current_text = ""
                current_size = None
                current_name = None
                current_bold = False

        elif tag == 'r':
            run_text_elem = child.find('a:t', NS)
            run_text = run_text_elem.text if run_text_elem is not None else ""

            if run_text.strip():
                rPr = child.find('a:rPr', NS)
                run_size = None
                run_name = None
                run_bold = False

                if rPr is not None:
                    sz = rPr.get('sz')
                    if sz:
                        run_size = int(sz) / 100
                    b = rPr.get('b')
                    if b == '1':
                        run_bold = True
                    latin = rPr.find('a:latin', NS)
                    if latin is not None:
                        run_name = latin.get('typeface')

                # Font size change → start new group
                if (run_size and current_size and
                        run_size != current_size and current_text.strip()):
                    groups.append({
                        "text": current_text.strip(),
                        "font_size": current_size,
                        "font_name": current_name,
                        "bold": current_bold,
                    })
                    current_text = ""

                current_text += run_text
                if run_size:
                    current_size = run_size
                if run_name:
                    current_name = run_name
                if run_bold:
                    current_bold = True

    # Flush last group
    if current_text.strip():
        groups.append({
            "text": current_text.strip(),
            "font_size": current_size,
            "font_name": current_name,
            "bold": current_bold,
        })

    return groups


def extract_images(slide) -> list[dict]:
    """
    Extract all images from a slide.
    Returns list of dicts with: name, content_type, blob, left, top, width, height
    """
    images = []
    for shape in slide.shapes:
        if shape.shape_type == 13:  # Picture
            try:
                images.append({
                    "name": shape.name,
                    "content_type": shape.image.content_type,
                    "blob": shape.image.blob,
                    "left": shape.left,
                    "top": shape.top,
                    "width": shape.width,
                    "height": shape.height,
                })
            except ValueError:
                # Linked images (not embedded) — note presence but no blob
                images.append({
                    "name": shape.name,
                    "content_type": None,
                    "blob": None,
                    "left": shape.left,
                    "top": shape.top,
                    "width": shape.width,
                    "height": shape.height,
                })
    return images
