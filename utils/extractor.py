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

    Preserves spaces between runs by checking whether the previous
    run's text already ends with whitespace or the next run starts
    with whitespace. If neither, inserts a space.
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

            if not run_text:
                continue

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

            # Insert space between runs if needed:
            # Only when current_text has content AND neither side
            # already has whitespace at the boundary.
            if current_text and run_text:
                if not current_text[-1].isspace() and not run_text[0].isspace():
                    current_text += " "

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


def extract_shapes_with_text(slide) -> list[dict]:
    """
    Extract text grouped by shape. Each shape becomes one entry with
    all its paragraphs joined. This preserves the shape-level grouping
    that's needed to distinguish title shapes from body shapes.

    Returns a list of dicts:
        shape_name, shape_idx (placeholder idx or None), text,
        font_size (of largest run), bold, left, top, width, height,
        is_placeholder, placeholder_type
    """
    shapes = []

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        # Collect all paragraph texts for this shape
        paras = []
        max_font_size = None
        any_bold = False

        for para in shape.text_frame.paragraphs:
            para_text = para.text.strip()
            if para_text:
                paras.append(para_text)

            for run in para.runs:
                if run.font.size:
                    sz = run.font.size.pt
                    if max_font_size is None or sz > max_font_size:
                        max_font_size = sz
                if run.font.bold:
                    any_bold = True

        full_text = "\n".join(paras)
        if not full_text.strip():
            continue

        # Placeholder info
        is_ph = shape.is_placeholder if hasattr(shape, 'is_placeholder') else False
        ph_idx = None
        ph_type = None
        if is_ph:
            ph_idx = shape.placeholder_format.idx
            ph_type = str(shape.placeholder_format.type) if shape.placeholder_format.type else None

        shapes.append({
            "shape_name": shape.name,
            "shape_idx": ph_idx,
            "text": full_text,
            "font_size": max_font_size,
            "bold": any_bold,
            "left": shape.left,
            "top": shape.top,
            "width": shape.width,
            "height": shape.height,
            "is_placeholder": is_ph,
            "placeholder_type": ph_type,
        })

    return shapes


def extract_images(slide) -> list[dict]:
    """
    Extract all images from a slide.

    Checks three sources:
        1. Picture shapes (shape_type 13) — the standard way
        2. Placeholder shapes with an embedded image (e.g., content
           placeholders where someone inserted a picture)
        3. Shapes with image fill (background images in shapes)

    Returns list of dicts with: name, content_type, blob, left, top, width, height
    """
    images = []
    seen_ids = set()  # Avoid duplicates

    for shape in slide.shapes:
        image_data = None

        # Method 1: Standard Picture shapes
        if shape.shape_type == 13:
            try:
                image_data = {
                    "name": shape.name,
                    "content_type": shape.image.content_type,
                    "blob": shape.image.blob,
                    "left": shape.left,
                    "top": shape.top,
                    "width": shape.width,
                    "height": shape.height,
                }
            except (ValueError, AttributeError):
                image_data = {
                    "name": shape.name,
                    "content_type": None,
                    "blob": None,
                    "left": shape.left,
                    "top": shape.top,
                    "width": shape.width,
                    "height": shape.height,
                }

        # Method 2: Placeholders or other shapes with .image property
        elif hasattr(shape, 'image'):
            try:
                blob = shape.image.blob
                if blob:
                    image_data = {
                        "name": shape.name,
                        "content_type": shape.image.content_type,
                        "blob": blob,
                        "left": shape.left,
                        "top": shape.top,
                        "width": shape.width,
                        "height": shape.height,
                    }
            except (ValueError, AttributeError):
                pass

        if image_data:
            # Deduplicate by position + size
            key = (image_data["left"], image_data["top"],
                   image_data["width"], image_data["height"])
            if key not in seen_ids:
                seen_ids.add(key)
                images.append(image_data)

    return images
