"""
Template utilities for creating output slides from the UQ template.
"""

import os
from pptx import Presentation

# Template lives alongside the app
TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "templates")
TEMPLATE_FILENAME = "preferred_template.pptx"


def open_template() -> Presentation:
    """Open a fresh copy of the UQ template presentation."""
    path = os.path.join(TEMPLATE_DIR, TEMPLATE_FILENAME)
    return Presentation(path)


def add_slide_from_layout(prs: Presentation, layout_index: int):
    """
    Add a new slide using the specified layout index from the template.
    Returns the new slide object.
    """
    layout = prs.slide_masters[0].slide_layouts[layout_index]
    return prs.slides.add_slide(layout)


def delete_slide(prs: Presentation, slide_index: int):
    """Delete a slide by index using XML manipulation."""
    rId = prs.slides._sldIdLst[slide_index].get(
        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    )
    prs.part.drop_rel(rId)
    sldId = prs.slides._sldIdLst[slide_index]
    prs.slides._sldIdLst.remove(sldId)


def delete_all_original_slides(prs: Presentation, num_new_slides: int = 1):
    """
    Delete all the original template slides, keeping only the newly added ones.
    The new slides are always appended at the end, so we delete from the
    beginning (in reverse order to preserve indices).
    """
    num_original = len(prs.slides) - num_new_slides
    for i in range(num_original - 1, -1, -1):
        delete_slide(prs, i)
