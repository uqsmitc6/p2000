"""
AI-powered slide classifier using Claude Vision API.

Called as a fallback when heuristic classification confidence is low.
Sends a rendered PNG screenshot of the slide to Claude Sonnet's vision
capability, which classifies the slide type based on its visual appearance.

This is far more accurate than text-only classification because it captures:
- Layout structure and whitespace
- Background colours and gradients
- Font sizes and visual hierarchy
- Images, logos, and decorative elements
- Overall "feel" of the slide (sparse divider vs dense content)

Pipeline:
    1. Render the entire PPTX to per-slide PNGs (done once, cached)
    2. For each ambiguous slide, send its PNG to Claude Vision
    3. Claude returns a JSON classification

Requires: LibreOffice (for rendering), anthropic SDK (for API calls)
"""

import base64
import json
import logging

logger = logging.getLogger("uqslide.classifier")


# Available layout types and their descriptions — used in the prompt
LAYOUT_DESCRIPTIONS = {
    "Cover 1": (
        "The first slide of a presentation. Contains the programme/course title, "
        "the presenting entity (e.g. 'UQ Business School Executive Education'), "
        "and optionally presenter names and dates. Usually a white or light "
        "background with a UQ purple graphic device or gradient."
    ),
    "Section Divider": (
        "A section break slide that separates major topics or modules. "
        "Typically has a section number (01, 02, Session 1, Block 1, etc.) "
        "and a short section title. Very sparse — usually 2-3 text elements "
        "max, no bullet points or body content. Often has a distinctive "
        "background colour or large typography."
    ),
    "Text with Image": (
        "A content slide with text on one side and an image/photo on the other. "
        "The image takes up a significant portion of the slide (roughly 1/3 to 2/3). "
        "The text side typically has a title and some body content. The image may "
        "be a photograph, illustration, stock image, or diagram. Distinguishable "
        "from Title and Content by the prominent placement of a large image "
        "alongside text, rather than text filling the full width."
    ),
    "Title and Table": (
        "A slide containing a data table. The table has visible rows and columns "
        "with text in cells. May also have a title above the table. Tables might "
        "contain schedules, comparison matrices, data summaries, rubrics, or "
        "structured information. The table is the dominant element on the slide."
    ),
    "Two Content": (
        "A slide with two distinct columns of content side by side, with roughly "
        "EQUAL column widths. Each column may contain bullet points, text, or "
        "lists. Often used for comparisons, pros/cons, before/after, or two "
        "parallel concepts. The key visual signal is two equally-sized text areas "
        "arranged horizontally with a title above."
    ),
    "Split Content": (
        "A slide with two columns of UNEQUAL width — one narrow (roughly 1/3) "
        "and one wide (roughly 2/3). The narrow column usually has a key concept, "
        "summary, or introduction, while the wide column has detailed content. "
        "Distinguished from Two Content by the clear asymmetry in column widths."
    ),
    "Quote": (
        "A quote slide featuring a prominent quotation with attribution. "
        "The quote is typically in large text, often with quotation marks, and "
        "the author/source is shown below in smaller text. May have a background "
        "image. Very sparse — usually just the quote and attribution, no bullets "
        "or body content."
    ),
    "Title Only": (
        "A slide with just a title and visual content (images, diagrams, charts, "
        "flowcharts, process diagrams) that fills the content area. Distinguished "
        "from Title and Content by having minimal or no body TEXT — the content "
        "is primarily visual. The title is the only meaningful text element."
    ),
    "Title and Content": (
        "A standard content slide with a title at the top and body content below. "
        "The body may contain bullet points, paragraphs, references, discussion "
        "questions, activity instructions, or any other teaching content. "
        "This is the most common slide type — the default for anything that "
        "isn't clearly a cover, divider, closing slide, image slide, table, "
        "two-column, or quote slide."
    ),
    "Thank You": (
        "A closing/contact slide at the end of the presentation. Contains "
        "'Thank you', 'Questions?', 'Contact', or similar closing text. "
        "May include presenter name, email, phone number, or job title. "
        "Often sparse with large text."
    ),
    "Acknowledgement of Country": (
        "A slide acknowledging Traditional Owners and custodians of the land. "
        "Contains phrases like 'Acknowledgement of Country', 'Traditional Owners', "
        "'Traditional Custodians', 'Elders past', or 'custodianship of the lands'. "
        "This is automatically replaced with a standard branded version."
    ),
    "References": (
        "A slide listing academic references, citations, bibliography, image credits, "
        "or sources. Contains DOIs, journal names, author names with years (2019), "
        "'et al.', 'Vol.', 'pp.', or image credit lines (Adobe Stock, Unsplash, etc.). "
        "Title is typically 'References', 'Selected References', 'Bibliography', "
        "'Sources', 'Image Credits', or 'Further Reading'."
    ),
    "Skip": (
        "A slide that doesn't fit any of the above types and should be "
        "excluded from the output. Examples: blank slides, housekeeping "
        "slides (restrooms, WiFi, Zoom details), break slides (Morning Tea, "
        "Lunch), or slides with only images/logos "
        "and no meaningful teaching content."
    ),
}


VISION_CLASSIFICATION_PROMPT = """You are classifying PowerPoint slides for a university presentation converter tool.

Look at this slide image and classify it as one of these types:

{layout_descriptions}

Context:
- This is slide {slide_number} of {total_slides} in the deck.
- The presentation is for a university course or executive education programme.

Based on the visual appearance, layout, and content of this slide, classify it.

Respond with ONLY a JSON object (no markdown, no code fences):
{{"type": "<one of: Cover 1, Acknowledgement of Country, Section Divider, Text with Image, Title and Table, Two Content, Split Content, Quote, Title Only, Title and Content, References, Thank You, Skip>", "confidence": <0.0 to 1.0>, "reason": "<brief explanation>"}}"""


def classify_slide_with_api(
    slide,
    slide_index: int,
    total_slides: int,
    api_key: str,
    model: str = "claude-sonnet-4-6",
    slide_image: bytes = None,
) -> dict:
    """
    Classify a slide using Claude Vision API.

    Args:
        slide: python-pptx slide object (used as fallback for text-only if no image)
        slide_index: 0-based index of this slide
        total_slides: total slides in the deck
        api_key: Anthropic API key
        model: Claude model to use
        slide_image: PNG bytes of the rendered slide. If None, falls back to
                     text-only classification.

    Returns:
        {"type": str, "confidence": float, "reason": str}
        or {"type": None, "error": str} on failure
    """
    try:
        import anthropic

        logger.info("Classifying slide %d/%d via %s (%s)",
                     slide_index + 1, total_slides, model,
                     "vision" if slide_image else "text-only")
        client = anthropic.Anthropic(api_key=api_key)

        layout_descs = "\n".join(
            f"- **{name}**: {desc}" for name, desc in LAYOUT_DESCRIPTIONS.items()
        )

        prompt_text = VISION_CLASSIFICATION_PROMPT.format(
            layout_descriptions=layout_descs,
            slide_number=slide_index + 1,
            total_slides=total_slides,
        )

        if slide_image:
            # Vision-based classification (preferred)
            image_b64 = base64.b64encode(slide_image).decode("utf-8")

            message = client.messages.create(
                model=model,
                max_tokens=200,
                messages=[{
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": "image/png",
                                "data": image_b64,
                            },
                        },
                        {
                            "type": "text",
                            "text": prompt_text,
                        },
                    ],
                }],
            )
        else:
            # Fallback: text-only classification (if rendering failed)
            text_desc = _build_text_description(slide, slide_index, total_slides)
            message = client.messages.create(
                model=model,
                max_tokens=200,
                messages=[{"role": "user", "content": text_desc}],
            )

        response_text = message.content[0].text.strip()

        # Strip markdown code fences if present
        if response_text.startswith("```"):
            lines = response_text.split("\n")
            # Remove first and last lines (```json and ```)
            lines = [l for l in lines if not l.strip().startswith("```")]
            response_text = "\n".join(lines).strip()

        # Parse JSON response
        result = json.loads(response_text)

        # Validate
        valid_types = list(LAYOUT_DESCRIPTIONS.keys())
        if result.get("type") not in valid_types:
            return {"type": None, "error": f"Invalid type: {result.get('type')}"}

        logger.info("Slide %d classified as '%s' (%.2f): %s",
                     slide_index + 1, result["type"],
                     float(result.get("confidence", 0.8)),
                     result.get("reason", ""))
        return {
            "type": result["type"],
            "confidence": float(result.get("confidence", 0.8)),
            "reason": result.get("reason", ""),
        }

    except json.JSONDecodeError as e:
        logger.error("Slide %d: failed to parse API response: %s | raw: %s",
                      slide_index + 1, e, response_text[:200] if 'response_text' in dir() else "N/A")
        return {"type": None, "error": f"Failed to parse API response: {e}"}
    except Exception as e:
        logger.error("Slide %d: API call failed: %s", slide_index + 1, e, exc_info=True)
        return {"type": None, "error": f"API call failed: {e}"}


def _build_text_description(slide, slide_index: int, total_slides: int) -> str:
    """
    Fallback: build a text description when no image is available.
    Used if LibreOffice rendering fails.
    """
    from utils.extractor import extract_text_elements, extract_images

    texts = extract_text_elements(slide)
    images = extract_images(slide)

    text_lines = []
    for t in texts:
        text = t["text"].strip()
        if not text:
            continue
        font_info = ""
        if t.get("font_size"):
            font_info += f" [{t['font_size']}pt"
            if t.get("bold"):
                font_info += ", bold"
            font_info += "]"
        if len(text) > 150:
            text = text[:150] + "..."
        text_lines.append(f'  - "{text}"{font_info}')

    if not text_lines:
        text_lines = ["  (no text content)"]

    layout_descs = "\n".join(
        f"- **{name}**: {desc}" for name, desc in LAYOUT_DESCRIPTIONS.items()
    )

    prompt = f"""You are classifying PowerPoint slides for a university presentation converter tool.

Given the following description of a slide's content, classify it as one of these types:

{layout_descs}

Slide number: {slide_index + 1} of {total_slides}
Layout name in source file: "{slide.slide_layout.name}"
Number of shapes: {len(slide.shapes)}
Number of images: {len(images)}

Text elements (in order of position, top to bottom):
{chr(10).join(text_lines)}

Respond with ONLY a JSON object (no markdown, no code fences):
{{"type": "<one of: Cover 1, Acknowledgement of Country, Section Divider, Text with Image, Title and Table, Two Content, Split Content, Quote, Title Only, Title and Content, References, Thank You, Skip>", "confidence": <0.0 to 1.0>, "reason": "<brief explanation>"}}"""

    return prompt


def classify_slides_batch(
    slides_with_indices: list,
    total_slides: int,
    api_key: str,
    model: str = "claude-sonnet-4-6",
    slide_images: dict = None,
) -> list:
    """
    Classify multiple slides. Each call is independent.

    Args:
        slides_with_indices: list of (slide, slide_index) tuples
        total_slides: total slides in the deck
        api_key: Anthropic API key
        model: model to use
        slide_images: dict mapping slide_index → PNG bytes (optional)

    Returns:
        list of classification results (same order as input)
    """
    results = []
    for slide, slide_index in slides_with_indices:
        image = slide_images.get(slide_index) if slide_images else None
        result = classify_slide_with_api(
            slide, slide_index, total_slides, api_key, model,
            slide_image=image,
        )
        results.append(result)
    return results
