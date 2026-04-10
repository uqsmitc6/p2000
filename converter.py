"""
Core conversion engine.

Takes an input PPTX, detects slide types, extracts content,
and produces a brand-compliant output PPTX using the UQ template.

Classification pipeline:
    1. Render all slides to PNG images (once, upfront) if API key provided
    2. Heuristics run first (free, instant)
    3. If confidence >= 0.7 → auto-convert (no API call)
    4. If confidence < 0.7 and API key provided → send slide image to
       Claude Vision for classification
    5. If no API key → fall back to heuristic result with flag

Tiers in the output report:
    - CONVERTED: confident match (heuristic >= 0.7 or API-confirmed)
    - FLAGGED: low confidence, converted but review recommended
    - SKIPPED: no match or classified as "Skip" by the API
"""

import io
from pptx import Presentation

from handlers import get_all_handlers, HANDLER_REGISTRY
from utils.template import open_template, add_slide_from_layout, delete_all_original_slides

# Confidence thresholds
CONFIDENT_THRESHOLD = 0.7   # No API call needed
CONVERT_THRESHOLD = 0.35    # Minimum to convert without API
# Below CONVERT_THRESHOLD without API → skip


def convert_presentation(
    input_bytes: bytes,
    api_key: str = None,
    model: str = "claude-sonnet-4-6",
    progress_callback=None,
) -> tuple[bytes, dict]:
    """
    Convert an uploaded PPTX to brand-compliant format.

    Args:
        input_bytes: Raw bytes of the uploaded .pptx file
        api_key: Optional Anthropic API key. If provided, Claude Vision
                 classifies ambiguous slides using rendered images.
        model: Claude model for classification (default: claude-sonnet-4-6)
        progress_callback: Optional callable(message: str) for status updates

    Returns:
        (output_bytes, report)
    """
    input_prs = Presentation(io.BytesIO(input_bytes))
    handlers = get_all_handlers()
    total_slides = len(input_prs.slides)

    report = {
        "slides_converted": 0,
        "slides_flagged": 0,
        "slides_skipped": 0,
        "api_calls": 0,
        "details": [],
    }

    # --- Pre-render slides to images if API key provided ---
    slide_images = {}
    if api_key:
        if progress_callback:
            progress_callback("Rendering slides to images for AI classification...")
        slide_images = _render_slide_images(input_bytes)
        if slide_images:
            if progress_callback:
                progress_callback(
                    f"Rendered {len(slide_images)} slide images. Classifying..."
                )
        else:
            if progress_callback:
                progress_callback(
                    "Could not render slide images. Using text-only fallback."
                )

    output_prs = open_template()
    new_slides_added = 0

    for slide_idx, slide in enumerate(input_prs.slides):

        if progress_callback:
            progress_callback(
                f"Processing slide {slide_idx + 1} of {total_slides}..."
            )

        # --- Step 1: Heuristic scoring ---
        scores = {}
        for name, handler in handlers.items():
            scores[name] = handler.detect(slide, slide_idx)

        best_name = max(scores, key=scores.get)
        best_confidence = scores[best_name]
        classification_method = "heuristic"

        # --- Step 2: API fallback for low-confidence slides ---
        if best_confidence < CONFIDENT_THRESHOLD and api_key:
            api_result = _classify_with_api(
                slide, slide_idx, total_slides, api_key, model,
                slide_image=slide_images.get(slide_idx),
            )

            if api_result and api_result.get("type"):
                report["api_calls"] += 1
                api_type = api_result["type"]
                api_confidence = api_result.get("confidence", 0.8)
                classification_method = "api"

                if api_type == "Skip":
                    # API says skip this slide
                    report["slides_skipped"] += 1
                    report["details"].append({
                        "slide": slide_idx + 1,
                        "status": "skipped",
                        "handler": None,
                        "confidence": api_confidence,
                        "content": None,
                        "preview": _get_slide_preview(slide),
                        "all_scores": scores,
                        "classification_method": "api",
                        "api_reason": api_result.get("reason", ""),
                    })
                    continue

                elif api_type in HANDLER_REGISTRY:
                    # API classified it — use that handler
                    best_name = api_type
                    best_confidence = api_confidence

        # --- Step 3: Convert or skip ---
        best_handler = handlers.get(best_name)
        slide_preview = _get_slide_preview(slide)

        if best_handler and best_confidence >= CONVERT_THRESHOLD:
            content = best_handler.extract_content(slide, slide_idx)

            new_slide = add_slide_from_layout(output_prs, best_handler.layout_index)
            best_handler.fill_slide(new_slide, content)
            new_slides_added += 1

            if best_confidence >= CONFIDENT_THRESHOLD:
                status = "converted"
            else:
                status = "flagged"
                report["slides_flagged"] += 1

            report["slides_converted"] += 1
            report["details"].append({
                "slide": slide_idx + 1,
                "status": status,
                "handler": best_handler.name,
                "confidence": best_confidence,
                "content": content,
                "preview": slide_preview,
                "all_scores": scores,
                "classification_method": classification_method,
            })
        else:
            report["slides_skipped"] += 1
            report["details"].append({
                "slide": slide_idx + 1,
                "status": "skipped",
                "handler": None,
                "confidence": best_confidence,
                "content": None,
                "preview": slide_preview,
                "all_scores": scores,
                "classification_method": classification_method,
                "reason": f"Best match '{best_name}' at {best_confidence:.2f}",
            })

    # Remove original template slides
    if new_slides_added > 0:
        delete_all_original_slides(output_prs, num_new_slides=new_slides_added)

    # Save
    output_buffer = io.BytesIO()
    output_prs.save(output_buffer)
    output_buffer.seek(0)

    return output_buffer.getvalue(), report


def _render_slide_images(input_bytes: bytes) -> dict:
    """
    Render all slides to PNG images using LibreOffice.

    Returns:
        dict mapping slide_index (0-based) → PNG bytes
        Empty dict if rendering fails.
    """
    try:
        from utils.renderer import render_slides_to_images
        images = render_slides_to_images(input_bytes, dpi=150)
        return {i: img for i, img in enumerate(images)}
    except Exception:
        return {}


def _classify_with_api(slide, slide_index, total_slides, api_key, model,
                       slide_image=None):
    """Call the AI classifier. Returns result dict or None on failure."""
    try:
        from utils.classifier import classify_slide_with_api
        return classify_slide_with_api(
            slide, slide_index, total_slides, api_key, model,
            slide_image=slide_image,
        )
    except Exception:
        return None


def _get_slide_preview(slide) -> str:
    """Get a short text preview of a slide for reporting."""
    texts = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            t = shape.text_frame.text.strip()
            if t and len(t) > 1:
                texts.append(t[:60])
    return " | ".join(texts[:3])[:120] if texts else "(no text)"


def convert_cover_only(input_bytes: bytes) -> tuple[bytes, dict]:
    """
    Simplified conversion — only processes the first slide as a cover.
    """
    input_prs = Presentation(io.BytesIO(input_bytes))

    if len(input_prs.slides) == 0:
        raise ValueError("The uploaded file contains no slides.")

    from handlers.cover1 import Cover1Handler
    handler = Cover1Handler()

    slide = input_prs.slides[0]
    content = handler.extract_content(slide, 0)

    output_prs = open_template()
    new_slide = add_slide_from_layout(output_prs, handler.layout_index)
    handler.fill_slide(new_slide, content)
    delete_all_original_slides(output_prs, num_new_slides=1)

    output_buffer = io.BytesIO()
    output_prs.save(output_buffer)
    output_buffer.seek(0)

    report = {
        "slides_converted": 1,
        "slides_flagged": 0,
        "slides_skipped": len(input_prs.slides) - 1,
        "api_calls": 0,
        "details": [{
            "slide": 1,
            "status": "converted",
            "handler": handler.name,
            "content": content,
        }],
    }

    return output_buffer.getvalue(), report
