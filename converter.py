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
import logging
from pptx import Presentation

from handlers import get_all_handlers, HANDLER_REGISTRY
from utils.template import open_template, add_slide_from_layout, delete_all_original_slides

logger = logging.getLogger("uqslide.converter")

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
    logger.info("Starting conversion (%d bytes)", len(input_bytes))

    input_prs = Presentation(io.BytesIO(input_bytes))
    handlers = get_all_handlers()
    total_slides = len(input_prs.slides)
    logger.info("Loaded presentation: %d slides", total_slides)

    # Detect programme name from source footers
    programme_name = _detect_programme_name(input_prs)
    if programme_name:
        logger.info("Detected programme name: '%s'", programme_name)

    report = {
        "slides_converted": 0,
        "slides_flagged": 0,
        "slides_skipped": 0,
        "api_calls": 0,
        "errors": [],
        "details": [],
    }

    # --- Pre-render slides to images if API key provided ---
    slide_images = {}
    if api_key:
        if progress_callback:
            progress_callback("Rendering slides to images for AI classification...")
        slide_images = _render_slide_images(input_bytes)
        if slide_images:
            logger.info("Rendered %d slide images for Vision classification", len(slide_images))
            if progress_callback:
                progress_callback(
                    f"Rendered {len(slide_images)} slide images. Classifying..."
                )
        else:
            logger.warning("Slide rendering failed — falling back to text-only classification")
            report["errors"].append("Slide rendering failed. Used text-only classification fallback.")
            if progress_callback:
                progress_callback(
                    "Could not render slide images. Using text-only fallback."
                )

    output_prs = open_template()
    new_slides_added = 0

    # --- Auto-insert Acknowledgement of Country as slide 2 ---
    # Pre-scan source to find (and later skip) any existing AoC slides
    aoc_handler = handlers.get("Acknowledgement of Country")
    aoc_source_indices = set()  # source slide indices to skip (already consumed)
    if aoc_handler:
        for i, s in enumerate(input_prs.slides):
            if aoc_handler.detect(s, i) >= 0.9:
                aoc_source_indices.add(i)
        # AoC will be inserted after the cover slide is processed (see below)
    aoc_inserted = False

    for slide_idx, slide in enumerate(input_prs.slides):

        # Skip source AoC slides — we auto-insert a branded one
        if slide_idx in aoc_source_indices:
            logger.info("Slide %d: skipping existing AoC (will auto-insert branded version)", slide_idx + 1)
            report["details"].append({
                "slide": slide_idx + 1,
                "status": "converted",
                "handler": "Acknowledgement of Country",
                "confidence": 0.95,
                "content": None,
                "preview": "Acknowledgement of Country (replaced with branded version)",
                "all_scores": {},
                "classification_method": "auto",
            })
            continue

        if progress_callback:
            progress_callback(
                f"Processing slide {slide_idx + 1} of {total_slides}..."
            )

        # --- Step 1: Heuristic scoring ---
        scores = {}
        for name, handler in handlers.items():
            try:
                scores[name] = handler.detect(slide, slide_idx)
            except Exception as e:
                logger.error("Slide %d: handler '%s' detect() failed: %s", slide_idx + 1, name, e)
                scores[name] = 0.0

        best_name = max(scores, key=scores.get)
        best_confidence = scores[best_name]
        classification_method = "heuristic"
        logger.debug("Slide %d: heuristic → %s (%.2f) | scores=%s",
                      slide_idx + 1, best_name, best_confidence, scores)

        # --- Step 2: API fallback for low-confidence slides ---
        if best_confidence < CONFIDENT_THRESHOLD and api_key:
            logger.info("Slide %d: low confidence (%.2f), calling Vision API", slide_idx + 1, best_confidence)
            api_result = _classify_with_api(
                slide, slide_idx, total_slides, api_key, model,
                slide_image=slide_images.get(slide_idx),
            )

            if api_result and api_result.get("type"):
                report["api_calls"] += 1
                api_type = api_result["type"]
                api_confidence = api_result.get("confidence", 0.8)
                classification_method = "api"
                logger.info("Slide %d: API → %s (%.2f) reason=%s",
                            slide_idx + 1, api_type, api_confidence,
                            api_result.get("reason", ""))

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
            elif api_result and api_result.get("error"):
                err_msg = f"Slide {slide_idx + 1}: API error — {api_result['error']}"
                logger.error(err_msg)
                report["errors"].append(err_msg)

        # --- Step 3: Convert or skip ---
        best_handler = handlers.get(best_name)
        slide_preview = _get_slide_preview(slide)

        if best_handler and best_confidence >= CONVERT_THRESHOLD:
            try:
                content = best_handler.extract_content(slide, slide_idx)

                # Some handlers dynamically choose their layout based on content
                if hasattr(best_handler, 'get_layout_index'):
                    layout_idx = best_handler.get_layout_index(content)
                else:
                    layout_idx = best_handler.layout_index

                new_slide = add_slide_from_layout(output_prs, layout_idx)
                best_handler.fill_slide(new_slide, content)

                # Populate footer and slide number on every converted slide
                _fill_footer_and_slide_num(
                    new_slide, best_handler, programme_name, slide_idx + 1
                )

                new_slides_added += 1

                # Auto-insert AoC after cover slide
                if not aoc_inserted and best_handler.name == "Cover 1" and aoc_handler:
                    try:
                        aoc_content = aoc_handler._get_standard_content()
                        aoc_slide = add_slide_from_layout(output_prs, aoc_handler.layout_index)
                        aoc_handler.fill_slide(aoc_slide, aoc_content)
                        _fill_footer_and_slide_num(aoc_slide, aoc_handler, programme_name, 2)
                        new_slides_added += 1
                        aoc_inserted = True
                        report["slides_converted"] += 1
                        report["details"].append({
                            "slide": "2 (auto)",
                            "status": "converted",
                            "handler": "Acknowledgement of Country",
                            "confidence": 1.0,
                            "content": aoc_content,
                            "preview": "Acknowledgement of Country (auto-inserted)",
                            "all_scores": {},
                            "classification_method": "auto",
                        })
                        logger.info("Auto-inserted Acknowledgement of Country as slide 2")
                    except Exception as e:
                        logger.error("Failed to insert AoC slide: %s", e)
                        report["errors"].append(f"AoC auto-insert failed: {e}")

                if best_confidence >= CONFIDENT_THRESHOLD:
                    status = "converted"
                else:
                    status = "flagged"
                    report["slides_flagged"] += 1

                report["slides_converted"] += 1
                logger.info("Slide %d: %s → %s (%.2f, %s)",
                            slide_idx + 1, status, best_handler.name,
                            best_confidence, classification_method)
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
            except Exception as e:
                err_msg = f"Slide {slide_idx + 1}: conversion failed ({best_handler.name}) — {e}"
                logger.error(err_msg, exc_info=True)
                report["errors"].append(err_msg)
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
                    "reason": f"Error: {e}",
                })
        else:
            logger.info("Slide %d: skipped (best='%s' at %.2f)",
                        slide_idx + 1, best_name, best_confidence)
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

    logger.info("Conversion complete: %d converted, %d flagged, %d skipped, %d API calls, %d errors",
                report["slides_converted"], report["slides_flagged"],
                report["slides_skipped"], report["api_calls"], len(report["errors"]))

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
    except Exception as e:
        logger.error("Slide rendering failed: %s", e, exc_info=True)
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
    except Exception as e:
        logger.error("API classifier failed for slide %d: %s", slide_index + 1, e, exc_info=True)
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


def _detect_programme_name(prs) -> str:
    """
    Detect the programme/course name for footer text.

    Strategy:
    1. Check footer placeholders (idx 17) on slides 2-5 for an existing
       programme name like "Executive Education" or "Negotiating For Success".
    2. Fall back to the cover slide title (largest text on slide 1).
    """
    if len(prs.slides) == 0:
        return ""

    # Strategy 1: Look for footer placeholder text on early slides
    for i in range(1, min(6, len(prs.slides))):
        slide = prs.slides[i]
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            if (hasattr(shape, 'is_placeholder') and shape.is_placeholder
                    and shape.placeholder_format.idx == 17):
                text = shape.text_frame.text.strip()
                if text and len(text) < 80:
                    return text

    # Strategy 2: Cover slide title (largest font)
    cover = prs.slides[0]

    # Check placeholder idx 0 first
    for shape in cover.shapes:
        if not shape.has_text_frame:
            continue
        if (hasattr(shape, 'is_placeholder') and shape.is_placeholder
                and shape.placeholder_format.idx in (0, 3, 15)):
            text = shape.text_frame.text.strip()
            if text and len(text) < 120:
                return text

    # Check shape named "Title" (some slides have non-placeholder titles)
    for shape in cover.shapes:
        if not shape.has_text_frame:
            continue
        if "title" in shape.name.lower():
            text = shape.text_frame.text.strip()
            if text and len(text) < 120:
                return text

    # Fallback: largest font on the first slide
    best_text = ""
    best_size = 0
    for shape in cover.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if run.font.size and run.font.size.pt > best_size:
                    best_size = run.font.size.pt
                    best_text = shape.text_frame.text.strip()

    return best_text if len(best_text) < 120 else ""


def _fill_footer_and_slide_num(slide, handler, programme_name: str, slide_number: int):
    """
    Populate footer and slide number placeholders on a converted slide.
    Uses the handler's placeholder map to find the right indices.
    """
    ph_map = handler.get_placeholder_map()
    placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}

    # Footer (programme name) — always override with detected programme name
    footer_idx = ph_map.get("footer")
    if footer_idx is not None and footer_idx in placeholders and programme_name:
        placeholders[footer_idx].text = programme_name

    # Slide number
    slide_num_idx = ph_map.get("slide_num")
    if slide_num_idx is not None and slide_num_idx in placeholders:
        placeholders[slide_num_idx].text = str(slide_number)


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
