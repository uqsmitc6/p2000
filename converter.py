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
) -> tuple[bytes, bytes | None, dict]:
    """
    Convert an uploaded PPTX to brand-compliant format.

    Args:
        input_bytes: Raw bytes of the uploaded .pptx file
        api_key: Optional Anthropic API key. If provided, Claude Vision
                 classifies ambiguous slides using rendered images.
        model: Claude model for classification (default: claude-sonnet-4-6)
        progress_callback: Optional callable(message: str) for status updates

    Returns:
        (output_bytes, review_bytes_or_None, report)
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
        slide_images, render_diag = _render_slide_images(input_bytes)
        if slide_images:
            logger.info("Rendered %d slide images for Vision classification", len(slide_images))
            if progress_callback:
                progress_callback(
                    f"Rendered {len(slide_images)} slide images. Classifying..."
                )
        else:
            logger.warning("Slide rendering failed — falling back to text-only classification")
            report["errors"].append(
                f"Slide rendering failed: {render_diag}. Used text-only classification fallback."
            )
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

    # --- Pre-scan for Thank You: only allow ONE closing slide ---
    # Find the last slide that looks like a Thank You (heuristic or position)
    # so we can reserve the Thank You layout for only that one.
    last_thank_you_idx = None
    ty_handler = handlers.get("Thank You")
    if ty_handler:
        for i in range(len(input_prs.slides) - 1, max(len(input_prs.slides) - 5, -1), -1):
            if i in aoc_source_indices:
                continue
            s = input_prs.slides[i]
            ty_score = ty_handler.detect(s, i)
            if ty_score >= 0.5:
                last_thank_you_idx = i
                break
        # If heuristics didn't find one, the last slide is the fallback
        if last_thank_you_idx is None:
            last_thank_you_idx = len(input_prs.slides) - 1
    logger.info("Thank You reserved for source slide %d", (last_thank_you_idx or -1) + 1)

    for slide_idx, slide in enumerate(input_prs.slides):

        # Skip source AoC slides — we auto-insert a branded one
        if slide_idx in aoc_source_indices:
            logger.info("Slide %d: skipping existing AoC (will auto-insert branded version)", slide_idx + 1)
            report["details"].append({
                "slide": slide_idx + 1,
                "status": "replaced",  # Not "converted" — no output slide for this source
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
                    # API says skip — but if the heuristic found a viable
                    # match above CONVERT_THRESHOLD, trust the heuristic.
                    # This is especially important when slide rendering failed
                    # (text-only fallback) and the slide has images but no text.
                    if best_confidence >= CONVERT_THRESHOLD:
                        logger.info("Slide %d: API said Skip but heuristic has %s at %.2f — keeping heuristic",
                                    slide_idx + 1, best_name, best_confidence)
                        # Don't skip — fall through to Step 3 with heuristic result
                    else:
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
                    # Guard: only allow Thank You for the reserved last slot
                    if api_type == "Thank You" and slide_idx != last_thank_you_idx:
                        logger.info("Slide %d: API suggested Thank You but reserved for slide %d — using Title and Content",
                                    slide_idx + 1, (last_thank_you_idx or -1) + 1)
                        best_name = "Title and Content"
                        best_confidence = api_confidence
                    else:
                        # API classified it — use that handler
                        best_name = api_type
                        best_confidence = api_confidence
            elif api_result and api_result.get("error"):
                err_msg = f"Slide {slide_idx + 1}: API error — {api_result['error']}"
                logger.error(err_msg)
                report["errors"].append(err_msg)

        # --- Guard: Thank You only for the reserved last slot ---
        if best_name == "Thank You" and slide_idx != last_thank_you_idx:
            logger.info("Slide %d: downgrading Thank You to Title and Content (reserved for slide %d)",
                        slide_idx + 1, (last_thank_you_idx or -1) + 1)
            best_name = "Title and Content"
            best_confidence = scores.get("Title and Content", 0.40)

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

                # Remove unused subtitle placeholders (removes dashed boxes)
                _cleanup_empty_subtitle(new_slide, best_handler, content)

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
    output_bytes = output_buffer.getvalue()

    logger.info("Conversion complete: %d converted, %d flagged, %d skipped, %d API calls, %d errors",
                report["slides_converted"], report["slides_flagged"],
                report["slides_skipped"], report["api_calls"], len(report["errors"]))

    # --- Post-conversion verification ---
    # Send original + branded slide pairs to Claude for QA checking.
    # Only runs if we have an API key AND source slide images were rendered.
    if api_key and slide_images:
        if progress_callback:
            progress_callback("Verifying converted slides...")

        verification_results = _verify_all_slides(
            input_bytes, output_bytes, slide_images, report, api_key, model,
            progress_callback,
        )
        report["verification"] = verification_results

        # Count issues
        v_passed = sum(1 for v in verification_results if v.get("pass") is True)
        v_issues = sum(1 for v in verification_results if v.get("pass") is False)
        v_errors = sum(1 for v in verification_results if v.get("pass") is None)
        report["verification_summary"] = {
            "total": len(verification_results),
            "passed": v_passed,
            "issues_found": v_issues,
            "errors": v_errors,
        }
        report["api_calls"] += len(verification_results)

        logger.info("Verification: %d passed, %d issues found, %d errors",
                     v_passed, v_issues, v_errors)

        if progress_callback:
            progress_callback(
                f"Verification complete: {v_passed} passed, {v_issues} issues found"
            )

    # --- Build review copy (interleaved original + branded) ---
    if progress_callback:
        progress_callback("Building review copy...")
    try:
        review_bytes = _build_review_copy(input_bytes, output_bytes, report)
    except Exception as e:
        logger.error("Failed to build review copy: %s", e, exc_info=True)
        report["errors"].append(f"Review copy failed: {e}")
        review_bytes = None

    return output_bytes, review_bytes, report


def _verify_all_slides(
    input_bytes: bytes,
    output_bytes: bytes,
    source_images: dict,
    report: dict,
    api_key: str,
    model: str,
    progress_callback=None,
) -> list:
    """
    Verify all converted slides by comparing source and output images.

    Args:
        input_bytes: original PPTX bytes
        output_bytes: branded PPTX bytes
        source_images: dict mapping source slide_index → PNG bytes
        report: conversion report (to read slide details)
        api_key: Anthropic API key
        model: Claude model to use
        progress_callback: optional status callback

    Returns:
        List of verification results, one per converted slide.
    """
    from utils.classifier import verify_slide_pair

    # Render the output slides
    output_images, out_diag = _render_slide_images(output_bytes)
    if not output_images:
        logger.warning("Could not render output slides for verification: %s", out_diag)
        return []

    total_output = len(output_images)
    results = []

    # Build a mapping from output slide index to source slide index + handler
    # The report["details"] list tracks which source slides were converted
    # and in what order they appear in the output.
    output_idx = 0
    for detail in report["details"]:
        if detail["status"] in ("converted", "flagged"):
            source_slide_num = detail["slide"]
            handler_name = detail.get("handler", "Unknown")

            # Source image: keyed by 0-based source index
            if isinstance(source_slide_num, int):
                source_idx = source_slide_num - 1
            else:
                # Auto-inserted slides (e.g. "2 (auto)") — skip verification
                output_idx += 1
                continue

            source_img = source_images.get(source_idx)
            output_img = output_images.get(output_idx)

            if source_img and output_img:
                if progress_callback:
                    progress_callback(
                        f"Verifying slide {output_idx + 1}/{total_output} "
                        f"(source slide {source_slide_num})..."
                    )

                v_result = verify_slide_pair(
                    source_image=source_img,
                    output_image=output_img,
                    slide_number=source_slide_num,
                    total_slides=total_output,
                    handler_name=handler_name,
                    api_key=api_key,
                    model=model,
                )
                v_result["source_slide"] = source_slide_num
                v_result["output_slide"] = output_idx + 1
                v_result["handler"] = handler_name
                results.append(v_result)

                # Attach to report detail for UI display
                detail["verification"] = v_result
            else:
                logger.debug("Skipping verification for source %d / output %d (missing image)",
                             source_slide_num, output_idx + 1)

            output_idx += 1

    return results


def _render_slide_images(input_bytes: bytes) -> tuple[dict, str]:
    """
    Render all slides to PNG images using LibreOffice.

    Returns:
        Tuple of (dict mapping slide_index → PNG bytes, diagnostic_message)
        Empty dict if rendering fails.
    """
    try:
        from utils.renderer import render_slides_to_images, is_libreoffice_available
        if not is_libreoffice_available():
            msg = "LibreOffice not available — check Docker image has libreoffice-impress installed"
            logger.error(msg)
            return {}, msg
        logger.info("LibreOffice available, starting render of %d bytes...", len(input_bytes))
        images, diag = render_slides_to_images(input_bytes, dpi=150)
        logger.info("Render complete: %d slide images produced. Diag: %s", len(images), diag)
        return {i: img for i, img in enumerate(images)}, diag
    except Exception as e:
        msg = f"Slide rendering failed: {e}"
        logger.error(msg, exc_info=True)
        return {}, msg


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


def _cleanup_empty_subtitle(slide, handler, content: dict):
    """
    Remove subtitle placeholder if the handler didn't populate it.
    Prevents the dashed template prompt box from showing on output slides.
    """
    ph_map = handler.get_placeholder_map()
    subtitle_idx = ph_map.get("subtitle")
    if subtitle_idx is None:
        return  # Handler has no subtitle placeholder

    # Check if content actually provided a subtitle
    if content and content.get("subtitle"):
        return  # Subtitle was filled — keep it

    # Remove the empty placeholder element from the slide XML
    placeholders = {ph.placeholder_format.idx: ph for ph in slide.placeholders}
    if subtitle_idx in placeholders:
        sp = placeholders[subtitle_idx]._element
        sp.getparent().remove(sp)


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


def _build_review_copy(
    input_bytes: bytes,
    output_bytes: bytes,
    report: dict,
) -> bytes | None:
    """
    Build an interleaved review PPTX: original slide → branded slide, repeated.

    This lets a human reviewer quickly flip through to compare each source
    slide against its branded counterpart.

    Strategy:
        1. Open both source and output presentations.
        2. Walk the report details to find converted/flagged slides.
        3. For each, copy the source slide then the output slide into a
           new blank presentation.

    Returns:
        PPTX bytes for the review copy, or None if building fails.
    """
    from pptx.util import Inches, Pt, Emu
    from copy import deepcopy
    from lxml import etree

    source_prs = Presentation(io.BytesIO(input_bytes))
    output_prs = Presentation(io.BytesIO(output_bytes))
    review_prs = Presentation()

    # Match slide dimensions to the output
    review_prs.slide_width = output_prs.slide_width
    review_prs.slide_height = output_prs.slide_height

    # Use a blank layout from the review presentation
    blank_layout = review_prs.slide_layouts[6]  # Typically "Blank"

    # Track which output slide index corresponds to each source slide
    output_idx = 0
    source_slides = list(source_prs.slides)
    output_slides = list(output_prs.slides)

    for detail in report["details"]:
        if detail["status"] not in ("converted", "flagged"):
            continue

        source_slide_num = detail["slide"]
        handler_name = detail.get("handler", "?")

        # Auto-inserted slides (e.g. AoC "2 (auto)") have no source original,
        # but still occupy an output slot. Include them branded-only.
        if not isinstance(source_slide_num, int):
            if output_idx < len(output_slides):
                _copy_slide_to(output_prs, output_slides[output_idx], review_prs, blank_layout,
                               label=f"AUTO-INSERTED — {handler_name}")
            output_idx += 1
            continue

        source_idx = source_slide_num - 1

        # --- Add source slide ---
        if source_idx < len(source_slides):
            _copy_slide_to(source_prs, source_slides[source_idx], review_prs, blank_layout,
                           label=f"ORIGINAL — Slide {source_slide_num}")

        # --- Add branded slide ---
        if output_idx < len(output_slides):
            _copy_slide_to(output_prs, output_slides[output_idx], review_prs, blank_layout,
                           label=f"BRANDED — Slide {source_slide_num} → {handler_name}")

        output_idx += 1

    if len(review_prs.slides) == 0:
        logger.warning("Review copy has no slides — skipping")
        return None

    buf = io.BytesIO()
    review_prs.save(buf)
    buf.seek(0)
    logger.info("Review copy built: %d slides", len(review_prs.slides))
    return buf.getvalue()


def _copy_slide_to(source_prs, source_slide, target_prs, blank_layout, label: str = ""):
    """
    Copy all shapes from a source slide into a new blank slide in the target
    presentation. Also adds a small label in the top-left corner.

    This is a simplified copy — it transfers shapes via XML cloning but
    does not copy slide backgrounds, master styles, or media relationships
    that live outside the shape tree. For review purposes this is sufficient.
    """
    from pptx.util import Pt, Emu, Inches
    from lxml import etree
    import copy as _copy

    new_slide = target_prs.slides.add_slide(blank_layout)

    # Copy the slide background if it has one
    src_bg = source_slide.background
    if src_bg is not None and src_bg._element is not None:
        bg_elem = src_bg._element
        # Check if it has fill (not just inherited)
        if bg_elem.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill') is not None or \
           bg_elem.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}gradFill') is not None:
            new_bg = _copy.deepcopy(bg_elem)
            new_slide.background._element.getparent().replace(
                new_slide.background._element, new_bg
            )

    # Copy shapes from the source slide's shape tree
    for shape in source_slide.shapes:
        try:
            el = _copy.deepcopy(shape._element)
            new_slide.shapes._spTree.append(el)

            # If the shape references an image, copy the image relationship
            blip_elems = el.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
            for blip in blip_elems:
                embed_attr = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                if embed_attr and embed_attr in source_slide.part.rels:
                    try:
                        src_rel = source_slide.part.rels[embed_attr]
                        src_part = src_rel.target_part
                        new_rid = new_slide.part.relate_to(src_part, src_rel.reltype)
                        blip.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', new_rid)
                    except Exception:
                        pass  # Image won't render but shape still copies
        except Exception as e:
            logger.debug("Could not copy shape '%s': %s", getattr(shape, 'name', '?'), e)

    # Add a small label textbox in the top-left corner
    if label:
        from pptx.util import Pt, Inches
        txBox = new_slide.shapes.add_textbox(
            Inches(0.2), Inches(0.05), Inches(5), Inches(0.3)
        )
        tf = txBox.text_frame
        tf.text = label
        for para in tf.paragraphs:
            for run in para.runs:
                run.font.size = Pt(10)
                run.font.bold = True
                from pptx.dml.color import RGBColor
                if "ORIGINAL" in label:
                    run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)  # Red
                elif "AUTO-INSERTED" in label:
                    run.font.color.rgb = RGBColor(0x00, 0x66, 0x99)  # Teal blue
                else:
                    run.font.color.rgb = RGBColor(0x51, 0x24, 0x7A)  # UQ Purple


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
