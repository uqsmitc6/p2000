"""
Slide-to-image renderer using LibreOffice headless.

Converts a PPTX file to PNG images (one per slide) using LibreOffice's
built-in export. This gives pixel-perfect rendering of themes, gradients,
fonts, and layout — exactly what Claude Vision needs for classification.

Images are written to a persistent directory on disk rather than held in
memory, to support large presentations (200MB+, 100+ slides).

Requires LibreOffice to be installed (available via system package or Docker).
"""

import io
import os
import glob
import logging
import shutil
import tempfile
import subprocess
from pathlib import Path

logger = logging.getLogger("uqslide.renderer")

# Persistent render cache directory. Uses Azure /home (persists across
# restarts) if available, otherwise falls back to /tmp.
_RENDER_BASE = None


def _get_render_base() -> str:
    """Return the base directory for rendered slide images."""
    global _RENDER_BASE
    if _RENDER_BASE is not None:
        return _RENDER_BASE

    candidates = [
        "/home/data/uq-slide-converter/renders",
        "/tmp/uq-slide-converter/renders",
    ]
    for path in candidates:
        try:
            os.makedirs(path, exist_ok=True)
            test_file = os.path.join(path, ".write_test")
            with open(test_file, "w") as f:
                f.write("ok")
            os.remove(test_file)
            _RENDER_BASE = path
            logger.info("Render base: %s", path)
            return path
        except OSError:
            continue

    # Last resort
    _RENDER_BASE = tempfile.mkdtemp(prefix="uqslide_renders_")
    logger.warning("Render base (fallback tmpdir): %s", _RENDER_BASE)
    return _RENDER_BASE


def render_slides_to_dir(
    input_bytes: bytes,
    dpi: int = 96,
    timeout: int = 120,
    session_id: str = None,
) -> tuple[str | None, int, str]:
    """
    Render all slides in a PPTX to PNG images on disk.

    This is the preferred entry point for large-file support. Images are
    written to a session-specific subdirectory and loaded on demand by
    the caller, avoiding holding hundreds of MBs in RAM.

    Args:
        input_bytes: Raw bytes of the .pptx file
        dpi: Resolution for export (96 = default, 72 = low-memory mode)
        timeout: Max seconds for LibreOffice conversion
        session_id: Optional unique session ID for the render directory.
                    If None, uses a timestamp-based name.

    Returns:
        Tuple of (render_dir, num_images, diagnostic_message).
        render_dir is None if rendering failed entirely.
        PNG files are named slide_000.png, slide_001.png, etc.
    """
    import time

    if session_id is None:
        session_id = f"render_{int(time.time() * 1000)}"

    render_dir = os.path.join(_get_render_base(), session_id)
    os.makedirs(render_dir, exist_ok=True)

    # Adaptive settings based on file size
    file_mb = len(input_bytes) / (1024 * 1024)
    if file_mb > 100:
        dpi = min(dpi, 72)
        timeout = max(timeout, 300)
        logger.info("Large file (%.0f MB): using low-DPI (%d) and extended timeout (%ds)",
                     file_mb, dpi, timeout)
    elif file_mb > 30:
        dpi = min(dpi, 96)
        timeout = max(timeout, 180)
        logger.info("Medium file (%.0f MB): timeout=%ds", file_mb, timeout)

    tmpdir = tempfile.mkdtemp(prefix="uqslide_lo_")
    diag = ""

    try:
        # Check available disk space
        disk = shutil.disk_usage(tmpdir)
        free_mb = disk.free / (1024 * 1024)
        logger.info("Render tmpdir: %s — free disk: %.0f MB", tmpdir, free_mb)
        if free_mb < 200:
            diag = f"Low disk space: {free_mb:.0f} MB free"
            logger.warning(diag)

        # Write input PPTX to temp file
        input_path = os.path.join(tmpdir, "input.pptx")
        with open(input_path, "wb") as f:
            f.write(input_bytes)
        logger.info("Wrote input PPTX: %.1f MB to %s", file_mb, input_path)

        # Try PDF route first (more reliable for per-slide output)
        num_images, pdf_diag = _render_via_pdf_to_dir(
            input_path, tmpdir, render_dir, dpi, timeout
        )
        if num_images > 0:
            diag = pdf_diag or f"Rendered {num_images} slide images via PDF route"
            return render_dir, num_images, diag

        # Fallback: direct PNG export (may produce only 1 image)
        num_images, png_diag = _render_direct_png_to_dir(
            input_path, tmpdir, render_dir, timeout
        )
        if num_images > 0:
            diag = png_diag or f"Rendered {num_images} slide images via direct PNG"
            return render_dir, num_images, diag

        diag = (pdf_diag or "") + " | " + (png_diag or "")
        logger.error("All render methods failed: %s", diag)
        return None, 0, diag

    except subprocess.TimeoutExpired:
        diag = f"LibreOffice timed out after {timeout} seconds"
        logger.error(diag)
        return None, 0, diag
    except FileNotFoundError:
        diag = "LibreOffice not found — is it installed?"
        logger.error(diag)
        return None, 0, diag
    except Exception as e:
        diag = f"Unexpected rendering error: {e}"
        logger.error(diag, exc_info=True)
        return None, 0, diag
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
        import gc
        gc.collect()


def _render_via_pdf_to_dir(
    input_path: str,
    tmpdir: str,
    render_dir: str,
    dpi: int,
    timeout: int,
) -> tuple[int, str]:
    """
    PPTX → PDF → per-slide PNGs written to render_dir.
    Returns (num_images, diagnostic_message).
    """
    pdf_dir = os.path.join(tmpdir, "pdf_out")
    os.makedirs(pdf_dir, exist_ok=True)

    env = os.environ.copy()
    env["HOME"] = tmpdir
    env["SAL_USE_VCLPLUGIN"] = "svp"

    # Step 1: PPTX → PDF
    cmd_pdf = [
        "libreoffice", "--headless", "--norestore",
        "--convert-to", "pdf",
        "--outdir", pdf_dir,
        input_path,
    ]

    logger.info("Running LibreOffice PDF export...")
    result = subprocess.run(
        cmd_pdf, capture_output=True, text=True, timeout=timeout, env=env
    )

    if result.returncode != 0:
        logger.warning("LibreOffice PDF export failed (code %d): %s",
                        result.returncode, (result.stderr or "")[:300])

    pdf_files = glob.glob(os.path.join(pdf_dir, "*.pdf"))
    if not pdf_files:
        return 0, f"PDF conversion produced no output (code {result.returncode})"

    pdf_path = pdf_files[0]
    pdf_size = os.path.getsize(pdf_path)
    logger.info("PDF created: %.1f MB", pdf_size / (1024 * 1024))

    # Step 2: PDF → PNGs via pdftoppm (best quality, writes to disk directly)
    try:
        cmd_png = [
            "pdftoppm", "-png",
            "-r", str(dpi),
            pdf_path,
            os.path.join(render_dir, "slide"),
        ]

        logger.info("Running pdftoppm (DPI=%d)...", dpi)
        ppm_result = subprocess.run(
            cmd_png, capture_output=True, text=True, timeout=timeout
        )

        if ppm_result.stderr:
            logger.info("pdftoppm stderr: %s", ppm_result.stderr[:300])

        # pdftoppm outputs slide-01.png, slide-02.png, etc.
        raw_files = sorted(glob.glob(os.path.join(render_dir, "slide-*.png")))
        if raw_files:
            # Rename to zero-indexed: slide_000.png, slide_001.png, ...
            for i, src in enumerate(raw_files):
                dst = os.path.join(render_dir, f"slide_{i:03d}.png")
                os.rename(src, dst)
            logger.info("pdftoppm produced %d slide images", len(raw_files))
            return len(raw_files), f"PDF route: {len(raw_files)} images via pdftoppm"
        else:
            logger.warning("pdftoppm produced no PNGs")
    except FileNotFoundError:
        logger.warning("pdftoppm not found")

    # Fallback: pdf2image Python library
    try:
        from pdf2image import convert_from_path
        pil_images = convert_from_path(pdf_path, dpi=dpi)
        for i, img in enumerate(pil_images):
            out_path = os.path.join(render_dir, f"slide_{i:03d}.png")
            img.save(out_path, format="PNG")
            img.close()  # Free memory immediately
        logger.info("pdf2image produced %d slide images", len(pil_images))
        return len(pil_images), f"PDF route: {len(pil_images)} images via pdf2image"
    except ImportError:
        logger.warning("pdf2image not available")
    except Exception as e:
        logger.warning("pdf2image failed: %s", e)

    return 0, "All PDF-to-PNG methods failed"


def _render_direct_png_to_dir(
    input_path: str,
    tmpdir: str,
    render_dir: str,
    timeout: int,
) -> tuple[int, str]:
    """
    Direct PPTX → PNG via LibreOffice. Fallback when PDF route fails.
    Returns (num_images, diagnostic_message).
    """
    env = os.environ.copy()
    env["HOME"] = tmpdir
    env["SAL_USE_VCLPLUGIN"] = "svp"

    cmd = [
        "libreoffice", "--headless", "--norestore",
        "--convert-to", "png",
        "--outdir", tmpdir,
        input_path,
    ]

    logger.info("Running LibreOffice direct PNG export...")
    result = subprocess.run(
        cmd, capture_output=True, text=True, timeout=timeout, env=env
    )

    png_files = sorted(glob.glob(os.path.join(tmpdir, "*.png")))
    if not png_files:
        return 0, f"Direct PNG export produced no files (code {result.returncode})"

    # Move to render_dir with standard naming
    for i, src in enumerate(png_files):
        dst = os.path.join(render_dir, f"slide_{i:03d}.png")
        shutil.move(src, dst)

    logger.info("Direct PNG export produced %d images", len(png_files))
    return len(png_files), f"Direct PNG: {len(png_files)} images"


# --- Legacy API (backward compatible) ---

def render_slides_to_images(
    input_bytes: bytes,
    dpi: int = 96,
    timeout: int = 120,
) -> tuple[list[bytes], str]:
    """
    Render all slides in a PPTX to PNG images (in memory).

    DEPRECATED for large files — use render_slides_to_dir() instead.
    Kept for backward compatibility with small files.

    Returns:
        Tuple of (list of PNG image bytes, diagnostic_message).
        Returns (empty list, error_message) if rendering failed.
    """
    render_dir, num_images, diag = render_slides_to_dir(
        input_bytes, dpi=dpi, timeout=timeout
    )
    if render_dir is None or num_images == 0:
        return [], diag

    images = []
    for i in range(num_images):
        path = os.path.join(render_dir, f"slide_{i:03d}.png")
        try:
            with open(path, "rb") as f:
                images.append(f.read())
        except FileNotFoundError:
            pass

    # Clean up render dir
    shutil.rmtree(render_dir, ignore_errors=True)

    return images, diag


def load_slide_image(render_dir: str, slide_index: int) -> bytes | None:
    """
    Load a single slide image from a render directory.

    This is the memory-efficient way to access rendered slides — load
    one at a time instead of all at once.

    Args:
        render_dir: Path returned by render_slides_to_dir()
        slide_index: Zero-based slide index

    Returns:
        PNG bytes, or None if the image doesn't exist.
    """
    path = os.path.join(render_dir, f"slide_{slide_index:03d}.png")
    try:
        with open(path, "rb") as f:
            return f.read()
    except FileNotFoundError:
        return None


def cleanup_render_dir(render_dir: str) -> None:
    """Remove a render directory and all its contents."""
    if render_dir and os.path.isdir(render_dir):
        shutil.rmtree(render_dir, ignore_errors=True)
        logger.info("Cleaned up render dir: %s", render_dir)


def is_libreoffice_available() -> bool:
    """Check if LibreOffice is installed and available."""
    try:
        result = subprocess.run(
            ["libreoffice", "--version"],
            capture_output=True,
            text=True,
            timeout=10,
        )
        if result.returncode == 0:
            logger.info("LibreOffice: %s", result.stdout.strip())
            return True
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass
    return False
