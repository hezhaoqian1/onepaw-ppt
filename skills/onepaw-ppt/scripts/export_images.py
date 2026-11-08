#!/usr/bin/env python3
"""Export a PPTD project as page images through the local neo-ppt editor for visual QA.

The editor runs in a headless Chrome/Edge driven over the DevTools protocol
(scripts/cdp.py, standard library only): the script picks 图片 in the editor's
export dialog, unzips the downloaded images, and lets the same browser stitch
all pages into one overview image that a multimodal model can review.

Editor web fonts are rendered with the system fonts the PPTX export swaps in
for them (export_pptx.FONT_FALLBACKS), so the preview matches the PPTX that
PowerPoint/WPS will show. Pass --keep-fonts to preview the web fonts instead.
"""

from __future__ import annotations

import argparse
import html
import json
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

# Embedded Pythons configured by a ._pth file (e.g. the one bundled with desktop
# agents such as OnePaw) do not put the script's directory on sys.path.
sys.path.insert(0, str(Path(__file__).resolve().parent))

from cdp import Browser, CDPError, Page  # noqa: E402
from export_pptx import (  # noqa: E402,F401  (is_image_zip is part of this module's API)
    EditorHost,
    ExportError,
    build_payload,
    find_manifest,
    is_image_zip,
    log,
    make_output_robust,
    temporary_directory,
)

IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".webp"}
OVERVIEW_COLUMNS = 3
OVERVIEW_THUMB_WIDTH = 640
OVERVIEW_LABEL_HEIGHT = 32
OVERVIEW_GAP = 12


def page_sort_key(path: Path) -> Tuple[int, str]:
    match = re.match(r"(\d+)", path.stem)
    return (int(match.group(1)) if match else sys.maxsize, path.name)


def unzip_images(archive_path: Path, pages_dir: Path) -> List[Path]:
    pages_dir.mkdir(parents=True, exist_ok=True)
    images: List[Path] = []
    with zipfile.ZipFile(archive_path) as archive:
        for info in archive.infolist():
            if info.is_dir() or Path(info.filename).suffix.lower() not in IMAGE_SUFFIXES:
                continue
            name = Path(info.filename).name
            if not name:
                continue
            target = pages_dir / name
            with archive.open(info) as source, target.open("wb") as out:
                shutil.copyfileobj(source, out)
            images.append(target)
    images.sort(key=page_sort_key)
    if not images:
        raise ExportError(f"no page images found in: {archive_path}")
    return images


def overview_html(images: Sequence[Path], base: Path) -> str:
    """A grid of labelled thumbnails (P1, P2, ...) referencing files under ``base``."""
    cells = []
    for index, path in enumerate(images, start=1):
        source = html.escape(path.relative_to(base).as_posix(), quote=True)
        cells.append(
            f'<figure><figcaption>P{index}</figcaption><img src="{source}" alt="P{index}"></figure>'
        )
    return (
        "<!doctype html><html><head><meta charset='utf-8'><style>"
        "html,body{margin:0;background:#e5e7eb}"
        f"main{{display:grid;grid-template-columns:repeat({OVERVIEW_COLUMNS},{OVERVIEW_THUMB_WIDTH}px);"
        f"gap:{OVERVIEW_GAP}px;padding:{OVERVIEW_GAP}px;width:max-content}}"
        "figure{margin:0}"
        f"figcaption{{height:{OVERVIEW_LABEL_HEIGHT - 4}px;margin-bottom:4px;background:#111827;"
        "color:#fff;font:600 18px/28px Arial,sans-serif;padding:0 8px}"
        f"img{{display:block;width:{OVERVIEW_THUMB_WIDTH}px;height:auto}}"
        "</style></head><body><main>" + "".join(cells) + "</main></body></html>"
    )


IMAGES_LOADED_JS = (
    "document.readyState === 'complete' && "
    "[...document.images].every((img) => img.complete && img.naturalWidth > 0)"
)
GRID_SIZE_JS = (
    "(() => { const r = document.querySelector('main').getBoundingClientRect();"
    " return [Math.ceil(r.width), Math.ceil(r.height)]; })()"
)


def stitch_overview(browser: Browser, images: Sequence[Path], output: Path) -> Path:
    """Render all page images into one JPEG grid with the headless browser."""
    base = output.parent
    document = base / f".{output.stem}.html"
    document.write_text(overview_html(images, base), encoding="utf-8")
    try:
        page = Page(browser, width=1280, height=720)
        page.navigate(document.resolve().as_uri())
        page.wait_for(IMAGES_LOADED_JS, 60, "the overview images to load")
        width, height = page.evaluate(GRID_SIZE_JS)
        page.call(
            "Emulation.setDeviceMetricsOverride",
            {"width": width, "height": height, "deviceScaleFactor": 1, "mobile": False},
        )
        page.screenshot(output, quality=85)
    finally:
        document.unlink(missing_ok=True)
    return output


def export_images(
    source: Path,
    output: Path,
    keep_download: bool = False,
    force: bool = False,
    map_fonts: bool = True,
) -> Dict[str, Any]:
    manifest = find_manifest(source)
    payload = build_payload(manifest)
    output = output.expanduser().resolve()
    if output.exists() and any(output.iterdir()) and not force:
        raise ExportError(
            f"output directory already exists (pass --force to replace it): {output}"
        )

    log(f"manifest: {manifest}")
    with temporary_directory(prefix="onepaw-ppt-images-") as temp_name:
        download_dir = Path(temp_name) / "downloads"
        download_dir.mkdir()
        with EditorHost(payload, font_fallbacks=map_fonts) as host:
            downloaded = host.export(download_dir, images=True, timeout=240)
            if output.exists():
                shutil.rmtree(output)
            output.mkdir(parents=True)
            images = unzip_images(downloaded, output / "pages")
            if keep_download:
                shutil.copy2(downloaded, output / "browser-raw.zip")
            if host.browser is None:
                raise ExportError("the headless browser closed unexpectedly")
            overview = stitch_overview(host.browser, images, output / "overview.jpg")

    page_paths = [entry["path"] for entry in payload["pages"]]
    mapping = [
        {
            "index": index,
            "image": f"pages/{path.name}",
            "page": page_paths[index - 1] if index - 1 < len(page_paths) else None,
        }
        for index, path in enumerate(images, start=1)
    ]
    return {
        "pages": len(images),
        "overview": str(overview),
        "output": str(output),
        "images": mapping,
        "exporter": "browser-local-editor",
        "fonts": "system-font substitutes, as in the PPTX" if map_fonts else "editor web fonts",
    }


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Export a PPTD project as page images via the local neo-ppt editor, unzip "
            "them, and stitch an overview image for visual QA."
        )
    )
    parser.add_argument("input", type=Path, help=".pptd manifest or project directory")
    parser.add_argument(
        "--output",
        "-o",
        type=Path,
        help="output directory (default: <project>/.qa-images)",
    )
    parser.add_argument(
        "--keep-browser-raw",
        action="store_true",
        help="also keep the downloaded images ZIP beside the overview",
    )
    parser.add_argument(
        "--keep-fonts",
        action="store_true",
        help="preview the editor web fonts instead of the PPTX system-font substitutes",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="replace an existing output directory",
    )
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    make_output_robust()
    args = parse_args(argv)
    try:
        manifest = find_manifest(args.input)
        output = args.output or manifest.parent / ".qa-images"
        summary = export_images(
            args.input, output, args.keep_browser_raw, args.force, map_fonts=not args.keep_fonts
        )
    except (ExportError, CDPError, OSError, subprocess.SubprocessError) as exc:
        print(f"onepaw-ppt image export failed: {exc}", file=sys.stderr)
        return 1
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
