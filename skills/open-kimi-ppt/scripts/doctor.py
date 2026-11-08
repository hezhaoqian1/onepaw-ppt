#!/usr/bin/env python3
"""Check that this computer can run the open-kimi-ppt skill, then prove it.

Usage:
    python scripts/doctor.py           # checks + a tiny PPTX export and image QA run
    python scripts/doctor.py --quick   # checks only (no test export)
    python scripts/doctor.py --json    # machine-readable result

Exit code 0 when PPTX export works. A missing browser only disables image QA,
so it is reported as a warning. Nothing is installed or changed.
"""

from __future__ import annotations

import argparse
import json
import os
import platform
import shutil
import sys
import tempfile
import time
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

sys.path.insert(0, str(Path(__file__).resolve().parent))

import cdp  # noqa: E402
import export_pptx  # noqa: E402

TEST_MANIFEST = """\
version: v2
title: open-kimi-ppt doctor
size: [960, 540]
pages:
  - pages/1.page
  - pages/2.page
"""
TEST_PAGE = """\
pageType: final
background: {{type: solid, color: "#FFFFFF"}}
elements:
  - elementId: title
    elementType: text
    bounds: [60, 60, 840, 140]
    content: {{fontSize: 36, fontFamily: MiSans, bold: true, color: "#111827", text: "环境检查 {index}<br/>Doctor page {index}"}}
"""


class Report:
    def __init__(self) -> None:
        self.items: List[Dict[str, Any]] = []

    def add(self, name: str, status: str, detail: str) -> None:
        self.items.append({"check": name, "status": status, "detail": detail})

    def status(self, name: str) -> Optional[str]:
        return next((item["status"] for item in self.items if item["check"] == name), None)


def timed(function: Callable[[], Any]) -> Any:
    started = time.monotonic()
    result = function()
    return result, time.monotonic() - started


def check_python(report: Report) -> None:
    version = platform.python_version()
    status = "ok" if sys.version_info >= (3, 8) else "fail"
    report.add("Python", status, f"{version} ({sys.executable})")


def check_yaml(report: Report) -> None:
    module = export_pptx.yaml
    origin = Path(getattr(module, "__file__", "") or "")
    vendored = "_vendor" in origin.parts
    source = "bundled copy" if vendored else "installed"
    report.add("PyYAML", "ok", f"{getattr(module, '__version__', '?')} ({source})")


def check_files(report: Report) -> None:
    problems = []
    try:
        editor = export_pptx.resolve_editor_root()
    except export_pptx.ExportError as exc:
        problems.append(str(exc))
        editor = None
    try:
        export_pptx.resolve_local_wasm()
    except export_pptx.ExportError as exc:
        problems.append(str(exc))
    if not export_pptx.LOCAL_EXPORT_MJS.is_file():
        problems.append(f"missing {export_pptx.LOCAL_EXPORT_MJS}")
    if problems:
        report.add("Skill files", "fail", "; ".join(problems))
    else:
        report.add("Skill files", "ok", f"editor, WASM and exporter present ({editor})")


def check_node(report: Report) -> None:
    try:
        _command, _env, label = export_pptx.node_runtime()
    except export_pptx.ExportError as exc:
        report.add("Node runtime", "fail", str(exc))
        return
    report.add("Node runtime", "ok", label)


def check_browser(report: Report) -> None:
    candidates = cdp.browser_candidates()
    if not candidates:
        report.add(
            "Browser",
            "warn",
            "no Chrome or Edge found: image QA is unavailable (install Google Chrome or "
            "Microsoft Edge, or set OKP_BROWSER)",
        )
        return
    try:
        browser, seconds = timed(cdp.launch_browser)
    except (cdp.CDPError, OSError) as exc:
        report.add("Browser", "warn", f"image QA is unavailable: {exc}")
        return
    try:
        version = ""
        if browser.cdp is not None:
            version = browser.cdp.call("Browser.getVersion").get("product", "")
    finally:
        browser.close()
    report.add("Browser", "ok", f"{version or 'started'} in {seconds:.1f}s ({browser.executable})")


def check_fonts(report: Report) -> None:
    if sys.platform != "win32":
        report.add("Fonts", "ok", "system fonts are mapped by PowerPoint/WPS on this platform")
        return
    fonts = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
    if (fonts / "msyh.ttc").exists() or (fonts / "msyh.ttf").exists():
        report.add("Fonts", "ok", "Microsoft YaHei installed")
    else:
        report.add("Fonts", "warn", "Microsoft YaHei not found: Chinese text will use another font")


def write_test_deck(root: Path) -> Path:
    (root / "pages").mkdir(parents=True)
    manifest = root / "doctor.pptd"
    manifest.write_text(TEST_MANIFEST, encoding="utf-8")
    for index in (1, 2):
        (root / "pages" / f"{index}.page").write_text(
            TEST_PAGE.format(index=index), encoding="utf-8"
        )
    return manifest


def smoke_pptx(report: Report, manifest: Path) -> None:
    output = manifest.with_suffix(".pptx")
    try:
        summary, seconds = timed(
            lambda: export_pptx.export_pptx(manifest, output, "fade", False, force=True)
        )
    except (export_pptx.ExportError, cdp.CDPError, OSError) as exc:
        report.add("PPTX export", "fail", str(exc))
        return
    ok = summary.get("slides") == 2 and summary.get("lineBreaks") == 2
    detail = (
        f"{summary.get('slides')} slides in {seconds:.1f}s via {summary.get('exporter')}; "
        f"fonts replaced: {summary.get('fontsReplaced') or {}}; line breaks: {summary.get('lineBreaks', 0)}"
    )
    report.add("PPTX export", "ok" if ok else "fail", detail)


def smoke_images(report: Report, manifest: Path) -> None:
    if report.status("Browser") != "ok":
        report.add("Image QA", "warn", "skipped: no usable browser")
        return
    import export_images

    try:
        summary, seconds = timed(
            lambda: export_images.export_images(manifest, manifest.parent / ".qa-images", force=True)
        )
    except (export_pptx.ExportError, cdp.CDPError, OSError) as exc:
        report.add("Image QA", "warn", f"image QA failed: {exc}")
        return
    ok = summary.get("pages") == 2 and Path(summary["overview"]).is_file()
    report.add(
        "Image QA",
        "ok" if ok else "warn",
        f"{summary.get('pages')} page images + overview in {seconds:.1f}s",
    )


def main(argv: Optional[List[str]] = None) -> int:
    export_pptx.make_output_robust()
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--quick", action="store_true", help="skip the test export")
    parser.add_argument("--json", action="store_true", help="print JSON instead of text")
    args = parser.parse_args(argv)

    report = Report()
    check_python(report)
    check_yaml(report)
    check_files(report)
    check_node(report)
    check_browser(report)
    check_fonts(report)
    if not args.quick:
        workdir = Path(tempfile.mkdtemp(prefix="okp-doctor-"))
        try:
            manifest = write_test_deck(workdir)
            smoke_pptx(report, manifest)
            smoke_images(report, manifest)
        finally:
            shutil.rmtree(workdir, ignore_errors=True)

    pptx_ok = report.status("Node runtime") == "ok" and report.status("Skill files") == "ok"
    if not args.quick:
        pptx_ok = report.status("PPTX export") == "ok"
    qa_ok = report.status("Image QA" if not args.quick else "Browser") == "ok"
    if pptx_ok and qa_ok:
        verdict = "ready: PPTX export and image QA both work"
    elif pptx_ok:
        verdict = "PPTX export works; image QA is unavailable (use the structural review instead)"
    else:
        verdict = "PPTX export does not work on this computer; see the failed checks above"

    if args.json:
        print(json.dumps({"ready": pptx_ok, "imageQA": qa_ok, "verdict": verdict, "checks": report.items},
                         ensure_ascii=False, indent=2))
    else:
        labels = {"ok": "[OK]  ", "warn": "[WARN]", "fail": "[FAIL]"}
        for item in report.items:
            print(f"{labels[item['status']]} {item['check']}: {item['detail']}")
        print(f"Result: {verdict}")
    return 0 if pptx_ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
