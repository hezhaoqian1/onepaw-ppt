#!/usr/bin/env python3
"""Export a PPTD project to PPTX.

Default path (preferred): local patched official WASM writer
  scripts/local-export/export-pptd.mjs --no-sign
  -> offline, no cookie, no signature API, no browser UI. Needs Node.js 18+;
  when Node is not installed, the Electron runtime of a desktop agent such as
  OnePaw is used instead (ELECTRON_RUN_AS_NODE).

Fallback / --browser path: the local neo-ppt editor driven in a headless
Chrome/Edge through the DevTools protocol (scripts/cdp.py, stdlib only).

Every export is post-processed for compatibility with PowerPoint and WPS:
fade transitions, real line breaks, and editor-only web fonts replaced by
look-alike system fonts (see FONT_FALLBACKS; disable with --keep-fonts).

Image QA (`export_images.py`) uses the same local editor host.
"""

from __future__ import annotations

import argparse
import base64
import html
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import uuid
import zipfile
import xml.etree.ElementTree as ET
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple
from urllib.parse import urlparse
from xml.sax.saxutils import escape as xml_escape

# Embedded Pythons configured by a ._pth file (e.g. the one bundled with desktop
# agents such as OnePaw) do not put the script's directory on sys.path.
sys.path.insert(0, str(Path(__file__).resolve().parent))

from cdp import Browser, CDPError, Page, launch_browser  # noqa: E402

SKILL_DIR = Path(__file__).resolve().parent.parent
IMAGE_MIME = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".svg": "image/svg+xml",
}
MAX_IMAGE_BYTES = 20 * 1024 * 1024
MAX_EMBEDDED_MEDIA_BYTES = 200 * 1024 * 1024
PPTX_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
)
FADE_TRANSITION_XML = (
    '<p:transition spd="fast" advClick="1"><p:fade/></p:transition>'
)
MIN_NODE_MAJOR = 18
NODE_INSTALL_HINT = "Install Node.js 18+ from https://nodejs.org, then retry."
EDITOR_MISSING_HINT = (
    "local neo-ppt editor not found. Re-run "
    "`npx open-kimi-ppt-skill install` (copies editor into the skill) "
    "or set OPEN_KIMI_PPT_EDITOR to the editor directory."
)


class ExportError(RuntimeError):
    pass


class QuietHandler(SimpleHTTPRequestHandler):
    def log_message(self, _format: str, *_args: Any) -> None:
        return


def log(message: str) -> None:
    print(f"[onepaw-ppt] {message}", file=sys.stderr, flush=True)


def _registry_path_entries() -> List[str]:
    """Machine + user PATH as configured in the Windows registry."""
    try:
        import winreg
    except ImportError:
        return []
    entries: List[str] = []
    for root, key in (
        (winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment"),
        (winreg.HKEY_CURRENT_USER, "Environment"),
    ):
        try:
            with winreg.OpenKey(root, key) as handle:
                value, _kind = winreg.QueryValueEx(handle, "Path")
        except OSError:
            continue
        entries.extend(os.path.expandvars(value).split(os.pathsep))
    return entries


def augment_path() -> None:
    """Let node be found when the host trims PATH.

    Desktop agents may run skill commands with a reduced PATH: OnePaw's shell,
    for example, has its bundled Python but no Node even when Node is installed.
    Known locations are appended (never prepended) so an explicit PATH still
    wins; child processes inherit the result.
    """
    entries = [entry for entry in os.environ.get("PATH", "").split(os.pathsep) if entry]
    seen = {os.path.normcase(os.path.normpath(entry)) for entry in entries}
    candidates: List[str] = []
    if sys.platform == "win32":
        candidates.extend(_registry_path_entries())
        for variable, child in (
            ("ProgramFiles", "nodejs"),
            ("ProgramFiles(x86)", "nodejs"),
            ("APPDATA", "npm"),
            ("NVM_SYMLINK", ""),
        ):
            base = os.environ.get(variable)
            if base:
                candidates.append(os.path.join(base, child) if child else base)
    else:
        candidates.extend(
            ["/opt/homebrew/bin", "/usr/local/bin", str(Path.home() / ".volta" / "bin")]
        )
    for candidate in candidates:
        candidate = candidate.strip().strip('"')
        if not candidate:
            continue
        key = os.path.normcase(os.path.normpath(candidate))
        if key in seen or not os.path.isdir(candidate):
            continue
        seen.add(key)
        entries.append(candidate)
    os.environ["PATH"] = os.pathsep.join(entries)


augment_path()


def run_command(
    command: Sequence[str],
    *,
    cwd: Optional[Path] = None,
    env: Optional[Dict[str, str]] = None,
    timeout: int = 90,
) -> subprocess.CompletedProcess[str]:
    """Capture merged stdout/stderr via a temp file.

    On Windows, a detached grandchild process can inherit a PIPE handle and
    prevent EOF, deadlocking ``subprocess.run(stdout=PIPE)``. Decoding with the
    system locale (GBK on zh-CN Windows) can also raise UnicodeDecodeError.
    Writing to a UTF-8 file avoids both failures.
    """
    handle, sink_path = tempfile.mkstemp(prefix="onepaw-ppt-", suffix=".log")
    os.close(handle)
    sink = Path(sink_path)
    output = ""
    try:
        with sink.open("w", encoding="utf-8", errors="replace") as out:
            returncode = subprocess.call(
                list(command),
                cwd=str(cwd) if cwd is not None else None,
                env=env,
                stdout=out,
                stderr=subprocess.STDOUT,
                timeout=timeout,
            )
        output = sink.read_text(encoding="utf-8", errors="replace")
    except subprocess.TimeoutExpired as exc:
        try:
            output = sink.read_text(encoding="utf-8", errors="replace")
        except OSError:
            output = ""
        raise subprocess.TimeoutExpired(
            cmd=list(command),
            timeout=timeout,
            output=output,
        ) from exc
    finally:
        try:
            sink.unlink(missing_ok=True)
        except OSError:
            # WinError 32: daemon may still hold the log file handle.
            pass
    return subprocess.CompletedProcess(list(command), returncode, output, None)


def temporary_directory(prefix: str) -> Any:
    # ignore_cleanup_errors avoids masking the real export error when a Windows
    # browser daemon still holds files under the temp tree (Python 3.10+).
    try:
        return tempfile.TemporaryDirectory(prefix=prefix, ignore_cleanup_errors=True)
    except TypeError:
        return tempfile.TemporaryDirectory(prefix=prefix)


def load_yaml_module() -> Any:
    """PyYAML, from site-packages or from the copy vendored with this skill."""
    try:
        import yaml as module
    except ImportError:
        vendor = str(Path(__file__).resolve().parent / "_vendor")
        if vendor not in sys.path:
            sys.path.append(vendor)
        import yaml as module
    return module


yaml = load_yaml_module()


def parse_node_version(output: str) -> Tuple[int, int, int]:
    match = re.search(r"v?(\d+)\.(\d+)\.(\d+)\b", output)
    if not match:
        raise ExportError(f"could not parse Node.js version from: {output.strip()}")
    return tuple(int(part) for part in match.groups())


def _electron_in(directory: Path) -> Optional[Path]:
    """The Electron app executable in ``directory``, if it is an Electron app."""
    if sys.platform == "darwin":
        if directory.suffix == ".app" and (directory / "Contents" / "Frameworks").is_dir():
            binaries = sorted((directory / "Contents" / "MacOS").glob("*"))
            return next((path for path in binaries if path.is_file()), None)
        return None
    markers = ("LICENSE.electron.txt", "resources/app.asar")
    if not any((directory / marker).exists() for marker in markers):
        return None
    if sys.platform == "win32":
        executables = [
            path
            for path in directory.glob("*.exe")
            if not path.name.lower().startswith(("uninstall", "elevate", "squirrel"))
        ]
    else:
        executables = [
            path
            for path in directory.iterdir()
            if path.is_file() and os.access(path, os.X_OK) and not path.suffix
            and path.name not in ("chrome-sandbox", "chrome_crashpad_handler")
        ]
    # The app binary is by far the largest file next to the Electron markers.
    return max(executables, key=lambda path: path.stat().st_size, default=None)


def find_electron() -> Optional[Path]:
    """An Electron runtime that can stand in for Node via ELECTRON_RUN_AS_NODE.

    Desktop agents (OnePaw and other Electron apps) bundle the Python that runs
    this script inside their install folder, e.g.
    ``<app>/resources/portable_py311/python.exe``. Their Electron binary embeds
    Node, so Node.js does not have to be installed on the machine.
    """
    explicit = os.environ.get("OKP_ELECTRON")
    if explicit and Path(explicit).is_file():
        return Path(explicit)
    for directory in list(Path(sys.executable).resolve().parents)[:5]:
        try:
            found = _electron_in(directory)
        except OSError:
            continue
        if found is not None:
            return found
    return None


_NODE: Optional[Tuple[List[str], Dict[str, str], str]] = None


def node_runtime() -> Tuple[List[str], Dict[str, str], str]:
    """Return (command prefix, extra env, description) for running a .mjs file."""
    global _NODE
    if _NODE is not None:
        return _NODE
    problems = []
    node = shutil.which("node")
    if node:
        try:
            process = run_command([node, "--version"], timeout=30)
            version = parse_node_version(process.stdout)
            if version[0] >= MIN_NODE_MAJOR:
                _NODE = ([node], {}, f"Node.js {'.'.join(map(str, version))}")
                return _NODE
            problems.append(
                f"Node.js {'.'.join(map(str, version))} at {node} is older than {MIN_NODE_MAJOR}"
            )
        except (ExportError, OSError, subprocess.SubprocessError) as exc:
            problems.append(f"{node} --version failed: {exc}")
    electron = find_electron()
    if electron is not None:
        env = {"ELECTRON_RUN_AS_NODE": "1"}
        try:
            process = run_command(
                [str(electron), "-e", "process.stdout.write(process.versions.node)"],
                env={**os.environ, **env},
                timeout=30,
            )
            version = parse_node_version(process.stdout)
            if version[0] >= MIN_NODE_MAJOR:
                label = f"Node {'.'.join(map(str, version))} via {electron.name}"
                _NODE = ([str(electron)], env, label)
                return _NODE
            problems.append(f"{electron} embeds Node {'.'.join(map(str, version))}")
        except (ExportError, OSError, subprocess.SubprocessError) as exc:
            problems.append(f"{electron} cannot run as Node: {exc}")
    detail = ("; ".join(problems) + ". ") if problems else ""
    raise ExportError(f"Node.js {MIN_NODE_MAJOR}+ was not found. {detail}{NODE_INSTALL_HINT}")


def ensure_nodejs() -> str:
    """Description of the Node runtime that will be used (raises if none)."""
    return node_runtime()[2]


def find_manifest(source: Path) -> Path:
    source = source.expanduser().resolve()
    if source.is_file():
        if source.suffix.lower() != ".pptd":
            raise ExportError(f"input must be a .pptd file or project directory: {source}")
        return source
    if not source.is_dir():
        raise ExportError(f"input does not exist: {source}")
    manifests = sorted(source.rglob("*.pptd"))
    if not manifests:
        raise ExportError(f"no .pptd manifest found under: {source}")
    if len(manifests) > 1:
        choices = "\n  ".join(str(path) for path in manifests[:20])
        raise ExportError(
            "multiple .pptd manifests found; pass one manifest explicitly:\n  " + choices
        )
    return manifests[0]


BR_TAG = re.compile(r"<br\s*/?>", re.IGNORECASE)


def find_dropped_line_breaks(manifest: Path) -> List[str]:
    """Pages using <br/>, which the editor's own PPTX writer drops."""
    project = manifest.parent
    hits = []
    for page in sorted(project.rglob("*.page")):
        try:
            count = len(BR_TAG.findall(page.read_text(encoding="utf-8")))
        except OSError:
            continue
        if count:
            hits.append(f"{page.relative_to(project).as_posix()} ({count})")
    return hits


def warn_dropped_line_breaks(manifest: Path, summary: Dict[str, Any]) -> None:
    """Only the browser exporter drops <br/>; the local exporter converts it."""
    if summary.get("exporter") != "browser-local-editor":
        return
    hits = find_dropped_line_breaks(manifest)
    if not hits:
        return
    message = (
        "the browser exporter drops <br/> line breaks, so those lines run together; "
        "split them into separate <p> paragraphs and re-export: " + ", ".join(hits)
    )
    log(f"warning: {message}")
    summary.setdefault("warnings", []).append(message)


def read_yaml_mapping(path: Path) -> Tuple[str, Dict[str, Any]]:
    text = path.read_text(encoding="utf-8")
    try:
        value = yaml.safe_load(text)
    except yaml.YAMLError as exc:
        raise ExportError(f"invalid YAML in {path}: {exc}") from exc
    if not isinstance(value, dict):
        raise ExportError(f"expected a YAML mapping in {path}")
    return text, value


def safe_project_path(root: Path, relative: str) -> Path:
    if not isinstance(relative, str) or not relative.strip():
        raise ExportError("page path must be a non-empty string")
    candidate = (root / relative).resolve()
    try:
        candidate.relative_to(root)
    except ValueError as exc:
        raise ExportError(f"project path escapes the PPTD directory: {relative}") from exc
    return candidate


def build_image_map(root: Path) -> Dict[str, str]:
    image_map: Dict[str, str] = {}
    total = 0
    for path in sorted(root.rglob("*")):
        if not path.is_file() or path.suffix.lower() not in IMAGE_MIME:
            continue
        size = path.stat().st_size
        if size > MAX_IMAGE_BYTES:
            log(f"skip local image over 20 MiB: {path.relative_to(root)}")
            continue
        if total + size > MAX_EMBEDDED_MEDIA_BYTES:
            raise ExportError(
                "local image payload exceeds 200 MiB; reduce media size or use remote URLs"
            )
        data = base64.b64encode(path.read_bytes()).decode("ascii")
        rel = path.relative_to(root).as_posix()
        image_map[rel] = f"data:{IMAGE_MIME[path.suffix.lower()]};base64,{data}"
        total += size
    if image_map:
        log(f"prepared {len(image_map)} local image resource(s), {total} bytes")
    return image_map


def build_payload(manifest: Path) -> Dict[str, Any]:
    manifest_text, manifest_data = read_yaml_mapping(manifest)
    if manifest_data.get("version") != "v2":
        raise ExportError("local PPTX export currently requires PPTD version: v2")
    page_paths = manifest_data.get("pages")
    if not isinstance(page_paths, list) or not page_paths:
        raise ExportError("PPTD manifest must contain a non-empty pages list")

    root = manifest.parent.resolve()
    pages: List[Dict[str, str]] = []
    for entry in page_paths:
        page_path = safe_project_path(root, entry)
        if not page_path.is_file():
            raise ExportError(f"missing page file: {entry}")
        page_text, page_data = read_yaml_mapping(page_path)
        if not isinstance(page_data.get("elements"), list):
            raise ExportError(f"page elements must be an array: {entry}")
        pages.append({"path": str(entry), "content": page_text})

    title = str(manifest_data.get("title") or manifest.stem)
    return {
        "id": f"local-export-{uuid.uuid4().hex}",
        "title": title,
        "manifestPath": manifest.name,
        "manifestContent": manifest_text,
        "pages": pages,
        "imageMap": build_image_map(root),
    }


EXPORT_DIALOG = ".export-dialog.visible"
DECK_READY_JS = """
(() => {
  const status = document.documentElement.dataset.deckStatus;
  if (status === 'error') throw new Error(window.exportHostError || 'the editor could not load the deck');
  return status === 'ready';
})()
""".strip()
ACTIVE_FORMAT_JS = """
(() => {
  const dialog = [...document.querySelectorAll('.export-dialog.visible')].pop();
  const active = dialog && dialog.querySelector('.radio-group-item.active');
  return active ? active.textContent.trim() : null;
})()
""".strip()
FONT_SWITCH_JS = """
(() => {
  const dialog = [...document.querySelectorAll('.export-dialog.visible')].pop();
  const sw = dialog && dialog.querySelector('[role=switch]');
  if (!sw) return null;
  const r = sw.getBoundingClientRect();
  return {
    checked: sw.getAttribute('aria-checked') === 'true',
    disabled: !!sw.disabled || sw.classList.contains('disabled'),
    x: r.left + r.width / 2,
    y: r.top + r.height / 2,
  };
})()
""".strip()


class EditorHost:
    """The local editor served for one payload plus a headless browser.

    Use as a context manager; ``export()`` drives the editor's export dialog and
    returns the downloaded file. The browser stays available as ``browser`` for
    follow-up work such as stitching the QA overview.
    """

    def __init__(self, payload: Dict[str, Any], *, font_fallbacks: bool = True) -> None:
        self.payload = payload
        self.font_fallbacks = font_fallbacks
        self.browser: Optional[Browser] = None
        self.server: Optional[ThreadingHTTPServer] = None
        self.thread: Optional[threading.Thread] = None
        self.url = ""

    def __enter__(self) -> "EditorHost":
        self.server, self.thread, self.url = serve_local_editor(self.payload)
        try:
            self.browser = launch_browser()
        except Exception:
            self._stop_server()
            raise
        log(f"headless browser: {self.browser.executable}")
        return self

    def __exit__(self, *_exc: Any) -> None:
        if self.browser is not None:
            self.browser.close()
            self.browser = None
        self._stop_server()

    def _stop_server(self) -> None:
        if self.server is not None:
            self.server.shutdown()
            self.server.server_close()
            self.server = None
        if self.thread is not None:
            self.thread.join(timeout=2)
            self.thread = None

    def open_editor(self, download_dir: Path) -> Page:
        if self.browser is None:
            raise ExportError("EditorHost must be used as a context manager")
        page = Page(self.browser)
        page.allow_downloads(download_dir)
        if self.font_fallbacks:
            page.call("Page.addScriptToEvaluateOnNewDocument", {"source": font_alias_script()})
        page.navigate(self.url)
        page.wait_for(DECK_READY_JS, 120, "the editor to load the deck")
        return page

    def export(
        self,
        download_dir: Path,
        *,
        images: bool,
        embed_fonts: Optional[bool] = None,
        timeout: float = 240.0,
    ) -> Path:
        page = self.open_editor(download_dir)
        page.click_text("导出")
        page.wait_for(
            f"!!document.querySelector({json.dumps(EXPORT_DIALOG)})", 20, "the export dialog"
        )
        if images:
            page.click_text("图片", scope=EXPORT_DIALOG)
            page.wait_for(f"{ACTIVE_FORMAT_JS} === '图片'", 10, "the image format option")
        if embed_fonts is not None:
            state = page.evaluate(FONT_SWITCH_JS)
            if state is None:
                if embed_fonts:
                    log("warning: the export dialog exposed no font switch")
            elif state["checked"] != embed_fonts:
                if state["disabled"]:
                    log("warning: the font switch is disabled for this deck")
                else:
                    page.click_at(state["x"], state["y"])
        started_at = time.time() - 1.0
        page.click_text("下载", scope=EXPORT_DIALOG)
        log("rendering page images in the local editor" if images else "generating PPTX in the local editor")
        accept = is_image_zip if images else is_pptx
        return find_download((download_dir,), timeout=timeout, accept=accept, since=started_at)


def is_image_zip(path: Path) -> bool:
    if not path.is_file() or path.name.endswith(".crdownload"):
        return False
    try:
        with zipfile.ZipFile(path) as archive:
            return any(
                Path(name).suffix.lower() in (".png", ".jpg", ".jpeg", ".webp")
                for name in archive.namelist()
            )
    except (OSError, zipfile.BadZipFile):
        return False


def is_pptx(path: Path) -> bool:
    if not path.is_file() or path.name.endswith(".crdownload"):
        return False
    try:
        with zipfile.ZipFile(path) as archive:
            if "ppt/presentation.xml" not in archive.namelist():
                return False
            content_types = archive.read("[Content_Types].xml")
            return PPTX_CONTENT_TYPE.encode("utf-8") in content_types
    except (OSError, KeyError, zipfile.BadZipFile):
        return False


def find_download(
    search_roots: Iterable[Path],
    timeout: float = 150.0,
    accept: Callable[[Path], bool] = is_pptx,
    *,
    since: Optional[float] = None,
) -> Path:
    deadline = time.monotonic() + timeout
    last_sizes: Dict[Path, int] = {}
    stable: Dict[Path, int] = {}
    while time.monotonic() < deadline:
        # Snapshot stats while collecting and tolerate races everywhere: the
        # search roots include the live Downloads folder, where Chrome renames
        # .crdownload files away between directory listing and stat().
        entries: List[Tuple[Path, float, int]] = []
        for root in search_roots:
            if not root.exists():
                continue
            for path in root.rglob("*"):
                if not path.is_file():
                    continue
                try:
                    info = path.stat()
                except OSError:
                    continue
                entries.append((path, info.st_mtime, info.st_size))
        for path, mtime, size in sorted(entries, key=lambda entry: entry[1], reverse=True):
            if since is not None and mtime < since:
                continue
            if size == last_sizes.get(path) and size > 0:
                stable[path] = stable.get(path, 0) + 1
            else:
                stable[path] = 0
            last_sizes[path] = size
            if stable[path] >= 1 and accept(path):
                return path
        time.sleep(0.5)
    visible = "\n  ".join(str(path) for path in last_sizes) or "(none)"
    raise ExportError(f"timed out waiting for download; observed files:\n  {visible}")


def replace_transition(slide_xml: bytes, transition: str) -> bytes:
    text = slide_xml.decode("utf-8")
    pattern = re.compile(
        r"<p:transition\b[^>]*(?:/>|>.*?</p:transition>)", re.DOTALL
    )
    text = pattern.sub("", text)
    if transition == "none":
        return text.encode("utf-8")

    # CT_Slide requires transition as a direct child after cSld/clrMapOvr and
    # before timing/extLst. Searching for the first p:extLst is incorrect:
    # shapes may contain their own nested extLst inside cSld, causing Office to
    # ignore a transition inserted there.
    color_map = re.search(
        r"<p:clrMapOvr\b[^>]*(?:/>|>.*?</p:clrMapOvr>)", text, re.DOTALL
    )
    common_slide = re.search(
        r"<p:cSld\b[^>]*(?:/>|>.*?</p:cSld>)", text, re.DOTALL
    )
    anchor = color_map or common_slide
    if anchor is None:
        raise ExportError("slide XML has no cSld/clrMapOvr insertion anchor")
    position = anchor.end()
    return (text[:position] + FADE_TRANSITION_XML + text[position:]).encode("utf-8")


def root_child_names(slide_xml: bytes) -> List[str]:
    try:
        root = ET.fromstring(slide_xml)
    except ET.ParseError as exc:
        raise ExportError(f"invalid slide XML: {exc}") from exc
    return [child.tag.rsplit("}", 1)[-1] for child in root]


def has_direct_fade_transition(slide_xml: bytes) -> bool:
    try:
        root = ET.fromstring(slide_xml)
    except ET.ParseError as exc:
        raise ExportError(f"invalid slide XML: {exc}") from exc
    transition = next(
        (child for child in root if child.tag.rsplit("}", 1)[-1] == "transition"),
        None,
    )
    if transition is None:
        return False
    return any(child.tag.rsplit("}", 1)[-1] == "fade" for child in transition)


def validate_transition_order(slide_xml: bytes, transition: str) -> None:
    names = root_child_names(slide_xml)
    transition_indexes = [index for index, name in enumerate(names) if name == "transition"]
    if transition == "none":
        if transition_indexes:
            raise ExportError("transition=none left a root-level transition")
        return
    if len(transition_indexes) != 1 or not has_direct_fade_transition(slide_xml):
        raise ExportError("slide does not contain exactly one root-level fade transition")
    transition_index = transition_indexes[0]
    for required_before in ("cSld", "clrMapOvr"):
        if required_before in names and names.index(required_before) > transition_index:
            raise ExportError(f"{required_before} appears after transition")
    for required_after in ("timing", "extLst"):
        if required_after in names and names.index(required_after) < transition_index:
            raise ExportError(f"{required_after} appears before transition")


# Fonts the editor offers that are not installed on ordinary Windows / Office /
# WPS machines. The local writer cannot embed them and PowerPoint ignores the
# copies the browser path embeds, so on the audience's computer they fall back
# to SimSun/Calibri and the layout breaks. Every export therefore swaps them for
# look-alike fonts that ship with Windows and Office (and that WPS and PowerPoint
# for Mac map). Image QA renders the same substitutes, so the preview matches
# the PPTX. Keys are the names the editor writes, plus the aliases it accepts.
FONT_FALLBACKS: Dict[str, str] = {
    # Chinese sans-serif
    "MiSans": "Microsoft YaHei",
    "Mi Sans": "Microsoft YaHei",
    "MiSans LC": "Microsoft YaHei",
    "Noto Sans SC": "Microsoft YaHei",
    "Noto Sans": "Microsoft YaHei",
    "思源黑体": "Microsoft YaHei",
    "Source Han Sans": "Microsoft YaHei",
    "Source Han Sans SC": "Microsoft YaHei",
    "SourceHanSans": "Microsoft YaHei",
    "阿里妈妈数黑体": "Microsoft YaHei",
    "得意黑": "Microsoft YaHei",
    "Smiley Sans": "Microsoft YaHei",
    "ZCOOL KuaiLe": "Microsoft YaHei",
    "站酷快乐体": "Microsoft YaHei",
    # Chinese serif
    "思源宋体": "SimSun",
    "Source Han Serif": "SimSun",
    "Source Han Serif SC": "SimSun",
    "SourceHanSerif": "SimSun",
    "Noto Serif SC": "SimSun",
    "霞鹜新致宋": "SimSun",
    "精品点阵体": "SimSun",
    # Chinese calligraphy / handwriting
    "阿里妈妈东方大楷": "KaiTi",
    "阿里妈妈刀隶体": "KaiTi",
    "飞波正点体": "KaiTi",
    "程荣光刻楷": "KaiTi",
    "站酷文艺体": "KaiTi",
    "LXGW Bright": "KaiTi",
    "LXGW WenKai": "KaiTi",
    "霞鹜文楷": "KaiTi",
    "Long Cang": "KaiTi",
    # Latin
    "Liter": "Arial",
    "HedvigLettersSans": "Arial",
    "Hedvig Letters Sans": "Arial",
    "QuattrocentoSans": "Arial",
    "Quattrocento Sans": "Arial",
    "Coda": "Arial",
    "Oranienbaum": "Georgia",
    "SortsMillGoudy": "Georgia",
    "Sorts Mill Goudy": "Georgia",
    "Unna": "Georgia",
    "Jersey15": "Impact",
    "Jersey 15": "Impact",
    "Jersey20Charted": "Impact",
    "Jersey 20 Charted": "Impact",
    "LuckiestGuy": "Impact",
    "Luckiest Guy": "Impact",
    "PressStart2P": "Consolas",
    "Press Start 2P": "Consolas",
}

# local() names of the substitutes for the browser preview: (regular, bold).
FONT_FACES: Dict[str, Tuple[Tuple[str, ...], Tuple[str, ...]]] = {
    "Microsoft YaHei": (
        ("Microsoft YaHei", "微软雅黑", "PingFang SC Regular"),
        ("Microsoft YaHei Bold", "PingFang SC Semibold"),
    ),
    "SimSun": (("SimSun", "宋体", "Songti SC Regular"), ()),
    "KaiTi": (("KaiTi", "楷体", "STKaiti", "Kaiti SC Regular"), ()),
    "Arial": (("Arial", "ArialMT"), ("Arial Bold", "Arial-BoldMT")),
    "Georgia": (("Georgia",), ("Georgia Bold", "Georgia-Bold")),
    "Impact": (("Impact",), ()),
    "Consolas": (("Consolas", "Menlo Regular"), ("Consolas Bold", "Menlo Bold")),
}

# Fonts that ship with Windows + Office (and WPS); anything else gets a warning
# because another computer may not have it.
STANDARD_FONTS = {
    name.lower()
    for name in (
        "Microsoft YaHei", "微软雅黑", "Microsoft YaHei UI", "Microsoft YaHei Light",
        "SimSun", "宋体", "NSimSun", "新宋体", "SimHei", "黑体", "KaiTi", "楷体",
        "FangSong", "仿宋", "DengXian", "等线", "DengXian Light", "等线 Light",
        "Microsoft JhengHei", "PingFang SC", "Arial", "Arial Black", "Calibri",
        "Calibri Light", "Cambria", "Candara", "Consolas", "Constantia", "Corbel",
        "Courier New", "Georgia", "Impact", "Segoe UI", "Segoe UI Light",
        "Segoe UI Semibold", "Segoe UI Black", "Tahoma", "Times New Roman",
        "Trebuchet MS", "Verdana", "Bahnschrift", "Symbol", "Wingdings",
        "Helvetica", "Helvetica Neue",
    )
}

TYPEFACE_ATTR = re.compile(
    r'(<a:(?:latin|ea|cs|sym|font|buFont)\b[^>]*?\btypeface=")([^"]*)(")'
)
TEXT_RUN = re.compile(
    r"<a:r>(<a:rPr\b[^>]*/>|<a:rPr\b(?:[^>]*[^/>])?>(?:(?!</a:r>).)*?</a:rPr>)?"
    r"<a:t(\s[^>]*)?>([^<]*)</a:t></a:r>",
    re.DOTALL,
)
NEWLINE = re.compile(r"\r\n|&#13;&#10;|&#x0*[dD];&#x0*[aA];|\r|\n|&#1[03];|&#x0*[aAdD];")
SLIDE_PART = re.compile(r"ppt/slides/slide\d+\.xml")


def _css_string(name: str) -> str:
    return '"' + name.replace("\\", "\\\\").replace('"', '\\"') + '"'


def font_alias_css() -> str:
    """@font-face rules that render each FONT_FALLBACKS font with its substitute."""
    rules = []
    for family, target in FONT_FALLBACKS.items():
        regular, bold = FONT_FACES[target]
        sources = ",".join(f"local({_css_string(name)})" for name in regular)
        rules.append(f"@font-face{{font-family:{_css_string(family)};src:{sources}}}")
        if bold:
            sources = ",".join(f"local({_css_string(name)})" for name in bold)
            rules.append(
                f"@font-face{{font-family:{_css_string(family)};font-weight:700;src:{sources}}}"
            )
    return "\n".join(rules)


FONT_ALIAS_SCRIPT = """
(() => {
  const css = __CSS__;
  const blocked = new Set(__FAMILIES__);
  const strip = (s) => String(s || '').trim().replace(/^["']|["']$/g, '').toLowerCase();
  // Keep the editor from registering its web fonts for these families at runtime.
  const add = FontFaceSet.prototype.add;
  FontFaceSet.prototype.add = function (face) {
    if (face && blocked.has(strip(face.family))) return this;
    return add.call(this, face);
  };
  const install = () => {
    const style = document.createElement('style');
    style.dataset.okp = 'font-fallbacks';
    style.textContent = css;
    (document.head || document.documentElement).appendChild(style);
  };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', install);
  else install();
})();
""".strip()


def font_alias_script() -> str:
    """Page script that renders editor web fonts with the PPTX substitutes."""
    families = sorted({name.lower() for name in FONT_FALLBACKS})
    return FONT_ALIAS_SCRIPT.replace(
        "__CSS__", json.dumps(font_alias_css(), ensure_ascii=False)
    ).replace("__FAMILIES__", json.dumps(families, ensure_ascii=False))


def map_fonts_in_xml(text: str, replaced: Dict[str, str]) -> str:
    def swap(match: "re.Match[str]") -> str:
        name = html.unescape(match.group(2))
        target = FONT_FALLBACKS.get(name)
        if target is None:
            return match.group(0)
        replaced[name] = target
        return f"{match.group(1)}{xml_escape(target)}{match.group(3)}"

    return TYPEFACE_ATTR.sub(swap, text)


def split_newline_runs(text: str) -> Tuple[str, int]:
    """Turn newlines inside <a:t> into real <a:br/> line breaks.

    PowerPoint happens to render a raw newline in <a:t>, but the schema wants
    <a:br>, and other viewers (WPS, Keynote, LibreOffice) may not honour it.
    """
    count = 0

    def split(match: "re.Match[str]") -> str:
        nonlocal count
        props, attrs, body = match.group(1) or "", match.group(2) or "", match.group(3)
        parts = NEWLINE.split(body)
        if len(parts) == 1:
            return match.group(0)
        count += len(parts) - 1
        line_break = f"<a:br>{props}</a:br>" if props else "<a:br/>"
        pieces = []
        for index, part in enumerate(parts):
            if index:
                pieces.append(line_break)
            if part:
                pieces.append(f"<a:r>{props}<a:t{attrs}>{part}</a:t></a:r>")
        return "".join(pieces)

    return TEXT_RUN.sub(split, text), count


def used_typefaces(text: str) -> List[str]:
    return [html.unescape(match.group(2)) for match in TYPEFACE_ATTR.finditer(text)]


def postprocess_pptx(pptx: Path, transition: str, *, map_fonts: bool = True) -> Dict[str, Any]:
    """Rewrite the exported PPTX in place for PowerPoint/WPS compatibility."""
    temporary = pptx.with_name(f".{pptx.name}.{uuid.uuid4().hex}.tmp")
    slide_count = 0
    line_breaks = 0
    replaced: Dict[str, str] = {}
    typefaces: Dict[str, None] = {}
    try:
        with zipfile.ZipFile(pptx, "r") as source, zipfile.ZipFile(temporary, "w") as target:
            target.comment = source.comment
            for info in source.infolist():
                data = source.read(info.filename)
                name = info.filename
                if name.startswith("ppt/") and name.endswith(".xml"):
                    text = data.decode("utf-8")
                    updated = map_fonts_in_xml(text, replaced) if map_fonts else text
                    updated, breaks = split_newline_runs(updated)
                    line_breaks += breaks
                    if SLIDE_PART.fullmatch(name):
                        typefaces.update(dict.fromkeys(used_typefaces(updated)))
                    if updated != text:
                        data = updated.encode("utf-8")
                if SLIDE_PART.fullmatch(name):
                    data = replace_transition(data, transition)
                    slide_count += 1
                target.writestr(info, data, compress_type=info.compress_type)
        if slide_count == 0:
            raise ExportError("exported PPTX contains no slide XML")
        temporary.replace(pptx)
    finally:
        temporary.unlink(missing_ok=True)
    unusual = sorted(
        name
        for name in typefaces
        if name
        and not name.startswith("+")
        and name.lower() not in STANDARD_FONTS
        and name not in FONT_FALLBACKS
    )
    return {
        "slides": slide_count,
        "lineBreaks": line_breaks,
        "fontsReplaced": dict(sorted(replaced.items())),
        "nonStandardFonts": unusual,
    }


def patch_transitions(pptx: Path, transition: str) -> int:
    return postprocess_pptx(pptx, transition, map_fonts=False)["slides"]


def apply_postprocess(
    pptx: Path, transition: str, map_fonts: bool, summary: Dict[str, Any]
) -> None:
    result = postprocess_pptx(pptx, transition, map_fonts=map_fonts)
    summary["transitionPatchedSlides"] = result["slides"]
    if result["lineBreaks"]:
        summary["lineBreaks"] = result["lineBreaks"]
    if result["fontsReplaced"]:
        summary["fontsReplaced"] = result["fontsReplaced"]
        log(
            "replaced fonts that other computers lack: "
            + ", ".join(f"{old} -> {new}" for old, new in result["fontsReplaced"].items())
        )
    if result["nonStandardFonts"]:
        message = (
            "these fonts are not standard on Windows/Office, so PowerPoint on another "
            "computer may substitute them: " + ", ".join(result["nonStandardFonts"])
        )
        log(f"warning: {message}")
        summary.setdefault("warnings", []).append(message)


def verify_output(pptx: Path, transition: str, expect_fonts: bool) -> Dict[str, Any]:
    if not is_pptx(pptx):
        raise ExportError(f"output is not a valid PPTX ZIP: {pptx}")
    with zipfile.ZipFile(pptx) as archive:
        broken = archive.testzip()
        if broken:
            raise ExportError(f"PPTX CRC check failed at: {broken}")
        slide_names = [
            name
            for name in archive.namelist()
            if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)
        ]
        slide_xml = {name: archive.read(name) for name in slide_names}
        for data in slide_xml.values():
            validate_transition_order(data, transition)
        transition_hits = sum(has_direct_fade_transition(data) for data in slide_xml.values())
        if transition == "fade" and transition_hits != len(slide_names):
            raise ExportError("fade transition was not written to every slide")
        fonts = [
            name
            for name in archive.namelist()
            if name.startswith("ppt/fonts/") and not name.endswith("/")
        ]
        if expect_fonts and not fonts:
            log(
                "warning: embed-fonts was enabled, but the official writer produced no font part"
            )
        return {
            "slides": len(slide_names),
            "fadeTransitions": transition_hits,
            "fontParts": len(fonts),
            "bytes": pptx.stat().st_size,
        }


def resolve_editor_root() -> Path:
    """Locate the offline neo-ppt mirror (package editor/ or skill-installed copy)."""
    env = os.environ.get("OPEN_KIMI_PPT_EDITOR")
    candidates: List[Path] = []
    if env:
        candidates.append(Path(env).expanduser().resolve())
    candidates.append(SKILL_DIR / "editor")
    # monorepo / npm package: skills/onepaw-ppt → package root
    candidates.append(SKILL_DIR.parent.parent / "editor")
    for candidate in candidates:
        if (candidate / "index.html").is_file():
            return candidate
    raise ExportError(EDITOR_MISSING_HINT)


def serve(
    directory: Path,
    *,
    entry: str = "index.html",
) -> Tuple[ThreadingHTTPServer, threading.Thread, str]:
    """Legacy static serve (tests / callers). Prefer serve_local_editor for exports."""
    handler = lambda *args, **kwargs: QuietHandler(  # noqa: E731
        *args, directory=str(directory), **kwargs
    )
    server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    host, port = server.server_address
    return server, thread, f"http://{host}:{port}/{entry.lstrip('/')}"


def serve_local_editor(
    payload: Dict[str, Any],
) -> Tuple[ThreadingHTTPServer, threading.Thread, str]:
    """Serve the offline editor and inject payload.json for headless export."""
    editor_root = resolve_editor_root()
    payload_bytes = json.dumps(payload, ensure_ascii=False).encode("utf-8")

    class LocalEditorHandler(QuietHandler):
        def __init__(self, *args: Any, **kwargs: Any) -> None:
            super().__init__(*args, directory=str(editor_root), **kwargs)

        def _is_payload(self) -> bool:
            path = urlparse(self.path).path
            return path in ("/payload.json", "payload.json")

        def do_GET(self) -> None:  # noqa: N802
            if self._is_payload():
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(payload_bytes)))
                self.send_header("Cache-Control", "no-store")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(payload_bytes)
                return
            return SimpleHTTPRequestHandler.do_GET(self)

        def do_HEAD(self) -> None:  # noqa: N802
            if self._is_payload():
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(payload_bytes)))
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                return
            return SimpleHTTPRequestHandler.do_HEAD(self)

    server = ThreadingHTTPServer(("127.0.0.1", 0), LocalEditorHandler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    host, port = server.server_address
    url = f"http://{host}:{port}/?ndExport=1"
    log(f"local editor host: {url} (root={editor_root})")
    return server, thread, url


LOCAL_EXPORT_DIR = Path(__file__).resolve().parent / "local-export"
LOCAL_EXPORT_MJS = LOCAL_EXPORT_DIR / "export-pptd.mjs"
# Canonical patched WASM: editor/neo-ppt/assets/ (repo/npm). Skill install copies it
# into local-export/pptd_wasm_bg.wasm for ~/.agents|~/.claude skills trees.
CANONICAL_WASM_NAME = "pptd_wasm_bg-DPPWdROu.wasm"


def resolve_local_wasm() -> Path:
    package_root = Path(__file__).resolve().parents[3]
    candidates = [
        LOCAL_EXPORT_DIR / "pptd_wasm_bg.wasm",
        LOCAL_EXPORT_DIR / CANONICAL_WASM_NAME,
        SKILL_DIR / "editor" / "neo-ppt" / "assets" / CANONICAL_WASM_NAME,
        package_root / "editor" / "neo-ppt" / "assets" / CANONICAL_WASM_NAME,
    ]
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    raise ExportError(
        "patched WASM not found. Expected "
        f"editor/neo-ppt/assets/{CANONICAL_WASM_NAME} (repo/npm package) or "
        f"{LOCAL_EXPORT_DIR / 'pptd_wasm_bg.wasm'} (after skill install)."
    )


def export_pptx_local(
    source: Path,
    output: Path,
    transition: str,
    force: bool = False,
    map_fonts: bool = True,
) -> Dict[str, Any]:
    """Export via local patched official WASM (no cookie / no browser UI)."""
    if not LOCAL_EXPORT_MJS.is_file():
        raise ExportError(f"local exporter missing: {LOCAL_EXPORT_MJS}")
    wasm_path = resolve_local_wasm()
    node_command, node_env, node_label = node_runtime()

    manifest = find_manifest(source)
    output = output.expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists() and not force:
        raise ExportError(f"output already exists (pass --force to replace it): {output}")

    log(f"local WASM export with {node_label}: {manifest} -> {output}")
    log(f"defaults: transition={transition} (no-sign / signature bypassed)")

    # Pass project directory so media paths resolve relative to the deck root.
    project_dir = manifest.parent if manifest.is_file() else source
    cmd = [
        *node_command,
        str(LOCAL_EXPORT_MJS),
        str(project_dir),
        "-o",
        str(output),
        "--no-sign",
        "--transition",
        transition if transition in ("fade", "none") else "fade",
        "--wasm",
        str(wasm_path),
    ]

    # OKP_PYTHON lets the exporter's last-resort YAML fallback reuse this
    # interpreter. Node prints UTF-8; decode explicitly so zh-CN Windows (GBK)
    # cannot choke.
    process = subprocess.run(
        cmd,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        timeout=300,
        env={**os.environ, **node_env, "OKP_PYTHON": sys.executable},
    )
    if process.returncode != 0:
        raise ExportError(
            f"local WASM export failed ({process.returncode}):\n{process.stdout[-4000:]}"
        )

    summary: Dict[str, Any] = {}
    apply_postprocess(output, transition, map_fonts, summary)
    summary = {**verify_output(output, transition, expect_fonts=False), **summary}
    summary["output"] = str(output)
    summary["exporter"] = "local-wasm-patched"
    log(f"local export ok: {output} ({output.stat().st_size} bytes)")
    return summary


def export_pptx(
    source: Path,
    output: Path,
    transition: str,
    embed_fonts: bool,
    keep_download: bool = False,
    force: bool = False,
    prefer_local: bool = True,
    map_fonts: bool = True,
) -> Dict[str, Any]:
    """Prefer local patched WASM; fall back to the local editor in a headless browser."""
    if prefer_local:
        try:
            return export_pptx_local(source, output, transition, force=force, map_fonts=map_fonts)
        except ExportError as exc:
            log(f"local WASM export unavailable ({exc}); falling back to the local browser editor")

    manifest = find_manifest(source)
    payload = build_payload(manifest)
    output = output.expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists() and not force:
        raise ExportError(f"output already exists (pass --force to replace it): {output}")
    # Embedded copies of the editor fonts are ignored by PowerPoint; with font
    # substitution on they would also no longer be referenced.
    if map_fonts:
        embed_fonts = False

    log(f"manifest: {manifest}")
    log(f"defaults: transition={transition}, embed_fonts={'on' if embed_fonts else 'off'}")

    with temporary_directory(prefix="onepaw-ppt-export-") as temp_name:
        download_dir = Path(temp_name) / "downloads"
        download_dir.mkdir()
        with EditorHost(payload, font_fallbacks=map_fonts) as host:
            downloaded = host.export(download_dir, images=False, embed_fonts=embed_fonts, timeout=180)
        shutil.copy2(downloaded, output)
        if keep_download:
            debug_copy = output.with_name(f"{output.stem}.browser-raw.pptx")
            if debug_copy.exists() and not force:
                raise ExportError(f"raw debug output already exists (pass --force): {debug_copy}")
            shutil.copy2(downloaded, debug_copy)

    summary: Dict[str, Any] = {}
    apply_postprocess(output, transition, map_fonts, summary)
    summary = {**verify_output(output, transition, embed_fonts), **summary}
    summary["output"] = str(output)
    summary["exporter"] = "browser-local-editor"
    return summary


def make_output_robust() -> None:
    """Print UTF-8 into pipes and never crash on an unencodable character.

    Agent hosts read tool output from a pipe and decode it as UTF-8, while
    Python on zh-CN Windows would write GBK there, garbling Chinese paths.
    Interactive consoles keep their own encoding.
    """
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream.isatty():
                stream.reconfigure(errors="backslashreplace")
            else:
                stream.reconfigure(encoding="utf-8", errors="backslashreplace")
        except (AttributeError, ValueError, OSError):
            pass


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Export a PPTD project to PPTX. "
            "Default: local patched official WASM (offline). "
            "Optional --browser uses the local neo-ppt editor in a headless browser."
        )
    )
    parser.add_argument("input", type=Path, help=".pptd manifest or project directory")
    parser.add_argument("--output", "-o", type=Path, help="output .pptx path")
    parser.add_argument(
        "--transition",
        choices=("fade", "none"),
        default="fade",
        help="slide transition written to every slide (default: fade)",
    )
    font_group = parser.add_mutually_exclusive_group()
    font_group.add_argument(
        "--embed-fonts",
        dest="embed_fonts",
        action="store_true",
        default=True,
        help="embed editor fonts (browser path with --keep-fonts only; default)",
    )
    font_group.add_argument(
        "--no-embed-fonts",
        dest="embed_fonts",
        action="store_false",
        help="disable font embedding (browser path)",
    )
    parser.add_argument(
        "--keep-fonts",
        action="store_true",
        help=(
            "keep editor web fonts such as MiSans instead of replacing them with system "
            "fonts (the PPTX then only looks right where those fonts are installed)"
        ),
    )
    parser.add_argument(
        "--browser",
        action="store_true",
        help="force the local editor (headless browser) path instead of Node WASM",
    )
    parser.add_argument(
        "--keep-browser-raw",
        action="store_true",
        help="also keep the unpatched browser download beside the output",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="replace an existing output file",
    )
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    make_output_robust()
    args = parse_args(argv)
    try:
        manifest = find_manifest(args.input)
        output = args.output or manifest.with_suffix(".pptx")
        summary = export_pptx(
            args.input,
            output,
            args.transition,
            args.embed_fonts,
            args.keep_browser_raw,
            args.force,
            prefer_local=not args.browser,
            map_fonts=not args.keep_fonts,
        )
        warn_dropped_line_breaks(manifest, summary)
    except (ExportError, CDPError, OSError, subprocess.SubprocessError) as exc:
        print(f"onepaw-ppt export failed: {exc}", file=sys.stderr)
        return 1
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
