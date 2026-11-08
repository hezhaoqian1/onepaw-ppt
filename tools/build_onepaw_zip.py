#!/usr/bin/env python3
"""Build a skill zip that can be uploaded to OnePaw (QwenPaw) via "upload zip".

Usage:
    python tools/build_onepaw_zip.py [--output-dir DIR]

Output: dist/onepaw-ppt-<version>.zip, containing a single top-level
``onepaw-ppt/`` folder (SKILL.md + reference + scripts + editor + WASM).

The skill tree is staged with the project's own installer, so the zip has the
same layout as ``open-kimi-ppt-skill install``. On top of that, SKILL.md is
adapted for agents that do not install into ``~/.agents/skills``:

* ``python3 ~/.agents/skills/onepaw-ppt/scripts/...`` -> ``python scripts/...``
  (paths relative to the skill directory, which OnePaw shows the agent);
* ``npx open-kimi-ppt-skill serve`` -> ``python scripts/serve.py --open`` (bundled
  editor server, started in the background with Node.js or the desktop agent's
  own Electron runtime, so it works offline and without npm).

The editor's legacy (``nomodule``) bundle is left out: the export scripts and
the manual editor only run in current Chrome/Edge, which never load it. That
removes about a quarter of the files and most of the skill scanner's findings.
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
SKILL_NAME = "onepaw-ppt"

PATH_NOTE = """\
> **Skill paths (OnePaw / uploaded-zip build).** Every `scripts/...` path in this file is
> relative to this skill's own directory, whose absolute path is shown to you with the skill.
> Run commands from that directory, e.g. `cd <this skill dir> && python scripts/export_pptx.py ...`.
> Use `python` on Windows and `python3` on macOS/Linux.
> Desktop agents may run your shell with a trimmed PATH, so `node --version` can fail even though the
> scripts work: they find Node.js themselves, or use the desktop app's built-in Electron runtime.
> Rely on `python scripts/doctor.py` instead of checking `node`/`npm` yourself, and never ask the
> user to install Node.js unless doctor reports that PPTX export fails.
> Image QA and PPTX export can take a few minutes on slow computers: when your shell tool accepts a
> timeout, pass at least 300 seconds for `export_images.py`, `export_pptx.py` and `doctor.py`.
> `python scripts/serve.py --open` starts the manual editor in the background and returns at once.
"""

SERVE_MJS = """\
#!/usr/bin/env node
// Local PPTD editor bundled with this skill (same as `open-kimi-ppt-skill serve`).
// Usage: node scripts/serve.mjs [--port 55173] [--open]
import { spawn } from "node:child_process";
import { startEditorServer } from "./editor-server.mjs";

let port = 55173;
let open = false;
const args = process.argv.slice(2);
for (let index = 0; index < args.length; index += 1) {
  if (args[index] === "--port") {
    port = Number(args[++index]);
    if (!Number.isInteger(port) || port < 1 || port > 65_535) {
      console.error("Error: --port must be an integer between 1 and 65535");
      process.exit(1);
    }
  } else if (args[index] === "--open") {
    open = true;
  } else {
    console.error(`Error: unknown argument: ${args[index]}`);
    process.exit(1);
  }
}

const { server, url } = await startEditorServer({ port });
console.log(`OnePaw PPT editor is running at ${url}`);
console.log("Press Ctrl+C to stop the server.");

if (open) {
  const command = process.platform === "darwin" ? "open" : process.platform === "win32" ? "cmd" : "xdg-open";
  const commandArgs = process.platform === "win32" ? ["/c", "start", "", url] : [url];
  const child = spawn(command, commandArgs, { detached: true, stdio: "ignore" });
  child.on("error", (error) => console.warn(`Could not open the browser: ${error.message}`));
  child.unref();
}

const shutdown = () => server.close(() => process.exit(0));
process.once("SIGINT", shutdown);
process.once("SIGTERM", shutdown);
"""

SERVE_PY = '''\
#!/usr/bin/env python3
"""Start the bundled PPTD editor in the background and print its URL.

Usage: python scripts/serve.py [--port 55173] [--open]

Returns as soon as the editor answers; the server keeps running. Node.js 18+ is
used when installed, otherwise the desktop agent's Electron runtime (same
lookup as export_pptx.py). Re-running it reuses an editor already on the port.
"""

import argparse
import os
import subprocess
import sys
import tempfile
import time
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from export_pptx import ExportError, make_output_robust, node_runtime  # noqa: E402

LOCAL = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def answering(url):
    try:
        with LOCAL.open(url, timeout=2) as response:
            return response.status == 200
    except OSError:
        return False


def main():
    make_output_robust()
    parser = argparse.ArgumentParser(description="Start the bundled PPTD editor in the background.")
    parser.add_argument("--port", type=int, default=55173)
    parser.add_argument("--open", action="store_true", help="open the editor in the default browser")
    args = parser.parse_args()
    url = f"http://127.0.0.1:{args.port}/"
    if answering(url):
        print(f"OnePaw PPT editor is already running at {url}")
        return 0
    try:
        command, env, label = node_runtime()
    except ExportError as exc:
        print(f"cannot start the editor: {exc}", file=sys.stderr)
        return 1
    script = Path(__file__).resolve().with_name("serve.mjs")
    argv = [*command, str(script), "--port", str(args.port)] + (["--open"] if args.open else [])
    log_path = Path(tempfile.gettempdir()) / f"onepaw-ppt-editor-{args.port}.log"
    options = {}
    if sys.platform == "win32":
        options["creationflags"] = (
            subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.CREATE_NO_WINDOW
        )
    else:
        options["start_new_session"] = True
    with log_path.open("w", encoding="utf-8") as log:
        subprocess.Popen(
            argv,
            stdin=subprocess.DEVNULL,
            stdout=log,
            stderr=subprocess.STDOUT,
            env={**os.environ, **env},
            close_fds=True,
            **options,
        )
    deadline = time.monotonic() + 20
    while time.monotonic() < deadline:
        if answering(url):
            print(f"OnePaw PPT editor is running at {url} ({label})")
            return 0
        time.sleep(0.3)
    print(f"the editor did not start; log: {log_path}", file=sys.stderr)
    print(log_path.read_text(encoding="utf-8", errors="replace")[-2000:], file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
'''


def stage_skill(staging: Path) -> Path:
    node = shutil.which("node")
    if not node:
        sys.exit("node (18+) is required to stage the skill")
    subprocess.run(
        [node, str(REPO / "bin" / "open-kimi-ppt-skill.js"), "install", "--target", str(staging)],
        check=True,
    )
    return staging / SKILL_NAME


def adapt_skill(skill: Path) -> None:
    # Developer tests are not needed at runtime and only add files for the scanner.
    shutil.rmtree(skill / "tests", ignore_errors=True)

    # Bundle the editor server; it resolves ../editor, i.e. <skill>/editor.
    shutil.copyfile(REPO / "lib" / "editor-server.js", skill / "scripts" / "editor-server.mjs")
    (skill / "scripts" / "serve.mjs").write_text(SERVE_MJS, encoding="utf-8", newline="\n")
    (skill / "scripts" / "serve.py").write_text(SERVE_PY, encoding="utf-8", newline="\n")

    # Legacy (nomodule) editor chunks: never loaded by current Chrome/Edge.
    for legacy in (skill / "editor").rglob("*-legacy-*"):
        legacy.unlink()

    skill_md = skill / "SKILL.md"
    # Normalise CRLF (git autocrlf checkouts on Windows) so the front matter parses.
    text = skill_md.read_text(encoding="utf-8").replace("\r\n", "\n")
    text = text.replace(f"python3 ~/.agents/skills/{SKILL_NAME}/scripts/", "python scripts/")
    text = text.replace("npx open-kimi-ppt-skill serve", "python scripts/serve.py --open")

    # Insert the path note right after the YAML front matter.
    if not text.startswith("---\n"):
        sys.exit("SKILL.md must start with YAML front matter")
    end = text.index("\n---\n", 4) + len("\n---\n")
    text = f"{text[:end]}\n{PATH_NOTE}{text[end:]}"

    leftovers = [marker for marker in ("~/.agents/skills", "npx open-kimi-ppt-skill") if marker in text]
    if leftovers:
        sys.exit(f"SKILL.md still references {leftovers}; update adapt_skill()")
    skill_md.write_text(text, encoding="utf-8", newline="\n")


def write_zip(skill: Path, destination: Path) -> int:
    destination.parent.mkdir(parents=True, exist_ok=True)
    # Skip local caches (e.g. __pycache__ or .ruff_cache left behind by tests and linters).
    files = sorted(
        path
        for path in skill.rglob("*")
        if path.is_file()
        and path.suffix != ".pyc"
        and not {"__pycache__", ".pytest_cache", ".ruff_cache"} & set(path.relative_to(skill).parts)
    )
    with zipfile.ZipFile(destination, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for path in files:
            archive.write(path, f"{SKILL_NAME}/{path.relative_to(skill).as_posix()}")
    return len(files)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--output-dir", type=Path, default=REPO / "dist")
    args = parser.parse_args()

    version = json.loads((REPO / "package.json").read_text(encoding="utf-8"))["version"]
    destination = args.output_dir.resolve() / f"onepaw-ppt-{version}.zip"

    with tempfile.TemporaryDirectory(prefix="onepaw-ppt-") as staging:
        skill = stage_skill(Path(staging))
        adapt_skill(skill)
        count = write_zip(skill, destination)

    size_mb = destination.stat().st_size / (1024 * 1024)
    print(f"Built {destination} ({count} files, {size_mb:.1f} MB)")


if __name__ == "__main__":
    main()
