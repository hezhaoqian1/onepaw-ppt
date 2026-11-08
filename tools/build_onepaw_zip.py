#!/usr/bin/env python3
"""Build a skill zip that can be uploaded to OnePaw (QwenPaw) via "upload zip".

Usage:
    python tools/build_onepaw_zip.py [--output-dir DIR]

Output: dist/onepaw-ppt-<version>.zip, containing a single top-level
``open-kimi-ppt/`` folder (SKILL.md + reference + scripts + editor + WASM).

The skill tree is staged with the project's own installer, so the zip has the
same layout as ``open-kimi-ppt-skill install``. On top of that, SKILL.md is
adapted for agents that do not install into ``~/.agents/skills``:

* ``python3 ~/.agents/skills/open-kimi-ppt/scripts/...`` -> ``python scripts/...``
  (paths relative to the skill directory, which OnePaw shows the agent);
* ``npx open-kimi-ppt-skill serve`` -> ``node scripts/serve.mjs`` (bundled, so
  the editor works offline and uses this build instead of the npm package).
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
SKILL_NAME = "open-kimi-ppt"

PATH_NOTE = """\
> **Skill paths (OnePaw / uploaded-zip build).** Every `scripts/...` path in this file is
> relative to this skill's own directory, whose absolute path is shown to you with the skill.
> Run commands from that directory, e.g. `cd <this skill dir> && python scripts/export_pptx.py ...`.
> Use `python` on Windows and `python3` on macOS/Linux. `node scripts/serve.mjs` runs until it is
> stopped, so start it as a background/detached process (Windows cmd: `start "" node scripts/serve.mjs --open`;
> macOS/Linux: `nohup node scripts/serve.mjs --open >/dev/null 2>&1 &`) instead of waiting on it.
> Desktop agents may run your shell with a trimmed PATH, so `node --version` can fail even though Node.js
> is installed. Before reporting Node as missing in step0, look for it (Windows: `C:\\Program Files\\nodejs\\node.exe`;
> macOS: `/opt/homebrew/bin/node`, `/usr/local/bin/node`) and prepend its folder to PATH in your commands.
> The export scripts find Node in these places by themselves.
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
console.log(`Open Kimi PPT editor is running at ${url}`);
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

    skill_md = skill / "SKILL.md"
    # Normalise CRLF (git autocrlf checkouts on Windows) so the front matter parses.
    text = skill_md.read_text(encoding="utf-8").replace("\r\n", "\n")
    text = text.replace(f"python3 ~/.agents/skills/{SKILL_NAME}/scripts/", "python scripts/")
    text = text.replace("npx open-kimi-ppt-skill serve", "node scripts/serve.mjs")

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
    # Skip local caches (e.g. __pycache__ left behind by running the tests).
    files = sorted(
        path
        for path in skill.rglob("*")
        if path.is_file()
        and path.suffix != ".pyc"
        and not {"__pycache__", ".pytest_cache"} & set(path.relative_to(skill).parts)
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
