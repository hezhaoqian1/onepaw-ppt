#!/usr/bin/env python3
"""Drive a headless Chrome / Edge through the DevTools protocol (stdlib only).

The export scripts use this to operate the local neo-ppt editor: open it, click
its export dialog, and collect the download. Nothing outside the standard
library is needed (no agent-browser, npm, websocket-client or Playwright), so
it works with embedded Pythons such as the one bundled with desktop agents.

The browser runs headless with a throwaway profile, so it never shows a window
and never touches the user's own browser profile.
"""

from __future__ import annotations

import base64
import json
import os
import shutil
import socket
import struct
import subprocess
import sys
import tempfile
import time
import urllib.parse
from pathlib import Path
from typing import Any, Callable, Dict, Iterator, List, Optional


class CDPError(RuntimeError):
    pass


# ---------------------------------------------------------------------------
# Minimal RFC 6455 client: enough for a local, unencrypted DevTools endpoint.
# ---------------------------------------------------------------------------


def _mask(data: bytes, key: bytes) -> bytes:
    if not data:
        return b""
    size = len(data)
    pad = (key * (size // 4 + 1))[:size]
    return (int.from_bytes(data, "big") ^ int.from_bytes(pad, "big")).to_bytes(size, "big")


class WebSocket:
    def __init__(self, url: str, timeout: float = 30.0) -> None:
        parsed = urllib.parse.urlparse(url)
        if parsed.scheme != "ws" or not parsed.hostname:
            raise CDPError(f"unsupported DevTools URL: {url}")
        host = parsed.hostname
        port = parsed.port or 80
        self.sock = socket.create_connection((host, port), timeout=timeout)
        self.buffer = bytearray()
        key = base64.b64encode(os.urandom(16)).decode("ascii")
        path = parsed.path or "/"
        if parsed.query:
            path += "?" + parsed.query
        request = (
            f"GET {path} HTTP/1.1\r\n"
            f"Host: {host}:{port}\r\n"
            "Upgrade: websocket\r\n"
            "Connection: Upgrade\r\n"
            f"Sec-WebSocket-Key: {key}\r\n"
            "Sec-WebSocket-Version: 13\r\n\r\n"
        )
        self.sock.sendall(request.encode("ascii"))
        while b"\r\n\r\n" not in self.buffer:
            chunk = self.sock.recv(4096)
            if not chunk:
                raise CDPError("DevTools websocket closed during handshake")
            self.buffer.extend(chunk)
        end = self.buffer.index(b"\r\n\r\n")
        status = bytes(self.buffer[: self.buffer.index(b"\r\n")])
        del self.buffer[: end + 4]
        if b" 101 " not in status + b" ":
            raise CDPError(f"DevTools websocket handshake failed: {status!r}")

    def _read(self, size: int) -> bytes:
        while len(self.buffer) < size:
            chunk = self.sock.recv(max(65536, size - len(self.buffer)))
            if not chunk:
                raise CDPError("DevTools websocket closed")
            self.buffer.extend(chunk)
        data = bytes(self.buffer[:size])
        del self.buffer[:size]
        return data

    def _send_frame(self, opcode: int, payload: bytes) -> None:
        header = bytearray([0x80 | opcode])
        size = len(payload)
        if size < 126:
            header.append(0x80 | size)
        elif size < 65536:
            header.append(0x80 | 126)
            header += struct.pack(">H", size)
        else:
            header.append(0x80 | 127)
            header += struct.pack(">Q", size)
        key = os.urandom(4)
        self.sock.sendall(bytes(header) + key + _mask(payload, key))

    def send_text(self, text: str) -> None:
        self._send_frame(0x1, text.encode("utf-8"))

    def recv_text(self) -> str:
        message = bytearray()
        while True:
            first, second = self._read(2)
            opcode = first & 0x0F
            size = second & 0x7F
            if size == 126:
                (size,) = struct.unpack(">H", self._read(2))
            elif size == 127:
                (size,) = struct.unpack(">Q", self._read(8))
            key = self._read(4) if second & 0x80 else None
            payload = self._read(size)
            if key:
                payload = _mask(payload, key)
            if opcode == 0x8:
                raise CDPError("browser closed the DevTools connection")
            if opcode == 0x9:
                self._send_frame(0xA, payload)
                continue
            if opcode == 0xA:
                continue
            message.extend(payload)
            if first & 0x80:
                return message.decode("utf-8", errors="replace")

    def close(self) -> None:
        try:
            self._send_frame(0x8, b"")
        except OSError:
            pass
        try:
            self.sock.close()
        except OSError:
            pass


class CDP:
    """One browser-level DevTools connection (flattened sessions)."""

    def __init__(self, ws_url: str) -> None:
        self.ws = WebSocket(ws_url)
        self.next_id = 0
        self.events: List[Dict[str, Any]] = []

    def _receive(self, deadline: float, what: str) -> Dict[str, Any]:
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise CDPError(f"timed out waiting for {what}")
        self.ws.sock.settimeout(remaining)
        try:
            return json.loads(self.ws.recv_text())
        except socket.timeout as exc:
            # A partial frame may be buffered now, so the connection is unusable.
            raise CDPError(f"timed out waiting for {what}") from exc

    def call(
        self,
        method: str,
        params: Optional[Dict[str, Any]] = None,
        session: Optional[str] = None,
        timeout: float = 60.0,
    ) -> Dict[str, Any]:
        self.next_id += 1
        message_id = self.next_id
        message: Dict[str, Any] = {"id": message_id, "method": method, "params": params or {}}
        if session:
            message["sessionId"] = session
        self.ws.send_text(json.dumps(message))
        deadline = time.monotonic() + timeout
        while True:
            data = self._receive(deadline, method)
            if data.get("id") == message_id:
                if "error" in data:
                    raise CDPError(f"{method} failed: {data['error']}")
                return data.get("result", {})
            if "method" in data:
                self.events.append(data)

    def wait_event(
        self,
        predicate: Callable[[Dict[str, Any]], bool],
        timeout: float,
        what: str,
    ) -> Dict[str, Any]:
        deadline = time.monotonic() + timeout
        while True:
            for index, event in enumerate(self.events):
                if predicate(event):
                    return self.events.pop(index)
            data = self._receive(deadline, what)
            if "method" in data:
                self.events.append(data)

    def close(self) -> None:
        self.ws.close()


# ---------------------------------------------------------------------------
# Browser discovery and launch
# ---------------------------------------------------------------------------


def _windows_app_paths(executable: str) -> Iterator[str]:
    try:
        import winreg
    except ImportError:
        return
    key = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{executable}"
    for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        try:
            with winreg.OpenKey(hive, key) as handle:
                value, _ = winreg.QueryValueEx(handle, "")
        except OSError:
            continue
        if value:
            yield os.path.expandvars(str(value)).strip('"')


def browser_candidates() -> List[str]:
    """Chromium browsers that can render the editor, most preferred first."""
    found: List[str] = []
    explicit = os.environ.get("OKP_BROWSER")
    if explicit:
        found.append(explicit)
    if sys.platform == "win32":
        bases = [
            os.environ.get(name)
            for name in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA", "ProgramW6432")
        ]
        relative = (
            r"Google\Chrome\Application\chrome.exe",
            r"Microsoft\Edge\Application\msedge.exe",
            r"Chromium\Application\chrome.exe",
        )
        for rel in relative:
            for base in bases:
                if base:
                    found.append(os.path.join(base, rel))
        for exe in ("chrome.exe", "msedge.exe"):
            found.extend(_windows_app_paths(exe))
    elif sys.platform == "darwin":
        for app, binary in (
            ("Google Chrome", "Google Chrome"),
            ("Microsoft Edge", "Microsoft Edge"),
            ("Chromium", "Chromium"),
            ("Brave Browser", "Brave Browser"),
        ):
            for root in ("/Applications", str(Path.home() / "Applications")):
                found.append(f"{root}/{app}.app/Contents/MacOS/{binary}")
    else:
        for name in (
            "google-chrome",
            "google-chrome-stable",
            "chromium",
            "chromium-browser",
            "microsoft-edge",
            "microsoft-edge-stable",
        ):
            path = shutil.which(name)
            if path:
                found.append(path)
    unique: List[str] = []
    seen = set()
    for path in found:
        key = os.path.normcase(os.path.abspath(path))
        if key in seen or not os.path.isfile(path):
            continue
        seen.add(key)
        unique.append(path)
    return unique


class Browser:
    """A headless browser with a temporary profile, closed by close()."""

    def __init__(self, executable: str, timeout: float = 30.0) -> None:
        self.executable = executable
        self.profile = Path(tempfile.mkdtemp(prefix="okp-browser-"))
        args = [
            executable,
            "--headless=new",
            f"--user-data-dir={self.profile}",
            "--remote-debugging-port=0",
            "--remote-allow-origins=*",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-extensions",
            "--disable-sync",
            "--disable-background-networking",
            "--disable-component-update",
            "--disable-features=Translate,MediaRouter",
            "--mute-audio",
            "--hide-scrollbars",
            "--window-size=1280,720",
            "about:blank",
        ]
        flags = 0
        if sys.platform == "win32":
            flags = subprocess.CREATE_NO_WINDOW | subprocess.CREATE_NEW_PROCESS_GROUP
        self.process = subprocess.Popen(
            args,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=flags,
        )
        self.cdp: Optional[CDP] = None
        try:
            self.cdp = CDP(self._wait_for_endpoint(timeout))
        except Exception:
            self.close()
            raise

    def _wait_for_endpoint(self, timeout: float) -> str:
        port_file = self.profile / "DevToolsActivePort"
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if self.process.poll() is not None:
                raise CDPError(
                    f"browser exited during startup (code {self.process.returncode}): "
                    f"{self.executable}"
                )
            try:
                lines = port_file.read_text(encoding="utf-8").split()
            except OSError:
                lines = []
            if len(lines) >= 2:
                return f"ws://127.0.0.1:{lines[0]}{lines[1]}"
            time.sleep(0.2)
        raise CDPError(
            f"browser did not open a DevTools port within {timeout:.0f}s "
            f"(remote debugging may be disabled by policy): {self.executable}"
        )

    def close(self) -> None:
        if self.cdp is not None:
            try:
                self.cdp.call("Browser.close", timeout=5)
            except (CDPError, OSError):
                pass
            self.cdp.close()
            self.cdp = None
        try:
            self.process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            self.process.kill()
            try:
                self.process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                pass
        for _ in range(10):
            shutil.rmtree(self.profile, ignore_errors=True)
            if not self.profile.exists():
                break
            time.sleep(0.3)


def launch_browser(timeout: float = 30.0) -> Browser:
    """Start the first candidate browser that exposes a DevTools endpoint."""
    candidates = browser_candidates()
    if not candidates:
        raise CDPError(
            "no Chrome or Edge found. Install Google Chrome or Microsoft Edge, or set "
            "OKP_BROWSER to the browser executable."
        )
    errors = []
    for executable in candidates:
        try:
            return Browser(executable, timeout=timeout)
        except (CDPError, OSError) as exc:
            errors.append(f"{executable}: {exc}")
    raise CDPError("could not start a headless browser:\n  " + "\n  ".join(errors))


# ---------------------------------------------------------------------------
# Page helpers
# ---------------------------------------------------------------------------

# Finds a visible clickable element by its exact label and returns the centre of
# its box. Buttons win over plain elements; the last match wins so dialogs that
# are appended to <body> beat the toolbar underneath.
FIND_ELEMENT_JS = r"""
((label, scope) => {
  const norm = (s) => (s || '').replace(/\s+/g, ' ').trim();
  const visible = (el) => {
    const r = el.getBoundingClientRect();
    if (r.width <= 0 || r.height <= 0) return false;
    const st = getComputedStyle(el);
    return st.visibility !== 'hidden' && st.display !== 'none' && st.pointerEvents !== 'none';
  };
  let root = document;
  if (scope) {
    const dialogs = [...document.querySelectorAll(scope)].filter(visible);
    if (!dialogs.length) return null;
    root = dialogs[dialogs.length - 1];
  }
  const named = (el) =>
    norm(el.innerText) === label || norm(el.getAttribute('aria-label')) === label || norm(el.title) === label;
  const buttons = [...root.querySelectorAll('button,[role=button],[role=menuitem],[role=tab],[role=radio]')]
    .filter((el) => visible(el) && named(el));
  const leaves = [...root.querySelectorAll('*')]
    .filter((el) => el.children.length === 0 && visible(el) && norm(el.textContent) === label);
  const el = buttons[buttons.length - 1] || leaves[leaves.length - 1];
  if (!el) return null;
  if (el.disabled || el.getAttribute('aria-disabled') === 'true') return { disabled: true };
  el.scrollIntoView({ block: 'center', inline: 'center' });
  const r = el.getBoundingClientRect();
  return { x: r.left + r.width / 2, y: r.top + r.height / 2 };
})
"""


class Page:
    """One tab in a Browser, with helpers for the editor automation."""

    def __init__(self, browser: Browser, width: int = 1280, height: int = 720) -> None:
        if browser.cdp is None:
            raise CDPError("browser is closed")
        self.cdp = browser.cdp
        target = self.cdp.call("Target.createTarget", {"url": "about:blank"})["targetId"]
        self.target = target
        self.session = self.cdp.call(
            "Target.attachToTarget", {"targetId": target, "flatten": True}
        )["sessionId"]
        self.call("Page.enable")
        self.call("Runtime.enable")
        self.call(
            "Emulation.setDeviceMetricsOverride",
            {"width": width, "height": height, "deviceScaleFactor": 1, "mobile": False},
        )

    def call(self, method: str, params: Optional[Dict[str, Any]] = None, timeout: float = 60.0) -> Dict[str, Any]:
        return self.cdp.call(method, params, session=self.session, timeout=timeout)

    def navigate(self, url: str, timeout: float = 60.0) -> None:
        result = self.call("Page.navigate", {"url": url}, timeout=timeout)
        if result.get("errorText"):
            raise CDPError(f"navigation to {url} failed: {result['errorText']}")

    def evaluate(self, expression: str, timeout: float = 30.0) -> Any:
        result = self.call(
            "Runtime.evaluate",
            {"expression": expression, "returnByValue": True, "awaitPromise": True},
            timeout=timeout,
        )
        if result.get("exceptionDetails"):
            details = result["exceptionDetails"]
            text = details.get("exception", {}).get("description") or details.get("text")
            raise CDPError(f"page script failed: {text}")
        return result.get("result", {}).get("value")

    def wait_for(self, expression: str, timeout: float, what: str, interval: float = 0.25) -> Any:
        deadline = time.monotonic() + timeout
        while True:
            value = self.evaluate(expression)
            if value:
                return value
            if time.monotonic() > deadline:
                raise CDPError(f"timed out after {timeout:.0f}s waiting for {what}")
            time.sleep(interval)

    def click_at(self, x: float, y: float) -> None:
        for kind, buttons in (("mouseMoved", 0), ("mousePressed", 1), ("mouseReleased", 0)):
            params: Dict[str, Any] = {"type": kind, "x": x, "y": y, "button": "left", "buttons": buttons}
            if kind != "mouseMoved":
                params["clickCount"] = 1
            self.call("Input.dispatchMouseEvent", params)

    def click_text(self, label: str, scope: Optional[str] = None, timeout: float = 20.0) -> None:
        """Click the visible element labelled ``label`` (optionally inside ``scope``)."""
        expression = f"{FIND_ELEMENT_JS}({json.dumps(label)}, {json.dumps(scope)})"
        deadline = time.monotonic() + timeout
        while True:
            point = self.evaluate(expression)
            if point and not point.get("disabled"):
                self.click_at(point["x"], point["y"])
                return
            if time.monotonic() > deadline:
                state = "disabled" if point else "not found"
                raise CDPError(f"could not click {label!r}: element {state}")
            time.sleep(0.3)

    def screenshot(self, path: Path, *, quality: int = 85, full_page: bool = False) -> Path:
        params: Dict[str, Any] = {"format": "jpeg", "quality": quality}
        if full_page:
            metrics = self.call("Page.getLayoutMetrics")
            size = metrics.get("cssContentSize") or metrics.get("contentSize")
            params["clip"] = {
                "x": 0,
                "y": 0,
                "width": size["width"],
                "height": size["height"],
                "scale": 1,
            }
            params["captureBeyondViewport"] = True
        data = self.call("Page.captureScreenshot", params, timeout=120)["data"]
        path.write_bytes(base64.b64decode(data))
        return path

    def allow_downloads(self, directory: Path) -> None:
        self.cdp.call(
            "Browser.setDownloadBehavior",
            {"behavior": "allow", "downloadPath": str(directory), "eventsEnabled": True},
        )
