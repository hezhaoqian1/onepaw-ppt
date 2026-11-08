#!/usr/bin/env python3
import importlib.util
import os
import tempfile
import time
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch


SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "export_pptx.py"
SPEC = importlib.util.spec_from_file_location("export_pptx", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC and SPEC.loader
SPEC.loader.exec_module(MODULE)


class ExportPptxTests(unittest.TestCase):
    def setUp(self):
        MODULE._NODE = None  # node_runtime() caches its result

    @patch.object(MODULE, "find_electron", return_value=None)
    @patch.object(MODULE, "run_command")
    @patch.object(MODULE.shutil, "which", return_value=None)
    def test_missing_nodejs_raises_clear_error(self, which, run_command, find_electron):
        with self.assertRaisesRegex(MODULE.ExportError, "Node.js 18\\+ was not found"):
            MODULE.ensure_nodejs()
        run_command.assert_not_called()

    @patch.object(MODULE, "find_electron", return_value=None)
    @patch.object(MODULE, "run_command")
    @patch.object(MODULE.shutil, "which", return_value="/bin/node")
    def test_old_nodejs_raises_clear_error(self, which, run_command, find_electron):
        run_command.return_value = MODULE.subprocess.CompletedProcess([], 0, "v16.20.2\n")
        with self.assertRaisesRegex(MODULE.ExportError, "16.20.2 at /bin/node is older than 18"):
            MODULE.ensure_nodejs()

    @patch.object(MODULE, "run_command")
    @patch.object(MODULE.shutil, "which", return_value="/bin/node")
    def test_installed_node_is_preferred(self, which, run_command):
        run_command.return_value = MODULE.subprocess.CompletedProcess([], 0, "v22.11.0\n")
        command, env, label = MODULE.node_runtime()
        self.assertEqual((command, env), (["/bin/node"], {}))
        self.assertIn("22.11.0", label)

    @patch.object(MODULE, "run_command")
    @patch.object(MODULE.shutil, "which", return_value=None)
    def test_electron_runtime_stands_in_for_missing_node(self, which, run_command):
        electron = Path("/apps/OnePaw/OnePaw.exe")
        run_command.return_value = MODULE.subprocess.CompletedProcess([], 0, "18.18.2")
        with patch.object(MODULE, "find_electron", return_value=electron):
            command, env, label = MODULE.node_runtime()
        self.assertEqual(command, [str(electron)])
        self.assertEqual(env, {"ELECTRON_RUN_AS_NODE": "1"})
        self.assertEqual(run_command.call_args.kwargs["env"]["ELECTRON_RUN_AS_NODE"], "1")
        self.assertIn("18.18.2", label)

    def test_find_electron_locates_app_around_bundled_python(self):
        with tempfile.TemporaryDirectory() as name:
            app = Path(name) / "OnePaw 智能助手"
            python = app / "resources" / "portable_py311" / "python.exe"
            python.parent.mkdir(parents=True)
            python.write_bytes(b"py")
            (app / "LICENSE.electron.txt").write_text("MIT")
            (app / "OnePaw.exe").write_bytes(b"x" * 4096)
            (app / "Uninstall OnePaw.exe").write_bytes(b"x" * 8192)
            (app / "elevate.exe").write_bytes(b"x" * 8192)
            with patch.object(MODULE.sys, "platform", "win32"), \
                    patch.object(MODULE.sys, "executable", str(python)), \
                    patch.dict(MODULE.os.environ, {}, clear=False):
                MODULE.os.environ.pop("OKP_ELECTRON", None)
                self.assertEqual(MODULE.find_electron(), (app / "OnePaw.exe").resolve())

    def test_find_electron_ignores_plain_python_installs(self):
        with tempfile.TemporaryDirectory() as name:
            python = Path(name) / "Python311" / "python.exe"
            python.parent.mkdir(parents=True)
            python.write_bytes(b"py")
            with patch.object(MODULE.sys, "executable", str(python)), \
                    patch.dict(MODULE.os.environ, {}, clear=False):
                MODULE.os.environ.pop("OKP_ELECTRON", None)
                self.assertIsNone(MODULE.find_electron())

    def test_parse_node_version(self):
        self.assertEqual(MODULE.parse_node_version("v22.11.0"), (22, 11, 0))
        self.assertEqual(MODULE.parse_node_version("18.20.4"), (18, 20, 4))

    def test_fade_is_inserted_before_timing(self):
        source = (
            b'<?xml version="1.0" encoding="UTF-8"?>'
            b'<p:sld xmlns:p="urn:test"><p:cSld><p:spTree><p:extLst/>'
            b'</p:spTree></p:cSld><p:clrMapOvr/><p:timing/><p:extLst/></p:sld>'
        )
        result_bytes = MODULE.replace_transition(source, "fade")
        result = result_bytes.decode("utf-8")
        self.assertIn("<p:transition", result)
        self.assertIn("<p:fade/>", result)
        self.assertGreater(result.index("<p:transition"), result.index("<p:clrMapOvr"))
        self.assertLess(result.index("<p:transition"), result.index("<p:timing"))
        MODULE.validate_transition_order(result_bytes, "fade")

    def test_existing_transition_is_replaced_or_removed(self):
        source = (
            b'<p:sld xmlns:p="urn:test"><p:cSld/>'
            b'<p:transition><p:wipe/></p:transition><p:extLst/></p:sld>'
        )
        faded = MODULE.replace_transition(source, "fade").decode("utf-8")
        self.assertNotIn("p:wipe", faded)
        self.assertEqual(faded.count("<p:transition"), 1)
        MODULE.validate_transition_order(faded.encode("utf-8"), "fade")
        cleared = MODULE.replace_transition(source, "none").decode("utf-8")
        self.assertNotIn("p:transition", cleared)
        MODULE.validate_transition_order(cleared.encode("utf-8"), "none")

    def test_nested_transition_is_relocated_to_slide_root(self):
        source = (
            b'<p:sld xmlns:p="urn:test"><p:cSld><p:spTree>'
            b'<p:transition><p:fade/></p:transition><p:extLst/>'
            b'</p:spTree></p:cSld><p:clrMapOvr/><p:extLst/></p:sld>'
        )
        result = MODULE.replace_transition(source, "fade")
        MODULE.validate_transition_order(result, "fade")
        self.assertEqual(MODULE.root_child_names(result), [
            "cSld", "clrMapOvr", "transition", "extLst"
        ])

    def test_patch_transitions_preserves_a_valid_zip(self):
        with tempfile.TemporaryDirectory() as name:
            deck = Path(name) / "test.pptx"
            with zipfile.ZipFile(deck, "w", zipfile.ZIP_DEFLATED) as archive:
                archive.writestr(
                    "[Content_Types].xml",
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                    '<Override PartName="/ppt/presentation.xml" '
                    f'ContentType="{MODULE.PPTX_CONTENT_TYPE}"/></Types>',
                )
                archive.writestr("ppt/presentation.xml", "<p:presentation xmlns:p=\"urn:test\"/>")
                archive.writestr(
                    "ppt/slides/slide1.xml",
                    '<p:sld xmlns:p="urn:test"><p:cSld/></p:sld>',
                )
            self.assertEqual(MODULE.patch_transitions(deck, "fade"), 1)
            with zipfile.ZipFile(deck) as archive:
                self.assertIsNone(archive.testzip())
                slide = archive.read("ppt/slides/slide1.xml")
                self.assertIn(b"<p:fade/>", slide)

    @patch.object(MODULE.subprocess, "call", return_value=0)
    def test_run_command_captures_utf8_via_temp_file(self, call):
        def write_sink(*_args, **kwargs):
            kwargs["stdout"].write("v22.11.0 · 中文\n")
            return 0

        call.side_effect = write_sink
        process = MODULE.run_command(["node", "--version"], timeout=5)
        self.assertEqual(process.returncode, 0)
        self.assertIn("22.11.0", process.stdout)
        self.assertEqual(call.call_args.kwargs["stderr"], MODULE.subprocess.STDOUT)

    def test_find_download_ignores_files_older_than_since(self):
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            old = root / "old.pptx"
            new = root / "new.pptx"
            for path in (old, new):
                with zipfile.ZipFile(path, "w") as archive:
                    archive.writestr(
                        "[Content_Types].xml",
                        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                        '<Override PartName="/ppt/presentation.xml" '
                        f'ContentType="{MODULE.PPTX_CONTENT_TYPE}"/></Types>',
                    )
                    archive.writestr("ppt/presentation.xml", "<p:presentation/>")

            older = time.time() - 60
            os.utime(old, (older, older))
            since = time.time() - 5
            found = MODULE.find_download([root], timeout=2.0, since=since)
            self.assertEqual(found.resolve(), new.resolve())

    def test_find_download_survives_files_vanishing_mid_scan(self):
        # Chrome renames "*.crdownload" files away between directory listing
        # and stat(); a vanished file must be skipped, not crash the export.
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            deck = root / "deck.pptx"
            with zipfile.ZipFile(deck, "w") as archive:
                archive.writestr(
                    "[Content_Types].xml",
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                    '<Override PartName="/ppt/presentation.xml" '
                    f'ContentType="{MODULE.PPTX_CONTENT_TYPE}"/></Types>',
                )
                archive.writestr("ppt/presentation.xml", "<p:presentation/>")
            ghost = root / "ghost.crdownload"
            ghost.write_bytes(b"partial download")

            real_stat = Path.stat
            seen = {"count": 0}

            def racy_stat(self, **kwargs):
                if self.name == "ghost.crdownload":
                    seen["count"] += 1
                    if seen["count"] > 1:
                        raise FileNotFoundError(2, "vanished mid-scan", str(self))
                return real_stat(self, **kwargs)

            with patch.object(Path, "stat", racy_stat):
                found = MODULE.find_download([root], timeout=2.0)
            self.assertEqual(found.resolve(), deck.resolve())

    def test_editor_fonts_are_mapped_to_system_fonts(self):
        replaced = {}
        xml = (
            '<a:rPr><a:latin typeface="Liter" pitchFamily="34"/><a:ea typeface="MiSans"/>'
            '<a:cs typeface="Segoe UI"/></a:rPr><p:font typeface="MiSans"/>'
        )
        result = MODULE.map_fonts_in_xml(xml, replaced)
        self.assertIn('<a:latin typeface="Arial" pitchFamily="34"/>', result)
        self.assertIn('<a:ea typeface="Microsoft YaHei"/>', result)
        self.assertIn('<a:cs typeface="Segoe UI"/>', result)
        # Embedded-font declarations are not text runs and stay untouched.
        self.assertIn('<p:font typeface="MiSans"/>', result)
        self.assertEqual(replaced, {"Liter": "Arial", "MiSans": "Microsoft YaHei"})

    def test_every_fallback_has_preview_faces(self):
        for target in MODULE.FONT_FALLBACKS.values():
            self.assertIn(target, MODULE.FONT_FACES)
            self.assertIn(target.lower(), MODULE.STANDARD_FONTS)
        css = MODULE.font_alias_css()
        self.assertIn('@font-face{font-family:"MiSans";src:local("Microsoft YaHei")', css)
        self.assertIn('font-family:"MiSans";font-weight:700;src:local("Microsoft YaHei Bold")', css)
        script = MODULE.font_alias_script()
        self.assertIn('"misans"', script)
        self.assertNotIn("__CSS__", script)

    def test_newlines_in_runs_become_line_breaks(self):
        xml = (
            '<a:p><a:r><a:rPr lang="zh-CN" sz="1200"><a:solidFill/></a:rPr>'
            "<a:t>线索\n跟进</a:t></a:r><a:r><a:rPr/><a:t>a&#10;b</a:t></a:r>"
            "<a:r><a:t>plain</a:t></a:r></a:p>"
        )
        result, count = MODULE.split_newline_runs(xml)
        self.assertEqual(count, 2)
        self.assertIn(
            '<a:r><a:rPr lang="zh-CN" sz="1200"><a:solidFill/></a:rPr><a:t>线索</a:t></a:r>'
            '<a:br><a:rPr lang="zh-CN" sz="1200"><a:solidFill/></a:rPr></a:br>'
            '<a:r><a:rPr lang="zh-CN" sz="1200"><a:solidFill/></a:rPr><a:t>跟进</a:t></a:r>',
            result,
        )
        self.assertIn("<a:r><a:rPr/><a:t>a</a:t></a:r><a:br><a:rPr/></a:br>", result)
        self.assertIn("<a:r><a:t>plain</a:t></a:r>", result)
        self.assertNotIn("\n", result)

    def test_postprocess_rewrites_deck_for_office(self):
        with tempfile.TemporaryDirectory() as name:
            deck = Path(name) / "deck.pptx"
            slide = (
                '<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><p:txBody><a:p><a:r>'
                '<a:rPr><a:latin typeface="Liter"/><a:ea typeface="HarmonyOS Sans SC"/></a:rPr>'
                "<a:t>一\n二</a:t></a:r></a:p></p:txBody></p:cSld></p:sld>"
            )
            with zipfile.ZipFile(deck, "w") as archive:
                archive.writestr("ppt/slides/slide1.xml", slide)
                archive.writestr(
                    "ppt/theme/theme1.xml",
                    '<a:theme xmlns:a="urn:a"><a:minorFont><a:latin typeface="MiSans"/>'
                    "</a:minorFont></a:theme>",
                )
            result = MODULE.postprocess_pptx(deck, "fade")
            self.assertEqual(result["slides"], 1)
            self.assertEqual(result["lineBreaks"], 1)
            self.assertEqual(result["fontsReplaced"], {"Liter": "Arial", "MiSans": "Microsoft YaHei"})
            self.assertEqual(result["nonStandardFonts"], ["HarmonyOS Sans SC"])
            with zipfile.ZipFile(deck) as archive:
                self.assertIsNone(archive.testzip())
                slide_out = archive.read("ppt/slides/slide1.xml").decode("utf-8")
                theme_out = archive.read("ppt/theme/theme1.xml").decode("utf-8")
            self.assertIn('typeface="Arial"', slide_out)
            self.assertIn("<a:br>", slide_out)
            self.assertIn("<p:fade/>", slide_out)
            self.assertIn('typeface="Microsoft YaHei"', theme_out)

            kept = MODULE.postprocess_pptx(deck, "fade", map_fonts=False)
            self.assertEqual(kept["fontsReplaced"], {})

    def test_line_break_warning_only_for_browser_exporter(self):
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            manifest = root / "deck.pptd"
            manifest.write_text("version: v2\n", encoding="utf-8")
            (root / "pages").mkdir()
            (root / "pages" / "1.page").write_text("text: a<br/>b\n", encoding="utf-8")
            local = {"exporter": "local-wasm-patched"}
            MODULE.warn_dropped_line_breaks(manifest, local)
            self.assertNotIn("warnings", local)
            browser = {"exporter": "browser-local-editor"}
            MODULE.warn_dropped_line_breaks(manifest, browser)
            self.assertIn("pages/1.page (1)", browser["warnings"][0])

    def test_vendored_yaml_is_used_when_pyyaml_is_missing(self):
        vendor = SCRIPT.parent / "_vendor"
        self.assertTrue((vendor / "yaml" / "__init__.py").is_file())
        is_yaml = lambda key: key == "yaml" or key.startswith("yaml.")  # noqa: E731
        saved = {key: value for key, value in MODULE.sys.modules.items() if is_yaml(key)}
        # Hide every installed PyYAML: no site-packages, no cached module.
        bare_path = [entry for entry in MODULE.sys.path if "site-packages" not in entry]
        try:
            for key in saved:
                del MODULE.sys.modules[key]
            with patch.object(MODULE.sys, "path", bare_path):
                module = MODULE.load_yaml_module()
            self.assertEqual(module.safe_load("pages: [a.page]"), {"pages": ["a.page"]})
            self.assertTrue(str(Path(module.__file__).resolve()).startswith(str(vendor)))
        finally:
            for key in [key for key in MODULE.sys.modules if is_yaml(key)]:
                del MODULE.sys.modules[key]
            MODULE.sys.modules.update(saved)

    def test_piped_output_is_utf8(self):
        import io

        stream = io.TextIOWrapper(io.BytesIO(), encoding="gbk")
        with patch.object(MODULE.sys, "stdout", stream), patch.object(MODULE.sys, "stderr", stream):
            MODULE.make_output_robust()
            print("✓ 完成 🎉")
        stream.flush()
        self.assertIn("✓ 完成 🎉".encode("utf-8"), stream.buffer.getvalue())


class CdpTests(unittest.TestCase):
    def setUp(self):
        import cdp

        self.cdp = cdp

    def test_mask_is_an_involution(self):
        key = b"\x01\x02\x03\x04"
        data = "导出 PPTX".encode("utf-8") * 50
        masked = self.cdp._mask(data, key)
        self.assertNotEqual(masked, data)
        self.assertEqual(self.cdp._mask(masked, key), data)
        self.assertEqual(self.cdp._mask(b"", key), b"")

    def test_websocket_reads_fragmented_and_large_frames(self):
        import struct

        payload = ("x" * 70000).encode("utf-8")
        frames = (
            bytes([0x01, 5]) + b"hello"  # text, not final
            + bytes([0x89, 0])  # ping in between
            + bytes([0x80, 127]) + struct.pack(">Q", len(payload)) + payload  # final continuation
        )

        class FakeSocket:
            def __init__(self, data):
                self.data = bytearray(data)
                self.sent = []

            def recv(self, size):
                chunk = bytes(self.data[:1000])
                del self.data[:1000]
                return chunk

            def sendall(self, data):
                self.sent.append(data)

        ws = self.cdp.WebSocket.__new__(self.cdp.WebSocket)
        ws.sock = FakeSocket(frames)
        ws.buffer = bytearray()
        self.assertEqual(ws.recv_text(), "hello" + "x" * 70000)
        self.assertEqual(ws.sock.sent[0][0], 0x8A)  # answered the ping with a pong

    def test_explicit_browser_comes_first(self):
        with tempfile.TemporaryDirectory() as name:
            browser = Path(name) / "my-browser.exe"
            browser.write_bytes(b"")
            with patch.dict(self.cdp.os.environ, {"OKP_BROWSER": str(browser)}):
                self.assertEqual(self.cdp.browser_candidates()[0], str(browser))


if __name__ == "__main__":
    unittest.main()
