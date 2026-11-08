#!/usr/bin/env python3
import importlib.util
import os
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

SPEC = importlib.util.spec_from_file_location("export_images", SCRIPTS_DIR / "export_images.py")
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC and SPEC.loader
SPEC.loader.exec_module(MODULE)


def make_images_zip(path: Path, names=("1.jpeg", "10.jpeg", "2.jpeg")) -> None:
    # 1x1 white JPEG, the smallest valid payload Pillow can open.
    import base64

    pixel = base64.b64decode(
        "/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAP//////////////////////////"
        "////////////////////////////////////////////2wBDAf//////////////////"
        "////////////////////////////////////////////wAARCAABAAEDASIAAhEBAxEB"
        "/8QAFQABAQAAAAAAAAAAAAAAAAAAAAX/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oADAMB"
        "AAIQAxAAAAGf/8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABBQJ//8QAFBEBAAAA"
        "AAAAAAAAAAAAAAAAAP/aAAgBAwEBPwF//8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAgB"
        "AgEBPwF//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQAGPwJ//8QAFBABAAAAAAAA"
        "AAAAAAAAAAAAAP/aAAgBAQABPyF//9oADAMBAAIAAwAAABCf/8QAFBEBAAAAAAAAAAAA"
        "AAAAAAAAAP/aAAgBAwEPEBB//8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAgBAgEPEBB/"
        "/8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPxB//9k="
    )
    with zipfile.ZipFile(path, "w") as archive:
        for name in names:
            archive.writestr(name, pixel)


class ExportImagesTests(unittest.TestCase):
    def test_page_sort_key_orders_numeric_stems(self):
        paths = [Path("10.jpeg"), Path("2.jpeg"), Path("cover.jpeg"), Path("1.jpeg")]
        ordered = sorted(paths, key=MODULE.page_sort_key)
        self.assertEqual(
            [path.name for path in ordered],
            ["1.jpeg", "2.jpeg", "10.jpeg", "cover.jpeg"],
        )

    def test_is_image_zip_accepts_image_entries_only(self):
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            good = root / "images.zip"
            make_images_zip(good)
            self.assertTrue(MODULE.is_image_zip(good))

            bad = root / "text.zip"
            with zipfile.ZipFile(bad, "w") as archive:
                archive.writestr("readme.txt", "hello")
            self.assertFalse(MODULE.is_image_zip(bad))
            self.assertFalse(MODULE.is_image_zip(root / "missing.zip"))

    def test_unzip_images_flattens_and_sorts(self):
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            archive_path = root / "images.zip"
            make_images_zip(archive_path, names=("1.jpeg", "10.jpeg", "2.jpeg", "note.txt"))
            images = MODULE.unzip_images(archive_path, root / "pages")
            self.assertEqual(
                [path.name for path in images], ["1.jpeg", "2.jpeg", "10.jpeg"]
            )

    def test_overview_html_labels_pages_in_order(self):
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            images = [root / "pages" / f"{index}.jpeg" for index in (1, 2, 10)]
            document = MODULE.overview_html(images, root)
            self.assertLess(document.index("P1<"), document.index("P2<"))
            self.assertLess(document.index("P2<"), document.index("P3<"))
            self.assertIn('src="pages/10.jpeg"', document)
            self.assertIn(f"repeat({MODULE.OVERVIEW_COLUMNS},{MODULE.OVERVIEW_THUMB_WIDTH}px)", document)

    @unittest.skipUnless(os.environ.get("OKP_E2E"), "set OKP_E2E=1 to drive a real browser")
    def test_stitch_overview_with_headless_browser(self):
        import cdp

        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            pages = root / "pages"
            pages.mkdir()
            images = []
            # 16:9 SVG pages keep the test free of image libraries.
            for index in range(1, 5):
                path = pages / f"{index}.svg"
                path.write_text(
                    "<svg xmlns='http://www.w3.org/2000/svg' width='320' height='180'>"
                    f"<rect width='320' height='180' fill='#{index * 2}0{index}0a0'/></svg>",
                    encoding="utf-8",
                )
                images.append(path)
            browser = cdp.launch_browser()
            try:
                overview = MODULE.stitch_overview(browser, images, root / "overview.jpg")
            finally:
                browser.close()
            data = overview.read_bytes()
            self.assertEqual(data[:2], b"\xff\xd8")  # JPEG
            # JPEG SOF0 marker carries height/width.
            index = data.index(b"\xff\xc0")
            height = int.from_bytes(data[index + 5:index + 7], "big")
            width = int.from_bytes(data[index + 7:index + 9], "big")
            self.assertEqual(width, 3 * MODULE.OVERVIEW_THUMB_WIDTH + 4 * MODULE.OVERVIEW_GAP)
            rows = 2
            cell = MODULE.OVERVIEW_LABEL_HEIGHT + 360
            self.assertEqual(height, rows * cell + (rows + 1) * MODULE.OVERVIEW_GAP)
            self.assertFalse(any(root.glob(".*.html")))


if __name__ == "__main__":
    unittest.main()
