import json
import os
from pathlib import Path
import sys
import tempfile
import threading
import time
import unittest
from unittest import mock
from urllib.error import HTTPError
from urllib.parse import quote
from urllib.request import Request, urlopen

from docx import Document
from PIL import Image

APP_DIR = Path(__file__).resolve().parent.parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

import config_store
import doc_generator
import html_generator
from pdf_images import PLOT_TAG_PREFIX
from scanner import ScanManager, list_directory
from server import create_server


def _write_png(path: Path, color=(40, 120, 200), size=(64, 48)) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    img = Image.new("RGB", size, color)
    img.save(path, format="PNG")


class ScannerStressTests(unittest.TestCase):
    def test_large_folder_1200_images_lazy_guardrail_and_cancel(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            # Create 1,200 tiny PNGs across 6 subfolders (200 each)
            for folder_idx in range(6):
                sub = root / f"group_{folder_idx}"
                sub.mkdir()
                for img_idx in range(200):
                    _write_png(sub / f"plot_{img_idx:03d}.png", size=(16, 12))

            # 1. Instant single-level listing
            listing = list_directory(str(root))
            self.assertEqual(len(listing["folders"]), 6)
            self.assertEqual(len(listing["files"]), 0)

            sub_listing = list_directory(str(root / "group_0"))
            self.assertEqual(len(sub_listing["files"]), 200)

            # 2. Guarded deep scan pauses at soft_limit=500, then finishes all 1,200
            mgr = ScanManager()
            scan_id = mgr.start_scan(str(root), soft_limit=500)

            deadline = time.time() + 15.0
            status = None
            while time.time() < deadline:
                status = mgr.get_status(scan_id)
                if status["status"] == "paused_limit":
                    break
                time.sleep(0.05)

            self.assertIsNotNone(status)
            self.assertEqual(status["status"], "paused_limit")
            self.assertEqual(status["total_files"], 500)

            # Incremental fetch check
            inc = mgr.get_status(scan_id, since_index=450)
            self.assertEqual(len(inc["new_files"]), 50)
            self.assertEqual(inc["next_index"], 500)

            mgr.continue_scan(scan_id, new_limit=0)
            while time.time() < deadline:
                status = mgr.get_status(scan_id)
                if status["status"] == "done":
                    break
                time.sleep(0.05)

            self.assertEqual(status["status"], "done")
            self.assertEqual(status["total_files"], 1200)

            # 3. Cancel paused scan immediately unblocks
            cancel_id = mgr.start_scan(str(root), soft_limit=100)
            while time.time() < deadline:
                st2 = mgr.get_status(cancel_id)
                if st2["status"] == "paused_limit":
                    break
                time.sleep(0.05)
            cancelled = mgr.cancel_scan(cancel_id)
            self.assertEqual(cancelled["status"], "cancelled")

    def test_scan_manager_prunes_finished_jobs_beyond_max(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            _write_png(root / "p.png")
            mgr = ScanManager()
            deadline = time.time() + 10.0
            for _ in range(mgr.MAX_JOBS + 5):
                sid = mgr.start_scan(str(root), soft_limit=0)
                while time.time() < deadline:
                    if mgr.get_status(sid)["status"] == "done":
                        break
                    time.sleep(0.01)
            self.assertLessEqual(len(mgr._jobs), mgr.MAX_JOBS)


class ConfigStoreTests(unittest.TestCase):
    def test_save_load_validate_and_remap(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            src_a = root / "run_1"
            src_b = root / "run_2"
            p1 = src_a / "sub" / "a.png"
            p2 = src_a / "b.png"
            _write_png(p1)
            _write_png(p2)

            # Mirror into run_2 for remap test
            _write_png(src_b / "sub" / "a.png")
            _write_png(src_b / "b.png")

            cfg_path = root / "my_report.plot2docx.json"
            state = {
                "name": "Quarterly Plots",
                "source_folder": str(src_a),
                "grouping_mode": "manual",
                "sections": [
                    {"id": "sec_1", "name": "Summary", "page_break": True, "plots": [str(p1)]},
                    {"id": "sec_2", "name": "Details", "page_break": False, "plots": [str(p2)]},
                ],
                "doc_options": {
                    "orientation": "portrait",
                    "num_cols": 1,
                    "max_width_cm": 12.5,
                    "max_height_cm": 8.0,
                    "export_pdf": False,
                    "last_output_path": str(root / "out.docx"),
                },
            }

            saved = config_store.save_config(str(cfg_path), state)
            self.assertTrue(cfg_path.is_file())
            self.assertTrue(saved["file_status"][str(p1)]["exists"])

            # Delete p2 to verify missing detection, modify p1 to verify mtime update
            old_mtime = saved["file_status"][str(p1)]["mtime"]
            time.sleep(0.05)
            _write_png(p1, color=(200, 40, 40))
            p2.unlink()

            loaded = config_store.load_config(str(cfg_path))
            self.assertEqual(loaded["config"]["name"], "Quarterly Plots")
            self.assertEqual(loaded["config"]["doc_options"]["num_cols"], 1)
            self.assertEqual(loaded["config"]["doc_options"]["max_width_cm"], 12.5)
            self.assertTrue(loaded["file_status"][str(p1)]["exists"])
            self.assertGreaterEqual(loaded["file_status"][str(p1)]["mtime"], old_mtime)
            self.assertFalse(loaded["file_status"][str(p2)]["exists"])

            # Remap to src_b where both plots exist
            remapped = config_store.remap_config_source(loaded["config"]["sections"], str(src_b))
            self.assertEqual(remapped["matched_count"], 2)
            self.assertEqual(remapped["missing_count"], 0)
            self.assertEqual(remapped["sections"][0]["plots"][0]["path"], str(src_b / "sub" / "a.png"))

    def test_remap_disambiguates_duplicate_basenames_and_sanitizes_rel_path(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            src_a = root / "run_a"
            src_b = root / "run_b"
            _write_png(src_a / "subA" / "plot.png")
            _write_png(src_a / "subB" / "plot.png")
            _write_png(src_b / "subA" / "plot.png")
            _write_png(src_b / "subB" / "plot.png")

            # Even without rel_path, passing old_source_folder resolves subA/plot.png vs subB/plot.png
            sections = [{
                "id": "sec_1",
                "name": "S1",
                "page_break": True,
                "plots": [
                    str(src_a / "subA" / "plot.png"),
                    str(src_a / "subB" / "plot.png"),
                    {"path": str(src_a / "subA" / "plot.png"), "rel_path": "../../etc/passwd"},
                ],
            }]
            remapped = config_store.remap_config_source(sections, str(src_b), old_source_folder=str(src_a))
            self.assertEqual(remapped["matched_count"], 3)
            self.assertEqual(remapped["sections"][0]["plots"][0]["path"], str(src_b / "subA" / "plot.png"))
            self.assertEqual(remapped["sections"][0]["plots"][1]["path"], str(src_b / "subB" / "plot.png"))
            # Traversal rel_path was rejected and replaced with safe path inside src_b
            self.assertEqual(remapped["sections"][0]["plots"][2]["path"], str(src_b / "subA" / "plot.png"))

            payload = config_store.build_config_payload({
                "source_folder": str(src_a),
                "sections": [{
                    "id": "sec_1",
                    "name": "S1",
                    "plots": [{"path": str(src_a / "subA" / "plot.png"), "rel_path": "..\\outside.png"}],
                }],
                "doc_options": {"num_cols": 99, "max_width_cm": -5, "max_height_cm": "abc"},
            })
            self.assertEqual(payload["doc_options"]["num_cols"], 2)
            self.assertIsNone(payload["doc_options"]["max_width_cm"])
            self.assertIsNone(payload["doc_options"]["max_height_cm"])
            self.assertEqual(payload["sections"][0]["plots"][0]["rel_path"], os.path.join("subA", "plot.png"))

    def test_nested_subsections_save_load_remap_and_depth_cap(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            src_a = root / "run_a"
            src_b = root / "run_b"
            p1 = src_a / "a.png"
            p2 = src_a / "sub1" / "b.png"
            p3 = src_a / "sub1" / "sub2" / "c.png"
            _write_png(p1)
            _write_png(p2)
            _write_png(p3)
            _write_png(src_b / "a.png")
            _write_png(src_b / "sub1" / "b.png")
            _write_png(src_b / "sub1" / "sub2" / "c.png")

            cfg_path = root / "nested.plot2docx.json"
            state = {
                "name": "Nested",
                "source_folder": str(src_a),
                "grouping_mode": "manual",
                "sections": [{
                    "id": "sec_1",
                    "name": "Top",
                    "page_break": True,
                    "plots": [str(p1)],
                    "subsections": [{
                        "id": "sec_2",
                        "name": "Level 1",
                        "page_break": False,
                        "plots": [str(p2)],
                        "subsections": [{
                            "id": "sec_3",
                            "name": "Level 2",
                            "page_break": False,
                            "plots": [str(p3)],
                            "subsections": [{
                                "id": "sec_4",
                                "name": "Too Deep",
                                "plots": [str(p3)],
                            }],
                        }],
                    }],
                }],
                "doc_options": {
                    "export_docx": True,
                    "export_html": True,
                    "html_embed_images": False,
                },
            }
            saved = config_store.save_config(str(cfg_path), state)
            self.assertTrue(saved["config"]["doc_options"]["export_html"])
            self.assertFalse(saved["config"]["doc_options"]["html_embed_images"])
            # Depth 3 should be dropped (capped at 2 subsection levels)
            lvl2 = saved["config"]["sections"][0]["subsections"][0]["subsections"][0]
            self.assertEqual(lvl2["subsections"], [])

            loaded = config_store.load_config(str(cfg_path))
            self.assertEqual(loaded["config"]["sections"][0]["subsections"][0]["name"], "Level 1")
            remapped = config_store.remap_config_source(
                loaded["config"]["sections"], str(src_b), old_source_folder=str(src_a)
            )
            self.assertEqual(remapped["matched_count"], 3)
            self.assertEqual(remapped["missing_count"], 0)
            self.assertEqual(
                remapped["sections"][0]["subsections"][0]["subsections"][0]["plots"][0]["path"],
                str(src_b / "sub1" / "sub2" / "c.png"),
            )


class DocGeneratorTests(unittest.TestCase):
    def test_generate_1col_and_2col_docx_respects_sections_and_limits(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            img1 = root / "p1.png"
            img2 = root / "p2.png"
            img3 = root / "p3.png"
            _write_png(img1, size=(200, 100))
            _write_png(img2, size=(200, 100))
            _write_png(img3, size=(200, 100))

            sections = [
                {"title": "Empty Section", "page_break": True, "plots": []},
                {"title": "Alpha", "page_break": True, "plots": [str(img1), str(img2)]},
                {"title": "Beta", "page_break": False, "plots": [str(img3)]},
            ]

            out_1col = root / "report_1col.docx"
            res1 = doc_generator.build_report_documents(
                sections,
                str(out_1col),
                orientation="portrait",
                num_cols=1,
                user_max_w_cm=10.0,
                user_max_h_cm=4.0,
                export_pdf=False,
            )
            self.assertTrue(res1["success"])
            self.assertEqual(res1["count"], 3)
            doc1 = Document(str(out_1col))
            self.assertEqual(len(doc1.inline_shapes), 3)
            texts1 = "\n".join(p.text for p in doc1.paragraphs)
            self.assertIn("Section 1: Alpha", texts1)
            self.assertIn("Section 2: Beta", texts1)
            self.assertNotIn("Empty Section", texts1)
            # Height cap of 4.0 cm on 2:1 image => width ~8.0 cm (within EMU rounding)
            self.assertAlmostEqual(doc1.inline_shapes[0].height.cm, 4.0, delta=0.1)
            self.assertAlmostEqual(doc1.inline_shapes[0].width.cm, 8.0, delta=0.1)

            out_2col = root / "report_2col.docx"
            res2 = doc_generator.build_report_documents(
                sections,
                str(out_2col),
                orientation="landscape",
                num_cols=2,
                user_max_w_cm=8.0,
                user_max_h_cm=3.5,
                export_pdf=False,
            )
            self.assertTrue(res2["success"])
            self.assertEqual(res2["count"], 3)
            doc2 = Document(str(out_2col))
            self.assertEqual(len(doc2.inline_shapes), 3)
            # Verify 2-column mode now respects user_max_h_cm (3.5 cm => 7.0 cm width)
            self.assertAlmostEqual(doc2.inline_shapes[0].height.cm, 3.5, delta=0.1)
            self.assertAlmostEqual(doc2.inline_shapes[0].width.cm, 7.0, delta=0.1)
            for idx, shape in enumerate(doc2.inline_shapes, start=1):
                self.assertEqual(shape._inline.docPr.get("descr"), f"{PLOT_TAG_PREFIX}{idx}")

    def test_exif_transpose_missing_pack_and_invalid_inputs(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            # Create a 200x100 JPEG with EXIF Orientation=6 (90 deg CW -> visual 100x200)
            jpg_path = root / "rotated.jpg"
            img = Image.new("RGB", (200, 100), (10, 90, 180))
            exif = img.getexif()
            exif[0x0112] = 6
            img.save(jpg_path, format="JPEG", exif=exif)

            valid_png = root / "valid.png"
            _write_png(valid_png, size=(200, 100))
            missing_png = root / "missing.png"

            out_mixed = root / "mixed_2col.docx"
            res_mixed = doc_generator.build_report_documents(
                [{"title": "Mixed", "page_break": True, "plots": [str(missing_png), str(jpg_path), str(valid_png)]}],
                str(out_mixed),
                orientation="landscape",
                num_cols=2,
            )
            self.assertTrue(res_mixed["success"])
            self.assertEqual(res_mixed["count"], 2)
            self.assertEqual(len(res_mixed["skipped"]), 1)
            doc_m = Document(str(out_mixed))
            self.assertEqual(len(doc_m.inline_shapes), 2)
            # Only 1 pair of rows (img + caption) for 2 valid items in 2-col table
            self.assertEqual(len(doc_m.tables[0].rows), 2)
            # Transposed JPEG is taller than wide (height > width)
            self.assertGreater(doc_m.inline_shapes[0].height.cm, doc_m.inline_shapes[0].width.cm)

            # All missing -> no empty .docx saved
            out_empty = root / "should_not_exist.docx"
            res_empty = doc_generator.build_report_documents(
                [{"title": "Empty", "page_break": True, "plots": [str(missing_png)]}],
                str(out_empty),
                num_cols=2,
            )
            self.assertFalse(res_empty["success"])
            self.assertFalse(out_empty.exists())

            # Invalid num_cols=0
            res_bad = doc_generator.build_report_documents(
                [{"title": "Bad", "page_break": True, "plots": [str(valid_png)]}],
                str(out_empty),
                num_cols=0,
            )
            self.assertFalse(res_bad["success"])
            self.assertFalse(out_empty.exists())

    def test_nested_subsections_use_level1_figure_numbering(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            imgs = [root / f"p{i}.png" for i in range(1, 5)]
            for im in imgs:
                _write_png(im, size=(120, 80))

            sections = [
                {
                    "title": "Overview",
                    "page_break": True,
                    "plots": [str(imgs[0])],
                    "subsections": [
                        {
                            "title": "Thermal",
                            "page_break": False,
                            "plots": [str(imgs[1])],
                            "subsections": [
                                {"title": "Steady", "page_break": False, "plots": [str(imgs[2])]},
                            ],
                        },
                        {"title": "Empty Sub", "page_break": False, "plots": []},
                    ],
                },
                {
                    "title": "Next Section",
                    "page_break": True,
                    "plots": [str(imgs[3])],
                    "subsections": [],
                },
            ]
            out_docx = root / "nested_report.docx"
            res = doc_generator.build_report_documents(sections, str(out_docx), num_cols=1)
            self.assertTrue(res["success"])
            self.assertEqual(res["count"], 4)

            doc = Document(str(out_docx))
            texts = [p.text for p in doc.paragraphs]
            joined = "\n".join(texts)
            self.assertIn("Section 1: Overview", joined)
            self.assertIn("1.1 Thermal", joined)
            self.assertIn("1.1.1 Steady", joined)
            self.assertIn("Section 2: Next Section", joined)
            self.assertNotIn("Empty Sub", joined)
            # Level-1 figure numbering: 1.1, 1.2, 1.3 across subsections, resetting to 2.1 in Section 2
            self.assertIn("Figure 1.1\t", joined)
            self.assertIn("Figure 1.2\t", joined)
            self.assertIn("Figure 1.3\t", joined)
            self.assertIn("Figure 2.1\t", joined)


class HtmlGeneratorTests(unittest.TestCase):
    def test_image_src_uri_and_html_export_combinations(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            p1 = root / "sub" / "plot 1.png"
            p2 = root / "sub" / "plot 2.png"
            _write_png(p1, size=(160, 100))
            _write_png(p2, size=(160, 100))

            # Relative path when inside same folder tree
            rel_src = html_generator.image_src_for_html(str(p1), str(root), embed_images=False)
            self.assertEqual(rel_src, "sub/plot%201.png")

            # UNC path -> file://pentagon/... URI
            unc_src = html_generator.image_src_for_html(r"\\pentagon\share\sub folder\plot #1.png", str(root), embed_images=False)
            self.assertTrue(unc_src.startswith("file://pentagon/share/"))
            self.assertIn("plot%20%231.png", unc_src)

            # Embedded base64 data URI
            data_src = html_generator.image_src_for_html(str(p1), str(root), embed_images=True)
            self.assertTrue(data_src.startswith("data:image/png;base64,"))

            sections = [
                {
                    "title": "Main",
                    "plots": [str(p1)],
                    "subsections": [{"title": "Child", "plots": [str(p1)]}],
                },
                {"title": "Second", "plots": [str(p2)], "subsections": []},
            ]

            # 1. HTML-only (no .docx written)
            html_only_base = root / "html_only.html"
            res_html = doc_generator.build_report_documents(
                sections,
                str(html_only_base),
                num_cols=2,
                export_docx=False,
                export_html=True,
                export_pdf=True,  # Ignored when export_docx=False
            )
            self.assertTrue(res_html["success"])
            self.assertIsNone(res_html["docx_path"])
            self.assertIsNone(res_html["pdf_path"])
            self.assertTrue(Path(res_html["html_path"]).is_file())
            self.assertFalse(root.joinpath("html_only.docx").exists())

            html_text = Path(res_html["html_path"]).read_text(encoding="utf-8")
            self.assertIn("Section 1: Main", html_text)
            self.assertIn("1.1 Child", html_text)
            self.assertIn("Figure 1.1", html_text)
            self.assertIn("Figure 1.2", html_text)
            self.assertIn("Figure 2.1", html_text)
            self.assertIn('loading="lazy"', html_text)

            # 2. Both DOCX + HTML
            both_base = root / "both_report.docx"
            res_both = doc_generator.build_report_documents(
                sections,
                str(both_base),
                num_cols=2,
                export_docx=True,
                export_html=True,
                html_embed_images=True,
            )
            self.assertTrue(res_both["success"])
            self.assertTrue(Path(res_both["docx_path"]).is_file())
            self.assertTrue(Path(res_both["html_path"]).is_file())
            self.assertIn("data:image/png;base64,", Path(res_both["html_path"]).read_text(encoding="utf-8"))

            # 3. Neither format checked -> fails cleanly
            res_none = doc_generator.build_report_documents(
                sections, str(both_base), export_docx=False, export_html=False
            )
            self.assertFalse(res_none["success"])


class ServerEndToEndTests(unittest.TestCase):
    def _request_json(self, base_url: str, path: str, payload: dict | None = None):
        if payload is None:
            req = Request(f"{base_url}{path}", method="GET")
        else:
            req = Request(
                f"{base_url}{path}",
                data=json.dumps(payload).encode("utf-8"),
                headers={"Content-Type": "application/json"},
                method="POST",
            )
        with urlopen(req, timeout=10) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))

    def test_http_endpoints_end_to_end(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            p1 = root / "plots" / "fig1.png"
            p2 = root / "plots" / "fig2.png"
            _write_png(p1, size=(120, 80))
            _write_png(p2, size=(120, 80))

            server = create_server(host="127.0.0.1", port=0, auto_exit=False)
            port = server.server_address[1]
            base_url = f"http://127.0.0.1:{port}"
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()

            try:
                # Static assets
                for asset in ("/", "/styles.css", "/app.js"):
                    with urlopen(f"{base_url}{asset}", timeout=5) as r:
                        self.assertEqual(r.status, 200)
                        self.assertGreater(len(r.read()), 100)

                # Health
                code, health = self._request_json(base_url, "/api/health")
                self.assertEqual(code, 200)
                self.assertTrue(health["ok"])

                # Lazy directory listing
                code, listing = self._request_json(
                    base_url, f"/api/fs/list?path={quote(str(root / 'plots'))}"
                )
                self.assertEqual(code, 200)
                self.assertEqual(len(listing["files"]), 2)

                # Scan start + poll
                _, scan_start = self._request_json(base_url, "/api/scan/start", {"path": str(root), "soft_limit": 500})
                scan_id = scan_start["scan_id"]
                for _ in range(50):
                    _, st = self._request_json(base_url, f"/api/scan/status?id={scan_id}&since=0")
                    if st["status"] == "done":
                        break
                    time.sleep(0.05)
                self.assertEqual(st["status"], "done")
                self.assertEqual(st["total_files"], 2)

                # Thumbnail endpoint
                with urlopen(f"{base_url}/api/thumb?path={quote(str(p1))}&size=96", timeout=5) as tr:
                    self.assertEqual(tr.status, 200)
                    self.assertTrue(tr.headers.get("Content-Type", "").startswith("image/"))
                    self.assertGreater(len(tr.read()), 50)

                # Config save & load
                cfg_file = root / "test.plot2docx.json"
                _, saved = self._request_json(base_url, "/api/config/save", {
                    "path": str(cfg_file),
                    "state": {
                        "name": "API Test",
                        "source_folder": str(root / "plots"),
                        "grouping_mode": "manual",
                        "sections": [{"id": "sec_1", "name": "S1", "page_break": True, "plots": [str(p1), str(p2)]}],
                        "doc_options": {"orientation": "landscape", "num_cols": 2},
                    },
                })
                self.assertTrue(Path(saved["path"]).is_file())

                _, loaded = self._request_json(base_url, "/api/config/load", {"path": str(cfg_file)})
                self.assertEqual(loaded["config"]["name"], "API Test")

                # Generate start + poll (both DOCX + HTML with nested subsections)
                out_docx = root / "api_report.docx"
                _, gen = self._request_json(base_url, "/api/generate/start", {
                    "sections": [{
                        "title": "S1",
                        "page_break": True,
                        "plots": [str(p1)],
                        "subsections": [{"title": "Sub1", "page_break": False, "plots": [str(p2)]}],
                    }],
                    "save_file": str(out_docx),
                    "orientation": "landscape",
                    "num_cols": 2,
                    "export_docx": True,
                    "export_html": True,
                    "export_pdf": False,
                })
                job_id = gen["job_id"]
                for _ in range(100):
                    _, gst = self._request_json(base_url, f"/api/generate/status?id={job_id}")
                    if gst["status"] != "running":
                        break
                    time.sleep(0.05)
                self.assertEqual(gst["status"], "done")
                self.assertTrue(out_docx.is_file())
                out_html = root / "api_report.html"
                self.assertTrue(out_html.is_file())

                # Verify invalid path returns 404 JSON
                with self.assertRaises(HTTPError) as ctx:
                    self._request_json(base_url, f"/api/fs/list?path={quote(str(root / 'does_not_exist'))}")
                self.assertEqual(ctx.exception.code, 404)

                # Cross-origin POST returns 403
                bad_origin_req = Request(
                    f"{base_url}/api/ping",
                    data=b"{}",
                    headers={"Content-Type": "application/json", "Origin": "https://evil.example"},
                    method="POST",
                )
                with self.assertRaises(HTTPError) as ctx_origin:
                    urlopen(bad_origin_req, timeout=5)
                self.assertEqual(ctx_origin.exception.code, 403)

                # Null numeric JSON field returns 400 (not 500)
                with self.assertRaises(HTTPError) as ctx_null:
                    self._request_json(base_url, "/api/scan/start", {"path": str(root), "soft_limit": None})
                self.assertEqual(ctx_null.exception.code, 400)

                # /api/open-file rejects non-document files, allows .html
                script_file = root / "danger.py"
                script_file.write_text("print(1)", encoding="utf-8")
                with self.assertRaises(HTTPError) as ctx_open:
                    self._request_json(base_url, "/api/open-file", {"path": str(script_file)})
                self.assertEqual(ctx_open.exception.code, 400)

                with mock.patch("server.os.startfile") as mocked_start:
                    code_open, open_res = self._request_json(base_url, "/api/open-file", {"path": str(out_html)})
                    self.assertEqual(code_open, 200)
                    self.assertTrue(open_res["ok"])
                    mocked_start.assert_called_once()

                # /api/config/remap supports old_source_folder
                alt_dir = root / "alt_run"
                _write_png(alt_dir / "plots" / "fig1.png", size=(120, 80))
                _, remapped = self._request_json(base_url, "/api/config/remap", {
                    "old_source_folder": str(root),
                    "new_source_folder": str(alt_dir),
                    "sections": [{"id": "sec_1", "name": "S1", "page_break": True, "plots": [str(p1)]}],
                })
                self.assertEqual(remapped["matched_count"], 1)
                self.assertEqual(remapped["sections"][0]["plots"][0]["path"], str(alt_dir / "plots" / "fig1.png"))
            finally:
                server.shutdown()
                server.server_close()


if __name__ == "__main__":
    unittest.main()
