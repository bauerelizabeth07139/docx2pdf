# -*- coding: utf-8 -*-
"""Unit tests for docx2pdf.core (stdlib unittest; no Office software needed)."""
import os
import sys
import tempfile
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx2pdf import core  # noqa: E402


def _touch(path):
    with open(path, "wb") as f:
        f.write(b"x")


class CollectFilesTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = self.tmp.name
        os.makedirs(os.path.join(self.root, "sub"))
        _touch(os.path.join(self.root, "a.docx"))
        _touch(os.path.join(self.root, "b.DOC"))
        _touch(os.path.join(self.root, "notes.txt"))
        _touch(os.path.join(self.root, "sub", "c.docx"))

    def test_single_file(self):
        src = os.path.join(self.root, "a.docx")
        self.assertEqual(core.collect_files([src]), [(os.path.abspath(src), "a.docx")])

    def test_dir_non_recursive(self):
        items = core.collect_files([self.root])
        names = sorted(name for _, name in items)
        self.assertEqual(names, ["a.docx", "b.DOC"])

    def test_dir_recursive(self):
        items = core.collect_files([self.root], recursive=True)
        names = sorted(name for _, name in items)
        self.assertEqual(names, ["a.docx", "b.DOC", "sub__c.docx"])

    def test_mixed_inputs_dedupe_not_required(self):
        items = core.collect_files([os.path.join(self.root, "a.docx"), self.root])
        self.assertGreaterEqual(len(items), 2)

    def test_missing_input_raises(self):
        with self.assertRaises(FileNotFoundError):
            core.collect_files([os.path.join(self.root, "nope.docx")])


class DiscoverKwpsConvertTest(unittest.TestCase):
    def test_env_override(self):
        with mock.patch.dict(os.environ, {"KWPS_CONVERT": r"C:\tools\kwpsconvert.exe"}, clear=False):
            with mock.patch("docx2pdf.core.os.path.isfile", return_value=True):
                self.assertEqual(core.discover_kwpsconvert(), r"C:\tools\kwpsconvert.exe")

    def test_env_override_missing_file(self):
        with mock.patch.dict(os.environ, {"KWPS_CONVERT": "C:/nope/kwpsconvert.exe"}, clear=False):
            with mock.patch("docx2pdf.core.os.path.isfile", return_value=False), \
                    mock.patch("docx2pdf.core.os.path.isdir", return_value=False), \
                    mock.patch("docx2pdf.core.os.walk", return_value=iter([])):
                self.assertIsNone(core.discover_kwpsconvert())


class ProbeComTest(unittest.TestCase):
    def _fake_win32(self):
        # core.py does `import win32com.client as win32`, which compiles to
        # IMPORT_NAME (returns top-level "win32com") + IMPORT_FROM "client"
        # (getattr on the top-level module), so the attribute must point at
        # the client mock too.
        win32 = mock.MagicMock()
        client = mock.MagicMock()
        win32.client = client
        client.Dispatch = win32.Dispatch
        return win32, client

    def test_probe_detects_and_quits_app(self):
        win32, client = self._fake_win32()
        app = win32.Dispatch.return_value
        with mock.patch.dict(sys.modules, {"win32com": win32, "win32com.client": client}):
            self.assertTrue(core.available_wps_com())
            win32.Dispatch.assert_called_once_with("KWPS.Application")
            app.Quit.assert_called_once()

    def test_probe_failure_returns_false(self):
        win32, client = self._fake_win32()
        win32.Dispatch.side_effect = RuntimeError("not registered")
        with mock.patch.dict(sys.modules, {"win32com": win32, "win32com.client": client}):
            self.assertFalse(core.available_word_com())

    def test_available_backends_order(self):
        with mock.patch("docx2pdf.core.available_wps_com", return_value=True), \
                mock.patch("docx2pdf.core.available_word_com", return_value=False), \
                mock.patch("docx2pdf.core.discover_kwpsconvert", return_value="C:/kw.exe"):
            self.assertEqual(core.available_backends(), ["wps", "wps_cli"])


class BatchConvertTest(unittest.TestCase):
    def test_skip_existing_without_overwrite(self):
        with tempfile.TemporaryDirectory() as tmp:
            _touch(os.path.join(tmp, "a.docx"))
            _touch(os.path.join(tmp, "a.pdf"))
            with mock.patch("docx2pdf.core.convert_file") as cf:
                results = core.batch_convert([os.path.join(tmp, "a.docx")], outdir=tmp)
            cf.assert_not_called()
            self.assertEqual(results, [])

    def test_convert_creates_outdir(self):
        with tempfile.TemporaryDirectory() as tmp:
            _touch(os.path.join(tmp, "a.docx"))
            outdir = os.path.join(tmp, "out")
            with mock.patch("docx2pdf.core.convert_file") as cf:
                results = core.batch_convert([os.path.join(tmp, "a.docx")], outdir=outdir)
            cf.assert_called_once()
            self.assertTrue(os.path.isdir(outdir))
            self.assertEqual(results, [(os.path.abspath(os.path.join(tmp, "a.docx")),
                                        os.path.join(os.path.abspath(outdir), "a.pdf"))])

    def test_overwrite_flag_converts_anyway(self):
        with tempfile.TemporaryDirectory() as tmp:
            _touch(os.path.join(tmp, "a.docx"))
            _touch(os.path.join(tmp, "a.pdf"))
            with mock.patch("docx2pdf.core.convert_file") as cf:
                results = core.batch_convert([os.path.join(tmp, "a.docx")], outdir=tmp,
                                             overwrite=True)
            self.assertEqual(len(results), 1)


class ConvertFileTest(unittest.TestCase):
    def test_unknown_backend_raises(self):
        with self.assertRaises(ValueError):
            core.convert_file("x.docx", "x.pdf", backend="bogus")

    def test_auto_with_no_backend_raises(self):
        with mock.patch("docx2pdf.core.available_wps_com", return_value=False), \
                mock.patch("docx2pdf.core.available_word_com", return_value=False), \
                mock.patch("docx2pdf.core.discover_kwpsconvert", return_value=None):
            with self.assertRaises(RuntimeError):
                core.convert_file("x.docx", "x.pdf", backend="auto")

    def test_auto_prefers_wps(self):
        with mock.patch("docx2pdf.core.available_wps_com", return_value=True), \
                mock.patch("docx2pdf.core._convert_com") as cc:
            core.convert_file("x.docx", "x.pdf", backend="auto")
            cc.assert_called_once_with("KWPS.Application", "x.docx", "x.pdf", False)


if __name__ == "__main__":
    unittest.main()
