# -*- coding: utf-8 -*-
"""Unit tests for docx2pdf.cli (stdlib unittest; no Office software needed)."""
import io
import os
import sys
import tempfile
import unittest
from unittest import mock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx2pdf import cli  # noqa: E402
from docx2pdf import __version__  # noqa: E402


def _touch(path):
    with open(path, "wb") as f:
        f.write(b"x")


class CliVersionTest(unittest.TestCase):
    def test_version_prints_prog_and_version(self):
        out = io.StringIO()
        with mock.patch("sys.stdout", out):
            with self.assertRaises(SystemExit) as cm:
                cli.main(["--version"])
        self.assertEqual(cm.exception.code, 0)
        self.assertIn("docx2pdf %s" % __version__, out.getvalue())

    def test_no_command_fails(self):
        with mock.patch("sys.stderr", new_callable=io.StringIO):
            with self.assertRaises(SystemExit):
                cli.main([])


class CliConvertTest(unittest.TestCase):
    def test_list_backends(self):
        with mock.patch("docx2pdf.cli.available_backends", return_value=["wps", "word"]), \
                mock.patch("docx2pdf.cli.discover_kwpsconvert", return_value=None), \
                mock.patch("sys.stdout", new_callable=io.StringIO) as out:
            rc = cli.main(["convert", "--list"])
        self.assertEqual(rc, 0)
        self.assertIn("wps, word", out.getvalue())

    def test_no_backend_error_exit_2(self):
        with mock.patch("docx2pdf.cli.available_backends", return_value=[]), \
                mock.patch("sys.stderr", new_callable=io.StringIO) as err:
            rc = cli.main(["convert", "x.docx"])
        self.assertEqual(rc, 2)
        self.assertIn("no supported backend", err.getvalue())

    def test_explicit_backend_not_available(self):
        with mock.patch("docx2pdf.cli.available_backends", return_value=["wps"]), \
                mock.patch("sys.stderr", new_callable=io.StringIO) as err:
            rc = cli.main(["convert", "x.docx", "--backend", "word"])
        self.assertEqual(rc, 2)
        self.assertIn("not available", err.getvalue())

    def test_single_file_output_pdf(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, "a.docx")
            _touch(src)
            out = os.path.join(tmp, "a.pdf")
            with mock.patch("docx2pdf.cli.available_backends", return_value=["word"]), \
                    mock.patch("docx2pdf.cli.convert_file") as cf, \
                    mock.patch("docx2pdf.cli.batch_convert") as bc, \
                    mock.patch("sys.stdout", new_callable=io.StringIO):
                rc = cli.main(["convert", src, "-o", out, "--backend", "word"])
            self.assertEqual(rc, 0)
            cf.assert_called_once_with(src, out, backend="word", visible=False)
            bc.assert_not_called()

    def test_existing_dir_output_routes_to_batch(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, "a.docx")
            _touch(src)
            outdir = os.path.join(tmp, "out")
            os.makedirs(outdir)
            with mock.patch("docx2pdf.cli.available_backends", return_value=["word"]), \
                    mock.patch("docx2pdf.cli.convert_file") as cf, \
                    mock.patch("docx2pdf.cli.batch_convert",
                               return_value=[(src, os.path.join(outdir, "a.pdf"))]) as bc, \
                    mock.patch("sys.stdout", new_callable=io.StringIO):
                rc = cli.main(["convert", src, "-o", outdir, "--backend", "word"])
            self.assertEqual(rc, 0)
            bc.assert_called_once()
            cf.assert_not_called()

    def test_convert_error_exit_1(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, "a.docx")
            _touch(src)
            with mock.patch("docx2pdf.cli.available_backends", return_value=["word"]), \
                    mock.patch("docx2pdf.cli.convert_file",
                               side_effect=RuntimeError("boom")), \
                    mock.patch("sys.stderr", new_callable=io.StringIO) as err:
                rc = cli.main(["convert", src, "-o", os.path.join(tmp, "a.pdf")])
            self.assertEqual(rc, 1)
            self.assertIn("boom", err.getvalue())


if __name__ == "__main__":
    unittest.main()
