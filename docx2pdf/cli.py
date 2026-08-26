# -*- coding: utf-8 -*-
"""Command-line interface for docx2pdf."""
import argparse
import os
import sys

from . import __version__
from .core import (
    available_backends,
    batch_convert,
    convert_file,
    discover_kwpsconvert,
)


def _cmd_convert(args):
    backends = available_backends()
    if not backends:
        print("error: no supported backend found.", file=sys.stderr)
        print(
            "Install Microsoft Word or WPS Office, then `pip install pywin32`.",
            file=sys.stderr,
        )
        return 2
    if args.list:
        print("available backends:", ", ".join(backends))
        kw = discover_kwpsconvert()
        if kw:
            print("kwpsconvert:", kw)
        return 0
    if not args.inputs:
        print("error: no input files given", file=sys.stderr)
        return 2
    backend = args.backend or "auto"
    if backend != "auto" and backend not in backends:
        print("error: backend '%s' not available (have: %s)" % (backend, ", ".join(backends)),
              file=sys.stderr)
        return 2
    try:
        if len(args.inputs) == 1 and os.path.isfile(args.inputs[0]) and args.output:
            # single-file mode, explicit output name
            convert_file(args.inputs[0], args.output, backend=backend, visible=args.visible)
            print("written:", args.output)
        else:
            results = batch_convert(
                args.inputs,
                outdir=args.output,
                recursive=args.recursive,
                backend=backend,
                visible=args.visible,
                overwrite=args.overwrite,
            )
            for src, out in results:
                print("written:", out)
        return 0
    except Exception as exc:  # noqa: BLE001
        print("error: %s" % exc, file=sys.stderr)
        return 1


def main(argv=None):
    p = argparse.ArgumentParser(
        prog="docx2pdf",
        description="Convert .docx to PDF using WPS Office / Microsoft Word.",
    )
    p.add_argument("--version", action="version", version=__version__)
    sub = p.add_subparsers(dest="command", required=True)

    cv = sub.add_parser("convert", help="convert doc/docx to pdf")
    cv.add_argument("inputs", nargs="*", help="input .doc/.docx files or directories")
    cv.add_argument("-o", "--output", help="output .pdf (single file) or output dir (multi)")
    cv.add_argument("--backend", choices=["auto", "wps", "word", "wps_cli"], default="auto")
    cv.add_argument("--recursive", action="store_true", help="recurse into subdirectories")
    cv.add_argument("--overwrite", action="store_true", help="overwrite existing PDFs")
    cv.add_argument("--visible", action="store_true", help="show office window (COM)")
    cv.add_argument("--list", action="store_true", help="list detected backends and exit")
    cv.set_defaults(func=_cmd_convert)

    args = p.parse_args(argv)
    try:
        return args.func(args)
    except Exception as exc:  # noqa: BLE001
        print("error: %s" % exc, file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())