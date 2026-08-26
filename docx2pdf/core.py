# -*- coding: utf-8 -*-
"""Core conversion logic: WPS / MS Word COM backends, batch & recursive."""

import os
import subprocess

# Common WPS install layouts where kwpsconvert.exe may live
_KWPS_ROOTS = [
    "%LOCALAPPDATA%\\Kingsoft\\WPS Office",
    "%PROGRAMFILES%\\Kingsoft",
    "%PROGRAMFILES(X86)%\\Kingsoft",
]


def _expand(path):
    return os.path.expandvars(os.path.expanduser(path))


def discover_kwpsconvert():
    """Return the path to kwpsconvert.exe if found on this machine."""
    env = os.environ.get("KWPS_CONVERT", "").strip()
    if env:
        for cand in env.split(os.pathsep):
            if os.path.isfile(_expand(cand)):
                return _expand(cand)
    for root in _KWPS_ROOTS:
        root = _expand(root)
        if not os.path.isdir(root):
            continue
        for dirpath, _dirs, files in os.walk(root):
            if "kwpsconvert.exe" in files:
                return os.path.join(dirpath, "kwpsconvert.exe")
    return None


def available_wps_com():
    try:
        import win32com.client as win32

        win32.Dispatch("KWPS.Application")
        return True
    except Exception:  # noqa: BLE001
        return False


def available_word_com():
    try:
        import win32com.client as win32

        win32.Dispatch("Word.Application")
        return True
    except Exception:  # noqa: BLE001
        return False


def available_backends():
    backends = []
    if available_wps_com():
        backends.append("wps")
    if available_word_com():
        backends.append("word")
    if discover_kwpsconvert():
        backends.append("wps_cli")
    return backends


def convert_file(src, out, backend="auto", visible=False):
    """Convert one .docx to PDF.

    Parameters
    ----------
    src      : input .docx path
    out      : output .pdf path
    backend  : 'auto' | 'wps' | 'word' | 'wps_cli'
    visible  : show the office application window (COM backends)
    """
    if backend == "auto":
        order = ["wps", "word", "wps_cli"]
        tried = []
        for b in order:
            if b == "wps" and available_wps_com():
                return _convert_com("KWPS.Application", src, out, visible)
            if b == "word" and available_word_com():
                return _convert_com("Word.Application", src, out, visible)
            if b == "wps_cli" and discover_kwpsconvert():
                return _convert_cli(src, out)
            tried.append(b)
        raise RuntimeError(
            "No PDF backend available. Tried: %s. Install WPS Office or "
            "Microsoft Word (and pywin32 via `pip install pywin32`)." % ", ".join(tried)
        )
    if backend == "wps":
        return _convert_com("KWPS.Application", src, out, visible)
    if backend == "word":
        return _convert_com("Word.Application", src, out, visible)
    if backend == "wps_cli":
        return _convert_cli(src, out)
    raise ValueError("unknown backend: %s" % backend)


def _convert_com(progid, src, out, visible):
    import win32com.client as win32

    app = win32.Dispatch(progid)
    app.Visible = visible
    try:
        doc = app.Documents.Open(os.path.abspath(src), ReadOnly=False)
        try:
            doc.ExportAsFixedFormat(os.path.abspath(out), 17)  # 17 = PDF
        finally:
            doc.Close(False)
    finally:
        try:
            app.Quit()
        except Exception:  # noqa: BLE001
            pass
    return out


def _convert_cli(src, out):
    exe = discover_kwpsconvert()
    if not exe:
        raise RuntimeError("kwpsconvert.exe not found (set KWPS_CONVERT or install WPS)")
    os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)
    subprocess.run(
        [exe, "word2pdf", os.path.abspath(src), "-o", os.path.abspath(out)],
        check=True,
    )
    return out


def collect_files(inputs, recursive=False):
    """Expand inputs (files or dirs) into a list of (src, desired_pdf_name)."""
    found = []
    for inp in inputs:
        if os.path.isfile(inp):
            _extract_docx(inp, found)
        elif os.path.isdir(inp):
            if recursive:
                for root, _dirs, files in os.walk(inp):
                    for fn in sorted(files):
                        if fn.lower().endswith(".docx") or fn.lower().endswith(".doc"):
                            _extract_docx(os.path.join(root, fn), found, rel=True, root=inp)
            else:
                for fn in sorted(os.listdir(inp)):
                    if fn.lower().endswith(".docx") or fn.lower().endswith(".doc"):
                        _extract_docx(os.path.join(inp, fn), found)
        else:
            raise FileNotFoundError(inp)
    return found


def _extract_docx(path, found, rel=False, root=None):
    name = os.path.basename(path)
    if rel and root:
        relpath = os.path.relpath(path, root)
        name = relpath.replace(os.sep, "__")
    found.append((os.path.abspath(path), name))


def batch_convert(inputs, outdir=None, recursive=False, backend="auto", visible=False,
                  overwrite=False):
    """Convert many files.

    Returns list of (src, out) actually written.
    """
    items = collect_files(inputs, recursive=recursive)
    results = []
    for src, name in items:
        base = os.path.splitext(name)[0]
        if outdir:
            os.makedirs(outdir, exist_ok=True)
            out = os.path.abspath(os.path.join(outdir, base + ".pdf"))
        else:
            out = os.path.join(os.path.dirname(src), base + ".pdf")
        if os.path.exists(out) and not overwrite:
            print("skip (exists): %s" % out)
            continue
        convert_file(src, out, backend=backend, visible=visible)
        results.append((src, out))
    return results