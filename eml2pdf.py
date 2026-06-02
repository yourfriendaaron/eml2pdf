#!/usr/bin/env python3
"""
eml2pdf — convert Outlook .eml files into PDFs using headless Chrome.

Usage:
    eml2pdf.py <file.eml> [<file2.eml> ...]   # convert specific files
    eml2pdf.py --watch-dir <dir>              # scan a dir for *.eml and convert all

Designed to be triggered by a launchd WatchPaths agent on ~/Downloads:
every new .eml that lands gets turned into a same-named .pdf, and the
original .eml is moved into <dir>/.eml-processed/ so it isn't converted twice.
"""

import sys
import os
import re
import base64
import quopri
import shutil
import subprocess
import tempfile
import time
import html as htmllib
from email import policy
from email.parser import BytesParser
from pathlib import Path
from datetime import datetime

CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
# WeasyPrint renders HTML/CSS to PDF directly (no browser) — ~10x faster to
# start than Chrome. It's the primary renderer; Chrome is the fallback for the
# rare email whose CSS WeasyPrint can't handle.
WEASYPRINT = "/opt/homebrew/bin/weasyprint"
LOG = Path.home() / "Library" / "Logs" / "eml2pdf.log"


def log(msg: str) -> None:
    line = f"{datetime.now().isoformat(timespec='seconds')}  {msg}"
    try:
        with open(LOG, "a") as f:
            f.write(line + "\n")
    except OSError:
        pass
    print(line, file=sys.stderr)


def wait_until_stable(path: Path, tries: int = 10, interval: float = 0.3) -> bool:
    """Wait until a file's size stops changing (i.e. the drag/copy finished)."""
    last = -1
    for _ in range(tries):
        try:
            size = path.stat().st_size
        except FileNotFoundError:
            return False
        if size == last and size > 0:
            return True
        last = size
        time.sleep(interval)
    return path.exists()


def header_table(msg) -> str:
    """Render From/To/Cc/Subject/Date as a header block above the body."""
    rows = []
    for label in ("From", "To", "Cc", "Subject", "Date"):
        val = msg.get(label)
        if val:
            val = htmllib.escape(str(val).replace("\n", " ").strip())
            rows.append(
                f'<tr><td style="padding:2px 10px 2px 0;color:#555;'
                f'font-weight:600;vertical-align:top;white-space:nowrap">{label}</td>'
                f'<td style="padding:2px 0">{val}</td></tr>'
            )
    if not rows:
        return ""
    return (
        '<table style="font:13px -apple-system,Helvetica,Arial,sans-serif;'
        'border-collapse:collapse;margin:0 0 16px 0;width:100%">'
        + "".join(rows)
        + "</table><hr style='border:none;border-top:1px solid #ddd;margin:0 0 16px'>"
    )


def collect_inline_images(msg) -> dict:
    """Map Content-ID -> data: URI for inline (cid:) images."""
    cid_map = {}
    for part in msg.walk():
        if part.get_content_maintype() != "image":
            continue
        cid = part.get("Content-ID")
        payload = part.get_payload(decode=True)
        if not payload:
            continue
        ctype = part.get_content_type()
        b64 = base64.b64encode(payload).decode("ascii")
        data_uri = f"data:{ctype};base64,{b64}"
        if cid:
            cid_map[cid.strip().strip("<>")] = data_uri
        # also key by filename in case the body references it that way
        fn = part.get_filename()
        if fn:
            cid_map[fn] = data_uri
    return cid_map


def get_body_html(msg) -> str:
    """Return best-available HTML body, falling back to wrapped plain text."""
    html_part = None
    text_part = None
    for part in msg.walk():
        ctype = part.get_content_type()
        disp = (part.get("Content-Disposition") or "").lower()
        if "attachment" in disp:
            continue
        if ctype == "text/html" and html_part is None:
            html_part = part
        elif ctype == "text/plain" and text_part is None:
            text_part = part

    if html_part is not None:
        payload = html_part.get_payload(decode=True) or b""
        charset = html_part.get_content_charset() or "utf-8"
        return payload.decode(charset, errors="replace")

    if text_part is not None:
        payload = text_part.get_payload(decode=True) or b""
        charset = text_part.get_content_charset() or "utf-8"
        text = payload.decode(charset, errors="replace")
        return "<pre style='font:13px ui-monospace,monospace;white-space:pre-wrap'>" \
            + htmllib.escape(text) + "</pre>"

    return "<p><i>(no readable body)</i></p>"


def embed_cids(html: str, cid_map: dict) -> str:
    """Replace src="cid:xxx" references with embedded data URIs."""
    def repl(m):
        quote, cid = m.group(1), m.group(2)
        key = cid.strip().strip("<>")
        uri = cid_map.get(key) or cid_map.get(key.split("@")[0])
        return f"src={quote}{uri}{quote}" if uri else m.group(0)

    return re.sub(r'src=(["\'])cid:([^"\']+)\1', repl, html, flags=re.IGNORECASE)


def ensure_charset(html: str) -> str:
    """Guarantee a UTF-8 declaration so renderers don't guess the encoding."""
    if re.search(r'<meta[^>]+charset', html, re.IGNORECASE):
        return html
    meta = '<meta charset="utf-8">'
    if re.search(r"<head[^>]*>", html, re.IGNORECASE):
        return re.sub(r"(<head[^>]*>)", r"\1" + meta, html, count=1, flags=re.IGNORECASE)
    if re.search(r"<html[^>]*>", html, re.IGNORECASE):
        return re.sub(r"(<html[^>]*>)", r"\1<head>" + meta + "</head>", html,
                      count=1, flags=re.IGNORECASE)
    return meta + html


def build_html(eml_path: Path) -> str:
    with open(eml_path, "rb") as f:
        msg = BytesParser(policy=policy.default).parse(f)

    body = get_body_html(msg)
    body = embed_cids(body, collect_inline_images(msg))

    # If the body is a full HTML doc, inject the header after <body>; else wrap it.
    head = header_table(msg)
    if re.search(r"<body[^>]*>", body, re.IGNORECASE):
        body = re.sub(r"(<body[^>]*>)", r"\1" + head, body, count=1, flags=re.IGNORECASE)
        return ensure_charset(body)
    return (
        "<!DOCTYPE html><html><head><meta charset='utf-8'>"
        "<style>body{margin:24px;font:14px -apple-system,Helvetica,Arial,sans-serif}"
        "img{max-width:100%}</style></head><body>"
        + head + body + "</body></html>"
    )


def render_weasyprint(tmp_html: str, pdf_path: Path) -> bool:
    """Fast path: render with WeasyPrint. Returns True on a non-empty PDF."""
    if not os.path.exists(WEASYPRINT):
        return False
    try:
        result = subprocess.run(
            [WEASYPRINT, "-e", "utf-8", tmp_html, str(pdf_path)],
            capture_output=True, timeout=30,
        )
    except subprocess.TimeoutExpired:
        return False
    return result.returncode == 0 and pdf_path.exists() and pdf_path.stat().st_size > 0


def render_chrome(tmp_html: str, pdf_path: Path) -> bool:
    """Fallback: render with headless Chrome (full browser engine)."""
    if not os.path.exists(CHROME):
        return False
    cmd = [
        CHROME,
        "--headless=new",
        "--disable-gpu",
        "--no-pdf-header-footer",
        "--no-sandbox",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-crash-reporter",
        "--disable-extensions",
        "--virtual-time-budget=4000",  # let remote assets/CSS settle, then stop
        f"--print-to-pdf={pdf_path}",
        f"--user-data-dir={tempfile.mkdtemp(prefix='eml2pdf-chrome-')}",
        f"file://{tmp_html}",
    ]
    # Chrome headless often lingers after writing the PDF. Rather than waiting
    # for a timeout, poll for the finished file (size stable) and then kill it.
    proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        deadline = time.time() + 45
        last_size = -1
        while time.time() < deadline:
            if proc.poll() is not None:
                break  # Chrome exited on its own
            if pdf_path.exists():
                size = pdf_path.stat().st_size
                if size > 0 and size == last_size:
                    break  # PDF written and stable
                last_size = size
            time.sleep(0.25)
    finally:
        if proc.poll() is None:
            proc.terminate()
            try:
                proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                proc.kill()
    return pdf_path.exists() and pdf_path.stat().st_size > 0


def convert(eml_path: Path) -> bool:
    eml_path = eml_path.resolve()
    if not wait_until_stable(eml_path):
        log(f"skip (vanished/empty): {eml_path.name}")
        return False

    pdf_path = eml_path.with_suffix(".pdf")
    if pdf_path.exists():
        # avoid clobbering; add a counter
        i = 2
        while eml_path.with_name(f"{eml_path.stem} ({i}).pdf").exists():
            i += 1
        pdf_path = eml_path.with_name(f"{eml_path.stem} ({i}).pdf")

    try:
        html = build_html(eml_path)
    except Exception as e:
        log(f"parse failed {eml_path.name}: {e}")
        return False

    with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False, encoding="utf-8") as tmp:
        tmp.write(html)
        tmp_html = tmp.name

    # Render to a temp PDF in /tmp (unprotected). The renderer subprocesses
    # (weasyprint/chrome) don't have TCC access to protected folders like
    # Downloads, so THIS process does the final write into the watched folder.
    tmp_pdf = Path(tempfile.mkdtemp(prefix="eml2pdf-out-")) / "out.pdf"
    try:
        renderer = "weasyprint"
        if not render_weasyprint(tmp_html, tmp_pdf):
            renderer = "chrome"
            render_chrome(tmp_html, tmp_pdf)
    finally:
        try:
            os.unlink(tmp_html)
        except OSError:
            pass

    if not tmp_pdf.exists() or tmp_pdf.stat().st_size == 0:
        log(f"no PDF produced for {eml_path.name} (tried weasyprint+chrome)")
        return False

    # Place the finished PDF into the watched folder (done by THIS process).
    try:
        shutil.move(str(tmp_pdf), str(pdf_path))
    except OSError as e:
        log(f"could not write PDF for {eml_path.name}: {e}")
        return False

    # archive the original .eml so it isn't reconverted
    archive = eml_path.parent / ".eml-processed"
    archive.mkdir(exist_ok=True)
    try:
        eml_path.rename(archive / eml_path.name)
    except OSError:
        pass

    log(f"OK [{renderer}]  {eml_path.name}  ->  {pdf_path.name}")
    return True


def main(argv) -> int:
    if not os.path.exists(WEASYPRINT) and not os.path.exists(CHROME):
        log("Neither WeasyPrint nor Chrome found; cannot render.")
        return 1

    if len(argv) >= 2 and argv[1] == "--watch-dir":
        for d in argv[2:]:
            dd = Path(d).expanduser()
            for eml in sorted(dd.glob("*.eml")):
                convert(eml)
        return 0

    files = argv[1:]
    if not files:
        print(__doc__)
        return 2
    for f in files:
        convert(Path(f).expanduser())
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
