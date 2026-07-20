"""
app.py — Standalone content editor tool (Flask).
Compares original vs edited page content and exports a before/after table in Word,
plus a clean HTML copy for local review. Navigation/footer/forms stay in the page
(to preserve layout) but their text is never made editable. Headings (H1, H2...)
are always treated as editable content, even inside a <header> wrapper.
Images keep their remote URLs (nothing is embedded/downloaded).
"""

import io
from urllib.parse import urljoin

from flask import Flask, request, send_file
from bs4 import BeautifulSoup, NavigableString
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.section import WD_ORIENT

app = Flask(__name__)

HARD_REMOVE_TAGS = {"script"}
SOFT_IGNORE_TAGS = {"footer", "nav", "form", "button"}
IGNORED_CLASS_ID_HINTS = ("nav", "menu", "footer", "cookie", "sidebar", "breadcrumb", "site-header", "topbar")

PAGE_SHELL = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Content Editor</title>
<link rel="icon" type="image/png" href="https://www.julienrio.com/images/logo.png">
<style>
body{{font-family:'Segoe UI',Arial,sans-serif;margin:0;background:#f4f5f7;}}
.__app_bar__{{background:#1a1a2e;color:#fff;padding:14px 20px;display:flex;gap:10px;align-items:center;
box-shadow:0 2px 8px rgba(0,0,0,.15);}}
.__app_bar__ img{{height:28px;width:28px;border-radius:4px;}}
.__app_bar__ input[type=text]{{flex:1;padding:9px 12px;border-radius:6px;border:none;font-size:14px;}}
.__app_bar__ button{{border:none;padding:9px 16px;border-radius:6px;cursor:pointer;font-size:13px;font-weight:600;
transition:opacity .15s;color:#fff;}}
.__app_bar__ button:hover{{opacity:.85;}}
.btn-load{{background:#e6007e;}}
.btn-export{{background:#0072ce;}}
.__frame_wrap__{{padding:16px;}}
iframe{{width:100%;height:80vh;border:1px solid #ddd;border-radius:8px;background:#fff;}}
.__hint__{{color:#555;font-size:13px;padding:0 20px 10px;}}
</style>
</head>
<body>
<div class="__app_bar__">
  <img src="https://www.julienrio.com/images/logo.png" alt="logo">
  <strong>Content Editor</strong>
  <input type="text" id="url_input" placeholder="https://your-site.com/your-page" value="{url_value}">
  <input type="hidden" id="dynamic_flag" value="{dynamic_checked}">
  <button class="btn-load" onclick="loadPage()">Load</button>
  <button class="btn-export" onclick="exportContent()">Download Word</button>
  <button class="btn-export" onclick="exportHtml()">Download HTML</button>
</div>
<div class="__hint__">Click any text (including titles) in the preview to edit it. Navigation, footers and forms are ignored. Download Word for a before/after summary, or HTML for a full local review copy.</div>
<div class="__frame_wrap__">
  <iframe id="content_frame" srcdoc="{iframe_srcdoc}"></iframe>
</div>
<script>
function loadPage(){{
  var url = document.getElementById('url_input').value;
  window.location = '/?url=' + encodeURIComponent(url) + '&dynamic=false';
}}
function getFrameParts(){{
  var iframeEl = document.getElementById('content_frame');
  var win = iframeEl.contentWindow;
  var doc = win.document;
  var originals = win.__ORIGINALS__ || {{}};
  return {{win: win, doc: doc, originals: originals}};
}}
function exportContent(){{
  var parts = getFrameParts();
  var data = [];
  parts.doc.querySelectorAll('[data-edit-id]').forEach(function(el){{
    var id = el.getAttribute('data-edit-id');
    var current = el.textContent;
    var original = (id in parts.originals) ? parts.originals[id] : current;
    data.push({{original: original, current: current}});
  }});
  fetch('/export/docx', {{
    method: 'POST',
    headers: {{'Content-Type': 'application/json'}},
    body: JSON.stringify({{blocks: data}})
  }}).then(function(resp){{ return resp.blob(); }})
    .then(function(blob){{
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'content_update.docx';
      a.click();
    }});
}}
function exportHtml(){{
  var parts = getFrameParts();
  var clone = parts.doc.documentElement.cloneNode(true);
  clone.querySelectorAll('[data-edit-id]').forEach(function(el){{
    el.removeAttribute('data-edit-id');
    el.removeAttribute('contenteditable');
    el.removeAttribute('style');
  }});
  clone.querySelectorAll('script').forEach(function(s){{ s.remove(); }});
  var htmlStr = '<!doctype html>\\n' + clone.outerHTML;
  var blob = new Blob([htmlStr], {{type: 'text/html'}});
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'edited_page.html';
  a.click();
}}
</script>
</body>
</html>"""


def fetch_html(url: str, dynamic: bool) -> str:
    if dynamic:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page()
            page.goto(url, wait_until="networkidle", timeout=30000)
            html = page.content()
            browser.close()
        return html
    import requests
    resp = requests.get(url, timeout=20, headers={"User-Agent": "EditorTool/1.0"})
    resp.raise_for_status()
    return resp.text


def absolutize(soup: BeautifulSoup, base_url: str) -> None:
    for tag, attr in [("img", "src"), ("a", "href"), ("link", "href"), ("source", "src")]:
        for el in soup.find_all(tag):
            if el.get(attr):
                el[attr] = urljoin(base_url, el[attr])


def has_ignored_hint(el) -> bool:
    identifier = " ".join([el.get("id", "")] + el.get("class", [])).lower()
    return any(h in identifier for h in IGNORED_CLASS_ID_HINTS)


def is_inside_heading(node) -> bool:
    parent = node.parent
    while parent is not None:
        if getattr(parent, "name", None) in ("h1", "h2", "h3", "h4", "h5", "h6"):
            return True
        parent = parent.parent
    return False


def is_editable_text(node) -> bool:
    if not isinstance(node, NavigableString):
        return False
    if not str(node).strip():
        return False
    if node.parent and getattr(node.parent, "name", None) == "img":
        return False

    inside_heading = is_inside_heading(node)

    parent = node.parent
    while parent is not None:
        tag = getattr(parent, "name", None)
        if tag in HARD_REMOVE_TAGS or tag in ("img", "svg"):
            return False
        if not inside_heading and tag in SOFT_IGNORE_TAGS:
            return False
        if not inside_heading and has_ignored_hint(parent):
            return False
        parent = parent.parent
    return True


def make_editable_html(url: str, dynamic: bool) -> str:
    html = fetch_html(url, dynamic)
    soup = BeautifulSoup(html, "html.parser")
    for bad in soup(list(HARD_REMOVE_TAGS)):
        bad.decompose()
    # Keep nav/form/button in the DOM for layout, but they'll never be made editable
    # (handled by is_editable_text below).
    absolutize(soup, url)

    counter = 0
    for text_node in list(soup.find_all(string=True)):
        if is_editable_text(text_node):
            span = soup.new_tag("span")
            span["data-edit-id"] = f"t{counter}"
            span["contenteditable"] = "true"
            span["style"] = "outline:1px dashed rgba(230,0,126,.4);outline-offset:1px;"
            span.string = str(text_node)
            text_node.replace_with(span)
            counter += 1

    init_script = soup.new_tag("script")
    init_script.string = """
    window.__ORIGINALS__ = {};
    document.querySelectorAll('[data-edit-id]').forEach(function(el){
      var id = el.getAttribute('data-edit-id');
      window.__ORIGINALS__[id] = el.textContent;
    });
    """
    if soup.body:
        soup.body.append(init_script)
    else:
        soup.append(init_script)

    return str(soup)


@app.route("/")
def index():
    url = request.args.get("url", "")
    dynamic = request.args.get("dynamic", "false") == "true"
    iframe_srcdoc = ""
    if url:
        try:
            iframe_srcdoc = make_editable_html(url, dynamic).replace('"', "&quot;")
        except Exception as e:
            iframe_srcdoc = f"<p style='padding:20px;color:red'>Error: {e}</p>".replace('"', "&quot;")
    return PAGE_SHELL.format(
        url_value=url,
        dynamic_checked="true" if dynamic else "false",
        iframe_srcdoc=iframe_srcdoc,
    )


def set_cell_text(cell, text, bold=False, color=None):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(text or "")
    run.font.size = Pt(10)
    run.bold = bold
    if color:
        run.font.color.rgb = color


@app.route("/export/docx", methods=["POST"])
def export_docx():
    payload = request.get_json()
    blocks = payload.get("blocks", [])

    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    new_width, new_height = section.page_height, section.page_width
    section.page_width = new_width
    section.page_height = new_height

    doc.add_heading("Content Update", level=1)
    doc.add_paragraph(
        "Left column: current content on the site. Right column: new content to paste into the CMS. "
        "Only rows with actual changes are shown."
    )

    table = doc.add_table(rows=1, cols=2)
    table.style = "Light Grid Accent 1"
    hdr = table.rows[0].cells
    set_cell_text(hdr[0], "Current content", bold=True)
    set_cell_text(hdr[1], "New content", bold=True)

    changed_color = RGBColor(0x00, 0x72, 0xCE)
    changes_found = 0

    for b in blocks:
        original = (b.get("original") or "").strip()
        current = (b.get("current") or "").strip()
        if not original and not current:
            continue
        if original == current:
            continue
        changes_found += 1
        row = table.add_row().cells
        set_cell_text(row[0], original)
        set_cell_text(row[1], current, bold=True, color=changed_color)

    if changes_found == 0:
        row = table.add_row().cells
        set_cell_text(row[0], "No changes detected.")
        set_cell_text(row[1], "")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name="content_update.docx",
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
