"""
app.py — Standalone content editor tool (Flask).
Compares original vs edited page content and exports a before/after table in Word.
Navigation, headers, footers and images are ignored — focus is on page copy only.
"""

import io
from urllib.parse import urljoin

from flask import Flask, request, send_file
from bs4 import BeautifulSoup, NavigableString
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.section import WD_ORIENT

app = Flask(__name__)

IGNORED_ANCESTOR_TAGS = {
    "script", "style", "noscript", "template",
    "nav", "header", "footer", "form", "button", "svg", "img",
}
IGNORED_CLASS_ID_HINTS = ("nav", "menu", "footer", "header", "cookie", "sidebar", "breadcrumb")

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
  <button class="btn-export" onclick="exportContent()">Download</button>
</div>
<div class="__hint__">Click any text in the preview to edit it. Then click Download to get a before/after Word table (navigation, headers, footers and images are ignored).</div>
<div class="__frame_wrap__">
  <iframe id="content_frame" srcdoc="{iframe_srcdoc}"></iframe>
</div>
<script>
function loadPage(){{
  var url = document.getElementById('url_input').value;
  window.location = '/?url=' + encodeURIComponent(url) + '&dynamic=false';
}}
function exportContent(){{
  var iframeEl = document.getElementById('content_frame');
  var win = iframeEl.contentWindow;
  var doc = win.document;
  var originals = win.__ORIGINALS__ || {{}};
  var data = [];
  doc.querySelectorAll('[data-edit-id]').forEach(function(el){{
    var id = el.getAttribute('data-edit-id');
    var current = el.textContent;
    var original = (id in originals) ? originals[id] : current;
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


def is_editable_text(node) -> bool:
    if not isinstance(node, NavigableString):
        return False
    if not str(node).strip():
        return False
    parent = node.parent
    while parent is not None:
        if getattr(parent, "name", None) in IGNORED_ANCESTOR_TAGS:
            return False
        if has_ignored_hint(parent):
            return False
        parent = parent.parent
    return True


def make_editable_html(url: str, dynamic: bool) -> str:
    html = fetch_html(url, dynamic)
    soup = BeautifulSoup(html, "html.parser")
    for bad in soup(["script", "nav", "header", "footer", "form"]):
        bad.decompose()
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
    
