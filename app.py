"""
app.py — Outil indépendant : application web autonome (Flask) qui fait tout en un seul endroit.
"""

import io
import json
from urllib.parse import urljoin

from flask import Flask, request, send_file, jsonify
from bs4 import BeautifulSoup, NavigableString
from docx import Document

app = Flask(__name__)

IGNORED_ANCESTOR_TAGS = {"script", "style", "noscript", "template"}

PAGE_SHELL = """<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8">
<title>Editeur de contenu web</title>
<style>
body{{font-family:sans-serif;margin:0;background:#f4f5f7;}}
.__app_bar__{{background:#111;color:#fff;padding:14px 20px;display:flex;gap:10px;align-items:center;}}
.__app_bar__ input[type=text]{{flex:1;padding:8px;border-radius:6px;border:none;}}
.__app_bar__ button{{background:#4f46e5;color:#fff;border:none;padding:8px 14px;border-radius:6px;cursor:pointer;}}
.__app_bar__ button:hover{{background:#4338ca;}}
.__frame_wrap__{{padding:16px;}}
iframe{{width:100%;height:80vh;border:1px solid #ddd;border-radius:8px;background:#fff;}}
.__hint__{{color:#555;font-size:13px;padding:0 20px 10px;}}
</style>
</head>
<body>
<div class="__app_bar__">
  <strong>🖊️ Editeur</strong>
  <input type="text" id="url_input" placeholder="https://mon-site.com/ma-page" value="{url_value}">
  <label style="color:#fff;font-size:13px;"><input type="checkbox" id="dynamic_flag" {dynamic_checked}> Site JS (React/Vue)</label>
  <button onclick="loadPage()">Charger</button>
  <button onclick="exportContent('md')">Export Markdown</button>
  <button onclick="exportContent('docx')">Export Word</button>
  <button onclick="exportContent('json')">Export JSON</button>
</div>
<div class="__hint__">Cliquez un texte dans l'aperçu pour le modifier. Cliquez une image pour changer son URL. Puis exportez.</div>
<div class="__frame_wrap__">
  <iframe id="content_frame" srcdoc="{iframe_srcdoc}"></iframe>
</div>
<script>
function loadPage(){{
  var url = document.getElementById('url_input').value;
  var dynamic = document.getElementById('dynamic_flag').checked;
  window.location = '/?url=' + encodeURIComponent(url) + '&dynamic=' + dynamic;
}}
function exportContent(fmt){{
  var frame = document.getElementById('content_frame').contentWindow.document;
  var data = [];
  frame.querySelectorAll('[data-edit-id]').forEach(function(el){{
    if(el.tagName === 'IMG'){{
      data.push({{type:'img', src: el.getAttribute('src'), alt: el.getAttribute('alt') || ''}});
    }} else {{
      var tag = el.getAttribute('data-tag');
      var text = el.textContent.trim();
      if(text) data.push({{type: tag, text: text}});
    }}
  }});
  fetch('/export/' + fmt, {{
    method: 'POST',
    headers: {{'Content-Type': 'application/json'}},
    body: JSON.stringify({{blocks: data}})
  }}).then(function(resp){{ return resp.blob(); }})
    .then(function(blob){{
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'contenu_edite.' + fmt;
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


def is_editable_text(node) -> bool:
    if not node.strip():
        return False
    parent = node.parent
    while parent is not None:
        if parent.name in IGNORED_ANCESTOR_TAGS:
            return False
        parent = parent.parent
    return True


def make_editable_html(url: str, dynamic: bool) -> str:
    html = fetch_html(url, dynamic)
    soup = BeautifulSoup(html, "html.parser")
    for bad in soup(["script"]):
        bad.decompose()
    absolutize(soup, url)

    counter = 0
    for text_node in list(soup.find_all(string=True)):
        if isinstance(text_node, NavigableString) and is_editable_text(str(text_node)):
            parent_tag = text_node.parent.name if text_node.parent else "p"
            heading_tag = parent_tag if parent_tag in ("h1", "h2", "h3", "h4") else "p"
            span = soup.new_tag("span")
            span["data-edit-id"] = f"t{counter}"
            span["data-tag"] = heading_tag
            span["contenteditable"] = "true"
            span["style"] = "outline:1px dashed rgba(79,70,229,.4);outline-offset:1px;"
            span.string = str(text_node)
            text_node.replace_with(span)
            counter += 1

    for img in soup.find_all("img"):
        img["data-edit-id"] = f"i{counter}"
        img["style"] = (img.get("style", "") or "") + ";cursor:pointer;outline:1px dashed rgba(79,70,229,.4);"
        img["onclick"] = "var u=prompt('Nouvelle URL image :', this.src); if(u){this.src=u;}"
        counter += 1

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
            iframe_srcdoc = f"<p style='padding:20px;color:red'>Erreur: {e}</p>".replace('"', "&quot;")
    return PAGE_SHELL.format(
        url_value=url,
        dynamic_checked="checked" if dynamic else "",
        iframe_srcdoc=iframe_srcdoc,
    )


@app.route("/export/<fmt>", methods=["POST"])
def export(fmt):
    payload = request.get_json()
    blocks = payload.get("blocks", [])

    if fmt == "json":
        buf = io.BytesIO(json.dumps(blocks, ensure_ascii=False, indent=2).encode("utf-8"))
        return send_file(buf, mimetype="application/json", as_attachment=True, download_name="contenu_edite.json")

    if fmt == "md":
        lines = []
        for b in blocks:
            if b["type"] in ("h1", "h2", "h3", "h4"):
                level = int(b["type"][1])
                lines.append(f"{'#' * level} {b['text']}")
            elif b["type"] == "img":
                lines.append(f"![{b.get('alt','')}]({b['src']})")
            else:
                lines.append(b.get("text", ""))
            lines.append("")
        buf = io.BytesIO("\n".join(lines).encode("utf-8"))
        return send_file(buf, mimetype="text/markdown", as_attachment=True, download_name="contenu_edite.md")

    if fmt == "docx":
        doc = Document()
        for b in blocks:
            if b["type"] in ("h1", "h2", "h3", "h4"):
                doc.add_heading(b["text"], level=int(b["type"][1]))
            elif b["type"] == "img":
                doc.add_paragraph(f"[Image: {b['src']}]")
            else:
                doc.add_paragraph(b.get("text", ""))
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                          as_attachment=True, download_name="contenu_edite.docx")

    return jsonify({"error": "format inconnu"}), 400


if __name__ == "__main__":
    app.run(debug=True, port=5000)