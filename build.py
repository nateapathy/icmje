"""Build index.html: inline fill.js and the base64 template into app.src.html."""
import base64, pathlib
d = pathlib.Path(__file__).parent
html = (d / "app.src.html").read_text()
html = html.replace("__FILL_JS__", (d / "fill.js").read_text())
html = html.replace("__TEMPLATE_B64__", base64.b64encode((d / "template.docx").read_bytes()).decode())
# Standalone document shell so the page works outside claude.ai (e.g. GitHub Pages).
head = ('<!doctype html><html lang="en"><head><meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
        '<style>:root{padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}'
        'body{margin:0}img{max-width:100%}[hidden]{display:none!important}</style></head><body>\n')
(d / "index.html").write_text(head + html + "\n</body></html>\n")
print("wrote index.html", len(html), "bytes")
