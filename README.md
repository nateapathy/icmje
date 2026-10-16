# ICMJE Disclosure Helper

A browser-only app that walks authors through the ICMJE Disclosure of Interest form and downloads the completed form as a Word file in the original layout.

Live: https://nateapathy.github.io/icmje/

- Answers stay in the browser (localStorage); nothing is sent anywhere.
- `template.docx` is the blank ICMJE form; `fill.js` fills it in the browser with JSZip.
- Edit `app.src.html` or `fill.js`, then run `python3 build.py` to regenerate `index.html`.

Unofficial helper; not affiliated with ICMJE.
