// Fills the ICMJE disclosure template (.docx) with a list of answers.
// Pure: takes JSZip, DOMParser and XMLSerializer so it runs in the browser and in node tests.
//
// answers = {
//   date: "YYYY-MM-DD", name, title, msno,
//   items: [ { none: bool, rows: [ [entityText, specText], ... ] } x 13 ],
//   certify: bool
// }
function makeIcmjeFiller(JSZip, DOMParser, XMLSerializer) {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
  const XMLNS = "http://www.w3.org/XML/1998/namespace";

  const kids = (el, name) => Array.from(el.childNodes).filter(n => n.nodeType === 1 && n.localName === name);
  const all = (el, ns, name) => Array.from(el.getElementsByTagNameNS(ns, name));

  function wEl(doc, name, attrs) {
    const e = doc.createElementNS(W, "w:" + name);
    for (const [k, v] of Object.entries(attrs || {})) e.setAttributeNS(W, "w:" + k, v);
    return e;
  }

  // Run properties matching the form's body text (Calibri Light via the major theme font).
  function bodyRPr(doc) {
    const rPr = wEl(doc, "rPr");
    rPr.appendChild(wEl(doc, "rFonts", { asciiTheme: "majorHAnsi", hAnsiTheme: "majorHAnsi", cstheme: "majorHAnsi" }));
    rPr.appendChild(wEl(doc, "szCs", { val: "22" }));
    return rPr;
  }

  function appendText(doc, p, text, rPr) {
    const r = wEl(doc, "r");
    if (rPr) r.appendChild(rPr);
    String(text).split(/\r?\n/).forEach((line, i) => {
      if (i) r.appendChild(wEl(doc, "br"));
      const t = wEl(doc, "t");
      t.setAttributeNS(XMLNS, "xml:space", "preserve");
      t.textContent = line;
      r.appendChild(t);
    });
    p.appendChild(r);
  }

  function clearRuns(p) {
    kids(p, "r").forEach(r => p.removeChild(r));
  }

  function setCellText(doc, tc, text) {
    const p = kids(tc, "p")[0];
    clearRuns(p);
    // Drop the small grey size the hint row's paragraph mark carries.
    const pPr = kids(p, "pPr")[0];
    const mark = pPr && kids(pPr, "rPr")[0];
    if (mark) ["sz", "szCs", "color"].forEach(n => kids(mark, n).forEach(x => mark.removeChild(x)));
    if (text) appendText(doc, p, text, bodyRPr(doc));
  }

  function fillSdt(doc, sdt, text, isoDate) {
    if (!text) return;
    const sdtPr = kids(sdt, "sdtPr")[0];
    kids(sdtPr, "showingPlcHdr").forEach(x => sdtPr.removeChild(x));
    if (isoDate) {
      const d = kids(sdtPr, "date")[0];
      if (d) d.setAttributeNS(W, "w:fullDate", isoDate + "T00:00:00Z");
    }
    const p = all(kids(sdt, "sdtContent")[0], W, "p")[0];
    clearRuns(p);
    appendText(doc, p, text, null);
  }

  function setCheckbox(sdt, on) {
    const c = sdt.getElementsByTagNameNS(W14, "checked")[0];
    c.setAttributeNS(W14, "w14:val", on ? "1" : "0");
    const t = all(sdt, W, "t")[0];
    t.textContent = on ? "☒" : "☐";
  }

  function stripIds(el) {
    all(el, W, "p").forEach(p => {
      p.removeAttributeNS(W14, "paraId");
      p.removeAttributeNS(W14, "textId");
    });
  }

  function fillEntityTable(doc, tbl, rows) {
    let trs = kids(tbl, "tr");
    while (trs.length < rows.length) {
      const last = trs[trs.length - 1];
      const clone = last.cloneNode(true);
      stripIds(clone);
      kids(clone, "tc").forEach(tc => setCellText(doc, tc, ""));
      last.parentNode.insertBefore(clone, last.nextSibling);
      trs = kids(tbl, "tr");
    }
    rows.forEach((row, i) => {
      const tcs = kids(trs[i], "tc");
      setCellText(doc, tcs[0], row[0] || "");
      setCellText(doc, tcs[1], row[1] || "");
    });
  }

  function fmtDate(iso) {
    const [y, m, d] = iso.split("-").map(Number);
    return `${m}/${d}/${y}`;
  }

  return async function fill(templateB64, a) {
    const zip = await JSZip.loadAsync(templateB64, { base64: true });
    const xml = await zip.file("word/document.xml").async("string");
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const body = doc.getElementsByTagNameNS(W, "body")[0];
    const header = kids(body, "tbl")[0];

    const hs = all(header, W, "sdt");
    if (hs.length !== 4) throw new Error("Template header has changed (expected 4 fields).");
    if (a.date) fillSdt(doc, hs[0], fmtDate(a.date), a.date);
    fillSdt(doc, hs[1], a.name);
    fillSdt(doc, hs[2], a.title);
    fillSdt(doc, hs[3], a.msno);

    const boxes = all(doc, W14, "checkbox").map(cb => cb.parentNode.parentNode);
    // Each item's entity table is the next table after its "None" checkbox table in the same cell.
    const tables = boxes.slice(0, 13).map(sdt => {
      let t = sdt;
      while (t && t.localName !== "tbl") t = t.parentNode;
      let n = t && t.nextSibling;
      while (n && !(n.nodeType === 1 && n.localName === "tbl")) n = n.nextSibling;
      return n;
    });
    if (boxes.length !== 14 || tables.some(t => !t)) throw new Error("Template layout has changed.");

    a.items.forEach((it, i) => {
      setCheckbox(boxes[i], !!it.none);
      if (!it.none && it.rows.length) fillEntityTable(doc, tables[i], it.rows);
    });
    setCheckbox(boxes[13], !!a.certify);

    let out = new XMLSerializer().serializeToString(doc);
    if (!out.startsWith("<?xml")) out = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + out;
    zip.file("word/document.xml", out);
    return zip;
  };
}
if (typeof module !== "undefined") module.exports = makeIcmjeFiller;
