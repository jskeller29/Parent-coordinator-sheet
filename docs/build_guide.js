// ======================================================================
// Builds the User Guide from docs/USER_GUIDE.md (the only file to edit):
//   docs/guide.html                   (the page the PDF is printed from)
//   docs/PC_Tracker_User_Guide.pdf    (needs Playwright + Chromium)
//   docs/PC_Tracker_User_Guide.docx   (needs the "docx" package: npm install docx)
//
//   node docs/build_guide.js
//
// A documentation helper only: it is NOT part of the Apps Script project.
// Look: the "District" style of the tracker's own screens (Styles.html):
// navy #183e69, Source Sans 3 text, Source Serif 4 headings, thin grey lines,
// 3px corners, no colored boxes or emoji of its own.
// Markdown it understands (keep USER_GUIDE.md to this): # title, ## part,
// ### section, #### sub-section, paragraphs, "- " and "1. " lists (continued
// lines indented), | tables |, > [!NOTE] / [!IMPORTANT] / [!WARNING] boxes,
// **bold**, *italic*, `code`, [text](link), <!-- notes that never print -->.
// ======================================================================
const fs = require("fs");
const path = require("path");

const DIR = __dirname;
const SRC = path.join(DIR, "USER_GUIDE.md");
const OUT_HTML = path.join(DIR, "guide.html");
const OUT_PDF = path.join(DIR, "PC_Tracker_User_Guide.pdf");
const OUT_DOCX = path.join(DIR, "PC_Tracker_User_Guide.docx");

// Same tokens as Styles.html
const C = { navy: "183e69", ink: "1b2430", muted: "566173", faint: "8a94a3", line: "d3d9e1", panel: "f3f5f8",
            soft: "e3e8ee", danger: "a3312a", dangerSoft: "f8e9e7" };
const FONT = "Source Sans 3", HEAD = "Source Serif 4";
const NOTE_LABEL = { NOTE: "Note", IMPORTANT: "Important", WARNING: "Warning" };

// ---------------------------------------------------------------- parse
function slug(text) {
  return text.toLowerCase().replace(/[^a-z0-9 -]/g, "").trim().replace(/\s+/g, "-");
}

function inline(text) {
  const out = [];
  const re = /\*\*(.+?)\*\*|\*(.+?)\*|`(.+?)`|\[(.+?)\]\((.+?)\)/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push({ text: text.slice(last, m.index) });
    if (m[1] !== undefined) out.push({ text: m[1], bold: true });
    else if (m[2] !== undefined) out.push({ text: m[2], italic: true });
    else if (m[3] !== undefined) out.push({ text: m[3], code: true });
    else out.push({ text: m[4], link: m[5] });
    last = re.lastIndex;
  }
  if (last < text.length) out.push({ text: text.slice(last) });
  return out;
}

function parse(md) {
  md = md.replace(/<!--[\s\S]*?-->/g, "");
  const lines = md.split(/\r?\n/);
  const doc = { title: "", front: [], blocks: [] };
  let i = 0;
  const isSpecial = l => /^(#{1,4} |>|\||\d+\. |- )/.test(l);

  // Title page: "# Title", then its paragraphs until the first "##"
  while (i < lines.length && !lines[i].startsWith("# ")) i++;
  doc.title = (lines[i++] || "# Guide").slice(2).trim();
  let para = [];
  for (; i < lines.length && !lines[i].startsWith("## "); i++) {
    if (lines[i].trim()) para.push(lines[i].trim());
    else if (para.length) { doc.front.push(para); para = []; }
  }
  if (para.length) doc.front.push(para);

  let olCount = 0;
  while (i < lines.length) {
    const l = lines[i];
    if (!l.trim()) { i++; continue; }
    const h = l.match(/^(#{2,4}) (.+)$/);
    if (h) {
      const level = h[1].length, text = h[2].trim();
      doc.blocks.push({ t: "h" + level, text, id: slug(text) });
      i++; continue;
    }
    if (l.startsWith(">")) {
      const body = [];
      while (i < lines.length && lines[i].startsWith(">")) body.push(lines[i++].replace(/^>\s?/, ""));
      const kind = (body[0].match(/^\[!(NOTE|IMPORTANT|WARNING)\]/) || [, "NOTE"])[1];
      if (/^\[!/.test(body[0])) body.shift();
      const paras = body.join("\n").split(/\n\s*\n/).map(p => inline(p.replace(/\s*\n\s*/g, " ").trim())).filter(p => p.length);
      doc.blocks.push({ t: "note", kind, paras });
      continue;
    }
    if (l.startsWith("|")) {
      const rows = [];
      while (i < lines.length && lines[i].startsWith("|")) {
        rows.push(lines[i++].trim().replace(/^\||\|$/g, "").split("|").map(c => inline(c.trim())));
      }
      doc.blocks.push({ t: "table", head: rows[0], rows: rows.slice(2) });
      continue;
    }
    const li = l.match(/^(\d+\.|-) (.*)$/);
    if (li) {
      const ordered = li[1] !== "-", items = [];
      while (i < lines.length) {
        const m = lines[i].match(/^(\d+\.|-) (.*)$/);
        if (m && (m[1] !== "-") === ordered) { items.push(m[2].trim()); i++; }
        else if (items.length && /^\s{2,}\S/.test(lines[i])) { items[items.length - 1] += " " + lines[i].trim(); i++; }
        else break;
      }
      doc.blocks.push({ t: ordered ? "ol" : "ul", items: items.map(inline), n: ordered ? ++olCount : 0 });
      continue;
    }
    const buf = [];
    while (i < lines.length && lines[i].trim() && !isSpecial(lines[i])) buf.push(lines[i++].trim());
    doc.blocks.push({ t: "p", runs: inline(buf.join(" ")) });
  }
  return doc;
}

// ---------------------------------------------------------------- HTML (and the PDF)
const esc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

function runsHtml(runs) {
  return runs.map(r => {
    let s = esc(r.text);
    if (r.code) s = `<code>${s}</code>`;
    if (r.bold) s = `<b>${s}</b>`;
    if (r.italic) s = `<i>${s}</i>`;
    if (r.link) s = `<a href="${esc(r.link)}">${s}</a>`;
    return s;
  }).join("");
}

function toHtml(doc, updated) {
  const parts = doc.blocks.filter(b => b.t === "h2");
  const toc = [];
  doc.blocks.forEach(b => {
    if (b.t === "h2") toc.push(`<li class="t2"><a href="#${b.id}">${esc(b.text)}</a></li>`);
    if (b.t === "h3") toc.push(`<li class="t3"><a href="#${b.id}">${esc(b.text)}</a></li>`);
  });
  const body = doc.blocks.map(b => {
    switch (b.t) {
      case "h2": return `<h2 id="${b.id}">${esc(b.text)}</h2>`;
      case "h3": return `<h3 id="${b.id}">${esc(b.text)}</h3>`;
      case "h4": return `<h4 id="${b.id}">${esc(b.text)}</h4>`;
      case "p": return `<p>${runsHtml(b.runs)}</p>`;
      case "ul": return `<ul>${b.items.map(it => `<li>${runsHtml(it)}</li>`).join("")}</ul>`;
      case "ol": return `<ol>${b.items.map(it => `<li>${runsHtml(it)}</li>`).join("")}</ol>`;
      case "note": return `<div class="note ${b.kind.toLowerCase()}"><div class="lab">${NOTE_LABEL[b.kind]}</div>${b.paras.map(p => `<p>${runsHtml(p)}</p>`).join("")}</div>`;
      case "table": return `<table><thead><tr>${b.head.map(c => `<th>${runsHtml(c)}</th>`).join("")}</tr></thead><tbody>` +
        b.rows.map(r => `<tr>${r.map(c => `<td>${runsHtml(c)}</td>`).join("")}</tr>`).join("") + `</tbody></table>`;
    }
    return "";
  }).join("\n");

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<!-- Made by docs/build_guide.js from docs/USER_GUIDE.md. Don't edit by hand. -->
<title>${esc(doc.title)}</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Source+Sans+3:ital,wght@0,400;0,600;0,700;1,400&family=Source+Serif+4:opsz,wght@8..60,600&display=swap">
<style>
  :root { --navy:#${C.navy}; --ink:#${C.ink}; --muted:#${C.muted}; --faint:#${C.faint}; --line:#${C.line}; --panel:#${C.panel};
          --soft:#${C.soft}; --danger:#${C.danger}; --danger-soft:#${C.dangerSoft}; }
  @page { size: Letter; margin: 0.8in 0.85in 0.85in;
          @bottom-left { content: "${esc(doc.title)}"; font: 8pt "${FONT}", sans-serif; color: #${C.faint}; }
          @bottom-right { content: counter(page); font: 8pt "${FONT}", sans-serif; color: #${C.faint}; } }
  @page :first { @bottom-left { content: none; } @bottom-right { content: none; } }
  * { box-sizing: border-box; }
  html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  body { margin: 0; font: 10.5pt/1.5 "${FONT}", "Segoe UI", Arial, sans-serif; color: var(--ink); }
  h1, h2, h3 { font-family: "${HEAD}", Georgia, serif; font-weight: 600; color: var(--ink); line-height: 1.25; }
  h2 { font-size: 19pt; color: var(--navy); margin: 0 0 14pt; padding-bottom: 6pt; border-bottom: 1px solid var(--line); break-before: page; break-after: avoid; }
  h3 { font-size: 13.5pt; margin: 20pt 0 6pt; break-after: avoid; }
  h4 { font-size: 10.5pt; font-weight: 700; margin: 14pt 0 4pt; break-after: avoid; }
  p { margin: 0 0 8pt; }
  ul, ol { margin: 0 0 9pt; padding-left: 20pt; }
  li { margin-bottom: 3pt; }
  a { color: var(--navy); }
  code { font: 9.5pt Consolas, monospace; background: var(--panel); padding: 0 3px; border-radius: 3px; }
  table { width: 100%; border-collapse: collapse; margin: 4pt 0 12pt; font-size: 9.5pt; break-inside: auto; }
  tr { break-inside: avoid; }
  th { text-align: left; font-weight: 600; background: var(--panel); border-bottom: 1px solid var(--line); padding: 5pt 7pt; }
  td { border-bottom: 1px solid var(--line); padding: 5pt 7pt; vertical-align: top; }
  .note { border-left: 3px solid var(--faint); background: var(--panel); border-radius: 0 3px 3px 0; padding: 7pt 10pt 1pt; margin: 6pt 0 12pt; break-inside: avoid; }
  .note.important { border-left-color: var(--navy); }
  .note.warning { border-left-color: var(--danger); background: var(--danger-soft); }
  .note .lab { font-size: 8pt; font-weight: 700; letter-spacing: .06em; text-transform: uppercase; color: var(--muted); margin-bottom: 2pt; }
  .note.warning .lab { color: var(--danger); }
  /* title page */
  .cover { padding-top: 2.3in; }
  .cover .kicker { font-size: 9pt; font-weight: 600; letter-spacing: .08em; text-transform: uppercase; color: var(--navy); }
  .cover h1 { font-size: 32pt; margin: 6pt 0 10pt; }
  .cover .sub { font-size: 14pt; color: var(--muted); margin: 0 0 22pt; }
  .cover .rule { width: 56px; height: 3px; background: var(--navy); margin: 0 0 22pt; }
  .cover .front p { font-size: 10pt; color: var(--muted); margin: 0 0 4pt; }
  .cover .date { margin-top: 18pt; font-size: 10pt; color: var(--muted); }
  /* contents */
  .contents h2 { break-before: page; }
  .jump { display: flex; flex-wrap: wrap; gap: 6pt; margin: 0 0 16pt; }
  .jump a { text-decoration: none; border: 1px solid var(--line); border-radius: 3px; padding: 4pt 9pt; font-size: 9.5pt; font-weight: 600; }
  .toc { list-style: none; padding: 0; margin: 0; }
  .toc li { margin: 0; }
  .toc .t2 { font-weight: 600; margin-top: 8pt; padding-top: 6pt; border-top: 1px solid var(--line); }
  .toc .t2:first-child { border-top: 0; margin-top: 0; padding-top: 0; }
  .toc .t3 { padding-left: 14pt; font-size: 9.5pt; }
  .toc a { color: var(--ink); text-decoration: none; }
  .toc .t3 a { color: var(--muted); }
</style>
</head>
<body>
<section class="cover">
  <div class="kicker">PC Tracker</div>
  <h1>${esc(doc.title)}</h1>
  ${doc.front[0] ? `<p class="sub">${runsHtml(inline(doc.front[0].join(" ")))}</p>` : ""}
  <div class="rule"></div>
  <div class="front">${doc.front.slice(1).map(p => p.map(l => `<p>${runsHtml(inline(l))}</p>`).join("")).join("")}</div>
  <div class="date">Updated ${esc(updated)}</div>
</section>
<section class="contents">
  <h2 id="contents">Contents</h2>
  <div class="jump">${parts.map(b => `<a href="#${b.id}">${esc(b.text.replace(/^Part \d+:\s*/, ""))}</a>`).join("")}</div>
  <ul class="toc">${toc.join("")}</ul>
</section>
${body}
</body>
</html>
`;
}

// Chromium writes the PDF's jump links as NAMED destinations, which simple viewers (Drive's
// and GitHub's previews) often ignore. This turns each one into a plain "page N, this
// height" link, which every viewer follows. Needs Python with pypdf (pip install pypdf);
// without it the PDF is kept as Chromium made it.
const FIX_LINKS_PY = `
import sys
from pypdf import PdfReader, PdfWriter
from pypdf.generic import ArrayObject, NameObject, NumberObject, NullObject
path = sys.argv[1]
w = PdfWriter(clone_from=PdfReader(path))
pages = {p.indirect_reference.idnum: p.indirect_reference for p in w.pages}
dests = {}
for name, d in w.named_destinations.items():
    page = d.page.idnum if hasattr(d.page, "idnum") else d.page
    if page in pages: dests[name] = (pages[page], d.left or 0, d.top)
fixed = 0
for page in w.pages:
    for a in (page.get("/Annots") or []):
        a = a.get_object()
        d = a.get("/Dest")
        if a.get("/Subtype") != "/Link" or d is None or isinstance(d, ArrayObject): continue
        key = str(d)
        if key not in dests: continue
        ref, left, top = dests[key]
        a[NameObject("/Dest")] = ArrayObject([ref, NameObject("/XYZ"), NumberObject(int(left)),
                                              NumberObject(int(top)) if top is not None else NullObject(), NumberObject(0)])
        fixed += 1
w.write(path)
print(fixed)
`;

function fixPdfLinks(pdfPath) {
  try {
    const n = require("child_process").execFileSync("python3", ["-c", FIX_LINKS_PY, pdfPath], { encoding: "utf8" }).trim();
    return n + " jump links made plain";
  } catch (e) {
    return "jump links left as made (needs python3 + pypdf)";
  }
}

async function toPdf(htmlPath, pdfPath) {
  let chromium;
  try { ({ chromium } = require("playwright")); } catch (e) { return "skipped (Playwright isn't installed)"; }
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    await page.goto("file://" + htmlPath, { waitUntil: "networkidle" });
    await page.evaluate(() => document.fonts.ready);
    await page.pdf({ path: pdfPath, format: "Letter", printBackground: true, preferCSSPageSize: true });
  } finally { await browser.close(); }
  return "written, " + fixPdfLinks(pdfPath);
}

// ---------------------------------------------------------------- Word (and Google Docs)
function toDocx(doc, updated) {
  const D = require("docx");
  const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, BorderStyle,
          ShadingType, PageBreak, LevelFormat, AlignmentType, Bookmark, InternalHyperlink, ExternalHyperlink,
          Footer, PageNumber, LineRuleType } = D;
  const PAGE_W = 12240 - 2 * 1224; // twips inside the margins
  const ids = new Map(); // heading id -> bookmark name (letters, digits, _)
  doc.blocks.forEach((b, n) => { if (/^h[234]$/.test(b.t)) ids.set(b.id, "s" + n); });

  const run = (r, o = {}) => new TextRun({ text: r.text, bold: r.bold || o.bold, italics: r.italic, font: r.code ? "Consolas" : undefined,
    size: o.size, color: o.color || (r.link ? C.navy : undefined), underline: r.link ? {} : undefined });
  const runs = (list, o) => list.map(r => {
    if (!r.link) return run(r, o);
    if (r.link.startsWith("#") && ids.has(r.link.slice(1))) return new InternalHyperlink({ anchor: ids.get(r.link.slice(1)), children: [run(r, o)] });
    return new ExternalHyperlink({ link: r.link, children: [run(r, o)] });
  });
  const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const thin = { style: BorderStyle.SINGLE, size: 4, color: C.line };

  const out = [];
  // Title page
  out.push(new Paragraph({ spacing: { before: 3000, after: 80 }, children: [new TextRun({ text: "PC TRACKER", bold: true, size: 18, color: C.navy, characterSpacing: 30 })] }));
  out.push(new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: doc.title, font: HEAD, size: 64, color: C.ink })] }));
  if (doc.front[0]) out.push(new Paragraph({ spacing: { after: 360 }, children: runs(inline(doc.front[0].join(" ")), { size: 28, color: C.muted }) }));
  out.push(new Paragraph({ spacing: { after: 360 }, border: { top: { style: BorderStyle.SINGLE, size: 18, color: C.navy, space: 1 } }, indent: { right: PAGE_W - 900 }, children: [] }));
  doc.front.slice(1).forEach(p => p.forEach(l => out.push(new Paragraph({ spacing: { after: 60 }, children: runs(inline(l), { size: 20, color: C.muted }) }))));
  out.push(new Paragraph({ spacing: { before: 300 }, children: [new TextRun({ text: "Updated " + updated, size: 20, color: C.muted })] }));

  // Contents (links to the bookmarks on the headings)
  out.push(new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, children: [new TextRun("Contents")] }));
  doc.blocks.forEach(b => {
    if (b.t !== "h2" && b.t !== "h3") return;
    const top = b.t === "h2";
    out.push(new Paragraph({ spacing: { before: top ? 160 : 0, after: top ? 40 : 20 }, indent: { left: top ? 0 : 360 },
      children: [new InternalHyperlink({ anchor: ids.get(b.id), children: [new TextRun({ text: b.text, bold: top, size: top ? 21 : 19, color: top ? C.ink : C.muted })] })] }));
  });

  doc.blocks.forEach(b => {
    if (b.t === "h2" || b.t === "h3" || b.t === "h4") {
      const level = { h2: HeadingLevel.HEADING_1, h3: HeadingLevel.HEADING_2, h4: HeadingLevel.HEADING_3 }[b.t];
      out.push(new Paragraph({ heading: level, pageBreakBefore: b.t === "h2", keepNext: true,
        children: [new Bookmark({ id: ids.get(b.id), children: [new TextRun(b.text)] })] }));
    } else if (b.t === "p") {
      out.push(new Paragraph({ spacing: { after: 140 }, children: runs(b.runs) }));
    } else if (b.t === "ul" || b.t === "ol") {
      b.items.forEach(it => out.push(new Paragraph({ spacing: { after: 60 },
        numbering: b.t === "ul" ? { reference: "bullets", level: 0 } : { reference: "steps", level: 0, instance: b.n }, children: runs(it) })));
      out.push(new Paragraph({ spacing: { after: 60 }, children: [] }));
    } else if (b.t === "note") {
      const bar = b.kind === "WARNING" ? C.danger : b.kind === "IMPORTANT" ? C.navy : C.faint;
      const fill = b.kind === "WARNING" ? C.dangerSoft : C.panel;
      const box = { border: { left: { style: BorderStyle.SINGLE, size: 18, color: bar, space: 8 } }, shading: { type: ShadingType.CLEAR, fill, color: "auto" }, indent: { left: 160, right: 160 } };
      out.push(new Paragraph(Object.assign({ spacing: { before: 80, after: 0 }, keepNext: true,
        children: [new TextRun({ text: NOTE_LABEL[b.kind].toUpperCase(), bold: true, size: 16, color: b.kind === "WARNING" ? C.danger : C.muted, characterSpacing: 20 })] }, box)));
      b.paras.forEach((p, k) => out.push(new Paragraph(Object.assign({ spacing: { after: k === b.paras.length - 1 ? 200 : 80 }, children: runs(p) }, box))));
    } else if (b.t === "table") {
      const cols = b.head.length;
      const widths = cols === 2 ? [0.32, 0.68] : cols === 3 ? [0.24, 0.52, 0.24] : Array(cols).fill(1 / cols);
      const cell = (content, header, k) => new TableCell({
        width: { size: Math.round(PAGE_W * widths[k]), type: WidthType.DXA },
        shading: header ? { type: ShadingType.CLEAR, fill: C.panel, color: "auto" } : undefined,
        margins: { top: 80, bottom: 80, left: 110, right: 110 },
        borders: { top: none, left: none, right: none, bottom: thin },
        children: [new Paragraph({ children: runs(content, { bold: header, size: 19 }) })]
      });
      out.push(new Table({ width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: widths.map(w => Math.round(PAGE_W * w)),
        borders: { top: none, left: none, right: none, bottom: none, insideHorizontal: none, insideVertical: none },
        rows: [new TableRow({ tableHeader: true, children: b.head.map((c, k) => cell(c, true, k)) })]
          .concat(b.rows.map(r => new TableRow({ cantSplit: true, children: r.map((c, k) => cell(c, false, k)) }))) }));
      out.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
    }
  });

  const heading = (size, color, before, after) => ({ run: { font: HEAD, size, color }, paragraph: { spacing: { before, after } } });
  const document = new Document({
    creator: "PC Tracker", title: doc.title, description: doc.title,
    styles: {
      // line 288 = 1.2 lines, "auto" so big text (the title) gets taller lines instead of overlapping
      default: { document: { run: { font: FONT, size: 21, color: C.ink }, paragraph: { spacing: { line: 288, lineRule: LineRuleType.AUTO } } } },
      paragraphStyles: [
        Object.assign({ id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true }, heading(36, C.navy, 0, 200),
          { paragraph: { spacing: { before: 0, after: 200 }, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.line, space: 6 } } } }),
        Object.assign({ id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true }, heading(27, C.ink, 360, 100)),
        Object.assign({ id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true },
          { run: { font: FONT, size: 21, bold: true, color: C.ink }, paragraph: { spacing: { before: 240, after: 60 } } })
      ]
    },
    numbering: { config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 220 } } } }] },
      { reference: "steps", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 260 } } } }] }
    ] },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1152, bottom: 1152, left: 1224, right: 1224 } }, titlePage: true },
      footers: {
        default: new Footer({ children: [new Paragraph({ children: [new TextRun({ text: doc.title + "    ", size: 16, color: C.faint }),
          new TextRun({ children: [PageNumber.CURRENT], size: 16, color: C.faint })] })] }),
        first: new Footer({ children: [new Paragraph({ children: [] })] })
      },
      children: out
    }]
  });
  return Packer.toBuffer(document);
}

// ---------------------------------------------------------------- run
(async () => {
  const doc = parse(fs.readFileSync(SRC, "utf8"));
  const updated = new Date().toLocaleDateString("en-US", { month: "long", year: "numeric" });
  fs.writeFileSync(OUT_HTML, toHtml(doc, updated));
  console.log("guide.html: written");
  console.log("PC_Tracker_User_Guide.pdf: " + await toPdf(OUT_HTML, OUT_PDF));
  try {
    fs.writeFileSync(OUT_DOCX, await toDocx(doc, updated));
    console.log("PC_Tracker_User_Guide.docx: written");
  } catch (e) {
    console.log("PC_Tracker_User_Guide.docx: skipped (" + (/Cannot find module 'docx'/.test(e.message) ? "npm install docx" : e.message) + ")");
  }
})();
