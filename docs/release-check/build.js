// Builds the release check from checks.js (the one source):
//   docs/release-check/release-check.html          the web checklist (published as a claude.ai page)
//   docs/release-check/PC_Tracker_Release_Check.pdf the printable checklist (needs Playwright; skipped without)
//   .github/ISSUE_TEMPLATE/release-check.md         the GitHub issue version
// Usage: node docs/release-check/build.js
const fs = require("fs");
const path = require("path");

const HERE = __dirname;
const ROOT = path.resolve(HERE, "..", "..");
const DATA = require("./checks.js");

// ---------- numbering (same as the page: stage index . check index) ----------
const ORDER = [];
DATA.stages.forEach((st, si) => st.checks.forEach((c, ci) => ORDER.push({ st, si, c, num: si + "." + (ci + 1) })));
const ids = new Set();
ORDER.forEach(o => {
  if (ids.has(o.c.id)) throw new Error("Two checks share the id " + o.c.id);
  ids.add(o.c.id);
  if (!/^[A-Za-z0-9_-]+$/.test(o.c.id)) throw new Error("Check id " + o.c.id + ": letters, digits, - and _ only");
  if (!o.c.title || !Array.isArray(o.c.do) || !Array.isArray(o.c.see)) throw new Error("Check " + o.c.id + " needs title, do, see");
});
const esc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

// ---------- 1. the web page ----------
function buildPage() {
  const tpl = fs.readFileSync(path.join(HERE, "page.template.html"), "utf8");
  const marker = "var DATA = __CHECKS_DATA__;";
  if (tpl.split(marker).length !== 2) throw new Error("page.template.html must have the DATA line exactly once");
  const LS = String.fromCharCode(0x2028), PS = String.fromCharCode(0x2029); // separators JSON allows but a script does not
  const json = JSON.stringify(DATA).replace(/</g, "\\u003c").split(LS).join("\\u2028").split(PS).join("\\u2029");
  const out = tpl.replace(marker, () => "var DATA = " + json + ";");
  fs.writeFileSync(path.join(HERE, "release-check.html"), out);
  return "release-check.html: written (" + ORDER.length + " checks)";
}

// ---------- 2. the GitHub issue template ----------
function buildIssueTemplate() {
  const L = [];
  L.push("---", "name: Release check", "about: Tick through every check on a copy before a new version goes to schools",
    'title: "Release check: (version date)"', "---", "");
  L.push("<!--", "Made by docs/release-check/build.js from docs/release-check/checks.js: change the checks THERE and",
    "run the build (this file, the web checklist and the printable PDF are made from it). One issue per release;",
    "ticks save by themselves; anything that fails goes in a comment.", "-->", "");
  L.push("The easier way to do this: the web checklist (steps, notes, a report Claude can read). Ask Claude for its link.");
  L.push("How to put the new version on a copy: [DEPLOY.md](https://github.com/jskeller29/Parent-coordinator-sheet/blob/Test-Version/DEPLOY.md), sections 1 to 3.");
  L.push("Test files: [test-data](https://github.com/jskeller29/Parent-coordinator-sheet/tree/Test-Version/test-data) (its README says what each one should show).");
  L.push("Something fails? Write it as a comment on this issue: the step, what you saw, a screenshot, and any red error from Extensions → Apps Script → **Executions**.");
  DATA.stages.forEach((st, si) => {
    L.push("", "### " + (st.safety ? "" : si + ". ") + st.title, "_" + st.where + "_");
    st.checks.forEach((c, ci) => {
      const num = si + "." + (ci + 1);
      let line = "- [ ] **" + num + " " + c.title + ".** " + c.do.join(" ") + " **You should see:** " + c.see.join(" ");
      if (c.why) line += " _Why: " + c.why + "_";
      if (c.timing) line += " (" + c.timing + ": ___)";
      L.push(line);
    });
  });
  const file = path.join(ROOT, ".github", "ISSUE_TEMPLATE", "release-check.md");
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, L.join("\n") + "\n");
  return ".github/ISSUE_TEMPLATE/release-check.md: written";
}

// ---------- 3. the printable PDF ----------
function printHtml() {
  const stages = DATA.stages.map((st, si) => {
    const checks = st.checks.map((c, ci) => {
      const num = si + "." + (ci + 1);
      return `<div class="check">
  <div class="ch"><span class="num">${num}</span><span class="t">${esc(c.title)}</span>
    <span class="boxes"><span class="box"></span>Works <span class="box"></span>Problem <span class="box"></span>Skip</span></div>
  <div class="cols"><div><p class="lbl">Do</p><ol>${c.do.map(x => "<li>" + esc(x) + "</li>").join("")}</ol></div>
  <div><p class="lbl">You should see</p><ul>${c.see.map(x => "<li>" + esc(x) + "</li>").join("")}</ul></div></div>
  ${c.why ? `<p class="why"><b>Why:</b> ${esc(c.why)}</p>` : ""}
  ${c.timing ? `<p class="timing">${esc(c.timing)}: <span class="blank"></span> seconds</p>` : ""}
  <div class="notes"><span>Notes</span><span class="rule"></span></div><div class="notes"><span></span><span class="rule"></span></div>
</div>`;
    }).join("\n");
    return `<section class="stage${st.safety ? " safety" : ""}"><h2>${st.safety ? "" : `<span class="sn">${si}</span>`}${esc(st.title)}</h2>
<p class="where">${esc(st.where)}</p>${checks}</section>`;
  }).join("\n");
  return `<!doctype html><html><head><meta charset="utf-8">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@400;600;700&family=Source+Serif+4:opsz,wght@8..60,600;8..60,700&display=swap">
<style>
  @page { size: Letter; margin: 0.55in 0.6in 0.65in; }
  :root { --ink: #1b2430; --muted: #566173; --line: #c9d1db; --accent: #183e69; --warn: #7a5300; --warn-soft: #fbf1dc; }
  body { font: 10pt/1.38 "Source Sans 3", "Segoe UI", Arial, sans-serif; color: var(--ink); margin: 0; }
  h1, h2 { font-family: "Source Serif 4", Georgia, serif; font-weight: 600; margin: 0; }
  .top { border-bottom: 2px solid var(--ink); padding-bottom: 8pt; margin-bottom: 4pt; }
  .eyebrow { font-size: 8pt; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; color: var(--muted); margin: 0 0 3pt; }
  h1 { font-size: 22pt; }
  .meta { display: flex; justify-content: space-between; gap: 12pt; margin-top: 6pt; font-size: 10pt; }
  .meta .blank { display: inline-block; width: 150pt; border-bottom: 1px solid var(--ink); }
  .how { margin: 8pt 0 0; padding-left: 14pt; color: var(--ink); }
  .how li { margin: 1pt 0; }
  .stage { margin-top: 14pt; }
  .stage h2 { font-size: 14pt; border-bottom: 1.5px solid var(--ink); padding-bottom: 3pt; break-after: avoid; }
  .stage h2 .sn { font-family: "Source Sans 3", Arial, sans-serif; font-weight: 700; color: var(--accent); margin-right: 7pt; }
  .where { margin: 3pt 0 4pt; color: var(--muted); break-after: avoid; }
  .stage.safety { background: var(--warn-soft); border: 1px solid var(--warn); padding: 8pt 10pt 4pt; }
  .stage.safety h2 { color: var(--warn); border-bottom-color: var(--warn); }
  .check { border-bottom: 1px solid var(--line); padding: 6pt 0 5pt; break-inside: avoid; }
  .ch { display: flex; align-items: baseline; gap: 8pt; }
  .num { font-weight: 700; color: var(--muted); min-width: 24pt; }
  .t { font-weight: 700; font-size: 10.5pt; flex: 1; }
  .boxes { white-space: nowrap; font-size: 9pt; color: var(--muted); }
  .box { display: inline-block; width: 8pt; height: 8pt; border: 1px solid var(--ink); margin: 0 3pt 0 8pt; vertical-align: -1pt; }
  .cols { display: grid; grid-template-columns: 1fr 1fr; gap: 10pt; margin: 3pt 0 0 32pt; }
  .lbl { margin: 0 0 1pt; font-size: 7pt; font-weight: 700; letter-spacing: .07em; text-transform: uppercase; color: var(--muted); }
  ol, ul { margin: 0; padding-left: 12pt; }
  li { margin: 0.5pt 0; }
  .why, .timing { margin: 3pt 0 0 32pt; color: var(--muted); font-size: 9pt; }
  .timing .blank { display: inline-block; width: 40pt; border-bottom: 1px solid var(--ink); }
  .notes { display: flex; gap: 6pt; margin: 6pt 0 0 32pt; font-size: 8.5pt; color: var(--muted); }
  .notes span:first-child { width: 28pt; }
  .notes .rule { flex: 1; border-bottom: 1px solid var(--line); }
</style></head><body>
<div class="top">
  <p class="eyebrow">PC Tracker · before a version goes to schools</p>
  <h1>${esc(DATA.title)}</h1>
  <div class="meta"><span>Version tested: <span class="blank"></span></span><span>Date: <span class="blank" style="width:90pt"></span></span></div>
  <ol class="how">
    <li>Go through the stages in order, on copies of the tracker, never on your live sheet.</li>
    <li>Tick Works, Problem or Skip for each check. For a Problem, write what you saw under Notes.</li>
    <li>Done? Tell Claude what failed (the check numbers and your notes), or tick the same checks in the web checklist so Claude can read them.</li>
  </ol>
</div>
${stages}
</body></html>`;
}

async function buildPdf() {
  let chromium;
  try { ({ chromium } = require("playwright")); } catch (e) { return "PC_Tracker_Release_Check.pdf: skipped (Playwright isn't installed)"; }
  const tmp = path.join(require("os").tmpdir(), "release-check-print-" + process.pid + ".html");
  fs.writeFileSync(tmp, printHtml());
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    await page.goto("file://" + tmp, { waitUntil: "networkidle" });
    await page.evaluate(() => document.fonts.ready);
    await page.pdf({
      path: path.join(HERE, "PC_Tracker_Release_Check.pdf"), format: "Letter", printBackground: true,
      displayHeaderFooter: true, headerTemplate: "<span></span>",
      footerTemplate: '<div style="width:100%;font:8px Arial,sans-serif;color:#8a94a3;padding:0 0.6in;display:flex;justify-content:space-between">' +
        '<span>' + esc(DATA.title) + "</span><span>Page <span class=\"pageNumber\"></span> of <span class=\"totalPages\"></span></span></div>",
      margin: { top: "0.55in", bottom: "0.65in", left: "0.6in", right: "0.6in" }
    });
  } finally {
    await browser.close();
    fs.unlinkSync(tmp);
  }
  return "PC_Tracker_Release_Check.pdf: written";
}

(async () => {
  console.log(buildPage());
  console.log(buildIssueTemplate());
  console.log(await buildPdf());
})().catch(e => { console.error(e); process.exit(1); });
