// Screenshots of the restyled sidebars with sample data (needs Playwright + Chromium;
// optional, not part of run_all). Usage: node tests/render_sidebars.js [outDir]   (default tests/out)
const { chromium } = require('playwright');
const fs = require('fs');
const repo = require('path').resolve(__dirname, '..') + '/';
const styles = fs.readFileSync(repo + 'Styles.html', 'utf8');
function page(name, stub) {
  let html = fs.readFileSync(repo + name + '.html', 'utf8').replace("<?!= include_('Styles') ?>", styles);
  html = html.replace('<head>', '<head><script>' + stub + '</script>');
  const out = (process.argv[2] || require('path').join(__dirname, 'out')) + '/' + name + '.test.html';
  fs.mkdirSync(require('path').dirname(out), { recursive: true });
  fs.writeFileSync(out, html);
  return out;
}
const runStub = (responses) => `
  window.RESP = ${JSON.stringify(responses)};
  window.google = { script: { host: { close(){} }, run: (function make(ok, fail) {
    return new Proxy({}, { get(_, fn) {
      if (fn === 'withSuccessHandler') return h => make(h, fail);
      if (fn === 'withFailureHandler') return h => make(ok, h);
      return function () { const r = window.RESP[fn]; setTimeout(() => ok && ok(typeof r === 'function' ? r.apply(null, arguments) : JSON.parse(JSON.stringify(r === undefined ? null : r))), 30); };
    } });
  })() } };`;
const progress = (cur, done) => ['Student Data', 'Parent Square', 'NYCSA', 'Site / Class', 'Your Lists'].map((l, i) => ({ label: l, current: i === cur, status: i < cur ? (done[i] || 'done') : 'todo' }));
const ats = { active: true, stepIndex: 0, stepCount: 5, step: { key: 'ats', zone: 'Student Data (ATS / Insight)', required: true, manual: false, lists: false,
  howTo: ['Go to <b>TeachHub</b> and sign in.', 'In the search bar, look up and open <b>Insight</b>.', 'Go to <b>Tools → Downloader</b>.', 'Select only <b>Guardian Contact Information</b>, then download.'] },
  tabName: 'Student Data', lastUpdated: { text: '11/04/26', days: 8 }, done: [], progress: progress(0, {}), notice: null, pending: false };
const lists = { active: true, stepIndex: 4, stepCount: 5, step: { key: 'lists', zone: 'Your Lists', required: false, manual: false, lists: true,
  howTo: ['Any other list you get: bus routes, trip forms, after school, PTA…', 'Each list gets its own tab.'] },
  tabName: 'Your Lists', lastUpdated: { text: '11/10/26', days: 2 }, progress: progress(4, { 1: 'skipped' }),
  done: [{ name: 'Student Data', rows: 120, skipped: false }, { name: 'Parent Square Data', rows: 0, skipped: true }, { name: 'NYCSA Data', rows: 88, skipped: false }],
  lists: [{ tabId: 1, name: 'Bus Routes', tabName: '📥 Bus Routes', by: 'student', ready: true, updated: '11/10/26', stats: { rows: 40, matched: 36, review: 4, skipped: 0 } },
          { tabId: 2, name: 'PTA Members', tabName: '📥 PTA Members', by: 'guardian', ready: false, updated: '', stats: null }], notice: null };
const setupView = { phase: 'setup', list: { tabId: 1, name: 'Bus Routes', tabName: '📥 Bus Routes', ready: false, title: 'District 75 Fall Routes' }, kept: false, headerRow: 4, rowCount: 40,
  columns: [{ col: 1, letter: 'A', header: 'Student', samples: ['Bart Simpson'], hint: { osisKnown: 0, n: 40, comma: false } },
            { col: 2, letter: 'B', header: 'OSIS', samples: ['123456789'], hint: { osisKnown: 38, n: 40, comma: false } },
            { col: 3, letter: 'C', header: 'Route', samples: ['R12'], hint: { osisKnown: 0, n: 40, comma: false } }],
  roles: { osis: 2, student: { col: 1 }, guardian: null, phone: 0, email: 0, 'class': 0 }, by: 'student',
  vars: [{ col: 3, header: 'Route', label: 'Bus Route', show: true, column: true, blank: { mode: 'text', text: 'No bus' }, notOnList: { mode: 'nothing', text: '' } }] };
const opt = (t, text, sub, c, l) => ({ target: t, text: text, sub: sub, confidence: c, level: l });
const matchView = { phase: 'match', by: 'student', stats: { rows: 40, matched: 34, review: 6, skipped: 0, none: 1 }, how: { OSIS: 30, Name: 4 }, levels: { high: 90, medium: 75, min: 60 },
  picker: [{ target: '123456789', text: '123456789 · Bart Simpson', sub: 'Homer Simpson and Marge Simpson' }],
  attention: [
    { i: 3, row: 8, label: 'Bart Simpsn', sub: 'Bus Route R12', status: 'guess', reason: 'Close spelling: last name “Simpsn” ≈ “Simpson”', target: '123456789', targetText: { text: '123456789 · Bart Simpson' }, confidence: 95, level: 'high', options: [] },
    { i: 4, row: 9, label: 'Jon Smith', sub: 'Bus Route R7', status: 'guess', reason: 'Close spelling: first name “Jon” ≈ “Jonathan”, last name “Smith” ≈ “Smyth”', target: '456789123', targetText: { text: '456789123 · Jonathan Smyth' }, confidence: 87, level: 'medium',
      options: [opt('345678912', '345678912 · Mark Smith', 'Jane Smith', 68, 'low')] },
    { i: 5, row: 10, label: 'Marc Smith', sub: 'Bus Route R7', status: 'pick', reason: '2 students are about as close to “Marc Smith”', confidence: null, level: '',
      options: [opt('345678912', '345678912 · Mark Smith', 'Jane Smith', 89, 'medium'), opt('345678913', '345678913 · Mary Smith', 'Jane Smith', 89, 'medium')] },
    { i: 6, row: 11, label: 'Lisa', sub: 'Bus Route R3', status: 'guess', reason: 'First name only', target: '123456790', targetText: { text: '123456790 · Lisa Simpson' }, confidence: 70, level: 'low', options: [] },
    { i: 7, row: 12, label: 'Peter Parkr', sub: 'Bus Route R9', status: 'none', reason: 'No close match on the Student List', confidence: null, level: '',
      options: [opt('111222333', '111222333 · Pete Parks', 'Mary Parks', 58, 'weak')] }
  ] };
(async () => {
  const dir = process.argv[2] || require('path').join(__dirname, 'out');
  fs.mkdirSync(dir, { recursive: true });
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 340, height: 900 }, deviceScaleFactor: 1 });
  const shoot = async (file, resp, name, act) => {
    const p = await ctx.newPage();
    const errs = []; p.on('pageerror', e => errs.push(e.message));
    await p.goto('file://' + page(file, runStub(resp)));
    await p.waitForTimeout(600);
    if (act) await act(p);
    await p.screenshot({ path: dir + '/' + name + '.png', fullPage: true });
    console.log(name, errs.length ? 'JS ERRORS: ' + errs.join(' | ') : 'ok');
    await p.close();
  };
  await shoot('InputDataSidebar', { idm_getState: ats, idm_check: ats }, 'idm_ats');
  await shoot('InputDataSidebar', { idm_getState: lists, idm_check: lists }, 'idm_lists');
  await shoot('InputDataSidebar', { idm_getState: { active: false, finished: true, anyImported: true, summary: ['✅ Student Data (ATS / Insight): 120 row(s) imported', '⏭️ Parent Square: skipped (old data kept)', '— NYCSA: not loaded (old data kept)', '✅ Your Lists: 1 list(s) set up'] }, idm_deleteLeftoverTab: { deleted: true } }, 'idm_done');
  await shoot('InputDataSidebar', { idm_getState: Object.assign({}, ats, { debug: ['+0ms lock ok', '+12ms temp tab has data: false'] }), idm_check: ats }, 'idm_debug');
  await shoot('YourListsSidebar', { od_getWizard: setupView }, 'yl_setup');
  await shoot('YourListsSidebar', { od_getWizard: setupView, od_matchPreview: matchView }, 'yl_match', async p => { await p.click('#nextBtn'); await p.waitForTimeout(200); await p.click('#nextBtn'); await p.waitForTimeout(500); });
  await shoot('YourListsSidebar', { od_getWizard: setupView }, 'yl_vars', async p => { await p.click('#nextBtn'); await p.waitForTimeout(300); });
  await browser.close();
})();
