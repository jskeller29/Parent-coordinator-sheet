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
      // idm_open (the Input Data Mode sidebar's first call) answers like idm_getState unless given
      if (fn === 'idm_open' && window.RESP.idm_open === undefined) fn = 'idm_getState';
      return function () { const r = window.RESP[fn]; if (r === '__HANG__') return; setTimeout(() => ok && ok(typeof r === 'function' ? r.apply(null, arguments) : JSON.parse(JSON.stringify(r === undefined ? null : r))), 30); };
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
  // 2026-11-12: a wrong paste cleared, step 5 building first / built, the end screen after it, the list upload
  const cleared = Object.assign({}, ats, { notice: { type: 'error', code: 'wrongType', cleared: true, fromUpload: false,
    message: 'This looks like <b>Parent Square</b> data, but this step is <b>Student Data (ATS / Insight)</b>. Use the Student Data (ATS / Insight) export instead (or Skip to the Parent Square step).' } });
  await shoot('InputDataSidebar', { idm_getState: cleared, idm_check: ats }, 'idm_cleared');
  const buildSteps = [{ label: 'Reading your data' }, { label: 'Student List' }, { label: 'Event Log' }, { label: 'Combined Contact Tracking' }, { label: 'Parents Divided' }, { label: 'Directory' }, { label: 'Phone Contacts' }, { label: 'Finishing up' }];
  await shoot('InputDataSidebar', { idm_getState: Object.assign({}, lists, { needsBuild: true }), idm_buildForLists: '__HANG__', getSyncStatus: { phase: 'idle', buildSteps: buildSteps, buildStep: 1 } },
    'idm_lists_building', async p => { await p.waitForTimeout(3000); });
  await shoot('InputDataSidebar', { idm_getState: Object.assign({}, lists, { listsBuild: { ok: true, message: 'Sheets built, so your lists can find your students.' } }) }, 'idm_lists_built');
  await shoot('InputDataSidebar', { idm_getState: { active: false, finished: true, anyImported: false, builtAlready: true, guideNext: true, summary: ['✅ Student Data (ATS / Insight): 120 row(s) imported', '— Your Lists: skipped'] }, idm_deleteLeftoverTab: { deleted: true } }, 'idm_done_built');
  // "Finish or exit" (2026-11-12): the exit panel with "Save, exit and build my sheets"
  const ps = Object.assign({}, ats, { stepIndex: 1, canBuild: true, step: Object.assign({}, ats.step, { key: 'ps', zone: 'Parent Square', required: false }), progress: progress(1, {}),
    done: [{ name: 'Student Data', rows: 120, skipped: false }] });
  await shoot('InputDataSidebar', { idm_getState: ps, idm_check: ps }, 'idm_exit_build', async p => { await p.click('#exitBtn'); await p.waitForTimeout(150); });
  await shoot('InputDataSidebar', { idm_getState: Object.assign({}, ps, { canBuild: false }), idm_check: Object.assign({}, ps, { canBuild: false }) }, 'idm_exit_nobuild', async p => { await p.click('#exitBtn'); await p.waitForTimeout(150); });
  // 2026-11-13: "Student Data (ATS) completed!" at the top of the next step; a list's tab icon
  const psDone = Object.assign({}, ps, { notice: { type: 'imported', message: '✅ Student Data (ATS / Insight): 523 row(s) moved to Raw Data.',
    done: { kind: 'ok', title: 'Student Data (ATS) completed!', sub: '523 rows moved to Raw Data. Next: Parent Square.' } } });
  await shoot('InputDataSidebar', { idm_getState: psDone, idm_check: ps }, 'idm_done_banner', async p => { await p.waitForTimeout(2600); }); // still there after a poll
  const listsIcons = Object.assign({}, lists, { lists: [Object.assign({}, lists.lists[0], { icon: '🚌', tabName: '🚌 Bus Routes' }), lists.lists[1]] });
  await shoot('InputDataSidebar', { idm_getState: listsIcons }, 'idm_lists_icons', async p => {
    await p.click('#listItems .li button.ico'); await p.click('#addListBtn'); await p.fill('#newListName', 'Field Trip');
    await p.click('#newListIcons button[data-icon="🎉"]'); await p.waitForTimeout(150); });
  await shoot('YourListsSidebar', { od_getWizard: { phase: 'paste', list: { tabId: 1, name: 'Bus Routes', tabName: '📥 Bus Routes', ready: false } }, od_wizardCheck: { pasted: false } }, 'yl_paste');
  await shoot('YourListsSidebar', { od_getWizard: setupView }, 'yl_setup');
  await shoot('YourListsSidebar', { od_getWizard: setupView, od_matchPreview: matchView }, 'yl_match', async p => { await p.click('#nextBtn'); await p.waitForTimeout(200); await p.click('#nextBtn'); await p.waitForTimeout(500); });
  await shoot('YourListsSidebar', { od_getWizard: setupView }, 'yl_vars', async p => { await p.click('#nextBtn'); await p.waitForTimeout(300); });
  // Step 4 View (2026-11-12): simple, then one line with Advanced options open
  const viewSetup = Object.assign({}, setupView, {
    columns: setupView.columns.concat([{ col: 4, letter: 'D', header: 'Stop', samples: ['Main St & 5th'], hint: { osisKnown: 0, n: 40, comma: false } }]),
    vars: setupView.vars.concat([{ col: 4, header: 'Stop', label: 'Stop', show: true, column: true, blank: { mode: 'text', text: 'No stop yet' }, notOnList: { mode: 'nothing', text: '' } }]),
    view: { layout: 'lines', boldNames: true, boldValues: false, italic: false, sep: 'comma', namesBefore: false, start: 'list', startText: '', columns: 'each', order: [] } });
  const viewMatch = Object.assign({}, matchView, { preview: [
    { name: 'Bart Simpson', row: ['Bart Simpson', '123456789', 'R12', 'Main St & 5th'] },
    { name: 'Lisa Simpson', row: ['Lisa Simpson', '567891234', 'R12', ''] }] });
  const toView = async p => { for (let i = 0; i < 3; i++) { await p.click('#nextBtn'); await p.waitForTimeout(300); } };
  await shoot('YourListsSidebar', { od_getWizard: viewSetup, od_matchPreview: viewMatch }, 'yl_view', toView);
  await shoot('YourListsSidebar', { od_getWizard: Object.assign({}, viewSetup, { view: Object.assign({}, viewSetup.view, { layout: 'oneLine', sep: 'dot', columns: 'one' }) }), od_matchPreview: viewMatch },
    'yl_view_advanced', async p => { await toView(p); await p.click('.pvnav button:last-child'); await p.waitForTimeout(150); });
  // 2026-11-18 (release check round 2 notes + ideas): pick what to update, Site / Class Save above Skip,
  // a problem at the top, the end screen's Build / Done + parents to add, the Student Overrides main page
  const startItems = ['Student Data (ATS)', 'Parent Square', 'NYCSA', 'Site / Class', 'Your Lists'].map((n, i) => ({ index: i, key: 'k' + i, name: n, required: i === 0,
    uses: 'What it is for.', where: i < 3 ? 'TeachHub → …' : '', fileHint: i === 0 ? 'its name starts with <b>general_</b>' : '',
    last: i === 3 ? null : { text: '10/0' + (i + 1) + '/26', days: 40 - i * 10 }, due: i < 2, picked: i < 2 }));
  const start = { active: true, overview: true, stepIndex: 0, stepCount: 5, start: startItems, step: ats.step, tabName: 'Start Here', progress: progress(-1, {}), notice: null };
  await shoot('InputDataSidebar', { idm_getState: start, idm_check: start }, 'idm_start_pick', async p => { await p.click('#startItems .si:nth-child(3)'); await p.waitForTimeout(150); });
  const site = Object.assign({}, ats, { stepIndex: 3, step: { key: 'siteclass', zone: 'Site / Class', required: false, manual: true, lists: false, howTo: ['Type each class\'s <b>Teacher Name</b>.'] },
    tabName: 'Site / Class', next: { index: 4, short: 'Your Lists' }, siteClass: { classes: 12, teachers: 8, sites: 5 },
    progress: progress(3, {}).map((x, i) => Object.assign(x, { picked: i !== 1 })) });
  await shoot('InputDataSidebar', { idm_getState: site, idm_check: site }, 'idm_site_footer');
  const wrong = Object.assign({}, cleared, { stepIndex: 1, step: Object.assign({}, ats.step, { key: 'ps', zone: 'Parent Square', required: false }), progress: progress(1, {}) });
  await shoot('InputDataSidebar', { idm_getState: wrong, idm_check: ps }, 'idm_error_top');
  await shoot('InputDataSidebar', { idm_getState: listsIcons }, 'idm_icon_hover', async p => { await p.hover('#listItems .li button.ico'); await p.waitForTimeout(250); });
  await shoot('InputDataSidebar', { idm_getState: { active: false, finished: true, anyImported: true, guideNext: false, summary: ['✅ Student Data (ATS / Insight): 120 row(s) imported', '— Parent Square: not loaded (old data kept)'] },
    idm_deleteLeftoverTab: { deleted: true }, idm_parentsToAddSummary: { count: 4, sample: ['Abe Simpson (Bart Simpson)', 'Patty Bouvier (Bart Simpson)', 'Grandmama Addams (Wednesday Addams)'] } }, 'idm_done_parents', async p => { await p.waitForTimeout(400); });
  const home = { toAdd: [
      { key: 'a', name: 'Grandmama Addams', relationship: 'Grandmother', student: 'Wednesday Addams', cls: 'Y12', osis: '111222333', phone: '(777) 777-7777', email: '', sources: ['NYCSA'] },
      { key: 'b', name: 'Abe Simpson', relationship: '', student: 'Bart Simpson', cls: 'Y41', osis: '123456789', phone: '(333) 333-3333', email: 'abe@x.org', sources: ['Parent Square', '👪 PTA'] },
      { key: 'c', name: 'Patty Bouvier', relationship: 'Aunt', student: 'Bart Simpson', cls: 'Y41', osis: '123456789', phone: '(444) 444-4444', email: '', sources: ['NYCSA'] }],
    skipped: 2, paused: false, error: '',
    changes: [{ date: '11/18/2026', type: 'Added Guardian (NYCSA)', kind: 'added', source: 'NYCSA', osis: '567891234', lastName: 'Simpson', firstName: 'Lisa', classCode: 'Y25', guardians: ['Selma Bouvier'] },
              { date: '11/02/2026', type: 'New Student', kind: 'new', source: '', osis: '998877665', lastName: 'Belcher', firstName: 'Tina', classCode: 'Y42', guardians: ['Bob Belcher'] },
              { date: '10/28/2026', type: 'Edit Existing Student', kind: 'edit', source: '', osis: '678912345', lastName: 'Parker', firstName: 'Peter', classCode: 'Y11', guardians: [], hidden: true }] };
  const dlg = await browser.newContext({ viewport: { width: 640, height: 680 }, deviceScaleFactor: 1 });
  const shootDlg = async (file, resp, name, act) => {
    const p = await dlg.newPage();
    const errs = []; p.on('pageerror', e => errs.push(e.message));
    await p.goto('file://' + page(file, runStub(resp)));
    await p.waitForTimeout(500);
    if (act) await act(p);
    await p.screenshot({ path: dir + '/' + name + '.png' });
    console.log(name, errs.length ? 'JS ERRORS: ' + errs.join(' | ') : 'ok');
    await p.close();
  };
  await shootDlg('ManageOverridesDialog', { getOverridesHome: home }, 'ov_home', async p => { await p.click('#addList .card:nth-child(2) input'); await p.waitForTimeout(100); });
  await shootDlg('ManageOverridesDialog', { getOverridesHome: home }, 'ov_home_changes', async p => { await p.evaluate(() => document.querySelector('.ov-body').scrollTo(0, 9999)); await p.waitForTimeout(100); });
  await shootDlg('ManageOverridesDialog', { getOverridesHome: Object.assign({}, home, { toAdd: [], skipped: 0 }) }, 'ov_home_empty');
  await shootDlg('NewStudentDialog', { getAllStudentsForEdit: [], getNewStudentPrefill: { osis: '', guardian: null, fromHome: true } }, 'ov_newstudent_back');
  await browser.close();
})();
