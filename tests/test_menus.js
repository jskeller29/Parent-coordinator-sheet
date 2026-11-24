// Menu layout "2+" (owner 2026-11-24: "the App Menu is clogged"; Menus.gs MENUS, buildMenus_).
//   🚀 App Menu: ✅ Getting Started (until its checklist is finished), 📥 Input Data Mode, 🏗️ Build
//     Sheets, 🧱 Build + Sync Contacts (phone sync on), ───, ⚙️ Settings, ❓ Help ▸, 🧰 More ▸
//   👪 Families: 🔍 Find a Family, ✨ New / Edit Student, 🗂️ Family Changes
//   📆 Events: 📋 New Event, 👁️ View Events
// Only labels moved: every item runs the same function as before.
const assert = require('assert');
const { loadProject } = require('./lib/load');
const { makeSpreadsheet, fakeSpreadsheetApp } = require('./lib/fakesheet');

const TABS = ["Contact Log", "Combined Contact Tracking", "PCAR", "Directory", "Raw Data", "Student List", "Notes", "Backend"];

/** Opens a sheet in a given state and returns the menus that were added to the menu bar. */
function openWith(o) {
  o = o || {};
  const ss = makeSpreadsheet(TABS);
  const app = fakeSpreadsheetApp(ss);
  const g = loadProject({ extra: { SpreadsheetApp: app }, owner: !!o.owner, debugFiles: !!o.debugFiles, noTrialFiles: !o.trial });
  g.contactLogOpenRepair_ = () => {};
  g.welcomeTabAfterSetup_ = () => {};
  g.isAutoMigrateOn_ = () => false;
  g.showWelcomeTabOnly_ = () => {};
  if (typeof g.ownerNewCopyCheck_ === 'function') g.ownerNewCopyCheck_ = () => {};
  g.studentListIsBuilt_ = () => !!o.built;
  const dp = g.PropertiesService.getDocumentProperties();
  dp.setProperty('ORIGINAL_SHEET_ID', 'FAKE_SS');
  if (o.setup !== false) dp.setProperty('SETUP_COMPLETE', 'true');
  Object.keys(o.props || {}).forEach(k => dp.setProperty(k, o.props[k]));
  g.onOpen({ source: ss });
  const bar = app._ui._menus.filter(m => m.inUi);
  return { g, app, ss, bar, menu: name => bar.find(m => m.name === name) };
}
/** What a menu shows: labels, "---" for a separator, "<name> ▸ [ ... ]" for a submenu. */
const show = m => m.entries.map(e => e.sep ? '---' : e.sub ? { [e.sub.name]: show(e.sub) } : e.label);
const fns = m => m.entries.reduce((a, e) => a.concat(e.sub ? fns(e.sub) : e.fn ? [e.fn] : []), []);
const eq = (a, b, msg) => assert.deepStrictEqual(JSON.parse(JSON.stringify(a)), JSON.parse(JSON.stringify(b)), msg);

const HELP = ["📖 User Guide", "🐞 Report a Bug", "🔔 Check for Updates"];
const MORE = ["🔓 Unlock Layout", "🎓 End of Year Rollover"];

// ---------------------------------------------------------------- before Initial Setup: unchanged
{
  const { bar } = openWith({ setup: false });
  eq(bar.map(m => m.name), ["🚀 App Menu"], "one menu");
  eq(show(bar[0]), ["🚨 Initial Setup (Run Once)"], "only Initial Setup");
  console.log("before setup: ok");
}

// ---------------------------------------------------------------- set up, not built yet
{
  const { bar, menu } = openWith({ built: false });
  eq(bar.map(m => m.name), ["🚀 App Menu"], "only the 🚀 App Menu before the first build");
  eq(show(menu("🚀 App Menu")), [
    "✅ Getting Started", "📥 Input Data Mode", "🏗️ Build Sheets", "---",
    "📦 Bring Over Last Year's Tracker", { "❓ Help": HELP }
  ], "the getting-started App Menu");
  console.log("not built: ok");
}

// ---------------------------------------------------------------- built: the full menus
{
  const { bar, menu } = openWith({ built: true });
  eq(bar.map(m => m.name), ["🚀 App Menu", "👪 Families", "📆 Events"], "three menus in the bar (no Z# file)");
  eq(show(menu("🚀 App Menu")), [
    "✅ Getting Started", "📥 Input Data Mode", "🏗️ Build Sheets", "---",
    "⚙️ Settings", { "❓ Help": HELP }, { "🧰 More": MORE }
  ], "App Menu: checklist not finished, phone sync off");
  eq(show(menu("👪 Families")), ["🔍 Find a Family", "✨ New / Edit Student", "🗂️ Family Changes"]);
  eq(show(menu("📆 Events")), ["📋 New Event", "👁️ View Events"]);
  // Only the labels moved: the same functions as the old menus (+ nothing lost)
  eq(fns(menu("🚀 App Menu")).concat(fns(menu("👪 Families")), fns(menu("📆 Events"))).sort(), [
    "openGuideChecklist", "openInputDataMode", "runBuildOnly", "openSettingsDialog",
    "showWelcomeDialog", "reportABug", "manualCheckForUpdates", "toggleSheetLock", "openRolloverWizard",
    "openFindFamily", "openNewStudentDialog", "openManageOverridesDialog",
    "createNewEventTable", "openViewEventsDialog"
  ].sort(), "every old item is still there, running the same function");
  console.log("built, full menus: ok");
}

// ---------------------------------------------------------------- phone sync on, layout unlocked
{
  const { menu } = openWith({ built: true, props: { syncToPhones: 'true', SHEET_LOCK_OFF: 'true' } });
  const app = show(menu("🚀 App Menu"));
  eq(app.slice(0, 4), ["✅ Getting Started", "📥 Input Data Mode", "🏗️ Build Sheets", "🧱 Build + Sync Contacts"], "Build + Sync right under Build");
  eq(app[app.length - 1], { "🧰 More": ["🔒 Lock Layout", "🎓 End of Year Rollover"] }, "unlocked: Lock Layout");
  console.log("sync on / unlocked: ok");
}

// ---------------------------------------------------------------- the checklist finished: ✅ Getting Started moves into ❓ Help
{
  [{ ZG_AUTO_OPEN: 'off' }, { ZG_GUIDE_DONE: 'true' }].forEach(props => {
    const { menu } = openWith({ built: true, props });
    eq(show(menu("🚀 App Menu")), [
      "📥 Input Data Mode", "🏗️ Build Sheets", "---",
      "⚙️ Settings", { "❓ Help": ["✅ Getting Started"].concat(HELP) }, { "🧰 More": MORE }
    ], "finished (" + Object.keys(props)[0] + "): in ❓ Help");
  });
  // Before the first build it stays on top even when finished (the new-user menu is short anyway)
  const { menu } = openWith({ built: false, props: { ZG_AUTO_OPEN: 'off' } });
  eq(show(menu("🚀 App Menu"))[0], "✅ Getting Started", "not built: still on top");

  // All four steps done: remembered once (true the first time); guideOnOpen_ remembers it too
  const { g } = openWith({ built: true });
  const done = { allDone: true };
  assert.strictEqual(g.zg_rememberDone_({ allDone: false }), false, "not all done: nothing");
  assert.strictEqual(g.__props.ZG_GUIDE_DONE, undefined);
  assert.strictEqual(g.zg_rememberDone_(done), true, "first time");
  assert.strictEqual(g.__props.ZG_GUIDE_DONE, 'true');
  assert.strictEqual(g.zg_rememberDone_(done), false, "then nothing new");
  delete g.__props.ZG_GUIDE_DONE;
  let opened = 0;
  g.openGuideChecklist = () => { opened++; };
  g.zg_checklist_ = () => ({ allDone: true, steps: [] });
  g.guideOnOpen_();
  assert.strictEqual(g.__props.ZG_GUIDE_DONE, 'true', "guideOnOpen_: all done is remembered");
  assert.strictEqual(opened, 0, "and the checklist doesn't open");
  console.log("Getting Started into ❓ Help once finished: ok");
}

// ---------------------------------------------------------------- "Finish setup": the alert names the right place, menus redrawn
{
  const { g, app } = openWith({ built: true });
  let redraws = 0;
  g.onOpen = () => { redraws++; };
  g.zg_finishSetup();
  assert.strictEqual(g.__props.ZG_AUTO_OPEN, 'off');
  assert(/🚀 App Menu → ❓ Help → ✅ Getting Started/.test(app._ui._alerts.pop()), "built: Help path");
  assert.strictEqual(redraws, 1, "built: the menus are redrawn (it moves now)");
  const b = openWith({ built: false });
  let r2 = 0;
  b.g.onOpen = () => { r2++; };
  b.g.zg_finishSetup();
  const alert = b.app._ui._alerts.pop();
  assert(/You can always open it again: 🚀 App Menu → ✅ Getting Started/.test(alert), "not built: on top");
  assert.strictEqual(r2, 0, "not built: no redraw (it stays on top)");
  console.log("Finish setup: ok");
}

// ---------------------------------------------------------------- a copy forgets "finished"
{
  const ss = makeSpreadsheet(TABS);
  const app = fakeSpreadsheetApp(ss);
  const g = loadProject({ extra: { SpreadsheetApp: app }, noTrialFiles: true });
  g.isAutoMigrateOn_ = () => false; g.showWelcomeTabOnly_ = () => {};
  const dp = g.PropertiesService.getDocumentProperties();
  dp.setProperty('ORIGINAL_SHEET_ID', 'ANOTHER_SHEET');
  dp.setProperty('SETUP_COMPLETE', 'true');
  dp.setProperty('ZG_GUIDE_DONE', 'true');
  g.onOpen({ source: ss });
  assert.strictEqual(g.__props.ZG_GUIDE_DONE, undefined, "the clone check deletes ZG_GUIDE_DONE");
  console.log("copy: ok");
}

// ---------------------------------------------------------------- owner test files
{
  // The owner's sheet: 🐞 Contact Log Debug is in the 🐞 Debug menu, not in the App Menu
  const own = openWith({ built: true, owner: true, debugFiles: true });
  assert(fns(own.menu("🐞 Debug")).indexOf("openContactLogDebug") !== -1, "in 🐞 Debug");
  assert.strictEqual(fns(own.menu("🚀 App Menu")).indexOf("openContactLogDebug"), -1, "not in the App Menu");
  // A test copy with the debug file but no z_MINE.gs: at the end of 🧰 More (❓ Help before the first build)
  const copy = openWith({ built: true, debugFiles: true });
  const more = show(copy.menu("🚀 App Menu")).pop()["🧰 More"];
  eq(more[more.length - 1], "🐞 Contact Log Debug", "no Debug menu: end of 🧰 More");
  const copy2 = openWith({ built: false, debugFiles: true });
  const help = show(copy2.menu("🚀 App Menu")).pop()["❓ Help"];
  eq(help[help.length - 1], "🐞 Contact Log Debug", "not built: end of ❓ Help");
  // A school's sheet (no z_ files): nowhere
  const school = openWith({ built: true });
  assert(school.bar.every(m => fns(m).indexOf("openContactLogDebug") === -1), "a school: no debug item");
  console.log("owner test files: ok");
}

// ---------------------------------------------------------------- naming rules (every menu a school sees)
{
  const { bar } = openWith({ built: true, trial: true, props: { syncToPhones: 'true' } });
  eq(bar.map(m => m.name), ["🚀 App Menu", "👪 Families", "📆 Events", "📇 Contact Log"], "with the optional Contact Log tool");
  const labels = [];
  const walk = m => m.entries.forEach(e => { if (e.sub) { labels.push(e.sub.name); walk(e.sub); } else if (e.label) labels.push(e.label); });
  bar.forEach(walk);
  labels.forEach(l => {
    assert(!/[()]/.test(l), "no explanation in brackets: " + l);
    assert(l.length <= 32, "short: " + l);
  });
  // One emoji = one meaning (the same item may show in two places, e.g. ✅ Getting Started)
  const firstEmoji = l => Array.from(l)[0];
  const seen = {};
  bar.map(m => m.name).concat(labels).forEach(l => {
    const e = firstEmoji(l);
    assert(!seen[e] || seen[e] === l, "emoji " + e + " used for both '" + seen[e] + "' and '" + l + "'");
    seen[e] = l;
  });
  // No custom menu called Help in the bar (ours sit right after Google's own Help)
  assert(bar.every(m => !/help/i.test(m.name)), "no Help menu in the bar");
  console.log("naming rules: ok");
}
