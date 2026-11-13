// Structural check of the whole project (no Google needed):
//   - every .gs file parses, and all of them together compile as ONE script (Apps Script
//     shares one global scope), in both file orders
//   - no top-level name is declared twice across files
//   - every HTML file the code opens exists, and every HTML file is used
//   - every google.script.run call in the HTML and every menu item points at a function
// Usage: node tests/check.js [baselineDir]   (baselineDir: an older copy of the repo,
// to list which top-level declarations changed). Needs: npm install (acorn).
const fs = require('fs'), path = require('path'), vm = require('vm');
let acorn;
try { acorn = require('acorn'); } catch (e) { console.error("Run `npm install` in tests/ first (needs acorn)."); process.exit(2); }
const ROOT = path.resolve(__dirname, '..');
const base = process.argv[2];
let problems = 0;
const report = (label, list) => { console.log(label + ':', list.length ? list : 'none'); problems += list.length; };

function load(d) {
  const files = fs.readdirSync(d).filter(f => f.endsWith('.gs'));
  const decls = new Map(), errors = [];
  for (const f of files) {
    const src = fs.readFileSync(path.join(d, f), 'utf8');
    let ast;
    try { ast = acorn.parse(src, { ecmaVersion: 2022, sourceType: 'script' }); }
    catch (e) { errors.push(f + ': ' + e.message); continue; }
    for (const node of ast.body) {
      let names = [];
      if (node.type === 'FunctionDeclaration') names = [node.id.name];
      else if (node.type === 'VariableDeclaration') names = node.declarations.map(x => x.id.name);
      for (const n of names) {
        if (!decls.has(n)) decls.set(n, []);
        decls.get(n).push({ file: f, kind: node.type === 'VariableDeclaration' ? node.kind : 'function', src: src.slice(node.start, node.end) });
      }
    }
  }
  return { files, decls, errors };
}

const cur = load(ROOT);
console.log('.gs files:', cur.files.length);
report('parse errors', cur.errors);
report('names declared twice', [...cur.decls].filter(([, l]) => l.length > 1).map(([n, l]) => n + ' in ' + l.map(x => x.file).join(', ')));
const all = cur.files.map(f => fs.readFileSync(path.join(ROOT, f), 'utf8'));
const compileErrors = [];
for (const order of [all, all.slice().reverse()]) {
  try { new vm.Script(order.join('\n;\n')); } catch (e) { compileErrors.push(e.message); }
}
report('whole-project compile errors', compileErrors);

const html = fs.readdirSync(ROOT).filter(f => f.endsWith('.html'));
const htmlNames = new Set(html.map(f => f.replace(/\.html$/, '')));
const gsText = all.join('\n');
const refs = new Set();
const refRes = [/createHtmlOutputFromFile\(\s*['"]([^'"]+)['"]/g, /createTemplateFromFile\(\s*['"]([^'"]+)['"]/g, /htmlFromFile_\(\s*['"]([^'"]+)['"]/g, /include_\(\s*['"]([^'"]+)['"]/g];
refRes.forEach(re => { let m; while ((m = re.exec(gsText))) refs.add(m[1]); });
html.forEach(h => { const t = fs.readFileSync(path.join(ROOT, h), 'utf8'); let m; const re = /include_\(\s*['"]([^'"]+)['"]/g; while ((m = re.exec(t))) refs.add(m[1]); });
report('HTML files opened but missing', [...refs].filter(n => !htmlNames.has(n)));
report('HTML files never opened', [...htmlNames].filter(n => !refs.has(n)));

const fnNames = new Set([...cur.decls].filter(([, l]) => l.some(x => x.kind === 'function')).map(([n]) => n));
const notServer = new Set(['getElementById', 'close', 'querySelector', 'focus', 'click', 'remove', 'add', 'toggle', 'setTimeout', 'push']);
const missing = new Set();
html.forEach(h => {
  const t = fs.readFileSync(path.join(ROOT, h), 'utf8');
  (t.match(/google\.script\.run[\s\S]*?;/g) || []).forEach(chain => {
    const calls = [...chain.matchAll(/\.\s*([A-Za-z_$][\w$]*)\s*\(/g)].map(x => x[1])
      .filter(n => !/^with(Success|Failure)Handler$|^withUserObject$/.test(n) && !notServer.has(n));
    const last = calls[calls.length - 1];
    if (last && !fnNames.has(last)) missing.add(h + ': ' + last);
  });
  for (const m of t.matchAll(/['"]((?:idm|od|zg|cld|cl|ma)_[A-Za-z]+)['"]/g)) if (!fnNames.has(m[1])) missing.add(h + ' (by name): ' + m[1]);
});
report('google.script.run calls with no function', [...missing]);
const menuMissing = [];
for (const m of gsText.matchAll(/addItem\(\s*(?:"[^"]*"|'[^']*'|[^,]+?)\s*,\s*["']([\w$]+)["']\s*\)/g)) if (!fnNames.has(m[1])) menuMissing.push(m[1]);
report('menu items with no function', menuMissing);

if (base) {
  const old = load(base);
  console.log('names removed:', [...old.decls.keys()].filter(n => !cur.decls.has(n)));
  console.log('names added:', [...cur.decls.keys()].filter(n => !old.decls.has(n)));
  const changed = [];
  for (const [n, l] of cur.decls) {
    const o = old.decls.get(n);
    if (o && o.map(x => x.src).join('\n') !== l.map(x => x.src).join('\n')) changed.push(n + ' (' + l.map(x => x.file).join(',') + ')');
  }
  console.log('declarations whose code changed:', changed);
}
console.log(problems ? '\nCHECK FAILED: ' + problems + ' problem(s)' : '\nCHECK OK');
process.exitCode = problems ? 1 : 0;
