// Runs every offline test (node tests/run_all.js, or `npm test` in tests/).
const { spawnSync } = require('child_process');
const path = require('path');
const tests = ['check.js', 'test_matching.js', 'test_fixes.js', 'test_sync.js'];
let failed = 0;
for (const t of tests) {
  const r = spawnSync(process.execPath, [path.join(__dirname, t)], { encoding: 'utf8' });
  const ok = r.status === 0;
  if (!ok) failed++;
  console.log((ok ? 'PASS ' : 'FAIL ') + t);
  if (!ok) console.log((r.stdout || '') + (r.stderr || ''));
}
console.log(failed ? '\n' + failed + ' test file(s) failed' : '\nAll tests passed');
process.exitCode = failed ? 1 : 0;
