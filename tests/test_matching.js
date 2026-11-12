// Your Lists matching (Input Data.gs part 2): who each list row is, with a confidence.
const assert = require('assert');
const { loadProject } = require('./lib/load');
const g = loadProject();

// Student List rows: OSIS, Student, Guardian, Site, Class
const students = [
  ["123456789", "Bart Simpson", "Homer Simpson and Marge Simpson", "Site 1", "Y25"],
  ["123456790", "Lisa Simpson", "Homer Simpson and Marge Simpson", "Site 1", "Y26"],
  ["234567891", "Wednesday Addams", "Morticia Addams", "Site 2", "Y41"],
  ["234567892", "Pugsley Addams", "Morticia Addams", "Site 2", "Y41"],
  ["345678912", "Mark Smith", "Jane Smith", "Site 3", "Y12"],
  ["345678913", "Mary Smith", "Jane Smith", "Site 3", "Y12"],
  ["456789123", "Jonathan Smyth", "Paul Smyth", "Site 3", "Y50"],
  ["567891234", "Maria Del Carmen Lopez", "Rosa Lopez", "Site 4", "Y60"],
  ["678912345", "Samuel Johnson", "Linda Johnson", "Site 4", "Y61"],
  ["789123456", "Diana Prince", "Hippolyta Prince", "Site 5", "Y70"]
];
const guardians = [
  { student_id: "123456789", guardianOutput: "Homer Simpson (Father)", mobile: "(212) 555-0101", email: "homer@x.com" },
  { student_id: "123456790", guardianOutput: "Homer Simpson (Father)", mobile: "(212) 555-0101", email: "homer@x.com" },
  { student_id: "123456789", guardianOutput: "Marge Simpson (Mother)", mobile: "(212) 555-0102", email: "" },
  { student_id: "234567891", guardianOutput: "Morticia Addams (Mother)", mobile: "(212) 555-0201", email: "" },
  { student_id: "345678912", guardianOutput: "Jane Smith (Mother)", mobile: "(212) 555-0301", email: "" },
  { student_id: "678912345", guardianOutput: "Linda Johnson (Mother)", mobile: "(212) 555-0401", email: "" }
];
const idx = g.od_index_(students, guardians);
const match = (by, roles, row) => {
  const r = g.od_matchAll_(g.od_fillSetup_({ by: by, roles: roles, picks: {} }), [row], idx)[0];
  return { status: r.status, who: r.target ? g.od_targetText_(by, r.target, idx).text : "", conf: r.confidence, level: g.od_confLevel_(r.confidence), n: (r.options || []).length };
};
const expect = (by, roles, row, status, whoPart, level) => {
  const r = match(by, roles, row);
  assert.strictEqual(r.status, status, JSON.stringify(row) + " -> " + JSON.stringify(r));
  if (whoPart) assert(r.who.indexOf(whoPart) !== -1, JSON.stringify(row) + " -> " + JSON.stringify(r));
  if (level) assert.strictEqual(r.level, level, JSON.stringify(row) + " -> " + JSON.stringify(r));
  return r;
};
const S = { student: { col: 1 } };
expect("student", S, ["Bart Simpson"], "ok", "Bart Simpson");
expect("student", S, ["SIMPSON BART"], "ok", "Bart Simpson");
expect("student", S, ["Simpson, Bart"], "ok", "Bart Simpson");
expect("student", S, ["Bart Simpsn"], "guess", "Bart Simpson", "high");          // typo
expect("student", S, ["Bort Simpson"], "guess", "Bart Simpson", "medium");       // typo in the first name
expect("student", S, ["Marc Smith"], "pick");                                    // Mark or Mary: too close to call
expect("student", S, ["Mary Smith"], "ok", "Mary Smith");
expect("student", S, ["Jon Smith"], "guess", "Jonathan Smyth", "medium");
expect("student", S, ["Maria Lopez"], "guess", "Maria Del Carmen Lopez", "high"); // part of a longer name
expect("student", S, ["Sam Johnson"], "guess", "Samuel Johnson", "high");        // short name
expect("student", S, ["Diana P."], "guess", "Diana Prince", "medium");           // last initial
expect("student", S, ["D Prince"], "guess", "Diana Prince", "medium");           // first initial
expect("student", S, ["Pugsly Adams"], "guess", "Pugsley Addams", "medium");
expect("student", S, ["Lisa"], "guess", "Lisa Simpson", "low");                  // first name only
expect("student", S, ["Peter Parker"], "none");
const SG = { student: { col: 1 }, guardian: { col: 2 } };
expect("student", SG, ["Mark Smth", "Jane Smith"], "guess", "Mark Smith", "high"); // parent matches
expect("student", SG, ["", "Morticia Adams"], "pick");                            // a parent of two
const G = { guardian: { col: 1 }, student: { col: 2 } };
expect("guardian", G, ["Homer Simpson", ""], "ok", "Homer Simpson");
expect("guardian", G, ["Homer Simpsen", ""], "guess", "Homer Simpson", "high");
expect("guardian", G, ["Marj Simpson", ""], "guess", "Marge Simpson", "medium");
expect("guardian", G, ["", "Wednesday Addams"], "guess", "Morticia Addams", "medium"); // the child's name
expect("guardian", G, ["Tom Jones", ""], "none");
console.log("matching: 22 rows as expected");

// Speed: 500 rows that all need a fuzzy look against 1,500 students
const firsts = "james mary john patricia robert jennifer michael linda william elizabeth david barbara richard susan joseph jessica thomas sarah charles karen maria jose luis ana carlos sofia diego valentina mateo camila".split(" ");
const lasts = "smith johnson williams brown jones garcia miller davis rodriguez martinez hernandez lopez gonzalez wilson anderson thomas taylor moore jackson martin lee perez thompson white harris sanchez clark ramirez lewis robinson walker young allen king wright scott torres nguyen hill flores green adams nelson baker hall rivera campbell mitchell carter roberts".split(" ");
const cap = s => s[0].toUpperCase() + s.slice(1);
let seed = 7; const rnd = n => { seed = (seed * 16807) % 2147483647; return seed % n; };
const many = [];
for (let i = 0; i < 1500; i++) many.push([String(200000000 + i), cap(firsts[rnd(firsts.length)]) + " " + cap(lasts[rnd(lasts.length)]), "", "", "Y1"]);
const big = g.od_index_(many, []);
const typo = s => { const i = 1 + rnd(s.length - 2); return s.slice(0, i) + s.slice(i + 1); };
const rows = [];
for (let i = 0; i < 500; i++) { const st = many[rnd(1500)][1].split(" "); rows.push([typo(st[0]) + " " + typo(st[1])]); }
const t = Date.now();
g.od_matchAll_(g.od_fillSetup_({ by: "student", roles: S, picks: {} }), rows, big);
const ms = Date.now() - t;
console.log("matching speed: 500 fuzzy rows x 1,500 students in " + ms + " ms");
assert(ms < 10000, "matching got much slower: " + ms + " ms");
