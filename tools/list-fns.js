// node tools/list-fns.js
const fs = require('fs');
const globs = ['library-si/src/**/*.js', 'dashboard/src/**/*.js'];
const {globSync} = require('glob');

function extractFunctions(txt) {
  const names = new Set();
  const reFunc = /function\s+([A-Za-z0-9_$]+)\s*\(/g;
  const reAssign = /([A-Za-z0-9_$]+)\s*=\s*function\s*\(/g;
  const reArrow = /const\s+([A-Za-z0-9_$]+)\s*=\s*\([^)]*\)\s*=>/g;
  let m;
  for (const re of [reFunc, reAssign, reArrow]) {
    while ((m = re.exec(txt))) names.add(m[1]);
  }
  return [...names];
}

const out = {};
for (const p of globs.flatMap(g => globSync(g))) {
  const txt = fs.readFileSync(p,'utf8');
  out[p] = extractFunctions(txt).sort();
}
fs.writeFileSync('functions-inventory.json', JSON.stringify(out, null, 2));
console.log('OK -> functions-inventory.json');
