// tests/check-syntax.js
// Runs `node --check` on every shared module and generator.

import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

function collect(dir) {
  const out = [];
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) out.push(...collect(full));
    else if (entry.name.endsWith('.js')) out.push(full);
  }
  return out;
}

const files = [
  ...collect(path.join(ROOT, 'shared')),
  ...collect(path.join(ROOT, 'generators')),
  path.join(ROOT, 'registry.js'),
];

let failed = 0;
for (const file of files) {
  try {
    execFileSync(process.execPath, ['--check', file], { stdio: 'pipe' });
  } catch (err) {
    failed += 1;
    console.error(`SYNTAX ERROR: ${path.relative(ROOT, file)}`);
    console.error(String(err.stderr || err.message));
  }
}

if (failed > 0) {
  console.error(`\n${failed} file(s) with syntax errors`);
  process.exit(1);
}
console.log(`OK: ${files.length} files pass syntax check`);
