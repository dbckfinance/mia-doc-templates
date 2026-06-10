// tests/validate-outputs.js
// Runs every generator with its fixture (or {}) and validates the output buffers
// are real Office/PDF files (magic bytes + minimum size).

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { loadRegistry } from '../registry.js';
import { assertValidOffice } from '../shared/cli.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const FIXTURES_DIR = path.join(ROOT, 'tests', 'fixtures');

function fixtureFor(id) {
  const fixturePath = path.join(FIXTURES_DIR, `${id}.json`);
  if (fs.existsSync(fixturePath)) {
    return JSON.parse(fs.readFileSync(fixturePath, 'utf8'));
  }
  return {};
}

const registry = await loadRegistry();
let passed = 0;
const failures = [];

for (const entry of registry) {
  const mod = await import(new URL(`../${entry.script}`, import.meta.url).href);
  try {
    const input = fixtureFor(entry.id);
    const out = await mod.generate(input);
    const files = Buffer.isBuffer(out) ? [{ buffer: out, ext: entry.outputType }] : out;

    if (!Array.isArray(files) || files.length === 0) {
      throw new Error('generate() returned no output');
    }
    for (const f of files) {
      if (!Buffer.isBuffer(f.buffer)) throw new Error('output is not a Buffer');
      if (f.buffer.length < 1000) throw new Error(`output suspiciously small (${f.buffer.length} bytes)`);
      assertValidOffice(f.buffer, f.ext);
    }
    passed += 1;
    console.log(`PASS  ${entry.id} (${files.map((f) => `${f.ext}:${f.buffer.length}b`).join(', ')})`);
  } catch (err) {
    failures.push({ id: entry.id, error: err.message });
    console.error(`FAIL  ${entry.id} — ${err.message}`);
  }
}

console.log(`\n${passed}/${registry.length} generators produce valid output`);
if (failures.length > 0) {
  console.error('\nFailures:');
  for (const f of failures) console.error(`  ${f.id}: ${f.error}`);
  process.exit(1);
}
