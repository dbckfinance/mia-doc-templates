// shared/cli.js
// Standard CLI wrapper for every generator:
//   node generators/<vertical>/<id>.js input.json          → ./artifacts/<name>.<ext>
//   cat input.json | node generators/<vertical>/<id>.js    → ./artifacts/<name>.<ext>
//
// The Cursor SDK only detects files inside ./artifacts, so all outputs land there.

import fs from 'node:fs';
import path from 'node:path';

export function readStdinSync() {
  try {
    return fs.readFileSync(0, 'utf8');
  } catch {
    return '';
  }
}

/** Read generator input from argv[2] (JSON file path) or stdin. */
export function readInput() {
  const argPath = process.argv[2];
  let raw = '';
  if (argPath && fs.existsSync(argPath)) {
    raw = fs.readFileSync(argPath, 'utf8');
  } else {
    raw = readStdinSync();
  }
  if (!raw.trim()) return {};
  try {
    return JSON.parse(raw);
  } catch (err) {
    console.error(`ERROR: invalid JSON input — ${err.message}`);
    process.exit(1);
  }
}

export function ensureArtifactsDir() {
  const dir = path.resolve('./artifacts');
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  return dir;
}

export function safeFileName(name, fallback = 'mia-document') {
  const slug = String(name || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9_-]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 80);
  return slug || fallback;
}

/** Validate Office Open XML output is a real ZIP container. */
export function assertValidOffice(buffer, ext) {
  if (['docx', 'xlsx', 'pptx'].includes(ext)) {
    if (!buffer || buffer.length < 4 || buffer[0] !== 0x50 || buffer[1] !== 0x4b) {
      throw new Error(`Generated .${ext} is not a valid Office document (not a ZIP container)`);
    }
  }
  if (ext === 'pdf') {
    const head = buffer.slice(0, 5).toString('latin1');
    if (!head.startsWith('%PDF')) {
      throw new Error('Generated .pdf does not start with %PDF header');
    }
  }
}

/**
 * Standard generator entrypoint. Call from each generator:
 *
 *   if (isCliInvocation(import.meta)) await runCli({ metadata, generate });
 *
 * `generate(input)` must return:
 *   Buffer                                      (single file, metadata.outputType ext)
 *   or [{ buffer, ext, suffix? }, ...]          (multi-format generators)
 */
export async function runCli({ metadata, generate }) {
  const input = readInput();
  const dir = ensureArtifactsDir();
  const base = safeFileName(input.fileName || metadata.id);

  console.log(`[${metadata.id}] generating "${base}" (${metadata.outputType})…`);
  const out = await generate(input);
  const files = Buffer.isBuffer(out)
    ? [{ buffer: out, ext: metadata.outputType }]
    : out;

  for (const f of files) {
    assertValidOffice(f.buffer, f.ext);
    const name = `${base}${f.suffix ? `-${f.suffix}` : ''}.${f.ext}`;
    const target = path.join(dir, name);
    fs.writeFileSync(target, f.buffer);
    console.log(`OK: wrote ${target} (${f.buffer.length} bytes)`);
  }
}

/** True when the module is being executed directly (not imported). */
export function isCliInvocation(meta) {
  if (!process.argv[1]) return false;
  try {
    const invoked = path.resolve(process.argv[1]).toLowerCase();
    const selfPath = decodeURIComponent(new URL(meta.url).pathname)
      .replace(/^\/([A-Za-z]:)/, '$1');
    const self = path.resolve(selfPath).toLowerCase();
    return invoked === self;
  } catch {
    return false;
  }
}
