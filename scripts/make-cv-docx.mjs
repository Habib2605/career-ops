#!/usr/bin/env node
// Renders cv.md to a .docx, so the Word CV is always a view of cv.md rather
// than a second copy that drifts. Content is never invented or reordered here:
// headings, bullets and paragraphs come straight from the file.
//
// Expects the cv.md shape AGENTS.md documents: "## Experience" with
// "### Company — Location", a bold "**Title**" line, then a dates line.

import { readFileSync, writeFileSync } from 'node:fs';
import { P, R, docx, inlineRuns } from './docx-lib.mjs';

const ACCENT = '1F4E79';
const MUTED = '555555';

export function render(md) {
  const lines = md.split(/\r?\n/);
  const body = [];
  let i = 0;
  let afterH3 = 0; // 1 = next non-empty line is the title, 2 = the dates line

  while (i < lines.length) {
    const raw = lines[i];
    const line = raw.trim();
    i += 1;

    if (!line) continue;

    // H1 — candidate name
    if (/^# /.test(line)) {
      body.push(P(R(line.replace(/^# /, ''), { bold: true, size: 34, color: ACCENT }), { after: 60 }));
      // The two lines after the name are the headline and the contact block.
      continue;
    }

    // H2 — section
    if (/^## /.test(line)) {
      body.push(P(R(line.replace(/^## /, '').toUpperCase(), { bold: true, size: 23, color: ACCENT }),
        { before: 280, after: 120, borderBottom: ACCENT, keepNext: true }));
      afterH3 = 0;
      continue;
    }

    // H3 — employer / institution
    if (/^### /.test(line)) {
      body.push(P(R(line.replace(/^### /, ''), { bold: true, size: 22 }),
        { before: 180, after: 20, keepNext: true }));
      afterH3 = 1;
      continue;
    }

    // Bullet
    if (/^[-*] /.test(line)) {
      body.push(P(R('•  ') + inlineRuns(line.replace(/^[-*] /, '')),
        { align: 'both', after: 80, indent: 227 }));
      afterH3 = 0;
      continue;
    }

    // The bold title line directly under an H3, then its dates line
    if (afterH3 === 1 && /^\*\*.+\*\*/.test(line)) {
      body.push(P(inlineRuns(line, { size: 21 }), { after: 20, keepNext: true }));
      afterH3 = 2;
      continue;
    }
    if (afterH3 === 2) {
      body.push(P(R(line, { italic: true, size: 19, color: MUTED }), { after: 100 }));
      afterH3 = 0;
      continue;
    }

    // Everything else: summary, headline, contact block, skills lines, prose
    body.push(P(inlineRuns(line), { align: 'both', after: 110 }));
    afterH3 = 0;
  }

  return docx(body);
}

if (process.argv.includes('--self-test')) {
  const buf = render([
    '# Jane Doe', '**Analyst**', 'City · a@b.c',
    '## Experience',
    '### Acme — Paris', '**Lead**', 'January 2020 – Present',
    '- Did a thing with **bold** & <angle> chars.',
    '## Skills', '**Tools:** QGIS, Power BI',
  ].join('\n'));
  const s = buf.toString('latin1');
  console.assert(buf.subarray(0, 2).toString() === 'PK', 'not a zip');
  console.assert(buf.readUInt32LE(buf.length - 22) === 0x06054b50, 'bad EOCD');
  console.assert(s.includes('word/document.xml'), 'missing document part');
  console.log('self-test OK');
  process.exit(0);
}

const si = process.argv.indexOf('--source');
const oi = process.argv.indexOf('--out');
const src = si !== -1 ? process.argv[si + 1] : 'cv.md';
const out = oi !== -1 ? process.argv[oi + 1] : 'output/cv.docx';
writeFileSync(out, render(readFileSync(src, 'utf8')));
console.log(`✅ DOCX generated: ${out}  (source: ${src})`);
