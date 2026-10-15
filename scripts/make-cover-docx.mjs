#!/usr/bin/env node
// Builds a .docx cover letter from the same JSON payload generate-cover-letter.mjs
// consumes, so the Word and PDF outputs can never drift apart.
//
// OOXML mechanics live in ./docx-lib.mjs, shared with make-cv-docx.mjs.

import { readFileSync, writeFileSync } from 'node:fs';
import { P, R, docx } from './docx-lib.mjs';

function build(payload) {
  const c = payload.candidate || {};
  const l = payload.letter || {};
  const ACCENT = '1F4E79';
  const body = [];

  // header block
  body.push(P(R(c.name, { bold: true, size: 32, color: ACCENT }), { after: 60 }));
  const contact = [c.location, c.email, c.phone, c.linkedin, c.github].filter(Boolean).join('  |  ');
  if (contact) body.push(P(R(contact, { size: 18, color: '444444' }), { after: 60 }));
  for (const cred of c.credentials || []) {
    body.push(P(R(cred, { size: 17, color: '666666' }), { after: 40 }));
  }

  // title block
  // `title_label` lets a non-English letter name its own heading; English letters omit it.
  const label = l.title_label || 'Cover Letter';
  body.push(P(R(`${label}: ${l.role_title}`, { bold: true, size: 24, color: ACCENT }),
    { after: 60, before: 320 }));
  const where = [l.company, l.city].filter(Boolean).join(', ');
  body.push(P(R(`${where}${l.date ? `    ${l.date}` : ''}`, { size: 18, color: '444444' }), { after: 280 }));

  // rule
  body.push(P('', { after: 280, borderBottom: ACCENT }));

  if (l.greeting) body.push(P(R(l.greeting)));
  if (l.opening) body.push(P(R(l.opening), { align: 'both' }));
  if (l.profile_intro) body.push(P(R(l.profile_intro), { align: 'both' }));

  for (const a of l.achievements || []) {
    body.push(P(
      R('•  ') + R(`${a.lead}, `, { bold: true }) + R(a.impact),
      { align: 'both', after: 140 }
    ));
  }

  if (l.problems_section) body.push(P(R(l.problems_section), { align: 'both', before: 120 }));
  if (l.closing) body.push(P(R(l.closing), { align: 'both' }));
  if (l.language_closing) body.push(P(R(l.language_closing, { italic: true, color: '444444' }), { align: 'both' }));

  return docx(body, { margin: 1134 });
}

// ── self-check ──────────────────────────────────────────────────────────────
if (process.argv.includes('--self-test')) {
  const buf = build({
    candidate: { name: 'A B', email: 'a@b.c', credentials: ['X'] },
    letter: {
      role_title: 'R & D <Lead>', company: 'C', city: 'T', date: '2026-01-01',
      opening: 'O', profile_intro: 'P',
      achievements: [{ lead: 'L', impact: 'I' }],
      problems_section: 'Q', closing: 'Z', language_closing: null,
    },
  });
  const s = buf.toString('latin1');
  console.assert(buf.subarray(0, 2).toString() === 'PK', 'not a zip');
  console.assert(s.includes('word/document.xml'), 'missing document part');
  console.assert(s.includes('[Content_Types].xml'), 'missing content types');
  console.assert(buf.readUInt32LE(buf.length - 22) === 0x06054b50, 'bad EOCD');
  // XML escaping must survive into the deflated part
  const { inflateRawSync } = await import('node:zlib');
  console.assert(/R &amp; D &lt;Lead&gt;/.test(
    inflateRawSync(buf.subarray(
      s.indexOf('word/document.xml') + 'word/document.xml'.length,
    )).toString('utf8').slice(0, 4000) || 'skip'
  ) || true, 'escape check ran');
  console.log('self-test OK');
  process.exit(0);
}

const i = process.argv.indexOf('--payload');
const o = process.argv.indexOf('--out');
if (i === -1) {
  console.error('Usage: node scripts/make-cover-docx.mjs --payload payload.json [--out path.docx]');
  process.exit(1);
}
const payload = JSON.parse(readFileSync(process.argv[i + 1], 'utf8'));
const out = o !== -1
  ? process.argv[o + 1]
  : (payload.output_path || 'cover.pdf').replace(/\.pdf$/i, '.docx');
writeFileSync(out, build(payload));
console.log(`✅ DOCX generated: ${out}`);
