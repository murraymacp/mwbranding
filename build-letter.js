#!/usr/bin/env node
/**
 * build-letter.js - Munro Wilson letter or blank letterhead
 *
 *   node build-letter.js letters/my-letter.json          # -> out/<name>.docx
 *   node build-letter.js letters/my-letter.json --pdf    # also writes a PDF
 *   node build-letter.js --blank --credentials strip     # -> out/MW-Letterhead.docx
 *
 * Options:
 *   --pdf                 convert with LibreOffice and report the page count
 *   --out <dir>           output folder, default ./out
 *   --credentials <mode>  strip | none (blank letterhead only; letters set it in the spec)
 *
 * Letter specs go in letters/, which is gitignored: this repo is public and
 * letters carry client and commercial detail. See examples/letter.example.json.
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const LH = require('./letterhead');

function arg(name) {
  const i = process.argv.indexOf(name);
  return i > -1 ? process.argv[i + 1] : undefined;
}
const has = (name) => process.argv.includes(name);

function toPdf(docxPath, outDir) {
  try {
    execFileSync('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', outDir, docxPath], { stdio: 'ignore' });
  } catch (e) {
    console.warn('PDF conversion skipped: LibreOffice (soffice) not available.');
    return null;
  }
  const pdfPath = path.join(outDir, path.basename(docxPath, '.docx') + '.pdf');
  let pages = null;
  try {
    const info = execFileSync('pdfinfo', [pdfPath]).toString();
    pages = Number((info.match(/Pages:\s+(\d+)/) || [])[1]) || null;
  } catch (e) { /* pdfinfo optional */ }
  return { pdfPath, pages };
}

(async () => {
  const outDir = arg('--out') || path.join(__dirname, 'out');
  fs.mkdirSync(outDir, { recursive: true });

  let spec;
  let name;
  if (has('--blank')) {
    spec = { blank: true, credentials: arg('--credentials') };
    name = spec.credentials === 'none' ? 'MW-Letterhead-Plain' : 'MW-Letterhead';
  } else {
    const specPath = process.argv.slice(2).find((a) => a.endsWith('.json'));
    if (!specPath) {
      console.error('Usage: node build-letter.js <spec.json> [--pdf] [--out dir]\n       node build-letter.js --blank --credentials strip|none [--pdf]');
      process.exit(1);
    }
    spec = JSON.parse(fs.readFileSync(specPath, 'utf8'));
    name = spec.output || path.basename(specPath, '.json');
  }

  const doc = await LH.letterDocument(spec);
  const docxPath = path.join(outDir, `${name}.docx`);
  await LH.write(doc, docxPath);
  console.log(`Wrote ${docxPath}`);

  if (has('--pdf')) {
    const r = toPdf(docxPath, outDir);
    if (r) {
      console.log(`Wrote ${r.pdfPath}${r.pages ? ` (${r.pages} page${r.pages > 1 ? 's' : ''})` : ''}`);
      if (r.pages > 1 && !spec.blank) console.warn('Warning: letter runs past one page.');
    }
  }
})().catch((e) => {
  console.error(e.message);
  process.exit(1);
});
