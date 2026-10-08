/**
 * brand.js - Munro Wilson Limited document brand module
 *
 * Every client-facing .docx should be built from these helpers rather than
 * from ad-hoc formatting. Change a value in brand.json and every document
 * follows.
 *
 * Usage:
 *   const B = require('./brand');
 *   const logoBuffer = await B.logo();
 *   const badges = await B.loadAccreditations();
 *   const doc = new Document(B.docShell({ children, logoBuffer, meta }));
 *
 * Keep this file and brand.json in the brand asset repo alongside the logos.
 */

const fs = require('fs');
const path = require('path');
const https = require('https');
const sharp = require('sharp');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ShadingType,
  Header, Footer, PageNumber, LevelFormat, VerticalAlign,
} = require('docx');

const T = JSON.parse(fs.readFileSync(path.join(__dirname, 'brand.json'), 'utf8'));

const C = T.colour;
const PAGE = T.page;
const FONT = T.type.family;
const pt = (n) => Math.round(n * 2);           // points -> half-points
const NONE = { style: BorderStyle.NONE, size: 0, color: 'auto' };
const NO_BORDERS = {
  top: NONE, bottom: NONE, left: NONE, right: NONE,
  insideHorizontal: NONE, insideVertical: NONE,
};

/* ------------------------------------------------------------------ assets */

/** Download a file from the brand asset repo, cached in /tmp per build host. */
function fetchAsset(file) {
  const cache = path.join('/tmp', 'mwbrand-' + file.replace(/[^\w.]/g, '_'));
  if (fs.existsSync(cache)) return Promise.resolve(fs.readFileSync(cache));

  const url = T.assets.raw + encodeURIComponent(file);
  return new Promise((resolve, reject) => {
    https.get(url, (res) => {
      if (res.statusCode !== 200) return reject(new Error(`${url} returned HTTP ${res.statusCode}`));
      const chunks = [];
      res.on('data', (d) => chunks.push(d));
      res.on('end', () => {
        const buf = Buffer.concat(chunks);
        fs.writeFileSync(cache, buf);
        resolve(buf);
      });
    }).on('error', reject);
  });
}

/**
 * Crop a PNG to the bounding box of its non-transparent pixels.
 *
 * Needed because supplier-issued marks arrive on oversized square canvases.
 * sharp's own .trim() keys off the top-left pixel, which would eat the black
 * field of the MCS mark, so the alpha bounding box is computed directly.
 */
async function trimToArtwork(buf) {
  const img = sharp(buf).ensureAlpha();
  const { data, info } = await img.raw().toBuffer({ resolveWithObject: true });
  const { width, height, channels } = info;

  let minX = width, minY = height, maxX = -1, maxY = -1;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      if (data[(y * width + x) * channels + 3] > 8) {
        if (x < minX) minX = x;
        if (x > maxX) maxX = x;
        if (y < minY) minY = y;
        if (y > maxY) maxY = y;
      }
    }
  }
  if (maxX < 0) return { buffer: buf, width, height };          // fully transparent, leave alone

  const out = await sharp(buf)
    .extract({ left: minX, top: minY, width: maxX - minX + 1, height: maxY - minY + 1 })
    .png()
    .toBuffer();
  const meta = await sharp(out).metadata();
  return { buffer: out, width: meta.width, height: meta.height };
}

/** Load a named asset, trimming and measuring it. Returns { buffer, width, height }. */
async function asset(key, { reversed = false } = {}) {
  const spec = T.assets[key] || (T.assets.accreditations || []).find((a) => a.key === key);
  if (!spec) throw new Error(`No asset "${key}" in brand.json`);

  const file = reversed ? spec.fileReversed : spec.file;
  if (!file) {
    throw new Error(
      `Asset "${key}" has no ${reversed ? 'reversed' : 'standard'} variant in the brand repo. ` +
      (spec.note || ''),
    );
  }

  const raw = await fetchAsset(file);
  if (spec.trim) return trimToArtwork(raw);
  const meta = await sharp(raw).metadata();
  return { buffer: raw, width: meta.width, height: meta.height };
}

/** Convenience wrapper: the primary Munro Wilson logo as a plain Buffer. */
async function logo(which = 'logoPrimary') {
  return (await asset(which)).buffer;
}

/* -------------------------------------------------------------- primitives */

const run = (text, o = {}) => new TextRun({
  text,
  font: FONT,
  size: pt(o.size || T.type.sizes.body),
  bold: !!o.bold,
  italics: !!o.italics,
  color: o.color || C.textBody,
});

function body(text, o = {}) {
  return new Paragraph({
    children: [run(text, o)],
    alignment: o.align || AlignmentType.LEFT,
    spacing: { after: o.after === undefined ? 120 : o.after, line: 260 },
  });
}

/** Full-width navy section band. The main structural marker of the document. */
function h1(text, o = {}) {
  return new Paragraph({
    children: [run(text, { size: T.type.sizes.h1, bold: true, color: C.white })],
    shading: { type: ShadingType.CLEAR, fill: C.navy, color: 'auto' },
    spacing: { before: o.pageBreakBefore ? 0 : 320, after: 200 },
    indent: { left: 110, right: 110 },
    outlineLevel: 0,
    pageBreakBefore: !!o.pageBreakBefore,
  });
}

/** Navy sub heading with a hairline rule under it. */
function h2(text) {
  return new Paragraph({
    children: [run(text, { size: T.type.sizes.h2, bold: true, color: C.navy })],
    spacing: { before: 260, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.tintMid, space: 4 } },
    outlineLevel: 1,
  });
}

function h3(text) {
  return new Paragraph({
    children: [run(text, { size: T.type.sizes.h3, bold: true, color: C.textBody })],
    spacing: { before: 180, after: 80 },
    outlineLevel: 2,
  });
}

function bullets(items) {
  return items.map((t) => new Paragraph({
    children: [run(t)],
    numbering: { reference: 'mw-bullets', level: 0 },
    spacing: { after: 60, line: 250 },
  }));
}

/**
 * Pale tinted callout with a navy left edge. Used for commitments,
 * clarifications and anything the reader must not skim past.
 */
function callout(lines) {
  const arr = Array.isArray(lines) ? lines : [lines];
  return new Table({
    width: { size: PAGE.contentWidthDxa, type: WidthType.DXA },
    columnWidths: [PAGE.contentWidthDxa],
    borders: {
      ...NO_BORDERS,
      left: { style: BorderStyle.SINGLE, size: 18, color: C.navy },
    },
    rows: [new TableRow({
      cantSplit: true,
      children: [new TableCell({
        width: { size: PAGE.contentWidthDxa, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: C.tintPale, color: 'auto' },
        margins: { top: 160, bottom: 160, left: 200, right: 200 },
        children: arr.map((t, i) => body(t, { after: i === arr.length - 1 ? 0 : 100 })),
      })],
    })],
  });
}

/** Two-column label/value table. The workhorse for parameter schedules. */
function paramTable(rows, labelWidth = 3000) {
  const valueWidth = PAGE.contentWidthDxa - labelWidth;
  return new Table({
    width: { size: PAGE.contentWidthDxa, type: WidthType.DXA },
    columnWidths: [labelWidth, valueWidth],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 2, color: C.tintMid },
      bottom: { style: BorderStyle.SINGLE, size: 2, color: C.tintMid },
      left: { style: BorderStyle.SINGLE, size: 2, color: C.tintMid },
      right: { style: BorderStyle.SINGLE, size: 2, color: C.tintMid },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: C.tintMid },
      insideVertical: { style: BorderStyle.SINGLE, size: 2, color: C.tintMid },
    },
    rows: rows.map(([label, value]) => new TableRow({
      cantSplit: true,
      children: [
        new TableCell({
          width: { size: labelWidth, type: WidthType.DXA },
          shading: { type: ShadingType.CLEAR, fill: C.tintPale, color: 'auto' },
          margins: { top: 90, bottom: 90, left: 140, right: 140 },
          verticalAlign: VerticalAlign.CENTER,
          children: [body(label, { bold: true, size: T.type.sizes.tableBody, after: 0 })],
        }),
        new TableCell({
          width: { size: valueWidth, type: WidthType.DXA },
          margins: { top: 90, bottom: 90, left: 140, right: 140 },
          verticalAlign: VerticalAlign.CENTER,
          children: [body(value, { size: T.type.sizes.tableBody, after: 0 })],
        }),
      ],
    })),
  });
}

/** N-column table with a navy header row. */
function matrixTable(headers, rows, widths) {
  const cols = widths || headers.map(() => Math.floor(PAGE.contentWidthDxa / headers.length));
  const diff = PAGE.contentWidthDxa - cols.reduce((a, b) => a + b, 0);
  cols[cols.length - 1] += diff;                       // widths must sum to table width

  const cell = (text, o = {}) => new TableCell({
    width: { size: o.w, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: o.fill || C.white, color: 'auto' },
    margins: { top: 90, bottom: 90, left: 140, right: 140 },
    verticalAlign: VerticalAlign.CENTER,
    children: [body(text, {
      bold: !!o.bold,
      color: o.color || C.textBody,
      size: T.type.sizes.tableBody,
      after: 0,
    })],
  });

  return new Table({
    width: { size: PAGE.contentWidthDxa, type: WidthType.DXA },
    columnWidths: cols,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 2, color: C.tintMid },
      bottom: { style: BorderStyle.SINGLE, size: 2, color: C.tintMid },
      left: { style: BorderStyle.SINGLE, size: 2, color: C.tintMid },
      right: { style: BorderStyle.SINGLE, size: 2, color: C.tintMid },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: C.tintMid },
      insideVertical: { style: BorderStyle.SINGLE, size: 2, color: C.tintMid },
    },
    rows: [
      new TableRow({
        tableHeader: true,
        cantSplit: true,
        children: headers.map((h, i) => cell(h, { w: cols[i], fill: C.navy, color: C.white, bold: true })),
      }),
      ...rows.map((r, ri) => new TableRow({
        cantSplit: true,
        children: r.map((v, i) => cell(v, {
          w: cols[i],
          fill: ri % 2 ? C.tintFaint : C.white,
          bold: i === 0,
        })),
      })),
    ],
  });
}

/* ---------------------------------------------------- accreditation strip */

/**
 * Load the accreditation marks ready for placement. Call once per build and
 * pass the result into accreditations() or credentials().
 *
 * `onDark` switches to reversed artwork where a variant exists, and throws
 * for any mark that has none, rather than silently placing black on navy.
 */
async function loadAccreditations({ only, onDark = false } = {}) {
  const list = T.assets.accreditations.filter((a) => !only || only.includes(a.key));
  return Promise.all(list.map(async (spec) => ({
    ...spec,
    ...(await asset(spec.key, { reversed: onDark })),
  })));
}

/**
 * Horizontal strip of accreditation marks, optically balanced.
 *
 * The marks are set to a common height rather than a common width, because
 * they range from a tall narrow B-Corp block to a wide NICEIC lockup. Each
 * carries a per-mark scale factor in brand.json to correct the remaining
 * imbalance. Widths follow from the trimmed aspect ratio, so nothing distorts.
 */
function accreditations(loaded, o = {}) {
  const cfg = T.assets.badgeStrip;
  const h = o.heightPt || cfg.heightPt;
  const gap = Math.round((o.gapPt || cfg.gapPt) * 20);          // points -> DXA

  const sized = loaded.map((m) => {
    const height = Math.round(h * (m.scale || 1));
    return { ...m, drawH: height, drawW: Math.round(height * (m.width / m.height)) };
  });

  const cols = [];
  sized.forEach((m, i) => {
    if (i) cols.push(gap);
    cols.push(Math.round(m.drawW * 20));
  });

  const cells = [];
  sized.forEach((m, i) => {
    if (i) {
      cells.push(new TableCell({
        width: { size: gap, type: WidthType.DXA },
        borders: NO_BORDERS,
        children: [new Paragraph({ children: [], spacing: { after: 0 } })],
      }));
    }
    cells.push(new TableCell({
      width: { size: Math.round(m.drawW * 20), type: WidthType.DXA },
      borders: NO_BORDERS,
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 0 },
        children: [new ImageRun({
          data: m.buffer,
          type: 'png',
          altText: { title: m.label, description: m.label, name: m.label },
          transformation: { width: m.drawW, height: m.drawH },
        })],
      })],
    }));
  });

  const strip = new Table({
    width: { size: cols.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: cols,
    alignment: AlignmentType.CENTER,
    borders: NO_BORDERS,
    rows: [new TableRow({ cantSplit: true, children: cells })],
  });

  const out = [strip];
  if (o.caption === undefined ? cfg.captionEnabled : o.caption) {
    out.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 0 },
      children: [run(cfg.caption, { size: T.type.sizes.caption, color: C.textMuted })],
    }));
  }
  return out;
}

/* --------------------------------------------------------- credentials */

/**
 * The "About Munro Wilson" block, at one of four weights.
 *
 * There is no default. Every build must state a mode, because the right
 * answer changes per document and the failure mode is silent: a capability
 * statement gets pasted into a two-page technical note and buries the point.
 *
 *   'none'   nothing at all. Correct for technical notes, fault reports,
 *            variation requests, and anything going to an existing client
 *            who already knows who you are.
 *   'strip'  accreditation marks only, no prose. Signals credentials
 *            without spending a section on them. Good default for quotes.
 *   'short'  two-sentence capability paragraph plus marks.
 *   'full'   capability statement, accreditations and track record.
 *            For tenders, prequalification and first approaches.
 */
function credentials(mode, badges, opts = {}) {
  const MODES = ['none', 'strip', 'short', 'full'];
  if (!MODES.includes(mode)) {
    throw new Error(
      `credentials(): mode must be one of ${MODES.join(', ')} - received ${JSON.stringify(mode)}. ` +
      'This is deliberately required rather than defaulted; choose per document.',
    );
  }
  if (mode === 'none') return [];

  const marks = (h) => accreditations(badges, { heightPt: h, caption: opts.caption });

  if (mode === 'strip') return marks(opts.heightPt || 30);

  const out = [];
  if (mode === 'short') {
    out.push(body('Munro Wilson Limited is a specialist renewables contractor based in Duns, Scottish Borders, delivering solar PV, EV charging and battery storage for commercial, industrial and public sector clients. The company is an NICEIC Approved Contractor, MCS Certified, and a Certified B Corporation.'));
    out.push(new Paragraph({ children: [], spacing: { after: 140 } }));
    out.push(...marks(opts.heightPt || 32));
    return out;
  }

  // full
  out.push(h1('About Munro Wilson Limited', { pageBreakBefore: opts.pageBreakBefore }));
  out.push(body('Munro Wilson Limited is a specialist renewables contractor based in Duns, Scottish Borders. Founded in 2013 and fully focused on renewables since 2020, the company delivers solar PV, EV charging and battery storage installations for commercial, industrial, community and public sector clients across Scotland and the north of England.'));
  out.push(body('The company brings together electrical design and installation expertise with practical experience of managing large-scale renewables projects from initial concept through to grid connection, commissioning and ongoing maintenance. All work is delivered in-house under CDM 2015, with direct site management and engineering oversight on every project.'));
  out.push(h2('Accreditations'));
  out.push(body('Munro Wilson Limited is an NICEIC Approved Contractor and MCS Certified installer, and is a Certified B Corporation. All installations are notified and certified under the relevant scheme.'));
  out.push(new Paragraph({ children: [], spacing: { after: 140 } }));
  out.push(...marks(opts.heightPt || 34));
  out.push(new Paragraph({ children: [], spacing: { after: 140 } }));

  if (opts.trackRecord !== false) {
    out.push(h2('Track Record'));
    out.push(body(opts.trackRecordIntro
      || 'Munro Wilson has delivered over {{INSTALLED_CAPACITY}} of installed solar capacity and generated more than {{GENERATION_TO_DATE}} of renewable electricity, working with businesses, community groups, charities and local organisations. Recent projects include:'));
    out.push(...bullets(opts.references || [
      '{{REFERENCE_PROJECT_1}}',
      '{{REFERENCE_PROJECT_2}}',
      '{{REFERENCE_PROJECT_3}}',
      '{{REFERENCE_PROJECT_4}}',
    ]));
    out.push(body('The company has worked with South of Scotland Enterprise (SOSE) and Local Energy Scotland on a number of community and public sector projects, and has established relationships with the Scottish grid operators and DNOs.'));
  }
  return out;
}

/* ------------------------------------------------------------------- cover */

/**
 * Cover page. Centred logo, navy title band, pale client band, detail table,
 * optional accreditation strip.
 * `details` is an array of [label, value] pairs.
 */
function cover({ logoBuffer, title, subtitle, clientName, siteAddress, details, badges }) {
  const band = (children, fill, opts = {}) => new Table({
    width: { size: PAGE.contentWidthDxa, type: WidthType.DXA },
    columnWidths: [PAGE.contentWidthDxa],
    borders: opts.borders || NO_BORDERS,
    rows: [new TableRow({
      cantSplit: true,
      children: [new TableCell({
        width: { size: PAGE.contentWidthDxa, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill, color: 'auto' },
        margins: { top: 300, bottom: 300, left: 200, right: 200 },
        children,
      })],
    })],
  });

  return [
    new Paragraph({ spacing: { after: 1400 }, children: [] }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 900 },
      children: [new ImageRun({
        data: logoBuffer,
        type: 'png',
        altText: { title: 'Munro Wilson', description: 'Munro Wilson Limited', name: 'Munro Wilson' },
        transformation: { width: 330, height: 78 },   // 1768:418 aspect preserved
      })],
    }),
    band([
      body(title, { size: T.type.sizes.coverTitle, bold: true, color: C.white, align: AlignmentType.CENTER, after: 80 }),
      body(subtitle, { size: T.type.sizes.coverSubtitle, color: C.white, align: AlignmentType.CENTER, after: 0 }),
    ], C.navy),
    new Paragraph({ spacing: { after: 240 }, children: [] }),
    band([
      body(clientName, { size: T.type.sizes.clientName, bold: true, color: C.navy, align: AlignmentType.CENTER, after: 60 }),
      body(siteAddress, { size: T.type.sizes.clientAddress, align: AlignmentType.CENTER, after: 0 }),
    ], C.tintPale, {
      borders: {
        ...NO_BORDERS,
        top: { style: BorderStyle.SINGLE, size: 8, color: C.navy },
        bottom: { style: BorderStyle.SINGLE, size: 8, color: C.navy },
      },
    }),
    new Paragraph({ spacing: { after: 300 }, children: [] }),
    paramTable(details, 3200),
    ...(badges && badges.length
      ? [
        new Paragraph({ spacing: { after: 500 }, children: [] }),
        ...accreditations(badges, { caption: true, heightPt: 52, gapPt: 28 }),
      ]
      : []),
    new Paragraph({ children: [], pageBreakBefore: true }),
  ];
}

/* --------------------------------------------------------- page furniture */

function header(logoBuffer, { documentTitle, clientName }) {
  const left = T.furniture.headerLeft
    .replace('{{DOCUMENT_TITLE}}', documentTitle)
    .replace('{{CLIENT_NAME}}', clientName);

  return new Header({
    children: [new Table({
      width: { size: PAGE.contentWidthDxa, type: WidthType.DXA },
      columnWidths: [7600, PAGE.contentWidthDxa - 7600],
      borders: { ...NO_BORDERS, bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
      rows: [new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            width: { size: 7600, type: WidthType.DXA },
            margins: { top: 40, bottom: 80, left: 0, right: 0 },
            verticalAlign: VerticalAlign.BOTTOM,
            children: [body(left, { size: T.type.sizes.small, color: C.textMuted, after: 0 })],
          }),
          new TableCell({
            width: { size: PAGE.contentWidthDxa - 7600, type: WidthType.DXA },
            margins: { top: 40, bottom: 60, left: 0, right: 0 },
            verticalAlign: VerticalAlign.BOTTOM,
            children: [new Paragraph({
              alignment: AlignmentType.RIGHT,
              spacing: { after: 0 },
              children: [new ImageRun({
                data: logoBuffer,
                type: 'png',
                altText: { title: 'Munro Wilson', description: 'Munro Wilson Limited', name: 'Munro Wilson' },
                transformation: { width: 90, height: 21 },
              })],
            })],
          }),
        ],
      })],
    })],
  });
}

function footer({ confidentiality, documentRef, issueDate, registrationLine }) {
  const kids = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 0 },
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.tintMid, space: 6 } },
      children: [
        run(`${confidentiality}  |  Ref: ${documentRef}  |  Issue: ${issueDate}  |  Page `,
          { size: T.type.sizes.small, color: C.textMuted }),
        new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: pt(T.type.sizes.small), color: C.textMuted }),
        run(' of ', { size: T.type.sizes.small, color: C.textMuted }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: pt(T.type.sizes.small), color: C.textMuted }),
      ],
    }),
  ];
  if (registrationLine) {
    kids.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 40, after: 0 },
      children: [run(registrationLine, { size: T.type.sizes.small, color: C.textMuted })],
    }));
  }
  return new Footer({ children: kids });
}

/* ---------------------------------------------------------------- assembly */

/** Document options: page setup, numbering and default run formatting. */
function docShell({ children, logoBuffer, meta }) {
  return {
    creator: 'Munro Wilson Limited',
    title: meta.documentTitle,
    styles: {
      default: {
        document: { run: { font: FONT, size: pt(T.type.sizes.body), color: C.textBody } },
      },
    },
    numbering: {
      config: [{
        reference: 'mw-bullets',
        levels: [{
          level: 0,
          format: LevelFormat.BULLET,
          text: '•',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 360, hanging: 200 } } },
        }],
      }],
    },
    sections: [{
      properties: {
        page: {
          size: { width: PAGE.widthDxa, height: PAGE.heightDxa },
          margin: {
            top: PAGE.marginDxa, right: PAGE.marginDxa,
            bottom: PAGE.marginDxa, left: PAGE.marginDxa,
            header: PAGE.headerDxa, footer: PAGE.footerDxa,
          },
        },
      },
      headers: { default: header(logoBuffer, meta) },
      footers: { default: footer(meta) },
      children,
    }],
  };
}

async function write(doc, outPath) {
  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(outPath, buf);
  return outPath;
}

module.exports = {
  tokens: T, colour: C, page: PAGE,
  asset, logo, fetchAsset, trimToArtwork,
  loadAccreditations, accreditations, credentials,
  run, body, h1, h2, h3, bullets, callout,
  paramTable, matrixTable, cover, header, footer, docShell, write,
};
