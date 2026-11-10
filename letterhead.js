/**
 * letterhead.js - Munro Wilson Limited letterhead and letter layout
 *
 * Builds on brand.js so letters share the colour, type, logo, accreditation
 * strip and statutory footer line used by quotes and tenders.
 *
 * First page: full letterhead (logo left, contact block right, navy rule),
 * accreditation strip and statutory registration line in the footer.
 * Continuation pages (if a letter ever runs over): compact logo header and the
 * registration line only.
 *
 * Usage:
 *   const LH = require('./letterhead');
 *   const doc = await LH.letterDocument(spec);   // spec described in README
 *   await LH.write(doc, 'out/letter.docx');
 */

const {
  Document, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, Header, Footer, PageNumber, VerticalAlign,
} = require('docx');
const B = require('./brand');

const T = B.tokens;
const C = B.colour;
const L = T.letter;
const FONT = T.type.family;
const pt = (n) => Math.round(n * 2);            // points -> half-points
const NONE = { style: BorderStyle.NONE, size: 0, color: 'auto' };
const NO_BORDERS = {
  top: NONE, bottom: NONE, left: NONE, right: NONE,
  insideHorizontal: NONE, insideVertical: NONE,
};

const CONTENT_W = T.page.widthDxa - 2 * L.marginSideDxa;
const LOGO_ASPECT = 418 / 1768;                  // MunroWFinal.png, height / width

/* -------------------------------------------------------------- text */

/**
 * Text runs with light inline markup: **bold** only. Anything heavier than
 * that belongs in a document, not a letter.
 */
function runs(text, o = {}) {
  return String(text).split(/(\*\*[^*]+\*\*)/).filter(Boolean).map((part) => {
    const bold = /^\*\*[^*]+\*\*$/.test(part);
    return new TextRun({
      text: bold ? part.slice(2, -2) : part,
      font: FONT,
      size: pt(o.size || L.bodyPt),
      bold: bold || !!o.bold,
      color: o.color || C.textBody,
    });
  });
}

function para(text, o = {}) {
  return new Paragraph({
    children: runs(text, o),
    alignment: o.align || AlignmentType.LEFT,
    spacing: { before: o.before || 0, after: o.after === undefined ? L.paragraphAfter : o.after, line: L.lineSpacing },
    keepNext: !!o.keepNext,
    keepLines: true,
  });
}

const spacer = (afterPt) => new Paragraph({ children: [], spacing: { after: Math.round(afterPt * 20), line: 240 } });

/** 2026-11-10 -> 10 November 2026. Defaults to today. */
function ukDate(value) {
  const d = value ? new Date(`${value}T12:00:00Z`) : new Date();
  if (Number.isNaN(d.getTime())) throw new Error(`Unrecognised date "${value}". Use YYYY-MM-DD.`);
  return d.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'Europe/London' });
}

/* ------------------------------------------------------- letterhead parts */

function contactLines() {
  const k = T.contact || {};
  const lines = [];
  if (Array.isArray(k.tradingAddress) && k.tradingAddress.length) lines.push(k.tradingAddress.join(', '));
  else if (k.tradingBase) lines.push(k.tradingBase);
  if (k.phone) lines.push(`T  ${k.phone}`);
  if (k.email) lines.push(`E  ${k.email}`);
  if (k.web) lines.push(k.web);
  return lines;
}

/** Full letterhead for page one: logo left, contact right, navy rule beneath. */
function firstPageHeader(logoBuffer) {
  const logoW = L.logoWidthPx;
  const logoH = Math.round(logoW * LOGO_ASPECT);
  const leftW = Math.round(CONTENT_W * 0.5);
  const rightW = CONTENT_W - leftW;
  const k = T.contact || {};

  const right = [];
  if (k.services) {
    right.push(new Paragraph({
      alignment: AlignmentType.RIGHT,
      spacing: { after: 40 },
      children: runs(k.services, { size: T.type.sizes.small, bold: true, color: C.navy }),
    }));
  }
  contactLines().forEach((line, i, all) => right.push(new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: i === all.length - 1 ? 0 : 20 },
    children: runs(line, { size: T.type.sizes.small, color: C.textMuted }),
  })));

  return new Header({
    children: [
      new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: [leftW, rightW],
        borders: { ...NO_BORDERS, bottom: { style: BorderStyle.SINGLE, size: 12, color: C.navy } },
        rows: [new TableRow({
          cantSplit: true,
          children: [
            new TableCell({
              width: { size: leftW, type: WidthType.DXA },
              borders: { ...NO_BORDERS, bottom: { style: BorderStyle.SINGLE, size: 12, color: C.navy } },
              margins: { top: 0, bottom: 200, left: 0, right: 0 },
              verticalAlign: VerticalAlign.BOTTOM,
              children: [new Paragraph({
                spacing: { after: 0 },
                children: [new ImageRun({
                  data: logoBuffer,
                  type: 'png',
                  altText: { title: 'Munro Wilson', description: 'Munro Wilson Limited', name: 'Munro Wilson' },
                  transformation: { width: logoW, height: logoH },
                })],
              })],
            }),
            new TableCell({
              width: { size: rightW, type: WidthType.DXA },
              borders: { ...NO_BORDERS, bottom: { style: BorderStyle.SINGLE, size: 12, color: C.navy } },
              margins: { top: 0, bottom: 200, left: 0, right: 0 },
              verticalAlign: VerticalAlign.BOTTOM,
              children: right,
            }),
          ],
        })],
      }),
    ],
  });
}

/** Compact header for any continuation page. */
function continuationHeader(logoBuffer, recipientName) {
  const w = L.continuationLogoWidthPx;
  const leftW = Math.round(CONTENT_W * 0.7);
  const rightW = CONTENT_W - leftW;
  const prefix = recipientName ? `${recipientName}  |  ` : '';

  return new Header({
    children: [new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [leftW, rightW],
      borders: { ...NO_BORDERS, bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
      rows: [new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            width: { size: leftW, type: WidthType.DXA },
            borders: { ...NO_BORDERS, bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
            margins: { top: 0, bottom: 80, left: 0, right: 0 },
            verticalAlign: VerticalAlign.BOTTOM,
            children: [new Paragraph({
              spacing: { after: 0 },
              children: [
                ...runs(`${prefix}Page `, { size: T.type.sizes.small, color: C.textMuted }),
                new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: pt(T.type.sizes.small), color: C.textMuted }),
                ...runs(' of ', { size: T.type.sizes.small, color: C.textMuted }),
                new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: pt(T.type.sizes.small), color: C.textMuted }),
              ],
            })],
          }),
          new TableCell({
            width: { size: rightW, type: WidthType.DXA },
            borders: { ...NO_BORDERS, bottom: { style: BorderStyle.SINGLE, size: 4, color: C.navy } },
            margins: { top: 0, bottom: 60, left: 0, right: 0 },
            verticalAlign: VerticalAlign.BOTTOM,
            children: [new Paragraph({
              alignment: AlignmentType.RIGHT,
              spacing: { after: 0 },
              children: [new ImageRun({
                data: logoBuffer,
                type: 'png',
                altText: { title: 'Munro Wilson', description: 'Munro Wilson Limited', name: 'Munro Wilson' },
                transformation: { width: w, height: Math.round(w * LOGO_ASPECT) },
              })],
            })],
          }),
        ],
      })],
    })],
  });
}

/**
 * Statutory disclosure (Companies Act 2006 trading disclosures) for letters.
 * Same content as brand.js footerRegistration(), but broken deliberately
 * over two lines: at letter width the single line wraps mid-postcode.
 */
function registrationLine() {
  const co = T.company;
  const style = { size: T.type.sizes.caption, color: C.textMuted };
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 140, after: 20 },
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.tintMid, space: 6 } },
      children: runs(`${co.registeredName} Registered in ${co.jurisdiction} no. ${co.companyNumber}.  VAT no. ${co.vat}`, style),
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 0 },
      children: runs(`Registered Office: ${co.registeredOffice}`, style),
    }),
  ];
}

/**
 * Letter footer. `credentials` is required, as in brand.js, and only the two
 * modes that suit a letterhead are accepted: 'strip' (accreditation marks)
 * or 'none'. A capability section has no place in a letter footer.
 */
function letterFooter(credentials, badges) {
  const MODES = ['strip', 'none'];
  if (!MODES.includes(credentials)) {
    throw new Error(
      `letterhead: credentials must be 'strip' or 'none' - received ${JSON.stringify(credentials)}. ` +
      'Required rather than defaulted, as in brand.js credentials().',
    );
  }
  const kids = [];
  if (credentials === 'strip') {
    kids.push(...B.accreditations(badges, { heightPt: L.badgeHeightPx, gapPt: L.badgeGapPt, caption: false }));
  }
  kids.push(...registrationLine());
  return new Footer({ children: kids });
}

/* -------------------------------------------------------------- letter */

/**
 * Body of a letter from a spec object. All fields except body are optional.
 *
 *   recipient   array of address lines
 *   date        YYYY-MM-DD, defaults to today
 *   ref         our reference, shown under the date
 *   subject     bold navy line above the salutation
 *   salutation  defaults to "Dear Sir/Madam,"
 *   body        array of paragraphs (**bold** supported)
 *   closing     defaults to "Yours faithfully,"
 *   signatory   { name, lines: [title, company] }
 *   enclosures  array of strings, listed after the signature as "Enc."
 */
function letterBody(spec) {
  if (!Array.isArray(spec.body) || !spec.body.length) throw new Error('letter spec needs a non-empty body array');
  const out = [];

  (spec.recipient || []).forEach((line, i, all) => out.push(para(line, { after: 0, keepNext: i < all.length - 1 })));
  if (spec.recipient && spec.recipient.length) out.push(spacer(14));

  out.push(para(ukDate(spec.date), { after: spec.ref ? 40 : L.paragraphAfter }));
  if (spec.ref) out.push(para(`Our ref: ${spec.ref}`, { color: C.textMuted, size: T.type.sizes.small + 0.5 }));

  if (spec.subject) out.push(para(spec.subject, { bold: true, color: C.navy, keepNext: true, before: 60 }));

  out.push(para(spec.salutation || 'Dear Sir/Madam,', { keepNext: true }));
  spec.body.forEach((p) => out.push(para(p)));

  // Closing, signature space and name are kept together so a signature never
  // lands alone at the top of a second page.
  out.push(para(spec.closing || 'Yours faithfully,', { keepNext: true, after: 0 }));
  out.push(new Paragraph({ children: [], keepNext: true, spacing: { after: Math.round(L.signatureGapPt * 20), line: 240 } }));
  const sig = spec.signatory || {};
  if (sig.name) out.push(para(sig.name, { bold: true, after: 0, keepNext: !!(sig.lines && sig.lines.length) }));
  (sig.lines || []).forEach((line, i, all) => out.push(para(line, { after: 0, color: C.textMuted, keepNext: i < all.length - 1 })));

  if (spec.enclosures && spec.enclosures.length) {
    out.push(spacer(12));
    spec.enclosures.forEach((e, i) => out.push(para(`${i === 0 ? 'Enc.' : '    '}  ${e}`, { after: 0, size: T.type.sizes.small + 0.5, color: C.textMuted })));
  }
  return out;
}

/**
 * Complete letter document. Pass `blank: true` for an empty letterhead to
 * type into by hand in Word.
 */
async function letterDocument(spec = {}) {
  const logoBuffer = await B.logo();
  const badges = spec.credentials === 'strip' ? await B.loadAccreditations() : [];
  const footer = letterFooter(spec.credentials, badges);

  const children = spec.blank
    ? [new Paragraph({ children: runs(''), spacing: { after: L.paragraphAfter, line: L.lineSpacing } })]
    : letterBody(spec);

  const recipientName = spec.recipient && spec.recipient[0];

  return new Document({
    creator: 'Munro Wilson Limited',
    title: spec.title || (spec.blank ? 'Munro Wilson Letterhead' : `Letter to ${recipientName || 'recipient'}`),
    styles: {
      default: {
        document: { run: { font: FONT, size: pt(L.bodyPt), color: C.textBody } },
      },
    },
    sections: [{
      properties: {
        titlePage: true,
        page: {
          size: { width: T.page.widthDxa, height: T.page.heightDxa },
          margin: {
            top: L.marginTopDxa, bottom: L.marginBottomDxa,
            left: L.marginSideDxa, right: L.marginSideDxa,
            header: L.headerDxa, footer: L.footerDxa,
          },
        },
      },
      headers: {
        first: firstPageHeader(logoBuffer),
        default: continuationHeader(logoBuffer, recipientName),
      },
      footers: {
        first: footer,
        default: new Footer({ children: registrationLine() }),
      },
      children,
    }],
  });
}

module.exports = {
  letterDocument, letterBody, firstPageHeader, continuationHeader, letterFooter, ukDate,
  write: B.write,
};
