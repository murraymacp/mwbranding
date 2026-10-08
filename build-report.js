/**
 * build-report.js - branding test
 *
 * Short four-section report built entirely from brand.js. Serves as the
 * proving run for the brand system: if this document looks right, the
 * module is working.
 *
 * Credentials mode is 'strip' - marks only, no capability statement. This is
 * a short internal report, so the About block would only get in the way.
 *
 *   node build-report.js
 */

const { Document, Paragraph, AlignmentType } = require('docx');
const B = require('./brand');

const META = {
  documentTitle: 'Consistent Brand Identity',
  clientName: 'Internal',
  confidentiality: 'Internal',
  documentRef: 'MW-BRAND-001',
  issueDate: '8 October 2026',
};

(async () => {
  const logoBuffer = await B.logo();
  const badges = await B.loadAccreditations();
  const kids = [];
  const push = (...x) => x.forEach((i) => kids.push(i));
  const gap = (after = 200) => new Paragraph({ children: [], spacing: { after } });

  push(...B.cover({
    logoBuffer,
    title: 'CONSISTENT BRAND IDENTITY',
    subtitle: 'Why it matters and how it is maintained',
    clientName: 'Internal Report',
    siteAddress: 'Munro Wilson Limited, Duns, Scottish Borders',
    details: [
      ['Prepared for', 'Munro Wilson Limited'],
      ['Document Reference', META.documentRef],
      ['Issue Date', META.issueDate],
      ['Status', 'For discussion'],
      ['Related', 'brand.json v1.1.0, brand.js document module'],
    ],
    badges,
  }));

  /* -------------------------------------------------------------- section 1 */

  push(B.h1('1. What Consistency Actually Buys'));
  push(B.body('Brand consistency is often argued for in aesthetic terms, which makes it easy to dismiss on a busy week. The practical case is narrower and stronger: consistency reduces the effort a reader spends working out who is speaking to them, and it removes a category of doubt that has nothing to do with the technical content.'));
  push(B.body('A commercial client receiving a quotation, a variation and a completion certificate over a six month period is reading three documents from three apparently different organisations if the layout, typeface and colour shift each time. Nothing in that is fatal on its own. The cumulative effect is that the documents read as assembled rather than issued, and assembled documents invite closer scrutiny of everything else in them.'));
  push(B.body('For a business tendering against national contractors, the document is frequently the only evidence of delivery capability available at the point of decision. Estimators and procurement teams reviewing a dozen submissions form an impression of operational discipline from the artefact in front of them, because that is the only thing they have.'));
  push(gap(160));
  push(B.callout('Consistency is not a design preference. It is the cheapest available signal that the same organisation, working to the same standards, produced every document a client receives.'));

  /* -------------------------------------------------------------- section 2 */

  push(B.h1('2. Where Inconsistency Comes From'));
  push(B.body('Inconsistency is almost never a decision. It is what happens when each document is rebuilt by hand from the nearest previous example, and the nearest previous example is whichever file happened to be open. Three specific mechanisms account for most of it.'));

  push(B.h2('2.1 Copy-and-adapt drift'));
  push(B.body('A new quotation starts as last month\'s quotation. Small adjustments accumulate across generations with no reference point to correct against, so the twentieth document in the chain shares little with the first. Each individual step is defensible and the aggregate is not.'));

  push(B.h2('2.2 Multiple copies of the same asset'));
  push(B.body('Logo files proliferate: one in a template, one in an email signature, one exported from the website at a different resolution, one embedded in a document years ago and never revisited. When the artwork is updated, some copies follow and others do not. There is no way to tell which is current because there is no canonical version.'));

  push(B.h2('2.3 Values held in memory rather than written down'));
  push(B.body('If the brand navy exists only as a colour someone picks from a recently-used swatch, it will drift. Colour, type size, margin and spacing values need to exist somewhere they can be read back and applied identically, or each application is a fresh judgement call.'));

  push(gap(140));
  push(B.matrixTable(
    ['Mechanism', 'Typical Symptom', 'Structural Fix'],
    [
      ['Copy-and-adapt drift', 'No two documents share a layout; section order varies by author', 'Generate from a template rather than from the last file'],
      ['Duplicated assets', 'Logo appears at different crops, resolutions or backgrounds', 'One asset repository, referenced by URL, never copied'],
      ['Undocumented values', 'Near-miss colours and inconsistent type sizes across documents', 'Machine-readable token file as the single source of truth'],
      ['Flattened transparency', 'Logo shows a black or white box on some backgrounds', 'Keep true alpha artwork; do not rely on renderer workarounds'],
    ],
    [2400, 3600, 4106],
  ));

  /* -------------------------------------------------------------- section 3 */

  push(B.h1('3. What a Brand System Replaces'));
  push(B.body('The distinction that matters is between a template and a system. A template is a file that gets copied, which means it is subject to the same drift as everything else. A system holds the brand values in one place and generates documents from them, so a change propagates rather than needing to be applied.'));
  push(B.body('Three components carry the whole thing:'));
  push(...B.bullets([
    'A token file holding every colour, type size, margin and spacing value as data rather than as formatting choices',
    'An asset repository holding one canonical copy of each logo and accreditation mark, referenced at build time rather than embedded and forgotten',
    'A module of document components, so a section heading or a pricing table is produced by a named function and cannot be styled two different ways',
  ]));
  push(B.body('The test of whether a system is working is simple. Changing the brand navy should be a one-line edit that every subsequent document picks up automatically. If it requires opening documents and restyling them, it is a template.'));
  push(gap(160));
  push(B.callout([
    'A worked example from this exercise: the accreditation marks supplied for this system arrived on inconsistent canvases, one with roughly forty percent transparent padding around the artwork.',
    'Placed by hand, they would have rendered at visibly different sizes and the discrepancy would have been repeated into every future document. Handled once in the module, by measuring each mark and scaling to a common height, the problem is solved permanently and invisibly.',
  ]));

  /* -------------------------------------------------------------- section 4 */

  push(B.h1('4. Maintaining It'));
  push(B.body('A brand system decays if the rules around it are informal. Four practices keep it intact, none of them onerous.'));

  push(B.h2('4.1 Single source, including the code'));
  push(B.body('Assets live in one repository and are taken from there rather than from an older document. The same rule applies to the generator code itself. Source that exists only as a download in someone\'s folder is one cleared directory away from being lost, and rebuilding it costs far more than committing it.'));

  push(B.h2('4.2 Complete the variant set'));
  push(B.body('An artwork set is only usable if it covers the backgrounds it will meet. Reversed artwork for dark backgrounds is not optional decoration; without it, the alternatives are placing a mark at inadequate contrast or improvising a workaround that will not survive being opened in different software.'));
  push(gap(120));
  push(B.matrixTable(
    ['Asset', 'Standard', 'Reversed', 'Status'],
    [
      ['Munro Wilson wordmark', 'Present', 'Missing', 'Cannot be placed on navy'],
      ['NICEIC Approved Contractor', 'Present', 'Missing', 'White backgrounds only'],
      ['MCS Certified', 'Present', 'Missing', 'White backgrounds only'],
      ['Certified B Corporation', 'Present', 'Present', 'Complete'],
    ],
    [3200, 1700, 1700, 3506],
  ));

  push(B.h2('4.3 Check contrast, not just appearance'));
  push(B.body('Readability is measurable rather than a matter of taste. A contrast ratio below 4.5:1 for body-sized text is a defect regardless of how it looks on a good screen, and it will be noticed first by whoever prints the document or reads it outdoors. The grey Munro Wilson wordmark against the brand navy measures 3.01:1, which is why it is currently restricted to white and pale backgrounds.'));

  push(B.h2('4.4 Decide the content weight per document'));
  push(B.body('Consistency of appearance does not mean identical content. A capability statement belongs in a tender or a first approach to a new client. In a fault report or a variation request it delays the point and reads as padding to a client who already knows the company. The system carries the section at four weights so the choice is made deliberately each time rather than inherited from whichever file was copied.'));
  push(gap(120));
  push(B.matrixTable(
    ['Mode', 'Content', 'Use For'],
    [
      ['none', 'Nothing', 'Technical notes, fault reports, variations, existing clients'],
      ['strip', 'Accreditation marks only', 'Quotations and routine commercial documents'],
      ['short', 'Two-sentence summary plus marks', 'Proposals to warm contacts'],
      ['full', 'Capability statement, accreditations, track record', 'Tenders, prequalification, first approaches'],
    ],
    [1400, 3600, 5106],
  ));

  push(gap(300));
  push(B.body('This report was generated from brand.js and brand.json with no manual formatting applied. It is itself the test: every heading, table, callout and mark on these pages was produced by a named component.',
    { italics: true, align: AlignmentType.CENTER }));
  push(gap(240));
  push(...B.credentials('strip', badges, { heightPt: 28 }));
  push(...B.registrationsBlock());

  const doc = new Document(B.docShell({ children: kids, logoBuffer, meta: META }));
  const out = await B.write(doc, __dirname + '/MW-Brand-Identity-Report.docx');
  console.log('written:', out);
})().catch((e) => { console.error(e); process.exit(1); });
