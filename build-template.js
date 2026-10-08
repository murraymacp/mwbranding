/**
 * build-template.js - produces MW-Document-Template.docx
 *
 * A de-identified skeleton of the Munro Wilson tender/proposal document.
 * Everything project or customer specific is replaced with a {{TOKEN}}.
 * Standing company content (capability statement, standards list, generic
 * exclusions) is retained verbatim because it is reusable on every job.
 *
 *   node build-template.js
 */

const { Document, Paragraph, AlignmentType } = require('docx');
const B = require('./brand');

const META = {
  documentTitle: '{{DOCUMENT_TITLE}}',
  clientName: '{{CLIENT_NAME}}',
  confidentiality: '{{CONFIDENTIALITY}}',
  documentRef: '{{DOCUMENT_REF}}',
  issueDate: '{{ISSUE_DATE}}',
};

(async () => {
  const logoBuffer = await B.logo();
  const badges = await B.loadAccreditations();
  const kids = [];
  const push = (...x) => x.forEach((i) => kids.push(i));
  const gap = (after = 200) => new Paragraph({ children: [], spacing: { after } });

  /* ---------------------------------------------------------------- cover */

  push(...B.cover({
    logoBuffer,
    title: '{{DOCUMENT_TITLE}}',
    subtitle: '{{DOCUMENT_SUBTITLE}}',
    clientName: '{{CLIENT_NAME}}',
    siteAddress: '{{SITE_ADDRESS}}',
    details: [
      ['Submitted to', '{{SUBMITTED_TO}}'],
      ['Submitted by', 'Munro Wilson Limited'],
      ['Document Reference', '{{DOCUMENT_REF}}'],
      ['Contract Basis', '{{CONTRACT_BASIS}}'],
      ['Issue Date', '{{ISSUE_DATE}}'],
      ['Offer Valid For', '{{VALIDITY_PERIOD}}'],
      ['Target Completion', '{{TARGET_COMPLETION}}'],
      ['Consultant / Specification', '{{CONSULTANT_REF}}'],
    ],
    badges,
  }));

  /* -------------------------------------------- standing capability content */

  // Mode is required, not defaulted. See credentials() in brand.js.
  push(...B.credentials('full', badges));

  /* ------------------------------------------------------------ section 1 */

  push(B.h1('1. Introduction'));
  push(B.body('We understand the requirement to be {{SCOPE_SUMMARY}} at {{SITE_NAME}}. This document forms the submission of Munro Wilson Limited to {{SUBMITTED_TO}} and addresses the {{PACKAGE_NAME}} package only.'));
  push(B.body('The full tender package has been reviewed, including {{REFERENCED_DOCUMENTS}}.'));
  push(B.body('This submission sets out the scope, the technical approach, the proposed programme, relevant experience and the health and safety approach. {{OPTION_COUNT}} independently priced options are also presented for consideration.'));
  push(gap(160));
  push(B.callout('{{HEADLINE_COMMITMENT}} - use this callout for the single statement the reader must not miss. Named suppliers, accreditation confirmations, or acceptance of a contractual obligation.'));

  /* ------------------------------------------------------------ section 2 */

  push(B.h1('2. System Overview'));
  push(B.paramTable([
    ['{{PARAMETER}}', '{{DETAIL}}'],
    ['Installed Capacity', '{{CAPACITY}}'],
    ['Key Equipment', '{{EQUIPMENT_SCHEDULE}}'],
    ['Output', '{{OUTPUT}}'],
    ['Annual Production', '{{ANNUAL_PRODUCTION}}'],
    ['Connection Arrangement', '{{CONNECTION}}'],
    ['Network Assumption', '{{NETWORK_TYPE}}'],
    ['Key Standards', '{{KEY_STANDARDS}}'],
    ['Target Commission Date', '{{TARGET_COMPLETION}}'],
  ]));

  /* ------------------------------------------------------------ section 3 */

  push(B.h1('3. Technical Approach'));
  push(B.h2('3.1 {{SUBSECTION_TITLE}}'));
  push(B.body('{{BODY_TEXT}} Body copy sits at 9.5pt Arial on a 13pt line. Keep paragraphs to four or five lines so the document stays scannable on site.'));
  push(B.h3('{{MINOR_HEADING}}'));
  push(B.body('Use the minor heading where a subsection needs internal structure but does not warrant its own number.'));

  push(B.h2('3.2 Risk or Comparison Matrix'));
  push(B.body('The matrix table carries a navy header row and alternating row fill. Use it for risk registers, option comparisons and anywhere the reader must weigh several items against the same criteria.'));
  push(gap(120));
  push(B.matrixTable(
    ['{{COLUMN_1}}', '{{COLUMN_2}}', '{{COLUMN_3}}', '{{COLUMN_4}}'],
    [
      ['{{ROW_LABEL}}', '{{VALUE}}', '{{VALUE}}', '{{COMMENT}}'],
      ['{{ROW_LABEL}}', '{{VALUE}}', '{{VALUE}}', '{{COMMENT}}'],
      ['{{ROW_LABEL}}', '{{VALUE}}', '{{VALUE}}', '{{COMMENT}}'],
    ],
    [2600, 1700, 1700, 4106],
  ));

  /* ------------------------------------------------------------ section 4 */

  push(B.h1('4. Scope of Works'));
  push(B.body('Lead with customer-visible items before technical and installation detail. Model numbers belong in the equipment schedule, not here.'));
  push(B.h2('4.1 Design and Engineering'));
  push(...B.bullets(['{{SCOPE_ITEM}}', '{{SCOPE_ITEM}}', '{{SCOPE_ITEM}}']));
  push(B.h2('4.2 Supply and Installation'));
  push(...B.bullets(['{{SCOPE_ITEM}}', '{{SCOPE_ITEM}}', '{{SCOPE_ITEM}}']));
  push(B.h2('4.3 Certification, Testing and Commissioning'));
  push(...B.bullets([
    'Full test and inspection to BS 7671 with certification issued on completion',
    'Commissioning records and handover documentation',
    'Operation and maintenance manual, as-built drawings and schematics',
    '{{ADDITIONAL_COMMISSIONING_ITEM}}',
  ]));

  /* ------------------------------------------------------------ section 5 */

  push(B.h1('5. Advantages'));
  push(B.body('State the customer-visible benefits of the proposed approach in plain terms. Keep this section short and specific to the scheme rather than general to the company.'));
  push(...B.bullets(['{{ADVANTAGE}}', '{{ADVANTAGE}}', '{{ADVANTAGE}}']));

  push(B.h1('6. Considerations'));
  push(B.body('Set out the constraints, dependencies and matters the client must decide on or accept. Raising these plainly is more credible than omitting them.'));
  push(...B.bullets(['{{CONSIDERATION}}', '{{CONSIDERATION}}', '{{CONSIDERATION}}']));

  /* ------------------------------------------------------------ section 7 */

  push(B.h1('7. Priced Options'));
  push(B.body('The following options are priced independently and can be accepted or declined independently of each other and of the base contract.'));
  push(gap(120));
  push(B.matrixTable(
    ['Ref', 'Option', 'Description'],
    [
      ['O-1', '{{OPTION_NAME}}', '{{OPTION_DESCRIPTION}}'],
      ['O-2', '{{OPTION_NAME}}', '{{OPTION_DESCRIPTION}}'],
      ['O-3', '{{OPTION_NAME}}', '{{OPTION_DESCRIPTION}}'],
    ],
    [900, 2400, 6806],
  ));

  /* ------------------------------------------------------------ section 8 */

  push(B.h1('8. Health and Safety'));
  push(B.body('All work is delivered under CDM 2015 with direct site management and engineering oversight. A project-specific RAMS is issued before mobilisation and reviewed with the principal contractor.'));
  push(gap(120));
  push(B.matrixTable(
    ['Risk', 'Control Measure'],
    [
      ['{{RISK}}', '{{CONTROL}}'],
      ['{{RISK}}', '{{CONTROL}}'],
      ['{{RISK}}', '{{CONTROL}}'],
    ],
    [3400, 6706],
  ));

  /* ------------------------------------------------------------ section 9 */

  push(B.h1('9. Price'));
  push(B.h2('9.1 Contract Basis'));
  push(B.body('This offer is submitted on the basis of {{CONTRACT_BASIS}}. Any proposed amendments will need to be agreed during the contracting stage.'));
  push(B.h2('9.2 Programme'));
  push(...B.bullets([
    'Lead time from order to start on site: {{LEAD_TIME}}',
    'Construction period on site: {{CONSTRUCTION_PERIOD}}, weather dependent',
    'Target commissioning: {{TARGET_COMPLETION}}',
    'Detailed programme to be agreed with {{SUBMITTED_TO}} following award',
  ]));
  push(B.h2('9.3 Validity'));
  push(B.body('This offer remains open for acceptance for {{VALIDITY_PERIOD}} from {{ISSUE_DATE}}. The price is fixed for the duration of the contract, subject to {{VARIATION_MECHANISM}}.'));

  /* ----------------------------------------------------------- section 10 */

  push(B.h1('10. Exclusions and Assumptions'));
  push(B.h2('10.1 Exclusions'));
  push(...B.bullets([
    '{{PROJECT_SPECIFIC_EXCLUSION}}',
    '{{PROJECT_SPECIFIC_EXCLUSION}}',
    'Builders work, making good and decoration',
    'Ground reinstatement, seeding and landscaping unless separately priced as an option',
    'Site welfare, security and temporary roads',
    'DNO network reinforcement works',
    'VAT',
  ]));
  push(B.h2('10.2 Assumptions'));
  push(...B.bullets([
    '{{PROJECT_SPECIFIC_ASSUMPTION}}',
    '{{PROJECT_SPECIFIC_ASSUMPTION}}',
    'Access to the working area will be available from a date to be agreed within the programme',
    'This offer remains open for acceptance for {{VALIDITY_PERIOD}} from {{ISSUE_DATE}}',
  ]));

  /* ----------------------------------------------------------- section 11 */

  push(B.h1('11. Applicable Standards'));
  push(B.body('Trim this list to those applicable to the package. It is retained in full as standing content because it is reusable on every job.'));
  push(...B.bullets([
    'BS 7671:2018+A2:2022 - Requirements for Electrical Installations (18th Edition)',
    'MIS3002 - Microgeneration Installation Standard: PV Systems',
    'MCS005 - Performance and Installation Standard for PV Modules',
    'IET Code of Practice for Grid-Connected Solar PV Systems',
    'IET Code of Practice for Electric Vehicle Charging Equipment Installation',
    'IET Code of Practice for Electrical Energy Storage Systems',
    'BS EN 62446:2016 - Grid Connected PV Systems: Testing, Documentation and Maintenance',
    'G98 / G99 - Requirements for the Connection of Generation Equipment to the Distribution Network',
    'G100 - Code of Practice for Export Limitation Systems',
    'RC62 - Fire Safety Design Guide for Solar PV Systems',
    'BS EN 62305:2011 - Protection Against Lightning',
    'BS EN 50521:2008+A1:2012 - Connectors for PV Systems',
    'BS EN 61215-2:2017 - Terrestrial PV Modules: Design Qualification and Type Approval',
    'BS EN 61730-1:2018 and BS EN 61730-2:2018 - PV Module Safety Qualification',
    'BS EN 1991-1-4 - Eurocode 1: Actions on Structures, Wind Actions',
    'CDM Regulations 2015 - Construction Design and Management',
    'Work at Height Regulations 2005',
  ]));

  push(gap(300));
  push(B.body('This document is submitted by Munro Wilson Limited in confidence and constitutes the {{SUBMISSION_TYPE}} to {{SUBMITTED_TO}} for {{SCOPE_SUMMARY}} at {{SITE_NAME}}. Valid for {{VALIDITY_PERIOD}} from {{ISSUE_DATE}}.',
    { italics: true, align: AlignmentType.CENTER }));
  push(...B.registrationsBlock());

  const doc = new Document(B.docShell({ children: kids, logoBuffer, meta: META }));
  const out = await B.write(doc, __dirname + '/MW-Document-Template.docx');
  console.log('written:', out);
})().catch((e) => { console.error(e); process.exit(1); });
