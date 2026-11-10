# Munro Wilson brand assets and document system

Single source of truth for Munro Wilson Limited brand assets and the code that
generates client-facing documents from them.

## Why the code lives here

The generator source was previously held only as downloaded files outside any
repository. It was lost when a build environment was cleared, and had to be
reconstructed. Assets and source belong together, under version control.

## Contents

### Assets

| File | Purpose | Notes |
|---|---|---|
| `MunroWFinal.png` | Primary wordmark | 1768x418 RGBA, true alpha. White and pale backgrounds only |
| `niceic_app_contractor.png` | NICEIC Approved Contractor | ~40% transparent padding; trimmed at build time |
| `image-808x1024.png` | MCS Certified | Should be renamed `mcs-certified-black.png` |
| `B-Corp-Logo-Black-RGB.png` | Certified B Corporation | Standard |
| `B-Corp-Logo-White-RGB.png` | Certified B Corporation | Reversed, for dark backgrounds |
| `MunroW - favicon.jpg` | Web favicon | Not usable in documents |

### Source

| File | Purpose |
|---|---|
| `brand.json` | Design tokens: colour, type, page setup, asset manifest, known issues |
| `brand.js` | Document component module for `docx` |
| `build-template.js` | Generates the blank document template |
| `build-report.js` | Generates the brand identity report, and serves as the proving run |
| `letterhead.js` | Letterhead and letter layout, built on `brand.js` |
| `build-letter.js` | Builds a letter from a JSON spec, or a blank letterhead |
| `examples/letter.example.json` | Letter spec showing every field |

## Usage

```bash
npm install docx sharp
node build-template.js      # -> MW-Document-Template.docx
node build-report.js        # -> MW-Brand-Identity-Report.docx

node build-letter.js letters/my-letter.json --pdf          # -> out/<output>.docx and .pdf
node build-letter.js --blank --credentials strip --pdf     # -> out/MW-Letterhead.docx, to type into in Word
```

## Letters

`letters/` and `out/` are gitignored. This repo is public and letters carry
client and commercial detail, so only the layout code lives here. Copy
`examples/letter.example.json` into `letters/` and edit it.

| Field | Notes |
|---|---|
| `credentials` | Required. `strip` (accreditation marks in the footer) or `none` |
| `recipient` | Array of address lines |
| `date` | `YYYY-MM-DD`, printed as 10 November 2026. Defaults to today |
| `ref` | Optional "Our ref" line under the date |
| `subject` | Optional bold navy line above the salutation |
| `salutation` | Defaults to "Dear Sir/Madam," |
| `body` | Array of paragraphs. `**bold**` supported |
| `closing` | Defaults to "Yours faithfully," |
| `signatory` | `{ "name": "...", "lines": ["title", "company"] }` |
| `enclosures` | Optional array, listed as "Enc." after the signature |
| `output` | File name without extension |

Page one carries the full letterhead: logo left, services and contact right,
navy rule, accreditation strip and statutory disclosure in the footer. If a
letter runs over, later pages get a compact header with page numbering.
`--pdf` reports the page count and warns when a letter passes one page.

Letterhead contact details come from the `contact` block in `brand.json`;
any null field is left off. Layout values are in the `letter` block.

Assets are fetched from this repo's raw URLs at build time and cached in
`/tmp`. Nothing is embedded in the source, so updating artwork here updates
every document generated afterwards.

## Components

`brand.js` exports: `cover`, `h1`, `h2`, `h3`, `body`, `bullets`, `callout`,
`paramTable`, `matrixTable`, `accreditations`, `credentials`, `header`,
`footer`, `docShell`.

### credentials(mode, badges)

The "About Munro Wilson" block, at four weights. **There is no default** and a
build throws unless a mode is named, because the right answer changes per
document.

| Mode | Content | Use for |
|---|---|---|
| `none` | Nothing | Technical notes, fault reports, variations, existing clients |
| `strip` | Accreditation marks only | Quotations and routine commercial documents |
| `short` | Two sentences plus marks | Proposals to warm contacts |
| `full` | Capability statement, accreditations, track record | Tenders, prequalification, first approaches |

## Company identity

All registration numbers live in the `company` block of `brand.json` and are
never typed into a document. Two components consume them:

- `footerRegistration()` builds the per-page footer line, carrying only the
  statutory minimum: registered name, part of the UK, registered number,
  registered office, VAT number.
- `registrationsBlock()` renders the full scheme list for the end of a
  document: NICEIC, Approved Body, MCS, OZEV, RECC, plus EORI and DUNS.

The split is deliberate. The Companies Act 2006 trading disclosure rules
cover the first set only; putting all eleven numbers in a footer would cost
roughly an eighth of every page to reference nobody reads in place.

## Rules

1. Take assets from here, never from an older document or the website.
2. Change a value in `brand.json`, not in a document.
3. Keep true alpha artwork. Do not rely on Word's Set Transparent Color.
4. Body-sized text needs 4.5:1 contrast against its background.

## Outstanding

- Letterhead shows "Duns, Scottish Borders" and the website only. Add a full
  trading address, office phone and general email to `contact` in
  `brand.json` when agreed.
- `brand.js` names image sizes `heightPt`, but docx-js `ImageRun` sizes are
  pixels at 96 dpi. The letter block uses `Px` names; brand.js still to align.

- No reversed Munro Wilson wordmark. Grey `#A7A9AC` on navy `#1F5C8B` measures
  3.01:1, so the logo cannot currently sit on a navy panel.
- No reversed NICEIC or MCS marks, so the badge strip is white-background only.
- Confirm the registered company name against the Companies House record: `brand.json`
  holds "Munro Wilson Ltd." but body copy says "Munro Wilson Limited". Trading
  disclosures require the name as registered.
  number and scheme registration numbers as worded on invoices.
- Quote convention requires each numbered section to start on a new page. The
  `h1()` component supports `pageBreakBefore` but the builds do not yet set it.
