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

## Usage

```bash
npm install docx sharp
node build-template.js      # -> MW-Document-Template.docx
node build-report.js        # -> MW-Brand-Identity-Report.docx
```

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

## Rules

1. Take assets from here, never from an older document or the website.
2. Change a value in `brand.json`, not in a document.
3. Keep true alpha artwork. Do not rely on Word's Set Transparent Color.
4. Body-sized text needs 4.5:1 contrast against its background.

## Outstanding

- No reversed Munro Wilson wordmark. Grey `#A7A9AC` on navy `#1F5C8B` measures
  3.01:1, so the logo cannot currently sit on a navy panel.
- No reversed NICEIC or MCS marks, so the badge strip is white-background only.
- `{{REGISTRATION_LINE}}` is unset: the footer needs the company number, VAT
  number and scheme registration numbers as worded on invoices.
- Quote convention requires each numbered section to start on a new page. The
  `h1()` component supports `pageBreakBefore` but the builds do not yet set it.
