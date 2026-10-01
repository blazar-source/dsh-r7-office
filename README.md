# dsh-r7-office

[![CI](https://github.com/blazar-source/dsh-r7-office/actions/workflows/ci.yml/badge.svg)](https://github.com/blazar-source/dsh-r7-office/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![Node.js](https://img.shields.io/badge/node-%3E%3D20.0.0-brightgreen.svg)](https://nodejs.org/)

**R7-Office (Р7-Офис) document processing plugin and Model Context Protocol (MCP) server for DeepSeek Harness.**

Lets an AI agent inspect, read, create, format, edit and convert `DOCX`, `XLSX`,
`PPTX` and `PDF` documents — through standard OOXML manipulation plus the R7
converter you already have installed — without emulating a mouse or keyboard.

> ## ⚠️ Unofficial community project
>
> This is an **independent, community-maintained** project. It is **not
> affiliated with, endorsed by, sponsored by, or supported by АО «Р7»
> (R7-Office) or DeepSeek**, and it is not an official R7-Office or DeepSeek
> product.
>
> *R7-Office* and *Р7-Офис* are trademarks of their respective owners and are
> used here only to describe what this software interoperates with.
>
> This repository **contains no R7-Office code, binaries or other assets**. It
> detects an R7-Office installation already present on the user's machine and
> drives it locally. You must install and license R7-Office yourself.
>
> ---

**Русская документация:** [README.ru.md](README.ru.md)

---

## What it does

- **Format-preserving editing** — replacements and edits keep the original XML
  styles, fonts, colours and numbering hierarchies. Untouched parts of a
  document are preserved byte for byte.
- **All four formats** — DOCX (headings, paragraphs, lists, tables, page
  breaks), XLSX (sheets, ranges, values, formulas), PPTX (slides, titles, text
  frames), PDF (conversion through the local R7 `x2t` engine).
- **Validation built in** — `r7_validate` checks package integrity before and
  after a change.
- **Two ways to run** — as a native DeepSeek Harness plugin, or as a standalone
  MCP server over stdio for any MCP client.
- **Live desktop bridge** — optionally drives a document the user has **open**
  in R7-Office Desktop, over a loopback WebSocket.

## Architecture

```text
DeepSeek Harness ──► dsh-r7-office plugin ──┐
                                            ├──► DOCX / XLSX / PPTX / PDF
Any MCP client ────► MCP server (stdio) ────┤
                                            │
R7-Office Desktop ◄── desktop bridge ◄──────┘
```

Details in [docs/architecture.md](docs/architecture.md); design decisions in
[docs/decisions/](docs/decisions/).

---

## Requirements

| | |
|---|---|
| Node.js | **>= 20.0.0** to run the plugin and MCP server. **>= 22.0.0** to run the full test suite (the bridge and CDP clients use the global `WebSocket`, which only exists from Node 22). |
| Runtime dependencies | **none** — the plugin has zero runtime dependencies |
| R7-Office Desktop | **optional.** Needed for `r7_convert` (→ PDF/HTML), `inspect` fidelity on exotic documents, and the whole desktop bridge. Without it the pure-OOXML tools still work. |
| Platform | Windows, Linux and macOS. R7 auto-detection covers the standard install locations of all three; the live desktop bridge is verified on Windows only (see [limitations](#known-limitations)). |

---

## Clean install from scratch

These steps assume an empty directory and a machine with Node.js 20+.

```bash
# 1. Get the code
git clone https://github.com/blazar-source/dsh-r7-office.git
cd dsh-r7-office

# 2. Install the optional dev dependencies (test-only: the MCP client SDK)
npm install

# 3. Verify the checkout — no R7-Office required for this step
npm test
```

`npm test` runs 445 tests: pure unit tests, OOXML round-trip regression tests,
file end-to-end workflows, security-policy tests and an external MCP client
smoke suite. Tests that need an R7-Office installation **skip themselves** with
a clear message instead of failing, so a clean machine gets a green run.

Then verify the integration you actually intend to use:

```bash
# R7-Office file pipeline (author → edit → validate → PDF). Skips if R7 is absent.
npm run test:live

# DeepSeek Harness plugin activation. Boots a fresh Harness and asserts the
# 28 r7_* tools reached the tool registry.
npm run test:harness -- --profile web
```

### As a DeepSeek Harness plugin

Install the package directory as a bundle:

```bash
# in the DeepSeek Harness UI: plugin_manager → install_bundle
# target: /absolute/path/to/dsh-r7-office
```

or from the CLI equivalent for your profile:

```bash
dsh plugin --profile <profile> add /absolute/path/to/dsh-r7-office
```

Confirm it activated — a fresh Harness boot prints:

```text
[r7-office] зарегистрировано инструментов: 27 (r7_inspect, r7_read, ...); desktop bridge port=7888, developerMode=false
```

> **Note.** Add the row **either** through `install_bundle` **or** by hand in
> the profile's `cordis.patch.yml` — never both. Two entries with the same
> `r7-office` id make the row fail to activate.

The plugin also registers a short usage section into the agent system prompt,
so the agent knows the tools and the intended `inspect → read → edit →
validate → convert` order.

### As a standalone MCP server

The MCP server speaks line-delimited JSON-RPC 2.0 on stdio.

```json
{
  "mcpServers": {
    "r7-office": {
      "command": "node",
      "args": ["/absolute/path/to/dsh-r7-office/src/mcp/cli.js"]
    }
  }
}
```

Put that in your client's configuration file (for example
`claude_desktop_config.json` or `.cursor/mcp.json`). Any stdio MCP client
works; interoperability is verified in `npm test` against the official
`@modelcontextprotocol/client` SDK.

You can also run it by hand:

```bash
echo '{"jsonrpc":"2.0","id":1,"method":"tools/list"}' | node src/mcp/cli.js
```

---

## Tools

| Tool | Purpose |
|---|---|
| `r7_inspect` | Document outline: headings, paragraphs, tables, sheets, slides, metadata |
| `r7_read` | Structured text, Markdown view, or spreadsheet cell ranges |
| `r7_create` | New DOCX / XLSX / PPTX. Refuses to replace an existing file unless `overwrite: true`; every `sheets[]` entry and name is honoured |
| `r7_edit` | Replace, restyle or delete one paragraph by index |
| `r7_replace` | Find and replace text, preserving run formatting |
| `r7_insert` | Insert paragraphs, headings, bullet items or page breaks |
| `r7_table` | Create, inspect or update tables and cells, including merge, borders, shading and column widths |
| `r7_docx_formatting` | Read the normalized formatting of every paragraph, run and table cell: style, font, size, colour, alignment, indents, spacing, line spacing, lists |
| `r7_docx_sections` | Read and set page size, orientation, margins, columns, page breaks and section breaks |
| `r7_docx_header_footer` | List, read, create or retitle headers and footers; page-number fields survive |
| `r7_docx_image` | Insert a PNG/JPEG/GIF at a size, keeping the aspect ratio |
| `r7_docx_hyperlink` | List, insert, retitle or remove hyperlinks and their relationships |
| `r7_sheet_read` | Read values, formulas or the normalized formatting of a sheet or range (e.g. `A1:D10`) |
| `r7_sheet_write` | Write cells or a 2-D matrix, addressed by sheet name or index; dates are stored as real date serials |
| `r7_sheet_format` | Format a cell or range: font, background, borders, alignment, wrap, number format (integer, decimal, currency, percent, date, datetime, custom), merge, column width and row height |
| `r7_sheet_add` | Add a worksheet to an existing workbook; other sheets are untouched |
| `r7_sheet_formula` | Insert or update a formula |
| `r7_slide_read` | Read a slide as a normalized structure: every object's id, type, geometry, text, font, fill, stroke, alignment and paragraphs |
| `r7_slide_create` | Create a deck, or **append** a slide built on one of the deck's own layouts, without altering the slides already there |
| `r7_slide_format` | Restyle or reposition one existing object in place: font, geometry, fill, border, alignment, lists, spacing, text |
| `r7_slide_edit` | Duplicate, move, reorder or delete slides |
| `r7_slide_object` | Add or remove an object: shape, text box or PNG/JPEG image |
| `r7_convert` | Convert via the R7 `x2t` engine (PDF, HTML, TXT, DOCX, XLSX, PPTX). PDF conversion uses the font list R7 generates, repairs a malformed ToUnicode count, and **refuses to return a PDF with no extractable text** instead of silently emitting a blank one. XLSX exports every worksheet by default (`allSheets`) |
| `r7_validate` | Check package integrity and XML health |
| `r7_pdf_inspect` | Report whether a PDF really contains text or was produced without fonts: per-font embedded flag and ToUnicode/Cyrillic map counts, text glyphs, empty fill operators, and a `text` / `outlined` / `mixed` verdict. `repair: true` fixes a malformed ToUnicode count in place |
| `r7_desktop_status` | Desktop bridge connection and effective security mode |
| `r7_desktop_selection` | Read or replace the selection in the open editor |
| `r7_desktop_exec` | Run a safe editor command, or raw DocScript in developer mode |

### Presentations

A deck is built on the layouts it already has. Adding a slide registers the
slide part, its relationship part, its `[Content_Types].xml` override, its
`<p:sldId>` entry and the presentation relationship; the layout, the master,
the theme, the notes and every slide that already existed keep their original
bytes. A placeholder written onto a slide stays a placeholder, so it keeps
inheriting its geometry and typography from the master.

Reading resolves that inheritance for you: `r7_slide_read` reports a title
slide's 60 pt layout title rather than the presentation's 18 pt default, and
reports the position a placeholder inherits from the layout or the master
instead of `null`.

```javascript
import { PptxEngine } from 'dsh-r7-office/r7'

const pptx = new PptxEngine()

// 1. Read: geometry, text, font, fill, stroke, alignment, paragraphs.
const before = await pptx.readSlide('Deck.pptx', 1)
const title = before.slide.objects.find(o => o.placeholder?.type === 'title')
const body = before.slide.objects.find(o => o.placeholder?.type === 'body')
console.log(title.x, title.width, title.font.family, title.font.size)

// 2. Add a slide on one of the deck's own layouts.
const added = await pptx.addSlide('Deck.pptx', {
  layoutType: 'obj',
  title: 'Ключевые выводы',
  paragraphs: [{ text: 'Выручка +18%', bullet: true }]
})

// 3. Restyle one object. Everything not named keeps its original bytes.
await pptx.formatObject('Deck.pptx', {
  slideIndex: added.slideIndex,
  objectId: title.id,
  font: { family: 'Georgia', size: 32, bold: true, color: '#1F6FEB' },
  alignment: 'center'
})

// 4. Add an object: fill, border, text, font, geometry — all in one call.
await pptx.addShape('Deck.pptx', {
  slideIndex: added.slideIndex,
  shape: 'rounded-rectangle',          // rectangle, ellipse, line, arrow, star, …
  x: '2cm', y: '10cm', width: '9.4cm', height: '8.2cm',
  fill: '#1F6FEB', fillTransparency: 0.1,
  line: '#0B3D91', lineWidth: 2,       // lineWidth is in points
  text: 'KPI 98%',
  font: { family: 'Arial', size: 24, bold: true, color: '#FFFFFF' },
  alignment: 'center', verticalAnchor: 'middle'
})

// 5. Insert and replace a picture.
const image = await pptx.addImage('Deck.pptx', {
  slideIndex: added.slideIndex, imagePath: 'chart.png', x: '4cm', y: '4cm', width: '24cm'
})
await pptx.formatObject('Deck.pptx', {
  slideIndex: added.slideIndex, objectId: image.objectId, imagePath: 'chart-v2.png'
})

// 6. Structure, and a PDF.
await pptx.duplicateSlide('Deck.pptx', 1)
await pptx.moveSlide('Deck.pptx', 2, 0)
await pptx.deleteSlide('Deck.pptx', 4)
console.log((await pptx.validate('Deck.pptx')).valid)
await pptx.toPdf('Deck.pptx', 'Deck.pdf')
```

Run the full acceptance deck:

```bash
node examples/pptx-acceptance.js            # writes into os.tmpdir()/r7-acceptance
node examples/pptx-acceptance.js ./out      # or wherever you like
```

**Measurements.** A bare number is EMU (the unit `r7_slide_read` returns), so a
value can be read, adjusted and written back unchanged. `"2cm"`, `"1in"`,
`"30px"` and `"24pt"` also work. `lineWidth` is the exception: a number below
100 is read as points, because "border: 1.5" means 1.5 pt to everyone who is
not holding a DrawingML specification.

**Colours.** `#RRGGBB`, `#AARRGGBB` and transparency as a separate 0..1 option.
`transparency` on a font and `fillTransparency` on a fill are deliberately
distinct, so a half-transparent caption colour cannot make the shape under it
see-through.

**Shapes.** `rectangle`, `rounded-rectangle`, `ellipse`, `circle`, `line`,
`arrow` (and `arrow-left/up/down/left-right`), `triangle`, `diamond`,
`pentagon`, `hexagon`, `octagon`, `star`, `chevron`, `plus`, `cloud`, `heart`,
`cylinder`, `cube`, `donut`, `pie`, `parallelogram`, `trapezoid` — or any
DrawingML preset name. `pptx.shapeCatalog()` lists them all.

---

## Typical agent flow

> "Take `Report.docx`, update section 3, keep the formatting, add a summary
> table and save a PDF."

```text
r7_inspect → r7_read → r7_replace (keeps formatting) → r7_table
           → r7_validate → r7_convert
```

Run it yourself:

```bash
node examples/report-scenario.js
```

Library use:

```javascript
import { DocxEngine, R7Adapter } from 'dsh-r7-office/r7'

const docx = new DocxEngine()
const adapter = new R7Adapter()

await docx.replaceText('Report.docx', 'draft text', 'approved text', {
  outputPath: 'Report_v2.docx'   // the original is never overwritten by default
})

await docx.table('Report_v2.docx', {
  action: 'create',
  rows: [['Objective', 'Timeline', 'Owner'], ['Deploy R7', 'Q2', 'IT']]
})

console.log((await docx.validate('Report_v2.docx')).valid)

await adapter.convert('Report_v2.docx', 'Report_v2.pdf')
```

---

## Security

`r7_desktop_exec` can execute code inside the user's open editor. Raw
DocScript is therefore **disabled by default**; only a fixed allowlist of safe
argument-driven commands runs in production. Enable raw execution only
deliberately:

```yaml
- id: r7-office
  name: 'dsh-r7-office'
  config:
    developerMode: true        # or export DSH_R7_DEVELOPER_MODE=1
```

`r7_desktop_status` always reports the effective mode. The desktop bridge binds
to `127.0.0.1` only.

See [SECURITY.md](SECURITY.md) for the full threat model and how to report a
vulnerability.

---

## Known limitations

- **Presentations: building layouts or masters is out of scope.** Slides are
  created on the layouts the deck already contains; a deck that ships none gets
  the generic title/body pair.
- **Presentations: SmartArt and charts are preserved, never authored.** The
  reader reports them (`type: "chart"`, `type: "graphicFrame"`) and every part
  behind them stays byte-identical through any edit, but the engine cannot
  create one. A hand-written SmartArt frame is just a `dgm:relIds` reference,
  and R7's own renderer dereferences the diagram parts behind it.
- **Presentations: image replacement stores a fresh media part** when the
  format changes or the media is shared, so the old part can be left
  unreferenced. Everything that still points at it keeps working.
- **Presentations: theme colours read back as tokens** (`scheme:accent1`),
  because no literal hex exists until the theme is resolved. Writing accepts
  literals only.
- **Live desktop bridge: Windows verified.** The bridge plugin itself is
  platform-neutral, but the automated live test drives R7-Office Desktop
  through the CEF DevTools protocol and is only verified against the Windows
  build (`Editors-2026.3.1`). Other platforms should work; they are untested.
- **R7 required for PDF.** `r7_convert` needs a local R7 installation for its
  `x2t` converter. Everything else works without it.
- **Format coverage.** Reading and editing cover the common OOXML surface
  listed above. Charts, embedded objects, SmartArt and tracked changes are
  preserved but not editable through these tools.
- **DOCX: not implemented in v0.1.0.** Replacing or resizing an *existing*
  image; assigning a table style (`tblStyle`); changing a hyperlink's target
  (retitle, or remove and insert instead); a convenience wrapper for per-section
  different headers (both can be read and preserved, and `r7_docx_header_footer`
  accepts `first`/`even` — for `first` also call `r7_docx_sections` with
  `titlePg`). Search is not revision-aware: `r7_replace` scans raw `<w:t>`, so
  text inside a tracked insertion is editable and paragraph indices count
  paragraphs inside deleted content — `<w:delText>` is never matched.
- **DOCX: comments, tracked changes, footnotes, endnotes, a table of contents,
  equations and embedded objects are preserved, never edited.** A test asserts
  they survive an ordinary edit with every part byte-identical.
- **DOCX to PDF: a thin horizontal line can appear through an italic subtitle.**
  A paragraph styled `Subtitle` (italic, centred, grey) may render with a faint
  line across it after `r7_convert`. The document contains no strikethrough,
  underline or paragraph border — `w:strike`, `w:u` and `w:pBdr` are all absent,
  and the same style applied to R7's own template renders cleanly through the
  same converter — so the line is introduced during PDF conversion and cannot
  be corrected from the document. It is cosmetic: the text itself is correct and
  extracts normally. Quantified at roughly 5 % more ink in that text band.
- **PDF rendering is verified visually, not by text extraction.** A PDF can
  carry the right text, the right page count and a clean `validate()` while
  drawing the wrong glyphs or a solid block. `scripts/visual-acceptance.py`
  renders the output with two independent engines and compares against a
  LibreOffice render of the source; run it when changing anything in the
  conversion path.
- **XLSX: not implemented in v0.1.0.** Cell styles are limited to the
  properties listed for `r7_sheet_format`; conditional formatting, data
  validation, charts, pivot tables and defined names are preserved but not
  editable.
- **PPTX: SmartArt, charts, animations and transitions are preserved, never
  authored.** The engine cannot create SmartArt (a hand-written frame is only a
  `dgm:relIds` reference), and editing their content is out of scope.
- **No concurrent editing.** File tools operate on a document at rest. For a
  document the user has open, use the desktop bridge; the file tools assume
  the file is not locked by another process.

---

## Development

```bash
npm test              # everything
npm run test:unit
npm run test:integration
npm run test:e2e
npm run test:live     # needs R7-Office installed
npm run test:harness  # boots a fresh DeepSeek Harness
npm run inspect:r7    # CDP probe of a running R7 Desktop
```

See [docs/development.md](docs/development.md) and
[docs/roadmap.md](docs/roadmap.md).

---

## License

MIT — see [LICENSE](LICENSE).

R7-Office is proprietary software owned by АО «Р7». This project redistributes
none of it.
