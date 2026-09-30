# dsh-r7-office

[![CI](https://github.com/OWNER/dsh-r7-office/actions/workflows/ci.yml/badge.svg)](https://github.com/OWNER/dsh-r7-office/actions/workflows/ci.yml)
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
git clone https://github.com/OWNER/dsh-r7-office.git
cd dsh-r7-office

# 2. Install the optional dev dependencies (test-only: the MCP client SDK)
npm install

# 3. Verify the checkout — no R7-Office required for this step
npm test
```

`npm test` runs 155 tests: pure unit tests, OOXML round-trip regression tests,
file end-to-end workflows, security-policy tests and an external MCP client
smoke suite. Tests that need an R7-Office installation **skip themselves** with
a clear message instead of failing, so a clean machine gets a green run.

Then verify the integration you actually intend to use:

```bash
# R7-Office file pipeline (author → edit → validate → PDF). Skips if R7 is absent.
npm run test:live

# DeepSeek Harness plugin activation. Boots a fresh Harness and asserts the
# 18 r7_* tools reached the tool registry.
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
[r7-office] зарегистрировано инструментов: 18 (r7_inspect, r7_read, ...); desktop bridge port=7888, developerMode=false
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
| `r7_table` | Create, inspect or update tables and cells |
| `r7_sheet_read` | Read values and formulas from a sheet or range (e.g. `A1:D10`) |
| `r7_sheet_write` | Write cells or a 2-D matrix, addressed by sheet name or index |
| `r7_sheet_add` | Add a worksheet to an existing workbook; other sheets are untouched |
| `r7_sheet_formula` | Insert or update a formula |
| `r7_slide_create` | Create a deck, or **append** a slide to an existing one without altering the slides already there |
| `r7_slide_edit` | Edit slide titles and text frames |
| `r7_convert` | Convert via the R7 `x2t` engine (PDF, HTML, TXT, DOCX, XLSX, PPTX) |
| `r7_validate` | Check package integrity and XML health |
| `r7_desktop_status` | Desktop bridge connection and effective security mode |
| `r7_desktop_selection` | Read or replace the selection in the open editor |
| `r7_desktop_exec` | Run a safe editor command, or raw DocScript in developer mode |

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

- **Live desktop bridge: Windows verified.** The bridge plugin itself is
  platform-neutral, but the automated live test drives R7-Office Desktop
  through the CEF DevTools protocol and is only verified against the Windows
  build (`Editors-2026.3.1`). Other platforms should work; they are untested.
- **R7 required for PDF.** `r7_convert` needs a local R7 installation for its
  `x2t` converter. Everything else works without it.
- **Format coverage.** Reading and editing cover the common OOXML surface
  listed above. Charts, embedded objects, SmartArt and tracked changes are
  preserved but not editable through these tools.
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
