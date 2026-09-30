# dsh-r7-office

[![CI](https://github.com/deepseek-ai/dsh-r7-office/actions/workflows/ci.yml/badge.svg)](https://github.com/deepseek-ai/dsh-r7-office/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![Node.js](https://img.shields.io/badge/node-%3E%3D20.0.0-brightgreen.svg)](https://nodejs.org/)

**R7-Office (Р7-Офис) document processing plugin & Model Context Protocol (MCP) server for DeepSeek Harness.**

Enables AI agents to inspect, read, generate, format, edit, and convert `DOCX`, `XLSX`, `PPTX`, and `PDF` documents directly through official R7-Office engines and pure Open Packaging Conventions (OPC) standards without keyboard/mouse emulation.

---

## 🇷🇺 Документация на русском
Полная документация на русском языке доступна в файле [README.ru.md](README.ru.md).

---

## Key Features

- **Format-Preserving Editing**: Text replacements and edits preserve original XML styles, fonts, colors, and paragraph numbering hierarchies.
- **Full Office Formats Coverage**:
  - **DOCX**: Headings, paragraphs, bullet lists, tables, cell borders/shading, page breaks, metadata.
  - **XLSX**: Worksheets, cell ranges (`A1:D10`), numbers, strings, formulas (`SUM`, arithmetic, etc.).
  - **PPTX**: Slides, titles, text frames, shape content.
  - **PDF**: 100% fidelity document conversion powered by local R7 `x2t` engine.
- **Safety & Validation**: Built-in `r7_validate` checks package integrity and schema compliance before and after modifications.
- **Dual Runtime Support**:
  - **DeepSeek Harness Plugin**: Mounts `r7_*` tools into `ctx.tools` and provides agent guidance.
  - **Standalone MCP Server**: Exposes standard JSON-RPC 2.0 tools for Claude Code, Cursor, Windsurf, or any MCP client.
- **R7 Desktop Live Bridge**: Connects directly to open R7 Desktop editor windows via local WebSocket to read selections or execute DocScript commands in real-time.

---

## Architecture

```text
DeepSeek Harness → dsh-r7-office Plugin → MCP Server → R7 Engines & x2t → DOCX/XLSX/PPTX/PDF
                                                   └─→ Desktop Bridge ──→ Open R7 Desktop Window
```

See [docs/architecture.md](docs/architecture.md) for full architectural documentation and [docs/decisions/](docs/decisions/) for Architecture Decision Records (ADRs).

---

## Installation

### 1. As a DeepSeek Harness Plugin

Add `dsh-r7-office` to your `cordis.patch.yml`:

```yaml
- insert:
    - id: r7-office
      name: 'dsh-r7-office'
```

Or install via DeepSeek Harness Plugin Manager:

```bash
dsh plugin install dsh-r7-office
```

### 2. Standalone MCP Server (Claude Code, Cursor, etc.)

Add to your MCP client configuration (e.g. `claude_desktop_config.json` or `.cursor/mcp.json`):

```json
{
  "mcpServers": {
    "r7-office": {
      "command": "node",
      "args": ["/path/to/dsh-r7-office/src/mcp/cli.js"]
    }
  }
}
```

---

## MCP Tools Reference

| Tool | Description |
|---|---|
| `r7_inspect` | Inspects document hierarchy, outline, headings, tables, sheets, and slides |
| `r7_read` | Reads structured paragraphs, markdown representation, or cell ranges |
| `r7_create` | Creates fresh DOCX/XLSX/PPTX documents from native R7 templates |
| `r7_edit` | Edits, updates text/style, or removes specific paragraphs in DOCX |
| `r7_replace` | Searches and replaces text in DOCX while strictly preserving all formatting |
| `r7_insert` | Inserts paragraphs, headings, lists, or page breaks at specified positions |
| `r7_table` | Creates, modifies, or inspects tables and cell contents |
| `r7_sheet_read` | Reads cell values and formulas from XLSX worksheets |
| `r7_sheet_write` | Writes individual cells or 2D data matrices to XLSX worksheets |
| `r7_sheet_formula` | Inserts and updates spreadsheet formulas in XLSX |
| `r7_slide_create` | Creates new slides in PPTX presentations |
| `r7_slide_edit` | Modifies titles and text frames on PPTX slides |
| `r7_convert` | Converts documents to PDF, HTML, TXT, DOCX, XLSX via R7 `x2t` engine |
| `r7_validate` | Validates document package integrity and XML health |
| `r7_desktop_status` | Checks connection status with active R7 Desktop window |
| `r7_desktop_selection`| Reads or replaces selected text in active R7 Desktop editor |
| `r7_desktop_exec` | Executes R7 DocScript / DocumentBuilder JS commands in live editor |

---

## Example Usage

### Typical Agent User Prompt
> "Take `Report.docx`, update section 3 with the new 2026 strategic plan, preserve original styling, add a summary table, validate the document, and export to PDF."

### Programmatic Scenario Flow
```javascript
import { DocxEngine, R7Adapter } from 'dsh-r7-office/r7'

const docx = new DocxEngine()
const adapter = new R7Adapter()

// 1. Inspect
const outline = await docx.inspect('Report.docx')

// 2. Replace text preserving formatting
await docx.replaceText(
  'Report.docx',
  'Draft section 3 text',
  'Approved strategic plan for 2026-2028.',
  { outputPath: 'Report_v2.docx' }
)

// 3. Add table
await docx.table('Report_v2.docx', {
  action: 'create',
  rows: [
    ['Objective', 'Timeline', 'Owner'],
    ['Deploy R7', 'Q2', 'IT Team']
  ]
})

// 4. Validate
const check = await docx.validate('Report_v2.docx')
console.log('Valid:', check.valid)

// 5. Convert to PDF
await adapter.convert('Report_v2.docx', 'Report_v2.pdf')
```

---

## Running Tests

```bash
# Run all unit, integration, and e2e test suites
npm test

# Run specific suite
npm run test:unit
npm run test:integration
npm run test:e2e
```

---

## License

MIT © DSH R7-Office Contributors
