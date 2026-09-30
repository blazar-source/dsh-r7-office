# Architecture & Technical Design: dsh-r7-office

## 1. System Overview

`dsh-r7-office` is a native DeepSeek Harness plugin and Model Context Protocol (MCP) server providing programmatic creation, reading, modification, formatting, and conversion of office documents (`DOCX`, `XLSX`, `PPTX`, `PDF`) using R7-Office (Р7-Офис) standards and engines.

```text
┌─────────────────────────────────────────────────────────────┐
│                 DeepSeek Harness (Cordis)                   │
│   ┌─────────────────────────────────────────────────────┐   │
│   │              dsh-r7-office (Plugin)                 │   │
│   │   - Registers r7_* tools on ctx.tools               │   │
│   │   - Injects agent guidance (systemPrompt)           │   │
│   └──────────────────────────┬──────────────────────────┘   │
└──────────────────────────────┼──────────────────────────────┘
                               │ JSON-RPC / MCP Transport
                               ▼
┌─────────────────────────────────────────────────────────────┐
│                      R7 MCP Server                          │
│                                                             │
│  [MCP Tools Interface (r7_*)]                               │
│  ├── File Operations Subsystem                              │
│  │   ├── DOCX Engine (XML/DOM preservation, paragraph/table)│
│  │   ├── XLSX Engine (Worksheets, cell ranges, formulas)    │
│  │   ├── PPTX Engine (Slide creation, shapes, text runs)    │
│  │   └── R7 x2t Engine Adapter (Native format conversion)   │
│  │                                                          │
│  └── Desktop Bridge Subsystem (Stage 2)                     │
│      ├── WebSocket Server / Bridge Controller               │
│      └── R7 Desktop sdkjs-plugins Integration               │
└──────────────────────────────┬──────────────────────────────┘
                               │ Local process / IPC / WS
                               ▼
┌─────────────────────────────────────────────────────────────┐
│                   Local Host Environment                    │
│   - R7-Office Desktop Editors (C:\Program Files\R7-Office\) │
│   - x2t converter binary (converter/x2t.exe)                │
│   - R7 Native blank templates (converter/empty/ru-RU)       │
│   - Desktop plugins directory (%LOCALAPPDATA%/R7-Office/...)│
└─────────────────────────────────────────────────────────────┘
```

## 2. Key Design Principles

1. **Format Preservation**: Modifications preserve existing XML schemas, styling, themes, fonts, and numbering hierarchies without destructive overwrites.
2. **Dynamic Host Discovery**: R7-Office installation is detected dynamically via Windows Registry, standard Program Files paths, and Linux/macOS standard directories. Zero hardcoded user paths or machine-specific assumptions.
3. **No Bundling of Proprietary Binaries**: Does not distribute proprietary R7 binaries (`x2t.exe`, DLLs). Uses installed local R7 tools when available, and pure standard OOXML engine with graceful fallbacks.
4. **Safety & Backup**: Operations enforce validation before and after modification. Original files are protected with backup mechanisms unless explicitly instructed to overwrite in-place.
5. **Dual Protocol Support**:
   - Direct Cordis Plugin (registers `r7_*` tools into Harness Host runtime)
   - Standard MCP Server (stdio and streamable-http JSON-RPC 2.0 interface for Claude Code, Cursor, Windsurf, or external MCP clients).

## 3. Tool Matrix

| Tool | Category | Description |
|---|---|---|
| `r7_inspect` | Inspection | Returns document hierarchy, metadata, sections, tables, sheets, slides |
| `r7_read` | Reading | Extracts structured text, paragraphs by index/selector, tables, markdown preview |
| `r7_create` | Generation | Creates fresh DOCX, XLSX, or PPTX documents from native templates |
| `r7_edit` | DOCX Modification | Modifies, replaces, or deletes paragraphs preserving original formatting runs |
| `r7_replace` | DOCX Text | Finds and replaces text patterns/regex while keeping font, color, bold/italic intact |
| `r7_insert` | DOCX Insertion | Inserts paragraphs, headings, bullet lists, or page breaks at designated positions |
| `r7_table` | DOCX Tables | Creates, updates, or inspects tables, row cells, borders and shading |
| `r7_sheet_read` | XLSX Data | Reads cell values, formulas, types from specific sheets and ranges (e.g. `A1:D10`) |
| `r7_sheet_write` | XLSX Data | Writes values, styles, numbers, dates to sheets and cell ranges |
| `r7_sheet_formula`| XLSX Math | Inserts or updates a formula in a cell |
| `r7_slide_create` | PPTX Layout | Appends new slides with selected layout templates |
| `r7_slide_edit` | PPTX Content | Edits text frames, titles, bullet points, and shape contents on slides |
| `r7_convert` | Conversion | Converts between DOCX/XLSX/PPTX and PDF/HTML/TXT via R7 `x2t` converter |
| `r7_validate` | Integrity | Validates document package integrity, XML schema validity, and repair checks |
| `r7_desktop_status` | Desktop Bridge | Reports the bridge connection and the effective security mode |
| `r7_desktop_selection` | Desktop Bridge | Reads or replaces the selected text in the active R7 Desktop editor |
| `r7_desktop_exec` | Desktop Bridge | Runs an allowlisted safe command, or arbitrary DocScript when developer mode is explicitly enabled |

### 3.1 Where formatting lives

There is no separate `r7_format` tool. Formatting is expressed where it is
unambiguous and cannot silently corrupt a document:

- paragraph-level formatting through `r7_edit` (`style`) and `r7_insert`
  (heading level, style),
- cell and table formatting through `r7_table` and `r7_sheet_write`,
- text-level formatting is **preserved** by `r7_replace` rather than rewritten,
  because rewriting runs is exactly how a find-and-replace loses fonts and
  colours.

A dedicated formatting tool was intentionally left out of the first release:
without a stable style-inspection story it invites destructive writes.

## 4. Technical Details of R7 Office Integration

### 4.1 R7 Detection & x2t Engine
R7-Office Desktop bundles `x2t.exe` inside its `converter/` subdirectory:
- Windows: `C:\Program Files\R7-Office\Editors-*\converter\x2t.exe` or `%LOCALAPPDATA%\Programs\R7-Office\...`
- Linux: `/opt/r7-office/desktopeditors/converter/x2t`
- macOS: `/Applications/R7-Office.app/Contents/Resources/converter/x2t`

`x2t.exe` performs fast conversion between Office Open XML and PDF/HTML with exact font rendering and layout metrics matching R7 Document Server.

### 4.2 Document XML Architecture
Documents are processed by unpacking the Open Packaging Conventions (OPC) zip container:
- `[Content_Types].xml`
- `_rels/.rels`
- `word/document.xml`, `word/styles.xml`, `word/numbering.xml`
- `xl/workbook.xml`, `xl/worksheets/sheet*.xml`, `xl/sharedStrings.xml`
- `ppt/presentation.xml`, `ppt/slides/slide*.xml`

Text editing operates at the `w:r` (run) and `w:p` (paragraph) level to avoid splitting style properties across XML elements during search and replace.
