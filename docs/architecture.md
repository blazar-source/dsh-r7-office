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
| `r7_inspect` | Inspection | Document hierarchy, metadata, sections, tables, sheet and slide lists, headers/footers, images, hyperlinks |
| `r7_read` | Reading | Structured text, markdown preview, cell ranges; with `includeStyles` a normalized formatting model instead of raw OOXML |
| `r7_create` | Generation | Creates fresh DOCX, XLSX, or PPTX documents. Refuses to replace an existing file unless `overwrite` is set; for XLSX every `sheets[]` entry and its name is honoured |
| `r7_edit` | DOCX Modification | Modifies, replaces, or deletes paragraphs preserving original formatting runs |
| `r7_replace` | DOCX Text | Finds and replaces text while keeping font, color, bold/italic intact |
| `r7_insert` | DOCX Insertion | Inserts paragraphs, headings, bullet lists, or page breaks at designated positions |
| `r7_table` | DOCX Tables | Creates, updates, or inspects tables: cells, merge/unmerge, borders, shading, column widths, row and column add/remove |
| `r7_docx_formatting` | DOCX Formatting | Normalized formatting of paragraphs, runs and cells: style, font family/size/color, alignment, indents, spacing, line spacing, lists |
| `r7_docx_sections` | DOCX Structure | Page size, orientation, margins, columns, page breaks and section breaks |
| `r7_docx_header_footer` | DOCX Structure | Lists, reads and edits headers and footers; page-number fields survive |
| `r7_docx_image` | DOCX Media | Inserts PNG/JPEG/GIF at a requested size, preserving the aspect ratio |
| `r7_docx_hyperlink` | DOCX Links | Lists, inserts, retitles and removes hyperlinks and their relationships |
| `r7_sheet_read` | XLSX Data | Reads values, formulas or normalized formatting from a sheet or range, addressed by name or index |
| `r7_sheet_write` | XLSX Data | Writes values, real date serials and formulas into a sheet or cell range |
| `r7_sheet_format` | XLSX Formatting | Formats a cell or range: font, fill, borders, alignment, wrap, number formats, merge, column width, row height |
| `r7_sheet_add` | XLSX Structure | Adds a worksheet to an existing workbook, registering every required package part; existing sheets are never rewritten |
| `r7_sheet_formula`| XLSX Math | Inserts or updates a formula in a cell |
| `r7_slide_read` | PPTX Reading | Normalized slide model: object id, type, geometry, rotation, z-order, text, font, fill, stroke, alignment, paragraphs, with layout/master inheritance resolved |
| `r7_slide_create` | PPTX Structure | Creates a deck, or appends a slide built on one of the deck's own layouts; existing slides are not modified |
| `r7_slide_format` | PPTX Content | Formats or repositions one object in place: font, geometry, fill, border, alignment, lists, spacing, text |
| `r7_slide_edit` | PPTX Structure | Duplicates, moves, reorders or deletes slides |
| `r7_slide_object` | PPTX Content | Adds or removes shapes, text boxes and images |
| `r7_convert` | Conversion | Converts between DOCX/XLSX/PPTX and PDF/HTML/TXT via R7 `x2t` converter |
| `r7_validate` | Integrity | Validates document package integrity, XML schema validity, and repair checks |
| `r7_desktop_status` | Desktop Bridge | Reports the bridge connection and the effective security mode |
| `r7_desktop_selection` | Desktop Bridge | Reads or replaces the selected text in the active R7 Desktop editor |
| `r7_desktop_exec` | Desktop Bridge | Runs an allowlisted safe command, or arbitrary DocScript when developer mode is explicitly enabled |

### 3.1 Where formatting lives

Formatting is exposed where it can be read back before it is written, because a
write without a read is how a document gets repainted:

- **DOCX** — `r7_docx_formatting` reads paragraph, run and cell formatting;
  `r7_edit` and `r7_insert` set paragraph-level style; `r7_table` sets cell
  shading, borders, alignment and widths. Text-level formatting is *preserved*
  by `r7_replace` rather than rewritten.
- **XLSX** — `r7_sheet_read({includeStyles})` reads the normalized model;
  `r7_sheet_format` writes font, fill, borders, alignment, number formats,
  merge, column width and row height.
- **PPTX** — `r7_slide_read` reads every object's geometry and typography with
  layout inheritance resolved; `r7_slide_format` patches an existing object and
  `r7_slide_object` adds one.

Every write is a *patch*: only the properties named are changed, and a property
that is not named keeps whatever the author set.

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
