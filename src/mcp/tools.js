import path from 'node:path'
import fs from 'node:fs'
import { DocxEngine } from '../r7/docx.js'
import { XlsxEngine } from '../r7/xlsx.js'
import { PptxEngine } from '../r7/pptx.js'
import { R7Adapter } from '../r7/adapter.js'
import { DesktopBridge, SAFE_COMMANDS } from './desktop-bridge.js'
import { buildDocxTools } from './tools-docx.js'
import { buildXlsxTools } from './tools-xlsx.js'
import { buildPptxTools } from './tools-pptx.js'
import { defaultOutput } from './tool-output.js'

/**
 * Builds all R7 Office MCP tools and connects them to the underlying engines.
 */
export function buildR7Tools(options = {}) {
  const adapter = new R7Adapter(options.r7Path)
  const docxEngine = new DocxEngine(adapter)
  const xlsxEngine = new XlsxEngine(adapter)
  const pptxEngine = new PptxEngine(adapter)
  const desktopBridge = options.desktopBridge || new DesktopBridge(options)

  function getEngineByExt(filePath) {
    const ext = path.extname(filePath).toLowerCase()
    if (ext === '.docx' || ext === '.docxf' || ext === '.doc') return docxEngine
    if (ext === '.xlsx' || ext === '.xlsm' || ext === '.xls') return xlsxEngine
    if (ext === '.pptx' || ext === '.pptm' || ext === '.ppt') return pptxEngine
    return null
  }

  return [
    {
      name: 'r7_inspect',
      description: 'Inspect document structure: outline, metadata, headings, tables, sheet list, slide list, sections, headers/footers, images and hyperlinks.',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to the document file (DOCX, XLSX, PPTX).' },
          includeStyles: { type: 'boolean', description: 'DOCX: also report the style definitions in use.' },
          includeSections: { type: 'boolean', description: 'DOCX: also report page size, orientation, margins and section breaks.' },
          includeHeadersFooters: { type: 'boolean', description: 'DOCX: also report header and footer parts.' },
          includeHyperlinks: { type: 'boolean', description: 'DOCX: also report hyperlinks and their targets.' },
          includeImages: { type: 'boolean', description: 'DOCX: also report embedded images.' },
          includeLists: { type: 'boolean', description: 'DOCX: also report numbering definitions.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        const engine = getEngineByExt(args.filePath)
        if (!engine) throw new Error(`Unsupported document extension: ${path.extname(args.filePath)}`)
        // Forward every option: the DOCX engine's extra reporting is opt-in, so
        // dropping the arguments here would silently hide it from an agent.
        return await engine.inspect(args.filePath, args)
      }
    },

    {
      name: 'r7_read',
      description: 'Read structured text, filtered paragraphs, markdown preview, spreadsheet ranges or slides, optionally with normalized formatting.',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to the document file.' },
          format: { type: 'string', enum: ['structured', 'markdown', 'text'], description: 'Output text representation (default: structured).' },
          fromParagraph: { type: 'integer', description: '0-based starting paragraph index for DOCX.' },
          count: { type: 'integer', description: 'Max paragraphs to return for DOCX.' },
          query: { type: 'string', description: 'Filter text by substring.' },
          includeStyles: {
            type: 'boolean',
            description: 'Return normalized formatting instead of raw OOXML. DOCX: a "formatting" block per paragraph (style, font family/size, bold, italic, underline, colour, alignment, indents, spacing before/after, line spacing, list/numbering). XLSX: a "styles" matrix plus merged ranges, row heights and column widths. Read this before changing a document.'
          },
          includeFormulas: { type: 'boolean', description: 'XLSX: also return the formulas matrix.' },
          sheetIndex: { type: 'integer', description: '0-based sheet index for XLSX.' },
          sheetName: { type: 'string', description: 'Sheet name for XLSX.' },
          range: { type: 'string', description: 'Cell range for XLSX (e.g. A1:D10).' },
          slideIndex: { type: 'integer', description: '0-based slide index for PPTX.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        const engine = getEngineByExt(args.filePath)
        if (!engine) throw new Error(`Unsupported document extension: ${path.extname(args.filePath)}`)
        return await engine.read(args.filePath, args)
      }
    },

    {
      name: 'r7_create',
      description: 'Create a NEW DOCX, XLSX or PPTX document. Refuses to replace an existing file unless overwrite is true.',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Output destination path (.docx, .xlsx, or .pptx).' },
          title: { type: 'string', description: 'Document title (DOCX/PPTX).' },
          paragraphs: {
            type: 'array',
            description: 'DOCX: initial paragraphs (strings, or objects with text/style/bold/italic/align).',
            items: { type: 'string' }
          },
          tables: {
            type: 'array',
            description: 'DOCX: list of tables, each with a 2D array of rows and cells.'
          },
          sheets: {
            type: 'array',
            description: 'XLSX: every worksheet to create, each { name, data } with a 2D data matrix. All entries and their names are honoured.',
            items: {
              type: 'object',
              properties: {
                name: { type: 'string', description: 'Worksheet name.' },
                data: { type: 'array', description: '2D matrix of cell values.' }
              }
            }
          },
          overwrite: { type: 'boolean', description: 'Set true to replace an existing file at filePath. Default false.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        const engine = getEngineByExt(args.filePath)
        if (!engine) throw new Error(`Unsupported document extension: ${path.extname(args.filePath)}`)
        return await engine.create(args.filePath, args)
      }
    },

    ...buildDocxTools(docxEngine),

    ...buildXlsxTools(xlsxEngine),

    ...buildPptxTools(pptxEngine),

    {
      name: 'r7_convert',
      description: 'Convert document to PDF, HTML, TXT, DOCX, XLSX, or PPTX using R7 native converter (x2t). XLSX to PDF exports EVERY worksheet by default, in tab order; pass sheetName (or allSheets: false) to export one sheet only.',
      parameters: {
        type: 'object',
        required: ['sourcePath', 'targetPath'],
        properties: {
          sourcePath: { type: 'string', description: 'Absolute or workspace-relative path to source document.' },
          targetPath: { type: 'string', description: 'Target destination file path (e.g. document.pdf, report.html).' },
          allSheets: { type: 'boolean', description: 'XLSX to PDF only: export every worksheet, in tab order. Default true. Set false to export just the active (first) sheet.' },
          sheetName: { type: 'string', description: 'XLSX to PDF only: export just this worksheet. Takes precedence over allSheets.' }
        }
      },
      output: defaultOutput(),
      async execute({ sourcePath, targetPath, allSheets, sheetName }) {
        // x2t renders one worksheet per run (whichever tab the workbook marks
        // active), so a spreadsheet render is driven per sheet by the XLSX
        // engine; every other conversion stays a plain x2t call.
        const sourceExt = path.extname(sourcePath).toLowerCase()
        const targetExt = path.extname(targetPath).toLowerCase()
        if (targetExt === '.pdf' && ['.xlsx', '.xlsm', '.xls'].includes(sourceExt)) {
          return await xlsxEngine.exportPdf(sourcePath, { outputPath: targetPath, allSheets, sheetName })
        }
        return await adapter.convert(sourcePath, targetPath)
      }
    },

    {
      name: 'r7_validate',
      description: 'Validate document package structure, XML integrity, and formatting health.',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to document file.' }
        }
      },
      output: defaultOutput(),
      async execute({ filePath }) {
        const engine = getEngineByExt(filePath)
        if (!engine) throw new Error(`Unsupported document extension: ${path.extname(filePath)}`)
        return await engine.validate(filePath)
      }
    },

    {
      name: 'r7_desktop_status',
      description: 'Check live connection status and security policy for open R7-Office Desktop windows via the local bridge.',
      parameters: {
        type: 'object',
        properties: {}
      },
      output: defaultOutput(),
      async execute() {
        return desktopBridge.getStatus()
      }
    },

    {
      name: 'r7_desktop_selection',
      description: 'Get or replace selected text in the currently active R7-Office Desktop editor window.',
      parameters: {
        type: 'object',
        required: ['action'],
        properties: {
          action: { type: 'string', enum: ['get', 'replace'], description: 'Get or replace active selection.' },
          text: { type: 'string', description: 'Text to paste when action is replace.' }
        }
      },
      output: defaultOutput(),
      async execute({ action, text }) {
        if (action === 'get') {
          return await desktopBridge.execute('getSelection')
        }
        if (action === 'replace') {
          return await desktopBridge.execute('replaceSelection', { text: text || '' })
        }
        throw new Error(`Invalid desktop selection action: ${action}`)
      }
    },

    {
      name: 'r7_desktop_exec',
      description: 'Execute safe built-in command or custom DocScript in live opened document in R7 Desktop. Arbitrary JS requires developerMode enabled.',
      parameters: {
        type: 'object',
        properties: {
          safeCommand: {
            type: 'string',
            enum: ['addParagraph', 'insertTable', 'setSelectedText', 'getSelectedText', 'getDocumentText', 'saveDocument', 'searchAndReplace'],
            description: 'Safe pre-validated command permitted in production mode.'
          },
          args: {
            type: 'object',
            description: 'Arguments for the safeCommand (e.g. { text: "...", style: "Heading1" }).'
          },
          code: {
            type: 'string',
            description: 'Arbitrary DocScript JavaScript code (ONLY permitted when developerMode: true is explicitly configured).'
          }
        }
      },
      output: defaultOutput(),
      async execute({ safeCommand, args = {}, code }) {
        if (safeCommand) {
          if (!SAFE_COMMANDS.has(safeCommand)) {
            throw new Error(`Unknown safeCommand: ${safeCommand}. Allowed: ${Array.from(SAFE_COMMANDS).join(', ')}`)
          }
          return await desktopBridge.execute('safeCommand', { command: safeCommand, args })
        }

        if (code) {
          return await desktopBridge.execute('callCommand', { code })
        }

        throw new Error('Must provide either safeCommand or code parameter.')
      }
    },

    {
      name: 'r7_pdf_inspect',
      description: 'Measure the text fidelity of a PDF: page and font inventory, how many glyphs are really drawn by text operators, how many glyph fills painted nothing, whether the text is extractable, and a verdict (text/outlined/mixed). Use it to prove a PDF export kept its text instead of silently losing it; set repair to fix malformed ToUnicode CMap entry counts that make R7-exported body text copy out as garbage.',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to the PDF file to inspect.' },
          repair: { type: 'boolean', description: 'Repair malformed ToUnicode CMap entry counts in place (rewrites filePath) before reporting. Default false.' }
        }
      },
      output: defaultOutput(),
      async execute({ filePath, repair = false }) {
        // Imported lazily so this entry is a pure append to this shared module.
        const { inspectPdf, repairToUnicodeCMaps } = await import('../r7/pdf-inspect.js')
        const repairs = repair ? repairToUnicodeCMaps(filePath).repairs : []
        return { ...inspectPdf(filePath), repairs }
      }
    }
  ]
}
