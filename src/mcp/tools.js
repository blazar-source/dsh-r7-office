import path from 'node:path'
import fs from 'node:fs'
import { DocxEngine } from '../r7/docx.js'
import { XlsxEngine } from '../r7/xlsx.js'
import { PptxEngine } from '../r7/pptx.js'
import { R7Adapter } from '../r7/adapter.js'
import { DesktopBridge, SAFE_COMMANDS } from './desktop-bridge.js'

function defaultOutput() {
  return {
    schema: {
      type: 'object',
      additionalProperties: true,
      properties: {
        ok: { type: 'boolean' },
        result: { type: 'string' }
      }
    },
    render: (_args, val) => [{
      type: 'text',
      text: typeof val === 'string' ? val : JSON.stringify(val, null, 2)
    }]
  }
}

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
      description: 'Inspect structure, outline, metadata, tables, sheets, and slides of an office document (DOCX, XLSX, PPTX).',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to the document file (DOCX, XLSX, PPTX).' }
        }
      },
      output: defaultOutput(),
      async execute({ filePath }) {
        const engine = getEngineByExt(filePath)
        if (!engine) throw new Error(`Unsupported document extension: ${path.extname(filePath)}`)
        return await engine.inspect(filePath)
      }
    },

    {
      name: 'r7_read',
      description: 'Read structured text, filtered paragraphs, markdown preview, spreadsheet ranges, or slide contents.',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to the document file.' },
          format: { type: 'string', enum: ['structured', 'markdown', 'text'], description: 'Output text representation (default: structured).' },
          fromParagraph: { type: 'integer', description: '0-based starting paragraph index for DOCX.' },
          count: { type: 'integer', description: 'Max paragraphs to return for DOCX.' },
          query: { type: 'string', description: 'Filter text by substring.' },
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

    {
      name: 'r7_edit',
      description: 'Edit, update text, style, or remove a specific paragraph in a DOCX document.',
      parameters: {
        type: 'object',
        required: ['filePath', 'paragraphIndex'],
        properties: {
          filePath: { type: 'string', description: 'Path to the DOCX document.' },
          paragraphIndex: { type: 'integer', description: '0-based index of the target paragraph (from r7_inspect/r7_read).' },
          newText: { type: 'string', description: 'New text content for the paragraph.' },
          style: { type: 'string', description: 'Paragraph style name (e.g. Normal, Heading1, Heading2).' },
          deleteParagraph: { type: 'boolean', description: 'Set true to delete this paragraph.' },
          outputPath: { type: 'string', description: 'Optional target path (defaults to overwriting in-place).' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        return await docxEngine.editParagraph(args.filePath, args.paragraphIndex, args.newText || '', args)
      }
    },

    {
      name: 'r7_replace',
      description: 'Search and replace text in DOCX while strictly preserving all existing run formatting, fonts, colors, and styles.',
      parameters: {
        type: 'object',
        required: ['filePath', 'search', 'replace'],
        properties: {
          filePath: { type: 'string', description: 'Path to the DOCX file.' },
          search: { type: 'string', description: 'Text substring or pattern to find.' },
          replace: { type: 'string', description: 'Replacement text.' },
          matchCase: { type: 'boolean', description: 'Case sensitive matching (default true).' },
          replaceAll: { type: 'boolean', description: 'Replace all occurrences (default true).' },
          outputPath: { type: 'string', description: 'Optional target path.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        return await docxEngine.replaceText(args.filePath, args.search, args.replace, args)
      }
    },

    {
      name: 'r7_insert',
      description: 'Insert paragraph, heading, bullet item, or page break at start, end, or relative to another paragraph.',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to the DOCX file.' },
          text: { type: 'string', description: 'Text to insert.' },
          style: { type: 'string', description: 'Style name (default Normal).' },
          isHeading: { type: 'boolean', description: 'Whether this is a heading.' },
          headingLevel: { type: 'integer', description: 'Heading level 1..6 (default 1).' },
          position: { type: 'string', enum: ['start', 'end', 'before', 'after'], description: 'Insertion anchor.' },
          targetIndex: { type: 'integer', description: '0-based index when position is before or after.' },
          pageBreak: { type: 'boolean', description: 'Insert a page break.' },
          outputPath: { type: 'string', description: 'Optional destination file path.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        return await docxEngine.insert(args.filePath, args)
      }
    },

    {
      name: 'r7_table',
      description: 'Create, inspect, or modify tables in DOCX documents (insert rows, update cell text, format borders).',
      parameters: {
        type: 'object',
        required: ['filePath', 'action'],
        properties: {
          filePath: { type: 'string', description: 'Path to the DOCX file.' },
          action: { type: 'string', enum: ['create', 'addRow', 'setCell', 'inspect'], description: 'Table action.' },
          tableIndex: { type: 'integer', description: '0-based table index for existing tables.' },
          rows: { type: 'array', description: '2D array of rows and cell values for create/addRow.' },
          cell: {
            type: 'object',
            properties: {
              row: { type: 'integer' },
              col: { type: 'integer' },
              value: { type: 'string' }
            },
            description: 'Cell coordinate and value for setCell action.'
          },
          position: { type: 'string', enum: ['start', 'end'], description: 'Position when creating a table.' },
          outputPath: { type: 'string', description: 'Optional target path.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        return await docxEngine.table(args.filePath, args)
      }
    },

    {
      name: 'r7_sheet_read',
      description: 'Read spreadsheet values and formulas from an XLSX worksheet or range (e.g. A1:D10).',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to XLSX file.' },
          sheetIndex: { type: 'integer', description: '0-based sheet index.' },
          sheetName: { type: 'string', description: 'Sheet name (takes precedence over sheetIndex).' },
          range: { type: 'string', description: 'Range reference (e.g. A1:C10).' },
          includeFormulas: {
            type: 'boolean',
            description: 'Also return the formulas matrix. Use it to verify a formula was written: a formula cell has no cached value until a spreadsheet engine recalculates the workbook.'
          }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        return await xlsxEngine.read(args.filePath, args)
      }
    },

    {
      name: 'r7_sheet_write',
      description: 'Write values or a 2D matrix into an XLSX worksheet, addressed by index or name.',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to XLSX file.' },
          sheetIndex: { type: 'integer', description: '0-based sheet index.' },
          sheetName: { type: 'string', description: 'Sheet name (takes precedence over sheetIndex).' },
          startCell: { type: 'string', description: 'Starting cell (e.g. A1).' },
          matrix: { type: 'array', description: '2D matrix of values to write.' },
          cells: {
            type: 'array',
            description: 'List of individual cell updates [{ ref: "A1", value: 123, formula: "=SUM(B1:B5)" }].'
          },
          outputPath: { type: 'string', description: 'Optional target path.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        return await xlsxEngine.write(args.filePath, args)
      }
    },

    {
      name: 'r7_sheet_add',
      description: 'Add a new worksheet to an existing XLSX workbook, optionally with initial data. Other sheets are left untouched.',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to an existing XLSX file.' },
          name: { type: 'string', description: 'New worksheet name (defaults to ЛистN; made unique automatically).' },
          index: { type: 'integer', description: 'Position in the tab order; appended when omitted.' },
          data: { type: 'array', description: 'Optional 2D matrix to write starting at A1.' },
          outputPath: { type: 'string', description: 'Optional target path.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        return await xlsxEngine.addSheet(args.filePath, args)
      }
    },

    {
      name: 'r7_sheet_formula',
      description: 'Insert or update a formula in an XLSX spreadsheet cell.',
      parameters: {
        type: 'object',
        required: ['filePath', 'cell', 'formula'],
        properties: {
          filePath: { type: 'string', description: 'Path to XLSX file.' },
          cell: { type: 'string', description: 'Target cell reference (e.g. D2).' },
          formula: { type: 'string', description: 'Excel formula string (e.g. =SUM(A2:C2) or =B2*1.2).' },
          sheetIndex: { type: 'integer', description: '0-based sheet index.' },
          sheetName: { type: 'string', description: 'Sheet name (takes precedence over sheetIndex).' },
          outputPath: { type: 'string', description: 'Optional target path.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        return await xlsxEngine.write(args.filePath, {
          outputPath: args.outputPath,
          sheetIndex: args.sheetIndex,
          sheetName: args.sheetName,
          cells: [{ ref: args.cell, formula: args.formula }]
        })
      }
    },

    {
      name: 'r7_slide_create',
      description: 'Create a new PPTX deck, or append a slide to an existing one. Appending never modifies the slides already present.',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to a PPTX file. A missing file is created; an existing one gets a new slide.' },
          title: { type: 'string', description: 'Title text for the new slide (or for slide 1 when creating a deck).' },
          baseSlideIndex: { type: 'integer', description: 'When appending: 0-based slide to clone the layout from (default 0).' },
          outputPath: { type: 'string', description: 'Optional target path.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        // Appending to an existing deck must never rebuild it from a template:
        // that is data loss, not creation.
        if (fs.existsSync(args.filePath)) {
          return await pptxEngine.addSlide(args.filePath, args)
        }
        return await pptxEngine.create(args.filePath, { title: args.title })
      }
    },

    {
      name: 'r7_slide_edit',
      description: 'Edit title and shape text in a PPTX slide.',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to PPTX file.' },
          slideIndex: { type: 'integer', description: '0-based slide index.' },
          title: { type: 'string', description: 'New slide title.' },
          search: { type: 'string', description: 'Search text on slide.' },
          replace: { type: 'string', description: 'Replacement text.' },
          outputPath: { type: 'string', description: 'Optional target path.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        return await pptxEngine.editSlide(args.filePath, args)
      }
    },

    {
      name: 'r7_convert',
      description: 'Convert document to PDF, HTML, TXT, DOCX, XLSX, or PPTX using R7 native converter (x2t).',
      parameters: {
        type: 'object',
        required: ['sourcePath', 'targetPath'],
        properties: {
          sourcePath: { type: 'string', description: 'Absolute or workspace-relative path to source document.' },
          targetPath: { type: 'string', description: 'Target destination file path (e.g. document.pdf, report.html).' }
        }
      },
      output: defaultOutput(),
      async execute({ sourcePath, targetPath }) {
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
    }
  ]
}
