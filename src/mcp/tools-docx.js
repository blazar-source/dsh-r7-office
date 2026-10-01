import { defaultOutput } from './tool-output.js'

/**
 * DOCX tool definitions beyond the format-agnostic inspect/read/create.
 *
 * @param {import('../r7/docx.js').DocxEngine} docxEngine
 * @returns {Array<object>}
 */
export function buildDocxTools(docxEngine) {
  return [
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
      description: 'Insert a paragraph, heading, bullet item, or page break at the start, the end, or relative to another paragraph.',
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
    }
  ]
}
