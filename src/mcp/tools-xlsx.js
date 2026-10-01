import { defaultOutput } from './tool-output.js'

/**
 * XLSX tool definitions.
 *
 * Kept in its own module so the spreadsheet surface can evolve — and be
 * reviewed — without touching the DOCX or PPTX definitions.
 *
 * @param {import('../r7/xlsx.js').XlsxEngine} xlsxEngine
 * @returns {Array<object>}
 */
export function buildXlsxTools(xlsxEngine) {
  return [
    {
      name: 'r7_sheet_read',
      description: 'Read spreadsheet values, formulas and formatting from a worksheet or range (e.g. A1:D10).',
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
          },
          includeStyles: {
            type: 'boolean',
            description: 'Also return normalized formatting for every cell (font, fill, border, alignment, numberFormat) plus merged ranges, row heights and column widths.'
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
      description: 'Write values, dates or formulas into an XLSX worksheet, addressed by index or name.',
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
            description: 'Individual cell updates: [{ ref, value, formula?, numberFormat?, date? }]. Set date: true (or a date numberFormat) with an ISO 8601 value to store a real date serial instead of text, and numberFormat to apply an Excel number format to the cell.'
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
      name: 'r7_sheet_format',
      description: 'Apply formatting to a cell or range in an XLSX worksheet: font, fill, borders, alignment, number format, merge, column width and row height. Values, formulas and unrelated styles are preserved.',
      parameters: {
        type: 'object',
        required: ['filePath', 'range'],
        properties: {
          filePath: { type: 'string', description: 'Path to XLSX file.' },
          sheetIndex: { type: 'integer', description: '0-based sheet index.' },
          sheetName: { type: 'string', description: 'Sheet name (takes precedence over sheetIndex).' },
          range: { type: 'string', description: 'Cell or range to format, e.g. "B2" or "A1:D10".' },
          font: {
            type: 'object',
            description: 'Font settings.',
            properties: {
              family: { type: 'string', description: 'Font name, e.g. "Liberation Sans".' },
              size: { type: 'number', description: 'Point size, e.g. 12.' },
              bold: { type: 'boolean' },
              italic: { type: 'boolean' },
              underline: { type: 'boolean', description: 'true, or a style such as "double".' },
              strike: { type: 'boolean' },
              color: { type: 'string', description: 'Font colour, e.g. "#FF0000".' }
            }
          },
          fill: {
            type: 'object',
            description: 'Cell background.',
            properties: {
              color: { type: 'string', description: 'Background colour, e.g. "#FFF2CC".' }
            }
          },
          border: {
            type: 'object',
            description: 'Borders. Each edge accepts a style name or { style, color }. Use "all" to set every edge at once. Styles: thin, medium, thick, dashed, dotted, double, hair, mediumDashed, dashDot, mediumDashDot, dashDotDot, mediumDashDotDot, slantDashDot.',
            properties: {
              all: { description: 'Apply one style to all four edges.' },
              top: { description: 'Top edge.' },
              bottom: { description: 'Bottom edge.' },
              left: { description: 'Left edge.' },
              right: { description: 'Right edge.' }
            }
          },
          alignment: {
            type: 'object',
            description: 'Text alignment.',
            properties: {
              horizontal: { type: 'string', enum: ['left', 'center', 'right', 'fill', 'justify'], description: 'Horizontal alignment.' },
              vertical: { type: 'string', enum: ['top', 'center', 'bottom'], description: 'Vertical alignment.' },
              wrapText: { type: 'boolean', description: 'Wrap text inside the cell.' },
              textRotation: { type: 'integer', description: 'Rotation in degrees (0..180).' }
            }
          },
          numberFormat: {
            type: 'object',
            description: 'Excel number format. Values are stored as numbers, never as preformatted text.',
            properties: {
              type: { type: 'string', enum: ['general', 'integer', 'decimal', 'currency', 'percent', 'date', 'datetime', 'time', 'text', 'custom'] },
              decimals: { type: 'integer', description: 'Decimal places for decimal/currency/percent.' },
              symbol: { type: 'string', description: 'Currency symbol (default ₽).' },
              thousands: { type: 'boolean', description: 'Use a thousands separator.' },
              code: { type: 'string', description: 'Explicit format code when type is "custom", e.g. "0.000".' }
            }
          },
          merge: { type: 'boolean', description: 'Merge the range (or unmerge when set with unmerge: true).' },
          unmerge: { type: 'boolean', description: 'Remove the merge covering this range.' },
          columnWidth: {
            type: 'object',
            description: 'Set the width of the columns spanned by the range.',
            properties: {
              width: { type: 'number', description: 'Width in characters; omit to auto-fit.' },
              auto: { type: 'boolean', description: 'Estimate a width that fits the content.' }
            }
          },
          rowHeight: { type: 'number', description: 'Height in points for every row in the range.' },
          outputPath: { type: 'string', description: 'Optional target path.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        return await xlsxEngine.format(args.filePath, args)
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
          numberFormat: { type: 'object', description: 'Optional number format for the cell, e.g. { type: "currency" }.' },
          outputPath: { type: 'string', description: 'Optional target path.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        return await xlsxEngine.write(args.filePath, {
          outputPath: args.outputPath,
          sheetIndex: args.sheetIndex,
          sheetName: args.sheetName,
          cells: [{
            ref: args.cell,
            formula: args.formula,
            numberFormat: args.numberFormat
          }]
        })
      }
    }
  ]
}
