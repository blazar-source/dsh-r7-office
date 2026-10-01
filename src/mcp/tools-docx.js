import { defaultOutput } from './tool-output.js'

/**
 * DOCX tool definitions beyond the format-agnostic inspect/read/create.
 *
 * The engine surface is deliberately finer-grained than these tools; what is
 * exposed here is the set of operations an agent is expected to reach for,
 * grouped so that no two tools are near-synonyms of each other:
 *
 *  - `r7_docx_formatting` — read the document's normalized formatting before
 *    changing anything (styles, fonts, alignment, indents, spacing, lists,
 *    table cell styles, sections, header/footer parts).
 *  - `r7_docx_sections` — read or change page setup, page breaks and section
 *    breaks, preserving every property the caller does not name.
 *  - `r7_docx_header_footer` — read, create and re-text headers and footers
 *    without destroying page-number fields.
 *  - `r7_docx_image` — insert a PNG/JPEG with an explicit display size, or
 *    list the images already in the package.
 *  - `r7_docx_hyperlink` — read, insert, re-text and remove hyperlinks with
 *    their relationships intact.
 *
 * @param {import('../r7/docx.js').DocxEngine} docxEngine
 * @returns {Array<object>}
 */
export function buildDocxTools(docxEngine) {
  /** Paragraph formatting, shared by the formatting reader and the writers. */
  const paragraphFormattingProperties = {
    style: { type: 'string', description: 'Paragraph style: Normal, Heading1…Heading6, Title, ListParagraph, or any style name/id the document defines.' },
    alignment: { type: 'string', enum: ['left', 'center', 'right', 'both'], description: 'Horizontal alignment ("both" is justified).' },
    indents: {
      type: 'object',
      description: 'Indents in points: { left, right, firstLine, hanging }.',
      properties: {
        left: { type: 'number' },
        right: { type: 'number' },
        firstLine: { type: 'number' },
        hanging: { type: 'number' }
      }
    },
    spacing: {
      type: 'object',
      description: 'Spacing: { before, after } in points and { line, lineRule } for line spacing, where lineRule "auto" makes line a multiple (1.5) and "exact"/"atLeast" make it points.',
      properties: {
        before: { type: 'number' },
        after: { type: 'number' },
        line: { type: 'number' },
        lineRule: { type: 'string', enum: ['auto', 'exact', 'atLeast'] }
      }
    },
    list: { type: 'string', enum: ['bullet', 'number'], description: 'Make the paragraph a bulleted or numbered list item. The numbering part and its relationship are created when the document has none.' },
    listLevel: { type: 'integer', description: 'List level 0..8 (default 0).' },
    pageBreakBefore: { type: 'boolean', description: 'Start the paragraph on a new page.' },
    keepNext: { type: 'boolean', description: 'Keep the paragraph with the next one.' },
    shading: { type: 'string', description: 'Paragraph background, e.g. "#FFF2CC".' },
    bold: { type: 'boolean' },
    italic: { type: 'boolean' },
    underline: { type: 'boolean', description: 'true for a single underline, or a style name such as double.' },
    strike: { type: 'boolean' },
    family: { type: 'string', description: 'Font name, e.g. "Georgia".' },
    size: { type: 'number', description: 'Font size in points.' },
    color: { type: 'string', description: 'Font colour, e.g. "#C00000".' },
    highlight: { type: 'string', description: 'Text highlight colour, e.g. "yellow".' },
    runs: {
      type: 'array',
      description: 'Mixed formatting inside one paragraph: [{ text, bold, italic, size, color, family, underline }].',
      items: { type: 'object' }
    }
  }

  const cellProperties = {
    value: { type: 'string', description: 'Cell text.' },
    text: { type: 'string', description: 'Cell text (alias of value).' },
    shading: { type: 'string', description: 'Cell background, e.g. "#DEEAF6". Use "none" to clear it.' },
    verticalAlign: { type: 'string', enum: ['top', 'center', 'bottom'], description: 'Vertical alignment inside the cell.' },
    alignment: { type: 'string', enum: ['left', 'center', 'right', 'both'], description: 'Horizontal alignment of the cell paragraphs.' },
    widthCm: { type: 'number', description: 'Cell width in centimetres.' },
    widthTwips: { type: 'number', description: 'Cell width in twips (1/1440 inch).' },
    clearBorders: { type: 'boolean', description: 'Remove the cell borders.' },
    borders: {
      type: 'object',
      description: 'Per-edge borders: { top, left, bottom, right }, each { style, sizePoints, color } or a style name such as "single".',
      properties: {
        top: { type: 'object' },
        left: { type: 'object' },
        bottom: { type: 'object' },
        right: { type: 'object' }
      }
    }
  }

  return [
    {
      name: 'r7_edit',
      description: 'Edit, update text, style, or remove a specific paragraph in a DOCX document. The paragraph keeps its own properties (section break, list numbering, spacing) and the new text inherits the previous first run\'s formatting unless you ask for different formatting.',
      parameters: {
        type: 'object',
        required: ['filePath', 'paragraphIndex'],
        properties: {
          filePath: { type: 'string', description: 'Path to the DOCX document.' },
          paragraphIndex: { type: 'integer', description: '0-based index of the target paragraph (from r7_inspect/r7_read).' },
          newText: { type: 'string', description: 'New text content for the paragraph.' },
          style: { type: 'string', description: 'Paragraph style name (e.g. Normal, Heading1, Heading2).' },
          preserveFormatting: { type: 'boolean', description: 'Keep the paragraph\'s existing character formatting on the new text (default true).' },
          formatting: { type: 'object', description: 'Character formatting for the new text: { bold, italic, underline, strike, family, size, color, highlight }.' },
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
      description: 'Insert a paragraph, heading, list item, or page break at the start, the end, or relative to another paragraph. The inserted paragraph can carry its own alignment, indents, spacing and font.',
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
          pageBreak: { type: 'boolean', description: 'Insert a page break (alone, or before the text).' },
          list: { type: 'string', enum: ['bullet', 'number'], description: 'Insert the paragraph as a list item.' },
          listLevel: { type: 'integer', description: 'List level 0..8 (default 0).' },
          ...paragraphFormattingProperties,
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
      description: 'Create, inspect, or modify tables in DOCX documents: add and remove rows and columns, merge and unmerge cells, set column widths, and format individual cells (shading, borders, alignment, width). Cell formatting never rewrites the text or the properties of neighbouring cells.',
      parameters: {
        type: 'object',
        required: ['filePath', 'action'],
        properties: {
          filePath: { type: 'string', description: 'Path to the DOCX file.' },
          action: {
            type: 'string',
            enum: ['create', 'inspect', 'addRow', 'removeRow', 'addColumn', 'removeColumn', 'setCell', 'formatCell', 'merge', 'unmerge', 'setColumnWidths'],
            description: 'Table action. "merge" and "unmerge" take merge { row, col, rows, cols }.'
          },
          tableIndex: { type: 'integer', description: '0-based table index for existing tables.' },
          rows: { type: 'array', description: '2D array of rows and cell values for create/addRow. A cell may be a string or an object with the cell properties.' },
          values: { type: 'array', description: 'Row (addRow) or column (addColumn) cell values.' },
          cell: {
            type: 'object',
            properties: {
              row: { type: 'integer' },
              col: { type: 'integer' },
              ...cellProperties
            },
            description: 'Cell coordinate and properties for setCell/formatCell.'
          },
          merge: {
            type: 'object',
            properties: {
              row: { type: 'integer', description: 'Top-left cell row (0-based).' },
              col: { type: 'integer', description: 'Top-left cell index within its row.' },
              rows: { type: 'integer', description: 'How many rows the merge covers (default 1).' },
              cols: { type: 'integer', description: 'How many cells it covers per row (default 1).' },
              axis: { type: 'string', enum: ['all', 'horizontal', 'vertical'], description: 'For unmerge: which merge to release.' }
            },
            description: 'Merge block for the merge/unmerge actions.'
          },
          rowIndex: { type: 'integer', description: 'Row to remove.' },
          columnIndex: { type: 'integer', description: 'Column to remove or where to insert one.' },
          index: { type: 'integer', description: 'Insertion index for addRow/addColumn.' },
          widthsTwips: { type: 'array', items: { type: 'number' }, description: 'Column widths in twips for create/setColumnWidths (and per-cell widths for addRow).' },
          widths: { type: 'array', items: { type: 'number' }, description: 'Alias of widthsTwips.' },
          shading: { type: 'string', description: 'Shading applied to every new cell of addRow/addColumn.' },
          verticalAlign: { type: 'string', enum: ['top', 'center', 'bottom'], description: 'Vertical alignment for new cells.' },
          alignment: { type: 'string', enum: ['left', 'center', 'right', 'both'], description: 'Alignment for new cells, or for the whole new table.' },
          borders: {
            description: 'Table borders for create: false for none, or { top, left, bottom, right, insideH, insideV } with { style, sizePoints, color }.',
            oneOf: [{ type: 'boolean' }, { type: 'object' }]
          },
          fixedLayout: { type: 'boolean', description: 'For setColumnWidths: pin the layout so the widths are honoured (default true).' },
          includeStyles: { type: 'boolean', description: 'For inspect: also return normalized cell formatting (widths, merges, shading, borders, alignment).' },
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
      name: 'r7_docx_formatting',
      description: 'Read the normalized formatting of a DOCX document before changing it: per-paragraph style, font family/size/bold/italic/underline/strike/colour, alignment, indents, spacing, list numbering, plus table cell styles (widths, merges, shading, borders, alignment), page/section setup and header/footer parts. Values are normalized (points, #RRGGBB), never raw OOXML.',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to the DOCX file.' },
          fromParagraph: { type: 'integer', description: '0-based paragraph to start from (default 0).' },
          count: { type: 'integer', description: 'How many paragraphs to return (default 200).' },
          includeRuns: { type: 'boolean', description: 'Include the per-run detail inside each paragraph (default true).' },
          includeTables: { type: 'boolean', description: 'Include the table cell styles (default true).' },
          includeSections: { type: 'boolean', description: 'Include page setup and section breaks (default true).' },
          includeHeadersFooters: { type: 'boolean', description: 'Include header/footer parts and their fields (default true).' },
          includeStylesList: { type: 'boolean', description: 'Also list every style the document defines.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        return await docxEngine.formatting(args.filePath, args)
      }
    },

    {
      name: 'r7_docx_sections',
      description: 'Read or change DOCX page setup and sections: page size, portrait/landscape, margins, columns, page-number restart, page breaks and section breaks. Changing one property preserves everything else in the section (headers, footers, columns, document grid, title-page setting).',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to the DOCX file.' },
          action: {
            type: 'string',
            enum: ['read', 'set', 'insertBreak', 'remove'],
            description: 'read (default) reports every section and page break; set changes one section; insertBreak starts a new section.'
          },
          sectionIndex: { type: 'integer', description: 'Which section to change (0-based; -1 is the final section, which is the default).' },
          orientation: { type: 'string', enum: ['portrait', 'landscape'], description: 'Page orientation.' },
          widthCm: { type: 'number', description: 'Explicit page width in centimetres.' },
          heightCm: { type: 'number', description: 'Explicit page height in centimetres.' },
          margins: {
            type: 'object',
            description: 'Margins in centimetres: { top, right, bottom, left, header, footer, gutter }.',
            properties: {
              top: { type: 'number' },
              right: { type: 'number' },
              bottom: { type: 'number' },
              left: { type: 'number' },
              header: { type: 'number' },
              footer: { type: 'number' },
              gutter: { type: 'number' }
            }
          },
          columns: { type: 'integer', description: 'Number of text columns.' },
          columnSpaceCm: { type: 'number', description: 'Space between columns, in centimetres.' },
          separator: { type: 'boolean', description: 'Draw a vertical line between columns.' },
          titlePg: { type: 'boolean', description: 'Different first-page header/footer.' },
          pageNumberStart: { type: 'integer', description: 'Restart page numbering at this value.' },
          pageNumberFormat: { type: 'string', enum: ['decimal', 'upperRoman', 'lowerRoman', 'upperLetter', 'lowerLetter', 'decimalZero'], description: 'Page number format.' },
          type: { type: 'string', enum: ['nextPage', 'continuous', 'evenPage', 'oddPage', 'nextColumn'], description: 'Section break type.' },
          afterParagraphIndex: { type: 'integer', description: 'For insertBreak: the last paragraph of the section being closed.' },
          page: { type: 'object', description: 'For insertBreak: page setup applied to the NEW (following) section.' },
          outputPath: { type: 'string', description: 'Write to a copy instead of editing in place.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        switch (args.action || 'read') {
          case 'read':
            return await docxEngine.sections(args.filePath)
          case 'set':
            return await docxEngine.setSection(args.filePath, args)
          case 'insertBreak': {
            if (args.afterParagraphIndex === undefined) {
              throw new Error('action "insertBreak" requires afterParagraphIndex')
            }
            return await docxEngine.insertSectionBreak(args.filePath, args)
          }
          case 'remove':
            throw new Error('Removing a section is not supported: a document must keep at least its final section.')
          default:
            throw new Error(`Unknown action "${args.action}". Use read, set or insertBreak.`)
        }
      }
    },

    {
      name: 'r7_docx_header_footer',
      description: 'Read, create or re-text DOCX headers and footers. A text change rewrites only the non-field runs of the paragraph and keeps every field (PAGE, NUMPAGES, DATE…) exactly where it was, so a page number is never destroyed. Page-number fields are added with the fldSimple form R7 and Word both understand.',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to the DOCX file.' },
          action: {
            type: 'string',
            enum: ['list', 'read', 'setText', 'create', 'addPageNumber', 'remove'],
            description: 'list (default) reports every header/footer part; read returns one part; setText changes its text; create makes it when missing; addPageNumber adds a PAGE field; remove deletes the part.'
          },
          kind: { type: 'string', enum: ['header', 'footer'], description: 'Which part (default header).' },
          type: { type: 'string', enum: ['default', 'first', 'even'], description: 'Which of the three header/footer kinds (default default).' },
          sectionIndex: { type: 'integer', description: 'Section that owns the part (0-based, default 0).' },
          text: { type: 'string', description: 'New text for the part.' },
          paragraphIndex: { type: 'integer', description: 'Which paragraph of the part to re-text (default 0).' },
          pageNumber: {
            description: 'Add a page-number field: true for PAGE, or a field instruction such as "NUMPAGES" or "PAGE of NUMPAGES".',
            oneOf: [{ type: 'boolean' }, { type: 'string' }]
          },
          position: { type: 'string', enum: ['start', 'end'], description: 'Where to put the field inside the paragraph (default end).' },
          separator: { type: 'string', description: 'Literal text placed before the field, e.g. "Page ".' },
          alignment: { type: 'string', enum: ['left', 'center', 'right', 'both'], description: 'Alignment for a newly created part.' },
          outputPath: { type: 'string', description: 'Write to a copy instead of editing in place.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        return await docxEngine.headerFooter(args.filePath, args)
      }
    },

    {
      name: 'r7_docx_image',
      description: 'Insert a PNG/JPEG/GIF into a DOCX document, or list the images already in it. Inserting creates the media part, its relationship and its content type, and keeps the picture\'s own aspect ratio when only one dimension is given. Existing images and relationships are never touched.',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to the DOCX file.' },
          action: { type: 'string', enum: ['insert', 'list'], description: 'insert (default) or list.' },
          imagePath: { type: 'string', description: 'Path to the image file (PNG, JPEG or GIF).' },
          data: { type: 'string', description: 'Base64 image data (a data: URL also works) instead of imagePath.' },
          paragraphIndex: {
            description: 'Where the picture goes: "end" (default), "start", "new" for its own paragraph, or a 0-based paragraph index.',
            oneOf: [{ type: 'integer' }, { type: 'string' }]
          },
          widthCm: { type: 'number', description: 'Display width in centimetres. The height is derived from the image\'s aspect ratio.' },
          heightCm: { type: 'number', description: 'Display height in centimetres.' },
          widthPx: { type: 'number', description: 'Display width in pixels at 96 dpi.' },
          heightPx: { type: 'number', description: 'Display height in pixels at 96 dpi.' },
          alt: { type: 'string', description: 'Alternative text stored with the picture.' },
          name: { type: 'string', description: 'Picture name shown in the document\'s selection pane.' },
          outputPath: { type: 'string', description: 'Write to a copy instead of editing in place.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        if ((args.action || 'insert') === 'list') {
          return { filePath: args.filePath, images: await docxEngine.images(args.filePath) }
        }
        if (!args.imagePath && !args.data) throw new Error('action "insert" requires imagePath or data')
        return await docxEngine.insertImage(args.filePath, args)
      }
    },

    {
      name: 'r7_docx_hyperlink',
      description: 'Read, insert, re-text or remove DOCX hyperlinks. Reading returns each link\'s text and target URL; inserting writes the relationship (external, deduplicated by URL) and a Word-compatible Hyperlink-styled run; re-texting keeps the relationship and the run formatting.',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to the DOCX file.' },
          action: { type: 'string', enum: ['list', 'insert', 'setText', 'remove'], description: 'list (default), insert, setText or remove.' },
          url: { type: 'string', description: 'Target URL for insert.' },
          anchor: { type: 'string', description: 'Internal bookmark/anchor name instead of a URL.' },
          text: { type: 'string', description: 'Link text (insert) or its replacement (setText).' },
          tooltip: { type: 'string', description: 'Tooltip shown on hover.' },
          index: { type: 'integer', description: 'Hyperlink index from action "list", for setText/remove.' },
          paragraphIndex: { type: 'integer', description: 'Paragraph the link joins (insert) or whose first link to edit/remove.' },
          newParagraph: { type: 'boolean', description: 'For insert: put the link in its own paragraph.' },
          alignment: { type: 'string', enum: ['left', 'center', 'right', 'both'], description: 'Alignment of the new paragraph.' },
          bold: { type: 'boolean' },
          italic: { type: 'boolean' },
          size: { type: 'number', description: 'Font size in points for the link text.' },
          color: { type: 'string', description: 'Font colour for the link text, e.g. "#0563C1".' },
          style: { type: 'string', description: 'Character style id for the link (default Hyperlink).' },
          outputPath: { type: 'string', description: 'Write to a copy instead of editing in place.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        return await docxEngine.hyperlink(args.filePath, args)
      }
    }
  ]
}
