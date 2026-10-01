import { test, describe, before, after } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import { DocxEngine } from '../../src/r7/docx.js'
import { R7Adapter } from '../../src/r7/adapter.js'
import { cleanup, tempDir } from './docx-fixtures.test.js'

const r7Available = (await new R7Adapter().detect()).installed

/**
 * Scenario 1: an agent must be able to see the document's formatting before
 * changing it, as normalized objects rather than raw OOXML, and the default
 * read/inspect output must stay exactly as it was.
 */
describe('DOCX formatting read-back', () => {
  const dir = tempDir('docx_format')
  const file = path.join(dir, 'форматирование.docx')
  const engine = new DocxEngine()
  let report

  before(async () => {
    await engine.create(file, {
      title: 'Документ с оформлением',
      paragraphs: [
        { text: 'Заголовок раздела', style: 'Heading2' },
        {
          text: 'Обычный абзац с отступами.',
          family: 'Georgia',
          size: 11,
          color: '#123456',
          alignment: 'both',
          indents: { left: 24, firstLine: 12 },
          spacing: { before: 6, after: 10, line: 1.5, lineRule: 'auto' }
        },
        { text: 'Акцентная строка', bold: true, italic: true, underline: true, size: 14, highlight: 'yellow' },
        { text: 'Пункт списка', list: 'bullet' },
        { text: 'Нумерованный пункт', list: 'number' }
      ],
      tables: [{
        widthsTwips: [3000, 2000],
        rows: [[
          { value: 'Ячейка', shading: '#DEEAF6', verticalAlign: 'center', alignment: 'center' },
          'B'
        ], ['C', 'D']]
      }],
      header: { text: 'Колонтитул' },
      footer: { text: 'Стр.', pageNumber: true }
    })
    report = await engine.formatting(file)
  })

  after(() => cleanup(dir))

  test('a heading is reported with its level and a resolved style name', () => {
    const heading = report.paragraphs.find(paragraph => paragraph.text === 'Заголовок раздела')
    assert.ok(heading, 'the heading paragraph must be present')
    assert.equal(heading.isHeading, true)
    assert.equal(heading.headingLevel, 2)
    assert.match(heading.style, /heading/i)
    assert.match(heading.styleName, /heading\s*2/i)
  })

  test('font family, size and colour are normalized to points and #RRGGBB', () => {
    const paragraph = report.paragraphs.find(entry => entry.text === 'Обычный абзац с отступами.')
    assert.equal(paragraph.runs.length, 1)
    const run = paragraph.runs[0]
    assert.equal(run.family, 'Georgia')
    assert.equal(run.size, 11)
    assert.equal(run.color, '#123456')
  })

  test('alignment, indents and spacing are normalized', () => {
    const paragraph = report.paragraphs.find(entry => entry.text === 'Обычный абзац с отступами.')
    assert.equal(paragraph.alignment, 'both')
    assert.equal(paragraph.indents.left, 24)
    assert.equal(paragraph.indents.firstLine, 12)
    assert.equal(paragraph.spacing.before, 6)
    assert.equal(paragraph.spacing.after, 10)
    assert.equal(paragraph.spacing.line, 1.5)
    assert.equal(paragraph.spacing.lineRule, 'auto')
  })

  test('bold, italic, underline, size and highlight are read per run', () => {
    const paragraph = report.paragraphs.find(entry => entry.text === 'Акцентная строка')
    const run = paragraph.runs[0]
    assert.equal(run.bold, true)
    assert.equal(run.italic, true)
    assert.equal(run.underline, true)
    assert.equal(run.underlineStyle, 'single')
    assert.equal(run.size, 14)
    assert.equal(run.highlight, 'yellow')
  })

  test('list paragraphs report their numbering format and marker', () => {
    const bullet = report.paragraphs.find(entry => entry.text === 'Пункт списка')
    const numbered = report.paragraphs.find(entry => entry.text === 'Нумерованный пункт')

    assert.ok(bullet.list, 'the bullet paragraph must be a list item')
    assert.equal(bullet.list.format, 'bullet')
    assert.equal(bullet.list.level, 0)
    assert.ok(bullet.list.marker, 'a bullet must carry a marker glyph')

    assert.ok(numbered.list)
    assert.equal(numbered.list.format, 'decimal')
  })

  test('table cell styles are reported per cell', () => {
    assert.equal(report.tables.length, 1)
    const table = report.tables[0]
    assert.deepEqual(table.columns, [3000, 2000])
    assert.equal(table.rows.length, 2)

    const first = table.rows[0].cells[0]
    assert.equal(first.text, 'Ячейка')
    assert.equal(first.shading, '#DEEAF6')
    assert.equal(first.verticalAlign, 'center')
    assert.equal(first.alignment, 'center')
    assert.equal(first.width.twips, 3000)

    // A cell the caller did not format must not invent formatting.
    const second = table.rows[0].cells[1]
    assert.equal(second.text, 'B')
    assert.equal(second.shading, null)
    assert.equal(second.verticalAlign, null)

    assert.ok(table.borders?.top?.style, 'the created table carries borders')
  })

  test('sections and header/footer parts are part of the formatting report', () => {
    assert.equal(report.sections.length, 1)
    assert.equal(report.sections[0].pageSize.orientation, 'portrait')

    const kinds = report.headersFooters.map(part => `${part.kind}:${part.type}`)
    assert.ok(kinds.includes('header:default'))
    assert.ok(kinds.includes('footer:default'))

    const footer = report.headersFooters.find(part => part.kind === 'footer')
    assert.equal(footer.hasPageNumberField, true)
    assert.match(footer.text, /\{PAGE\}/)
  })

  test('the report documents the units of every normalized number', () => {
    assert.ok(report.styleVocabulary.units.indents)
    assert.ok(report.styleVocabulary.units.colour)
    assert.ok(report.styleVocabulary.paragraphKeys.includes('list'))
  })

  test('r7_read returns the same formatting when includeStyles is requested', async () => {
    const read = await engine.read(file, { includeStyles: true })
    const paragraph = read.paragraphs.find(entry => entry.text === 'Обычный абзац с отступами.')
    assert.equal(paragraph.formatting.runs[0].family, 'Georgia')
    assert.equal(paragraph.formatting.alignment, 'both')
  })

  test('inspect reports formatting only when asked', async () => {
    const plain = await engine.inspect(file)
    for (const entry of plain.outline) {
      assert.equal('formatting' in entry, false, 'the default outline must not grow new keys')
    }
    assert.equal('sections' in plain, false)

    const rich = await engine.inspect(file, { includeStyles: true, includeSections: true })
    const heading = rich.outline.find(entry => entry.style === 'Heading1')
    assert.ok(heading.formatting, 'includeStyles must attach normalized formatting')
    assert.equal(rich.sections.length, 1)
  })

  test('the default read output is unchanged', async () => {
    const structured = await engine.read(file, { format: 'structured' })
    assert.deepEqual(Object.keys(structured).sort(), ['format', 'paragraphs', 'totalParagraphs'])
    assert.deepEqual(Object.keys(structured.paragraphs[0]).sort(), ['index', 'style', 'text'])

    const markdown = await engine.read(file, { format: 'markdown' })
    assert.deepEqual(Object.keys(markdown).sort(), ['content', 'format', 'returned', 'totalParagraphs'])
    assert.ok(markdown.content.includes('Заголовок раздела'))
  })

  test('the document still validates and R7 renders it', async (t) => {
    const validation = await engine.validate(file)
    assert.equal(validation.valid, true, JSON.stringify(validation.errors))
    assert.equal(validation.details.parts.headersFooters.length, 2)

    if (!r7Available) {
      t.skip('R7 not installed')
      return
    }
    const adapter = new R7Adapter()
    const pdf = path.join(dir, 'форматирование.pdf')
    await adapter.convert(file, pdf)
    assert.equal(fs.readFileSync(pdf).subarray(0, 5).toString(), '%PDF-')
    assert.ok(fs.statSync(pdf).size > 1000)
  })
})
