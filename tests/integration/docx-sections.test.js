import { test, describe, before, after } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import { DocxEngine } from '../../src/r7/docx.js'
import { R7Adapter } from '../../src/r7/adapter.js'
import { archiveOf, cleanup, diffMembers, tempDir } from './docx-fixtures.test.js'

const r7Available = (await new R7Adapter().detect()).installed

/**
 * Scenario 2: page size, orientation, margins, page breaks and section breaks —
 * read and written, with everything the caller did not name preserved.
 */
describe('DOCX page setup and sections', () => {
  const dir = tempDir('docx_sections')
  const file = path.join(dir, 'sections.docx')
  const engine = new DocxEngine()

  before(async () => {
    await engine.create(file, {
      title: 'Разделы и страницы',
      paragraphs: [
        'Первый абзац первого раздела.',
        'Второй абзац первого раздела.',
        'Абзац после разрыва.',
        'Последний абзац документа.'
      ],
      header: { text: 'Колонтитул' },
      footer: { text: 'Стр.', pageNumber: true }
    })
  })

  after(() => cleanup(dir))

  test('the document reports one final A4 portrait section with its margins', async () => {
    const result = await engine.sections(file)
    assert.equal(result.sectionCount, 1)
    const section = result.sections[0]
    assert.equal(section.kind, 'final')
    assert.equal(section.paragraphIndex, null)
    assert.equal(section.pageSize.orientation, 'portrait')
    assert.equal(section.pageSize.widthCm, 21)
    assert.equal(section.pageSize.heightCm, 29.7)
    assert.equal(section.pageSize.assumed, false)
    // R7's own template margins: 2 cm top, 1.5 cm right, 2 cm bottom, 3 cm left.
    assert.equal(section.margins.top.cm, 2)
    assert.equal(section.margins.right.cm, 1.5)
    assert.equal(section.margins.bottom.cm, 2)
    assert.equal(section.margins.left.cm, 3)
    assert.equal(section.headers.default.partName, 'word/header1.xml')
    assert.equal(section.footers.default.partName, 'word/footer1.xml')
  })

  test('landscape swaps the page dimensions and keeps the margins', async () => {
    const set = await engine.setSection(file, { orientation: 'landscape' })
    assert.equal(set.section.pageSize.orientation, 'landscape')
    assert.equal(set.section.pageSize.widthCm, 29.7)
    assert.equal(set.section.pageSize.heightCm, 21)
    assert.equal(set.section.margins.left.cm, 3, 'an unnamed margin must survive')
    assert.equal(set.section.headers.default.relId !== null, true, 'header references must survive')
    assert.equal(set.section.footers.default.relId !== null, true, 'footer references must survive')

    const back = await engine.setSection(file, { orientation: 'portrait' })
    assert.equal(back.section.pageSize.orientation, 'portrait')
    assert.equal(back.section.pageSize.widthCm, 21)
  })

  test('margins are set in centimetres', async () => {
    const set = await engine.setSection(file, {
      margins: { top: 2.5, right: 1.5, bottom: 2.5, left: 1.5, header: 1.25, footer: 1.25 }
    })
    assert.equal(set.section.margins.top.cm, 2.5)
    assert.equal(set.section.margins.left.cm, 1.5)
    assert.equal(set.section.margins.header.cm, 1.25)
    assert.equal(set.section.margins.gutter.cm, 0, 'an absent gutter reads back as 0, not as lost')
  })

  test('a margin given as twips keeps its exact OOXML value', async () => {
    const set = await engine.setSection(file, { margins: { left: { twips: 1234 } } })
    assert.equal(set.section.margins.left.twips, 1234)
  })

  test('columns, title page and page numbering are recorded', async () => {
    const set = await engine.setSection(file, {
      columns: 2,
      columnSpaceCm: 1,
      separator: true,
      titlePg: true,
      pageNumberStart: 5,
      pageNumberFormat: 'lowerRoman'
    })
    assert.equal(set.section.columns.count, 2)
    assert.equal(set.section.columns.separator, true)
    assert.equal(set.section.columns.spaceCm, 1)
    assert.equal(set.section.titlePg, true)
    assert.equal(set.section.pageNumbering.start, 5)
    assert.equal(set.section.pageNumbering.format, 'lowerRoman')

    // The columns setting is removed again but the page numbering survives.
    const cleared = await engine.setSection(file, { columns: 1, separator: false, titlePg: false })
    assert.equal(cleared.section.columns.count, 1)
    assert.equal(cleared.section.titlePg, false)
    assert.equal(cleared.section.pageNumbering.start, 5)
  })

  test('an invalid orientation or margin name is refused', async () => {
    await assert.rejects(() => engine.setSection(file, { orientation: 'sideways' }), /Invalid orientation/)
    await assert.rejects(() => engine.setSection(file, { margins: { middle: 2 } }), /Unknown margin/)
    await assert.rejects(() => engine.setSection(file, { sectionIndex: 7 }), /Section index out of range/)
  })

  test('editing page setup leaves every other package part byte-identical', async () => {
    const before = await archiveOf(file)
    await engine.setSection(file, { orientation: 'landscape' })
    const after = await archiveOf(file)
    const diff = diffMembers(before, after, ['word/document.xml'])
    assert.deepEqual(diff.changed, [], 'only document.xml may change')
    assert.deepEqual(diff.added, [])
    assert.deepEqual(diff.removed, [])
  })

  test('page breaks are inserted and reported', async () => {
    await engine.insert(file, { position: 'end', pageBreak: true })
    await engine.insert(file, { position: 'end', text: 'С новой страницы', pageBreakBefore: true })

    const breaks = await engine.pageBreaks(file)
    assert.ok(breaks.length >= 2, `expected at least two page breaks, got ${breaks.length}`)
    assert.ok(breaks.some(entry => entry.kind === 'break'), 'an explicit w:br must be reported')
    assert.ok(breaks.some(entry => entry.kind === 'pageBreakBefore'), 'pageBreakBefore must be reported')
  })

  test('a section break splits the document and can flip the following section', async () => {
    const source = path.join(dir, 'split.docx')
    await engine.create(source, {
      paragraphs: ['Абзац 1', 'Абзац 2', 'Абзац 3'],
      header: { text: 'Колонтитул' }
    })

    const result = await engine.insertSectionBreak(source, {
      afterParagraphIndex: 1,
      type: 'nextPage',
      page: { orientation: 'landscape', margins: { left: 1 } }
    })
    assert.equal(result.sectionCount, 2)

    const sections = (await engine.sections(source)).sections
    assert.equal(sections.length, 2)

    // The first section ends at the inserted break paragraph — which is a real
    // paragraph in the body and shifts the paragraph indices after it.
    assert.equal(sections[0].kind, 'paragraph')
    assert.equal(sections[0].paragraphIndex, 2)
    assert.equal(sections[0].pageSize.orientation, 'portrait')
    assert.equal(sections[0].type, 'nextPage')
    assert.equal(sections[0].headers.default.partName, 'word/header1.xml')

    assert.equal(sections[1].kind, 'final')
    assert.equal(sections[1].paragraphIndex, null)
    assert.equal(sections[1].pageSize.orientation, 'landscape')
    assert.equal(sections[1].margins.left.cm, 1)

    // The body-level section must still be the last child of the body.
    const zip = await archiveOf(source)
    const docXml = zip.getText('word/document.xml')
    assert.match(docXml, /<\/w:p>\s*<w:sectPr[^>]*>[\s\S]*<\/w:sectPr><\/w:body>/)
    assert.equal((await engine.validate(source)).valid, true)
  })

  test('each section is addressable on its own', async () => {
    const source = path.join(dir, 'two-sections.docx')
    await engine.create(source, { paragraphs: ['A', 'B', 'C'] })
    await engine.insertSectionBreak(source, { afterParagraphIndex: 0, page: { orientation: 'landscape' } })

    const before = (await engine.sections(source)).sections
    await engine.setSection(source, { sectionIndex: 0, margins: { left: 4 } })
    const after = (await engine.sections(source)).sections

    assert.equal(after[0].margins.left.cm, 4)
    assert.equal(after[1].margins.left.cm, before[1].margins.left.cm, 'the other section must not move')
    assert.equal(after[1].pageSize.orientation, 'landscape')
    assert.equal(after[0].pageSize.orientation, 'portrait')

    const negative = await engine.setSection(source, { sectionIndex: -1, margins: { left: 2 } })
    assert.equal(negative.sectionIndex, 1, '-1 addresses the final section')
    assert.equal(negative.section.margins.left.cm, 2)
  })

  test('an out-of-range paragraph is refused for a section break', async () => {
    await assert.rejects(
      () => engine.insertSectionBreak(file, { afterParagraphIndex: 999 }),
      /Paragraph index out of range/
    )
  })

  test('sections survive an ordinary text edit', async () => {
    const source = path.join(dir, 'sections-edit.docx')
    await engine.create(source, {
      paragraphs: ['Первый абзац для замены.', 'Второй абзац.'],
      page: { orientation: 'landscape' }
    })
    await engine.insertSectionBreak(source, { afterParagraphIndex: 0, page: { orientation: 'portrait' } })

    const before = (await engine.sections(source)).sections
    const beforeDoc = (await archiveOf(source)).getText('word/document.xml')

    const result = await engine.replaceText(source, 'Первый абзац для замены.', 'Заменённый абзац.')
    assert.ok(result.matchesCount > 0)

    const after = (await engine.sections(source)).sections
    assert.deepEqual(after, before, 'page and section settings must be identical after an edit')

    const afterDoc = (await archiveOf(source)).getText('word/document.xml')
    assert.match(afterDoc, /<w:pgSz[^>]*w:w="16838"/, 'the landscape section keeps its page size')
    assert.ok(afterDoc.length > 0 && beforeDoc.length > 0)
  })

  test('R7 reopens the multi-section document and renders it', async (t) => {
    if (!r7Available) {
      t.skip('R7 not installed')
      return
    }
    const source = path.join(dir, 'render-sections.docx')
    await engine.create(source, {
      paragraphs: ['Первый раздел.', 'Второй раздел после разрыва.'],
      header: { text: 'Колонтитул' },
      footer: { text: 'Стр.', pageNumber: true }
    })
    await engine.insertSectionBreak(source, { afterParagraphIndex: 0, page: { orientation: 'landscape' } })

    const pdf = path.join(dir, 'render-sections.pdf')
    await new R7Adapter().convert(source, pdf)
    const head = fs.readFileSync(pdf).subarray(0, 5).toString()
    assert.equal(head, '%PDF-')
    assert.equal((await engine.validate(source)).valid, true)
  })
})
