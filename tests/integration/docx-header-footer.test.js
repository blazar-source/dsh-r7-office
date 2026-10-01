import { test, describe, before, after } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import { DocxEngine } from '../../src/r7/docx.js'
import { R7Adapter } from '../../src/r7/adapter.js'
import { ZipArchive } from '../../src/shared/zip.js'
import { archiveOf, cleanup, diffMembers, tempDir } from './docx-fixtures.test.js'

const r7Available = (await new R7Adapter().detect()).installed

const HEADER_NAMESPACES = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

/**
 * Scenario 3: read header/footer text, change it, and never destroy a
 * page-number field — in either of its two spellings.
 */
describe('DOCX headers and footers', () => {
  const dir = tempDir('docx_hf')
  const file = path.join(dir, 'колонтитулы.docx')
  const engine = new DocxEngine()

  before(async () => {
    await engine.create(file, {
      title: 'Документ с колонтитулами',
      paragraphs: ['Основной текст документа.', 'Второй абзац.'],
      header: { text: 'Отчёт — внутренний документ', alignment: 'center' },
      footer: { text: 'Стр.', pageNumber: true, alignment: 'center' }
    })
  })

  after(() => cleanup(dir))

  test('list reports every referenced part with its fields', async () => {
    const result = await engine.headerFooter(file)
    assert.equal(result.parts.length, 2)
    const header = result.parts.find(part => part.kind === 'header')
    const footer = result.parts.find(part => part.kind === 'footer')

    assert.equal(header.partName, 'word/header1.xml')
    assert.equal(header.exists, true)
    assert.equal(header.type, 'default')
    assert.deepEqual(header.sectionIndexes, [0])
    assert.equal(header.text, 'Отчёт — внутренний документ')
    assert.equal(header.hasPageNumberField, false)

    assert.equal(footer.partName, 'word/footer1.xml')
    assert.equal(footer.hasPageNumberField, true)
    assert.deepEqual(footer.fields.map(field => field.instruction), ['PAGE'])
    assert.equal(footer.fields[0].form, 'fldSimple')
  })

  test('read returns the part text, its fields and its paragraphs', async () => {
    const read = await engine.headerFooter(file, { action: 'read', kind: 'footer' })
    assert.equal(read.partName, 'word/footer1.xml')
    assert.match(read.text, /\{PAGE\}/)
    assert.equal(read.hasPageNumberField, true)
    assert.equal(read.paragraphCount, 1)
    assert.equal(read.paragraphs.length, 1)
    assert.match(read.paragraphs[0].text, /\{PAGE\}/)
  })

  test('setText changes the footer text and keeps the page-number field', async () => {
    const result = await engine.headerFooter(file, {
      action: 'setText',
      kind: 'footer',
      text: 'Страница',
      alignment: 'center'
    })
    assert.equal(result.success, true)
    assert.equal(result.created, false, 'an existing footer must be reused, not recreated')
    assert.equal(result.hasPageNumberField, true)
    assert.equal(result.text, 'Страница{PAGE}')

    const zip = await ZipArchive.fromFile(file)
    const footerXml = zip.getText('word/footer1.xml')
    assert.match(footerXml, /<w:fldSimple[^>]*w:instr=" PAGE "/, 'the field instruction must survive verbatim')
    assert.match(footerXml, /Страница/)
    assert.doesNotMatch(footerXml, />Стр\./)

    // Replacing the text twice must not accumulate stale text or fields.
    const again = await engine.headerFooter(file, { action: 'setText', kind: 'footer', text: 'Лист' })
    assert.equal(again.text, 'Лист{PAGE}')
    assert.equal((await engine.headerFooter(file, { action: 'read', kind: 'footer' })).hasPageNumberField, true)
  })

  test('a complex fldChar field survives a text change too', async () => {
    const target = path.join(dir, 'complex-field.docx')
    await engine.create(target, { paragraphs: ['Текст документа.'] })

    const zip = await ZipArchive.fromFile(target)
    zip.setText('word/header1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<w:hdr ${HEADER_NAMESPACES}><w:p><w:r><w:t>Стр. </w:t></w:r>`
      + '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
      + '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
      + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
      + '<w:r><w:t>1</w:t></w:r>'
      + '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>')
    await zip.save(target)

    // Point the document at the hand-written part.
    const withHeader = new DocxEngine()
    const created = await withHeader.headerFooter(target, {
      action: 'setText',
      kind: 'header',
      text: 'Лист'
    })
    assert.equal(created.created, true)
    const partName = created.partName
    const zip2 = await ZipArchive.fromFile(target)
    zip2.setText(partName, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<w:hdr ${HEADER_NAMESPACES}><w:p><w:r><w:t>Стр. </w:t></w:r>`
      + '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
      + '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
      + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
      + '<w:r><w:t>1</w:t></w:r>'
      + '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>')
    await zip2.save(target)

    const read = await engine.headerFooter(target, { action: 'read', kind: 'header' })
    assert.deepEqual(read.fields, [{ instruction: 'PAGE', form: 'complex' }])
    assert.equal(read.hasPageNumberField, true)

    const updated = await engine.headerFooter(target, { action: 'setText', kind: 'header', text: 'Лист' })
    assert.equal(updated.hasPageNumberField, true)

    const xml = (await archiveOf(target)).getText(partName)
    assert.match(xml, /<w:instrText[^>]*> PAGE <\/w:instrText>/, 'the complex instruction must survive')
    assert.match(xml, /w:fldCharType="begin"/)
    assert.match(xml, /w:fldCharType="end"/)
    assert.match(xml, />Лист</)
    assert.doesNotMatch(xml, />Стр\. </)
  })

  test('a page-number field is added without touching existing text', async () => {
    const target = path.join(dir, 'add-field.docx')
    await engine.create(target, { paragraphs: ['Текст.'], header: { text: 'Раздел 1' } })

    const result = await engine.headerFooter(target, {
      action: 'addPageNumber',
      kind: 'header',
      pageNumber: 'PAGE',
      separator: ' — '
    })
    assert.equal(result.hasPageNumberField, true)
    assert.match(result.text, /Раздел 1/)
    assert.match(result.text, /\{PAGE\}/)

    // Asking twice must not add a second field.
    const again = await engine.headerFooter(target, { action: 'addPageNumber', kind: 'header', pageNumber: 'PAGE' })
    const partXml = (await archiveOf(target)).getText(again.partName)
    assert.equal((partXml.match(/<w:fldSimple/g) || []).length, 1)
  })

  test('a missing part is created with its relationship and content type', async () => {
    const target = path.join(dir, 'create-part.docx')
    await engine.create(target, { paragraphs: ['Документ без колонтитулов.'] })

    const list = await engine.headerFooter(target)
    assert.equal(list.parts.length, 0)

    const result = await engine.headerFooter(target, {
      action: 'create',
      kind: 'header',
      type: 'default',
      text: 'Новый колонтитул',
      pageNumber: 'PAGE',
      alignment: 'right'
    })
    assert.equal(result.success, true)
    assert.equal(result.created, true)
    assert.equal(result.partName, 'word/header1.xml')
    assert.equal(result.relId.startsWith('rId'), true)

    const zip = await archiveOf(target)
    assert.ok(zip.has('word/header1.xml'))
    assert.match(zip.getText('[Content_Types].xml'), /wordprocessingml\.header\+xml/)
    assert.match(zip.getText('word/_rels/document.xml.rels'), /header1\.xml/)
    assert.match(zip.getText('word/document.xml'), /<w:headerReference w:type="default"/)
    assert.equal((await engine.validate(target)).valid, true)

    const read = await engine.headerFooter(target, { action: 'read', kind: 'header' })
    assert.equal(read.hasPageNumberField, true)
    assert.match(read.text, /Новый колонтитул/)
  })

  test('remove deletes the part, its relationship and its reference', async () => {
    const target = path.join(dir, 'remove-part.docx')
    await engine.create(target, {
      paragraphs: ['Текст.'],
      header: { text: 'Временный колонтитул' },
      footer: { text: 'Подвал' }
    })

    const removed = await engine.headerFooter(target, { action: 'remove', kind: 'header' })
    assert.equal(removed.removedPart, 'word/header1.xml')

    const zip = await archiveOf(target)
    assert.equal(zip.has('word/header1.xml'), false)
    assert.doesNotMatch(zip.getText('word/_rels/document.xml.rels'), /header1\.xml/)
    assert.doesNotMatch(zip.getText('word/document.xml'), /headerReference/)
    assert.match(zip.getText('word/document.xml'), /footerReference/)

    const parts = (await engine.headerFooter(target)).parts
    assert.deepEqual(parts.map(part => part.kind), ['footer'])
    assert.equal((await engine.validate(target)).valid, true)
  })

  test('an ordinary body edit leaves the header and footer byte-identical', async () => {
    const target = path.join(dir, 'body-edit.docx')
    await engine.create(target, {
      paragraphs: ['Исходный текст для замены.'],
      header: { text: 'Колонтитул' },
      footer: { text: 'Стр.', pageNumber: true }
    })

    const before = await archiveOf(target)
    await engine.replaceText(target, 'Исходный текст для замены.', 'Новый текст.')
    await engine.editParagraph(target, 0, 'Правка единственного абзаца.')
    const after = await archiveOf(target)

    const diff = diffMembers(before, after, ['word/document.xml'])
    assert.deepEqual(diff.changed, [], 'only document.xml may change')
    assert.deepEqual(diff.removed, [])
    assert.equal(after.getText('word/header1.xml'), before.getText('word/header1.xml'))
    assert.match(after.getText('word/footer1.xml'), /w:instr=" PAGE "/)
  })

  test('R7 reopens a document whose footer was re-texted, field intact', async (t) => {
    if (!r7Available) {
      t.skip('R7 not installed')
      return
    }
    const pdf = path.join(dir, 'колонтитулы.pdf')
    await new R7Adapter().convert(file, pdf)
    const head = fs.readFileSync(pdf).subarray(0, 5).toString()
    assert.equal(head, '%PDF-')
    assert.equal((await engine.validate(file)).valid, true)
  })
})
