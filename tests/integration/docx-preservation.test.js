import { test, describe, before, after } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import { DocxEngine } from '../../src/r7/docx.js'
import { R7Adapter } from '../../src/r7/adapter.js'
import { ZipArchive } from '../../src/shared/zip.js'
import { archiveOf, cleanup, diffMembers, makePng, sameBytes, tempDir } from './docx-fixtures.test.js'

const r7Available = (await new R7Adapter().detect()).installed

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
const R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

/**
 * Markup that belongs to features this engine deliberately does not implement
 * and must therefore never damage: comments, tracked changes, footnotes, a
 * table of contents, an equation, an embedded object and a complex field.
 */
const EXOTIC_PARAGRAPHS = [
  '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t xml:space="preserve">Абзац с примечанием.</w:t></w:r>'
    + '<w:r><w:commentReference w:id="1"/></w:r></w:p>',
  '<w:p><w:ins w:id="7" w:author="Автор" w:date="2026-01-01T00:00:00Z">'
    + '<w:r><w:t xml:space="preserve">Вставленный текст.</w:t></w:r></w:ins>'
    + '<w:del w:id="8" w:author="Автор" w:date="2026-01-01T00:00:00Z">'
    + '<w:r><w:delText xml:space="preserve">Удалённый текст.</w:delText></w:r></w:del></w:p>',
  '<w:p><w:r><w:t xml:space="preserve">Сноска</w:t></w:r>'
    + '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p>',
  '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    + '<w:r><w:instrText xml:space="preserve"> TOC \\o &quot;1-3&quot; \\h </w:instrText></w:r>'
    + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    + '<w:r><w:t xml:space="preserve">Содержание появится после обновления поля.</w:t></w:r>'
    + '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
  '<w:p><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
    + '<m:r><m:t>x</m:t></m:r><m:r><m:t>=</m:t></m:r><m:r><m:t>1</m:t></m:r></m:oMath></w:p>',
  '<w:p><w:r><w:object w:dxaOrig="1440" w:dyaOrig="1440">'
    + '<o:OLEObject xmlns:o="urn:schemas-microsoft-com:office:office" Type="Embed" ProgID="Excel.Sheet.12" '
    + 'ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1" r:id="rIdOle1"/>'
    + '</w:object></w:r></w:p>',
  '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    + '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
    + '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
]

const EXOTIC_MARKERS = [
  '<w:commentRangeStart w:id="1"/>',
  '<w:commentReference w:id="1"/>',
  '<w:ins w:id="7"',
  '<w:del w:id="8"',
  '<w:delText xml:space="preserve">Удалённый текст.</w:delText>',
  '<w:footnoteReference w:id="2"/>',
  '<w:instrText xml:space="preserve"> TOC \\o &quot;1-3&quot; \\h </w:instrText>',
  '<m:oMath',
  '<o:OLEObject',
  'ProgID="Excel.Sheet.12"'
]

/**
 * Author a document that carries every out-of-scope feature the brief names,
 * by adding the parts and the markup the engine must preserve verbatim.
 *
 * @param {DocxEngine} engine
 * @param {string} file
 */
async function authorDocumentWithExoticParts(engine, file) {
  await engine.create(file, {
    title: 'Полный документ',
    paragraphs: ['Первый абзац.', 'Целевой абзац для замены.', 'Абзац для правки.', 'Последний абзац.'],
    tables: [{ rows: [['A', 'B'], ['C', 'D']], widthsTwips: [2000, 2000] }],
    header: { text: 'Колонтитул' },
    footer: { text: 'Стр.', pageNumber: true }
  })

  const zip = await ZipArchive.fromFile(file)
  const docXml = zip.getText('word/document.xml')
  const sections = docXml.lastIndexOf('<w:sectPr')
  zip.setText('word/document.xml', docXml.slice(0, sections) + EXOTIC_PARAGRAPHS.join('') + docXml.slice(sections))

  zip.setText('word/comments.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments ${W_NS}><w:comment w:id="1" w:author="Автор" w:date="2026-01-01T00:00:00Z">`
    + '<w:p><w:r><w:t>Комментарий рецензента.</w:t></w:r></w:p></w:comment></w:comments>')

  zip.setText('word/footnotes.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes ${W_NS}><w:footnote w:id="0"><w:p><w:r><w:t>Сноска автора.</w:t></w:r></w:p></w:footnote>`
    + '<w:footnote w:id="2"><w:p><w:r><w:t>Текст сноски.</w:t></w:r></w:p></w:footnote></w:footnotes>')

  zip.setBuffer('word/embeddings/oleObject1.bin', Buffer.from('embedded object placeholder', 'utf8'))

  const rels = zip.getText('word/_rels/document.xml.rels')
  zip.setText('word/_rels/document.xml.rels', rels
    .replace('</Relationships>', '<Relationship Id="rIdComments" '
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" '
      + `Target="comments.xml"/><Relationship Id="rIdFootnotes" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" '
      + `Target="footnotes.xml"/><Relationship Id="rIdOle1" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" '
      + 'Target="embeddings/oleObject1.bin" TargetMode="Internal"/></Relationships>'))

  const contentTypes = zip.getText('[Content_Types].xml')
  zip.setText('[Content_Types].xml', contentTypes.replace('</Types>',
    '<Override PartName="/word/comments.xml" '
    + 'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
    + '<Override PartName="/word/footnotes.xml" '
    + 'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
    + '<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>'
    + '</Types>'))

  await zip.save(file)
  return file
}

/** Assert that every marker is still present verbatim. */
function assertMarkersIntact(docXml, markers = EXOTIC_MARKERS) {
  for (const marker of markers) {
    assert.ok(docXml.includes(marker), `markup must survive verbatim: ${marker}`)
  }
}

describe('DOCX editing never damages what it did not change', () => {
  const dir = tempDir('docx_preserve')
  const engine = new DocxEngine()

  before(() => {
    fs.mkdirSync(dir, { recursive: true })
  })

  after(() => cleanup(dir))

  test('the full edit flow leaves every untouched part byte-identical', async () => {
    const source = path.join(dir, 'flow-source.docx')
    await engine.create(source, {
      title: 'Отчёт за квартал',
      paragraphs: [
        { text: 'Раздел 1. Введение', style: 'Heading1' },
        'Вводный абзац отчёта.',
        { text: 'Раздел 2. Показатели', style: 'Heading1' },
        'Абзац с показателями, который надо обновить.',
        { text: 'Пункт списка', list: 'bullet' }
      ],
      tables: [{ rows: [['Показатель', 'Значение'], ['Выручка', '120.5']], widthsTwips: [4000, 3000] }],
      header: { text: 'Внутренний документ', alignment: 'center' },
      footer: { text: 'Стр.', pageNumber: true, alignment: 'center' }
    })
    await engine.insertImage(source, {
      buffer: makePng(4, 2),
      widthCm: 4,
      paragraphIndex: 'new',
      alt: 'Диаграмма'
    })

    // 1. inspect
    const inspected = await engine.inspect(source, { includeStyles: true, includeHeadersFooters: true })
    assert.ok(inspected.paragraphsCount > 5)
    assert.equal(inspected.tablesCount, 1)

    // 2. read
    const markdown = await engine.read(source, { format: 'markdown' })
    assert.ok(markdown.content.includes('Отчёт за квартал'))

    // 3. change one paragraph, 4. replace text, 5. change one table cell —
    //    all into a COPY.
    const copy = path.join(dir, 'flow-copy.docx')
    const targetIndex = (await engine.read(source)).paragraphs
      .find(entry => entry.text === 'Абзац с показателями, который надо обновить.').index

    await engine.editParagraph(source, targetIndex, 'Обновлённый абзац показателей.', { outputPath: copy })
    await engine.replaceText(copy, 'Вводный абзац отчёта.', 'Новый вводный абзац.')
    await engine.table(copy, { action: 'setCell', tableIndex: 0, cell: { row: 1, col: 1, value: '150.8' } })

    // 6. reopen and verify only document.xml moved.
    const before = await archiveOf(source)
    const after = await archiveOf(copy)
    const diff = diffMembers(before, after, ['word/document.xml'])
    assert.deepEqual(diff.changed, [], 'only document.xml may differ')
    assert.deepEqual(diff.added, [])
    assert.deepEqual(diff.removed, [])
    assert.ok(sameBytes(after.getBuffer('word/media/image1.png'), before.getBuffer('word/media/image1.png')))
    assert.match(after.getText('word/footer1.xml'), /w:instr=" PAGE "/)

    const reopened = await engine.read(copy)
    const texts = reopened.paragraphs.map(entry => entry.text)
    assert.ok(texts.includes('Обновлённый абзац показателей.'))
    assert.ok(texts.includes('Новый вводный абзац.'))
    assert.ok(!texts.includes('Вводный абзац отчёта.'))

    const table = await engine.table(copy, { action: 'inspect', tableIndex: 0 })
    assert.deepEqual(table.data, [['Показатель', 'Значение'], ['Выручка', '150.8']])
    assert.equal((await engine.validate(copy)).valid, true)

    // The original file is untouched by an edit that named another output.
    const original = await engine.read(source)
    assert.ok(original.paragraphs.some(entry => entry.text === 'Вводный абзац отчёта.'))
  })

  test('comments, tracked changes, footnotes, a TOC, an equation and an OLE object survive an ordinary edit', async () => {
    const source = path.join(dir, 'exotic.docx')
    await authorDocumentWithExoticParts(engine, source)

    const before = await archiveOf(source)
    const documentXmlBefore = before.getText('word/document.xml')
    assertMarkersIntact(documentXmlBefore)

    const copy = path.join(dir, 'exotic-copy.docx')
    const targetIndex = (await engine.read(source)).paragraphs
      .find(entry => entry.text === 'Целевой абзац для замены.').index
    const editIndex = (await engine.read(source)).paragraphs
      .find(entry => entry.text === 'Абзац для правки.').index

    await engine.editParagraph(source, targetIndex, 'Заменённый целевой абзац.', { outputPath: copy })
    await engine.replaceText(copy, 'Первый абзац.', 'Изменённый первый абзац.')
    await engine.editParagraph(copy, editIndex, 'Отредактированный абзац для правки.')
    await engine.table(copy, { action: 'setCell', tableIndex: 0, cell: { row: 0, col: 1, value: 'B2' } })

    const after = await archiveOf(copy)

    // Every part other than the main document must be byte-identical.
    const diff = diffMembers(before, after, ['word/document.xml'])
    assert.deepEqual(diff.changed, [], 'no other part may change')
    assert.deepEqual(diff.added, [])
    assert.deepEqual(diff.removed, [])
    assert.ok(sameBytes(after.getBuffer('word/comments.xml'), before.getBuffer('word/comments.xml')))
    assert.ok(sameBytes(after.getBuffer('word/footnotes.xml'), before.getBuffer('word/footnotes.xml')))
    assert.ok(sameBytes(after.getBuffer('word/embeddings/oleObject1.bin'), before.getBuffer('word/embeddings/oleObject1.bin')))

    // …and the out-of-scope markup inside the document itself is untouched.
    const documentXmlAfter = after.getText('word/document.xml')
    assertMarkersIntact(documentXmlAfter)
    assert.ok(documentXmlAfter.includes('Изменённый первый абзац.'), 'the edit itself landed')
    assert.equal((await engine.validate(copy)).valid, true)

    // Reading the document still reports the edited text and nothing broken.
    const read = await engine.read(copy)
    assert.ok(read.paragraphs.some(entry => entry.text === 'Отредактированный абзац для правки.'))
  })

  test('a document carrying those parts round-trips byte-for-byte when nothing is edited', async () => {
    const source = path.join(dir, 'exotic-readwrite.docx')
    await authorDocumentWithExoticParts(engine, source)

    const original = fs.readFileSync(source)
    const zip = await ZipArchive.fromFile(source)
    const rewritten = path.join(dir, 'exotic-readwrite-2.docx')
    await zip.save(rewritten)

    assert.ok(sameBytes(original, fs.readFileSync(rewritten)),
      'opening and writing a package back must not alter a single byte')
  })

  test('editing one paragraph leaves the header, footer, numbering and theme parts frozen', async () => {
    const source = path.join(dir, 'frozen.docx')
    await engine.create(source, {
      paragraphs: ['Первый.', { text: 'Второй пункт', list: 'bullet' }],
      header: { text: 'Колонтитул' },
      footer: { text: 'Стр.', pageNumber: true }
    })

    const before = await archiveOf(source)
    await engine.editParagraph(source, 0, 'Первый абзац изменён.')
    const after = await archiveOf(source)

    const diff = diffMembers(before, after, ['word/document.xml'])
    assert.deepEqual(diff.changed, [])
    for (const name of ['word/header1.xml', 'word/footer1.xml', 'word/numbering.xml', 'word/theme/theme1.xml']) {
      if (before.has(name)) assert.ok(sameBytes(before.getBuffer(name), after.getBuffer(name)), `${name} must be frozen`)
    }
  })

  test('the R7-authored document, when one is supplied, goes through the same flow', async (t) => {
    const authored = process.env.R7_AUTHORED_DOCX
    if (!authored || !fs.existsSync(authored)) {
      t.skip('set R7_AUTHORED_DOCX to a document authored in R7-Office Desktop to run this check')
      return
    }

    const before = await archiveOf(authored)
    const copy = path.join(dir, 'r7-authored-copy.docx')

    const inspection = await engine.inspect(authored)
    assert.ok(inspection.paragraphsCount > 0)

    const paragraphs = (await engine.read(authored)).paragraphs
    const firstText = paragraphs.find(entry => entry.text.trim().length > 3)
    assert.ok(firstText, 'the supplied document must carry readable text')

    await engine.editParagraph(authored, firstText.index, `${firstText.text} (правка)`, { outputPath: copy })
    await engine.table(copy, {
      action: 'setCell',
      tableIndex: 0,
      cell: { row: 0, col: 0, value: 'Обновлённая ячейка' }
    })

    const after = await archiveOf(copy)
    const diff = diffMembers(before, after, ['word/document.xml'])
    assert.deepEqual(diff.changed, [], 'nothing but document.xml may change in an R7-authored file')
    assert.deepEqual(diff.added, [])
    assert.deepEqual(diff.removed, [])

    const reopened = await engine.read(copy)
    assert.ok(reopened.paragraphs.some(entry => entry.text.includes('(правка)')))
    assert.equal((await engine.validate(copy)).valid, true)
  })

  test('R7 reopens the edited copy and renders it to PDF', async (t) => {
    if (!r7Available) {
      t.skip('R7 not installed')
      return
    }
    const source = path.join(dir, 'render-source.docx')
    await engine.create(source, {
      title: 'Отчёт для рендера',
      paragraphs: [
        { text: 'Раздел', style: 'Heading1' },
        { text: 'Пункт', list: 'bullet' },
        { text: 'Нумерованный', list: 'number' }
      ],
      tables: [{ rows: [['A', 'B'], ['1', '2']] }],
      header: { text: 'Колонтитул' },
      footer: { text: 'Стр.', pageNumber: true }
    })
    await engine.insertImage(source, { buffer: makePng(8, 4), widthCm: 5, paragraphIndex: 'new' })
    await engine.hyperlink(source, { action: 'insert', text: 'Сайт', url: 'https://r7-office.ru/', newParagraph: true })
    await engine.setSection(source, { orientation: 'landscape', margins: { left: 2 } })

    const pdf = path.join(dir, 'render-source.pdf')
    await new R7Adapter().convert(source, pdf)
    assert.equal(fs.readFileSync(pdf).subarray(0, 5).toString(), '%PDF-')
    assert.ok(fs.statSync(pdf).size > 1000)
  })
})
