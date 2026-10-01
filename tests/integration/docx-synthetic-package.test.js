import { test, describe, before, after } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import { DocxEngine } from '../../src/r7/docx.js'
import { ZipArchive } from '../../src/shared/zip.js'
import { cleanup, tempDir } from './docx-fixtures.test.js'

/**
 * The synthetic package: what the engine builds when no R7 template is
 * installed (`_createFallbackBlankDocx`). A machine without R7-Office — every
 * CI runner — takes this path on every `create()`, so it is exercised here
 * explicitly rather than merely tolerated as a fallback.
 *
 * The engine is handed an adapter that reports "no template", which forces the
 * synthetic path on every host, R7 installed or not. That is why this file
 * needs no skip: it asserts what the engine itself decides, never what R7's
 * template happens to contain. The XLSX suite pins its own no-template workbook
 * the same way.
 */
describe('DOCX package built without an R7 template', () => {
  const noTemplateAdapter = { getTemplatePath: async () => null }
  const engine = new DocxEngine(noTemplateAdapter)
  const dir = tempDir('docx_synthetic')
  const file = path.join(dir, 'синтетика.docx')

  before(async () => {
    await engine.create(file, {
      title: 'Заголовок документа',
      paragraphs: ['Первый абзац.', 'Второй абзац.']
    })
  })

  after(() => cleanup(dir))

  test('create() writes a minimal package that validates on its own', async () => {
    assert.equal(fs.existsSync(file), true, 'the file was written')

    const zip = await ZipArchive.fromFile(file)
    for (const member of ['[Content_Types].xml', '_rels/.rels', 'word/document.xml', 'word/_rels/document.xml.rels']) {
      assert.ok(zip.has(member), `${member} must be present in the hand-built package`)
    }
    // The package describes the main part it actually contains, both as a
    // content-type override and as a package relationship.
    assert.match(zip.getText('[Content_Types].xml'), /PartName="\/word\/document\.xml"/)
    assert.match(zip.getText('_rels/.rels'), /Target="word\/document\.xml"/)
    assert.match(zip.getText('word/document.xml'), /<w:body>[\s\S]*<\/w:body>/)

    const validation = await engine.validate(file)
    assert.equal(validation.valid, true, `validate() rejected the package: ${validation.errors.join('; ')}`)
  })

  test('the body carries the title and both paragraphs', async () => {
    const read = await engine.read(file)
    assert.deepEqual(
      read.paragraphs.map(paragraph => paragraph.text),
      ['Заголовок документа', 'Первый абзац.', 'Второй абзац.']
    )
  })

  test('the section it declares is the A4 portrait the engine assumes', async () => {
    const result = await engine.sections(file)
    assert.equal(result.sectionCount, 1, 'the hand-built package declares one final section')
    const section = result.sections[0]
    assert.equal(section.kind, 'final')
    assert.equal(section.paragraphIndex, null)
    // The synthetic `<w:sectPr/>` states no `w:pgSz`, so the engine falls back
    // to A4 portrait and says so.
    assert.equal(section.pageSize.assumed, true)
    assert.equal(section.pageSize.orientation, 'portrait')
    assert.equal(section.pageSize.widthCm, 21)
    assert.equal(section.pageSize.heightCm, 29.7)
    // Nothing set a margin either, and absence is reported as absence.
    assert.equal(section.margins.top.twips, null)
    assert.equal(section.margins.top.cm, null)
  })

  test('setSection writes the geometry the caller asked for, from scratch', async () => {
    const set = await engine.setSection(file, {
      orientation: 'landscape',
      margins: { top: 2.5, right: 1.5, bottom: 2.5, left: 1.5, header: 1.25, footer: 1.25 }
    })
    assert.equal(set.section.pageSize.orientation, 'landscape')
    assert.equal(set.section.pageSize.widthCm, 29.7)
    assert.equal(set.section.pageSize.heightCm, 21)
    assert.equal(set.section.pageSize.assumed, false, 'the page size is declared from now on')
    assert.equal(set.section.margins.top.cm, 2.5)
    assert.equal(set.section.margins.left.cm, 1.5)
    assert.equal(set.section.margins.header.cm, 1.25)

    // A fresh open must see the same geometry, not just the return value.
    const reread = (await engine.sections(file)).sections[0]
    assert.equal(reread.pageSize.orientation, 'landscape')
    assert.equal(reread.pageSize.heightCm, 21)
    assert.equal(reread.margins.left.cm, 1.5)
    assert.equal(reread.pageSize.assumed, false)
  })

  test('a basic edit round-trip keeps the package valid and readable', async () => {
    const replaced = await engine.replaceText(file, 'Первый абзац.', 'Изменённый абзац.')
    assert.ok(replaced.matchesCount > 0, 'the replace found its text')

    await engine.insert(file, { position: 'end', text: 'Добавленный абзац.' })

    const texts = (await engine.read(file)).paragraphs.map(paragraph => paragraph.text)
    assert.ok(texts.includes('Изменённый абзац.'), 'the replacement is visible')
    assert.ok(!texts.includes('Первый абзац.'), 'the original text is gone')
    assert.equal(texts.at(-1), 'Добавленный абзац.', 'the inserted paragraph is last')

    const validation = await engine.validate(file)
    assert.equal(validation.valid, true, `validate() rejected the edited package: ${validation.errors.join('; ')}`)
  })

  test('the hand-built package has no styles part, so the styles it references are undefined', async () => {
    // `ensureStyleDefinitions` can only append to an existing `word/styles.xml`;
    // it returns unchanged when there is none. The fallback creates no styles
    // part, so on a machine without R7 a created document references Heading1
    // (and, once a hyperlink is inserted, Hyperlink) with nothing to define it.
    // This test is the tripwire for that boundary: a fallback that starts
    // shipping a styles part fails here, and this suite must be revisited.
    const zip = await ZipArchive.fromFile(file)
    assert.equal(zip.has('word/styles.xml'), false, 'the fallback package ships no styles part')
    assert.match(zip.getText('word/document.xml'), /<w:pStyle w:val="Heading1"\/>/)
  })
})
