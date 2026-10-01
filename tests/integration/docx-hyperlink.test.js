import { test, describe, before, after } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import { DocxEngine } from '../../src/r7/docx.js'
import { R7Adapter } from '../../src/r7/adapter.js'
import { HAS_R7_TEMPLATES, requiresR7 } from '../helpers/r7-gate.js'
import { archiveOf, cleanup, diffMembers, tempDir } from './docx-fixtures.test.js'

/** How many hyperlink relationships an archive declares. */
function hyperlinkRelationships(zip) {
  const rels = zip.getText('word/_rels/document.xml.rels')
  return (rels.match(/relationships\/hyperlink/g) || []).length
}

/**
 * Scenario 6: read, insert, re-text and remove hyperlinks, keeping the
 * document/relationship pair in step.
 */
describe('DOCX hyperlinks', () => {
  const dir = tempDir('docx_link')
  const file = path.join(dir, 'ссылки.docx')
  const engine = new DocxEngine()

  before(async () => {
    await engine.create(file, {
      title: 'Документ со ссылками',
      paragraphs: ['Внешняя ссылка: ', 'Второй абзац с текстом.']
    })
  })

  after(() => cleanup(dir))

  test('inserting a link writes an external relationship and a Hyperlink run', async () => {
    const result = await engine.hyperlink(file, {
      action: 'insert',
      text: 'R7-Office',
      url: 'https://r7-office.ru/',
      paragraphIndex: 1
    })
    assert.equal(result.success, true)
    assert.match(result.relId, /^rId\d+$/)
    assert.equal(result.paragraphIndex, 1)

    const zip = await archiveOf(file)
    const rels = zip.getText('word/_rels/document.xml.rels')
    assert.match(rels, new RegExp(`Id="${result.relId}"[^>]*Type="[^"]*relationships/hyperlink"`))
    assert.match(rels, /Target="https:\/\/r7-office\.ru\/" TargetMode="External"/)
    assert.match(rels, /Target="https:\/\/r7-office\.ru\/"/)

    const docXml = zip.getText('word/document.xml')
    assert.match(docXml, new RegExp(`<w:hyperlink r:id="${result.relId}">`))
    assert.match(docXml, /<w:rStyle w:val="Hyperlink"\/>/)
    // The run references the Hyperlink character style, and defining it is the
    // engine's job, not the template's: R7's own styles.xml ships no such style,
    // so `ensureStyleDefinitions` appends one on the R7 path. The engine can
    // only extend a styles part that already exists, though, and the package it
    // builds without an R7 template has none — there the reference dangles.
    // `docx-synthetic-package.test.js` exercises that boundary on its own.
    const stylesXml = zip.getText('word/styles.xml')
    if (HAS_R7_TEMPLATES) {
      assert.match(stylesXml, /w:styleId="Hyperlink"/)
    } else {
      assert.equal(stylesXml, null, 'the no-template package has no styles part to append the style to')
    }
  })

  test('the link is readable back with its text and target', async () => {
    const list = await engine.hyperlink(file, { action: 'list' })
    assert.equal(list.hyperlinks.length, 1)
    const link = list.hyperlinks[0]
    assert.equal(link.text, 'R7-Office')
    assert.equal(link.url, 'https://r7-office.ru/')
    assert.equal(link.runStyle, 'Hyperlink')
    assert.equal(link.paragraphIndex, 1)
    assert.equal(link.anchor, null)
  })

  test('re-texting a link keeps its relationship, target and formatting', async () => {
    const before = await engine.hyperlink(file, { action: 'list' })
    const relId = before.hyperlinks[0].relId

    const result = await engine.hyperlink(file, { action: 'setText', index: 0, text: 'Сайт R7-Office' })
    assert.equal(result.index, 0)
    assert.equal(result.url, 'https://r7-office.ru/')

    const after = await engine.hyperlink(file, { action: 'list' })
    assert.equal(after.hyperlinks.length, 1)
    assert.equal(after.hyperlinks[0].text, 'Сайт R7-Office')
    assert.equal(after.hyperlinks[0].relId, relId, 'the relationship must not be recreated')
    assert.equal(after.hyperlinks[0].runStyle, 'Hyperlink', 'the run formatting must survive')

    const zip = await archiveOf(file)
    assert.match(zip.getText('word/document.xml'), new RegExp(`r:id="${relId}"`))
    assert.equal(hyperlinkRelationships(zip), 1)
  })

  test('the same URL reuses its relationship instead of duplicating it', async () => {
    const before = await engine.hyperlink(file, { action: 'list' })
    const relId = before.hyperlinks[0].relId

    const inserted = await engine.hyperlink(file, {
      action: 'insert',
      text: 'Ещё раз',
      url: 'https://r7-office.ru/',
      paragraphIndex: 'end'
    })
    assert.equal(inserted.relId, relId, 'an identical target must reuse the relationship')

    const zip = await archiveOf(file)
    assert.equal(hyperlinkRelationships(zip), 1)
    assert.equal((await engine.hyperlink(file, { action: 'list' })).hyperlinks.length, 2,
      'both links are still separate hyperlink elements')
  })

  test('a different URL gets its own relationship', async () => {
    await engine.hyperlink(file, {
      action: 'insert',
      text: 'Документация',
      url: 'https://docs.r7-office.ru/',
      newParagraph: true
    })
    const zip = await archiveOf(file)
    assert.equal(hyperlinkRelationships(zip), 2)
    const links = (await engine.hyperlink(file, { action: 'list' })).hyperlinks
    assert.deepEqual(links.map(link => link.url).sort(), [
      'https://docs.r7-office.ru/',
      'https://r7-office.ru/',
      'https://r7-office.ru/'
    ].sort())
  })

  test('a link can be inserted as its own paragraph, even after a table', async () => {
    const target = path.join(dir, 'linked-after-table.docx')
    await engine.create(target, {
      paragraphs: ['Перед таблицей.'],
      tables: [{ rows: [['A', 'B'], ['C', 'D']] }]
    })

    const result = await engine.hyperlink(target, {
      action: 'insert',
      text: 'Ссылка после таблицы',
      url: 'https://example.com/',
      newParagraph: true,
      alignment: 'center'
    })
    const zip = await archiveOf(target)
    const docXml = zip.getText('word/document.xml')
    const tableEnd = docXml.indexOf('</w:tbl>')
    const linkAt = docXml.indexOf('<w:hyperlink')
    assert.ok(linkAt > tableEnd, 'the link paragraph must be outside the table')
    assert.match(docXml, /<w:p><w:pPr><w:jc w:val="center"\/><\/w:pPr><w:hyperlink/)

    const link = (await engine.hyperlink(target, { action: 'list' })).hyperlinks.find(entry => entry.index === 0)
    assert.equal(link.paragraphIndex, result.paragraphIndex)
    assert.equal(link.url, 'https://example.com/')
  })

  test('an internal anchor needs no relationship', async () => {
    const target = path.join(dir, 'anchor.docx')
    await engine.create(target, { paragraphs: ['Смотри раздел'] })

    const result = await engine.hyperlink(target, {
      action: 'insert',
      text: 'раздел 2',
      anchor: 'section2',
      paragraphIndex: 0
    })
    assert.equal(result.relId, null)

    const link = (await engine.hyperlink(target, { action: 'list' })).hyperlinks[0]
    assert.equal(link.anchor, 'section2')
    assert.equal(link.url, null)
    assert.equal((await engine.validate(target)).valid, true)
  })

  test('removing a link keeps its visible text as a plain run', async () => {
    const target = path.join(dir, 'remove-link.docx')
    await engine.create(target, { paragraphs: ['Ссылка: '] })
    await engine.hyperlink(target, {
      action: 'insert',
      text: 'Удаляемая ссылка',
      url: 'https://removed.example/',
      paragraphIndex: 0
    })

    const removed = await engine.hyperlink(target, { action: 'remove', index: 0 })
    assert.equal(removed.index, 0)

    const zip = await archiveOf(target)
    const docXml = zip.getText('word/document.xml')
    assert.doesNotMatch(docXml, /<w:hyperlink/)
    assert.match(docXml, /Удаляемая ссылка/, 'the text stays in the document')
    assert.equal((await engine.hyperlink(target, { action: 'list' })).hyperlinks.length, 0)
    assert.equal((await engine.validate(target)).valid, true)
  })

  test('an existing link survives inserting and editing another one', async () => {
    const target = path.join(dir, 'two-links.docx')
    await engine.create(target, { paragraphs: ['Ссылки: '] })
    const first = await engine.hyperlink(target, {
      action: 'insert',
      text: 'Первая',
      url: 'https://first.example/',
      paragraphIndex: 0
    })
    await engine.hyperlink(target, {
      action: 'insert',
      text: 'Вторая',
      url: 'https://second.example/',
      newParagraph: true
    })

    await engine.hyperlink(target, { action: 'setText', index: 1, text: 'Вторая правленая' })
    const links = (await engine.hyperlink(target, { action: 'list' })).hyperlinks
    assert.equal(links.length, 2)
    assert.equal(links[0].text, 'Первая')
    assert.equal(links[0].relId, first.relId)
    assert.equal(links[0].url, 'https://first.example/')
    assert.equal(links[1].text, 'Вторая правленая')
  })

  test('a link needs text and a target, and an unknown index is refused', async () => {
    await assert.rejects(
      () => engine.hyperlink(file, { action: 'insert', text: 'Без цели' }),
      /needs a url or an anchor/
    )
    await assert.rejects(
      () => engine.hyperlink(file, { action: 'insert', url: 'https://x.example/' }),
      /needs text/
    )
    await assert.rejects(() => engine.hyperlink(file, { action: 'setText', index: 99, text: 'x' }), /not found/)
    await assert.rejects(() => engine.hyperlink(file, { action: 'nope' }), /Unsupported hyperlink action/)
  })

  test('inserting a link changes only the document, its relationships and its styles', async () => {
    const target = path.join(dir, 'byte-diff.docx')
    await engine.create(target, { paragraphs: ['Текст со ссылкой позже.'] })

    const before = await archiveOf(target)
    await engine.hyperlink(target, {
      action: 'insert',
      text: 'Ссылка',
      url: 'https://byte-diff.example/',
      paragraphIndex: 0
    })
    const after = await archiveOf(target)

    const diff = diffMembers(before, after, [
      'word/document.xml',
      'word/_rels/document.xml.rels',
      'word/styles.xml'
    ])
    assert.deepEqual(diff.changed, [], 'no other part may change')
    assert.deepEqual(diff.added, [])
    assert.deepEqual(diff.removed, [])
  })

  test('R7 reopens the document with links and renders it', async (t) => {
    if (requiresR7(t)) return
    const pdf = path.join(dir, 'ссылки.pdf')
    await new R7Adapter().convert(file, pdf)
    assert.equal(fs.readFileSync(pdf).subarray(0, 5).toString(), '%PDF-')
    assert.equal((await engine.validate(file)).valid, true)
  })
})
