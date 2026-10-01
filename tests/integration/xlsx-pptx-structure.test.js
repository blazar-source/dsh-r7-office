/**
 * Structure-level tests for the XLSX and PPTX engines.
 *
 * These cover the operations that change a workbook's or a presentation's
 * *part graph* (adding a worksheet, appending a slide) rather than the content
 * of an existing part. Getting this wrong is how a document silently loses
 * data or becomes unopenable, so every case here asserts both the new
 * structure and that previously existing parts were left alone.
 */

import { test, describe, before, after } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import { XlsxEngine } from '../../src/r7/xlsx.js'
import { PptxEngine } from '../../src/r7/pptx.js'
import { DocxEngine } from '../../src/r7/docx.js'
import { R7Adapter } from '../../src/r7/adapter.js'
import { ZipArchive } from '../../src/shared/zip.js'
import { requiresR7 } from '../helpers/r7-gate.js'

const tmpDir = path.join(os.tmpdir(), `dsh_r7_struct_${Date.now()}`)

function sameBytes(a, b) {
  if (!Buffer.isBuffer(a) || !Buffer.isBuffer(b)) return false
  return Buffer.compare(a, b) === 0
}

function looksLikePdf(filePath) {
  return fs.readFileSync(filePath).subarray(0, 1024).includes(Buffer.from('%PDF-'))
}

describe('Workbook and presentation structure', () => {
  let adapter

  before(async () => {
    fs.mkdirSync(tmpDir, { recursive: true })
    adapter = new R7Adapter()
  })

  after(() => {
    try { fs.rmSync(tmpDir, { recursive: true, force: true }) } catch { /* ignore */ }
  })

  describe('XLSX: multiple worksheets', () => {
    const book = path.join(tmpDir, 'multi.xlsx')

    test('r7_create honours every sheet name and all sheet data', async () => {
      const xlsx = new XlsxEngine(adapter)
      await xlsx.create(book, {
        sheets: [
          { name: 'Доходы', data: [['Статья', 'Сумма'], ['Продажи', 1000]] },
          { name: 'Расходы', data: [['Статья', 'Сумма'], ['Аренда', 400]] },
          { name: 'Итоги', data: [['Показатель', 'Значение'], ['Прибыль', 600]] }
        ]
      })

      const info = await xlsx.inspect(book)
      assert.equal(info.sheetsCount, 3, 'three sheets exist')
      assert.deepEqual(info.sheets.map(s => s.name), ['Доходы', 'Расходы', 'Итоги'])

      // Every sheet must carry its own data, not the first sheet's data.
      const first = await xlsx.read(book, { sheetName: 'Доходы', range: 'A1:B2' })
      assert.equal(first.data[0][0], 'Статья')
      assert.equal(first.data[1][0], 'Продажи')

      const second = await xlsx.read(book, { sheetName: 'Расходы', range: 'A1:B2' })
      assert.equal(second.data[1][0], 'Аренда')

      const third = await xlsx.read(book, { sheetName: 'Итоги', range: 'A1:B2' })
      assert.equal(third.data[1][0], 'Прибыль')
    })

    test('the produced workbook is a valid package', async () => {
      const xlsx = new XlsxEngine(adapter)
      const validation = await xlsx.validate(book)
      assert.equal(validation.valid, true, JSON.stringify(validation.errors))
    })

    test('R7 accepts the multi-sheet workbook', async (t) => {
      if (requiresR7(t)) return
      const pdf = path.join(tmpDir, 'multi.pdf')
      await adapter.convert(book, pdf)
      assert.ok(looksLikePdf(pdf), 'x2t rendered the workbook')
    })
  })

  describe('XLSX: adding a worksheet to an existing workbook', () => {
    const book = path.join(tmpDir, 'append.xlsx')

    test('setup: a one-sheet workbook', async () => {
      const xlsx = new XlsxEngine(adapter)
      await xlsx.create(book, {
        sheets: [{ name: 'Основной', data: [['A', 'B'], [1, 2]] }]
      })
      assert.equal((await xlsx.inspect(book)).sheetsCount, 1)
    })

    test('addSheet appends a named worksheet', async () => {
      const xlsx = new XlsxEngine(adapter)
      const res = await xlsx.addSheet(book, { name: 'Дополнительный' })
      assert.equal(res.success, true)

      const info = await xlsx.inspect(book)
      assert.equal(info.sheetsCount, 2)
      assert.deepEqual(info.sheets.map(s => s.name), ['Основной', 'Дополнительный'])
    })

    test('the existing worksheet keeps its data and its tab', async () => {
      const xlsx = new XlsxEngine(adapter)
      const data = await xlsx.read(book, { sheetName: 'Основной', range: 'A1:B2' })
      assert.deepEqual(data.data, [['A', 'B'], [1, 2]], 'original cells survived')

      // Tab order and sheet names live in workbook.xml; the original worksheet
      // part must still be the one the workbook points at.
      const info = await xlsx.inspect(book)
      assert.deepEqual(info.sheets.map(s => s.name), ['Основной', 'Дополнительный'])
    })

    test('the new worksheet accepts writes and reads back', async () => {
      const xlsx = new XlsxEngine(adapter)
      await xlsx.write(book, {
        sheetName: 'Дополнительный',
        matrix: [['X', 'Y'], [10, 20]]
      })
      const data = await xlsx.read(book, { sheetName: 'Дополнительный', range: 'A1:B2' })
      assert.equal(data.data[0][0], 'X')
      assert.equal(data.data[1][1], 20)
    })

    test('the workbook with the added sheet stays valid', async () => {
      const xlsx = new XlsxEngine(adapter)
      const validation = await xlsx.validate(book)
      assert.equal(validation.valid, true, JSON.stringify(validation.errors))
    })

    test('a cross-sheet formula is written and readable back', async () => {
      const xlsx = new XlsxEngine(adapter)
      await xlsx.write(book, {
        sheetName: 'Дополнительный',
        cells: [{ ref: 'D1', formula: '=SUM(Основной!A2:B2)' }]
      })

      const read = await xlsx.read(book, {
        sheetName: 'Дополнительный',
        range: 'A1:D1',
        includeFormulas: true
      })
      assert.equal(read.formulas[0][3], 'SUM(Основной!A2:B2)',
        'the formula is stored and reported, even though no engine has cached a value yet')
      assert.equal(read.formulaCount, 1)

      // Without the flag the matrix stays a plain grid of values.
      const plain = await xlsx.read(book, { sheetName: 'Дополнительный', range: 'A1:D1' })
      assert.equal(plain.formulas, undefined)
    })

    test('R7 accepts the workbook after a sheet was added', async (t) => {
      if (requiresR7(t)) return
      const pdf = path.join(tmpDir, 'append.pdf')
      await adapter.convert(book, pdf)
      assert.ok(looksLikePdf(pdf))
    })
  })

  describe('PPTX: appending slides', () => {
    const deck = path.join(tmpDir, 'deck.pptx')

    // `create` does have a no-template fallback (_createFallbackBlankPptx), but
    // that package carries no slide master, layout or theme — and `addSlide`
    // builds the new slide FROM a layout, so it can only run against a deck made
    // from R7's template. The three `addSlide` cases below skip themselves
    // without it, as does the final x2t render; `setup` and `validate` stay
    // portable, so the fallback deck still gets created and checked.
    test('setup: a deck with one slide', async () => {
      const pptx = new PptxEngine(adapter)
      await pptx.create(deck, { title: 'Первый слайд' })
      const info = await pptx.inspect(deck)
      assert.equal(info.slidesCount, 1)
      assert.equal(info.slides[0].title, 'Первый слайд')
    })

    test('addSlide appends instead of replacing the deck', async (t) => {
      if (requiresR7(t)) return
      const pptx = new PptxEngine(adapter)
      const before = await ZipArchive.fromFile(deck)
      const slide1Before = before.getBuffer('ppt/slides/slide1.xml')

      const res = await pptx.addSlide(deck, { title: 'Второй слайд' })
      assert.equal(res.success, true)
      assert.equal(res.slideCount, 2)

      const after = await ZipArchive.fromFile(deck)
      const info = await pptx.inspect(deck)
      assert.equal(info.slidesCount, 2, 'the original slide is still there')
      assert.equal(info.slides[0].title, 'Первый слайд', 'slide 1 kept its text')
      assert.equal(info.slides[1].title, 'Второй слайд', 'slide 2 has the new text')

      // Appending must not rewrite the existing slide part.
      assert.ok(sameBytes(slide1Before, after.getBuffer('ppt/slides/slide1.xml')),
        'slide1.xml is byte-identical after appending')
    })

    test('a third slide appends cleanly', async (t) => {
      if (requiresR7(t)) return
      const pptx = new PptxEngine(adapter)
      await pptx.addSlide(deck, { title: 'Третий слайд' })
      const info = await pptx.inspect(deck)
      assert.equal(info.slidesCount, 3)
      assert.deepEqual(info.slides.map(s => s.title), ['Первый слайд', 'Второй слайд', 'Третий слайд'])
    })

    test('the appended slides registered every required package part', async (t) => {
      if (requiresR7(t)) return
      const zip = await ZipArchive.fromFile(deck)
      const contentTypes = zip.getText('[Content_Types].xml')
      const presRels = zip.getText('ppt/_rels/presentation.xml.rels')
      const presentation = zip.getText('ppt/presentation.xml')

      for (const n of [2, 3]) {
        assert.ok(zip.has(`ppt/slides/slide${n}.xml`), `slide${n}.xml exists`)
        assert.ok(zip.has(`ppt/slides/_rels/slide${n}.xml.rels`), `slide${n} rels exists`)
        assert.ok(contentTypes.includes(`/ppt/slides/slide${n}.xml`), `slide${n} declared in [Content_Types]`)
        assert.ok(presRels.includes(`slides/slide${n}.xml`), `slide${n} related from presentation.xml`)
        assert.ok(presentation.includes(`<p:sldId `), 'sldIdLst entries exist')
      }

      // sldId values must be unique or PowerPoint/R7 refuses the file.
      const ids = [...presentation.matchAll(/<p:sldId id="(\d+)"/g)].map(m => m[1])
      assert.equal(ids.length, 3)
      assert.equal(new Set(ids).size, 3, `sldId values are unique: ${ids.join(',')}`)

      // Relationship ids referenced by sldIdLst must exist in the rels part.
      const rids = [...presentation.matchAll(/<p:sldId id="\d+" r:id="([^"]+)"/g)].map(m => m[1])
      for (const rid of rids) {
        assert.ok(presRels.includes(`Id="${rid}"`), `relationship ${rid} is declared`)
      }
    })

    test('the deck with appended slides is still valid', async () => {
      const pptx = new PptxEngine(adapter)
      const validation = await pptx.validate(deck)
      assert.equal(validation.valid, true, JSON.stringify(validation.errors))
    })

    test('R7 opens and renders the deck after slides were appended', async (t) => {
      if (requiresR7(t)) return
      const pdf = path.join(tmpDir, 'deck.pdf')
      await adapter.convert(deck, pdf)
      assert.ok(looksLikePdf(pdf))
      assert.ok(fs.statSync(pdf).size > 500)
    })
  })

  describe('create refuses to destroy an existing document', () => {
    test('DOCX create refuses to overwrite without an explicit flag', async () => {
      const target = path.join(tmpDir, 'guard.docx')
      const docx = new DocxEngine(adapter)
      await docx.create(target, { paragraphs: ['original content'] })

      await assert.rejects(
        () => docx.create(target, { paragraphs: ['replacement'] }),
        /already exists/
      )

      const text = await docx.read(target, { format: 'markdown' })
      assert.ok(text.content.includes('original content'), 'the original survived')
    })

    test('DOCX create overwrites when explicitly asked', async () => {
      const target = path.join(tmpDir, 'guard2.docx')
      const docx = new DocxEngine(adapter)
      await docx.create(target, { paragraphs: ['first'] })
      await docx.create(target, { paragraphs: ['second'], overwrite: true })
      const text = await docx.read(target, { format: 'markdown' })
      assert.ok(text.content.includes('second'))
    })

    test('XLSX create refuses to overwrite without an explicit flag', async () => {
      const target = path.join(tmpDir, 'guard.xlsx')
      const xlsx = new XlsxEngine(adapter)
      await xlsx.create(target, { sheets: [{ name: 'S', data: [['keep me']] }] })
      await assert.rejects(() => xlsx.create(target, { sheets: [{ name: 'S', data: [['nope']] }] }), /already exists/)
      const data = await xlsx.read(target, { range: 'A1' })
      assert.equal(data.data[0][0], 'keep me')
    })

    test('PPTX create refuses to overwrite without an explicit flag', async () => {
      const target = path.join(tmpDir, 'guard.pptx')
      const pptx = new PptxEngine(adapter)
      await pptx.create(target, { title: 'keep me' })
      await assert.rejects(() => pptx.create(target, { title: 'nope' }), /already exists/)
      const info = await pptx.inspect(target)
      assert.equal(info.slides[0].title, 'keep me')
    })
  })
})
