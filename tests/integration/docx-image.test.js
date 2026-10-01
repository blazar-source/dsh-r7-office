import { test, describe, before, after } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import { DocxEngine } from '../../src/r7/docx.js'
import { R7Adapter } from '../../src/r7/adapter.js'
import { emuToCm, probeImage } from '../../src/r7/docx-media.js'
import { archiveOf, cleanup, diffMembers, makeJpegHeader, makePng, sameBytes, tempDir } from './docx-fixtures.test.js'

const r7Available = (await new R7Adapter().detect()).installed

/**
 * Scenario 4: insert a picture with an explicit display size, keep its aspect
 * ratio, and never damage an image or a relationship that is already there.
 */
describe('DOCX images', () => {
  const dir = tempDir('docx_image')
  const file = path.join(dir, 'picture.docx')
  const engine = new DocxEngine()
  const png4x2 = makePng(4, 2)

  before(async () => {
    await engine.create(file, {
      title: 'Документ с иллюстрацией',
      paragraphs: ['Первый абзац.', 'Второй абзац.'],
      // The document therefore *ends* with a table cell paragraph, which is
      // exactly the anchor a careless "append at the end" would choose.
      tables: [{ rows: [['A', 'B'], ['C', 'D']] }]
    })
  })

  after(() => cleanup(dir))

  test('inserting a picture writes the media part, its relationship and its content type', async () => {
    const result = await engine.insertImage(file, { buffer: png4x2, widthCm: 4, alt: 'Точка' })
    assert.equal(result.success, true)
    assert.equal(result.mediaPartName, 'word/media/image1.png')
    assert.equal(result.format, 'png')
    assert.equal(result.scaled, 'width')

    const zip = await archiveOf(file)
    assert.ok(zip.has('word/media/image1.png'), 'the media part must exist')
    assert.ok(sameBytes(zip.getBuffer('word/media/image1.png'), png4x2), 'the bytes must be stored verbatim')
    assert.match(zip.getText('word/_rels/document.xml.rels'), /media\/image1\.png/)
    assert.match(zip.getText('[Content_Types].xml'), /Extension="png" ContentType="image\/png"/)

    const docXml = zip.getText('word/document.xml')
    assert.match(docXml, /<w:drawing><wp:inline/)
    assert.match(docXml, /<a:blip[^>]*r:embed="rId\d+"/)
    assert.match(docXml, /<wp:extent cx="1440000" cy="720000"\/>/, '4 cm x 2 cm in EMU')

    assert.equal((await engine.validate(file)).valid, true)
  })

  test('the picture is discoverable with its geometry', async () => {
    const images = await engine.images(file)
    assert.equal(images.length, 1)
    const image = images[0]
    assert.equal(image.mediaPartName, 'word/media/image1.png')
    assert.equal(image.format ?? 'png', 'png')
    assert.equal(image.widthCm, 4)
    assert.equal(image.heightCm, 2)
    assert.equal(image.aspectRatio, 2)
    assert.equal(image.description, 'Точка')
  })

  test('naming one dimension preserves the image aspect ratio', async () => {
    const target = path.join(dir, 'aspect.docx')
    await engine.create(target, { paragraphs: ['Картинка ниже.'] })

    const wide = await engine.insertImage(target, { buffer: makePng(4, 2), widthCm: 6 })
    assert.equal(wide.widthCm, 6)
    assert.equal(wide.heightCm, 3)
    assert.equal(wide.aspectRatio, 2)

    const scaled = await engine.insertImage(target, {
      buffer: makePng(4, 2),
      paragraphIndex: 'new',
      heightCm: 2
    })
    assert.equal(scaled.heightCm, 2)
    assert.equal(scaled.widthCm, 4)
    assert.equal(scaled.scaled, 'height')

    // Pixels at 96 dpi are just another way of naming the same dimension.
    const pixels = await engine.insertImage(target, { buffer: makePng(4, 2), paragraphIndex: 'new', widthPx: 192 })
    assert.equal(pixels.pixelWidth, 4)
    assert.equal(pixels.widthCm, emuToCm(192 * 9525))
    assert.equal(pixels.heightCm, emuToCm(96 * 9525))
  })

  test('naming both dimensions uses exactly those values', async () => {
    const target = path.join(dir, 'both-dims.docx')
    await engine.create(target, { paragraphs: ['Растянутая картинка.'] })

    const result = await engine.insertImage(target, {
      buffer: makePng(4, 2),
      widthCm: 3,
      heightCm: 5
    })
    assert.equal(result.widthCm, 3)
    assert.equal(result.heightCm, 5)
    assert.equal(result.scaled, 'both')
  })

  test('a picture is inserted into the paragraph that was named', async () => {
    const target = path.join(dir, 'positioned.docx')
    await engine.create(target, { paragraphs: ['Первый.', 'Второй.', 'Третий.'] })

    await engine.insertImage(target, { buffer: makePng(2, 2), paragraphIndex: 1, widthCm: 2 })
    const images = await engine.images(target)
    assert.equal(images[0].paragraphIndex, 1)

    const zip = await archiveOf(target)
    const paragraphs = docXmlParagraphs(zip.getText('word/document.xml'))
    assert.match(paragraphs[1], /<w:drawing>/)
    assert.doesNotMatch(paragraphs[0], /<w:drawing>/)
  })

  test('an image appended to a document that ends with a table gets its own body paragraph', async () => {
    const target = path.join(dir, 'after-table.docx')
    await engine.create(target, {
      paragraphs: ['Абзац.'],
      tables: [{ rows: [['A', 'B'], ['C', 'D']] }]
    })

    const result = await engine.insertImage(target, { buffer: makePng(2, 2), paragraphIndex: 'end', widthCm: 2 })
    const zip = await archiveOf(target)
    const docXml = zip.getText('word/document.xml')

    // The drawing must be outside the table.
    const tableEnd = docXml.indexOf('</w:tbl>')
    const drawingAt = docXml.indexOf('<w:drawing>')
    assert.ok(drawingAt > tableEnd, 'the picture must follow the table, not live inside its last cell')
    assert.equal(imagesInTable(docXml), 0)
    assert.equal(result.paragraphIndex > 0, true)
  })

  test('a second image gets its own part and leaves the first untouched', async () => {
    const target = path.join(dir, 'two-images.docx')
    await engine.create(target, { paragraphs: ['Две картинки.'] })

    const first = makePng(4, 2, [10, 20, 30])
    const second = makePng(8, 8, [200, 100, 50])
    await engine.insertImage(target, { buffer: first, widthCm: 3 })
    const before = await archiveOf(target)

    const inserted = await engine.insertImage(target, { buffer: second, paragraphIndex: 'new', widthCm: 3 })
    assert.equal(inserted.mediaPartName, 'word/media/image2.png')

    const after = await archiveOf(target)
    assert.ok(sameBytes(after.getBuffer('word/media/image1.png'), first), 'the first image must be untouched')
    assert.ok(sameBytes(after.getBuffer('word/media/image2.png'), second))
    assert.equal(images2Count(after.getText('word/document.xml')), 2)
    assert.equal(before.getText('word/_rels/document.xml.rels').includes('image2.png'), false)

    const images = await engine.images(target)
    assert.deepEqual(images.map(image => image.mediaPartName), ['word/media/image1.png', 'word/media/image2.png'])
  })

  test('a JPEG is stored with the JPEG content type', async () => {
    const target = path.join(dir, 'jpeg.docx')
    await engine.create(target, { paragraphs: ['JPEG.'] })

    const jpeg = makeJpegHeader(320, 200)
    const result = await engine.insertImage(target, { buffer: jpeg, widthCm: 5 })
    assert.equal(result.mediaPartName, 'word/media/image1.jpeg')
    assert.equal(result.pixelWidth, 320)
    assert.equal(result.pixelHeight, 200)
    assert.equal(probeImage(jpeg).format, 'jpeg')

    const zip = await archiveOf(target)
    assert.match(zip.getText('[Content_Types].xml'), /Extension="jpeg" ContentType="image\/jpeg"/)
    assert.ok(sameBytes(zip.getBuffer('word/media/image1.jpeg'), jpeg))
  })

  test('an image survives an ordinary text edit byte-identically', async () => {
    const target = path.join(dir, 'edit-around-image.docx')
    await engine.create(target, { paragraphs: ['Исходный текст для замены.'] })
    await engine.insertImage(target, { buffer: makePng(4, 2), widthCm: 3, paragraphIndex: 'new' })

    const before = await archiveOf(target)
    await engine.replaceText(target, 'Исходный текст для замены.', 'Заменённый текст.')
    const after = await archiveOf(target)

    const diff = diffMembers(before, after, ['word/document.xml'])
    assert.deepEqual(diff.changed, [])
    assert.deepEqual(diff.removed, [])
    assert.ok(sameBytes(after.getBuffer('word/media/image1.png'), before.getBuffer('word/media/image1.png')))
    assert.match(after.getText('word/document.xml'), /<w:drawing>/)
  })

  test('a missing image source and an impossible paragraph are refused', async () => {
    await assert.rejects(() => engine.insertImage(file, { widthCm: 2 }), /imagePath, data \(base64\) or buffer/)
    await assert.rejects(
      () => engine.insertImage(file, { buffer: png4x2, paragraphIndex: 99 }),
      /Paragraph index out of range/
    )
    await assert.rejects(() => engine.insertImage(file, { imagePath: path.join(dir, 'нет.png') }), /does not exist/)
  })

  test('R7 reopens the document with the picture and renders it', async (t) => {
    if (!r7Available) {
      t.skip('R7 not installed')
      return
    }
    const pdf = path.join(dir, 'picture.pdf')
    await new R7Adapter().convert(file, pdf)
    assert.equal(fs.readFileSync(pdf).subarray(0, 5).toString(), '%PDF-')
    assert.ok(fs.statSync(pdf).size > 1000)
  })
})

/** Inner XML of every `<w:p>` in document order. */
function docXmlParagraphs(docXml) {
  const result = []
  const regex = /<w:p(?=[\s>])[^>]*>([\s\S]*?)<\/w:p>/g
  let match
  while ((match = regex.exec(docXml)) !== null) result.push(match[1])
  return result
}

/** How many drawings sit inside a `<w:tbl>`. */
function imagesInTable(docXml) {
  const tables = docXml.match(/<w:tbl(?=[\s>])[\s\S]*?<\/w:tbl>/g) || []
  return tables.reduce((count, table) => count + (table.match(/<w:drawing>/g) || []).length, 0)
}

function images2Count(docXml) {
  return (docXml.match(/<w:drawing>/g) || []).length
}
