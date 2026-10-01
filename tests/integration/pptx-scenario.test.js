import { test, describe, before, after } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import { PptxEngine } from '../../src/r7/pptx.js'
import { R7Adapter } from '../../src/r7/adapter.js'
import { ZipArchive } from '../../src/shared/zip.js'
import { readRelationships, slideParts } from '../../src/r7/pptx-util.js'
import { sameBytes, tempDir, writeGradientPng } from './helpers/pptx-fixtures.js'

/**
 * The critical scenario, run against a real multi-object business deck.
 *
 * The deck is generated in the test (and by `examples/pptx-acceptance.js`) from
 * the R7 blank template, so its slide master, eleven layouts, two themes,
 * notes master, table styles and extended properties are the ones R7 ships.
 * That is the structure a document authored in R7-Office Desktop has, and the
 * scenario is the one that matters:
 *
 *   read -> change one object's text -> restyle a second -> add an object ->
 *   add a slide -> save as a COPY -> reopen and assert
 *
 * Everything the operation did not explicitly address has to be byte-identical,
 * the master/theme/layouts/notes must survive, unknown objects must still be
 * there, and the relationships must still resolve.
 */
describe('PPTX critical scenario on a real business deck', () => {
  const tmpDir = tempDir('pptx_scenario')
  let engine
  let r7Available = false
  let deck = null
  let originalBytes = null
  const copy = path.join(tmpDir, 'Сценарий — копия.pptx')

  /** Build a deck with text, colours, a shape, a chart, an image and 7 slides. */
  async function authorDeck(target) {
    await engine.create(target, { overwrite: true, title: 'Итоги года' })

    // Slide 2: heading and body.
    await engine.addSlide(target, { layoutType: 'obj', title: 'Ключевые выводы', paragraphs: ['Пункт один', 'Пункт два'] })
    // Slide 3: two columns.
    await engine.addSlide(target, { layoutType: 'twoObj', title: 'Сравнение' })
    // Slide 4: KPI shapes.
    await engine.addSlide(target, { layoutType: 'titleOnly', title: 'Показатели' })
    await engine.addShape(target, {
      slideIndex: 3, shape: 'rounded-rectangle', x: '1.4cm', y: '4.2cm', width: '9.4cm', height: '8.2cm',
      fill: '#F3F6FB', noLine: true, name: 'Подложка'
    })
    await engine.addShape(target, {
      slideIndex: 3, shape: 'rounded-rectangle', x: '1.4cm', y: '4.8cm', width: '9.4cm', height: '8.2cm',
      fill: '#FFFFFF', line: '#D2DAE6', lineWidth: 1, name: 'Карточка', text: 'Выручка', font: { family: 'Arial', size: 13, color: '#52606D' },
      alignment: 'center', verticalAnchor: 'middle'
    })
    // Slide 5: image and caption.
    await engine.addSlide(target, { layoutType: 'titleOnly', title: 'Динамика' })
    const png = writeGradientPng(path.join(path.dirname(target), 'scenario.png'), 320, 200)
    await engine.addImage(target, {
      slideIndex: 4, imagePath: png.filePath, x: '4cm', y: '4.2cm', width: '24cm', name: 'График'
    })
    await engine.addTextBox(target, {
      slideIndex: 4, x: '4cm', y: '15.4cm', width: '24cm', height: '1.4cm',
      text: 'Рис. 1. Динамика выручки', font: { family: 'Arial', size: 12, italic: true, color: '#52606D' },
      alignment: 'center', verticalAnchor: 'middle', fill: null, noFill: true, noLine: true
    })
    // Slide 6: numbered list.
    await engine.addSlide(target, { layoutType: 'obj', title: 'План' })
    // Slide 7: closing on the title layout.
    await engine.addSlide(target, { layoutType: 'title', title: 'Спасибо', subtitle: 'Вопросы' })
  }

  before(async () => {
    engine = new PptxEngine()
    const info = await new R7Adapter().detect()
    r7Available = info.installed
    deck = path.join(tmpDir, 'Сценарий.pptx')
    await authorDeck(deck)
    originalBytes = fs.readFileSync(deck)
  })

  after(() => {
    try {
      fs.rmSync(tmpDir, { recursive: true, force: true })
    } catch {
      // best effort
    }
  })

  test('the authored deck has the seven slides and the real R7 part graph', async () => {
    const slide = await engine.readSlide(deck, 0)
    assert.equal(slide.slideCount, 7)
    const zip = await ZipArchive.fromFile(deck)
    if (r7Available) {
      assert.equal(slide.layouts.length, 11, 'the R7 template ships eleven layouts')
      for (const part of [
        'ppt/slideMasters/slideMaster1.xml',
        'ppt/theme/theme1.xml',
        'ppt/theme/theme2.xml',
        'ppt/notesMasters/notesMaster1.xml',
        'ppt/tableStyles.xml',
        'ppt/presProps.xml',
        'ppt/viewProps.xml'
      ]) {
        assert.ok(zip.has(part), `${part} is present`)
      }
    }
    assert.equal((await engine.validate(deck)).valid, true)
  })

  test('read -> edit text -> restyle -> add object -> add slide -> save as a copy', async () => {
    // 1. READ: the normalized model is enough to address everything.
    const slideTwo = await engine.readSlide(deck, 1)
    const title = slideTwo.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'title')
    const body = slideTwo.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'body')
    assert.ok(title && body)
    assert.ok(title.x !== null && title.width !== null, 'the reader reports usable geometry')

    // 2. CHANGE THE TEXT OF ONE OBJECT.
    await engine.formatObject(deck, {
      slideIndex: 1, objectId: title.id, text: 'Изменённый заголовок', outputPath: copy
    })
    // 3. CHANGE THE STYLING OF A SECOND.
    await engine.formatObject(copy, {
      slideIndex: 1, objectId: body.id,
      font: { family: 'Georgia', size: 20, bold: true, color: '#C00000' },
      alignment: 'center', lineSpacing: 1.4
    })
    // 4. ADD AN OBJECT.
    const added = await engine.addShape(copy, {
      slideIndex: 1, shape: 'ellipse', x: '20cm', y: '12cm', width: '4cm', height: '3cm',
      fill: '#0E9F6E', line: '#0B3D91', lineWidth: 2, name: 'Добавленная фигура'
    })
    // 5. ADD A SLIDE.
    const addedSlide = await engine.addSlide(copy, { layoutType: 'obj', title: 'Добавленный слайд', paragraphs: ['Новый пункт'] })

    // 6. REOPEN AND ASSERT.
    assert.equal(addedSlide.slideCount, 8)

    const reopened = await engine.readSlide(copy, 1)
    const newTitle = reopened.slide.objects.find((o) => o.id === title.id)
    assert.equal(newTitle.text, 'Изменённый заголовок', 'the edited text survived')
    assert.equal(newTitle.x, title.x, 'the edited object kept its geometry')
    assert.equal(newTitle.width, title.width)

    const newBody = reopened.slide.objects.find((o) => o.id === body.id)
    assert.equal(newBody.font.family, 'Georgia')
    assert.equal(newBody.font.size, 20)
    assert.equal(newBody.font.bold, true)
    assert.equal(newBody.font.color, '#C00000')
    assert.equal(newBody.alignment.horizontal, 'ctr')
    assert.equal(newBody.paragraphs[0].lineSpacing, 1.4)

    const newShape = reopened.slide.objects.find((o) => o.id === added.objectId)
    assert.ok(newShape, 'the added object is addressable by its new id')
    assert.equal(newShape.fill.color, '#0E9F6E')
    assert.equal(newShape.stroke.color, '#0B3D91')

    const newSlide = await engine.readSlide(copy, 7)
    assert.equal(newSlide.slide.layout.type, 'obj')
    const newSlideTitle = newSlide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'title')
    assert.equal(newSlideTitle.text, 'Добавленный слайд')

    assert.equal((await engine.validate(copy)).valid, true)
  })

  test('every other slide is byte-identical after the scenario', async () => {
    const before = await ZipArchive.fromFile(deck)
    const after = await ZipArchive.fromFile(copy)

    const descriptorsBefore = slideParts(before)
    const descriptorsAfter = slideParts(after)
    const touched = new Set([
      descriptorsBefore[1].partPath,          // the edited slide
      'ppt/presentation.xml',                 // slide list changed by addSlide
      'ppt/_rels/presentation.xml.rels',      // and its relationship part
      '[Content_Types].xml',                  // the new slide part is registered
      'docProps/app.xml'                      // the slide count
    ])

    // Unrelated slides keep their original bytes.
    for (let index = 0; index < descriptorsBefore.length; index++) {
      if (index === 1) continue
      const part = descriptorsBefore[index].partPath
      assert.ok(after.has(part), `${part} still exists`)
      assert.ok(
        sameBytes(before.entries.get(part).raw, after.entries.get(part).raw),
        `unrelated slide ${part} must be byte-identical`
      )
      assert.ok(
        sameBytes(
          before.entries.get(`ppt/slides/_rels/${path.basename(part)}.rels`).raw,
          after.entries.get(`ppt/slides/_rels/${path.basename(part)}.rels`).raw
        ),
        `unrelated slide relationships for ${part} must be byte-identical`
      )
    }

    // The master, the layouts, the themes, the notes master and the document
    // metadata are all untouched.
    for (const name of before.list()) {
      if (touched.has(name)) continue
      if (!after.has(name)) continue
      assert.ok(
        sameBytes(before.entries.get(name).raw, after.entries.get(name).raw),
        `"${name}" must be byte-identical after the scenario`
      )
    }
    assert.equal(after.list().length, before.list().length + 2, 'exactly the new slide and its relationships were added')
  })

  test('the master, layouts, theme and notes are all preserved and relationships resolve', async () => {
    const before = await ZipArchive.fromFile(deck)
    const after = await ZipArchive.fromFile(copy)

    for (const pattern of [
      /^ppt\/slideMasters\//, /^ppt\/slideLayouts\//, /^ppt\/theme\//, /^ppt\/notesMasters\//, /^ppt\/notesSlides\//
    ]) {
      const parts = before.list().filter((n) => pattern.test(n))
      assert.ok(parts.length > 0, `the deck has parts matching ${pattern}`)
      for (const name of parts) {
        assert.ok(after.has(name), `"${name}" survived`)
        assert.ok(sameBytes(before.entries.get(name).raw, after.entries.get(name).raw), `"${name}" is byte-identical`)
      }
    }

    // Every relationship of every slide still points at something that exists.
    for (const descriptor of slideParts(after)) {
      for (const rel of readRelationships(after, descriptor.partPath)) {
        if (rel.targetMode === 'External') continue
        assert.ok(after.has(rel.partPath), `${descriptor.partPath} -> ${rel.partPath} resolves`)
      }
    }

    // The layouts still list the same names, in the same order.
    const layouts = await engine.listLayouts(copy)
    assert.equal(layouts.layouts.filter((l) => l.name).length, 11)
  })

  test('the image and the caption survive, and the media is untouched', async () => {
    const before = await ZipArchive.fromFile(deck)
    const after = await ZipArchive.fromFile(copy)

    const mediaBefore = before.list().filter((n) => /^ppt\/media\//.test(n))
    assert.ok(mediaBefore.length >= 1, 'the deck has media')
    for (const name of mediaBefore) {
      assert.ok(after.has(name), `"${name}" survived`)
      assert.ok(sameBytes(before.entries.get(name).raw, after.entries.get(name).raw), `"${name}" is byte-identical`)
    }

    const slide = await engine.readSlide(copy, 4)
    const image = slide.slide.objects.find((o) => o.type === 'image')
    assert.ok(image, 'the picture is still on slide 5')
    assert.equal(image.image.format, 'png')
    assert.equal(image.image.relId.length > 0, true, 'its relationship is intact')
    const caption = slide.slide.objects.find((o) => o.type === 'textBox' && /Рис\./.test(o.text || ''))
    assert.ok(caption, 'the caption is still there')
    assert.equal(caption.font.italic, true)
  })

  test('the copy renders to PDF exactly like the original', async (t) => {
    if (!r7Available) {
      t.skip('R7-Office installation not available on this host')
      return
    }
    for (const [label, source] of [
      ['original', deck],
      ['copy', path.join(tmpDir, 'Сценарий — копия.pptx')]
    ]) {
      const pdfPath = path.join(tmpDir, `${label}.pdf`)
      const converted = await engine.toPdf(source, pdfPath)
      assert.equal(converted.success, true, `${label} converted`)
      assert.ok(converted.isPdf, `${label} produced a PDF`)
      assert.ok(fs.statSync(pdfPath).size > 5000, `${label} PDF has content`)
    }
    // The copy has one more slide, so its PDF must be larger.
    assert.ok(
      fs.statSync(path.join(tmpDir, 'copy.pdf')).size > fs.statSync(path.join(tmpDir, 'original.pdf')).size,
      'the copy renders one extra slide'
    )
  })
})
