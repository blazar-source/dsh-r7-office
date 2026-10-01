import { test, describe, before, after } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import { PptxEngine } from '../../src/r7/pptx.js'
import { ZipArchive } from '../../src/shared/zip.js'
import { extractElements, getAttribute } from '../../src/shared/xml.js'
import { slideParts, openingTag, firstElement } from '../../src/r7/pptx-util.js'
import { tempDir } from './helpers/pptx-fixtures.js'

/**
 * Structural integrity of a written deck.
 *
 * OOXML requires `p:cNvPr/@id` to be unique inside a slide and forbids two
 * shapes claiming the same placeholder (`p:ph type` + `idx`). A renderer that
 * meets either draws one of the two and clips the other, which is exactly what
 * a user reports as "the title and the text are on top of each other".
 *
 * These tests assert the structure directly, because the read model cannot see
 * the defect: two objects that share an id are both reported, and a caller that
 * looks one up by id silently gets the first.
 */
describe('PPTX structural integrity', () => {
  const tmpDir = tempDir('pptx_structure')
  let engine

  before(() => {
    engine = new PptxEngine()
  })

  after(() => {
    try {
      fs.rmSync(tmpDir, { recursive: true, force: true })
    } catch {
      // best effort
    }
  })

  /** A deck with a title slide and one "title and content" slide. */
  async function twoSlideDeck(name) {
    const target = path.join(tmpDir, name)
    await engine.create(target, { overwrite: true, title: 'Годовой отчёт 2026' })
    await engine.addSlide(target, {
      layoutType: 'obj',
      title: 'Ключевые выводы',
      paragraphs: [{ text: 'Первый пункт', bullet: true }]
    })
    return target
  }

  /** Every top-level object of a slide, as markup. */
  async function shapeElements(target, slideIndex) {
    const zip = await ZipArchive.fromFile(target)
    const descriptor = slideParts(zip)[slideIndex]
    const tree = firstElement(zip.getText(descriptor.partPath), 'p:spTree')
    assert.ok(tree, 'the slide has a shape tree')
    const tags = ['p:sp', 'p:pic', 'p:graphicFrame', 'p:cxnSp', 'p:grpSp']
    const found = []
    for (const tag of tags) {
      for (const el of extractElements(tree, tag)) found.push({ tag, xml: el.outerXml, at: el.index })
    }
    return found.sort((a, b) => a.at - b.at).map((e) => e.xml)
  }

  function shapeId(xml) {
    return Number(getAttribute(openingTag(firstElement(xml, 'p:cNvPr') || ''), 'id') || 0)
  }

  function placeholderKeyOf(xml) {
    const ph = firstElement(xml, 'p:ph')
    if (!ph) return null
    const type = getAttribute(ph, 'type') || 'body'
    const idx = getAttribute(ph, 'idx')
    return `${type}/${idx === null ? '' : idx}`
  }

  // ------------------------------------------------- locating, not duplicating

  test('writing text into the title-slide subtitle keeps one subTitle shape', async () => {
    const target = await twoSlideDeck('subtitle.pptx')
    const before = await engine.readSlide(target, 0)
    const subtitle = before.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'subTitle')
    assert.ok(subtitle, 'precondition: the title slide carries a subtitle placeholder')
    const countBefore = (await shapeElements(target, 0)).length

    await engine.formatObject(target, {
      slideIndex: 0,
      objectId: subtitle.id,
      text: 'Итоги года, ключевые показатели и планы на следующий период'
    })

    const shapes = await shapeElements(target, 0)
    assert.equal(shapes.length, countBefore, 'a text replacement must not add a shape')
    const keys = shapes.map(placeholderKeyOf).filter(Boolean)
    assert.equal(keys.filter((k) => k === 'subTitle/1').length, 1, `one subTitle, got: ${keys.join(', ')}`)
    const updated = await engine.readSlide(target, 0)
    assert.match(
      updated.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'subTitle').text,
      /Итоги года/
    )
  })

  test('replacing the paragraphs of a body placeholder keeps one shape per placeholder', async () => {
    const target = await twoSlideDeck('body.pptx')
    const before = await engine.readSlide(target, 1)
    const body = before.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'body')
    assert.ok(body, 'precondition: the content layout carries a body placeholder')
    const countBefore = (await shapeElements(target, 1)).length

    await engine.formatObject(target, {
      slideIndex: 1,
      objectId: body.id,
      paragraphs: [
        { text: 'Выручка выросла', bullet: true },
        { text: 'Маржа удержана', bullet: true }
      ]
    })

    const shapes = await shapeElements(target, 1)
    assert.equal(shapes.length, countBefore, 'rewriting the text body must not add a shape')
    const keys = shapes.map(placeholderKeyOf).filter(Boolean)
    assert.equal(new Set(keys).size, keys.length, `no placeholder may repeat, got: ${keys.join(', ')}`)
  })

  test('replacing the text of a free text box does not add a shape', async () => {
    const target = await twoSlideDeck('textbox.pptx')
    const box = await engine.addTextBox(target, {
      slideIndex: 1, x: 500000, y: 4000000, width: 4000000, height: 900000, text: 'Надпись'
    })
    const countBefore = (await shapeElements(target, 1)).length

    await engine.formatObject(target, { slideIndex: 1, objectId: box.objectId, text: 'Другая надпись' })

    const shapes = await shapeElements(target, 1)
    assert.equal(shapes.length, countBefore, 'a text box must not be duplicated by its own edit')
    const updated = await engine.readSlide(target, 1)
    const object = updated.slide.objects.find((o) => o.id === box.objectId)
    assert.equal(object.text, 'Другая надпись')
    assert.equal(object.x, 500000, 'the original geometry survives')
    assert.equal(object.placeholder, null, 'a text box must not become a placeholder')
  })

  test('repeated edits never multiply the shapes on a slide', async () => {
    const target = await twoSlideDeck('repeated.pptx')
    const countBefore = (await shapeElements(target, 1)).length
    for (const text of ['Раз', 'Два', 'Три']) {
      const slide = await engine.readSlide(target, 1)
      const body = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'body')
      await engine.formatObject(target, { slideIndex: 1, objectId: body.id, text })
    }
    assert.equal((await shapeElements(target, 1)).length, countBefore, 'three edits, still one body shape')
  })

  // --------------------------------------------------------- the validator

  test('validateStructure is clean on a deck built through the public API', async () => {
    const target = await twoSlideDeck('clean.pptx')
    const slide = await engine.readSlide(target, 1)
    const body = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'body')
    await engine.formatObject(target, { slideIndex: 1, objectId: body.id, text: 'Проверка' })
    await engine.addShape(target, {
      slideIndex: 1, shape: 'rounded-rectangle', name: 'Плашка',
      x: 1000000, y: 5000000, width: 3000000, height: 1000000, text: 'Плашка'
    })

    const report = await engine.validateStructure(target)
    assert.equal(report.valid, true, `expected no structural errors, got: ${report.errors.join('; ')}`)
    assert.deepEqual(report.errors, [])
    assert.equal(report.details.slides.length, 2)
    for (const slideReport of report.details.slides) {
      assert.deepEqual(slideReport.duplicateShapeIds, [])
      assert.deepEqual(slideReport.duplicatePlaceholders, [])
      assert.deepEqual(slideReport.zeroIds, [])
    }
  })

  test('validateStructure reports a duplicated shape and a duplicated placeholder', async () => {
    const target = await twoSlideDeck('duplicated.pptx')
    const corrupt = path.join(tmpDir, 'duplicated-corrupt.pptx')
    const zip = await ZipArchive.fromFile(target)
    const descriptor = slideParts(zip)[1]
    const xml = zip.getText(descriptor.partPath)
    const tree = firstElement(xml, 'p:spTree')
    const firstBody = extractElements(tree, 'p:sp')
      .map((el) => el.outerXml)
      .find((sp) => /type="body"/.test(sp))
    assert.ok(firstBody, 'precondition: the slide has a body placeholder')
    // A copy of an existing shape: same cNvPr id, same placeholder.
    zip.setText(descriptor.partPath, xml.replace('</p:spTree>', `${firstBody}</p:spTree>`))
    await zip.save(corrupt)

    const report = await engine.validateStructure(corrupt)
    assert.equal(report.valid, false)
    assert.ok(
      report.errors.some((e) => /duplicate shape id/i.test(e)),
      `expected a duplicate id error, got: ${report.errors.join('; ')}`
    )
    assert.ok(
      report.errors.some((e) => /duplicate placeholder/i.test(e)),
      `expected a duplicate placeholder error, got: ${report.errors.join('; ')}`
    )
    const slideReport = report.details.slides[1]
    assert.equal(slideReport.duplicateShapeIds.length, 1)
    assert.equal(slideReport.duplicatePlaceholders.length, 1)
  })

  test('validateStructure rejects id 0 and a shape that reuses the group id', async () => {
    const target = await twoSlideDeck('badids.pptx')
    const zeroed = path.join(tmpDir, 'zero-id.pptx')
    const regrouped = path.join(tmpDir, 'group-id.pptx')

    const zip = await ZipArchive.fromFile(target)
    const descriptor = slideParts(zip)[1]
    const xml = zip.getText(descriptor.partPath)
    const groupId = Number(
      getAttribute(openingTag(firstElement(firstElement(xml, 'p:nvGrpSpPr'), 'p:cNvPr')), 'id')
    )
    assert.equal(groupId, 1, 'the shape tree group is id 1')

    zip.setText(descriptor.partPath, xml.replace('<p:cNvPr id="3"', '<p:cNvPr id="0"'))
    await zip.save(zeroed)
    const zeroReport = await engine.validateStructure(zeroed)
    assert.equal(zeroReport.valid, false)
    assert.ok(
      zeroReport.errors.some((e) => /id 0/i.test(e)),
      `expected an id 0 error, got: ${zeroReport.errors.join('; ')}`
    )

    zip.setText(descriptor.partPath, xml.replace('<p:cNvPr id="3"', `<p:cNvPr id="${groupId}"`))
    await zip.save(regrouped)
    const groupReport = await engine.validateStructure(regrouped)
    assert.equal(groupReport.valid, false)
    assert.ok(
      groupReport.errors.some((e) => /group/i.test(e)),
      `expected a group-id collision, got: ${groupReport.errors.join('; ')}`
    )
  })

  test('validate() folds the structural report into its own result', async () => {
    const target = await twoSlideDeck('validate-structure.pptx')
    const corrupt = path.join(tmpDir, 'validate-structure-corrupt.pptx')
    const zip = await ZipArchive.fromFile(target)
    const descriptor = slideParts(zip)[1]
    const xml = zip.getText(descriptor.partPath)
    const sp = extractElements(firstElement(xml, 'p:spTree'), 'p:sp')[0].outerXml
    zip.setText(descriptor.partPath, xml.replace('</p:spTree>', `${sp}</p:spTree>`))
    await zip.save(corrupt)

    const validation = await engine.validate(corrupt)
    assert.equal(validation.valid, false)
    assert.ok(validation.details.structure, 'validate() reports the structural detail')
    assert.equal(validation.details.structure.valid, false)
    assert.ok(validation.errors.some((e) => /duplicate shape id/i.test(e)))
  })

  // ------------------------------------------------------------ title sizing

  test('a title written with text carries normAutofit', async () => {
    const target = await twoSlideDeck('autofit.pptx')
    const slide = await engine.readSlide(target, 1)
    const title = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'title')
    await engine.formatObject(target, {
      slideIndex: 1,
      objectId: title.id,
      text: 'Итоги года, ключевые показатели и планы на следующий период',
      font: { family: 'Arial', size: 30 }
    })

    const shapes = await shapeElements(target, 1)
    const titleShape = shapes.find((sp) => placeholderKeyOf(sp) === 'title/')
    assert.ok(titleShape, 'the title shape is still on the slide')
    assert.match(titleShape, /<a:bodyPr[^>]*>\s*<a:normAutofit\/>/, 'the title body asks the renderer to fit its text')
    assert.doesNotMatch(titleShape, /<a:spAutoFit\/>/, 'a title must not grow its own box')
    // A title never freezes its geometry: the box still comes from the layout.
    assert.doesNotMatch(titleShape, /<a:xfrm>/, 'the title must not write a transform')
  })

  test('a title cannot be inflated by body-level paragraph spacing', async () => {
    const target = await twoSlideDeck('title-spacing.pptx')
    const slide = await engine.readSlide(target, 1)
    const title = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'title')
    await engine.formatObject(target, {
      slideIndex: 1,
      objectId: title.id,
      text: 'Ключевые выводы года',
      font: { family: 'Arial', size: 30 },
      // Body-list options asked of a title: a title's box comes from the layout
      // and does not grow, so spacing inside it is drawn over what follows.
      lineSpacing: 3,
      spaceBefore: 40,
      spaceAfter: 40
    })

    const titleShape = (await shapeElements(target, 1)).find((sp) => placeholderKeyOf(sp) === 'title/')
    assert.ok(titleShape, 'the title shape is still on the slide')
    assert.match(titleShape, /Ключевые выводы года/)
    assert.doesNotMatch(titleShape, /<a:lnSpc/, 'explicit line spacing must not inflate a title')
    assert.doesNotMatch(titleShape, /<a:spcBef/, 'space-before must not inflate a title')
    assert.doesNotMatch(titleShape, /<a:spcAft/, 'space-after must not inflate a title')
    assert.match(titleShape, /<a:normAutofit\/>/, 'the renderer is still told to fit the text')
  })

  test('a long Cyrillic title fits inside the placeholder box it inherits', async () => {
    const target = await twoSlideDeck('long-title.pptx')
    const longTitle = 'Итоги года, ключевые показатели и планы на следующий период развития'
    const slide = await engine.readSlide(target, 1)
    const title = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'title')
    await engine.formatObject(target, {
      slideIndex: 1,
      objectId: title.id,
      text: longTitle,
      font: { family: 'Arial', size: 30 }
    })

    const after = await engine.readSlide(target, 1)
    const updated = after.slide.objects.find((o) => o.id === title.id)
    assert.equal(updated.text, longTitle)
    assert.equal(updated.positionValid, true, 'the title reports a usable box')

    const widthPoints = updated.width / 12700
    const heightPoints = updated.height / 12700
    const size = updated.font.size
    assert.ok(size > 0, 'the title reports its font size')

    // A deliberately pessimistic extent: a Cyrillic glyph is taken as 0.62 em
    // wide and a line as 1.25 times the font size, both worse than Arial
    // actually renders, so a title that fits here fits in a real renderer.
    const perLine = Math.max(1, Math.floor(widthPoints / (size * 0.62)))
    const lines = Math.ceil(longTitle.length / perLine)
    const estimated = lines * size * 1.25
    assert.ok(
      estimated <= heightPoints,
      `the title needs an estimated ${Math.round(estimated)}pt but its box is ${Math.round(heightPoints)}pt`
    )

    const shapes = await shapeElements(target, 1)
    assert.match(
      shapes.find((sp) => placeholderKeyOf(sp) === 'title/'),
      /<a:normAutofit\/>/,
      'and the renderer is still told to fit the text'
    )
  })

  // ------------------------------------------------------- materialising

  test('a placeholder materialised from the layout gets a fresh shape id', async () => {
    const target = path.join(tmpDir, 'materialise.pptx')
    await engine.create(target, { overwrite: true, title: 'Материализация' })
    await engine.addSlide(target, { layoutType: 'titleOnly', title: 'Заголовок' })

    const before = await engine.readSlide(target, 1)
    const inherited = before.slide.objects.find((o) => !o.onSlide && o.placeholder)
    assert.ok(inherited, 'precondition: the layout declares a placeholder the slide does not')
    const countBefore = (await shapeElements(target, 1)).length

    await engine.formatObject(target, {
      slideIndex: 1,
      objectId: inherited.id,
      placeholderType: inherited.placeholder.type,
      font: { family: 'Verdana', size: 14 }
    })

    const shapes = await shapeElements(target, 1)
    assert.equal(shapes.length, countBefore + 1, 'exactly one shape was materialised')
    const ids = shapes.map(shapeId)
    assert.equal(new Set(ids).size, ids.length, `materialised ids must be unique, got: ${ids.join(', ')}`)
    assert.ok(!ids.includes(0), 'no materialised shape may carry id 0')
    const report = await engine.validateStructure(target)
    assert.equal(report.valid, true, `expected a clean deck, got: ${report.errors.join('; ')}`)
  })
})
