import { test, describe, before, after } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import { PptxEngine } from '../../src/r7/pptx.js'
import { ZipArchive } from '../../src/shared/zip.js'
import { requiresR7 } from '../helpers/r7-gate.js'
import { readRelationships, slideLayouts, slideParts, resolvePartPath } from '../../src/r7/pptx-util.js'
import { tempDir, injectChartAndSmartArt, sameBytes, addSlideAnywhere } from './helpers/pptx-fixtures.js'

/**
 * Slide lifecycle, layout reuse and — the point of the whole exercise —
 * preservation of everything the engine was not asked to touch.
 */
describe('PPTX slides, layouts and preservation', () => {
  const tmpDir = tempDir('pptx_slides')
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

  test('creating a deck refuses to overwrite without permission', async () => {
    const target = path.join(tmpDir, 'refuse.pptx')
    await engine.create(target, { title: 'Первый' })
    await assert.rejects(() => engine.create(target, { title: 'Второй' }), /Refusing to overwrite/)
    await engine.create(target, { title: 'Третий', overwrite: true })
    const slide = await engine.readSlide(target, 0)
    // With R7 the title is a placeholder declared by the title layout; the
    // synthetic package has no layout, so the title is the plain shape it
    // starts with. Both carry the text the last create() was given.
    const title = slide.slide.objects.find((o) => o.placeholder && (o.placeholder.type === 'ctrTitle' || o.placeholder.type === 'title'))
      || slide.slide.objects.find((o) => o.name === 'Title')
    assert.equal(title.text, 'Третий')
  })

  test('adding a slide builds it on a real layout of the deck and registers every part', async (t) => {
    // Its subject is the layout graph and the slide's layout relationship.
    if (requiresR7(t)) return
    const target = path.join(tmpDir, 'add.pptx')
    await engine.create(target, { title: 'Заголовок' })

    const layoutsBefore = await engine.listLayouts(target)
    const added = await engine.addSlide(target, {
      layoutType: 'obj',
      title: 'Заголовок раздела',
      paragraphs: ['Первый пункт', 'Второй пункт']
    })
    assert.equal(added.slideIndex, 1)
    assert.equal(added.slideCount, 2)
    assert.equal(added.layout.type, 'obj')

    const layoutsAfter = await engine.listLayouts(target)
    assert.equal(layoutsAfter.layouts.length, layoutsBefore.layouts.length, 'no layout was created')

    const zip = await ZipArchive.fromFile(target)
    const descriptors = slideParts(zip)
    assert.equal(descriptors.length, 2)

    // The slide part exists, is registered in [Content_Types].xml, and links a layout.
    const part = descriptors[1].partPath
    assert.ok(zip.has(part), 'the slide part exists')
    assert.ok(zip.has(`ppt/slides/_rels/${path.basename(part)}.rels`), 'its relationship part exists')
    assert.match(zip.getText('[Content_Types].xml'), new RegExp(`PartName="/${part.replace(/\//g, '\\/')}"`))

    const rels = readRelationships(zip, part)
    const layoutRel = rels.find((r) => /slideLayout$/.test(r.type))
    assert.ok(layoutRel, 'the slide links a layout')
    assert.equal(layoutRel.partPath, added.layout.partPath ?? layoutRel.partPath)
    assert.ok(zip.has(layoutRel.partPath), 'the linked layout exists in the package')

    const presentationRels = readRelationships(zip, 'ppt/presentation.xml')
    assert.ok(
      presentationRels.some((r) => r.partPath === part),
      'the presentation declares a relationship to the new slide'
    )
  })

  test('the layout is chosen by index, name and type, and a bad name is rejected', async (t) => {
    // Layout lookup by index, name and type only exists with R7's layouts.
    if (requiresR7(t)) return
    const target = path.join(tmpDir, 'layout-choice.pptx')
    await engine.create(target, { title: 'Заголовок' })
    const layouts = await engine.listLayouts(target)

    const byIndex = await engine.addSlide(target, { layoutIndex: 2, title: 'По индексу' })
    assert.equal(byIndex.layout.index, layouts.layouts[2].index)
    assert.equal(byIndex.layout.name, layouts.layouts[2].name)

    const byName = await engine.addSlide(target, { layoutName: 'Пустой слайд', title: 'По имени' })
    assert.equal(byName.layout.name, layouts.layouts.find((l) => l.name === 'Пустой слайд').name)

    const byType = await engine.addSlide(target, { layoutType: 'twoObj', title: 'По типу' })
    assert.equal(byType.layout.type, 'twoObj')

    await assert.rejects(() => engine.addSlide(target, { layoutName: 'Несуществующий' }), /No layout named/)
    await assert.rejects(() => engine.addSlide(target, { layoutType: 'nope' }), /No layout of type/)
    await assert.rejects(() => engine.addSlide(target, { layoutIndex: 99 }), /Layout index 99 not found/)
  })

  test('a slide added on a title layout gets title and subtitle placeholders', async (t) => {
    // The ctrTitle and subTitle placeholders come from R7's title layout.
    if (requiresR7(t)) return
    const target = path.join(tmpDir, 'title-slide.pptx')
    await engine.create(target, { title: 'Заголовок' })
    const added = await engine.addSlide(target, {
      layoutType: 'title',
      title: 'Титульный заголовок',
      subtitle: 'Подзаголовок презентации'
    })
    assert.equal(added.layout.type, 'title')

    const slide = await engine.readSlide(target, 1)
    const title = slide.slide.objects.find((o) => o.placeholder && (o.placeholder.type === 'ctrTitle' || o.placeholder.type === 'title'))
    const subtitle = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'subTitle')
    assert.equal(title.text, 'Титульный заголовок')
    assert.ok(subtitle, 'the subtitle placeholder was created')
    assert.equal(subtitle.text, 'Подзаголовок презентации')
  })

  test('a deck built only on blank layouts has no content placeholders', async (t) => {
    // "Blank layout" is an R7 template concept.
    if (requiresR7(t)) return
    const target = path.join(tmpDir, 'blank.pptx')
    await engine.create(target, { title: 'Заголовок' })
    await engine.addSlide(target, { layoutType: 'blank' })
    const slide = await engine.readSlide(target, 1)
    const content = slide.slide.objects.filter(
      (o) => o.onSlide && o.placeholder && o.placeholder.type !== 'dt'
        && o.placeholder.type !== 'ftr' && o.placeholder.type !== 'sldNum'
    )
    assert.equal(content.length, 0, 'a blank slide starts empty')
  })

  test('deleting a slide removes its parts and keeps the others valid', async (t) => {
    // The four slides are built on layouts and addressed by their placeholders.
    if (requiresR7(t)) return
    const target = path.join(tmpDir, 'delete.pptx')
    await engine.create(target, { title: 'Слайд 1' })
    await engine.addSlide(target, { layoutType: 'obj', title: 'Слайд 2' })
    await engine.addSlide(target, { layoutType: 'obj', title: 'Слайд 3' })
    await engine.addSlide(target, { layoutType: 'obj', title: 'Слайд 4' })

    const deletedPart = (await engine.readSlide(target, 1)).slide.partPath
    const result = await engine.deleteSlide(target, 1)
    assert.equal(result.slideCount, 3)
    assert.ok(result.removedParts.includes(deletedPart))

    const zip = await ZipArchive.fromFile(target)
    assert.equal(zip.has(deletedPart), false, 'the slide part is gone')
    assert.equal(zip.has(`ppt/slides/_rels/${path.basename(deletedPart)}.rels`), false, 'its relationships are gone')
    assert.doesNotMatch(
      zip.getText('[Content_Types].xml'),
      new RegExp(deletedPart.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&')),
      'the content-type override is gone'
    )

    const validation = await engine.validate(target)
    assert.equal(validation.valid, true, JSON.stringify(validation.errors))

    const remaining = await engine.readSlide(target, 0)
    assert.deepEqual(
      remaining.slides.map((s) => s.slideNumber),
      [2, 3]
    )
    // Reading the deck after the deletion must resolve the remaining slides by
    // position, not by the part number the file happens to have.
    const second = await engine.readSlide(target, 1)
    const title = second.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'title')
    assert.equal(title.text, 'Слайд 3')
  })

  test('deleting a slide that has notes removes the notes part too', async (t) => {
    if (requiresR7(t)) return
    const target = path.join(tmpDir, 'delete-notes.pptx')
    await engine.create(target, { title: 'С заметками' })
    await engine.addSlide(target, { layoutType: 'obj', title: 'Второй' })
    // The R7 blank template ships notesSlide1 for its first slide.
    const zip = await ZipArchive.fromFile(target)
    const descriptors = slideParts(zip)
    const rels = readRelationships(zip, descriptors[0].partPath)
    const notesRel = rels.find((r) => /notesSlide$/.test(r.type))
    if (!notesRel) {
      t.skip('this R7 template has no notes slide')
      return
    }
    assert.ok(zip.has(notesRel.partPath), 'precondition: the notes part exists')

    await engine.deleteSlide(target, 0)

    const after = await ZipArchive.fromFile(target)
    assert.equal(after.has(notesRel.partPath), false, 'the notes part went with its slide')
    assert.equal((await engine.readSlide(target, 0)).slideCount, 1)
  })

  test('duplicating a slide gives the copy fresh shape ids and no notes', async (t) => {
    // The duplicate is expected to keep the source slide's layout link, which
    // only a layout-bearing deck has.
    if (requiresR7(t)) return
    const target = path.join(tmpDir, 'duplicate.pptx')
    await engine.create(target, { title: 'Оригинал' })
    await engine.addSlide(target, { layoutType: 'obj', title: 'Копируемый', paragraphs: ['Текст'] })
    await engine.addShape(target, { slideIndex: 1, shape: 'rectangle', x: 100000, y: 100000, width: 500000, height: 500000, name: 'Фигура' })

    const result = await engine.duplicateSlide(target, 1)
    assert.equal(result.sourceIndex, 1)
    assert.equal(result.slideIndex, 2)
    assert.equal(result.slideCount, 3)

    const original = await engine.readSlide(target, 1)
    const copy = await engine.readSlide(target, 2)
    assert.equal(copy.slide.objects.filter((o) => o.onSlide).length, original.slide.objects.filter((o) => o.onSlide).length)

    // Shape ids must be unique within the deck: two shapes with the same id
    // make R7 refuse to edit either of them.
    const originalIds = original.slide.objects.filter((o) => o.onSlide).map((o) => o.id)
    const copyIds = copy.slide.objects.filter((o) => o.onSlide).map((o) => o.id)
    for (const id of copyIds) {
      assert.equal(originalIds.includes(id), false, `id ${id} must not be reused in a duplicate`)
    }

    const zip = await ZipArchive.fromFile(target)
    const copyRels = readRelationships(zip, copy.slide.partPath)
    assert.equal(copyRels.some((r) => /notesSlide$/.test(r.type)), false, 'a notes part cannot belong to two slides')
    assert.ok(copyRels.some((r) => /slideLayout$/.test(r.type)), 'the copy keeps the layout link')
    assert.equal((await engine.validate(target)).valid, true)
  })

  test('moving and reordering slides changes only the presentation order', async (t) => {
    // The slides are built on layouts and told apart by their placeholder title.
    if (requiresR7(t)) return
    const target = path.join(tmpDir, 'reorder.pptx')
    await engine.create(target, { title: 'A' })
    await engine.addSlide(target, { layoutType: 'obj', title: 'B' })
    await engine.addSlide(target, { layoutType: 'obj', title: 'C' })

    const before = await ZipArchive.fromFile(target)
    const orderBefore = slideParts(before).map((d) => d.partPath)

    await engine.moveSlide(target, 2, 0)

    const after = await ZipArchive.fromFile(target)
    const orderAfter = slideParts(after).map((d) => d.partPath)
    assert.deepEqual(orderAfter, [orderBefore[2], orderBefore[0], orderBefore[1]])

    // Nothing but the presentation's own slide list may change: no slide part,
    // no layout, no relationship.
    for (const name of before.list()) {
      if (name === 'ppt/presentation.xml') continue
      assert.ok(
        sameBytes(before.entries.get(name).raw, after.entries.get(name).raw),
        `"${name}" must be byte-identical after a reorder`
      )
    }

    const titles = []
    for (let i = 0; i < 3; i++) {
      const slide = await engine.readSlide(target, i)
      const title = slide.slide.objects.find((o) => o.placeholder && (o.placeholder.type === 'title' || o.placeholder.type === 'ctrTitle'))
      titles.push(title.text)
    }
    assert.deepEqual(titles, ['C', 'A', 'B'])
  })

  test('reorderSlides accepts a full permutation and rejects a bad one', async (t) => {
    if (requiresR7(t)) return
    const target = path.join(tmpDir, 'permutation.pptx')
    await engine.create(target, { title: 'A' })
    await engine.addSlide(target, { layoutType: 'obj', title: 'B' })
    await engine.addSlide(target, { layoutType: 'obj', title: 'C' })

    await engine.reorderSlides(target, { order: [2, 1, 0] })
    const titles = []
    for (let i = 0; i < 3; i++) {
      const slide = await engine.readSlide(target, i)
      titles.push(slide.slide.objects.find((o) => o.placeholder && (o.placeholder.type === 'title' || o.placeholder.type === 'ctrTitle')).text)
    }
    assert.deepEqual(titles, ['C', 'B', 'A'])

    await assert.rejects(() => engine.reorderSlides(target, { order: [0, 0, 1] }), /permutation/)
    await assert.rejects(() => engine.reorderSlides(target, { order: [0, 1] }), /must list all 3 slides/)
  })

  test('a slide can be cloned from an existing slide instead of a layout', async () => {
    const target = path.join(tmpDir, 'clone.pptx')
    await engine.create(target, { title: 'Источник' })
    // Slide 2 exists in either world: on R7's content layout, or — the fallback
    // having no layouts — as a clone of slide 1. Cloning it again is the point
    // of the test and needs no layout.
    await addSlideAnywhere(engine, target, { layoutType: 'obj', title: 'Шаблон', paragraphs: ['Строка'] })
    const added = await engine.addSlide(target, { baseSlideIndex: 1 })

    const source = await engine.readSlide(target, 1)
    const clone = await engine.readSlide(target, 2)
    assert.equal(clone.slide.layout.partPath, source.slide.layout.partPath, 'the clone shares the layout')
    assert.equal(
      clone.slide.objects.filter((o) => o.onSlide).length,
      source.slide.objects.filter((o) => o.onSlide).length
    )
    assert.equal(added.layout?.partPath ?? source.slide.layout.partPath, source.slide.layout.partPath)
  })

  test('an out-of-range slide index is refused everywhere, with the real count', async () => {
    const target = path.join(tmpDir, 'range.pptx')
    await engine.create(target, { title: 'Один' })
    await assert.rejects(() => engine.readSlide(target, 5), /out of range/)
    await assert.rejects(() => engine.deleteSlide(target, 5), /out of range/)
    await assert.rejects(() => engine.duplicateSlide(target, 5), /out of range/)
    await assert.rejects(
      () => engine.addShape(target, { slideIndex: 5, shape: 'rectangle', x: 0, y: 0, width: 100, height: 100 }),
      /out of range/
    )
  })

  test('notes, master, theme and layouts are preserved by an edit', async (t) => {
    // It asserts on the parts R7's template contributes (master, layouts,
    // theme, notes master, table styles), which the fallback does not have.
    if (requiresR7(t)) return
    const target = path.join(tmpDir, 'preserve-deck.pptx')
    await engine.create(target, { title: 'Презентация' })
    await engine.addSlide(target, { layoutType: 'obj', title: 'Слайд два', paragraphs: ['Текст'] })

    const before = await ZipArchive.fromFile(target)
    const protectedParts = before.list().filter((name) => (
      /^ppt\/slideMasters\//.test(name)
      || /^ppt\/slideLayouts\//.test(name)
      || /^ppt\/theme\//.test(name)
      || /^ppt\/notesMasters\//.test(name)
      || /^ppt\/notesSlides\//.test(name)
      || /^ppt\/tableStyles\.xml$/.test(name)
      || /^ppt\/presProps\.xml$/.test(name)
      || /^ppt\/viewProps\.xml$/.test(name)
      || name === 'docProps/core.xml'
    ))
    assert.ok(protectedParts.length > 15, 'the deck really has these parts')

    const slide = await engine.readSlide(target, 1)
    const body = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'body')
    await engine.formatObject(target, { slideIndex: 1, objectId: body.id, font: { family: 'Georgia', size: 20 }, color: '#1F3864' })
    await engine.addShape(target, {
      slideIndex: 1, shape: 'rounded-rectangle', x: 500000, y: 500000, width: 2000000, height: 800000, fill: '#FFC000', text: 'KPI'
    })

    const after = await ZipArchive.fromFile(target)
    for (const name of protectedParts) {
      assert.ok(after.has(name), `"${name}" still exists`)
      assert.ok(
        sameBytes(before.entries.get(name).raw, after.entries.get(name).raw),
        `"${name}" must be byte-identical after a slide edit`
      )
    }

    // The layout link of the edited slide is untouched, so it still follows the master.
    const relsBefore = readRelationships(before, slide.slide.partPath).map((r) => `${r.type}|${r.target}`)
    const relsAfter = readRelationships(after, slide.slide.partPath).map((r) => `${r.type}|${r.target}`)
    assert.deepEqual(relsAfter, relsBefore, 'the slide keeps its relationships')
  })

  test('a chart, a SmartArt graphic frame and an embedded object survive an edit', async () => {
    const target = path.join(tmpDir, 'chart-smartart.pptx')
    await engine.create(target, { title: 'С диаграммой' })
    // The second slide only has to exist and to carry an editable object: the
    // preservation of the injected chart and SmartArt frame does not depend on
    // which layout the engine built it from.
    await addSlideAnywhere(engine, target, { layoutType: 'obj', title: 'Данные', paragraphs: ['Строка'] })

    const descriptors = slideParts(await ZipArchive.fromFile(target))
    const slidePart = descriptors[1].partPath
    const zip = await ZipArchive.fromFile(target)
    const injected = injectChartAndSmartArt(zip, slidePart)
    await zip.save(target)

    const before = await ZipArchive.fromFile(target)
    const beforeSlideXml = before.getText(slidePart)
    assert.match(beforeSlideXml, /graphicData uri="[^"]*chart"/, 'precondition: the chart frame is present')
    assert.match(beforeSlideXml, /graphicData uri="[^"]*diagram"/, 'precondition: the SmartArt frame is present')

    // Read the deck: the unknown objects must be reported, not dropped.
    const slide = await engine.readSlide(target, 1)
    const chart = slide.slide.objects.find((o) => o.type === 'chart')
    const smartArt = slide.slide.objects.find((o) => o.type === 'graphicFrame' && /diagram/.test(o._element))
    assert.ok(chart, 'the chart is reported as a chart')
    assert.ok(smartArt, 'the SmartArt frame is reported')

    // Change the text of one object and add another.
    const title = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'title')
      || slide.slide.objects.find((o) => o.onSlide && o.type === 'shape')
    await engine.formatObject(target, { slideIndex: 1, objectId: title.id, text: 'Изменённый заголовок', bold: true })
    await engine.addShape(target, {
      slideIndex: 1, shape: 'ellipse', x: 100000, y: 100000, width: 300000, height: 300000, fill: '#00AA00'
    })

    const after = await ZipArchive.fromFile(target)
    const afterSlideXml = after.getText(slidePart)
    assert.match(afterSlideXml, /graphicData uri="[^"]*chart"/, 'the chart frame survived')
    assert.match(afterSlideXml, /graphicData uri="[^"]*diagram"/, 'the SmartArt frame survived')
    assert.ok(afterSlideXml.includes('Изменённый заголовок'), 'the text change landed')
    assert.ok(afterSlideXml.includes('rId900'), 'the SmartArt relationship references are untouched')

    // The chart part itself and its own relationships are byte-identical.
    assert.ok(sameBytes(before.getBuffer(injected.chartPart), after.getBuffer(injected.chartPart)))
    for (const name of before.list()) {
      if (name === slidePart || name === 'ppt/slides/_rels/slide2.xml.rels') continue
      assert.ok(
        sameBytes(before.entries.get(name).raw, after.entries.get(name).raw),
        `"${name}" must be untouched by a text edit on a deck with a chart`
      )
    }

    assert.equal(after.getText(`${injected.chartPart}`), injected.chartXml)
    assert.equal((await engine.validate(target)).valid, true)
  })

  test('the chart deck still renders in R7 after the edit', async (t) => {
    if (requiresR7(t)) return
    const source = path.join(tmpDir, 'render-chart.pptx')
    await engine.create(source, { title: 'Диаграмма' })
    await engine.addSlide(source, { layoutType: 'obj', title: 'Данные', paragraphs: ['Строка'] })
    const descriptors = slideParts(await ZipArchive.fromFile(source))
    const zip = await ZipArchive.fromFile(source)
    const injected = injectChartAndSmartArt(zip, descriptors[1].partPath)
    // The SmartArt frame is a bare `dgm:relIds` with no diagram parts behind
    // it (building a real SmartArt data model is out of scope for v0.1.0), and
    // R7's renderer dereferences those relationships. It is removed for this
    // one check so the PDF proves the *chart* survived, while the preservation
    // test above keeps it and proves the XML does.
    const slideXml = zip.getText(descriptors[1].partPath)
    const smartArtFrame = slideXml.match(/<p:graphicFrame>(?:(?!<\/p:graphicFrame>)[\s\S])*diagram[\s\S]*?<\/p:graphicFrame>/)
    assert.ok(smartArtFrame, 'the SmartArt frame was injected')
    zip.setText(descriptors[1].partPath, slideXml.replace(smartArtFrame[0], ''))
    await zip.save(source)

    const read = await engine.readSlide(source, 1)
    const title = read.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'title')
    await engine.formatObject(source, { slideIndex: 1, objectId: title.id, text: 'После правки', color: '#C00000' })

    const pdfPath = path.join(tmpDir, 'chart.pdf')
    const converted = await engine.toPdf(source, pdfPath)
    assert.equal(converted.success, true)
    assert.ok(converted.isPdf)
    assert.ok(fs.statSync(pdfPath).size > 500)
    assert.equal(injected.chartPart, 'ppt/charts/chart1.xml')
  })

  test('deleting a slide leaves the media other slides use in place', async () => {
    const target = path.join(tmpDir, 'shared-media.pptx')
    await engine.create(target, { title: 'Медиа' })
    await addSlideAnywhere(engine, target, { layoutType: 'obj', title: 'С картинкой' })
    const { writeGradientPng } = await import('./helpers/pptx-fixtures.js')
    const png = writeGradientPng(path.join(tmpDir, 'shared.png'), 120, 90)
    const added = await engine.addImage(target, { slideIndex: 1, imagePath: png.filePath, x: 0, y: 0, width: 1000000 })

    // A second slide referencing the same media part, as a duplicated slide would.
    await engine.duplicateSlide(target, 1)
    const copy = await engine.readSlide(target, 2)
    assert.equal(copy.slide.objects.find((o) => o.type === 'image').image.mediaPath, added.image.mediaPath)

    const zip = await ZipArchive.fromFile(target)
    const descriptors = slideParts(zip)
    const mediaPath = added.image.mediaPath

    await engine.deleteSlide(target, 1)

    const after = await ZipArchive.fromFile(target)
    assert.ok(after.has(mediaPath), 'media another slide still shows is kept')
    assert.equal((await engine.readSlide(target, 1)).slide.objects.filter((o) => o.type === 'image').length, 1)
    assert.equal((await engine.validate(target)).valid, true, 'the surviving reference still resolves')
  })

  test('a dangling slide relationship is reported by validate and dropped by the reader', async () => {
    const target = path.join(tmpDir, 'broken-rel.pptx')
    await engine.create(target, { title: 'A' })
    await addSlideAnywhere(engine, target, { layoutType: 'obj', title: 'B' })

    const zip = await ZipArchive.fromFile(target)
    // Point the second slide relationship at a part that does not exist.
    const relsPath = 'ppt/_rels/presentation.xml.rels'
    const rels = zip.getText(relsPath).replace('Target="slides/slide2.xml"', 'Target="slides/slide99.xml"')
    zip.setText(relsPath, rels)
    await zip.save(target)

    // `slideParts` reports the deck's declared order, so the broken entry is
    // still visible there; validation names it, and reading refuses to invent
    // a slide for it.
    const descriptors = slideParts(await ZipArchive.fromFile(target))
    assert.equal(descriptors.length, 2)

    const validation = await engine.validate(target)
    assert.equal(validation.valid, false)
    assert.ok(validation.errors.some((e) => /missing part|not in the package/.test(e)), JSON.stringify(validation.errors))

    await assert.rejects(() => engine.readSlide(target, 1), /Slide part not found/)
  })

  test('every layout the deck reports can actually host a slide', async (t) => {
    // Its whole subject is the set of layouts the R7 template ships.
    if (requiresR7(t)) return
    const target = path.join(tmpDir, 'every-layout.pptx')
    await engine.create(target, { title: 'Макеты' })
    const { layouts } = await engine.listLayouts(target)
    assert.ok(layouts.length >= 5, `expected the R7 template layouts, got ${layouts.length}`)

    for (const layout of layouts) {
      const added = await engine.addSlide(target, { layoutIndex: layout.index, title: `На макете: ${layout.name}` })
      assert.equal(added.layout.name, layout.name, `layout ${layout.index} was used`)
    }

    const validation = await engine.validate(target)
    assert.equal(validation.valid, true, JSON.stringify(validation.errors))

    const final = await engine.readSlide(target, 0)
    assert.equal(final.slideCount, layouts.length + 1)
    assert.equal(final.layouts.length, layouts.length, 'no layout was created or lost')
  })

  test('the whole deck with every layout renders in R7', async (t) => {
    if (requiresR7(t)) return
    const target = path.join(tmpDir, 'all-layouts-render.pptx')
    await engine.create(target, { title: 'Макеты' })
    const { layouts } = await engine.listLayouts(target)
    for (const layout of layouts) {
      await engine.addSlide(target, { layoutIndex: layout.index, title: layout.name, subtitle: 'Подзаголовок' })
    }
    const pdfPath = path.join(tmpDir, 'all-layouts.pdf')
    const converted = await engine.toPdf(target, pdfPath)
    assert.equal(converted.success, true)
    assert.ok(fs.statSync(pdfPath).size > 5000, 'a multi-page PDF was produced')
  })

  test('the engine never rewrites a part it was not asked to change, across a full workflow', async (t) => {
    // It addresses a shape by the id R7's content layout assigns and checks the
    // slide's layout relationship, so it needs the template.
    if (requiresR7(t)) return
    const target = path.join(tmpDir, 'workflow.pptx')
    await engine.create(target, { title: 'Исходный заголовок' })
    await engine.addSlide(target, { layoutType: 'obj', title: 'Раздел', paragraphs: ['Пункт'] })
    const settled = await ZipArchive.fromFile(target)

    const layoutBefore = slideLayouts(settled)
    const descriptorsBefore = slideParts(settled)
    const partTwo = descriptorsBefore[1].partPath

    await engine.formatObject(target, { slideIndex: 1, objectId: 2, bold: true })
    await engine.addTextBox(target, { slideIndex: 1, x: 0, y: 0, width: 1000000, height: 500000, text: 'Добавлено' })

    const after = await ZipArchive.fromFile(target)
    // The layout on disk is identical, and the resolved layout list did not move.
    assert.deepEqual(
      slideLayouts(after).map((l) => `${l.partPath}|${l.name}|${l.type}`),
      layoutBefore.map((l) => `${l.partPath}|${l.name}|${l.type}`)
    )
    for (const name of settled.list()) {
      if (name === partTwo) continue
      assert.ok(
        sameBytes(settled.entries.get(name).raw, after.entries.get(name).raw),
        `"${name}" must be untouched`
      )
    }

    // And the layout relationship of the edited slide is byte-identical.
    const relsBefore = readRelationships(settled, partTwo).find((r) => /slideLayout$/.test(r.type))
    const relsAfter = readRelationships(after, partTwo).find((r) => /slideLayout$/.test(r.type))
    assert.deepEqual(relsAfter, relsBefore)
    assert.equal(resolvePartPath(partTwo, relsAfter.target), relsAfter.partPath)
  })
})
