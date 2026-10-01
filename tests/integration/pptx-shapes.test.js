import { test, describe, before, after } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import { PptxEngine } from '../../src/r7/pptx.js'
import { ZipArchive } from '../../src/shared/zip.js'
import { requiresR7, HAS_R7_TEMPLATES } from '../helpers/r7-gate.js'
import { tempDir, writeGradientPng, writeSolidPng, addSlideAnywhere, clearSlide } from './helpers/pptx-fixtures.js'

/**
 * Shapes, fills, strokes, z-order and images.
 *
 * Every shape the engine advertises is exercised, because a catalogue that
 * lists a name the writer cannot emit is worse than a short catalogue.
 */
describe('PPTX shapes, paint and images', () => {
  const tmpDir = tempDir('pptx_shapes')
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

  /**
   * The objects the author put on the slide: everything the engine created,
   * plus the content placeholders, minus the graphic frames and the date,
   * footer and slide-number fields that belong to the layout.
   */
  function authoredObjects(slide) {
    const decorative = new Set(['dt', 'ftr', 'sldNum', 'hdr'])
    return slide.objects.filter((o) => {
      if (!o.onSlide) return false
      if (o.type === 'graphicFrame' || o.type === 'table' || o.type === 'chart') return false
      if (o.placeholder && decorative.has(o.placeholder.type)) return false
      return true
    })
  }

  /**
   * A one-slide deck whose second slide starts empty, so shapes are the only
   * objects on it.
   *
   * With R7 the second slide is built on the template's blank layout. Without
   * R7 the package carries no layouts, so the engine can only clone the first
   * slide; the clone is stripped of the title shape it inherits so the test
   * still sees nothing but what it added.
   */
  async function blankDeck(name) {
    const target = path.join(tmpDir, name)
    await engine.create(target, { overwrite: true, title: 'Фигуры' })
    await addSlideAnywhere(engine, target, { layoutType: 'blank' })
    if (!HAS_R7_TEMPLATES) await clearSlide(engine, target, 1)
    return target
  }

  const EVERY_SHAPE = [
    'rectangle', 'rounded-rectangle', 'ellipse', 'circle', 'line', 'arrow',
    'triangle', 'diamond', 'pentagon', 'hexagon', 'octagon', 'star',
    'chevron', 'plus', 'cloud', 'heart', 'cylinder', 'cube', 'donut', 'pie',
    'parallelogram', 'trapezoid', 'arrow-left', 'arrow-up', 'arrow-down'
  ]

  test('every advertised shape type is written and reads back with its geometry', async () => {
    const target = await blankDeck('all-shapes.pptx')
    const created = []
    let x = 300000
    for (const shape of EVERY_SHAPE) {
      const result = await engine.addShape(target, {
        slideIndex: 1,
        shape,
        x,
        y: 1000000,
        width: 800000,
        height: 600000,
        fill: '#4C78A8',
        lineColor: '#1F3A5F',
        lineWidth: 1
      })
      created.push({ shape, id: result.objectId })
      x += 900000
      if (x > 11000000) x = 300000
    }

    const slide = await engine.readSlide(target, 1)
    const onSlide = slide.slide.objects.filter((o) => o.onSlide)
    assert.equal(onSlide.filter((o) => !o.placeholder).length, EVERY_SHAPE.length, 'every shape produced one object')
    for (const { shape, id } of created) {
      const object = slide.slide.objects.find((o) => o.id === id)
      assert.ok(object, `${shape} (id ${id}) is on the slide`)
      assert.equal(object.positionValid, true, `${shape} has usable geometry`)
      assert.equal(object.width, 800000, `${shape} width`)
      assert.equal(object.height, 600000, `${shape} height`)
    }
  })

  test('the shape catalogue rejects an unknown name with a helpful list', async () => {
    const target = await blankDeck('bad-shape.pptx')
    await assert.rejects(
      () => engine.addShape(target, { slideIndex: 1, shape: 'dodecahedron', x: 0, y: 0, width: 100, height: 100 }),
      /Unsupported shape type "dodecahedron"/
    )
  })

  test('a rounded rectangle keeps its preset geometry', async () => {
    const target = await blankDeck('roundrect.pptx')
    const created = await engine.addShape(target, {
      slideIndex: 1, shape: 'rounded-rectangle', x: 100000, y: 100000, width: 2000000, height: 1000000
    })
    const slide = await engine.readSlide(target, 1)
    const object = slide.slide.objects.find((o) => o.id === created.objectId)
    assert.equal(object.preset, 'roundRect')
  })

  test('fill, transparency, stroke colour and width are written and read back', async () => {
    const target = await blankDeck('paint.pptx')
    const created = await engine.addShape(target, {
      slideIndex: 1,
      shape: 'rectangle',
      x: 500000, y: 500000, width: 3000000, height: 2000000,
      fill: '#1F6FEB',
      fillTransparency: 0.25,
      line: '#0B3D91',
      lineWidth: 2
    })

    const slide = await engine.readSlide(target, 1)
    const object = slide.slide.objects.find((o) => o.id === created.objectId)
    assert.equal(object.fill.kind, 'solid')
    assert.equal(object.fill.color, '#1F6FEB')
    // Alpha is an 8-bit channel, so 0.25 comes back as 0.251.
    assert.ok(Math.abs(object.fill.transparency - 0.25) < 0.01, 'fill transparency is about 0.25')
    assert.equal(object.stroke.color, '#0B3D91')
    assert.equal(object.stroke.width, 25400, '2 pt in EMU')
    assert.equal(object.stroke.style, 'solid')
  })

  test('a shape can be repainted, restroked and made transparent after creation', async () => {
    const target = await blankDeck('repaint.pptx')
    const created = await engine.addShape(target, {
      slideIndex: 1, shape: 'ellipse', x: 500000, y: 500000, width: 2000000, height: 2000000, fill: '#FF0000'
    })

    await engine.formatObject(target, {
      slideIndex: 1, objectId: created.objectId,
      fill: '#00AA55', fillTransparency: 0.5, line: '#000000', lineWidth: 3, lineStyle: 'dash'
    })

    const slide = await engine.readSlide(target, 1)
    const object = slide.slide.objects.find((o) => o.id === created.objectId)
    assert.equal(object.fill.color, '#00AA55')
    // Alpha is an 8-bit channel, so 0.5 comes back as 0.498.
    assert.ok(Math.abs(object.fill.transparency - 0.5) < 0.01, 'fill transparency is about 0.5')
    assert.equal(object.stroke.color, '#000000')
    assert.equal(object.stroke.width, 38100, '3 pt in EMU')
    assert.equal(object.stroke.style, 'dash')
  })

  test('noFill and noLine remove paint without deleting the object', async () => {
    const target = await blankDeck('nofill.pptx')
    const created = await engine.addShape(target, {
      slideIndex: 1, shape: 'rectangle', x: 500000, y: 500000, width: 1000000, height: 1000000,
      fill: '#FF0000', line: '#000000', lineWidth: 2
    })

    await engine.formatObject(target, { slideIndex: 1, objectId: created.objectId, noFill: true, noLine: true })

    const again = (await engine.readSlide(target, 1)).slide.objects.find((o) => o.id === created.objectId)
    assert.ok(again, 'the object exists')
    assert.equal(again.fill.kind, 'none')
    assert.equal(again.stroke.none, true)
  })

  test('an em dash and an arrow glyph survive as a bullet character', async () => {
    const target = await blankDeck('glyph.pptx')
    const created = await engine.addTextBox(target, {
      slideIndex: 1, x: 100000, y: 100000, width: 4000000, height: 1000000,
      paragraphs: [{ text: 'Стрелка', bullet: '→' }]
    })
    const slide = await engine.readSlide(target, 1)
    const object = slide.slide.objects.find((o) => o.id === created.objectId)
    assert.equal(object.paragraphs[0].bulletCharacter, '→')
  })

  test('z-order controls the object list order and can be changed later', async () => {
    const target = await blankDeck('zorder.pptx')
    const first = await engine.addShape(target, {
      slideIndex: 1, shape: 'rectangle', x: 100000, y: 100000, width: 1000000, height: 1000000, name: 'Первый'
    })
    const second = await engine.addShape(target, {
      slideIndex: 1, shape: 'rectangle', x: 200000, y: 200000, width: 1000000, height: 1000000, name: 'Второй'
    })

    let slide = await engine.readSlide(target, 1)
    let ordered = authoredObjects(slide.slide)
    assert.deepEqual(ordered.map((o) => o.name), ['Первый', 'Второй'])

    // Insert a third object at z-order 0, which is behind both of them.
    await engine.addShape(target, {
      slideIndex: 1, shape: 'ellipse', x: 0, y: 0, width: 500000, height: 500000, name: 'Фон', zOrder: 0
    })
    slide = await engine.readSlide(target, 1)
    ordered = authoredObjects(slide.slide)
    assert.deepEqual(ordered.map((o) => o.name), ['Фон', 'Первый', 'Второй'])
    // `zOrder` is the position in the shape tree, so inserting behind the two
    // existing shapes has to put the new one first and leave the others after.
    assert.ok(ordered[0].zOrder < ordered[1].zOrder, 'the new object is behind')
    assert.ok(ordered[1].zOrder < ordered[2].zOrder, 'the original order is preserved')

    // The first object is still addressable by id after the insertion.
    const stillThere = slide.slide.objects.find((o) => o.id === first.objectId)
    assert.equal(stillThere.name, 'Первый')
    assert.equal(slide.slide.objects.find((o) => o.id === second.objectId).name, 'Второй')
  })

  test('a line with an arrow head keeps its preset and arrow', async () => {
    const target = await blankDeck('arrow.pptx')
    const created = await engine.addShape(target, {
      slideIndex: 1, shape: 'line', x: 500000, y: 500000, width: 3000000, height: 0,
      line: '#C00000', lineWidth: 2, arrows: true
    })
    const zip = await ZipArchive.fromFile(target)
    const { slideParts } = await import('../../src/r7/pptx-util.js')
    const xml = zip.getText(slideParts(zip)[1].partPath)
    assert.match(xml, /<a:tailEnd type="triangle"/, 'the arrow is at the pointed end')
    const slide = await engine.readSlide(target, 1)
    const object = slide.slide.objects.find((o) => o.id === created.objectId)
    assert.ok(object, 'the line reads back')
  })

  test('a text box has no fill and no border unless asked for one', async () => {
    const target = await blankDeck('textbox-plain.pptx')
    const created = await engine.addTextBox(target, {
      slideIndex: 1, x: 100000, y: 100000, width: 3000000, height: 1000000, text: 'Надпись'
    })
    const slide = await engine.readSlide(target, 1)
    const object = slide.slide.objects.find((o) => o.id === created.objectId)
    assert.equal(object.fill.kind, 'none')
    assert.equal(object.stroke.none, true)
    assert.equal(object.text, 'Надпись')
  })

  test('a text box with a background and a border is possible too', async () => {
    const target = await blankDeck('textbox-filled.pptx')
    const created = await engine.addTextBox(target, {
      slideIndex: 1, x: 100000, y: 100000, width: 3000000, height: 1000000, text: 'Надпись',
      fill: '#FFF2CC', line: '#BF9000', lineWidth: 1
    })
    const slide = await engine.readSlide(target, 1)
    const object = slide.slide.objects.find((o) => o.id === created.objectId)
    assert.equal(object.fill.color, '#FFF2CC')
    assert.equal(object.stroke.color, '#BF9000')
  })

  test('removeObject takes a shape off the slide and leaves the rest alone', async () => {
    const target = await blankDeck('remove.pptx')
    const keep = await engine.addShape(target, {
      slideIndex: 1, shape: 'rectangle', x: 100000, y: 100000, width: 1000000, height: 1000000, name: 'Оставить'
    })
    const drop = await engine.addShape(target, {
      slideIndex: 1, shape: 'ellipse', x: 2000000, y: 100000, width: 1000000, height: 1000000, name: 'Убрать'
    })

    await engine.removeObject(target, { slideIndex: 1, objectId: drop.objectId })

    const slide = await engine.readSlide(target, 1)
    assert.equal(slide.slide.objects.find((o) => o.id === drop.objectId), undefined)
    assert.ok(slide.slide.objects.find((o) => o.id === keep.objectId), 'the other shape survives')
  })

  test('removing an object that exists only in the layout is refused, not guessed', async (t) => {
    // Its subject is the layout graph: the placeholder lives on a real R7
    // layout, and without one the slide inherits nothing to refuse.
    if (requiresR7(t)) return
    const target = path.join(tmpDir, 'layout-only.pptx')
    await engine.create(target, { overwrite: true, title: 'Только заголовок' })
    await engine.addSlide(target, { layoutType: 'titleOnly', title: 'Заголовок' })
    const slide = await engine.readSlide(target, 1)
    const inherited = slide.slide.objects.find((o) => !o.onSlide)
    assert.ok(inherited, 'the layout contributes a placeholder the slide does not carry')
    await assert.rejects(
      () => engine.removeObject(target, { slideIndex: 1, objectId: inherited.id, placeholderType: inherited.placeholder.type }),
      /exists only in the layout/
    )
  })

  test('styling a layout-only placeholder materialises it on the slide', async (t) => {
    // Materialising is defined against a layout placeholder, so it needs R7.
    if (requiresR7(t)) return
    const target = path.join(tmpDir, 'materialise.pptx')
    await engine.create(target, { overwrite: true, title: 'Материализация' })
    await engine.addSlide(target, { layoutType: 'titleOnly', title: 'Заголовок' })

    const before = await engine.readSlide(target, 1)
    const inherited = before.slide.objects.find((o) => !o.onSlide && o.placeholder)
    assert.ok(inherited, 'precondition: the layout declares a placeholder the slide does not')

    await engine.formatObject(target, {
      slideIndex: 1,
      objectId: inherited.id,
      placeholderType: inherited.placeholder.type,
      font: { family: 'Verdana', size: 14 },
      color: '#333333'
    })

    const after = await engine.readSlide(target, 1)
    const materialised = after.slide.objects.find(
      (o) => o.onSlide && o.placeholder && o.placeholder.type === inherited.placeholder.type
    )
    assert.ok(materialised, 'the placeholder now lives on the slide')
    assert.equal(materialised.font.family, 'Verdana')
  })

  // ------------------------------------------------------------------ images

  test('a PNG is inserted with its natural size and reads back with its media part', async () => {
    const target = await blankDeck('image.pptx')
    const png = writeGradientPng(path.join(tmpDir, 'gradient.png'), 240, 160)
    const created = await engine.addImage(target, { slideIndex: 1, imagePath: png.filePath, x: 1000000, y: 800000, width: 3000000 })

    assert.ok(created.image, 'the result names the media part')
    const slide = await engine.readSlide(target, 1)
    const object = slide.slide.objects.find((o) => o.id === created.objectId)
    assert.equal(object.type, 'image')
    assert.equal(object.image.format, 'png')
    assert.equal(object.image.bytes, png.bytes)
    assert.equal(object.width, 3000000)
    assert.equal(object.height, 2000000, '240x160 is a 3:2 ratio')
    assert.equal(object.x, 1000000)
    assert.equal(object.y, 800000)

    const zip = await ZipArchive.fromFile(target)
    assert.ok(zip.has(created.image.mediaPath), 'the media part exists in the package')
    assert.ok(zip.getBuffer(created.image.mediaPath).equals(fs.readFileSync(png.filePath)), 'bytes are stored verbatim')

    const contentTypes = zip.getText('[Content_Types].xml')
    assert.match(contentTypes, /Extension="png"/, 'the PNG content type is registered')
  })

  test('the natural aspect ratio is kept when only the height is given', async () => {
    const target = await blankDeck('image-ratio.pptx')
    const png = writeGradientPng(path.join(tmpDir, 'ratio.png'), 400, 100)
    const created = await engine.addImage(target, { slideIndex: 1, imagePath: png.filePath, x: 0, y: 0, height: 1000000 })
    const slide = await engine.readSlide(target, 1)
    const object = slide.slide.objects.find((o) => o.id === created.objectId)
    assert.equal(object.height, 1000000)
    assert.equal(object.width, 4000000, '4:1 source ratio')
  })

  test('an image with no size is placed at its natural pixel size, scaled to fit the slide', async () => {
    const target = await blankDeck('image-natural.pptx')
    const png = writeGradientPng(path.join(tmpDir, 'natural.png'), 300, 200)
    const created = await engine.addImage(target, { slideIndex: 1, imagePath: png.filePath, x: 500000, y: 500000 })
    const slide = await engine.readSlide(target, 1)
    const object = slide.slide.objects.find((o) => o.id === created.objectId)
    // 300 px at 96 dpi is 2857500 EMU, which fits a 12192000 EMU slide.
    assert.equal(object.width, 2857500)
    assert.equal(object.height, 1905000)
  })

  test('lockAspectRatio false stretches an image to the requested box', async () => {
    const target = await blankDeck('image-stretch.pptx')
    const png = writeSolidPng(path.join(tmpDir, 'solid.png'), [10, 20, 30], 300, 200)
    const created = await engine.addImage(target, {
      slideIndex: 1, imagePath: png.filePath, x: 0, y: 0, width: 4000000, height: 1000000, lockAspectRatio: false
    })
    const slide = await engine.readSlide(target, 1)
    const object = slide.slide.objects.find((o) => o.id === created.objectId)
    assert.equal(object.width, 4000000)
    assert.equal(object.height, 1000000)
  })

  test('an image can be moved and resized without touching its bytes', async () => {
    const target = await blankDeck('image-move.pptx')
    const png = writeGradientPng(path.join(tmpDir, 'move.png'), 120, 90)
    const created = await engine.addImage(target, { slideIndex: 1, imagePath: png.filePath, x: 0, y: 0, width: 1000000 })
    const before = await ZipArchive.fromFile(target)

    await engine.formatObject(target, {
      slideIndex: 1, objectId: created.objectId, x: 2000000, y: 3000000, width: 2000000, height: 1500000
    })

    const after = await ZipArchive.fromFile(target)
    assert.equal(
      before.entries.get(created.image.mediaPath).raw.toString('base64'),
      after.entries.get(created.image.mediaPath).raw.toString('base64'),
      'moving an image must not rewrite its data'
    )

    const slide = await engine.readSlide(target, 1)
    const object = slide.slide.objects.find((o) => o.id === created.objectId)
    assert.equal(object.x, 2000000)
    assert.equal(object.y, 3000000)
    assert.equal(object.width, 2000000)
    assert.equal(object.height, 1500000)
  })

  test('replaceImage swaps the picture and keeps the relationship valid', async () => {
    const target = await blankDeck('image-replace.pptx')
    const firstPng = writeGradientPng(path.join(tmpDir, 'first.png'), 200, 200)
    const created = await engine.addImage(target, {
      slideIndex: 1, imagePath: firstPng.filePath, x: 500000, y: 500000, width: 2000000
    })

    const replacementPng = writeSolidPng(path.join(tmpDir, 'replacement.png'), [200, 30, 30], 100, 100)
    const result = await engine.formatObject(target, {
      slideIndex: 1, objectId: created.objectId, imagePath: replacementPng.filePath
    })
    assert.ok(result.image, 'the result reports the replacement')

    const zip = await ZipArchive.fromFile(target)
    const mediaPath = result.image.mediaPath
    assert.ok(zip.has(mediaPath), 'the replacement media part exists')
    assert.equal(zip.getBuffer(mediaPath).length, replacementPng.bytes)

    const slide = await engine.readSlide(target, 1)
    const object = slide.slide.objects.find((o) => o.id === created.objectId)
    assert.equal(object.image.mediaPath, mediaPath)
    assert.equal(object.image.format, 'png')

    // The relationship the slide declares must still resolve.
    const { resolvePartPath, readRelationships } = await import('../../src/r7/pptx-util.js')
    const rels = readRelationships(zip, slide.slide.partPath)
    const imageRel = rels.find((r) => r.id === object.image.relId)
    assert.ok(imageRel, 'the image relationship exists')
    assert.equal(imageRel.partPath, mediaPath)
    assert.equal(resolvePartPath(slide.slide.partPath, imageRel.target), mediaPath)
  })

  test('replacing an image keeps every other part of the package frozen', async () => {
    const target = await blankDeck('image-preserve.pptx')
    const firstPng = writeGradientPng(path.join(tmpDir, 'before.png'), 200, 200)
    const created = await engine.addImage(target, {
      slideIndex: 1, imagePath: firstPng.filePath, x: 500000, y: 500000, width: 2000000
    })
    const before = await ZipArchive.fromFile(target)

    const replacementPng = writeSolidPng(path.join(tmpDir, 'after.png'), [1, 2, 3], 100, 100)
    await engine.formatObject(target, { slideIndex: 1, objectId: created.objectId, imagePath: replacementPng.filePath })
    const after = await ZipArchive.fromFile(target)

    const { slideParts } = await import('../../src/r7/pptx-util.js')
    const editedSlide = slideParts(after)[1].partPath
    const changed = new Set([editedSlide, created.image.mediaPath, 'ppt/slides/_rels/slide2.xml.rels'])
    for (const name of before.list()) {
      if (changed.has(name)) continue
      assert.equal(
        before.entries.get(name).raw.toString('base64'),
        after.entries.get(name).raw.toString('base64'),
        `"${name}" must be untouched by an image replacement`
      )
    }
  })

  test('an unsupported image format is refused with the supported list', async () => {
    const target = await blankDeck('image-bad.pptx')
    const bogus = path.join(tmpDir, 'not-an-image.png')
    fs.writeFileSync(bogus, Buffer.from('this is definitely not an image'))
    await assert.rejects(
      () => engine.addImage(target, { slideIndex: 1, imagePath: bogus, x: 0, y: 0, width: 100000 }),
      /Unsupported image format/
    )
  })

  test('a missing image path is reported, not silently skipped', async () => {
    const target = await blankDeck('image-missing.pptx')
    await assert.rejects(
      () => engine.addImage(target, { slideIndex: 1, imagePath: path.join(tmpDir, 'nope.png'), x: 0, y: 0, width: 100000 }),
      /Image not found/
    )
  })

  test('the deck with shapes and an image renders to PDF in R7', async (t) => {
    if (requiresR7(t)) return
    const target = await blankDeck('render-shapes.pptx')
    await engine.addShape(target, {
      slideIndex: 1, shape: 'rounded-rectangle', x: 500000, y: 500000, width: 3000000, height: 1500000,
      fill: '#1F6FEB', line: '#0B3D91', lineWidth: 2, text: 'Показатель 1', color: '#FFFFFF', size: 20, alignment: 'center'
    })
    await engine.addShape(target, {
      slideIndex: 1, shape: 'ellipse', x: 4000000, y: 500000, width: 1500000, height: 1500000, fill: '#E45756'
    })
    await engine.addShape(target, {
      slideIndex: 1, shape: 'line', x: 500000, y: 2500000, width: 5000000, height: 0, line: '#000000', lineWidth: 2, arrows: true
    })
    const png = writeGradientPng(path.join(tmpDir, 'render.png'), 200, 150)
    await engine.addImage(target, { slideIndex: 1, imagePath: png.filePath, x: 6000000, y: 1000000, width: 3000000 })

    const pdfPath = path.join(tmpDir, 'shapes.pdf')
    const converted = await engine.toPdf(target, pdfPath)
    assert.equal(converted.success, true)
    assert.ok(converted.isPdf, 'the converter produced a PDF')
    assert.ok(fs.statSync(pdfPath).size > 1000, 'the PDF has real content')
  })

  test('a deck full of shapes still validates and reopens', async () => {
    const target = await blankDeck('validate-shapes.pptx')
    for (const shape of ['rectangle', 'rounded-rectangle', 'ellipse', 'triangle', 'star']) {
      await engine.addShape(target, {
        slideIndex: 1, shape, x: 500000, y: 500000, width: 1000000, height: 1000000, fill: '#8899AA'
      })
    }
    const validation = await engine.validate(target)
    assert.equal(validation.valid, true, JSON.stringify(validation.errors))
    const slide = await engine.readSlide(target, 1)
    assert.equal(slide.slide.objects.filter((o) => o.onSlide && !o.placeholder).length, 5)
  })

  /**
   * `a:srgbClr/@val` is ST_HexColorRGB: exactly six hex digits. Writing an
   * eight-digit AARRGGBB there is invalid, and renderers disagree about the
   * extra byte — LibreOffice draws the colour, while R7's converter read it as
   * black and painted an opaque black block over the slide. Opacity belongs in
   * a child a:alpha, in thousandths of a percent.
   */
  test('colours are written as six-digit RGB with alpha as a child element', async () => {
    const target = await blankDeck('colour-validity.pptx')
    await engine.addShape(target, {
      slideIndex: 1, shape: 'rectangle', x: 500000, y: 500000, width: 2000000, height: 1000000,
      fill: { color: '#F3F6FB' }, line: { color: '#D2DAE6', width: 1 }
    })
    await engine.addShape(target, {
      slideIndex: 1, shape: 'rectangle', x: 3000000, y: 500000, width: 2000000, height: 1000000,
      fill: { color: '#1F6FEB', transparency: 0.4 }, noLine: true
    })

    const zip = await ZipArchive.fromFile(target)
    const xml = zip.getText('ppt/slides/slide2.xml')

    const values = [...xml.matchAll(/<a:srgbClr\s+val="([^"]*)"/g)].map((m) => m[1])
    assert.ok(values.length > 0, 'the slide should carry srgbClr elements')
    for (const v of values) {
      assert.match(v, /^[0-9A-Fa-f]{6}$/, `srgbClr val must be exactly 6 hex digits, got "${v}"`)
    }

    // The translucent fill must express its opacity as a child, not in `val`.
    assert.match(xml, /<a:srgbClr\s+val="1F6FEB"><a:alpha\s+val="60000"\/><\/a:srgbClr>/,
      'a 0.4 transparency must be written as <a:alpha val="60000"/> inside a 6-digit srgbClr')

    // And it must read back as the same colour and transparency.
    const slide = await engine.readSlide(target, 1)
    const painted = slide.slide.objects.filter((o) => o.onSlide && !o.placeholder)
    const translucent = painted.find((o) => o.fill && o.fill.transparency > 0.3)
    assert.ok(translucent, 'the translucent shape should be reported with its transparency')
    assert.ok(Math.abs(translucent.fill.transparency - 0.4) < 0.01,
      `expected transparency ~0.4, got ${translucent.fill.transparency}`)
  })
})

