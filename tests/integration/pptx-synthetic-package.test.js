import { test, describe, after } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import { PptxEngine } from '../../src/r7/pptx.js'
import { ZipArchive } from '../../src/shared/zip.js'
import { tempDir, writeGradientPng } from './helpers/pptx-fixtures.js'

/**
 * The synthetic package, exercised on its own.
 *
 * When no R7-Office template can be found the engine does not refuse to work:
 * `create()` writes a minimal but well-formed PPTX package of its own. That
 * path is what a clean machine (and CI) actually runs, so it is tested here
 * directly rather than being tolerated by tests that wish R7 were installed.
 *
 * What the fallback genuinely supports is asserted below. What it does NOT
 * support is asserted too, with the engine's own error: the package carries no
 * slide layout, master or theme, so a slide cannot be built *on a layout*. The
 * documented alternative — cloning an existing slide with `baseSlideIndex` —
 * does work and is covered here.
 *
 * The engine is driven through the `R7_OFFICE_DISABLED` detection seam so the
 * fallback is also exercised on a developer machine that has R7 installed;
 * otherwise this file would silently stop testing anything there.
 */
describe('PPTX on a machine without R7: the synthetic package', () => {
  const tmpDir = tempDir('pptx_synthetic')

  /**
   * A fresh synthetic deck and the engine that produced it.
   *
   * `R7_OFFICE_DISABLED` must be set while the engine's first detection runs:
   * `R7Adapter.detect()` caches its answer on the adapter instance, so once
   * `create()` has asked, this engine stays on the fallback path for its whole
   * life. The flag is restored as soon as the deck exists, so nothing else in
   * this process is affected. On a machine without R7 the flag is already set,
   * and the deck is the same either way.
   */
  async function syntheticDeck(name, title = 'Синтетическая колода') {
    const previous = process.env.R7_OFFICE_DISABLED
    process.env.R7_OFFICE_DISABLED = '1'
    try {
      const engine = new PptxEngine()
      const target = path.join(tmpDir, name)
      await engine.create(target, { overwrite: true, title })
      return { engine, target }
    } finally {
      if (previous === undefined) delete process.env.R7_OFFICE_DISABLED
      else process.env.R7_OFFICE_DISABLED = previous
    }
  }

  after(() => {
    try {
      fs.rmSync(tmpDir, { recursive: true, force: true })
    } catch {
      // best effort
    }
  })

  test('create() writes a self-contained package and invents no R7 part', async () => {
    const { target } = await syntheticDeck('package.pptx')
    const zip = await ZipArchive.fromFile(target)

    for (const part of [
      '[Content_Types].xml',
      '_rels/.rels',
      'ppt/presentation.xml',
      'ppt/_rels/presentation.xml.rels',
      'ppt/slides/slide1.xml'
    ]) {
      assert.ok(zip.has(part), `${part} is present`)
    }

    for (const part of [
      'ppt/slideLayouts/slideLayout1.xml',
      'ppt/slideMasters/slideMaster1.xml',
      'ppt/theme/theme1.xml',
      'ppt/tableStyles.xml'
    ]) {
      assert.equal(zip.has(part), false, `${part} must not be invented`)
    }
  })

  test('the package validates, and the missing layout link is reported as a warning', async () => {
    const { engine, target } = await syntheticDeck('validate.pptx')
    const validation = await engine.validate(target)

    assert.equal(validation.valid, true, JSON.stringify(validation.errors))
    assert.deepEqual(validation.errors, [])
    // Honest reporting: the deck is valid OOXML but it really has no layout, and
    // the validator says so instead of staying quiet about it.
    assert.ok(
      validation.warnings.some((w) => /no slideLayout relationship/.test(w)),
      `expected a layout warning, got: ${JSON.stringify(validation.warnings)}`
    )
    assert.equal(validation.details.layouts, 0)
  })

  test('readSlide reports the slide, its size, and the title create() was given', async () => {
    const { engine, target } = await syntheticDeck('read.pptx', 'Заголовок синтетики')
    const slide = await engine.readSlide(target, 0)

    assert.equal(slide.slideCount, 1)
    assert.deepEqual(slide.layouts, [], 'the fallback declares no layouts')
    assert.equal(slide.slide.layout.partPath, null)
    assert.equal(slide.slide.size.width, 12192000)
    assert.equal(slide.slide.size.height, 6858000)

    const title = slide.slide.objects.find((o) => o.onSlide)
    assert.ok(title, 'the slide carries the title shape create() writes')
    assert.equal(title.text, 'Заголовок синтетики')
    assert.equal(title.name, 'Title')
  })

  test('inspect() and listLayouts() agree the deck is one slide with no layouts', async () => {
    const { engine, target } = await syntheticDeck('inspect.pptx', 'Осмотр')
    const inspection = await engine.inspect(target)

    assert.equal(inspection.type, 'pptx')
    assert.equal(inspection.slidesCount, 1)
    assert.deepEqual(inspection.layouts, [])
    assert.equal(inspection.slides[0].title, 'Осмотр')

    const listed = await engine.listLayouts(target)
    assert.deepEqual(listed.layouts, [])
  })

  test('the object layer — shapes, text boxes and images — works without a layout', async () => {
    const { engine, target } = await syntheticDeck('objects.pptx')

    const shape = await engine.addShape(target, {
      slideIndex: 0, shape: 'rounded-rectangle',
      x: 100000, y: 200000, width: 3000000, height: 1000000,
      fill: '#1F6FEB', line: '#0B3D91', lineWidth: 2
    })
    const box = await engine.addTextBox(target, {
      slideIndex: 0, x: 0, y: 2000000, width: 2000000, height: 500000, text: 'Надпись'
    })
    const png = writeGradientPng(path.join(tmpDir, 'picture.png'), 120, 90)
    const picture = await engine.addImage(target, {
      slideIndex: 0, imagePath: png.filePath, x: 4000000, y: 0, width: 1200000
    })

    const slide = await engine.readSlide(target, 0)
    const painted = slide.slide.objects.find((o) => o.id === shape.objectId)
    assert.ok(painted, 'the shape reads back')
    assert.equal(painted.preset, 'roundRect')
    assert.equal(painted.fill.color, '#1F6FEB')
    assert.equal(painted.stroke.color, '#0B3D91')
    assert.equal(painted.positionValid, true)

    assert.equal(slide.slide.objects.find((o) => o.id === box.objectId).text, 'Надпись')

    const image = slide.slide.objects.find((o) => o.id === picture.objectId)
    assert.equal(image.type, 'image')
    assert.equal(image.image.format, 'png')
    assert.equal(image.width, 1200000)
    assert.equal(image.height, 900000, 'the source ratio is kept')

    // Adding objects must not invalidate the package the fallback wrote.
    const validation = await engine.validate(target)
    assert.equal(validation.valid, true, JSON.stringify(validation.errors))
  })

  test('a slide cannot be built on a layout, and the error says exactly that', async () => {
    const { engine, target } = await syntheticDeck('no-layout.pptx')

    // This is the fallback's real limitation, not a test accident: the package
    // has no layouts to build on. Asking for one must fail loudly and name the
    // documented way out rather than writing a broken slide.
    await assert.rejects(
      () => engine.addSlide(target, { layoutType: 'blank' }),
      /No slide layout available in this presentation\. Pass baseSlideIndex to clone an existing slide\./
    )
    await assert.rejects(
      () => engine.addSlide(target, { title: 'Заголовок' }),
      /No slide layout available/
    )
    // And nothing was half-written by the refused calls.
    assert.equal((await engine.readSlide(target, 0)).slideCount, 1)
  })

  test('a slide CAN be added by cloning an existing one, the documented fallback', async () => {
    const { engine, target } = await syntheticDeck('clone.pptx')

    const added = await engine.addSlide(target, { baseSlideIndex: 0 })
    assert.equal(added.slideCount, 2)
    assert.equal(added.partPath, 'ppt/slides/slide2.xml')

    const source = await engine.readSlide(target, 0)
    const clone = await engine.readSlide(target, 1)
    assert.equal(clone.slideCount, 2)
    assert.equal(clone.slide.layout.partPath, null, 'no layout is invented for the clone')
    assert.equal(
      clone.slide.objects.filter((o) => o.onSlide).length,
      source.slide.objects.filter((o) => o.onSlide).length,
      'the clone carries the same objects as its source'
    )
    // Two objects sharing a shape id would make an editor refuse to touch
    // either of them, so the clone's ids must be fresh.
    const sourceIds = new Set(source.slide.objects.filter((o) => o.onSlide).map((o) => o.id))
    for (const object of clone.slide.objects.filter((o) => o.onSlide)) {
      assert.equal(sourceIds.has(object.id), false, `id ${object.id} must not be reused`)
    }
  })

  test('the slide lifecycle works on the synthetic deck without rewriting it', async () => {
    const { engine, target } = await syntheticDeck('lifecycle.pptx')
    await engine.addSlide(target, { baseSlideIndex: 0 })

    const duplicated = await engine.duplicateSlide(target, 1)
    assert.equal(duplicated.slideCount, 3)

    const moved = await engine.moveSlide(target, 2, 0)
    assert.equal(moved.slideCount, 3)
    assert.deepEqual(moved.order, ['ppt/slides/slide3.xml', 'ppt/slides/slide1.xml', 'ppt/slides/slide2.xml'])

    const deleted = await engine.deleteSlide(target, 0)
    assert.equal(deleted.slideCount, 2)

    const validation = await engine.validate(target)
    assert.equal(validation.valid, true, JSON.stringify(validation.errors))
    assert.equal((await engine.readSlide(target, 0)).slideCount, 2)
  })
})
