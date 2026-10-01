import { test, describe, before, after } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import { PptxEngine } from '../../src/r7/pptx.js'
import { ZipArchive } from '../../src/shared/zip.js'
import { requiresR7, R7_AVAILABLE, HAS_R7_TEMPLATES } from '../helpers/r7-gate.js'
import { tempDir, addSlideAnywhere, clearSlide } from './helpers/pptx-fixtures.js'

/**
 * Text formatting, geometry and alignment.
 *
 * Every assertion goes through the public read model rather than through raw
 * XML, because that model is what an agent sees: if `readSlide` disagrees with
 * what was written, the feature is broken however tidy the XML looks.
 */
describe('PPTX text formatting and geometry', () => {
  const tmpDir = tempDir('pptx_text')
  let engine
  const deck = path.join(tmpDir, 'Текст.pptx')

  /**
   * A deck with one prepared slide carrying a heading, a body and a free text
   * box — as far as this host can produce it.
   *
   * With R7 the slide is built on the content layout, so the heading and the
   * body placeholder are real and the tests that address them use it. Without
   * R7 the package has no layouts at all: the second slide is a clone of the
   * first with its inherited shape removed, which leaves the free text box as
   * the only object to format. Tests whose subject is a placeholder are gated
   * with `requiresR7` rather than reading a placeholder that is not there.
   */
  async function freshDeck(name) {
    const target = path.join(tmpDir, name)
    await engine.create(target, { overwrite: true, title: 'Исходный заголовок' })
    await addSlideAnywhere(engine, target, {
      layoutType: 'obj',
      title: 'Заголовок раздела',
      paragraphs: [
        { text: 'Первый пункт', bullet: true },
        { text: 'Второй пункт', bullet: true, level: 1 }
      ]
    })
    if (!HAS_R7_TEMPLATES) await clearSlide(engine, target, 1)
    const textBox = await engine.addTextBox(target, {
      slideIndex: 1,
      x: 500000, y: 4000000, width: 4000000, height: 900000,
      text: 'Текст в надписи',
      font: { family: 'Arial', size: 14 }
    })
    return { target, textBoxId: textBox.objectId }
  }

  before(async () => {
    engine = new PptxEngine()
    if (!R7_AVAILABLE) return
    await engine.create(deck, { overwrite: true, title: 'Текстовая презентация' })
    await engine.addSlide(deck, {
      layoutType: 'obj',
      title: 'Форматирование',
      paragraphs: ['Строка один', 'Строка два']
    })
  })

  after(() => {
    try {
      fs.rmSync(tmpDir, { recursive: true, force: true })
    } catch {
      // best effort
    }
  })

  test('the reader reports the real font family and size a title inherits', async (t) => {
    // The inherited size and family come from R7's title layout.
    if (requiresR7(t)) return
    const slide = await engine.readSlide(deck, 0)
    const title = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'ctrTitle')
    assert.ok(title, 'the title slide exposes its title placeholder')
    assert.ok(title.font.family, 'the title reports a font family')
    assert.ok(title.font.size > 0, `the title reports a real size, got ${title.font.size}`)
    assert.equal(title.positionValid, true, 'the title has usable geometry')
  })

  test('font family, size, bold, italic, underline and colour are written and read back', async (t) => {
    if (requiresR7(t)) return
    const { target } = await freshDeck('font.pptx')
    const before = await engine.readSlide(target, 1)
    const title = before.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'title')

    await engine.formatObject(target, {
      slideIndex: 1,
      objectId: title.id,
      font: {
        family: 'Georgia',
        size: 32,
        bold: true,
        italic: true,
        underline: true,
        color: '#C00000'
      }
    })

    const after = await engine.readSlide(target, 1)
    const updated = after.slide.objects.find((o) => o.id === title.id)
    assert.equal(updated.font.family, 'Georgia')
    assert.equal(updated.font.size, 32)
    assert.equal(updated.font.bold, true)
    assert.equal(updated.font.italic, true)
    assert.equal(updated.font.underline, 'sng')
    assert.equal(updated.font.color, '#C00000')
  })

  test('a font name also reaches the complex-script slot, so Cyrillic changes', async (t) => {
    if (requiresR7(t)) return
    const { target } = await freshDeck('font-cs.pptx')
    const before = await engine.readSlide(target, 1)
    const body = before.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'body')
    await engine.formatObject(target, { slideIndex: 1, objectId: body.id, font: { family: 'PT Sans' } })

    const zip = await ZipArchive.fromFile(target)
    const slideXml = zip.getText(await partPathOf(target, 1))
    assert.match(slideXml, /<a:latin typeface="PT Sans"\/>/)
    assert.match(slideXml, /<a:cs typeface="PT Sans"\/>/, 'Cyrillic must not fall back to the theme font')
  })

  test('underline accepts a style name and false removes it', async (t) => {
    if (requiresR7(t)) return
    const { target } = await freshDeck('underline.pptx')
    const before = await engine.readSlide(target, 1)
    const title = before.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'title')

    await engine.formatObject(target, { slideIndex: 1, objectId: title.id, underline: 'double' })
    let slide = await engine.readSlide(target, 1)
    assert.equal(slide.slide.objects.find((o) => o.id === title.id).font.underline, 'dbl')

    await engine.formatObject(target, { slideIndex: 1, objectId: title.id, underline: false })
    slide = await engine.readSlide(target, 1)
    assert.equal(slide.slide.objects.find((o) => o.id === title.id).font.underline, false)
  })

  test('an unsupported underline style is rejected instead of silently ignored', async (t) => {
    if (requiresR7(t)) return
    const { target } = await freshDeck('underline-bad.pptx')
    const before = await engine.readSlide(target, 1)
    const title = before.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'title')
    await assert.rejects(
      () => engine.formatObject(target, { slideIndex: 1, objectId: title.id, underline: 'squiggly' }),
      /Unsupported underline style/
    )
  })

  test('paragraph alignment is applied and read back', async (t) => {
    if (requiresR7(t)) return
    const { target } = await freshDeck('align.pptx')
    const before = await engine.readSlide(target, 1)
    const title = before.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'title')

    for (const [requested, expected] of [['center', 'ctr'], ['right', 'r'], ['justify', 'just'], ['left', 'l']]) {
      await engine.formatObject(target, { slideIndex: 1, objectId: title.id, alignment: requested })
      const slide = await engine.readSlide(target, 1)
      const object = slide.slide.objects.find((o) => o.id === title.id)
      assert.equal(object.alignment.horizontal, expected, `alignment ${requested}`)
      assert.equal(object.paragraphs[0].alignment, expected)
    }
  })

  test('vertical anchor is applied to the text body and reported', async (t) => {
    if (requiresR7(t)) return
    const { target } = await freshDeck('anchor.pptx')
    const before = await engine.readSlide(target, 1)
    const title = before.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'title')
    for (const [requested, expected] of [['bottom', 'b'], ['center', 'ctr'], ['top', 't']]) {
      await engine.formatObject(target, { slideIndex: 1, objectId: title.id, verticalAnchor: requested })
      const slide = await engine.readSlide(target, 1)
      assert.equal(slide.slide.objects.find((o) => o.id === title.id).alignment.vertical, expected)
    }
  })

  test('position and size are written in EMU and read back exactly', async () => {
    const { target, textBoxId } = await freshDeck('geometry.pptx')
    await engine.formatObject(target, {
      slideIndex: 1,
      objectId: textBoxId,
      x: 1234567,
      y: 765432,
      width: 3456789,
      height: 987654
    })

    const slide = await engine.readSlide(target, 1)
    const object = slide.slide.objects.find((o) => o.id === textBoxId)
    assert.equal(object.x, 1234567)
    assert.equal(object.y, 765432)
    assert.equal(object.width, 3456789)
    assert.equal(object.height, 987654)
    assert.equal(object.positionValid, true)
  })

  test('position and size accept unit suffixes', async () => {
    const { target, textBoxId } = await freshDeck('units.pptx')
    await engine.formatObject(target, { slideIndex: 1, objectId: textBoxId, x: '2cm', y: '1in', width: '10cm', height: '20pt' })
    const slide = await engine.readSlide(target, 1)
    const object = slide.slide.objects.find((o) => o.id === textBoxId)
    assert.equal(object.x, 720000, '2 cm')
    assert.equal(object.y, 914400, '1 inch')
    assert.equal(object.width, 3600000, '10 cm')
    assert.equal(object.height, 254000, '20 pt')
  })

  test('rotation is written in 1/60000 degrees and read back', async () => {
    const { target, textBoxId } = await freshDeck('rotation.pptx')
    await engine.formatObject(target, { slideIndex: 1, objectId: textBoxId, rotation: 45 })
    const slide = await engine.readSlide(target, 1)
    assert.equal(slide.slide.objects.find((o) => o.id === textBoxId).rotation, 45)
  })

  test('a rotated object keeps its rotation when only its size changes', async () => {
    const { target, textBoxId } = await freshDeck('rotation-keep.pptx')
    await engine.formatObject(target, { slideIndex: 1, objectId: textBoxId, rotation: 30 })
    await engine.formatObject(target, { slideIndex: 1, objectId: textBoxId, width: 2000000 })
    const slide = await engine.readSlide(target, 1)
    const object = slide.slide.objects.find((o) => o.id === textBoxId)
    assert.equal(object.rotation, 30)
    assert.equal(object.width, 2000000)
  })

  test('bulleted and numbered lists are distinguished when read back', async (t) => {
    if (requiresR7(t)) return
    const target = path.join(tmpDir, 'lists.pptx')
    await engine.create(target, { overwrite: true, title: 'Списки' })
    await engine.addSlide(target, {
      layoutType: 'obj',
      title: 'Списки',
      paragraphs: [
        { text: 'Маркированный', bullet: true },
        { text: 'Нумерованный', numbered: true },
        { text: 'Без знака', bullet: false },
        { text: 'Тире', bullet: '–' }
      ]
    })

    const slide = await engine.readSlide(target, 1)
    const body = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'body')
    assert.equal(body.paragraphs.length, 4)
    assert.equal(body.paragraphs[0].bullet, 'bullet')
    assert.equal(body.paragraphs[0].bulletCharacter, '•')
    assert.equal(body.paragraphs[1].bullet, 'numbered')
    assert.equal(body.paragraphs[1].bulletType, 'arabicPeriod')
    assert.equal(body.paragraphs[2].bullet, 'none')
    assert.equal(body.paragraphs[3].bullet, 'bullet')
    assert.equal(body.paragraphs[3].bulletCharacter, '–')
  })

  test('line spacing and paragraph spacing survive a round-trip', async (t) => {
    if (requiresR7(t)) return
    const { target } = await freshDeck('spacing.pptx')
    const before = await engine.readSlide(target, 1)
    const body = before.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'body')

    await engine.formatObject(target, {
      slideIndex: 1,
      objectId: body.id,
      lineSpacing: 1.5,
      spaceBefore: 12,
      spaceAfter: 6
    })

    const slide = await engine.readSlide(target, 1)
    const object = slide.slide.objects.find((o) => o.id === body.id)
    assert.equal(object.paragraphs[0].lineSpacing, 1.5)
    assert.equal(object.paragraphs[0].spaceBefore, 12)
    assert.equal(object.paragraphs[0].spaceAfter, 6)
  })

  test('paragraphIndex formats one paragraph and leaves its siblings alone', async (t) => {
    if (requiresR7(t)) return
    const { target } = await freshDeck('one-paragraph.pptx')
    const before = await engine.readSlide(target, 1)
    const body = before.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'body')

    await engine.formatObject(target, {
      slideIndex: 1,
      objectId: body.id,
      paragraphIndex: 1,
      alignment: 'right',
      bold: true,
      color: '#008000'
    })

    const slide = await engine.readSlide(target, 1)
    const object = slide.slide.objects.find((o) => o.id === body.id)
    assert.equal(object.paragraphs[1].alignment, 'r')
    assert.notEqual(object.paragraphs[0].alignment, 'r', 'the first paragraph must keep its alignment')
  })

  test('an out-of-range paragraphIndex fails loudly', async (t) => {
    if (requiresR7(t)) return
    const { target } = await freshDeck('one-paragraph-bad.pptx')
    const before = await engine.readSlide(target, 1)
    const body = before.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'body')
    await assert.rejects(
      () => engine.formatObject(target, { slideIndex: 1, objectId: body.id, paragraphIndex: 99, bold: true }),
      /out of range/
    )
  })

  test('mixed formatting inside one paragraph is preserved as runs', async (t) => {
    if (requiresR7(t)) return
    const target = path.join(tmpDir, 'runs.pptx')
    await engine.create(target, { overwrite: true, title: 'Прогоны' })
    await engine.addSlide(target, {
      layoutType: 'obj',
      title: 'Прогоны',
      paragraphs: [{
        runs: [
          { text: 'Жирный', bold: true, color: '#C00000' },
          { text: ' и обычный', bold: false }
        ]
      }]
    })

    const slide = await engine.readSlide(target, 1)
    const body = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'body')
    assert.equal(body.paragraphs[0].runs.length, 2)
    assert.equal(body.paragraphs[0].runs[0].text, 'Жирный')
    assert.equal(body.paragraphs[0].runs[0].bold, true)
    assert.equal(body.paragraphs[0].runs[0].color, '#C00000')
    assert.equal(body.paragraphs[0].runs[1].text, ' и обычный')
    assert.equal(body.paragraphs[0].runs[1].bold, false)
  })

  test('text replacement keeps the object and only changes its text', async () => {
    const { target, textBoxId } = await freshDeck('replace-text.pptx')
    await engine.formatObject(target, { slideIndex: 1, objectId: textBoxId, text: 'Новый текст надписи' })
    const slide = await engine.readSlide(target, 1)
    const object = slide.slide.objects.find((o) => o.id === textBoxId)
    assert.equal(object.text, 'Новый текст надписи')
    assert.equal(object.x, 500000, 'geometry must survive a text replacement')
  })

  test('editSlide still replaces text by search and reports the match count', async (t) => {
    if (requiresR7(t)) return
    const { target } = await freshDeck('search-replace.pptx')
    const result = await engine.editSlide(target, {
      slideIndex: 1,
      search: 'Первый пункт',
      replace: 'Первый пункт (изменён)'
    })
    assert.ok(result.success)
    assert.equal(result.matchesCount, 1)
    const slide = await engine.readSlide(target, 1)
    const body = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'body')
    assert.match(body.text, /изменён/)
  })

  test('editing a presentation preserves every part it did not need to change', async () => {
    const { target, textBoxId } = await freshDeck('untouched.pptx')
    const before = await ZipArchive.fromFile(target)
    const editedPart = await partPathOf(target, 1)

    await engine.formatObject(target, { slideIndex: 1, objectId: textBoxId, bold: true, color: '#123456' })

    const after = await ZipArchive.fromFile(target)
    assert.deepEqual(after.list(), before.list(), 'no part may appear or disappear')
    for (const name of before.list()) {
      if (name === editedPart) {
        assert.notEqual(
          before.getBuffer(name).toString('base64'),
          after.getBuffer(name).toString('base64'),
          'the edited slide must change'
        )
        continue
      }
      assert.equal(
        before.entries.get(name).raw.toString('base64'),
        after.entries.get(name).raw.toString('base64'),
        `"${name}" must be byte-identical after a text restyle`
      )
    }
  })

  test('the deck reopens after every edit and still validates', async () => {
    const { target, textBoxId } = await freshDeck('reopen.pptx')
    await engine.formatObject(target, { slideIndex: 1, objectId: textBoxId, size: 20, bold: true })
    const validation = await engine.validate(target)
    assert.equal(validation.valid, true, JSON.stringify(validation.errors))
    const reopening = await engine.inspect(target)
    assert.equal(reopening.slidesCount, 2)
  })

  test('the edited deck renders to PDF in R7', async (t) => {
    if (requiresR7(t)) return
    const { target, textBoxId } = await freshDeck('render.pptx')
    await engine.formatObject(target, { slideIndex: 1, objectId: textBoxId, size: 28, bold: true, color: '#1F6FEB' })
    const pdfPath = path.join(tmpDir, 'render.pdf')
    const result = await engine.toPdf(target, pdfPath)
    assert.equal(result.success, true)
    assert.ok(fs.statSync(pdfPath).size > 500, 'a real PDF was produced')
    assert.ok(fs.readFileSync(pdfPath).subarray(0, 5).toString('latin1') === '%PDF-')
  })
})

/** The slide part backing a deck position, so a test can address raw markup. */
async function partPathOf(filePath, slideIndex) {
  const zip = await ZipArchive.fromFile(filePath)
  const { slideParts } = await import('../../src/r7/pptx-util.js')
  const descriptors = slideParts(zip)
  if (!descriptors[slideIndex]) throw new Error(`no slide ${slideIndex}`)
  return descriptors[slideIndex].partPath
}
