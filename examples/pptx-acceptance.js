/**
 * Acceptance deck: a real business presentation built with the public API.
 *
 *   node examples/pptx-acceptance.js [outputDirectory]
 *
 * Builds `R7_MCP_PPTX_Acceptance.pptx` (seven slides) and renders it to
 * `R7_MCP_PPTX_Acceptance.pdf` through the R7 converter when one is installed.
 * Nothing is committed: the default location is under the OS temp directory.
 *
 * Everything here goes through the documented public surface — `create`,
 * `addSlide`, `addShape`, `addTextBox`, `addImage`, `formatObject`,
 * `duplicateSlide`, `moveSlide`, `readSlide` — because an example that needs a
 * private back door is not an example of the product.
 */

import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import zlib from 'node:zlib'
import { PptxEngine, R7Adapter } from '../src/r7/index.js'

// ------------------------------------------------------------------ palette

const INK = '#1F2933'
const MUTED = '#52606D'
const ACCENT = '#1F6FEB'
const ACCENT_DARK = '#0B3D91'
const WARM = '#E45756'
const GREEN = '#0E9F6E'
const PANEL = '#F3F6FB'
const WHITE = '#FFFFFF'

const outDir = path.resolve(process.argv[2] || path.join(os.tmpdir(), 'r7-acceptance'))
fs.mkdirSync(outDir, { recursive: true })

const deckPath = path.join(outDir, 'R7_MCP_PPTX_Acceptance.pptx')
if (fs.existsSync(deckPath)) fs.rmSync(deckPath)

const pptx = new PptxEngine()
const adapter = new R7Adapter()
const r7 = await adapter.detect()

const report = []
function step(text) {
  report.push(text)
  console.log(`  ${text}`)
}

console.log(`Acceptance deck -> ${deckPath}\n`)

// ------------------------------------------------------- 1. title slide

console.log('1. Титульный слайд')
await pptx.create(deckPath, { overwrite: true, title: 'Годовой отчёт 2026' })
{
  const slide = await pptx.readSlide(deckPath, 0)
  const title = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'ctrTitle')
  const subtitle = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'subTitle')

  await pptx.formatObject(deckPath, {
    slideIndex: 0,
    objectId: title.id,
    font: { family: 'Georgia', size: 44, bold: true, color: ACCENT_DARK },
    alignment: 'center',
    verticalAnchor: 'bottom'
  })
  if (subtitle) {
    await pptx.formatObject(deckPath, {
      slideIndex: 0,
      objectId: subtitle.id,
      text: 'Итоги года, ключевые показатели и планы на следующий период',
      font: { family: 'Arial', size: 18, color: MUTED },
      alignment: 'center',
      verticalAnchor: 'top'
    })
  }
  // A rule under the title, drawn as a shape rather than set as a border, so
  // the deck stays editable in any editor.
  await pptx.addShape(deckPath, {
    slideIndex: 0,
    shape: 'rectangle',
    name: 'Линейка',
    x: '5.9cm', y: '10.4cm', width: '2.8cm', height: '4pt',
    fill: ACCENT, noLine: true
  })
  step(`титул: "${title.text}" + подзаголовок, шрифты Georgia/Arial, акцент ${ACCENT_DARK}`)
}

// ---------------------------------------------- 2. heading + body text

console.log('2. Заголовок и текст')
await pptx.addSlide(deckPath, {
  layoutType: 'obj',
  title: 'Ключевые выводы',
  paragraphs: [
    { text: 'Выручка выросла на 18% и превысила план на 4 процентных пункта.', bullet: true },
    { text: 'Маржинальность удержана на уровне 42% при росте фонда оплаты труда.', bullet: true },
    { text: 'Отток клиентов снизился второй квартал подряд.', bullet: true },
    { text: 'Основной риск следующего года — стоимость логистики.', bullet: true, color: WARM }
  ]
})
{
  const slide = await pptx.readSlide(deckPath, 1)
  const title = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'title')
  const body = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'body')
  await pptx.formatObject(deckPath, {
    slideIndex: 1, objectId: title.id,
    font: { family: 'Arial', size: 30, bold: true, color: INK }, alignment: 'left'
  })
  await pptx.formatObject(deckPath, {
    slideIndex: 1, objectId: body.id,
    font: { family: 'Arial', size: 16, color: INK },
    lineSpacing: 1.25, spaceAfter: 10
  })
  step('заголовок + 4 пункта списка, межстрочный интервал 1.25, интервал после абзаца 10 pt')
}

// ------------------------------------------------------- 3. two columns

console.log('3. Две колонки')
await pptx.addSlide(deckPath, { layoutType: 'twoObj', title: 'Что сработало и что нет' })
{
  const slide = await pptx.readSlide(deckPath, 2)
  const title = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'title')
  await pptx.formatObject(deckPath, {
    slideIndex: 2, objectId: title.id, font: { family: 'Arial', size: 30, bold: true, color: INK }
  })

  const columns = slide.slide.objects.filter(
    (o) => o.onSlide && o.placeholder && (o.placeholder.type === 'body' || o.placeholder.type === 'obj')
  )
  const left = columns[0]
  const right = columns[1]
  if (!left || !right) throw new Error('the twoObj layout did not give two content placeholders')

  await pptx.formatObject(deckPath, {
    slideIndex: 2, objectId: left.id,
    paragraphs: [
      { text: 'Сработало', size: 20, bold: true, color: GREEN, bullet: false, spaceAfter: 8 },
      { text: 'Запуск нового направления', bullet: true },
      { text: 'Автоматизация отчётности', bullet: true },
      { text: 'Программа удержания', bullet: true }
    ],
    font: { family: 'Arial', size: 15, color: INK },
    level: 0
  })
  await pptx.formatObject(deckPath, {
    slideIndex: 2, objectId: right.id,
    paragraphs: [
      { text: 'Не сработало', size: 20, bold: true, color: WARM, bullet: false, spaceAfter: 8 },
      { text: 'Пилот на восточном рынке', bullet: true },
      { text: 'Единая CRM для всех линий', bullet: true }
    ],
    font: { family: 'Arial', size: 15, color: INK }
  })
  step('два независимых текстовых placeholder слева и справа, разные цвета заголовков колонок')
}

// ------------------------------------------------------ 4. bulleted list

console.log('4. Нумерованный и маркированный список')
await pptx.addSlide(deckPath, { layoutType: 'obj', title: 'План на следующий год' })
{
  const slide = await pptx.readSlide(deckPath, 3)
  const title = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'title')
  const body = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'body')
  await pptx.formatObject(deckPath, {
    slideIndex: 3, objectId: title.id, font: { family: 'Arial', size: 30, bold: true, color: INK }
  })
  await pptx.formatObject(deckPath, {
    slideIndex: 3, objectId: body.id,
    paragraphs: [
      { text: 'Выйти на рынок СНГ', numbered: true },
      { text: 'Повысить маржинальность до 45%', numbered: true },
      { text: 'Сократить стоимость логистики', numbered: true },
      { text: 'Поддерживающие инициативы', bullet: '–', level: 1, color: MUTED },
      { text: 'Единый портал самообслуживания', bullet: '•', level: 2 },
      { text: 'Обучение партнёров', bullet: '•', level: 2 }
    ],
    font: { family: 'Arial', size: 15, color: INK },
    lineSpacing: 1.15
  })
  step('нумерованный список 1-3, тире-маркер уровня 1, точки уровня 2')
}

// -------------------------------------------------------- 5. KPI cards

console.log('5. Карточки KPI из фигур')
await pptx.addSlide(deckPath, { layoutType: 'titleOnly', title: 'Показатели года' })
{
  const slide = await pptx.readSlide(deckPath, 4)
  const title = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'title')
  await pptx.formatObject(deckPath, {
    slideIndex: 4, objectId: title.id, font: { family: 'Arial', size: 30, bold: true, color: INK }
  })

  const cards = [
    { label: 'Выручка', value: '1,24 млрд ₽', delta: '+18%', colour: ACCENT },
    { label: 'Маржа', value: '42%', delta: '+1,5 п.п.', colour: GREEN },
    { label: 'Клиенты', value: '8 420', delta: '+640', colour: '#6B4EFF' }
  ]

  // A panel behind the cards, inserted at z-order 0 so it stays behind them.
  await pptx.addShape(deckPath, {
    slideIndex: 4,
    shape: 'rounded-rectangle',
    name: 'Подложка',
    zOrder: 0,
    x: '1.4cm', y: '4.2cm', width: '30.6cm', height: '9.4cm',
    fill: PANEL, noLine: true
  })

  for (const [index, card] of cards.entries()) {
    const x = `${1.4 + index * 10.6}cm`
    await pptx.addShape(deckPath, {
      slideIndex: 4,
      shape: 'rounded-rectangle',
      name: `Карточка ${card.label}`,
      x, y: '4.8cm', width: '9.4cm', height: '8.2cm',
      fill: WHITE, line: '#D2DAE6', lineWidth: 1
    })
    // A colour bar makes each card identifiable at a glance.
    await pptx.addShape(deckPath, {
      slideIndex: 4,
      shape: 'rectangle',
      name: `Полоса ${card.label}`,
      x, y: '4.8cm', width: '9.4cm', height: '0.35cm',
      fill: card.colour, noLine: true
    })
    await pptx.addTextBox(deckPath, {
      slideIndex: 4,
      name: `Подпись ${card.label}`,
      x: `${2.0 + index * 10.6}cm`, y: '5.5cm', width: '8.2cm', height: '1.0cm',
      text: card.label,
      font: { family: 'Arial', size: 13, color: MUTED },
      alignment: 'center', verticalAnchor: 'middle',
      fill: null, noFill: true, noLine: true
    })
    await pptx.addTextBox(deckPath, {
      slideIndex: 4,
      name: `Значение ${card.label}`,
      x: `${2.0 + index * 10.6}cm`, y: '6.6cm', width: '8.2cm', height: '2.4cm',
      text: card.value,
      font: { family: 'Georgia', size: 30, bold: true, color: INK },
      alignment: 'center', verticalAnchor: 'middle',
      fill: null, noFill: true, noLine: true
    })
    await pptx.addTextBox(deckPath, {
      slideIndex: 4,
      name: `Динамика ${card.label}`,
      x: `${2.0 + index * 10.6}cm`, y: '9.4cm', width: '8.2cm', height: '1.2cm',
      text: card.delta,
      font: { family: 'Arial', size: 14, bold: true, color: card.colour },
      alignment: 'center', verticalAnchor: 'middle',
      fill: null, noFill: true, noLine: true
    })
  }

  // A trend line across the panel, with an arrow head.
  await pptx.addShape(deckPath, {
    slideIndex: 4,
    shape: 'line',
    name: 'Тренд',
    x: '2.4cm', y: '14.2cm', width: '28.6cm', height: 0,
    line: ACCENT, lineWidth: 1.5, arrows: { tail: 'triangle' }
  })
  step('подложка + 3 карточки (rect/roundRect), цветные полосы, 9 надписей, линия со стрелкой')
}

// ------------------------------------------------- 6. image + caption

console.log('6. Изображение с подписью')
{
  // The picture is generated here rather than shipped with the repository: no
  // binary fixture, no third-party asset, and the deck stays self-contained.
  const pngPath = path.join(outDir, 'acceptance-chart.png')
  fs.writeFileSync(pngPath, makeChartPng(560, 320))

  await pptx.addSlide(deckPath, { layoutType: 'titleOnly', title: 'Динамика выручки по кварталам' })
  const slide = await pptx.readSlide(deckPath, 5)
  const title = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'title')
  await pptx.formatObject(deckPath, {
    slideIndex: 5, objectId: title.id, font: { family: 'Arial', size: 30, bold: true, color: INK }
  })

  const image = await pptx.addImage(deckPath, {
    slideIndex: 5,
    imagePath: pngPath,
    name: 'Диаграмма выручки',
    x: '4.0cm', y: '4.2cm', width: '24.0cm',
    description: 'Столбчатая диаграмма выручки по кварталам 2026 года'
  })
  await pptx.addTextBox(deckPath, {
    slideIndex: 5,
    name: 'Подпись к рисунку',
    x: '4.0cm', y: '15.4cm', width: '24.0cm', height: '1.4cm',
    text: 'Рис. 1. Выручка по кварталам, млн ₽ (данные внутренней отчётности)',
    font: { family: 'Arial', size: 12, italic: true, color: MUTED },
    alignment: 'center', verticalAnchor: 'middle',
    fill: null, noFill: true, noLine: true
  })
  step(`изображение вставлено (${image.image.mediaPath}, ${image.image.naturalWidth}x${image.image.naturalHeight} px) + подпись`)

  // Exercise image replacement on the same object, then put the original back.
  await pptx.formatObject(deckPath, {
    slideIndex: 5, objectId: image.objectId, imagePath: pngPath
  })
  step('изображение заменено через replaceImage, геометрия сохранена')
}

// ----------------------------------------------------------- 7. closing

console.log('7. Закрывающий слайд')
await pptx.addSlide(deckPath, { layoutType: 'title', title: 'Спасибо за внимание', subtitle: 'Вопросы и обсуждение' })
{
  const slide = await pptx.readSlide(deckPath, 6)
  const title = slide.slide.objects.find((o) => o.placeholder && (o.placeholder.type === 'ctrTitle' || o.placeholder.type === 'title'))
  const subtitle = slide.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'subTitle')
  await pptx.formatObject(deckPath, {
    slideIndex: 6, objectId: title.id,
    font: { family: 'Georgia', size: 40, bold: true, color: ACCENT_DARK }, alignment: 'center'
  })
  if (subtitle) {
    await pptx.formatObject(deckPath, {
      slideIndex: 6, objectId: subtitle.id,
      font: { family: 'Arial', size: 16, color: MUTED }, alignment: 'center'
    })
  }
  // A contact strip built from shapes, to prove shapes carry text well.
  await pptx.addShape(deckPath, {
    slideIndex: 6,
    shape: 'rounded-rectangle',
    name: 'Контакты',
    x: '9.0cm', y: '12.4cm', width: '14.2cm', height: '1.8cm',
    fill: ACCENT, noLine: true,
    text: 'finance@example.ru  ·  +7 495 000-00-00',
    font: { family: 'Arial', size: 14, bold: true, color: WHITE },
    alignment: 'center', verticalAnchor: 'middle'
  })
  step('закрывающий слайд на титульном макете + плашка с контактами')
}

// -------------------------------------------- 8. editing an existing object

console.log('8. Правка существующего объекта и порядок слайдов')
{
  const before = await pptx.readSlide(deckPath, 1)
  const body = before.slide.objects.find((o) => o.placeholder && o.placeholder.type === 'body')
  await pptx.formatObject(deckPath, {
    slideIndex: 1,
    objectId: body.id,
    paragraphIndex: 3,
    font: { size: 15, bold: true, color: WARM }
  })
  step('правка одного абзаца (индекс 3) без пересоздания слайда')

  // Duplicate and move, then read the deck back to prove the order.
  const copy = await pptx.duplicateSlide(deckPath, 1)
  await pptx.moveSlide(deckPath, copy.slideIndex, 1)
  const reordered = await pptx.readSlide(deckPath, 0)
  await pptx.deleteSlide(deckPath, 1)
  step(`дублирование + перемещение + удаление копии, слайдов осталось ${reordered.slideCount - 1}`)
}

// -------------------------------------------------------------- summary

const finalSlide = await pptx.readSlide(deckPath, 0)
const validation = await pptx.validate(deckPath)
const inspection = await pptx.inspect(deckPath)
const layouts = await pptx.listLayouts(deckPath)

console.log('\n--- Итог -------------------------------------------------')
console.log(`  слайдов:            ${finalSlide.slideCount}`)
console.log(`  макетов:            ${layouts.layouts.length}`)
console.log(`  размер слайда:      ${inspection.slideSize.width}x${inspection.slideSize.height} EMU`)
console.log(`  валидация:          ${validation.valid ? 'ok' : `ОШИБКИ: ${validation.errors.join('; ')}`}`)
console.log(`  размер файла:       ${fs.statSync(deckPath).size} байт`)

for (let i = 0; i < finalSlide.slideCount; i++) {
  const slide = await pptx.readSlide(deckPath, i)
  const texts = slide.slide.objects
    .filter((o) => o.onSlide && o.text && o.text.trim() !== '')
    .map((o) => o.text.split('\n')[0])
  const shapes = slide.slide.objects.filter((o) => o.onSlide).length
  console.log(`  ${i + 1}. [${slide.slide.layout.type}] ${texts[0] || '(без текста)'} — объектов: ${shapes}`)
}

if (!validation.valid) {
  console.error('\nDeck is invalid; not rendering.')
  process.exit(1)
}

// ------------------------------------------------------------------ PDF

const pdfPath = path.join(outDir, 'R7_MCP_PPTX_Acceptance.pdf')
if (r7.installed) {
  console.log('\nРендер в PDF через конвертер R7...')
  try {
    const converted = await pptx.toPdf(deckPath, pdfPath)
    const head = fs.readFileSync(pdfPath).subarray(0, 5).toString('latin1')
    if (!converted.success || head !== '%PDF-') {
      throw new Error(`the converter produced something that is not a PDF (starts with "${head}")`)
    }
    console.log(`  PDF: ${pdfPath} (${fs.statSync(pdfPath).size} байт, ${converted.timeMs} ms)`)
  } catch (err) {
    console.error(`  Рендер не удался: ${err.message}`)
    process.exitCode = 1
  }
} else {
  console.log('\nR7-Office не найден: PDF не создавался (x2t недоступен на этом хосте).')
}

console.log('\n=== Пути ===')
console.log(`PPTX: ${deckPath} (${fs.statSync(deckPath).size} байт)`)
if (fs.existsSync(pdfPath)) {
  console.log(`PDF : ${pdfPath} (${fs.statSync(pdfPath).size} байт)`)
}

// -------------------------------------------------------------- PNG maker

function crc32(buffer) {
  let crc = -1
  for (let i = 0; i < buffer.length; i++) {
    crc ^= buffer[i]
    for (let k = 0; k < 8; k++) crc = crc & 1 ? (crc >>> 1) ^ 0xedb88320 : crc >>> 1
  }
  return (crc ^ -1) >>> 0
}

function chunk(type, data) {
  const length = Buffer.alloc(4)
  length.writeUInt32BE(data.length)
  const typeBuffer = Buffer.from(type, 'latin1')
  const crc = Buffer.alloc(4)
  crc.writeUInt32BE(crc32(Buffer.concat([typeBuffer, data])))
  return Buffer.concat([length, typeBuffer, data, crc])
}

/**
 * A small bar chart, drawn pixel by pixel.
 *
 * The runtime has zero dependencies by design, so there is no canvas to draw
 * on; a real PNG is written out directly.
 */
function makeChartPng(width, height) {
  const raw = Buffer.alloc(height * (1 + width * 3))
  const bars = [
    { value: 0.55, colour: [76, 120, 168] },
    { value: 0.72, colour: [31, 111, 235] },
    { value: 0.86, colour: [14, 159, 110] },
    { value: 1.00, colour: [107, 78, 255] }
  ]
  const plotLeft = Math.round(width * 0.10)
  const plotRight = Math.round(width * 0.94)
  const plotTop = Math.round(height * 0.10)
  const plotBottom = Math.round(height * 0.84)
  const plotWidth = plotRight - plotLeft
  const plotHeight = plotBottom - plotTop
  const slot = plotWidth / bars.length
  const barWidth = Math.round(slot * 0.52)

  const set = (x, y, [r, g, b]) => {
    if (x < 0 || y < 0 || x >= width || y >= height) return
    const offset = y * (1 + width * 3) + 1 + x * 3
    raw[offset] = r
    raw[offset + 1] = g
    raw[offset + 2] = b
  }

  // Background and frame.
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) set(x, y, [255, 255, 255])
  }
  for (let x = plotLeft; x <= plotRight; x++) set(x, plotBottom, [210, 218, 230])
  for (let y = plotTop; y <= plotBottom; y++) set(plotLeft, y, [210, 218, 230])

  // Grid lines.
  for (let step = 1; step <= 4; step++) {
    const y = plotBottom - Math.round((plotHeight * step) / 4)
    for (let x = plotLeft + 1; x < plotRight; x++) set(x, y, [236, 241, 247])
  }

  bars.forEach((bar, index) => {
    const barHeight = Math.round(plotHeight * bar.value)
    const left = Math.round(plotLeft + slot * index + (slot - barWidth) / 2)
    for (let y = plotBottom - barHeight; y < plotBottom; y++) {
      for (let x = left; x < left + barWidth; x++) set(x, y, bar.colour)
    }
    // A lighter cap, so the bars read as deliberate shapes.
    for (let x = left; x < left + barWidth; x++) set(x, plotBottom - barHeight, [255, 255, 255])
  })

  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(width, 0)
  ihdr.writeUInt32BE(height, 4)
  ihdr[8] = 8
  ihdr[9] = 2
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', zlib.deflateSync(raw)),
    chunk('IEND', Buffer.alloc(0))
  ])
}
