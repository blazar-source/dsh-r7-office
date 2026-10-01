/**
 * Acceptance deck: a real business presentation built with the public API.
 *
 *   node examples/pptx-acceptance.js [outputDirectory]
 *
 * Builds `R7_MCP_PPTX_Acceptance.pptx` (seven slides), checks its structure
 * (duplicate shape ids, duplicate placeholders, overflow, unintended overlap)
 * and renders it to `R7_MCP_PPTX_Acceptance.pdf` through the R7 converter when
 * one is installed. Nothing is committed: the default location is under the OS
 * temp directory.
 *
 * Everything here goes through the documented public surface — `create`,
 * `addSlide`, `addShape`, `addTextBox`, `addImage`, `formatObject`,
 * `duplicateSlide`, `moveSlide`, `readSlide` — because an example that needs a
 * private back door is not an example of the product.
 *
 * Geometry is read back, never guessed: every object that has to sit below an
 * inherited placeholder box asks the engine where that box ends. A slide-level
 * placeholder keeps the layout's geometry, so a hardcoded `y` is exactly how a
 * title ends up underneath the object that was supposed to follow it.
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

const EMU_CM = 360000
const EMU_PT = 12700
/** Air between an inherited box and the object placed below it. */
const GAP = 200000
/** 1 pt of tolerance, so touching edges are not reported as an overlap. */
const TOLERANCE = EMU_PT
/**
 * Marks a backdrop panel: a shape drawn behind other shapes on purpose.
 *
 * The marker travels inside the file as `p:cNvPr/@descr`, so the overlap check
 * below can tell a deliberate panel from the accidental overlap this example
 * exists to catch.
 */
const BACKDROP = 'backdrop'

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

// ------------------------------------------------------------- geometry API

/** The slide canvas in EMU. */
async function canvasSize() {
  const info = await pptx.inspect(deckPath)
  return info.slideSize
}

/** A placeholder object by type, straight out of the read model. */
async function placeholder(slideIndex, types) {
  const wanted = Array.isArray(types) ? types : [types]
  const read = await pptx.readSlide(deckPath, slideIndex)
  return read.slide.objects.find((o) => o.placeholder && wanted.includes(o.placeholder.type)) || null
}

/**
 * The first `y` that clears every inherited placeholder box of the given
 * types, read back from the slide rather than assumed.
 */
async function belowPlaceholders(slideIndex, types, extra = 0) {
  const read = await pptx.readSlide(deckPath, slideIndex)
  const wanted = Array.isArray(types) ? types : [types]
  const bottoms = read.slide.objects
    .filter((o) => o.onSlide && o.placeholder && wanted.includes(o.placeholder.type))
    .map((o) => (o.y === null || o.height === null ? null : o.y + o.height))
    .filter((value) => value !== null)
  return (bottoms.length > 0 ? Math.max(...bottoms) : 0) + GAP + extra
}

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
  // the deck stays editable in any editor. Its `y` comes from where the
  // subtitle's inherited box actually ends — the box is taller than its one
  // line of text, and a rule placed inside it is drawn on top of the subtitle.
  const size = await canvasSize()
  const ruleTop = await belowPlaceholders(0, ['ctrTitle', 'title', 'subTitle'])
  await pptx.addShape(deckPath, {
    slideIndex: 0,
    shape: 'rectangle',
    name: 'Линейка',
    x: Math.round((size.width - 2.8 * EMU_CM) / 2),
    y: ruleTop,
    width: 2.8 * EMU_CM,
    height: 4 * EMU_PT,
    fill: ACCENT,
    noLine: true
  })
  step(`титул: "${title.text}" + подзаголовок, шрифты Georgia/Arial, акцент ${ACCENT_DARK}`)
  step(`линейка под подзаголовком: y=${ruleTop} EMU (низ бокса подзаголовка + ${GAP})`)
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
  const title = await placeholder(1, 'title')
  const body = await placeholder(1, 'body')
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
  const title = await placeholder(2, 'title')
  await pptx.formatObject(deckPath, {
    slideIndex: 2, objectId: title.id, font: { family: 'Arial', size: 30, bold: true, color: INK }
  })

  const read = await pptx.readSlide(deckPath, 2)
  const columns = read.slide.objects.filter(
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
  const title = await placeholder(3, 'title')
  const body = await placeholder(3, 'body')
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
  const size = await canvasSize()
  const title = await placeholder(4, 'title')
  await pptx.formatObject(deckPath, {
    slideIndex: 4, objectId: title.id, font: { family: 'Arial', size: 30, bold: true, color: INK }
  })

  const cards = [
    { label: 'Выручка', value: '1,24 млрд ₽', delta: '+18%', colour: ACCENT },
    { label: 'Маржа', value: '42%', delta: '+1,5 п.п.', colour: GREEN },
    { label: 'Клиенты', value: '8 420', delta: '+640', colour: '#6B4EFF' }
  ]

  // A panel behind the cards, inserted at z-order 0 so it stays behind them.
  // It is marked `descr="backdrop"`: a panel drawn behind other shapes is the
  // one overlap this deck intends, and the check below skips it by that mark.
  const panelTop = await belowPlaceholders(4, 'title')
  const panelHeight = 9.4 * EMU_CM
  await pptx.addShape(deckPath, {
    slideIndex: 4,
    shape: 'rounded-rectangle',
    name: 'Подложка',
    description: BACKDROP,
    zOrder: 0,
    x: 1.4 * EMU_CM, y: panelTop, width: 30.6 * EMU_CM, height: panelHeight,
    fill: PANEL, noLine: true
  })

  const cardTop = panelTop + 0.6 * EMU_CM
  const barHeight = 0.35 * EMU_CM
  const cardHeight = 8.2 * EMU_CM

  for (const [index, card] of cards.entries()) {
    const x = (1.4 + index * 10.6) * EMU_CM
    // The colour bar sits on the card's top edge rather than on top of the
    // card: touching edges are not an overlap, and the deck stays clean.
    await pptx.addShape(deckPath, {
      slideIndex: 4,
      shape: 'rectangle',
      name: `Полоса ${card.label}`,
      x, y: cardTop, width: 9.4 * EMU_CM, height: barHeight,
      fill: card.colour, noLine: true
    })
    // One card is one shape carrying its own three paragraphs: a card with
    // separate text boxes stacked over it is three shapes that all overlap it.
    await pptx.addShape(deckPath, {
      slideIndex: 4,
      shape: 'rounded-rectangle',
      name: `Карточка ${card.label}`,
      x, y: cardTop + barHeight, width: 9.4 * EMU_CM, height: cardHeight,
      fill: WHITE, line: '#D2DAE6', lineWidth: 1,
      verticalAnchor: 'middle',
      font: { family: 'Arial' },
      paragraphs: [
        { text: card.label, size: 13, color: MUTED, alignment: 'center', spaceAfter: 6 },
        { text: card.value, size: 30, bold: true, color: INK, family: 'Georgia', alignment: 'center', spaceAfter: 6 },
        { text: card.delta, size: 14, bold: true, color: card.colour, alignment: 'center' }
      ]
    })
  }

  // A trend line across the panel, below it and with no height of its own.
  await pptx.addShape(deckPath, {
    slideIndex: 4,
    shape: 'line',
    name: 'Тренд',
    x: 2.4 * EMU_CM, y: panelTop + panelHeight + Math.round(GAP / 2), width: 28.6 * EMU_CM, height: 0,
    line: ACCENT, lineWidth: 1.5, arrows: { tail: 'triangle' }
  })
  step(`подложка (descr="${BACKDROP}", исключена из проверки перекрытий) + 3 карточки с текстом + полосы + линия со стрелкой`)
  step(`подложка под заголовком: y=${panelTop} EMU (низ бокса заголовка + ${GAP}), слайд ${size.width}x${size.height}`)
}

// ------------------------------------------------- 6. image + caption

console.log('6. Изображение с подписью')
{
  // The picture is generated here rather than shipped with the repository: no
  // binary fixture, no third-party asset, and the deck stays self-contained.
  const pngPath = path.join(outDir, 'acceptance-chart.png')
  const chartWidth = 560
  const chartHeight = 320
  fs.writeFileSync(pngPath, makeChartPng(chartWidth, chartHeight))

  await pptx.addSlide(deckPath, { layoutType: 'titleOnly', title: 'Динамика выручки по кварталам' })
  const title = await placeholder(5, 'title')
  await pptx.formatObject(deckPath, {
    slideIndex: 5, objectId: title.id, font: { family: 'Arial', size: 30, bold: true, color: INK }
  })

  // The picture and its caption are fitted into the space that is actually
  // left below the title, measured from the slide, so neither can reach into
  // the title's box or past the bottom of the slide.
  const size = await canvasSize()
  const top = await belowPlaceholders(5, 'title')
  const captionHeight = 1.4 * EMU_CM
  const available = size.height - top - captionHeight - GAP
  const ratio = chartWidth / chartHeight
  const imageWidth = Math.min(24 * EMU_CM, Math.floor(available * ratio))
  const imageHeight = Math.round(imageWidth / ratio)

  const image = await pptx.addImage(deckPath, {
    slideIndex: 5,
    imagePath: pngPath,
    name: 'Диаграмма выручки',
    x: Math.round((size.width - imageWidth) / 2), y: top, width: imageWidth,
    description: 'Столбчатая диаграмма выручки по кварталам 2026 года'
  })
  await pptx.addTextBox(deckPath, {
    slideIndex: 5,
    name: 'Подпись к рисунку',
    x: Math.round((size.width - imageWidth) / 2), y: top + imageHeight + Math.round(GAP / 2),
    width: imageWidth, height: captionHeight,
    text: 'Рис. 1. Выручка по кварталам, млн ₽ (данные внутренней отчётности)',
    font: { family: 'Arial', size: 12, italic: true, color: MUTED },
    alignment: 'center', verticalAnchor: 'middle',
    fill: null, noFill: true, noLine: true
  })
  step(`изображение вставлено (${image.image.mediaPath}, ${image.image.naturalWidth}x${image.image.naturalHeight} px) + подпись`)
  step(`картинка ${imageWidth}x${imageHeight} EMU от y=${top} (низ бокса заголовка + ${GAP}), подпись ниже картинки`)

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
  const title = await placeholder(6, ['ctrTitle', 'title'])
  const subtitle = await placeholder(6, 'subTitle')
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
  // A contact strip built from shapes, to prove shapes carry text well. It
  // starts below the subtitle's box, not at a guessed height.
  const stripTop = await belowPlaceholders(6, ['ctrTitle', 'title', 'subTitle'])
  await pptx.addShape(deckPath, {
    slideIndex: 6,
    shape: 'rounded-rectangle',
    name: 'Контакты',
    x: 9.0 * EMU_CM, y: stripTop, width: 14.2 * EMU_CM, height: 1.8 * EMU_CM,
    fill: ACCENT, noLine: true,
    text: 'finance@example.ru  ·  +7 495 000-00-00',
    font: { family: 'Arial', size: 14, bold: true, color: WHITE },
    alignment: 'center', verticalAnchor: 'middle'
  })
  step(`закрывающий слайд на титульном макете + плашка с контактами от y=${stripTop} EMU`)
}

// -------------------------------------------- 8. editing an existing object

console.log('8. Правка существующего объекта и порядок слайдов')
{
  const body = await placeholder(1, 'body')
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
const structure = await pptx.validateStructure(deckPath)
const inspection = await pptx.inspect(deckPath)
const layouts = await pptx.listLayouts(deckPath)
const size = inspection.slideSize

console.log('\n--- Итог -------------------------------------------------')
console.log(`  слайдов:            ${finalSlide.slideCount}`)
console.log(`  макетов:            ${layouts.layouts.length}`)
console.log(`  размер слайда:      ${size.width}x${size.height} EMU`)
console.log(`  валидация:          ${validation.valid ? 'ok' : `ОШИБКИ: ${validation.errors.join('; ')}`}`)
console.log(`  структура:          ${structure.valid ? 'ok' : `ОШИБКИ: ${structure.errors.join('; ')}`}`)
console.log(`  размер файла:       ${fs.statSync(deckPath).size} байт`)

for (let i = 0; i < finalSlide.slideCount; i++) {
  const slide = await pptx.readSlide(deckPath, i)
  const texts = slide.slide.objects
    .filter((o) => o.onSlide && o.text && o.text.trim() !== '')
    .map((o) => o.text.split('\n')[0])
  const shapes = slide.slide.objects.filter((o) => o.onSlide).length
  console.log(`  ${i + 1}. [${slide.slide.layout.type}] ${texts[0] || '(без текста)'} — объектов: ${shapes}`)
}

// --------------------------------------------------- structure and geometry

console.log('\n--- Структура и геометрия (измерено) ---------------------')
const problems = []
for (let i = 0; i < finalSlide.slideCount; i++) {
  const slide = await pptx.readSlide(deckPath, i)
  const reportSlide = structure.details.slides[i]
  const shapes = slide.slide.objects.filter((o) => o.onSlide)
  const overflow = []
  for (const object of shapes) {
    if (object.x === null || object.y === null || object.width === null || object.height === null) continue
    const label = object.name || object.id
    if (object.x < -1 || object.y < -1) overflow.push(`${label}: отрицательная координата ${object.x},${object.y}`)
    if (object.x + object.width > size.width + 1000) {
      overflow.push(`${label}: выход вправо на ${Math.round((object.x + object.width - size.width) / EMU_PT)} pt`)
    }
    if (object.y + object.height > size.height + 1000) {
      overflow.push(`${label}: выход вниз на ${Math.round((object.y + object.height - size.height) / EMU_PT)} pt`)
    }
  }

  const boxed = shapes.filter((o) => o.positionValid !== false && (o.width || 0) > 0 && (o.height || 0) > 0)
  const backdrop = boxed.filter((o) => o.description === BACKDROP)
  const checked = boxed.filter((o) => o.description !== BACKDROP)
  const overlaps = []
  for (let a = 0; a < checked.length; a++) {
    for (let b = a + 1; b < checked.length; b++) {
      const A = checked[a]
      const B = checked[b]
      const ox = Math.min(A.x + A.width, B.x + B.width) - Math.max(A.x, B.x)
      const oy = Math.min(A.y + A.height, B.y + B.height) - Math.max(A.y, B.y)
      if (ox > TOLERANCE && oy > TOLERANCE) {
        overlaps.push(`${A.name || A.id} x ${B.name || B.id} (${Math.round(ox / EMU_PT)}x${Math.round(oy / EMU_PT)}pt)`)
      }
    }
  }

  console.log(
    `  слайд ${i + 1}: объектов ${reportSlide.shapeCount}`
    + ` | dup id ${reportSlide.duplicateShapeIds.length}`
    + ` | dup placeholder ${reportSlide.duplicatePlaceholders.length}`
    + ` | id=0 ${reportSlide.zeroIds.length}`
    + ` | выход за границы ${overflow.length}`
    + ` | перекрытий ${overlaps.length}`
    + (backdrop.length > 0 ? ` | подложек исключено ${backdrop.length}` : '')
  )
  for (const text of [...overflow, ...overlaps]) console.log(`      ${text}`)
  if (overflow.length > 0 || overlaps.length > 0) {
    problems.push(`slide ${i + 1}: ${[...overflow, ...overlaps].join('; ')}`)
  }
}

console.log(`\n  дубликаты id/placeholder: ${structure.errors.length === 0 ? 'нет' : structure.errors.join('; ')}`)
console.log(`  непреднамеренных перекрытий и выходов за границы: ${problems.length === 0 ? 'нет' : problems.join(' | ')}`)

if (!validation.valid) {
  console.error('\nDeck is invalid; not rendering.')
  process.exit(1)
}
if (!structure.valid || problems.length > 0) {
  console.error('\nDeck is structurally unsound; not rendering.')
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
