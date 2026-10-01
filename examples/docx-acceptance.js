/**
 * Acceptance example: build a real business document with the public DOCX API.
 *
 * The point of this script is that it uses nothing but `DocxEngine` and
 * `R7Adapter` — no private helpers, no hand-written XML — and still produces a
 * 4–6 page document that exercises every capability the engine claims:
 *
 *   title block · headings 1–3 · body text · bulleted list · numbered list ·
 *   table with a shaded header row · chart image · hyperlink · header/footer
 *   with a live page-number field · varied indents and spacing · a page break ·
 *   two sections (portrait, then landscape) · R7 PDF rendering
 *
 * Run:            node examples/docx-acceptance.js
 * Custom output:  node examples/docx-acceptance.js D:\out\Acceptance.docx
 *
 * The generated `.docx` and `.pdf` are never committed: the default output
 * directory is inside the system temp directory.
 */

import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import zlib from 'node:zlib'
import { DocxEngine, R7Adapter } from '../src/r7/index.js'

/* ------------------------------------------------------------------ *
 * A tiny PNG encoder, so the example ships no binary asset.
 * ------------------------------------------------------------------ */

const CRC_TABLE = (() => {
  const table = new Int32Array(256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    table[n] = c
  }
  return table
})()

function crc32(buffer) {
  let crc = -1
  for (let i = 0; i < buffer.length; i++) crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ buffer[i]) & 0xff]
  return (crc ^ -1) >>> 0
}

function pngChunk(type, data) {
  const length = Buffer.alloc(4)
  length.writeUInt32BE(data.length)
  const body = Buffer.concat([Buffer.from(type, 'latin1'), data])
  const crc = Buffer.alloc(4)
  crc.writeUInt32BE(crc32(body))
  return Buffer.concat([length, body, crc])
}

/** A real 8-bit truecolour PNG of `width` x `height` pixels drawn by `paint`. */
function makePng(width, height, paint) {
  const raw = Buffer.alloc(height * (1 + width * 3))
  for (let y = 0; y < height; y++) {
    const rowStart = y * (1 + width * 3)
    raw[rowStart] = 0
    for (let x = 0; x < width; x++) {
      const [r, g, b] = paint(x, y)
      const at = rowStart + 1 + x * 3
      raw[at] = r
      raw[at + 1] = g
      raw[at + 2] = b
    }
  }

  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(width, 0)
  ihdr.writeUInt32BE(height, 4)
  ihdr[8] = 8
  ihdr[9] = 2

  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', zlib.deflateSync(raw)),
    pngChunk('IEND', Buffer.alloc(0))
  ])
}

/** A simple bar chart: three bars on a light background with a baseline. */
function renderChart(width = 480, height = 260) {
  const bars = [
    { value: 0.55, color: [31, 111, 235] },
    { value: 0.78, color: [17, 148, 84] },
    { value: 1.0, color: [200, 80, 40] }
  ]
  const chartBottom = height - 30
  const chartTop = 24
  const slotWidth = Math.floor(width / bars.length)

  return makePng(width, height, (x, y) => {
    if (y > chartBottom) return [120, 120, 120]
    if (y > chartBottom - 2) return [60, 60, 60]

    for (let i = 0; i < bars.length; i++) {
      const barHeight = Math.round((chartBottom - chartTop) * bars[i].value)
      const left = i * slotWidth + Math.floor(slotWidth * 0.25)
      const right = (i + 1) * slotWidth - Math.floor(slotWidth * 0.25)
      if (x >= left && x < right && y >= chartBottom - barHeight && y <= chartBottom) {
        return bars[i].color
      }
    }
    return [255, 255, 255]
  })
}

/* ------------------------------------------------------------------ *
 * The document
 * ------------------------------------------------------------------ */

const outDir = process.argv[2]
  ? path.dirname(path.resolve(process.argv[2]))
  : path.join(os.tmpdir(), 'r7-acceptance')
const docxPath = process.argv[2]
  ? path.resolve(process.argv[2])
  : path.join(outDir, 'R7_MCP_DOCX_Acceptance.docx')
const pdfPath = path.join(outDir, 'R7_MCP_DOCX_Acceptance.pdf')

fs.mkdirSync(outDir, { recursive: true })

const engine = new DocxEngine()
const adapter = new R7Adapter()

console.log('=== R7 MCP DOCX acceptance document ===\n')

console.log('1. Creating the document with the public engine API...')
await engine.create(docxPath, {
  overwrite: true,
  title: 'Коммерческое предложение',
  paragraphs: [
    { text: 'Внедрение корпоративной системы документооборота', style: 'Subtitle' },
    {
      text: 'ООО «Ромашка» · Департамент цифровизации · 12 февраля 2026 г.',
      alignment: 'center',
      size: 10,
      color: '#595959',
      spacing: { after: 18 }
    },

    { text: '1. Резюме предложения', style: 'Heading1' },
    {
      text: 'Компания Р7-Офис предлагает внедрить единый контур работы с документами: '
        + 'тексты, таблицы, презентации и PDF обрабатываются в одном приложении, без потери '
        + 'разметки при обмене файлами с контрагентами. Ниже приведены состав работ, сроки и '
        + 'экономический эффект для подразделений заказчика.',
      alignment: 'both',
      indents: { firstLine: 35 },
      spacing: { after: 10, line: 1.15, lineRule: 'auto' }
    },
    {
      text: 'Ключевые выгоды проекта:',
      spacing: { before: 8, after: 4 },
      bold: true
    },
    { text: 'Единый формат DOCX/XLSX/PPTX без конвертеров и «поехавшей» вёрстки.', list: 'bullet' },
    { text: 'Работа с документами прямо из агентной системы через MCP-интерфейс.', list: 'bullet' },
    { text: 'Соответствие требованиям импортозамещения и размещение в закрытом контуре.', list: 'bullet' },
    { text: 'Сохранение привычных горячих клавиш и шаблонов Microsoft Office.', list: 'bullet' },

    { text: '2. Состав работ', style: 'Heading1' },
    { text: '2.1. Подготовительный этап', style: 'Heading2' },
    {
      text: 'На подготовительном этапе проводится аудит текущего документооборота, собираются '
        + 'типовые шаблоны и описываются сценарии согласования. Результат этапа — техническое '
        + 'задание, согласованное всеми подразделениями.',
      alignment: 'both',
      indents: { left: 18, firstLine: 35 },
      spacing: { after: 8, line: 1.15, lineRule: 'auto' }
    },
    { text: 'Последовательность работ:', spacing: { before: 6, after: 4 } },
    { text: 'Аудит текущих процессов и реестра шаблонов.', list: 'number' },
    { text: 'Проектирование структуры папок, прав и маршрутов согласования.', list: 'number' },
    { text: 'Установка и настройка серверной части на площадке заказчика.', list: 'number' },
    { text: 'Миграция архива документов и обучение сотрудников.', list: 'number' },

    { text: '2.2. Результаты по кварталам', style: 'Heading2' },
    {
      text: 'В таблице приведён план-график работ с ответственными подразделениями и плановыми '
        + 'показателями. Значения уточняются после аудита.',
      indents: { left: 18 },
      spacing: { after: 8 }
    },
    {
      text: 'Показатель «55 %» означает долю обработанных документов от планового объёма в первом '
        + 'квартале пилота: в контур были включены два отдела, а часть исторических документов '
        + 'осталась в старой системе. К третьему кварталу в контур переводятся все подразделения, '
        + 'включая филиалы, и показатель выходит на 100 %.',
      alignment: 'both',
      indents: { left: 18, firstLine: 35 },
      spacing: { after: 10, line: 1.15, lineRule: 'auto' }
    },
    {
      text: 'Отклонения от плана фиксируются в еженедельном статусе и разбираются на совещании '
        + 'рабочей группы. Каждое отклонение сопровождается решением: сдвиг срока, перераспределение '
        + 'ответственных или изменение объёма работ в пределах утверждённого бюджета.',
      alignment: 'both',
      indents: { left: 18, firstLine: 35 },
      spacing: { after: 12, line: 1.15, lineRule: 'auto' }
    },

    { text: '3. Динамика показателей', style: 'Heading1' },
    {
      text: 'На диаграмме показан рост количества обработанных документов по кварталам: '
        + '55 %, 78 % и 100 % от планового объёма. Пилотный контур вышел на плановую мощность '
        + 'на третьем месяце эксплуатации.',
      alignment: 'both',
      indents: { firstLine: 35 },
      spacing: { after: 12, line: 1.15, lineRule: 'auto' }
    },
    {
      text: 'Нагрузка распределяется неравномерно: пики приходятся на конец квартала, когда '
        + 'подразделения готовят отчётность, и на период договорной кампании. Пиковая нагрузка '
        + 'на 40 % выше средней, поэтому мощности серверной части рассчитывались с двойным запасом, '
        + 'а пилотный контур тестировался именно на пиковых сценариях, а не на среднесуточных.',
      alignment: 'both',
      indents: { firstLine: 35, right: 18 },
      spacing: { after: 10, line: 1.15, lineRule: 'auto' }
    },

    { text: '4. Порядок взаимодействия', style: 'Heading1' },
    { text: '4.1. Каналы связи', style: 'Heading3' },
    {
      text: 'Все вопросы по проекту ведутся через единый почтовый ящик и еженедельные статусы. '
        + 'Техническая документация доступна на портале поддержки.',
      indents: { right: 24 },
      spacing: { after: 8 }
    },
    {
      text: 'Справочные материалы и описания форматов опубликованы на сайте r7-office.ru.',
      spacing: { after: 12 }
    },

    { text: '4.2. Регламент сопровождения', style: 'Heading3' },
    {
      text: 'После вывода системы в промышленную эксплуатацию заказчику передаётся регламент '
        + 'сопровождения: порядок регистрации обращений, сроки реакции по критичности, перечень '
        + 'работ по обновлению и правила резервного копирования. Регламент согласуется со службой '
        + 'эксплуатации и пересматривается не реже одного раза в год.',
      alignment: 'both',
      indents: { firstLine: 35 },
      spacing: { after: 8, line: 1.15, lineRule: 'auto' }
    },
    {
      text: 'Критичность обращений и сроки реакции:',
      spacing: { before: 6, after: 4 },
      bold: true
    },
    {
      text: 'Блокирующая ошибка — реакция в течение двух часов, работа без обхода до устранения.',
      indents: { left: 24, hanging: 12 },
      spacing: { after: 4 }
    },
    {
      text: 'Существенная ошибка — реакция в течение рабочего дня, обход предлагается заказчику.',
      indents: { left: 24, hanging: 12 },
      spacing: { after: 4 }
    },
    {
      text: 'Незначительное замечание — включается в ближайшее плановое обновление.',
      indents: { left: 24, hanging: 12 },
      spacing: { after: 12 }
    },

    { text: '5. Экономический эффект', style: 'Heading1' },
    {
      text: 'Расчёт эффекта построен на сокращении времени подготовки типовых документов и на отказе '
        + 'от ручного переноса данных между редакторами. Замеры проводились на пилотной группе из '
        + 'тридцати сотрудников в течение двух месяцев: среднее время согласования одного договора '
        + 'сократилось с 4 часов 20 минут до 1 часа 45 минут.',
      alignment: 'both',
      indents: { firstLine: 35 },
      spacing: { after: 10, line: 1.15, lineRule: 'auto' }
    },
    {
      text: 'Дополнительный эффект даёт единый поиск по архиву: сотрудники перестают запрашивать '
        + 'актуальную версию документа по почте, а значит исчезает целый класс ошибок, связанных '
        + 'с работой над устаревшей копией файла.',
      alignment: 'both',
      indents: { firstLine: 35 },
      spacing: { after: 10, line: 1.15, lineRule: 'auto' }
    },

    { text: '6. Заключение', style: 'Heading1' },
    {
      text: 'Предложенный план покрывает полный цикл — от аудита до сопровождения — и рассчитан на '
        + 'три квартала. Для старта проекта достаточно согласовать техническое задание и выделить '
        + 'ответственных со стороны заказчика.',
      alignment: 'both',
      indents: { firstLine: 35 },
      spacing: { after: 10, line: 1.15, lineRule: 'auto' }
    },
    { text: 'Ближайшие шаги:', spacing: { before: 6, after: 4 } },
    { text: 'Согласовать состав пилотной группы.', list: 'number' },
    { text: 'Подтвердить площадку для установки серверной части.', list: 'number' },
    { text: 'Назначить дату старта аудита.', list: 'number' },

    { text: 'Приложение А. Оборудование и лицензии', style: 'Heading1', pageBreak: true },
    {
      text: 'Приложение оформлено в альбомной ориентации: широкие таблицы спецификации '
        + 'читаются без горизонтальной прокрутки.',
      alignment: 'both',
      indents: { firstLine: 35 },
      spacing: { after: 10 }
    }
  ],
  tables: [
    {
      widthsTwips: [1800, 3400, 2200, 2000],
      rows: [
        [
          { value: 'Квартал', shading: '#DEEAF6', alignment: 'center', verticalAlign: 'center' },
          { value: 'Работы', shading: '#DEEAF6', alignment: 'center', verticalAlign: 'center' },
          { value: 'Ответственный', shading: '#DEEAF6', alignment: 'center', verticalAlign: 'center' },
          { value: 'Показатель', shading: '#DEEAF6', alignment: 'center', verticalAlign: 'center' }
        ],
        ['Q1 2026', 'Аудит и проектирование', 'Аналитик', '55 %'],
        ['Q2 2026', 'Пилотный контур в 2 отделах', 'Внедренец', '78 %'],
        ['Q3 2026', 'Тиражирование на филиалы', 'Группа внедрения', '100 %'],
        [{ value: 'Итого', shading: '#F2F2F2', alignment: 'right' },
          { value: 'Единый контур документооборота', shading: '#F2F2F2' },
          { value: '—', shading: '#F2F2F2', alignment: 'center' },
          { value: '3 квартала', shading: '#F2F2F2', alignment: 'center' }]
      ]
    },
    {
      widthsTwips: [3600, 1800, 1800, 2200],
      rows: [
        [
          { value: 'Показатель', shading: '#E2EFDA', alignment: 'center' },
          { value: 'До внедрения', shading: '#E2EFDA', alignment: 'center' },
          { value: 'После внедрения', shading: '#E2EFDA', alignment: 'center' },
          { value: 'Эффект', shading: '#E2EFDA', alignment: 'center' }
        ],
        ['Время согласования договора', '4 ч 20 мин', '1 ч 45 мин', '−60 %'],
        ['Ошибки версий документов, в месяц', '18', '3', '−83 %'],
        ['Стоимость печати и курьеров, в год', '620 000 ₽', '210 000 ₽', '−66 %']
      ]
    }
  ],
  // A portrait A4 page with comfortable margins.
  page: { orientation: 'portrait', margins: { top: 2.5, right: 2, bottom: 2.5, left: 2.5 } },
  header: { text: 'ООО «Ромашка» · Коммерческое предложение', alignment: 'center' },
  footer: { text: 'Стр.', pageNumber: 'PAGE', alignment: 'center' }
})
console.log(`   created: ${docxPath}`)

console.log('\n2. Inserting the chart image (14 cm wide, aspect ratio preserved)...')
const image = await engine.insertImage(docxPath, {
  buffer: renderChart(),
  widthCm: 14,
  paragraphIndex: 'new',
  alt: 'Рост количества обработанных документов по кварталам',
  name: 'Динамика показателей'
})
console.log(`   ${image.mediaPartName} · ${image.widthCm} x ${image.heightCm} cm · ${image.scaled}`)

console.log('\n3. Adding an external hyperlink...')
const link = await engine.hyperlink(docxPath, {
  action: 'insert',
  text: 'r7-office.ru — описание форматов и загрузки',
  url: 'https://r7-office.ru/',
  newParagraph: true,
  alignment: 'center'
})
console.log(`   relationship ${link.relId} -> https://r7-office.ru/`)

console.log('\n4. Splitting the appendix into its own landscape section...')
const outline = await engine.read(docxPath, { format: 'structured' })
const appendix = outline.paragraphs.find(entry => entry.text.startsWith('Приложение А'))
if (!appendix) throw new Error('the appendix heading was not found')

// A section break closes the section that ends at the *previous* paragraph, so
// every paragraph from the appendix onwards belongs to the landscape section.
const section = await engine.insertSectionBreak(docxPath, {
  afterParagraphIndex: appendix.index - 1,
  type: 'nextPage',
  page: { orientation: 'landscape', margins: { top: 1.5, right: 1.5, bottom: 1.5, left: 1.5 } }
})
console.log(`   sections: ${section.sectionCount}, break after paragraph ${appendix.index - 1}`)

console.log('\n5. Appending the equipment specification table to the landscape appendix...')
await engine.table(docxPath, {
  action: 'create',
  position: 'end',
  widthsTwips: [3000, 2600, 2000, 1800, 2000],
  rows: [
    [
      { value: 'Позиция', shading: '#DEEAF6', alignment: 'center' },
      { value: 'Назначение', shading: '#DEEAF6', alignment: 'center' },
      { value: 'Количество', shading: '#DEEAF6', alignment: 'center' },
      { value: 'Срок поставки', shading: '#DEEAF6', alignment: 'center' },
      { value: 'Ответственный', shading: '#DEEAF6', alignment: 'center' }
    ],
    ['Сервер приложений', 'Серверная часть Р7-Офис', '2', 'Q2 2026', 'ИТ-служба'],
    ['Хранилище 8 ТБ', 'Архив документов и резервные копии', '1', 'Q2 2026', 'ИТ-служба'],
    ['Лицензии на 250 мест', 'Рабочие места сотрудников', '1 пакет', 'Q2 2026', 'Закупки'],
    ['Модуль конвертации PDF', 'Печать и выгрузка документов', '1', 'Q3 2026', 'ИТ-служба']
  ]
})
console.log('   specification table added')

console.log('\n6. Reading the result back through the public API...')
const inspection = await engine.inspect(docxPath, {
  includeStyles: true,
  includeSections: true,
  includeHeadersFooters: true,
  includeHyperlinks: true,
  includeImages: true
})
const sections = await engine.sections(docxPath)
const lists = await engine.lists(docxPath)
const links = await engine.hyperlink(docxPath, { action: 'list' })
const images = await engine.images(docxPath)
const tables = await engine.tables(docxPath)

console.log(`   paragraphs: ${inspection.paragraphsCount}, headings: ${inspection.headingsCount}, tables: ${inspection.tablesCount}`)
console.log(`   list items: ${lists.length} (${lists.filter(item => item.format === 'bullet').length} bullet, ${lists.filter(item => item.format === 'decimal').length} numbered)`)
console.log(`   pages/sections: ${sections.sectionCount} — ${sections.sections.map(s => `${s.kind}/${s.pageSize.orientation}`).join(', ')}`)
console.log(`   page breaks: ${sections.pageBreaks.length}`)
console.log(`   images: ${images.map(i => `${i.mediaPartName} ${i.widthCm}x${i.heightCm}cm`).join(', ')}`)
console.log(`   hyperlinks: ${links.hyperlinks.map(l => `${l.text} -> ${l.url}`).join(', ')}`)
for (const part of inspection.headersFooters) {
  console.log(`   ${part.kind} (${part.type}): "${part.text}" page-number=${part.hasPageNumberField}`)
}
console.log(`   table 1: ${tables[0].rows} rows x ${tables[0].cols} cols`)
for (const row of tables[0].data) console.log(`     | ${row.join(' | ')}`)
console.log(`   table 2: ${tables[1].rows} rows x ${tables[1].cols} cols`)
for (const row of tables[1].data) console.log(`     | ${row.join(' | ')}`)

const validation = await engine.validate(docxPath)
if (!validation.valid) throw new Error(`the document does not validate: ${validation.errors.join('; ')}`)
console.log(`   valid: true (${fs.statSync(docxPath).size} bytes, ${validation.details.fileEntries} package parts)`)

console.log('\n7. Rendering to PDF through the R7 adapter...')
const info = await adapter.detect()
let pdfSize = null
if (!info.installed) {
  console.log('   R7-Office is not installed on this host: PDF rendering skipped.')
} else {
  const started = Date.now()
  await adapter.convert(docxPath, pdfPath)
  const head = fs.readFileSync(pdfPath).subarray(0, 5).toString()
  if (head !== '%PDF-') {
    throw new Error(`the renderer did not produce a PDF (header was "${head}")`)
  }
  pdfSize = fs.statSync(pdfPath).size
  const pageCount = (fs.readFileSync(pdfPath).toString('latin1').match(/\/Type\s*\/Page[^s]/g) || []).length
  console.log(`   ${pdfPath} · ${pdfSize} bytes · ~${pageCount} pages · ${Date.now() - started} ms`)
  console.log('   PDF header verified: %PDF-')
}

console.log('\n=== Result ===')
console.log(`DOCX: ${docxPath} (${fs.statSync(docxPath).size} bytes)`)
if (pdfSize !== null) console.log(`PDF:  ${pdfPath} (${pdfSize} bytes)`)
else console.log('PDF:  not produced (R7-Office not detected)')
console.log('Both files live outside the repository and are not committed.')
