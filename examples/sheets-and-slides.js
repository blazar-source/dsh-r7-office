/**
 * Example: spreadsheets and presentations.
 *
 * Shows the two things that separate a real office document from a flat text
 * file — a workbook keeps several named worksheets, and a deck keeps several
 * slides. Both are edited structurally: adding a sheet or a slide registers a
 * new part in the package and leaves the parts that already exist untouched.
 *
 * Run: node examples/sheets-and-slides.js
 */

import path from 'node:path'
import os from 'node:os'
import fs from 'node:fs'
import { XlsxEngine, PptxEngine, R7Adapter } from '../src/r7/index.js'

const outDir = path.join(os.tmpdir(), 'dsh-r7-example-sheets-slides')
fs.mkdirSync(outDir, { recursive: true })

const xlsx = new XlsxEngine()
const pptx = new PptxEngine()
const adapter = new R7Adapter()

// ---------------------------------------------------------------- spreadsheet

const book = path.join(outDir, 'Бюджет.xlsx')

console.log('1. Workbook with three named worksheets...')
await xlsx.create(book, {
  // The example regenerates its own output on every run. Creating a document
  // over an existing file is refused unless this is explicit.
  overwrite: true,
  sheets: [
    { name: 'Доходы', data: [['Статья', 'Сумма'], ['Продажи', 1200000], ['Услуги', 340000]] },
    { name: 'Расходы', data: [['Статья', 'Сумма'], ['ФОТ', 620000], ['Аренда', 180000]] },
    { name: 'Итоги', data: [['Показатель', 'Значение'], ['Прибыль', '']] }
  ]
})

let info = await xlsx.inspect(book)
console.log('   sheets:', info.sheets.map(s => s.name).join(', '))

console.log('\n2. Reading each worksheet by name...')
for (const name of ['Доходы', 'Расходы']) {
  const sheet = await xlsx.read(book, { sheetName: name, range: 'A1:B3' })
  console.log(`   ${name}:`, JSON.stringify(sheet.data))
}

console.log('\n3. Cross-sheet formula in Итоги...')
await xlsx.write(book, {
  sheetName: 'Итоги',
  cells: [{ ref: 'B2', formula: '=SUM(Доходы!B2:B3)-SUM(Расходы!B2:B3)' }]
})
const totals = await xlsx.read(book, { sheetName: 'Итоги', range: 'A1:B2', includeFormulas: true })
console.log('   stored formula at B2:', totals.formulas[1][1])
console.log('   (the cached value appears once R7 opens and recalculates the workbook)')

console.log('\n4. Adding a fourth worksheet to the existing book...')
await xlsx.addSheet(book, {
  name: 'Прогноз',
  data: [['Квартал', 'Ожидание'], ['Q1', 400000], ['Q2', 520000]]
})
info = await xlsx.inspect(book)
console.log('   sheets now:', info.sheets.map(s => s.name).join(', '))

console.log('\n5. Validating the workbook...')
console.log('   valid:', (await xlsx.validate(book)).valid)

// ----------------------------------------------------------------- presentation

const deck = path.join(outDir, 'Презентация.pptx')

console.log('\n6. New presentation...')
await pptx.create(deck, { title: 'Отчёт за квартал', overwrite: true })

console.log('\n7. Appending slides (the first slide is not modified)...')
await pptx.addSlide(deck, { title: 'Финансовые показатели' })
await pptx.addSlide(deck, { title: 'Планы на следующий квартал' })

info = await pptx.inspect(deck)
console.log('   slides:')
for (const slide of info.slides) {
  console.log(`     ${slide.slideNumber}. ${slide.title}`)
}

console.log('\n8. Editing one slide...')
await pptx.editSlide(deck, { slideIndex: 1, title: 'Финансовые показатели (обновлено)' })
info = await pptx.inspect(deck)
console.log('   slide 2 is now:', info.slides[1].title)

console.log('\n9. Validating the presentation...')
console.log('   valid:', (await pptx.validate(deck)).valid)

// -------------------------------------------------------------------- rendering

const r7Info = await adapter.detect()
if (r7Info.installed) {
  console.log('\n10. Rendering both through the R7 engine...')
  for (const [src, name] of [[book, 'book'], [deck, 'deck']]) {
    const pdf = path.join(outDir, `${name}.pdf`)
    const res = await adapter.convert(src, pdf)
    console.log(`   ${path.basename(src)} -> ${path.basename(pdf)} (${fs.statSync(pdf).size} bytes, ${res.timeMs}ms)`)
  }
} else {
  console.log('\n10. R7-Office not detected — skipping native rendering.')
}

console.log(`\nOutput: ${outDir}`)
