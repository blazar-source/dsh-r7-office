/**
 * Acceptance workbook: `R7_MCP_XLSX_Acceptance.xlsx`.
 *
 * Builds a three-sheet finance workbook using only the public engine API
 * (`create`, `write`, `format`, `read`, `inspect`, `validate`), then proves the
 * result the way a caller would: read it back with formatting, check the stored
 * values, and render it through the installed R7 engine.
 *
 * The workbook is deliberately the awkward case — real dates stored as serials,
 * currency and percent number formats, a merged title, borders on a data block,
 * cross-sheet formulas, mixed alignment and per-column widths — because that is
 * what breaks a spreadsheet writer.
 *
 * It also carries a print layout for every sheet (fit-to-width so no column
 * spills onto a page of its own, landscape, A4, a print area over the used
 * range and the title rows repeating), and the render at the end exports ALL
 * THREE sheets into one PDF, in tab order — the default for the XLSX engine.
 *
 * Run: node examples/xlsx-acceptance.js [outputDirectory]
 * Default output directory: <os.tmpdir()>/r7-acceptance
 */

import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { XlsxEngine, R7Adapter } from '../src/r7/index.js'
import { serialToDate } from '../src/r7/xlsx-styles.js'

const outDir = process.argv[2] || path.join(os.tmpdir(), 'r7-acceptance')
fs.mkdirSync(outDir, { recursive: true })

const bookPath = path.join(outDir, 'R7_MCP_XLSX_Acceptance.xlsx')
const pdfPath = path.join(outDir, 'R7_MCP_XLSX_Acceptance.pdf')

const INCOME = 'Доходы'
const EXPENSES = 'Расходы'
const SUMMARY = 'Итоги'

const xlsx = new XlsxEngine()
const adapter = new R7Adapter()

const checks = []
function check(name, ok, detail = '') {
  checks.push({ name, ok, detail })
  console.log(`   ${ok ? 'OK  ' : 'FAIL'} ${name}${detail ? ` — ${detail}` : ''}`)
}

const money = { type: 'currency', symbol: '₽', decimals: 2, thousands: true }
const dataBorder = { all: { style: 'thin', color: '#B4C6E7' } }
const totalBorder = { all: { style: 'thin', color: '#8EAADB' } }

// ------------------------------------------------------------------ the data

const incomeRows = [
  ['2026-01-15', 'Продажи лицензий', 850000, 'Годовые контракты'],
  ['2026-02-15', 'Услуги внедрения', 420000, 'Три проекта'],
  ['2026-02-28', 'Техподдержка', 165000, 'Продления'],
  ['2026-03-15', 'Продажи лицензий', 910000, 'Квартальный пик'],
  ['2026-03-31', 'Обучение', 120000, 'Два потока']
]

const expenseRows = [
  ['2026-01-20', 'ФОТ', 740000, 'Январь'],
  ['2026-02-10', 'Аренда офиса', 180000, 'I квартал'],
  ['2026-02-20', 'ФОТ', 740000, 'Февраль'],
  ['2026-03-05', 'Маркетинг', 260000, 'Кампании'],
  ['2026-03-20', 'Лицензии ПО', 95000, 'Продление'],
  ['2026-03-31', 'Командировки', 64000, 'Поездки']
]

const incomeTotalRow = incomeRows.length + 3   // header is row 2, data starts at row 3
const expenseTotalRow = expenseRows.length + 3
const incomeLastDataRow = incomeTotalRow - 1
const expenseLastDataRow = expenseTotalRow - 1

console.log('1. Creating the workbook with three worksheets...')
await xlsx.create(bookPath, {
  overwrite: true,
  sheets: [
    { name: INCOME, data: [] },
    { name: EXPENSES, data: [] },
    { name: SUMMARY, data: [] }
  ]
})

// ------------------------------------------------------------------ Доходы

console.log(`2. Filling "${INCOME}"...`)
const incomeCells = [
  { ref: 'A1', value: 'Доходы за I квартал 2026 года' },
  { ref: 'A2', value: 'Дата' },
  { ref: 'B2', value: 'Категория' },
  { ref: 'C2', value: 'Сумма, ₽' },
  { ref: 'D2', value: 'Комментарий' }
]
incomeRows.forEach((row, i) => {
  const r = i + 3
  incomeCells.push({ ref: `A${r}`, value: row[0], date: true })
  incomeCells.push({ ref: `B${r}`, value: row[1] })
  incomeCells.push({ ref: `C${r}`, value: row[2], numberFormat: money })
  incomeCells.push({ ref: `D${r}`, value: row[3] })
})
incomeCells.push({ ref: `A${incomeTotalRow}`, value: 'Итого' })
incomeCells.push({ ref: `B${incomeTotalRow}`, value: 'за квартал' })
incomeCells.push({
  ref: `C${incomeTotalRow}`,
  formula: `SUM(C3:C${incomeLastDataRow})`,
  numberFormat: money
})
incomeCells.push({ ref: `D${incomeTotalRow}`, value: '5 категорий' })
await xlsx.write(bookPath, { sheetName: INCOME, cells: incomeCells })

console.log(`3. Formatting "${INCOME}"...`)
await xlsx.format(bookPath, {
  sheetName: INCOME,
  range: 'A1:D1',
  merge: true,
  font: { bold: true, size: 14, color: '#FFFFFF' },
  fill: { color: '#2F5597' },
  alignment: { horizontal: 'center', vertical: 'center' },
  rowHeight: 26
})
await xlsx.format(bookPath, {
  sheetName: INCOME,
  range: 'A2:D2',
  font: { bold: true, color: '#FFFFFF' },
  fill: { color: '#4472C4' },
  border: dataBorder,
  alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
  rowHeight: 30
})
await xlsx.format(bookPath, { sheetName: INCOME, range: 'A2:D8', border: dataBorder })
await xlsx.format(bookPath, {
  sheetName: INCOME,
  range: `A3:A${incomeLastDataRow}`,
  numberFormat: { type: 'date' },
  alignment: { horizontal: 'center' }
})
await xlsx.format(bookPath, {
  sheetName: INCOME,
  range: `C3:C${incomeTotalRow}`,
  numberFormat: money,
  alignment: { horizontal: 'right' }
})
await xlsx.format(bookPath, {
  sheetName: INCOME,
  range: `A${incomeTotalRow}:D${incomeTotalRow}`,
  font: { bold: true },
  fill: { color: '#D9E2F3' },
  border: totalBorder
})
await xlsx.format(bookPath, { sheetName: INCOME, range: 'A1:A7', columnWidth: { width: 13 } })
await xlsx.format(bookPath, { sheetName: INCOME, range: 'B1:B7', columnWidth: { width: 24 } })
await xlsx.format(bookPath, { sheetName: INCOME, range: 'C1:C7', columnWidth: { width: 17 } })
await xlsx.format(bookPath, { sheetName: INCOME, range: 'D1:D7', columnWidth: { width: 30 } })

// ----------------------------------------------------------------- Расходы

console.log(`4. Filling "${EXPENSES}"...`)
const expenseCells = [
  { ref: 'A1', value: 'Расходы за I квартал 2026 года' },
  { ref: 'A2', value: 'Дата' },
  { ref: 'B2', value: 'Статья' },
  { ref: 'C2', value: 'Сумма, ₽' },
  { ref: 'D2', value: 'Комментарий' }
]
expenseRows.forEach((row, i) => {
  const r = i + 3
  expenseCells.push({ ref: `A${r}`, value: row[0], date: true })
  expenseCells.push({ ref: `B${r}`, value: row[1] })
  expenseCells.push({ ref: `C${r}`, value: row[2], numberFormat: money })
  expenseCells.push({ ref: `D${r}`, value: row[3] })
})
expenseCells.push({ ref: `A${expenseTotalRow}`, value: 'Итого' })
expenseCells.push({ ref: `B${expenseTotalRow}`, value: 'за квартал' })
expenseCells.push({
  ref: `C${expenseTotalRow}`,
  formula: `SUM(C3:C${expenseLastDataRow})`,
  numberFormat: money
})
expenseCells.push({ ref: `D${expenseTotalRow}`, value: '6 статей' })
await xlsx.write(bookPath, { sheetName: EXPENSES, cells: expenseCells })

console.log(`5. Formatting "${EXPENSES}"...`)
await xlsx.format(bookPath, {
  sheetName: EXPENSES,
  range: 'A1:D1',
  merge: true,
  font: { bold: true, size: 14, color: '#FFFFFF' },
  fill: { color: '#833C0C' },
  alignment: { horizontal: 'center', vertical: 'center' },
  rowHeight: 26
})
await xlsx.format(bookPath, {
  sheetName: EXPENSES,
  range: 'A2:D2',
  font: { bold: true, color: '#FFFFFF' },
  fill: { color: '#C55A11' },
  border: { all: { style: 'thin', color: '#FFFFFF' } },
  alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
  rowHeight: 30
})
await xlsx.format(bookPath, { sheetName: EXPENSES, range: 'A2:D9', border: dataBorder })
await xlsx.format(bookPath, {
  sheetName: EXPENSES,
  range: `A3:A${expenseLastDataRow}`,
  numberFormat: { type: 'date' },
  alignment: { horizontal: 'center' }
})
await xlsx.format(bookPath, {
  sheetName: EXPENSES,
  range: `C3:C${expenseTotalRow}`,
  numberFormat: money,
  alignment: { horizontal: 'right' }
})
await xlsx.format(bookPath, {
  sheetName: EXPENSES,
  range: `A${expenseTotalRow}:D${expenseTotalRow}`,
  font: { bold: true },
  fill: { color: '#FBE5D6' },
  border: totalBorder
})
await xlsx.format(bookPath, { sheetName: EXPENSES, range: 'A1:A7', columnWidth: { width: 13 } })
await xlsx.format(bookPath, { sheetName: EXPENSES, range: 'B1:B7', columnWidth: { width: 22 } })
await xlsx.format(bookPath, { sheetName: EXPENSES, range: 'C1:C7', columnWidth: { width: 17 } })
await xlsx.format(bookPath, { sheetName: EXPENSES, range: 'D1:D7', columnWidth: { width: 26 } })

// ------------------------------------------------------------------- Итоги

console.log(`6. Filling "${SUMMARY}" with cross-sheet formulas...`)
await xlsx.write(bookPath, {
  sheetName: SUMMARY,
  cells: [
    { ref: 'A1', value: 'Итоги по проекту за I квартал 2026 года' },
    { ref: 'A2', value: 'Показатель' },
    { ref: 'B2', value: 'План, ₽' },
    { ref: 'C2', value: 'Факт, ₽' },
    { ref: 'D2', value: 'Отклонение, ₽' },
    { ref: 'E2', value: '% выполнения' },

    { ref: 'A3', value: 'Доходы' },
    { ref: 'B3', value: 2350000, numberFormat: money },
    { ref: 'C3', formula: `${INCOME}!C${incomeTotalRow}`, numberFormat: money },
    { ref: 'D3', formula: 'C3-B3', numberFormat: money },
    { ref: 'E3', formula: 'C3/B3', numberFormat: { type: 'percent', decimals: 1 } },

    { ref: 'A4', value: 'Расходы' },
    { ref: 'B4', value: 2100000, numberFormat: money },
    { ref: 'C4', formula: `${EXPENSES}!C${expenseTotalRow}`, numberFormat: money },
    { ref: 'D4', formula: 'C4-B4', numberFormat: money },
    { ref: 'E4', formula: 'C4/B4', numberFormat: { type: 'percent', decimals: 1 } },

    { ref: 'A5', value: 'Прибыль' },
    { ref: 'B5', formula: 'B3-B4', numberFormat: money },
    { ref: 'C5', formula: 'C3-C4', numberFormat: money },
    { ref: 'D5', formula: 'C5-B5', numberFormat: money },
    { ref: 'E5', formula: 'C5/B5', numberFormat: { type: 'percent', decimals: 1 } },

    { ref: 'A6', value: 'Итого' },
    { ref: 'B6', formula: 'SUM(B3:B5)', numberFormat: money },
    { ref: 'C6', formula: 'SUM(C3:C5)', numberFormat: money },
    { ref: 'D6', formula: 'SUM(D3:D5)', numberFormat: money },
    { ref: 'E6', formula: 'AVERAGE(E3:E5)', numberFormat: { type: 'percent', decimals: 1 } }
  ]
})

console.log(`7. Formatting "${SUMMARY}"...`)
await xlsx.format(bookPath, {
  sheetName: SUMMARY,
  range: 'A1:E1',
  merge: true,
  font: { bold: true, size: 15, color: '#FFFFFF' },
  fill: { color: '#1F4E79' },
  alignment: { horizontal: 'center', vertical: 'center' },
  rowHeight: 30
})
await xlsx.format(bookPath, {
  sheetName: SUMMARY,
  range: 'A2:E2',
  font: { bold: true, color: '#FFFFFF' },
  fill: { color: '#2F5597' },
  border: { all: { style: 'thin', color: '#9DC3E6' } },
  alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
  rowHeight: 34
})
await xlsx.format(bookPath, { sheetName: SUMMARY, range: 'A2:E6', border: { all: { style: 'thin', color: '#9DC3E6' } } })
await xlsx.format(bookPath, { sheetName: SUMMARY, range: 'B3:D6', numberFormat: money, alignment: { horizontal: 'right' } })
await xlsx.format(bookPath, { sheetName: SUMMARY, range: 'E3:E6', alignment: { horizontal: 'right' } })
await xlsx.format(bookPath, {
  sheetName: SUMMARY,
  range: 'A6:E6',
  font: { bold: true },
  fill: { color: '#DDEBF7' }
})
await xlsx.format(bookPath, { sheetName: SUMMARY, range: 'A1:A7', columnWidth: { width: 22 } })
await xlsx.format(bookPath, { sheetName: SUMMARY, range: 'B1:B7', columnWidth: { width: 15 } })
await xlsx.format(bookPath, { sheetName: SUMMARY, range: 'C1:C7', columnWidth: { width: 15 } })
await xlsx.format(bookPath, { sheetName: SUMMARY, range: 'D1:D7', columnWidth: { width: 16 } })
await xlsx.format(bookPath, { sheetName: SUMMARY, range: 'E1:E7', columnWidth: { width: 14 } })

// -------------------------------------------------------------- verification

console.log('\n8. Verifying the workbook through the engine...')
const info = await xlsx.inspect(bookPath)
check('three worksheets exist', info.sheetsCount === 3, info.sheets.map((s) => s.name).join(', '))
check('sheet names are the requested ones',
  JSON.stringify(info.sheets.map((s) => s.name)) === JSON.stringify([INCOME, EXPENSES, SUMMARY]))

const income = await xlsx.read(bookPath, {
  sheetName: INCOME,
  range: `A1:D${incomeTotalRow}`,
  includeStyles: true,
  includeFormulas: true
})
check('the title is merged across the data columns',
  JSON.stringify(income.merged) === JSON.stringify(['A1:D1']), JSON.stringify(income.merged))
check('the title is bold, sized and filled',
  income.styles[0][0].font.bold === true
  && income.styles[0][0].font.size === 14
  && income.styles[0][0].font.color === 'FFFFFFFF'
  && income.styles[0][0].fill.color === 'FF2F5597')
check('a date column holds real date serials',
  income.data[2][0] === 46037 && income.styles[2][0].numberFormatCode === 'DD.MM.YYYY',
  `A3 = ${income.data[2][0]} (${serialToDate(income.data[2][0]).date})`)
check('a currency column holds numbers with a currency format',
  typeof income.data[2][2] === 'number'
  && income.styles[2][2].numberFormatCode.includes('₽'),
  `C3 = ${income.data[2][2]}, code ${income.styles[2][2].numberFormatCode}`)
check('the totals row carries a SUM formula',
  income.formulas[incomeTotalRow - 1][2] === `SUM(C3:C${incomeLastDataRow})`,
  income.formulas[incomeTotalRow - 1][2])
check('the totals row keeps its currency format',
  income.styles[incomeTotalRow - 1][2].numberFormatCode.includes('₽'))
check('the header row has borders on every edge',
  ['left', 'right', 'top', 'bottom'].every((e) => income.styles[1][0].border[e].style === 'thin'))
check('the header row centers and wraps',
  income.styles[1][0].alignment.horizontal === 'center'
  && income.styles[1][0].alignment.wrapText === true)
check('column widths are the requested ones',
  JSON.stringify(income.columnWidths) === JSON.stringify([
    { min: 1, max: 1, width: 13 },
    { min: 2, max: 2, width: 24 },
    { min: 3, max: 3, width: 17 },
    { min: 4, max: 4, width: 30 }
  ]),
  JSON.stringify(income.columnWidths))
check('row heights are the requested ones',
  income.rowHeights['1'] === 26 && income.rowHeights['2'] === 30, JSON.stringify(income.rowHeights))

const summary = await xlsx.read(bookPath, {
  sheetName: SUMMARY,
  range: 'A1:E6',
  includeStyles: true,
  includeFormulas: true
})
check('the summary reads the other worksheets',
  summary.formulas[2][2] === `${INCOME}!C${incomeTotalRow}`
  && summary.formulas[3][2] === `${EXPENSES}!C${expenseTotalRow}`,
  `${summary.formulas[2][2]} / ${summary.formulas[3][2]}`)
check('the summary keeps its money and percent formats',
  summary.styles[2][1].numberFormatCode.includes('₽')
  && summary.styles[2][4].numberFormatCode === '0.0%',
  `${summary.styles[2][1].numberFormatCode} / ${summary.styles[2][4].numberFormatCode}`)
check('the summary title is merged over all five columns',
  JSON.stringify(summary.merged) === JSON.stringify(['A1:E1']))
check('the summary header wraps its text',
  summary.styles[1][4].alignment.wrapText === true
  && summary.styles[1][4].alignment.horizontal === 'center')

const validation = await xlsx.validate(bookPath)
check('the package validates', validation.valid === true, validation.errors.join('; '))

const expenses = await xlsx.read(bookPath, { sheetName: EXPENSES, range: `A1:D${expenseTotalRow}` })
check('the expense sheet is independent of the income sheet',
  expenses.data[0][0] === 'Расходы за I квартал 2026 года'
  && expenses.data[2][1] === 'ФОТ')

// -------------------------------------------------------------- print layout

console.log('\n9. Setting the print layout on every sheet...')

// fitToWidth: 1 with fitToHeight: 0 is the fix for "column D lands on page 2":
// the table is squeezed to one page wide while staying as many pages tall as it
// needs. The engine writes <sheetPr><pageSetUpPr fitToPage="1"/></sheetPr> with
// it — without that switch a renderer ignores fitToWidth entirely.
const printLayout = {
  orientation: 'landscape',
  fitToWidth: 1,
  fitToHeight: 0,
  paperSize: 'A4',
  margins: { left: 0.5, right: 0.5, top: 0.6, bottom: 0.6 },
  centerHorizontally: true
}

const layouts = [
  { sheetName: INCOME, printArea: `A1:D${incomeTotalRow}` },
  { sheetName: EXPENSES, printArea: `A1:D${expenseTotalRow}` },
  { sheetName: SUMMARY, printArea: 'A1:E6' }
]

for (const target of layouts) {
  const result = await xlsx.setPageSetup(bookPath, {
    ...printLayout,
    ...target,
    printTitles: '1:2'
  })
  check(`"${target.sheetName}" fits one page wide with the header rows repeating`,
    result.pageSetup.fitToPage === true
    && result.pageSetup.fitToWidth === 1
    && result.pageSetup.fitToHeight === 0
    && result.printArea === `${target.sheetName}!${target.printArea.replace(/([A-Z])(\d+)/g, '$$$1$$$2')}`
    && result.printTitles === `${target.sheetName}!$1:$2`,
    `fit ${result.pageSetup.fitToWidth}x${result.pageSetup.fitToHeight} on ${result.pageSetup.paperSizeName}, `
    + `area ${result.printArea}, titles ${result.printTitles}`)
}

const laidOut = await xlsx.inspect(bookPath)
check('every sheet reports its print layout back',
  laidOut.sheets.every((s) => s.pageSetup.fitToPage === true
    && s.pageSetup.fitToWidth === 1
    && s.pageSetup.fitToHeight === 0
    && s.pageSetup.orientation === 'landscape'),
  laidOut.sheets.map((s) => `${s.name}: ${s.pageSetup.orientation} ${s.pageSetup.fitToWidth}x${s.pageSetup.fitToHeight}`).join('; '))

// ----------------------------------------------------------------- rendering

console.log('\n10. Rendering EVERY sheet through the R7 engine...')
const r7 = await adapter.detect()
let renderedPdf = null
if (!r7.installed) {
  check('R7 x2t is installed', false, 'not detected on this host — the PDF was not produced')
} else {
  const started = Date.now()
  const exported = await xlsx.exportPdf(bookPath, { outputPath: pdfPath })
  renderedPdf = pdfPath
  const head = fs.readFileSync(pdfPath).subarray(0, 1024)
  check('the workbook renders to a PDF that starts with %PDF-',
    head.includes(Buffer.from('%PDF-')),
    `${fs.statSync(pdfPath).size} bytes in ${Date.now() - started}ms`)

  check('allSheets is the default and every worksheet was exported',
    exported.allSheets === true
    && exported.sheets.length === 3
    && exported.sheets.every((s) => s.name),
    exported.sheets.map((s) => s.name).join(' → '))

  check('each sheet fits on a single page, so no column spills onto its own page',
    exported.sheets.every((s) => s.pages === 1),
    exported.sheets.map((s) => `${s.name}: ${s.pages}p`).join('; '))

  check('the PDF holds one page per sheet, in tab order',
    exported.pageCount === 3,
    `${exported.pageCount} pages for ${exported.sheets.length} sheets`)
}

// ------------------------------------------------------------------- summary

const failed = checks.filter((c) => !c.ok)
console.log(`\n${checks.length - failed.length}/${checks.length} checks passed`)
for (const f of failed) console.log(`  FAILED: ${f.name} — ${f.detail}`)

const xlsxSize = fs.statSync(bookPath).size
console.log(`\nWorkbook: ${bookPath} (${xlsxSize} bytes)`)
if (renderedPdf) console.log(`PDF:      ${renderedPdf} (${fs.statSync(renderedPdf).size} bytes)`)
else console.log('PDF:      not produced (R7-Office is not installed)')

if (failed.length > 0) {
  process.exitCode = 1
}
