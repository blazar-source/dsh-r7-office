/**
 * Formatting tests for the XLSX engine.
 *
 * Every case authors its own workbook at run time — nothing here depends on a
 * committed fixture — and then asks the engine what it actually did. The point
 * of the suite is that formatting is *additive*: a value, a formula, another
 * worksheet or another package member must come out of a format call exactly as
 * it went in.
 */

import { test, describe, before, after } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import { XlsxEngine } from '../../src/r7/xlsx.js'
import { R7Adapter } from '../../src/r7/adapter.js'
import { ZipArchive } from '../../src/shared/zip.js'
import { requiresR7 } from '../helpers/r7-gate.js'
import { Stylesheet, dateToSerial, serialToDate, normalizeArgb } from '../../src/r7/xlsx-styles.js'
import {
  colToIndex,
  indexToCol,
  parseRef,
  buildRef,
  parseRange,
  expandRange,
  normalizeRange,
  attr,
  setCellStyle,
  setCellContent,
  ensureCell,
  getCellStyle,
  getMergedCells,
  setMergedCells,
  getColumnWidths,
  setColumnWidth,
  getRowHeights,
  setRowHeight,
  estimateColumnWidth
} from '../../src/r7/xlsx-worksheet.js'

const tmpDir = path.join(os.tmpdir(), `dsh_r7_xlsx_format_${Date.now()}`)
const engine = new XlsxEngine()
let adapter = null

before(async () => {
  fs.mkdirSync(tmpDir, { recursive: true })
  adapter = new R7Adapter()
})

after(() => {
  try { fs.rmSync(tmpDir, { recursive: true, force: true }) } catch { /* ignore */ }
})

/** Path of a fresh copy of the base workbook, so tests never share a file. */
function copyTo(name) {
  const target = path.join(tmpDir, name)
  fs.copyFileSync(path.join(tmpDir, 'base.xlsx'), target)
  return target
}

/**
 * One cell of every parallel matrix, addressed the way a caller would.
 * The reader must have been asked for an A1-anchored range.
 */
function cell(result, ref) {
  const { col, row } = parseRef(ref)
  const anchor = parseRef(result.range.split(':')[0])
  assert.equal(`${anchor.col}${anchor.row}`, '00', `cell() needs an A1-anchored range, got ${result.range}`)
  return {
    value: result.data?.[row]?.[col],
    style: result.styles?.[row]?.[col],
    formula: result.formulas?.[row]?.[col]
  }
}

async function sheetXmlOf(filePath, sheetPath = 'xl/worksheets/sheet1.xml') {
  const zip = await ZipArchive.fromFile(filePath)
  return zip.getText(sheetPath)
}

function looksLikePdf(filePath) {
  return fs.readFileSync(filePath).subarray(0, 1024).includes(Buffer.from('%PDF-'))
}

describe('Worksheet XML primitives', () => {
  test('column letters and references round-trip', () => {
    assert.equal(colToIndex('A'), 0)
    assert.equal(colToIndex('Z'), 25)
    assert.equal(colToIndex('AA'), 26)
    assert.equal(colToIndex('ab'), 27)
    assert.equal(indexToCol(0), 'A')
    assert.equal(indexToCol(25), 'Z')
    assert.equal(indexToCol(26), 'AA')
    assert.equal(indexToCol(701), 'ZZ')
    for (const n of [0, 1, 25, 26, 27, 51, 701, 702]) {
      assert.equal(colToIndex(indexToCol(n)), n, `index ${n} survives both directions`)
    }
    assert.equal(buildRef(0, 0), 'A1')
    assert.equal(buildRef(9, 3), 'D10')
    assert.deepEqual(parseRef('$B$2'), { colLetters: 'B', col: 1, row: 1 })
    assert.throws(() => parseRef('ZZ'), /Invalid cell reference/)
    assert.throws(() => parseRef('1A'), /Invalid cell reference/)
  })

  test('ranges expand, normalize and reject a reversed end', () => {
    assert.deepEqual(expandRange('A1:B2'), ['A1', 'B1', 'A2', 'B2'])
    assert.equal(normalizeRange('b2:d5'), 'B2:D5')
    assert.equal(normalizeRange('C3'), 'C3')
    assert.deepEqual(parseRange('A1:D10').end, { colLetters: 'D', col: 3, row: 9 })
    assert.throws(() => parseRange('D10:A1'), /must not precede/)
  })

  test('setCellStyle keeps the payload and creates a missing cell in order', () => {
    const sheet = '<worksheet><sheetData>'
      + '<row r="1"><c r="A1"><v>1</v></c><c r="C1" t="str"><f>SUM(A1)</f><v>1</v></c></row>'
      + '</sheetData></worksheet>'

    const styled = setCellStyle(sheet, 'B1', 7)
    assert.ok(styled.includes('<c r="B1" s="7"/>'), 'the missing cell was created with its format')
    assert.ok(styled.indexOf('r="A1"') < styled.indexOf('r="B1"') && styled.indexOf('r="B1"') < styled.indexOf('r="C1"'),
      'the new cell is inserted in column order')

    const restyled = setCellStyle(styled, 'C1', 9)
    const restyledTag = restyled.match(/<c r="C1"[^>]*>/)[0]
    assert.ok(/\bs="9"/.test(restyledTag), 'the new format is applied')
    assert.ok(/\bt="str"/.test(restyledTag), 'the existing type attribute is kept')
    assert.ok(restyled.includes('<f>SUM(A1)</f>'), 'the formula is untouched')

    assert.equal(getCellStyle(sheet, 'C1'), null, 'a cell with no s is reported as unstyled')
    assert.equal(getCellStyle(sheet, 'B1'), null, 'a cell that does not exist is reported as unstyled')
    assert.equal(getCellStyle(restyled, 'C1'), 9)

    // Format 0 is the default, so it is expressed by the absence of the attribute.
    assert.ok(!setCellStyle(restyled, 'C1', 0).includes('s="0"'))
  })

  test('setCellContent replaces the payload, keeps the style and never swallows a neighbour', () => {
    const sheet = '<worksheet><sheetData>'
      + '<row r="1"><c r="A1"><v>1</v></c><c r="B1" s="3"/></row>'
      + '<row r="2"><c r="A2" t="inlineStr"><is><t>keep</t></is></c><c r="B2"><v>2</v></c></row>'
      + '</sheetData></worksheet>'

    const written = setCellContent(sheet, 'B1', { content: '<is><t>X</t></is>', type: 'inlineStr' })
    assert.ok(written.includes('<c r="B1" s="3" t="inlineStr"><is><t>X</t></is></c>'),
      'the empty styled cell was reopened for its payload and kept its format')
    assert.ok(written.includes('<row r="2"><c r="A2" t="inlineStr"><is><t>keep</t></is></c><c r="B2"><v>2</v></c></row>'),
      'the next row is intact — a self-closing cell must not consume it')

    // Writing a number must drop a stale string type.
    const numeric = setCellContent(sheet, 'A1', { content: '<v>42</v>' })
    assert.ok(numeric.includes('<c r="A1"><v>42</v></c>'))
  })

  test('merged ranges are added, removed and kept after sheetData', () => {
    const sheet = '<worksheet><dimension ref="A1:C3"/><sheetData>'
      + '<row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'

    const merged = setMergedCells(sheet, { add: ['a1:c1'] })
    assert.deepEqual(getMergedCells(merged), ['A1:C1'], 'the reference is normalized')
    assert.ok(merged.indexOf('</sheetData>') < merged.indexOf('<mergeCells'),
      'mergeCells follows sheetData, as the schema requires')
    assert.ok(merged.includes('count="1"'))

    const two = setMergedCells(merged, { add: ['A2:C2'] })
    assert.deepEqual(getMergedCells(two), ['A1:C1', 'A2:C2'])

    const removed = setMergedCells(two, { remove: ['A1:C1'] })
    assert.deepEqual(getMergedCells(removed), ['A2:C2'])
    assert.ok(!removed.includes('A1:C1'))

    const none = setMergedCells(removed, { remove: ['A2:C2'] })
    assert.deepEqual(getMergedCells(none), [])
    assert.ok(!none.includes('mergeCells'))
  })

  test('column widths split overlapping spans and keep unrelated ones', () => {
    const sheet = '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'

    const wide = setColumnWidth(sheet, { min: 1, max: 3, width: 20 })
    assert.deepEqual(getColumnWidths(wide), [{ min: 1, max: 3, width: 20 }])

    const split = setColumnWidth(wide, { min: 2, max: 2, width: 5 })
    assert.deepEqual(getColumnWidths(split), [
      { min: 1, max: 1, width: 20 },
      { min: 2, max: 2, width: 5 },
      { min: 3, max: 3, width: 20 }
    ], 'the neighbours keep the author\'s width')

    const untouched = setColumnWidth(wide, { min: 7, max: 7, width: 9 })
    assert.deepEqual(getColumnWidths(untouched), [
      { min: 1, max: 3, width: 20 },
      { min: 7, max: 7, width: 9 }
    ])

    assert.ok(split.indexOf('<cols>') < split.indexOf('<sheetData>'), 'cols precedes sheetData')
    assert.throws(() => setColumnWidth(sheet, { min: 1, max: 1, width: 0 }), /Invalid column width/)
    assert.throws(() => setColumnWidth(sheet, { min: 1, max: 1, width: 'wide' }), /Invalid column width/)
  })

  test('row heights are set, read back and preserved', () => {
    const sheet = '<worksheet><sheetData>'
      + '<row r="1"><c r="A1"><v>1</v></c></row>'
      + '<row r="2" ht="12.5" customHeight="1"><c r="A2"><v>2</v></c></row>'
      + '</sheetData></worksheet>'

    const tall = setRowHeight(sheet, 1, 30)
    assert.equal(getRowHeights(tall)['1'], 30)
    assert.equal(getRowHeights(tall)['2'], 12.5, 'an unrelated row keeps its height')
    assert.ok(tall.includes('ht="30" customHeight="1"'))
    assert.throws(() => setRowHeight(sheet, 1, -4), /Invalid row height/)
  })

  test('ensureCell only creates what is missing', () => {
    const sheet = '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'
    const withCell = ensureCell(sheet, 'A1')
    assert.equal(withCell, sheet, 'an existing cell is left exactly as it was')
    assert.ok(ensureCell(sheet, 'C1').includes('<c r="C1"/>'))
  })

  test('attributes are matched by their whole name, not as a substring', () => {
    // R7's own writer emits customWidth before width, and row spans before the
    // cells, so a substring search reads the wrong attribute.
    const col = '<col customWidth="1" min="1" max="1" width="14"/>'
    assert.equal(attr(col, 'width'), '14')
    assert.equal(attr(col, 'customWidth'), '1')
    assert.equal(attr(col, 'min'), '1')
    assert.equal(attr('<c r="A1" spans="1:2"/>', 's'), null, 'spans is not a cell style')
    assert.equal(attr('<c r="A1" t="s" s="3"/>', 't'), 's')

    const sheet = '<worksheet><cols><col customWidth="1" min="1" max="1" width="14"/>'
      + '<col customWidth="1" min="2" max="2" width="22"/></cols>'
      + '<sheetData><row r="1" spans="1:2"><c r="A1" t="s"><v>0</v></c><c r="B1" s="3"/></row></sheetData></worksheet>'
    assert.deepEqual(getColumnWidths(sheet), [
      { min: 1, max: 1, width: 14 },
      { min: 2, max: 2, width: 22 }
    ], 'the widths are read, not the customWidth flag')
    assert.equal(getCellStyle(sheet, 'A1'), null)
    assert.equal(getCellStyle(sheet, 'B1'), 3)
  })

  test('editing a cell keeps the row height and the row attributes', () => {
    const sheet = '<worksheet><sheetData>'
      + '<row r="1" ht="28.5" customHeight="1" spans="1:2"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>'
      + '</sheetData></worksheet>'

    const styled = setCellStyle(sheet, 'B1', 4)
    assert.deepEqual(getRowHeights(styled), { 1: 28.5 }, 'the height survives a style edit')
    assert.ok(/\bht="28.5"/.test(styled))
    assert.ok(/\bcustomHeight="1"/.test(styled))
    assert.ok(/\bspans="1:2"/.test(styled), 'attributes this module does not model are kept')

    const created = ensureCell(styled, 'C1')
    assert.deepEqual(getRowHeights(created), { 1: 28.5 }, 'the height survives a new cell')

    const retitled = setRowHeight(created, 1, 40)
    assert.deepEqual(getRowHeights(retitled), { 1: 40 }, 'a new height replaces the old one')
    assert.ok(!/\bht="28.5"/.test(retitled), 'the old height attribute is gone')
    assert.equal((retitled.match(/\bht=/g) || []).length, 1, 'the height is written exactly once')
  })

  test('declared number formats below 164 are re-emitted and never renumbered', () => {
    // R7 allocates its own custom formats from 160, not from 164, so ">= 164"
    // is not what makes a declaration custom — being declared is.
    const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<numFmts count="2"><numFmt numFmtId="160" formatCode="DD.MM.YYYY"/>'
      + '<numFmt numFmtId="161" formatCode="0.00 &quot;₽&quot;"/></numFmts>'
      + '<fonts count="1"><font><sz val="11"/></font></fonts>'
      + '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
      + '<borders count="1"><border/></borders>'
      + '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
      + '<xf numFmtId="160" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>'
      + '</styleSheet>'

    const sheet = new Stylesheet(xml)
    const added = sheet.ensureCellXf({ base: 1, font: { bold: true } })
    const out = sheet.toXml()

    assert.ok(out.includes('<numFmt numFmtId="160" formatCode="DD.MM.YYYY"/>'), '160 survives verbatim')
    assert.ok(out.includes('<numFmt numFmtId="161"'), '161 survives')

    const reparsed = new Stylesheet(out)
    assert.equal(reparsed.describeNumberFormatCode(160), 'DD.MM.YYYY')
    assert.equal(reparsed.describeCellXf(1).numberFormatCode, 'DD.MM.YYYY', 'the date cell keeps its format')
    assert.equal(reparsed.describeCellXf(added).numberFormatCode, 'DD.MM.YYYY',
      'the new format inherits the base number format')

    const fresh = new Stylesheet(xml)
    assert.ok(fresh.ensureNumFmt({ type: 'custom', code: '0.000' }) > 161,
      'a new id never collides with an id the workbook already uses')
  })

  test('estimateColumnWidth grows with the content', () => {
    assert.equal(estimateColumnWidth([]), 8, 'an empty column keeps the floor width')
    assert.ok(estimateColumnWidth([{ text: 'x'.repeat(30) }]) >= 30)
    assert.ok(estimateColumnWidth([{ text: 'Доходы за март' }]) > estimateColumnWidth([{ text: 'Дата' }]))
  })
})

describe('XLSX formatting', () => {
  before(async () => {
    await engine.create(path.join(tmpDir, 'base.xlsx'), {
      sheets: [
        {
          name: 'Данные',
          data: [
            ['Категория', 'Сумма', 'Дата'],
            ['Маркетинг', 50000, ''],
            ['Офис', 30000, '']
          ]
        },
        { name: 'Прочее', data: [['нетронуто', 7]] },
        { name: 'Пустой', data: [] }
      ]
    })
    // A formula cell that no engine has calculated yet.
    await engine.write(path.join(tmpDir, 'base.xlsx'), {
      sheetName: 'Данные',
      cells: [{ ref: 'C2', formula: '=B2*2' }]
    })
  })

  describe('font', () => {
    test('family, size and bold are applied and read back', async () => {
      const book = copyTo('font.xlsx')
      const res = await engine.format(book, {
        sheetName: 'Данные',
        range: 'B1',
        font: { family: 'Arial', size: 14, bold: true }
      })
      assert.equal(res.success, true)
      assert.equal(res.cellsFormatted, 1)

      const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C1', includeStyles: true })
      const font = cell(read, 'B1').style.font
      assert.equal(font.family, 'Arial')
      assert.equal(font.size, 14)
      assert.equal(font.bold, true)
      assert.equal(cell(read, 'A1').style.font.bold, false, 'a neighbouring cell is not affected')
    })

    test('italic, underline and strike are applied and read back', async () => {
      const book = copyTo('font-style.xlsx')
      await engine.format(book, {
        sheetName: 'Данные',
        range: 'A2:A3',
        font: { italic: true, underline: true, strike: true }
      })

      const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C3', includeStyles: true })
      const font = cell(read, 'A2').style.font
      assert.equal(font.italic, true)
      assert.equal(font.underline, 'single')
      assert.equal(font.strike, true)
      assert.equal(cell(read, 'A3').style.font.italic, true, 'every cell of the range got the format')
      assert.equal(cell(read, 'A1').style.font.italic, false, 'the row outside the range did not')

      await engine.format(book, { sheetName: 'Данные', range: 'B2', font: { underline: 'double' } })
      const double = await engine.read(book, { sheetName: 'Данные', range: 'A1:C3', includeStyles: true })
      assert.equal(cell(double, 'B2').style.font.underline, 'double')
    })

    test('font colour is stored as ARGB', async () => {
      const book = copyTo('font-color.xlsx')
      await engine.format(book, { sheetName: 'Данные', range: 'B1', font: { color: '#C00000' } })

      const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C1', includeStyles: true })
      assert.equal(cell(read, 'B1').style.font.color, normalizeArgb('#C00000'))
      assert.equal(cell(read, 'B1').style.font.color, 'FFC00000')
    })
  })

  describe('fill', () => {
    test('a background colour is applied and read back', async () => {
      const book = copyTo('fill.xlsx')
      await engine.format(book, { sheetName: 'Данные', range: 'A1:C1', fill: { color: '#FFF2CC' } })

      const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C3', includeStyles: true })
      assert.equal(cell(read, 'A1').style.fill.patternType, 'solid')
      assert.equal(cell(read, 'A1').style.fill.color, 'FFFFF2CC')
      assert.equal(cell(read, 'C1').style.fill.color, 'FFFFF2CC', 'the whole range is filled')
      assert.equal(cell(read, 'A2').style.fill, null, 'the row below is not')
      const stylesXml = (await ZipArchive.fromFile(book)).getText('xl/styles.xml')
      assert.ok(stylesXml.includes('<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/>'),
        'the fill is declared in the stylesheet')
    })
  })

  describe('borders', () => {
    test('each edge is stored with its own style and colour', async () => {
      const book = copyTo('border-edges.xlsx')
      await engine.format(book, {
        sheetName: 'Данные',
        range: 'A1:C1',
        border: { top: 'thin', bottom: { style: 'double', color: '#FF0000' } }
      })

      const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C1', includeStyles: true })
      const border = cell(read, 'A1').style.border
      assert.equal(border.top.style, 'thin')
      assert.equal(border.top.color, 'FF000000', 'an unspecified border colour defaults to black')
      assert.equal(border.bottom.style, 'double')
      assert.equal(border.bottom.color, 'FFFF0000')
      assert.equal(border.left, undefined, 'an edge nobody asked for stays absent')
      assert.equal(border.right, undefined)
    })

    test('"all" applies one style to every edge', async () => {
      const book = copyTo('border-all.xlsx')
      await engine.format(book, {
        sheetName: 'Данные',
        range: 'A2:B2',
        border: { all: { style: 'medium', color: '#336699' } }
      })

      const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C3', includeStyles: true })
      const border = cell(read, 'A2').style.border
      for (const edge of ['left', 'right', 'top', 'bottom']) {
        assert.equal(border[edge].style, 'medium', `${edge} is medium`)
        assert.equal(border[edge].color, 'FF336699', `${edge} carries the requested colour`)
      }
    })

    test('adding one edge keeps the edges a cell already had', async () => {
      const book = copyTo('border-merge.xlsx')
      await engine.format(book, { sheetName: 'Данные', range: 'B2', border: { all: 'thin' } })
      await engine.format(book, { sheetName: 'Данные', range: 'B2', border: { bottom: 'thick' } })

      const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C3', includeStyles: true })
      const border = cell(read, 'B2').style.border
      assert.equal(border.bottom.style, 'thick', 'the new edge wins')
      assert.equal(border.top.style, 'thin', 'the untouched edges survive')
      assert.equal(border.left.style, 'thin')
      assert.equal(border.right.style, 'thin')
    })
  })

  describe('alignment', () => {
    test('horizontal, vertical and wrapText are applied and read back', async () => {
      const book = copyTo('alignment.xlsx')
      await engine.format(book, {
        sheetName: 'Данные',
        range: 'A1:C1',
        alignment: { horizontal: 'center', vertical: 'top', wrapText: true }
      })

      const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C1', includeStyles: true })
      const alignment = cell(read, 'A1').style.alignment
      assert.equal(alignment.horizontal, 'center')
      assert.equal(alignment.vertical, 'top')
      assert.equal(alignment.wrapText, true)
      assert.equal(cell(read, 'C1').style.alignment.horizontal, 'center', 'the whole range is aligned')
    })

    test('textRotation is stored', async () => {
      const book = copyTo('rotation.xlsx')
      await engine.format(book, { sheetName: 'Данные', range: 'A2', alignment: { textRotation: 45 } })

      const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C3', includeStyles: true })
      assert.equal(Number(cell(read, 'A2').style.alignment.textRotation), 45)
    })
  })

  describe('number formats', () => {
    test('every documented format is stored as a code and the value stays a number', async () => {
      const book = copyTo('number-formats.xlsx')

      await engine.write(book, {
        sheetName: 'Данные',
        cells: [
          { ref: 'A1', value: 1234.5, numberFormat: { type: 'integer' } },
          { ref: 'B1', value: 1234.5, numberFormat: { type: 'decimal', decimals: 3 } },
          { ref: 'C1', value: 1234.5, numberFormat: { type: 'currency', symbol: '₽', thousands: true } },
          { ref: 'D1', value: 0.256, numberFormat: { type: 'percent', decimals: 1 } },
          { ref: 'E1', value: 1234.5, numberFormat: { type: 'decimal' } },
          { ref: 'F1', value: 0.25, numberFormat: { type: 'percent' } },
          { ref: 'A2', value: 46096, numberFormat: { type: 'date' } },
          { ref: 'B2', value: 46096.5, numberFormat: { type: 'datetime' } },
          { ref: 'C2', value: 12.5, numberFormat: { type: 'custom', code: '0.000' } }
        ]
      })

      const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:F3', includeStyles: true })
      const xml = await sheetXmlOf(book)

      // A code the built-in table already defines reuses its id, and is reported
      // by the friendly name; anything else is declared and reported as a code.
      assert.equal(cell(read, 'A1').style.numberFormatId, 1, 'the built-in integer id is reused')
      assert.equal(cell(read, 'A1').style.numberFormat, 'integer')
      assert.equal(cell(read, 'A1').style.numberFormatCode, '0')
      assert.equal(cell(read, 'A1').style.numberFormatName, 'integer')

      assert.equal(cell(read, 'E1').style.numberFormatId, 2, 'the built-in decimal id is reused')
      assert.equal(cell(read, 'E1').style.numberFormat, 'decimal')
      assert.equal(cell(read, 'E1').style.numberFormatCode, '0.00')

      assert.equal(cell(read, 'F1').style.numberFormatId, 9, 'the built-in percent id is reused')
      assert.equal(cell(read, 'F1').style.numberFormat, 'percent')
      assert.equal(cell(read, 'F1').style.numberFormatCode, '0%')

      assert.equal(cell(read, 'B1').style.numberFormat, '0.000')
      assert.equal(cell(read, 'B1').style.numberFormatName, null, 'a custom code has no built-in name')
      assert.ok(cell(read, 'B1').style.numberFormatId >= 164, 'a custom code gets a custom id')

      assert.ok(cell(read, 'C1').style.numberFormat.includes('₽'), 'the currency code carries the symbol')
      assert.equal(cell(read, 'C1').style.numberFormatCode, cell(read, 'C1').style.numberFormat)

      assert.equal(cell(read, 'D1').style.numberFormat, '0.0%')
      assert.equal(cell(read, 'A2').style.numberFormat, 'DD.MM.YYYY')
      assert.equal(cell(read, 'B2').style.numberFormat, 'DD.MM.YYYY HH:MM:SS')
      assert.equal(cell(read, 'C2').style.numberFormat, '0.000', 'a custom code is stored verbatim')

      for (const ref of ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'A2', 'B2', 'C2']) {
        assert.equal(typeof cell(read, ref).value, 'number', `${ref} is a number, never preformatted text`)
      }

      const stylesXml = (await ZipArchive.fromFile(book)).getText('xl/styles.xml')
      const codes = [...stylesXml.matchAll(/formatCode="([^"]*)"/g)].map((m) => m[1])
      assert.ok(codes.some((c) => c.includes('₽')), `a currency numFmt is declared: ${codes.join(' | ')}`)
      assert.ok(!xml.includes('<is><t>1234'), 'numbers are not written as inline strings')
      assert.ok(!/<c r="C1"[^>]*t="inlineStr"/.test(xml), 'a formatted number is not typed as a string')
    })

    test('a number format replaces the previous one but keeps the rest of the format', async () => {
      const book = copyTo('number-format-replace.xlsx')
      await engine.format(book, { sheetName: 'Данные', range: 'B2', font: { bold: true }, fill: { color: '#DDEEFF' } })
      await engine.format(book, { sheetName: 'Данные', range: 'B2', numberFormat: { type: 'currency' } })

      const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C3', includeStyles: true })
      const style = cell(read, 'B2').style
      assert.ok(style.numberFormat.includes('₽'))
      assert.equal(style.font.bold, true, 'the font survived the number format')
      assert.equal(style.fill.color, 'FFDDEEFF', 'the fill survived the number format')
    })
  })

  describe('real dates', () => {
    test('an ISO date is stored as a serial number with a date format', async () => {
      const book = copyTo('date.xlsx')
      await engine.write(book, {
        sheetName: 'Данные',
        cells: [{ ref: 'C2', value: '2026-03-15', date: true }]
      })

      const xml = await sheetXmlOf(book)
      assert.ok(!xml.includes('2026-03-15'), 'the date is never stored as text')
      assert.ok(xml.includes(`<v>${dateToSerial('2026-03-15')}</v>`), 'the serial number is stored')

      const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C3', includeStyles: true })
      assert.equal(cell(read, 'C2').value, 46096)
      assert.equal(typeof cell(read, 'C2').value, 'number')
      assert.equal(cell(read, 'C2').style.numberFormat, 'DD.MM.YYYY')
      assert.equal(serialToDate(cell(read, 'C2').value).date, '2026-03-15', 'the serial round-trips')
    })

    test('a value with a time part gets a datetime format', async () => {
      const book = copyTo('datetime.xlsx')
      const iso = '2026-03-15T08:30:00Z'
      await engine.write(book, {
        sheetName: 'Данные',
        cells: [{ ref: 'C2', value: iso, date: true }]
      })

      const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C3', includeStyles: true })
      assert.equal(cell(read, 'C2').style.numberFormat, 'DD.MM.YYYY HH:MM:SS')
      assert.ok(Math.abs(cell(read, 'C2').value - dateToSerial(iso)) < 1e-9)
      assert.equal(serialToDate(cell(read, 'C2').value).datetime, '2026-03-15T08:30:00Z')
      assert.ok(!(await sheetXmlOf(book)).includes('08:30'), 'the time is not stored as text')
    })

    test('a caller-supplied format wins over the date default', async () => {
      const book = copyTo('date-format-override.xlsx')
      await engine.write(book, {
        sheetName: 'Данные',
        cells: [{ ref: 'C2', value: '2026-03-15', date: true, numberFormat: { type: 'custom', code: 'YYYY-MM-DD' } }]
      })

      const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C3', includeStyles: true })
      assert.equal(cell(read, 'C2').style.numberFormat, 'YYYY-MM-DD')
      assert.equal(cell(read, 'C2').value, 46096)
    })
  })

  describe('merges, widths and heights', () => {
    test('a range can be merged and unmerged', async () => {
      const book = copyTo('merge.xlsx')
      const merged = await engine.format(book, { sheetName: 'Данные', range: 'A1:C1', merge: true })
      assert.deepEqual(merged.merges, ['A1:C1'])

      const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C1', includeStyles: true })
      assert.deepEqual(read.merged, ['A1:C1'])

      const xml = await sheetXmlOf(book)
      assert.ok(xml.includes('<mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells>'))
      assert.ok(xml.indexOf('</sheetData>') < xml.indexOf('<mergeCells count'), 'mergeCells follows sheetData')

      // Unmerging through a single cell inside the block removes the block.
      const unmerged = await engine.format(book, { sheetName: 'Данные', range: 'B1', unmerge: true })
      assert.deepEqual(unmerged.merges, [])
      const after = await engine.read(book, { sheetName: 'Данные', range: 'A1:C1', includeStyles: true })
      assert.deepEqual(after.merged, [])
      assert.ok(!(await sheetXmlOf(book)).includes('mergeCell'))
    })

    test('unmerging leaves merges the range does not touch', async () => {
      const book = copyTo('merge-keep.xlsx')
      await engine.format(book, { sheetName: 'Данные', range: 'A1:C1', merge: true })
      await engine.format(book, { sheetName: 'Данные', range: 'A3:C3', merge: true })
      const res = await engine.format(book, { sheetName: 'Данные', range: 'A1:C1', unmerge: true })
      assert.deepEqual(res.merges, ['A3:C3'])
    })

    test('column widths and row heights are applied to the whole range', async () => {
      const book = copyTo('dimensions.xlsx')
      const res = await engine.format(book, {
        sheetName: 'Данные',
        range: 'A1:C1',
        columnWidth: { width: 25 },
        rowHeight: 30
      })

      assert.deepEqual(res.columnWidths, [{ min: 1, max: 3, width: 25 }])
      assert.deepEqual(res.rowHeights, { 1: 30 })

      const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C1', includeStyles: true })
      assert.deepEqual(read.columnWidths, [{ min: 1, max: 3, width: 25 }])
      assert.equal(read.rowHeights[1], 30)
      assert.equal(read.rowHeights[2], undefined, 'only the rows of the range are set')
    })

    test('every row of a multi-row range gets the height', async () => {
      const book = copyTo('heights.xlsx')
      const res = await engine.format(book, { sheetName: 'Данные', range: 'A2:C3', rowHeight: 22.5 })
      assert.deepEqual(res.rowHeights, { 2: 22.5, 3: 22.5 })
    })

    test('{ auto: true } estimates a width from the cell text', async () => {
      const book = copyTo('auto-width.xlsx')
      const res = await engine.format(book, {
        sheetName: 'Данные',
        range: 'A1:A3',
        columnWidth: { auto: true }
      })

      assert.equal(res.columnWidths.length, 1)
      assert.ok(res.columnWidths[0].width >= 10,
        `"Категория" needs more than the 8-character floor, got ${res.columnWidths[0].width}`)
      assert.ok(res.columnWidths[0].width < 255)

      const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C1', includeStyles: true })
      assert.equal(read.columnWidths[0].min, 1)
      assert.equal(read.columnWidths[0].max, 1, 'only the addressed column was widened')
    })
  })

  describe('additive behaviour', () => {
    test('pre-existing content and its formatting are unchanged', async () => {
      const book = copyTo('preexisting.xlsx')
      await engine.format(book, { sheetName: 'Данные', range: 'B2', font: { bold: true, color: '#00B050' } })

      const before = await engine.read(book, {
        sheetName: 'Данные', range: 'A1:C3', includeStyles: true, includeFormulas: true
      })

      await engine.format(book, {
        sheetName: 'Данные', range: 'A1:C1', fill: { color: '#DDEEFF' }, border: { all: 'thin' }
      })

      const after = await engine.read(book, {
        sheetName: 'Данные', range: 'A1:C3', includeStyles: true, includeFormulas: true
      })

      assert.deepEqual(after.data, before.data, 'no value moved')
      assert.deepEqual(after.formulas, before.formulas, 'no formula changed')
      assert.deepEqual(cell(after, 'B2').style, cell(before, 'B2').style, 'the unrelated format is untouched')
      assert.deepEqual(cell(after, 'A2').style, cell(before, 'A2').style)
      assert.deepEqual(cell(after, 'C2').style, cell(before, 'C2').style)

      const other = await engine.read(book, { sheetName: 'Прочее', range: 'A1:B1' })
      assert.deepEqual(other.data[0], ['нетронуто', 7], 'another worksheet is untouched')
    })

    test('the original stylesheet entries keep their index and their XML', async () => {
      const before = await ZipArchive.fromFile(path.join(tmpDir, 'base.xlsx'))
      const book = copyTo('styles-preserved.xlsx')
      await engine.format(book, {
        sheetName: 'Данные',
        range: 'A1:C1',
        font: { bold: true, size: 13, color: '#123456' },
        fill: { color: '#ABCDEF' },
        border: { all: 'thin' },
        alignment: { horizontal: 'center' },
        numberFormat: { type: 'decimal', decimals: 4 }
      })

      const after = await ZipArchive.fromFile(book)
      const was = new Stylesheet(before.getText('xl/styles.xml'))
      const now = new Stylesheet(after.getText('xl/styles.xml'))

      for (const table of ['fonts', 'fills', 'borders', 'cellXfs', 'cellStyleXfs', 'numFmts']) {
        assert.ok(now[table].length >= was[table].length, `${table} only grew`)
        for (let i = 0; i < was[table].length; i++) {
          assert.equal(now[table][i].raw, was[table][i].raw,
            `${table}[${i}] is byte-identical, so every cell that refers to it looks the same`)
        }
      }

      // Format 0 is what every unstyled cell in the workbook points at.
      assert.equal(now.describeCellXf(0).index, 0)
      assert.equal(now.describeCellXf(0).font.bold, false)
    })

    test('untouched package members stay byte-identical', async () => {
      const before = await ZipArchive.fromFile(path.join(tmpDir, 'base.xlsx'))
      const book = copyTo('members.xlsx')

      await engine.format(book, {
        sheetName: 'Данные',
        range: 'A1:C1',
        font: { bold: true, size: 14, color: '#FFFFFF' },
        fill: { color: '#4F81BD' },
        border: { all: 'thin' },
        alignment: { horizontal: 'center', wrapText: true },
        merge: true,
        columnWidth: { width: 22 },
        rowHeight: 28
      })
      await engine.format(book, { sheetName: 'Данные', range: 'B2:B3', numberFormat: { type: 'currency' } })

      const after = await ZipArchive.fromFile(book)

      // A workbook without the R7 template has no styles part until the first
      // format call creates one; that is the only member allowed to appear.
      const expected = before.has('xl/styles.xml') ? before.list() : [...before.list(), 'xl/styles.xml']
      assert.deepEqual(after.list(), expected, 'no member was added other than the styles part')

      const edited = new Set(['xl/worksheets/sheet1.xml', 'xl/styles.xml'])
      if (!before.has('xl/styles.xml')) {
        // Without the R7 template the package starts with no stylesheet at all,
        // so the first format call also has to DECLARE the part it creates: the
        // content types and the workbook relationships necessarily change with
        // it. A templated workbook already names both of them, so there only the
        // worksheet and the stylesheet move.
        edited.add('[Content_Types].xml')
        edited.add('xl/_rels/workbook.xml.rels')
      }
      for (const name of before.list()) {
        assert.ok(after.has(name), `${name} is still there`)
        if (edited.has(name)) continue
        assert.equal(Buffer.compare(before.getBuffer(name), after.getBuffer(name)), 0,
          `${name} is byte-identical`)
      }
      assert.notEqual(
        Buffer.compare(before.getBuffer('xl/worksheets/sheet1.xml'), after.getBuffer('xl/worksheets/sheet1.xml')),
        0,
        'the edited worksheet is the one part that did change'
      )
    })

    test('formatting a formula cell keeps the formula and the cached value', async () => {
      const book = copyTo('formula.xlsx')
      await engine.format(book, {
        sheetName: 'Данные',
        range: 'C2',
        fill: { color: '#FCE4D6' },
        font: { bold: true },
        numberFormat: { type: 'decimal', decimals: 2 }
      })

      const read = await engine.read(book, {
        sheetName: 'Данные', range: 'A1:C3', includeFormulas: true, includeStyles: true
      })
      assert.equal(cell(read, 'C2').formula, 'B2*2', 'the formula survives')
      assert.equal(cell(read, 'C2').style.fill.color, 'FFFCE4D6')
      assert.equal(cell(read, 'C2').style.numberFormat, 'decimal')
      assert.equal(cell(read, 'C2').style.numberFormatCode, '0.00')

      const xml = await sheetXmlOf(book)
      assert.ok(xml.includes('<f>B2*2</f>'))
      assert.ok(/<c r="C2"[^>]*s="\d+"><f>B2\*2<\/f>/.test(xml),
        'the formula cell carries the format and nothing else changed')
    })

    test('a merged title keeps the value of its top-left cell', async () => {
      const book = copyTo('merged-title.xlsx')
      await engine.write(book, { sheetName: 'Данные', cells: [{ ref: 'A1', value: 'Итоговый отчёт' }] })
      await engine.format(book, {
        sheetName: 'Данные',
        range: 'A1:C1',
        merge: true,
        font: { bold: true, size: 16 },
        fill: { color: '#D9E2F3' },
        alignment: { horizontal: 'center' }
      })

      const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C1', includeStyles: true })
      assert.equal(cell(read, 'A1').value, 'Итоговый отчёт')
      assert.deepEqual(read.merged, ['A1:C1'])
      assert.equal(cell(read, 'A1').style.font.bold, true)
      assert.equal(cell(read, 'A1').style.alignment.horizontal, 'center')
    })

    test('a styled but empty cell exists and carries its format', async () => {
      const book = copyTo('empty-styled.xlsx')
      await engine.format(book, {
        sheetName: 'Пустой',
        range: 'A1:B2',
        font: { bold: true },
        alignment: { horizontal: 'center' }
      })

      const read = await engine.read(book, { sheetName: 'Пустой', range: 'A1:B2', includeStyles: true })
      assert.equal(read.data[0][0], '')
      assert.equal(cell(read, 'A1').style.font.bold, true)
      assert.equal(cell(read, 'B2').style.font.bold, true)
      assert.ok((await sheetXmlOf(book, (await engine.inspect(book)).sheets[2].sheetPath))
        .includes('<c r="B2" s="'))
    })
  })

  describe('the read side', () => {
    test('an existing cell without a style is the default, a missing cell is null', async () => {
      const book = copyTo('read-styles.xlsx')
      const read = await engine.read(book, { sheetName: 'Прочее', range: 'A1:D1', includeStyles: true })

      assert.equal(read.styles[0][0].index, 0, 'a cell with no s attribute is cell format 0')
      assert.equal(read.styles[0][0].numberFormat, 'general')
      assert.equal(read.styles[0][0].numberFormatName, 'general')
      assert.equal(read.styles[0][0].numberFormatCode, 'General')
      assert.equal(read.styles[0][3], null, 'there is no cell in D1')
      assert.equal(read.data[0][3], '')
      assert.ok(read.styleVocabulary.numberFormatExamples.includes('currency'))
    })

    test('the sheet name is echoed only when the caller used it', async () => {
      const book = copyTo('read-name.xlsx')
      const byName = await engine.read(book, { sheetName: 'Данные', range: 'A1' })
      assert.equal(byName.sheetName, 'Данные')

      const byIndex = await engine.read(book, { sheetIndex: 0, range: 'A1' })
      assert.equal(byIndex.sheetName, undefined)
      assert.equal(byIndex.sheet, byName.sheet, 'both address the same part')
    })

    test('includeStyles reports the whole sheet picture', async () => {
      const book = copyTo('read-picture.xlsx')
      await engine.format(book, {
        sheetName: 'Данные', range: 'A1:C1', merge: true, rowHeight: 20, columnWidth: { width: 18 }
      })

      const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C3', includeStyles: true })
      assert.deepEqual(read.merged, ['A1:C1'])
      assert.deepEqual(read.rowHeights, { 1: 20 })
      assert.deepEqual(read.columnWidths, [{ min: 1, max: 3, width: 18 }])
      assert.equal(read.styles.length, 3, 'one style row per data row')
      assert.equal(read.styles[0].length, 3, 'one style per column')

      const plain = await engine.read(book, { sheetName: 'Данные', range: 'A1:C3' })
      assert.equal(plain.styles, undefined, 'no styles are reported unless asked for')
    })
  })

  describe('rejecting nonsense', () => {
    test('a missing or malformed range fails loudly', async () => {
      const book = copyTo('errors.xlsx')
      await assert.rejects(() => engine.format(book, { sheetName: 'Данные' }), /requires a range/)
      await assert.rejects(() => engine.format(book, { sheetName: 'Данные', range: '' }), /requires a range/)
      await assert.rejects(() => engine.format(book, { sheetName: 'Данные', range: 'ZZ' }), /Invalid cell reference/)
      await assert.rejects(() => engine.format(book, { sheetName: 'Данные', range: 'C3:A1' }), /must not precede/)
    })

    test('an unknown sheet and an impossible width fail loudly', async () => {
      const book = copyTo('errors2.xlsx')
      await assert.rejects(
        () => engine.format(book, { sheetName: 'Нет такого листа', range: 'A1' }),
        /Sheet not found/
      )
      await assert.rejects(
        () => engine.format(book, { sheetName: 'Данные', range: 'A1', columnWidth: {} }),
        /columnWidth requires/
      )
      await assert.rejects(
        () => engine.format(book, { sheetName: 'Данные', range: 'A1', rowHeight: 0 }),
        /Invalid row height/
      )
      await assert.rejects(
        () => engine.format(book, { sheetName: 'Данные', range: 'A1', numberFormat: { type: 'wat' } }),
        /Unsupported number format/
      )
    })

    test('a failed format leaves the file untouched', async () => {
      const book = copyTo('errors3.xlsx')
      const before = fs.readFileSync(book)
      await assert.rejects(() => engine.format(book, { sheetName: 'Данные', range: 'A1', columnWidth: {} }))
      assert.equal(Buffer.compare(before, fs.readFileSync(book)), 0, 'the workbook on disk did not change')
    })
  })

  describe('R7 round trip', () => {
    test('R7 opens and renders a formatted workbook', async (t) => {
      if (requiresR7(t)) return

      const book = copyTo('r7-render.xlsx')
      await engine.write(book, {
        sheetName: 'Данные',
        cells: [
          { ref: 'A1', value: 'Отчёт за март 2026' },
          { ref: 'C2', value: '2026-03-15', date: true },
          { ref: 'C3', value: '2026-03-16', date: true }
        ]
      })
      await engine.format(book, {
        sheetName: 'Данные',
        range: 'A1:C1',
        merge: true,
        font: { bold: true, size: 15, color: '#FFFFFF' },
        fill: { color: '#4472C4' },
        alignment: { horizontal: 'center' },
        rowHeight: 26
      })
      await engine.format(book, {
        sheetName: 'Данные',
        range: 'B2:B3',
        numberFormat: { type: 'currency', symbol: '₽', thousands: true },
        border: { all: 'thin' },
        alignment: { horizontal: 'right' }
      })
      await engine.format(book, { sheetName: 'Данные', range: 'A1:C3', columnWidth: { width: 20 } })

      const pdf = path.join(tmpDir, 'r7-render.pdf')
      await adapter.convert(book, pdf)
      assert.ok(looksLikePdf(pdf), 'x2t rendered the formatted workbook')
      assert.ok(fs.statSync(pdf).size > 500)
    })
  })
})

describe('write: formats and real dates', () => {
  before(async () => {
    if (!fs.existsSync(path.join(tmpDir, 'base.xlsx'))) {
      await engine.create(path.join(tmpDir, 'base.xlsx'), {
        sheets: [{ name: 'Данные', data: [['Категория', 'Сумма'], ['Маркетинг', 50000], ['Офис', 30000]] }]
      })
    }
  })

  test('write applies a per-cell number format and keeps the value numeric', async () => {
    const book = copyTo('write-format.xlsx')
    const res = await engine.write(book, {
      sheetName: 'Данные',
      cells: [
        { ref: 'B1', value: 1234.5, numberFormat: { type: 'currency', symbol: '₽' } },
        { ref: 'B2', value: 1234.5, numberFormat: { type: 'integer' } }
      ]
    })
    assert.equal(res.updatedCells, 2)

    const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C3', includeStyles: true })
    assert.equal(cell(read, 'B1').value, 1234.5)
    assert.ok(cell(read, 'B1').style.numberFormat.includes('₽'))
    assert.equal(cell(read, 'B2').style.numberFormat, 'integer')
    assert.equal(cell(read, 'B1').value, 1234.5)
  })

  test('a date written with date: true becomes a serial, not text', async () => {
    const book = copyTo('write-date.xlsx')
    await engine.write(book, {
      sheetName: 'Данные',
      cells: [{ ref: 'C1', value: '2026-12-31', date: true }]
    })

    const xml = await sheetXmlOf(book)
    assert.ok(!xml.includes('2026-12-31'))
    const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C1', includeStyles: true })
    assert.equal(cell(read, 'C1').value, dateToSerial('2026-12-31'))
    assert.equal(serialToDate(cell(read, 'C1').value).date, '2026-12-31')
    assert.equal(cell(read, 'C1').style.numberFormat, 'DD.MM.YYYY')
  })

  test('write keeps the format a cell already had', async () => {
    const book = copyTo('write-keeps-style.xlsx')
    await engine.format(book, {
      sheetName: 'Данные', range: 'A2', font: { bold: true, color: '#0070C0' }, fill: { color: '#FFF2CC' }
    })
    const formatted = await engine.read(book, { sheetName: 'Данные', range: 'A1:C3', includeStyles: true })

    await engine.write(book, { sheetName: 'Данные', cells: [{ ref: 'A2', value: 'Обновлено' }] })

    const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C3', includeStyles: true })
    assert.equal(cell(read, 'A2').value, 'Обновлено')
    assert.deepEqual(cell(read, 'A2').style, cell(formatted, 'A2').style, 'the format survived the write')
  })

  test('writing an empty value clears the payload but not the format', async () => {
    const book = copyTo('write-empty.xlsx')
    await engine.format(book, { sheetName: 'Данные', range: 'B2', font: { bold: true } })
    await engine.write(book, { sheetName: 'Данные', cells: [{ ref: 'B2' }] })

    const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C3', includeStyles: true })
    assert.equal(cell(read, 'B2').value, '')
    assert.equal(cell(read, 'B2').style.font.bold, true)
  })

  test('writing into an empty styled cell does not swallow the next row', async () => {
    const book = copyTo('write-selfclosing.xlsx')
    // Formatting a blank cell creates the element R7 writes as `<c r="D1" s="N"/>`.
    await engine.format(book, { sheetName: 'Данные', range: 'D1', font: { bold: true } })
    assert.ok((await sheetXmlOf(book)).includes('<c r="D1" s="'))
    assert.ok((await sheetXmlOf(book)).includes('<c r="D1" s="')
      && /<c r="D1" s="\d+"\/>/.test(await sheetXmlOf(book)), 'the empty cell is self-closing')

    await engine.write(book, { sheetName: 'Данные', cells: [{ ref: 'D1', value: 'X' }] })

    const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:D3' })
    assert.equal(read.data[0][3], 'X')
    assert.deepEqual(read.data[1], ['Маркетинг', 50000, '', ''], 'row 2 is intact')
    assert.deepEqual(read.data[2], ['Офис', 30000, '', ''], 'row 3 is intact')
  })

  test('a formula is written without a stale type attribute', async () => {
    const book = copyTo('write-formula.xlsx')
    await engine.write(book, { sheetName: 'Данные', cells: [{ ref: 'C1', formula: '=SUM(B1:B2)', numberFormat: { type: 'currency' } }] })

    const xml = await sheetXmlOf(book)
    assert.ok(xml.includes('<f>SUM(B1:B2)</f>'))
    assert.ok(!/<c r="C1"[^>]*t="inlineStr"/.test(xml), 'a formula cell is not typed as an inline string')

    const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C1', includeFormulas: true, includeStyles: true })
    assert.equal(cell(read, 'C1').formula, 'SUM(B1:B2)')
    assert.ok(cell(read, 'C1').style.numberFormat.includes('₽'))
  })

  test('a formula with a cached value keeps both', async () => {
    const book = copyTo('write-formula-value.xlsx')
    await engine.write(book, { sheetName: 'Данные', cells: [{ ref: 'C2', formula: '=B2*2', value: 100000 }] })

    const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:C2', includeFormulas: true })
    assert.equal(cell(read, 'C2').formula, 'B2*2')
    assert.equal(cell(read, 'C2').value, 100000)
  })

  test('writing several cells of one call keeps every cell independent', async () => {
    const book = copyTo('write-many.xlsx')
    await engine.write(book, {
      sheetName: 'Данные',
      cells: [
        { ref: 'E1', value: 'итог' },
        { ref: 'C2', value: '2026-01-01', date: true },
        { ref: 'C3', value: 0.5, numberFormat: { type: 'percent' } },
        { ref: 'A1', value: 'Обновлено' }
      ]
    })

    const read = await engine.read(book, { sheetName: 'Данные', range: 'A1:E3', includeStyles: true })
    assert.equal(cell(read, 'A1').value, 'Обновлено')
    assert.equal(cell(read, 'E1').value, 'итог')
    assert.equal(cell(read, 'C2').value, dateToSerial('2026-01-01'))
    assert.equal(cell(read, 'C3').value, 0.5)
    assert.equal(cell(read, 'C3').style.numberFormat, 'percent')
    assert.equal(cell(read, 'C3').style.numberFormatCode, '0%')
    assert.equal(cell(read, 'A2').value, 'Маркетинг', 'the untouched rows survived')
  })

  test('R7 accepts a workbook written with dates and formats', async (t) => {
    if (requiresR7(t)) return
    const book = copyTo('write-r7.xlsx')
    await engine.write(book, {
      sheetName: 'Данные',
      cells: [
        { ref: 'A1', value: 'Дата' },
        { ref: 'B1', value: 'Сумма' },
        { ref: 'A2', value: '2026-02-01', date: true },
        { ref: 'B2', value: 1999.9, numberFormat: { type: 'currency', symbol: '₽' } }
      ]
    })
    const pdf = path.join(tmpDir, 'write-r7.pdf')
    await adapter.convert(book, pdf)
    assert.ok(looksLikePdf(pdf))
  })
})

describe('workbook without the R7 template', () => {
  // The synthetic-package path, exercised ON PURPOSE on every host: the stub
  // adapter answers `getTemplatePath` with null, so `create` builds the package
  // itself even where R7-Office is installed. Nothing here may be gated on R7 —
  // this is the group that keeps the no-template fallback honest.
  const stubAdapter = { getTemplatePath: async () => null }
  const standalone = new XlsxEngine(stubAdapter)
  const dir = path.join(tmpDir, 'standalone')

  before(() => {
    fs.mkdirSync(dir, { recursive: true })
  })

  test('the hand-built package still gets a registered styles part', async () => {
    const book = path.join(dir, 'fallback.xlsx')
    await standalone.create(book, { sheets: [{ name: 'Лист1', data: [['A', 1]] }] })

    const zipBefore = await ZipArchive.fromFile(book)
    assert.ok(!zipBefore.has('xl/styles.xml'), 'the fallback package starts without a stylesheet')

    await standalone.format(book, {
      range: 'A1:B1',
      font: { bold: true, color: '#FF0000' },
      fill: { color: '#FFFF00' },
      border: { all: 'thin' }
    })
    await standalone.write(book, { cells: [{ ref: 'C1', value: '2026-01-02', date: true }] })

    const zip = await ZipArchive.fromFile(book)
    assert.ok(zip.has('xl/styles.xml'), 'the stylesheet was created')
    assert.ok(zip.getText('[Content_Types].xml').includes('PartName="/xl/styles.xml"'))
    assert.ok(zip.getText('xl/_rels/workbook.xml.rels').includes('relationships/styles'))

    const validation = await standalone.validate(book)
    assert.equal(validation.valid, true, JSON.stringify(validation.errors))

    const read = await standalone.read(book, { range: 'A1:C1', includeStyles: true })
    assert.equal(read.styles[0][0].font.bold, true)
    assert.equal(read.styles[0][0].fill.color, 'FFFFFF00')
    assert.equal(read.styles[0][2].numberFormat, 'DD.MM.YYYY')
    assert.equal(read.data[0][2], dateToSerial('2026-01-02'))
  })

  test('page setup round-trips through setPageSetup, read and inspect', async () => {
    const book = path.join(dir, 'layout.xlsx')
    await standalone.create(book, {
      sheets: [
        { name: 'Данные', data: [['Категория', 'Сумма'], ['Маркетинг', 50000]] },
        { name: 'Прочее', data: [['нетронуто', 7]] }
      ]
    })
    const cellsBefore = await standalone.read(book, { sheetName: 'Данные', range: 'A1:B2' })

    const result = await standalone.setPageSetup(book, {
      sheetName: 'Данные',
      orientation: 'landscape',
      fitToWidth: 1,
      fitToHeight: 0,
      paperSize: 'A4',
      margins: { left: 0.5, right: 0.5, top: 0.6, bottom: 0.6 },
      printArea: 'A1:B2'
    })

    assert.equal(result.pageSetup.fitToPage, true, 'the sheet gained the fitToPage switch')
    assert.equal(result.pageSetup.orientation, 'landscape')
    assert.equal(result.pageSetup.paperSizeName, 'A4')
    assert.equal(result.printArea, 'Данные!$A$1:$B$2')

    const read = await standalone.read(book, { sheetName: 'Данные', includeStyles: true })
    assert.equal(read.pageSetup.orientation, 'landscape')
    assert.equal(read.pageSetup.fitToWidth, 1)
    assert.equal(read.pageSetup.margins.left, 0.5)
    assert.equal(read.printArea, 'Данные!$A$1:$B$2')
    assert.deepEqual((await standalone.read(book, { sheetName: 'Данные', range: 'A1:B2' })).data, cellsBefore.data,
      'the page setup did not move a single value')

    const info = await standalone.inspect(book)
    const sheet = info.sheets.find((s) => s.name === 'Данные')
    assert.equal(sheet.pageSetup.paperSize, 9)
    assert.equal(sheet.printArea, 'Данные!$A$1:$B$2')
    assert.equal(info.sheets.find((s) => s.name === 'Прочее').printArea, null,
      'a print area on one sheet is not reported for another')

    const validation = await standalone.validate(book)
    assert.equal(validation.valid, true, JSON.stringify(validation.errors))
  })

  test('an edit to the hand-built package leaves every other member byte-identical', async () => {
    const base = path.join(dir, 'preserve-base.xlsx')
    await standalone.create(base, {
      sheets: [
        { name: 'Лист1', data: [['A', 1]] },
        { name: 'Лист2', data: [['B', 2]] }
      ]
    })
    const before = await ZipArchive.fromFile(base)
    assert.ok(!before.has('xl/styles.xml'), 'the synthetic package starts with no stylesheet')

    const book = path.join(dir, 'preserve.xlsx')
    fs.copyFileSync(base, book)
    await standalone.format(book, {
      range: 'A1:B1', font: { bold: true }, fill: { color: '#FFFF00' }, merge: true, rowHeight: 22
    })

    const after = await ZipArchive.fromFile(book)
    // Registering the stylesheet the package did not have is the only reason a
    // member other than the edited worksheet moves, and it adds exactly one.
    const changed = after.list()
      .filter((name) => !before.has(name) || !after.getBuffer(name).equals(before.getBuffer(name)))
      .sort()
    assert.deepEqual(changed, [
      '[Content_Types].xml',
      'xl/_rels/workbook.xml.rels',
      'xl/styles.xml',
      'xl/worksheets/sheet1.xml'
    ])

    for (const name of ['_rels/.rels', 'xl/workbook.xml', 'xl/worksheets/sheet2.xml']) {
      assert.ok(after.getBuffer(name).equals(before.getBuffer(name)), `${name} must stay byte-identical`)
    }

    const validation = await standalone.validate(book)
    assert.equal(validation.valid, true, JSON.stringify(validation.errors))
  })

  test('formatting works on a self-closing sheetData', async () => {
    const book = path.join(dir, 'empty.xlsx')
    await standalone.create(book, { sheets: [{ name: 'Лист1', data: [] }] })
    assert.ok((await ZipArchive.fromFile(book)).getText('xl/worksheets/sheet1.xml').includes('<sheetData/>'))

    const res = await standalone.format(book, { range: 'B2', font: { italic: true } })
    assert.equal(res.cellsFormatted, 1)

    const read = await standalone.read(book, { range: 'A1:B2', includeStyles: true })
    assert.equal(read.data[1][1], '')
    assert.equal(read.styles[1][1].font.italic, true, 'the created cell carries its format')
    assert.equal(read.styles[0][0], null, 'there is no cell in A1')

    const xml = (await ZipArchive.fromFile(book)).getText('xl/worksheets/sheet1.xml')
    assert.ok(xml.includes('<c r="B2" s="'))
    assert.ok(xml.includes('<row r="2">'))
  })

  test('R7 accepts the standalone workbook', async (t) => {
    if (requiresR7(t)) return
    const book = path.join(dir, 'fallback.xlsx')
    const pdf = path.join(dir, 'fallback.pdf')
    await adapter.convert(book, pdf)
    assert.ok(looksLikePdf(pdf), 'x2t accepted the workbook this engine built from scratch')
  })
})
