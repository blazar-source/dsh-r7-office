/**
 * XLSX print layout and the all-sheet PDF export.
 *
 * Two defects are pinned down here:
 *
 *   1. `fitToWidth` on its own does nothing — a renderer ignores it unless the
 *      sheet also carries `<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>`, and
 *      putting the print elements in the wrong place in the worksheet makes them
 *      invisible too. The layout tests therefore assert the element order, not
 *      just the presence of the values.
 *   2. R7's `x2t` writes the PDF for exactly ONE worksheet per run — whichever
 *      tab the workbook marks active. The export path renders each sheet from a
 *      scratch copy and merges the results, so the tests check the merged page
 *      count and prove page N really came from sheet N.
 *
 * The merge itself is exercised on hand-built PDFs as well, so it is covered on
 * a host with no R7 install (the conversion tests skip there).
 */

import { test, describe, before, after } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import zlib from 'node:zlib'
import { XlsxEngine } from '../../src/r7/xlsx.js'
import { R7Adapter } from '../../src/r7/adapter.js'
import { ZipArchive } from '../../src/shared/zip.js'
import { getPageSetup, setPageSetup, setTabSelected } from '../../src/r7/xlsx-worksheet.js'
import { countPdfPages, mergePdfs, parsePdfDocument, setWorkbookActiveTab } from '../../src/r7/xlsx-pdf.js'

const tmpDir = path.join(os.tmpdir(), `dsh_r7_xlsx_pdf_${Date.now()}`)
const engine = new XlsxEngine()
let r7Available = false

const SHEETS = ['Доходы', 'Расходы', 'Итоги']

before(async () => {
  fs.mkdirSync(tmpDir, { recursive: true })
  r7Available = (await new R7Adapter().detect()).installed
})

after(() => {
  try { fs.rmSync(tmpDir, { recursive: true, force: true }) } catch { /* ignore */ }
})

/** A workbook with three sheets of distinct, checkable content. */
async function buildWorkbook(filePath, sheetNames = SHEETS) {
  await engine.create(filePath, {
    overwrite: true,
    sheets: sheetNames.map((name) => ({ name, data: [] }))
  })
  for (const [i, name] of sheetNames.entries()) {
    await engine.write(filePath, {
      sheetName: name,
      cells: [
        { ref: 'A1', value: `Заголовок ${i + 1}` },
        { ref: 'A2', value: 'Дата' },
        { ref: 'B2', value: 'Категория' },
        { ref: 'C2', value: 'Сумма' },
        { ref: 'D2', value: 'Комментарий' },
        { ref: 'A3', value: 1000 + i },
        { ref: 'B3', value: `Статья ${i + 1}` },
        { ref: 'C3', value: 250000 + i },
        { ref: 'D3', value: `Примечание ${i + 1}` }
      ]
    })
  }
  return filePath
}

/** The worksheet XML of a part. */
async function sheetXmlOf(filePath, sheetPath) {
  const zip = await ZipArchive.fromFile(filePath)
  return zip.getText(sheetPath)
}

/** Member names whose bytes differ between two files. */
async function changedMembers(beforePath, afterPath) {
  const before = await ZipArchive.fromFile(beforePath)
  const after = await ZipArchive.fromFile(afterPath)
  return after.list().filter((name) => !after.getBuffer(name).equals(before.getBuffer(name)))
}

/** The index of each tag in a worksheet XML, for order assertions. */
function positions(xml, tags) {
  return tags.map((tag) => {
    const at = xml.search(new RegExp(`<${tag}(?=[\\s/>])`))
    assert.notEqual(at, -1, `expected <${tag}> in the worksheet`)
    return { tag, at }
  })
}

/**
 * A minimal PDF with one page per marker, built by hand.
 *
 * Handing the merger known input keeps its object copying and page ordering
 * testable without a converter, and `dangling` plants a reference that has no
 * target so the failure path can be checked too.
 */
function buildPdf(markers, { dangling = null } = {}) {
  const objects = new Map()
  const pageRefs = []
  let next = 3

  for (const marker of markers) {
    const pageNum = next++
    const contentNum = next++
    const stream = `BT /F1 12 Tf 20 100 Td (${marker}) Tj ET`
    objects.set(contentNum, `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`)
    const annots = dangling === null ? '' : ` /Annots [ ${dangling} 0 R ]`
    objects.set(pageNum, `<< /Type /Page /Parent 2 0 R /MediaBox [ 0 0 200 200 ] /Resources << >>${annots} /Contents ${contentNum} 0 R >>`)
    pageRefs.push(`${pageNum} 0 R`)
  }
  objects.set(1, '<< /Type /Catalog /Pages 2 0 R >>')
  objects.set(2, `<< /Type /Pages /Count ${markers.length} /Kids [ ${pageRefs.join(' ')} ] >>`)

  const parts = [Buffer.from('%PDF-1.4\n%\u00e2\u00e3\u00cf\u00d3\n', 'latin1')]
  const offsets = new Map()
  let offset = parts[0].length
  for (let num = 1; num < next; num++) {
    offsets.set(num, offset)
    const text = Buffer.from(`${num} 0 obj\n${objects.get(num)}\nendobj\n`, 'latin1')
    parts.push(text)
    offset += text.length
  }

  const xrefOffset = offset
  const entries = [`xref\n0 ${next}\n`, '0000000000 65535 f \n']
  for (let num = 1; num < next; num++) {
    entries.push(`${String(offsets.get(num)).padStart(10, '0')} 00000 n \n`)
  }
  entries.push(`trailer\n<< /Size ${next} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF\n`)
  parts.push(Buffer.from(entries.join(''), 'latin1'))
  return Buffer.concat(parts)
}

/** Each page's decoded content stream, in page order. */
function pageContentStreams(buffer) {
  const document = parsePdfDocument(buffer)
  return document.pages.map((page) => {
    const contents = page.value.map.get('/Contents')
    const refs = contents.kind === 'array' ? contents.items : [contents]
    const chunks = refs.map((ref) => {
      const value = document.objects.get(ref.num)
      assert.ok(value && value.kind === 'stream', 'a page content must be a stream')
      const filter = value.dict.map.get('/Filter')
      const filterName = filter && filter.kind === 'token' ? filter.text : null
      const raw = Buffer.from(value.data, 'latin1')
      return filterName === '/FlateDecode' ? zlib.inflateSync(raw) : raw
    })
    return Buffer.concat(chunks).toString('latin1')
  })
}

// ---------------------------------------------------------------------------
// Page setup XML
// ---------------------------------------------------------------------------

describe('Worksheet page setup', () => {
  test('fitToWidth is written together with the fitToPage switch that activates it', () => {
    const xml = '<?xml version="1.0"?><worksheet xmlns="urn:x"><dimension ref="A1:D8"/><sheetData></sheetData></worksheet>'
    const out = setPageSetup(xml, { fitToWidth: 1, fitToHeight: 0 })

    assert.match(out, /<sheetPr><pageSetUpPr fitToPage="1"\/><\/sheetPr>/,
      'fitToWidth without fitToPage="1" is silently ignored by every renderer')
    assert.match(out, /<pageSetup fitToWidth="1" fitToHeight="0"\/>/)
    // 0 means "as many pages tall as it takes", not "one page".
    assert.ok(!/fitToHeight="1"/.test(out))
  })

  test('the print elements land in worksheet schema order', () => {
    const xml = '<?xml version="1.0"?><worksheet xmlns="urn:x"><dimension ref="A1:D8"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData></sheetData><mergeCells count="1"><mergeCell ref="A1:D1"/></mergeCells></worksheet>'
    const out = setPageSetup(xml, {
      fitToWidth: 1,
      fitToHeight: 0,
      orientation: 'landscape',
      margins: { left: 0.5 },
      centerHorizontally: true
    })

    const found = positions(out, ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'sheetData', 'mergeCells', 'printOptions', 'pageMargins', 'pageSetup'])
    for (let i = 1; i < found.length; i++) {
      assert.ok(found[i - 1].at < found[i].at,
        `<${found[i].tag}> must follow <${found[i - 1].tag}> (CT_Worksheet order)`)
    }
  })

  test('an existing sheetPr keeps its children and gains pageSetUpPr last', () => {
    const xml = '<?xml version="1.0"?><worksheet xmlns="urn:x"><sheetPr><tabColor rgb="FFFF0000"/><outlinePr summaryBelow="1"/></sheetPr><dimension ref="A1"/><sheetData/></worksheet>'
    const out = setPageSetup(xml, { fitToWidth: 1 })
    assert.match(out, /<sheetPr><tabColor rgb="FFFF0000"\/><outlinePr summaryBelow="1"\/><pageSetUpPr fitToPage="1"\/><\/sheetPr>/)
  })

  test('a second page-setup call updates in place instead of adding elements', () => {
    const xml = '<?xml version="1.0"?><worksheet xmlns="urn:x"><dimension ref="A1"/><sheetData/></worksheet>'
    const once = setPageSetup(xml, { fitToWidth: 1, fitToHeight: 0, orientation: 'landscape', margins: { left: 0.5 } })
    const twice = setPageSetup(once, { fitToWidth: 2, paperSize: 'A4' })

    assert.equal((twice.match(/<pageSetup/g) || []).length, 1)
    assert.equal((twice.match(/<pageMargins/g) || []).length, 1)
    assert.equal((twice.match(/<pageSetUpPr/g) || []).length, 1)
    assert.match(twice, /fitToWidth="2"/)
    // fitToHeight was not mentioned the second time, so it keeps its old value.
    assert.match(twice, /fitToHeight="0"/)
    assert.match(twice, /orientation="landscape"/)
    assert.match(twice, /paperSize="9"/)
  })

  test('orientation, margins, paper and centring round-trip through getPageSetup', () => {
    const xml = '<?xml version="1.0"?><worksheet xmlns="urn:x"><dimension ref="A1"/><sheetData/></worksheet>'
    const out = setPageSetup(xml, {
      orientation: 'landscape',
      fitToWidth: 1,
      fitToHeight: 0,
      paperSize: 'A4',
      margins: { left: 0.5, right: 0.25, top: 0.6, bottom: 0.6 },
      centerHorizontally: true,
      centerVertically: false
    })
    const read = getPageSetup(out)

    assert.equal(read.fitToPage, true)
    assert.equal(read.orientation, 'landscape')
    assert.equal(read.fitToWidth, 1)
    assert.equal(read.fitToHeight, 0)
    assert.equal(read.paperSize, 9)
    assert.equal(read.paperSizeName, 'A4')
    assert.deepEqual(read.margins, { left: 0.5, right: 0.25, top: 0.6, bottom: 0.6, header: 0.3, footer: 0.3 },
      'pageMargins requires every attribute, so the missing ones take the defaults')
    assert.equal(read.centerHorizontally, true)
    assert.equal(read.centerVertically, null, 'false removes the attribute rather than writing 0')
  })

  test('a sheet with no page setup reports nulls, not guesses', () => {
    const read = getPageSetup('<?xml version="1.0"?><worksheet xmlns="urn:x"><sheetData/></worksheet>')
    assert.equal(read.fitToPage, null)
    assert.equal(read.orientation, null)
    assert.equal(read.fitToWidth, null)
    assert.equal(read.margins, null)
    assert.equal(read.paperSizeName, null)
  })

  test('fitToPage: false drops fitToWidth/fitToHeight rather than leaving them inert', () => {
    const xml = '<?xml version="1.0"?><worksheet xmlns="urn:x"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1"/><sheetData/><pageSetup fitToWidth="1" fitToHeight="0"/></worksheet>'
    const out = setPageSetup(xml, { fitToPage: false, fitToWidth: 1, fitToHeight: 1, scale: 80 })
    const read = getPageSetup(out)

    assert.equal(read.fitToWidth, null)
    assert.equal(read.fitToHeight, null)
    assert.equal(read.scale, 80, 'a fixed scale is what remains once fitting is off')
    assert.ok(!/pageSetUpPr/.test(out))
  })

  test('an inert request is rejected loudly', () => {
    const xml = '<?xml version="1.0"?><worksheet xmlns="urn:x"><sheetData/></worksheet>'
    assert.throws(() => setPageSetup(xml, { orientation: 'sideways' }), /Invalid orientation/)
    assert.throws(() => setPageSetup(xml, { fitToWidth: -1 }), /Invalid fitToWidth/)
    assert.throws(() => setPageSetup(xml, { fitToWidth: 1.5 }), /Invalid fitToWidth/)
    assert.throws(() => setPageSetup(xml, { scale: 5 }), /Invalid scale/)
    assert.throws(() => setPageSetup(xml, { paperSize: 'A99' }), /Unknown paper size/)
    assert.throws(() => setPageSetup(xml, { margins: { left: -1 } }), /Invalid margin/)
  })

  test('setTabSelected marks one tab and clears the others', () => {
    const xml = '<?xml version="1.0"?><worksheet xmlns="urn:x"><dimension ref="A1"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetData/></worksheet>'
    assert.match(setTabSelected(xml, false), /<sheetView workbookViewId="0"\/>/)
    assert.match(setTabSelected(xml, true), /<sheetView tabSelected="1" workbookViewId="0"\/>/)
    assert.equal(setTabSelected(xml, false).includes('tabSelected'), false)

    const bare = '<?xml version="1.0"?><worksheet xmlns="urn:x"><dimension ref="A1"/><sheetData/></worksheet>'
    assert.match(setTabSelected(bare, true), /<dimension ref="A1"\/><sheetViews><sheetView tabSelected="1" workbookViewId="0"\/><\/sheetViews>/)
    assert.equal(setTabSelected(bare, false), bare, 'nothing to deselect is not an edit')
  })

  test('setWorkbookActiveTab replaces the tab instead of adding a second one', () => {
    const xml = '<?xml version="1.0"?><workbook xmlns="urn:x"><workbookPr/><bookViews><workbookView xWindow="0" activeTab="0" windowWidth="9"/></bookViews><sheets><sheet name="A" sheetId="1"/></sheets></workbook>'
    const out = setWorkbookActiveTab(xml, 2)
    assert.equal((out.match(/activeTab=/g) || []).length, 1)
    assert.match(out, /activeTab="2"/)
    assert.ok(!/activeTab="0"/.test(out))
  })

  test('setWorkbookActiveTab creates the bookViews container when there is none', () => {
    const xml = '<?xml version="1.0"?><workbook xmlns="urn:x"><workbookPr/><sheets><sheet name="A" sheetId="1"/></sheets><calcPr calcId="1"/></workbook>'
    const out = setWorkbookActiveTab(xml, 1)
    assert.match(out, /<workbookPr\/><bookViews><workbookView activeTab="1"\/><\/bookViews><sheets>/,
      'bookViews must precede sheets')
  })
})

// ---------------------------------------------------------------------------
// Print area and print titles
// ---------------------------------------------------------------------------

describe('Print area and print titles', () => {
  const book = path.join(tmpDir, 'areas.xlsx')

  before(async () => {
    await buildWorkbook(book)
  })

  test('printArea and printTitles round-trip through setPageSetup, read and inspect', async () => {
    const moved = path.join(tmpDir, 'areas-roundtrip.xlsx')
    fs.copyFileSync(book, moved)

    const result = await engine.setPageSetup(moved, {
      sheetName: 'Расходы',
      orientation: 'landscape',
      fitToWidth: 1,
      fitToHeight: 0,
      paperSize: 'A4',
      margins: { left: 0.5, right: 0.5, top: 0.6, bottom: 0.6 },
      printArea: 'A1:D4',
      printTitles: '1:2'
    })

    assert.equal(result.pageSetup.orientation, 'landscape')
    assert.equal(result.pageSetup.fitToWidth, 1)
    assert.equal(result.pageSetup.fitToHeight, 0)
    assert.equal(result.pageSetup.paperSizeName, 'A4')
    assert.equal(result.printArea, 'Расходы!$A$1:$D$4')
    assert.equal(result.printTitles, 'Расходы!$1:$2')
    assert.deepEqual(result.warnings, [])

    const read = await engine.read(moved, { sheetName: 'Расходы', includeStyles: true })
    assert.equal(read.pageSetup.orientation, 'landscape')
    assert.equal(read.pageSetup.fitToWidth, 1)
    assert.equal(read.pageSetup.margins.left, 0.5)
    assert.equal(read.printArea, 'Расходы!$A$1:$D$4')
    assert.equal(read.printTitles, 'Расходы!$1:$2')

    const info = await engine.inspect(moved)
    const expenses = info.sheets.find((s) => s.name === 'Расходы')
    assert.equal(expenses.pageSetup.orientation, 'landscape')
    assert.equal(expenses.pageSetup.paperSize, 9)
    assert.equal(expenses.printArea, 'Расходы!$A$1:$D$4')
    assert.equal(expenses.printTitles, 'Расходы!$1:$2')
    assert.equal(info.sheets.find((s) => s.name === 'Итоги').printArea, null,
      'a print area on one sheet is not reported for another')
  })

  test('a sheet name that needs quoting is quoted, and its apostrophes doubled', async () => {
    const quoted = path.join(tmpDir, 'quoted.xlsx')
    await buildWorkbook(quoted, ["Мой лист", "It's mine"])

    const first = await engine.setPageSetup(quoted, { sheetName: 'Мой лист', printArea: 'A1:B2' })
    assert.equal(first.printArea, "'Мой лист'!$A$1:$B$2")

    const second = await engine.setPageSetup(quoted, { sheetName: "It's mine", printArea: 'B2:C3' })
    assert.equal(second.printArea, "'It''s mine'!$B$2:$C$3")
    assert.equal(first.printArea, "'Мой лист'!$A$1:$B$2", 'the other sheet keeps its area')
  })

  test('several print areas and a rows+columns title pair are accepted', async () => {
    const multi = path.join(tmpDir, 'multi.xlsx')
    await buildWorkbook(multi, ['Лист1'])

    const result = await engine.setPageSetup(multi, {
      sheetName: 'Лист1',
      printArea: ['A1:B2', 'D1:E2'],
      printTitles: { rows: '1:1', columns: 'A:A' }
    })
    assert.equal(result.printArea, 'Лист1!$A$1:$B$2,Лист1!$D$1:$E$2')
    assert.equal(result.printTitles, 'Лист1!$1:$1,Лист1!$A:$A')

    const comma = await engine.setPageSetup(multi, { sheetName: 'Лист1', printTitles: '2:2,B:B' })
    assert.equal(comma.printTitles, 'Лист1!$2:$2,Лист1!$B:$B')
  })

  test('a print area is removed with null and unrelated defined names survive', async () => {
    const names = path.join(tmpDir, 'names.xlsx')
    await buildWorkbook(names, ['Лист1'])
    await engine.setPageSetup(names, { sheetName: 'Лист1', printArea: 'A1:B2' })

    // A named range of the caller's own, added the way a spreadsheet would.
    const zip = await ZipArchive.fromFile(names)
    const workbookXml = zip.getText('xl/workbook.xml')
    assert.match(workbookXml, /<definedNames>/, 'the print area created the container')
    zip.setText('xl/workbook.xml', workbookXml.replace('</definedNames>',
      '<definedName name="Ставка">Лист1!$C$3</definedName></definedNames>'))
    await zip.save(names)

    const kept = await engine.setPageSetup(names, { sheetName: 'Лист1', margins: { left: 0.4 } })
    assert.equal(kept.printArea, 'Лист1!$A$1:$B$2', 'an untouched print area is reported, not dropped')
    assert.match((await ZipArchive.fromFile(names)).getText('xl/workbook.xml'), /name="Ставка"/)

    const removed = await engine.setPageSetup(names, { sheetName: 'Лист1', printArea: null })
    assert.equal(removed.printArea, null)
    const after = (await ZipArchive.fromFile(names)).getText('xl/workbook.xml')
    assert.ok(!/Print_Area/.test(after))
    assert.match(after, /name="Ставка"/, 'the caller\'s own defined name is preserved')
    assert.match(after, /<definedNames><definedName name="Ставка">/, 'definedNames keeps a valid position')
    assert.equal((after.match(/<definedNames>/g) || []).length, 1, 'the container is rebuilt, not duplicated')
  })

  test('an unknown sheet name and a reversed range are refused', async () => {
    const bad = path.join(tmpDir, 'bad.xlsx')
    await buildWorkbook(bad, ['Лист1'])
    await assert.rejects(() => engine.setPageSetup(bad, { sheetName: 'Нет такого', fitToWidth: 1 }),
      /Sheet not found/)
    await assert.rejects(() => engine.setPageSetup(bad, { sheetName: 'Лист1', printArea: 'D8:A1' }),
      /end must not precede the start/)
  })
})

// ---------------------------------------------------------------------------
// Byte preservation
// ---------------------------------------------------------------------------

describe('Byte preservation', () => {
  test('only the named worksheet part changes — and the workbook part only for printArea/printTitles', async () => {
    const base = path.join(tmpDir, 'preserve-base.xlsx')
    await buildWorkbook(base)

    const layoutOnly = path.join(tmpDir, 'preserve-layout.xlsx')
    fs.copyFileSync(base, layoutOnly)
    await engine.setPageSetup(layoutOnly, {
      sheetName: 'Расходы',
      fitToWidth: 1,
      fitToHeight: 0,
      orientation: 'landscape',
      margins: { left: 0.5 },
      centerHorizontally: true
    })
    assert.deepEqual(await changedMembers(base, layoutOnly), ['xl/worksheets/sheet2.xml'],
      'a page-setup edit must touch exactly one package member')

    const withArea = path.join(tmpDir, 'preserve-area.xlsx')
    fs.copyFileSync(base, withArea)
    await engine.setPageSetup(withArea, { sheetName: 'Расходы', fitToWidth: 1, printArea: 'A1:D4' })
    assert.deepEqual((await changedMembers(base, withArea)).sort(),
      ['xl/workbook.xml', 'xl/worksheets/sheet2.xml'])

    // The untouched sheets are byte-identical, not merely equivalent.
    const before = await ZipArchive.fromFile(base)
    const after = await ZipArchive.fromFile(withArea)
    for (const part of ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet3.xml', 'xl/styles.xml', '[Content_Types].xml']) {
      assert.ok(after.getBuffer(part).equals(before.getBuffer(part)), `${part} must stay byte-identical`)
    }
    assert.ok(after.getBuffer('xl/worksheets/sheet2.xml').length > before.getBuffer('xl/worksheets/sheet2.xml').length)
  })

  test('format() can write a page setup without touching a single cell value', async () => {
    const book = path.join(tmpDir, 'format-layout.xlsx')
    await buildWorkbook(book, ['Лист1'])
    const before = await engine.read(book, { sheetName: 'Лист1', range: 'A1:D3', includeFormulas: true })

    const result = await engine.format(book, {
      sheetName: 'Лист1',
      pageSetup: { fitToWidth: 1, fitToHeight: 0, orientation: 'landscape' }
    })

    assert.equal(result.cellsFormatted, 0)
    assert.equal(result.pageSetup.fitToWidth, 1)
    const after = await engine.read(book, { sheetName: 'Лист1', range: 'A1:D3', includeFormulas: true })
    assert.deepEqual(after.data, before.data)
    assert.deepEqual(after.formulas, before.formulas)
  })

  test('format() still requires a range when no page setup is given', async () => {
    const book = path.join(tmpDir, 'format-range.xlsx')
    await buildWorkbook(book, ['Лист1'])
    await assert.rejects(() => engine.format(book, { sheetName: 'Лист1', font: { bold: true } }),
      /format requires a range/)
  })
})

// ---------------------------------------------------------------------------
// PDF assembly
// ---------------------------------------------------------------------------

describe('PDF assembly', () => {
  test('countPdfPages walks the page tree', () => {
    assert.equal(countPdfPages(buildPdf(['one'])), 1)
    assert.equal(countPdfPages(buildPdf(['one', 'two', 'three'])), 3)
    assert.equal(countPdfPages(Buffer.from('not a pdf at all')), 0)
  })

  test('merging keeps every page, in order, with its content intact', () => {
    const merged = mergePdfs([buildPdf(['alpha']), buildPdf(['beta', 'gamma']), buildPdf(['delta'])])
    assert.equal(countPdfPages(merged), 4)

    const contents = pageContentStreams(merged)
    assert.equal(contents.length, 4)
    for (const [i, marker] of ['alpha', 'beta', 'gamma', 'delta'].entries()) {
      assert.ok(contents[i].includes(`(${marker})`), `page ${i + 1} still draws ${marker}, got ${contents[i]}`)
    }
    assert.equal(new Set(contents).size, 4, 'no page is a duplicate of another')
  })

  test('a merged file is a well-formed PDF with one catalog and one page tree', () => {
    const merged = mergePdfs([buildPdf(['a']), buildPdf(['b'])])
    const text = merged.toString('latin1')
    assert.ok(text.startsWith('%PDF-'))
    assert.ok(text.trimEnd().endsWith('%%EOF'))
    assert.equal((text.match(/\/Type \/Catalog/g) || []).length, 1,
      'the inputs\' own catalogs are not copied — a second one would be unreachable')
    assert.equal((text.match(/\/Type \/Pages/g) || []).length, 1)
    assert.equal((text.match(/startxref/g) || []).length, 1)
  })

  test('a single document is returned byte-identical', () => {
    const single = buildPdf(['solo'])
    assert.ok(mergePdfs([single]).equals(single))
  })

  test('a reference with no target is refused instead of written', () => {
    assert.throws(() => mergePdfs([buildPdf(['a'], { dangling: 999 }), buildPdf(['b'])]),
      /has no target object/)
  })

  test('a file this module cannot parse is refused with a reason', () => {
    assert.equal(countPdfPages(Buffer.from('%PDF-1.7\n')), 0, 'an unparsable file has no page tree to count')
    assert.throws(() => parsePdfDocument(Buffer.from('short')), /too short/)
    assert.throws(() => parsePdfDocument(Buffer.from('%PDF-1.4\n/ObjStm\n')), /ObjStm/)
    assert.throws(() => parsePdfDocument(Buffer.from('%PDF-1.4\n1 0 obj\n<< /Type /Foo >>\nendobj\n')), /no document catalog/)
  })
})

// ---------------------------------------------------------------------------
// All-sheet export (needs R7)
// ---------------------------------------------------------------------------

describe('All-sheet PDF export', () => {
  const book = path.join(tmpDir, 'export.xlsx')
  const single = path.join(tmpDir, 'export-one.xlsx')
  let allSheetsResult = null

  before(async () => {
    await buildWorkbook(book)
    await buildWorkbook(single, ['Один лист'])
    for (const name of SHEETS) {
      await engine.setPageSetup(book, {
        sheetName: name,
        orientation: 'landscape',
        fitToWidth: 1,
        fitToHeight: 0,
        paperSize: 'A4',
        printArea: 'A1:D3'
      })
    }
    await engine.setPageSetup(single, { sheetName: 'Один лист', fitToWidth: 1, fitToHeight: 0, orientation: 'landscape' })
  })

  test('exportPdf refuses to run without an R7 converter', async () => {
    const stub = new XlsxEngine({ detect: async () => ({ installed: false, x2tPath: null }) })
    await assert.rejects(() => stub.exportPdf(book, { outputPath: path.join(tmpDir, 'never.pdf') }),
      /x2t converter is not available/)
  })

  test('every worksheet is exported by default, in tab order', async (t) => {
    if (!r7Available) {
      t.skip('R7 not installed')
      return
    }
    const out = path.join(tmpDir, 'export-all.pdf')
    allSheetsResult = await engine.exportPdf(book, { outputPath: out })

    assert.equal(allSheetsResult.allSheets, true, 'allSheets defaults to true')
    assert.deepEqual(allSheetsResult.sheets.map((s) => s.name), SHEETS, 'sheets are exported in tab order')
    assert.equal(allSheetsResult.pageCount, 3)
    assert.deepEqual(allSheetsResult.sheets.map((s) => s.pages), [1, 1, 1],
      'fitToWidth: 1 keeps each sheet on one page')
    assert.ok(fs.existsSync(out))
    assert.ok(fs.readFileSync(out).subarray(0, 1024).includes(Buffer.from('%PDF-')))

    // Page N must really be sheet N: a page's drawing commands are compared,
    // byte for byte, with a single-sheet render of that same worksheet.
    const mergedPages = pageContentStreams(fs.readFileSync(out))
    assert.equal(mergedPages.length, 3)
    assert.equal(new Set(mergedPages).size, 3, 'each page of the merged PDF is a different sheet')
    for (const [i, name] of SHEETS.entries()) {
      const one = path.join(tmpDir, `export-only-${i}.pdf`)
      await engine.exportPdf(book, { outputPath: one, sheetName: name })
      const [onlyPage] = pageContentStreams(fs.readFileSync(one))
      assert.equal(mergedPages[i], onlyPage, `page ${i + 1} of the merged PDF is worksheet "${name}"`)
    }
  })

  test('a named sheet exports on its own', async (t) => {
    if (!r7Available) {
      t.skip('R7 not installed')
      return
    }
    const out = path.join(tmpDir, 'export-named.pdf')
    const result = await engine.exportPdf(book, { outputPath: out, sheetName: 'Расходы' })

    assert.equal(result.allSheets, false)
    assert.equal(result.sheets.length, 1)
    assert.equal(result.sheets[0].name, 'Расходы')
    assert.equal(result.pageCount, 1)
  })

  test('allSheets: false exports the first sheet only', async (t) => {
    if (!r7Available) {
      t.skip('R7 not installed')
      return
    }
    const out = path.join(tmpDir, 'export-first.pdf')
    const result = await engine.exportPdf(book, { outputPath: out, allSheets: false })

    assert.equal(result.allSheets, false)
    assert.deepEqual(result.sheets.map((s) => s.name), [SHEETS[0]])
  })

  test('the page count grows with the number of sheets', async (t) => {
    if (!r7Available) {
      t.skip('R7 not installed')
      return
    }
    assert.ok(allSheetsResult, 'the three-sheet export ran first')
    const one = await engine.exportPdf(single, { outputPath: path.join(tmpDir, 'export-single.pdf') })

    assert.equal(one.pageCount, 1)
    assert.equal(allSheetsResult.pageCount, 3)
    assert.ok(allSheetsResult.pageCount > one.pageCount)
    assert.equal(allSheetsResult.pageCount, allSheetsResult.sheets.length,
      'with fitToWidth the page count is the sheet count')
  })

  test('an unknown sheet name fails before the converter runs', async () => {
    await assert.rejects(() => engine.exportPdf(book, { outputPath: path.join(tmpDir, 'no.pdf'), sheetName: 'Нет' }),
      /Sheet not found/)
  })

  test('the acceptance workbook shape renders with every sheet present', async (t) => {
    if (!r7Available) {
      t.skip('R7 not installed')
      return
    }
    // The same shape as examples/xlsx-acceptance.js: three sheets, wide comment
    // column, merged title, fit on one page.
    const acceptance = path.join(tmpDir, 'acceptance.xlsx')
    await engine.create(acceptance, {
      overwrite: true,
      sheets: SHEETS.map((name) => ({ name, data: [] }))
    })
    for (const [i, name] of SHEETS.entries()) {
      await engine.write(acceptance, {
        sheetName: name,
        cells: [
          { ref: 'A1', value: `Отчёт ${name} за I квартал 2026 года` },
          { ref: 'A2', value: 'Дата' },
          { ref: 'B2', value: 'Категория' },
          { ref: 'C2', value: 'Сумма, ₽' },
          { ref: 'D2', value: 'Комментарий' },
          { ref: 'A3', value: 46037 + i, date: true },
          { ref: 'B3', value: 'Продажи лицензий' },
          { ref: 'C3', value: 850000, numberFormat: { type: 'currency', symbol: '₽' } },
          { ref: 'D3', value: 'Годовые контракты' }
        ]
      })
      await engine.format(acceptance, { sheetName: name, range: 'A1:D1', merge: true, font: { bold: true } })
      await engine.format(acceptance, { sheetName: name, range: 'A1:D1', columnWidth: { width: 30 } })
      await engine.setPageSetup(acceptance, {
        sheetName: name,
        orientation: 'landscape',
        fitToWidth: 1,
        fitToHeight: 0,
        printArea: 'A1:D3',
        centerHorizontally: true
      })
    }

    const out = path.join(tmpDir, 'acceptance.pdf')
    const result = await engine.exportPdf(acceptance, { outputPath: out })

    assert.equal(result.sheets.length, 3)
    assert.equal(result.pageCount, 3)
    assert.ok(result.sheets.every((s) => s.pages === 1), 'no sheet spills onto a second page')
    assert.ok(fs.statSync(out).size > 500)
  })
})
