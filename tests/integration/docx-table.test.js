import { test, describe, before, after } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import { DocxEngine } from '../../src/r7/docx.js'
import { R7Adapter } from '../../src/r7/adapter.js'
import { extractElements } from '../../src/shared/xml.js'
import { archiveOf, cleanup, tempDir } from './docx-fixtures.test.js'

const r7Available = (await new R7Adapter().detect()).installed

/** Rows of the first table in a document. */
async function rowsOf(file) {
  const zip = await archiveOf(file)
  const table = extractElements(zip.getText('word/document.xml'), 'w:tbl')[0]
  return extractElements(table.outerXml, 'w:tr').map(row => row.outerXml)
}

/**
 * Scenario 5: merges, column widths, cell styling and row/column surgery.
 */
describe('DOCX tables', () => {
  const dir = tempDir('docx_table')
  const file = path.join(dir, 'table.docx')
  const engine = new DocxEngine()

  before(async () => {
    await engine.create(file, {
      title: 'Документ с таблицей',
      paragraphs: ['Таблица ниже.'],
      tables: [{
        widthsTwips: [2000, 3000, 4000],
        rows: [
          [{ value: 'A', shading: '#DEEAF6' }, 'B', 'C'],
          ['D', 'E', 'F'],
          ['G', 'H', 'I']
        ]
      }]
    })
  })

  after(() => cleanup(dir))

  test('a created table carries a grid, per-cell widths and borders', async () => {
    const result = await engine.table(file, { action: 'inspect', tableIndex: 0, includeStyles: true })
    assert.equal(result.rowCount, 3)
    assert.equal(result.colCount, 3)
    assert.deepEqual(result.data, [['A', 'B', 'C'], ['D', 'E', 'F'], ['G', 'H', 'I']])

    const formatting = result.formatting
    assert.deepEqual(formatting.columns, [2000, 3000, 4000])
    assert.equal(formatting.layout, 'fixed')
    assert.ok(formatting.borders?.top?.style)
    assert.equal(formatting.rows[0].cells[0].width.twips, 2000)
    assert.equal(formatting.rows[0].cells[0].shading, '#DEEAF6')
    assert.equal(formatting.rows[2].cells[2].width.twips, 4000)
  })

  test('merging horizontally spans the grid and keeps the first cell text', async () => {
    await engine.table(file, { action: 'merge', tableIndex: 0, merge: { row: 0, col: 0, cols: 2 } })

    const result = await engine.table(file, { action: 'inspect', tableIndex: 0, includeStyles: true })
    assert.equal(result.data[0].length, 2, 'two cells were merged into one')
    assert.equal(result.data[0][0], 'A', 'the merged cell keeps the first cell text')
    assert.equal(result.formatting.rows[0].cells[0].gridSpan, 2)
    assert.equal(result.formatting.rows[0].cells[1].text, 'C', 'the neighbour is untouched')
    assert.deepEqual(result.data[2], ['G', 'H', 'I'], 'other rows are untouched')
  })

  test('merging vertically marks one restart and the continuation cells', async () => {
    await engine.table(file, { action: 'merge', tableIndex: 0, merge: { row: 1, col: 1, rows: 2 } })

    const result = await engine.table(file, { action: 'inspect', tableIndex: 0, includeStyles: true })
    assert.equal(result.formatting.rows[1].cells[1].vMerge, 'restart')
    assert.equal(result.formatting.rows[2].cells[1].vMerge, 'continue')
    assert.equal(result.data[1][1], 'E')
    assert.equal(result.formatting.rows[1].cells[0].vMerge, null)
  })

  test('unmerging restores the covered cells', async () => {
    const horizontal = await engine.table(file, {
      action: 'unmerge',
      tableIndex: 0,
      merge: { row: 0, col: 0, axis: 'horizontal' }
    })
    assert.equal(horizontal.success, true)

    let result = await engine.table(file, { action: 'inspect', tableIndex: 0, includeStyles: true })
    assert.equal(result.data[0].length, 3, 'the covered columns are restored')
    assert.equal(result.formatting.rows[0].cells[0].gridSpan, 1)
    assert.equal(result.formatting.rows[0].cells[0].text, 'A')

    await engine.table(file, { action: 'unmerge', tableIndex: 0, merge: { row: 1, col: 1, axis: 'vertical' } })
    result = await engine.table(file, { action: 'inspect', tableIndex: 0, includeStyles: true })
    assert.equal(result.formatting.rows[1].cells[1].vMerge, null)
    assert.equal(result.formatting.rows[2].cells[1].vMerge, null, 'the continuation cell is released too')
  })

  test('column widths land in the grid and in every cell', async () => {
    const target = path.join(dir, 'widths.docx')
    await engine.create(target, {
      paragraphs: ['Ширины.'],
      tables: [{ rows: [['A', 'B'], ['C', 'D']], widthsTwips: [1000, 1000] }]
    })

    await engine.table(target, { action: 'setColumnWidths', tableIndex: 0, widthsTwips: [1500, 2500] })
    const result = await engine.table(target, { action: 'inspect', tableIndex: 0, includeStyles: true })
    assert.deepEqual(result.formatting.columns, [1500, 2500])
    for (const row of result.formatting.rows) {
      assert.deepEqual(row.cells.map(cell => cell.width.twips), [1500, 2500])
    }

    // A spanned cell receives the sum of the columns it covers.
    await engine.table(target, { action: 'merge', tableIndex: 0, merge: { row: 0, col: 0, cols: 2 } })
    await engine.table(target, { action: 'setColumnWidths', tableIndex: 0, widthsTwips: [2000, 3000] })
    const spanned = await engine.table(target, { action: 'inspect', tableIndex: 0, includeStyles: true })
    assert.equal(spanned.formatting.rows[0].cells[0].gridSpan, 2)
    assert.equal(spanned.formatting.rows[0].cells[0].width.twips, 5000)
    assert.deepEqual(spanned.formatting.columns, [2000, 3000])
  })

  test('a cell takes shading, borders, alignment, vertical alignment and width', async () => {
    const target = path.join(dir, 'cell-format.docx')
    await engine.create(target, {
      paragraphs: ['Стиль ячейки.'],
      tables: [{ rows: [['A', 'B'], ['C', 'D']], widthsTwips: [2000, 2000] }]
    })

    const result = await engine.table(target, {
      action: 'formatCell',
      tableIndex: 0,
      cell: {
        row: 1,
        col: 1,
        shading: '#FFF2CC',
        verticalAlign: 'center',
        alignment: 'right',
        widthCm: 3,
        borders: { top: { style: 'double', sizePoints: 2, color: '#FF0000' }, left: { style: 'dashed' } }
      }
    })
    assert.equal(result.success, true)

    const cell = (await engine.table(target, { action: 'inspect', tableIndex: 0, includeStyles: true }))
      .formatting.rows[1].cells[1]
    assert.equal(cell.shading, '#FFF2CC')
    assert.equal(cell.verticalAlign, 'center')
    assert.equal(cell.alignment, 'right')
    assert.equal(cell.width.cm, 3)
    assert.equal(cell.borders.top.style, 'double')
    assert.equal(cell.borders.top.sizeEighthsOfPoint, 16)
    assert.equal(cell.borders.top.color, '#FF0000')
    assert.equal(cell.borders.left.style, 'dashed')

    // The neighbour cell is untouched.
    const neighbour = (await engine.table(target, { action: 'inspect', tableIndex: 0, includeStyles: true }))
      .formatting.rows[1].cells[0]
    assert.equal(neighbour.shading, null)
    assert.equal(neighbour.borders, null)
  })

  test('changing only a cell text keeps that cell formatting', async () => {
    const target = path.join(dir, 'cell-text.docx')
    await engine.create(target, {
      paragraphs: ['Текст ячейки.'],
      tables: [{ rows: [['A', 'B'], ['C', 'D']], widthsTwips: [2000, 2000] }]
    })
    await engine.table(target, {
      action: 'formatCell',
      tableIndex: 0,
      cell: { row: 0, col: 0, shading: '#F8CBAD', alignment: 'center' }
    })

    await engine.table(target, { action: 'setCell', tableIndex: 0, cell: { row: 0, col: 0, value: 'Обновлено' } })
    const cell = (await engine.table(target, { action: 'inspect', tableIndex: 0, includeStyles: true }))
      .formatting.rows[0].cells[0]
    assert.equal(cell.text, 'Обновлено')
    assert.equal(cell.shading, '#F8CBAD', 'setCell must not repaint the cell')
    assert.equal(cell.alignment, 'center')
  })

  test('rows are added at the requested index and removed by index', async () => {
    const target = path.join(dir, 'rows.docx')
    await engine.create(target, {
      paragraphs: ['Строки.'],
      tables: [{ rows: [['A', 'B'], ['C', 'D']], widthsTwips: [2000, 3000] }]
    })

    await engine.table(target, { action: 'addRow', tableIndex: 0, index: 1, values: ['X', 'Y'] })
    let result = await engine.table(target, { action: 'inspect', tableIndex: 0, includeStyles: true })
    assert.deepEqual(result.data, [['A', 'B'], ['X', 'Y'], ['C', 'D']])
    assert.deepEqual(result.formatting.rows[1].cells.map(cell => cell.width.twips), [2000, 3000],
      'a new row follows the table grid')

    const rowsBefore = await rowsOf(target)
    await engine.table(target, { action: 'removeRow', tableIndex: 0, rowIndex: 1 })
    result = await engine.table(target, { action: 'inspect', tableIndex: 0 })
    assert.deepEqual(result.data, [['A', 'B'], ['C', 'D']])

    const rowsAfter = await rowsOf(target)
    assert.ok(rowsAfter.includes(rowsBefore[0]), 'untouched rows keep their exact XML')
    assert.ok(rowsAfter.includes(rowsBefore[2]))
  })

  test('columns are added and removed with the grid kept in sync', async () => {
    const target = path.join(dir, 'columns.docx')
    await engine.create(target, {
      paragraphs: ['Столбцы.'],
      tables: [{ rows: [['A', 'B'], ['C', 'D']], widthsTwips: [2000, 2000] }]
    })

    await engine.table(target, {
      action: 'addColumn',
      tableIndex: 0,
      index: 1,
      values: ['c0', 'c1'],
      widthTwips: 1000,
      shading: '#E2EFDA'
    })
    let result = await engine.table(target, { action: 'inspect', tableIndex: 0, includeStyles: true })
    assert.deepEqual(result.data, [['A', 'c0', 'B'], ['C', 'c1', 'D']])
    assert.deepEqual(result.formatting.columns, [2000, 1000, 2000])
    assert.equal(result.formatting.rows[0].cells[1].shading, '#E2EFDA')

    await engine.table(target, { action: 'removeColumn', tableIndex: 0, columnIndex: 1 })
    result = await engine.table(target, { action: 'inspect', tableIndex: 0, includeStyles: true })
    assert.deepEqual(result.data, [['A', 'B'], ['C', 'D']])
    assert.deepEqual(result.formatting.columns, [2000, 2000])
  })

  test('impossible table operations fail loudly instead of corrupting the table', async () => {
    const target = path.join(dir, 'errors.docx')
    await engine.create(target, { paragraphs: ['Ошибки.'], tables: [{ rows: [['A', 'B'], ['C', 'D']] }] })

    await assert.rejects(
      () => engine.table(target, { action: 'merge', merge: { row: 0, col: 0, cols: 5 } }),
      /Merge covers columns/
    )
    await assert.rejects(
      () => engine.table(target, { action: 'merge', merge: { row: 0, col: 0, rows: 9 } }),
      /Merge covers rows/
    )
    await assert.rejects(() => engine.table(target, { action: 'removeRow', rowIndex: 4 }), /Row index out of range/)
    await assert.rejects(() => engine.table(target, { action: 'removeColumn', columnIndex: 9 }), /Column index out of range/)
    await assert.rejects(() => engine.table(target, { action: 'nonsense' }), /Unsupported table action/)
    assert.equal((await engine.validate(target)).valid, true)
  })

  test('a table survives an unrelated paragraph edit and still reads back', async () => {
    const before = await engine.table(file, { action: 'inspect', tableIndex: 0 })

    await engine.replaceText(file, 'Таблица ниже.', 'Таблица изменена.')
    await engine.table(file, { action: 'setCell', tableIndex: 0, cell: { row: 2, col: 2, value: 'обновлено' } })

    const after = await engine.table(file, { action: 'inspect', tableIndex: 0 })
    assert.equal(after.rowCount, before.rowCount)
    assert.equal(after.colCount, before.colCount)
    assert.equal(after.data[0][0], 'A', 'unrelated cells are unchanged')
    assert.equal(after.data[2][2], 'обновлено')
    assert.equal((await engine.validate(file)).valid, true)
  })

  test('R7 reopens the edited table and renders it to PDF', async (t) => {
    if (!r7Available) {
      t.skip('R7 not installed')
      return
    }
    const pdf = path.join(dir, 'table.pdf')
    await new R7Adapter().convert(file, pdf)
    assert.equal(fs.readFileSync(pdf).subarray(0, 5).toString(), '%PDF-')
  })
})
