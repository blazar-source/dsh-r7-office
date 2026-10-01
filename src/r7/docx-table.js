/**
 * DOCX tables: grid geometry, merges, cell formatting and row/column surgery.
 *
 * Tables are addressed by **element position**: `row` is the 0-based `<w:tr>`
 * index and `col` is the 0-based `<w:tc>` index inside that row. That is the
 * only addressing scheme that stays meaningful on a table that is already
 * merged, because a merged cell occupies one element while covering several
 * visual columns — a visual-grid model would have to invent phantom cells that
 * do not exist in the file.
 *
 * Every mutation rebuilds only the touched `<w:tr>` / `<w:tc>` elements and
 * splices them back into the table, so unrelated rows (and every other part of
 * the document) keep their exact bytes.
 *
 * Child elements inside `<w:tcPr>`, `<w:trPr>` and `<w:tblPr>` are written in
 * schema order, because both Word and R7 validate it.
 */

import {
  childElements,
  expandSelfClosing,
  removeChildElements,
  upsertChildOrdered
} from './docx-parts.js'
import {
  buildBorderEdge,
  cmToTwips,
  extractTableFormatting,
  pointsToTwips
} from './docx-styles.js'
import { extractElements, extractTextFromXml, getAttribute } from '../shared/xml.js'

/** Child order inside `<w:tcPr>`, per the WordprocessingML schema. */
export const TCPR_ORDER = [
  'w:cnfStyle',
  'w:tcW',
  'w:gridSpan',
  'w:hMerge',
  'w:vMerge',
  'w:tcBorders',
  'w:shd',
  'w:noWrap',
  'w:tcMar',
  'w:textDirection',
  'w:tcFitText',
  'w:vAlign',
  'w:hideMark'
]

/** Child order inside `<w:trPr>`, per the WordprocessingML schema. */
export const TRPR_ORDER = [
  'w:cnfStyle',
  'w:divId',
  'w:gridBefore',
  'w:gridAfter',
  'w:wBefore',
  'w:wAfter',
  'w:cantSplit',
  'w:trHeight',
  'w:tblHeader',
  'w:tblCellSpacing',
  'w:jc',
  'w:hidden'
]

/** Child order inside `<w:tblPr>`, per the WordprocessingML schema. */
export const TBLEPR_ORDER = [
  'w:tblStyle',
  'w:tblpPr',
  'w:tblOverlap',
  'w:bidiVisual',
  'w:tblStyleRowBandSize',
  'w:tblStyleColBandSize',
  'w:tblW',
  'w:jc',
  'w:tblCellSpacing',
  'w:tblInd',
  'w:tblBorders',
  'w:shd',
  'w:tblLayout',
  'w:tblCellMar',
  'w:tblLook',
  'w:tblCaption',
  'w:tblDescription'
]

/** Border edges a cell can carry. */
const CELL_BORDER_EDGES = ['top', 'left', 'bottom', 'right']

/** An empty cell, used when an unmerge restores covered columns. */
const EMPTY_CELL = '<w:tc><w:p/></w:tc>'

/** Rows of a table, with their spans inside the table XML. */
function tableRows(tblXml) {
  return extractElements(tblXml, 'w:tr')
}

/** Cells of one row, with their spans inside the row XML. */
function rowCells(rowXml) {
  return extractElements(rowXml, 'w:tc')
}

/**
 * Rebuild a `<w:tr>` from a list of cell element XML strings.
 * @param {string} rowXml
 * @param {string[]} cells
 * @returns {string}
 */
export function rebuildRow(rowXml, cells) {
  const open = /<w:tr\b[^>]*>/.exec(rowXml)?.[0] ?? '<w:tr>'
  const trPr = childElements(rowXml).find(element => element.tag === 'w:trPr')
  return `${open}${trPr ? trPr.xml : ''}${cells.join('')}</w:tr>`
}

/**
 * Rebuild a `<w:tbl>` from a list of row element XML strings.
 * @param {string} tblXml
 * @param {string[]} rows
 * @returns {string}
 */
export function rebuildTable(tblXml, rows) {
  const open = /<w:tbl\b[^>]*>/.exec(tblXml)?.[0] ?? '<w:tbl>'
  const head = childElements(tblXml)
    .filter(element => element.tag !== 'w:tr')
    .map(element => element.xml)
    .join('')
  return `${open}${head}${rows.join('')}</w:tbl>`
}

/** Replace one row element inside a table. */
function replaceRow(tblXml, rowIndex, newRowXml) {
  const rows = tableRows(tblXml)
  if (rowIndex < 0 || rowIndex >= rows.length) {
    throw new Error(`Row index out of range: ${rowIndex} (table has ${rows.length} rows)`)
  }
  const target = rows[rowIndex]
  return tblXml.slice(0, target.index) + newRowXml + tblXml.slice(target.index + target.outerXml.length)
}

/**
 * Replace, add or remove ordered children inside a properties container
 * (`<w:tcPr>`, `<w:trPr>`, `<w:tblPr>`), creating the container when needed.
 *
 * @param {string} xml - the element that owns the container
 * @param {object} changes - tag to element XML, or null to remove that child
 * @param {string[]} order - the container's schema order
 * @param {string} containerTag - e.g. `w:tcPr`
 * @returns {string}
 */
export function upsertProperties(xml, changes, order, containerTag) {
  const container = extractElements(xml, containerTag)[0]
  let inner = container ? container.innerXml : ''
  for (const [tag, childXml] of Object.entries(changes)) {
    inner = upsertChildOrdered(inner, tag, childXml, order)
  }

  const element = inner.length > 0 ? `<${containerTag}>${inner}</${containerTag}>` : ''
  if (container) {
    return xml.slice(0, container.index) + element + xml.slice(container.index + container.outerXml.length)
  }
  if (!element) return xml

  // A new `<w:tcPr>` precedes the cell's paragraphs, a new `<w:trPr>` its cells
  // and a new `<w:tblPr>` the grid and rows.
  const anchorTag = containerTag === 'w:tcPr' ? 'w:p' : (containerTag === 'w:trPr' ? 'w:tc' : 'w:tblGrid')
  const anchor = extractElements(xml, anchorTag)[0]
  if (anchor) return xml.slice(0, anchor.index) + element + xml.slice(anchor.index)

  const ownerTag = containerTag === 'w:tcPr' ? 'w:tc' : (containerTag === 'w:trPr' ? 'w:tr' : 'w:tbl')
  const close = `</${ownerTag}>`
  if (!xml.includes(close)) throw new Error(`Cannot add ${containerTag}: ${close} not found`)
  return xml.replace(close, `${element}${close}`)
}

/**
 * Build a `<w:tc>` cell element.
 *
 * @param {object} spec
 * @param {string} [spec.text]
 * @param {number} [spec.widthTwips]
 * @param {number} [spec.gridSpan]
 * @param {'restart'|'continue'|null} [spec.vMerge]
 * @param {string} [spec.shading] - `#RRGGBB`
 * @param {object} [spec.borders]
 * @param {object} [spec.margins] - points per edge
 * @param {string} [spec.verticalAlign]
 * @param {string} [spec.alignment] - paragraph alignment
 * @returns {string}
 */
export function buildCellXml(spec = {}) {
  const changes = {}
  if (spec.widthTwips !== undefined && spec.widthTwips !== null) {
    changes['w:tcW'] = `<w:tcW w:w="${Math.round(spec.widthTwips)}" w:type="dxa"/>`
  }
  if (spec.gridSpan !== undefined && spec.gridSpan !== null && Number(spec.gridSpan) > 1) {
    changes['w:gridSpan'] = `<w:gridSpan w:val="${Math.round(spec.gridSpan)}"/>`
  }
  if (spec.vMerge) {
    changes['w:vMerge'] = spec.vMerge === 'restart' ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>'
  }
  if (spec.borders && typeof spec.borders === 'object') {
    const edges = CELL_BORDER_EDGES.filter(edge => spec.borders[edge] !== undefined && spec.borders[edge] !== null)
    if (edges.length > 0) {
      changes['w:tcBorders'] = `<w:tcBorders>${edges
        .map(edge => buildBorderEdge(edge, spec.borders[edge]))
        .join('')}</w:tcBorders>`
    }
  }
  if (spec.shading) {
    changes['w:shd'] = `<w:shd w:val="clear" w:color="auto" w:fill="${String(spec.shading).replace(/^#/, '')}"/>`
  }
  if (spec.margins && typeof spec.margins === 'object') {
    const edges = ['top', 'left', 'bottom', 'right']
      .filter(edge => spec.margins[edge] !== undefined && spec.margins[edge] !== null)
    if (edges.length > 0) {
      changes['w:tcMar'] = `<w:tcMar>${edges
        .map(edge => `<w:${edge} w:w="${pointsToTwips(spec.margins[edge])}" w:type="dxa"/>`)
        .join('')}</w:tcMar>`
    }
  }
  if (spec.verticalAlign) changes['w:vAlign'] = `<w:vAlign w:val="${spec.verticalAlign}"/>`

  const inner = TCPR_ORDER.filter(tag => changes[tag]).map(tag => changes[tag]).join('')
  const tcPr = inner ? `<w:tcPr>${inner}</w:tcPr>` : ''
  const alignment = spec.alignment ? `<w:jc w:val="${spec.alignment}"/>` : ''
  const spacing = spec.paragraphSpacingBefore !== undefined && spec.paragraphSpacingBefore !== null
    ? `<w:spacing w:before="${pointsToTwips(spec.paragraphSpacingBefore)}"/>`
    : ''
  const pPr = (alignment || spacing) ? `<w:pPr>${spacing}${alignment}</w:pPr>` : ''
  const text = spec.text === undefined || spec.text === null ? '' : String(spec.text)
  const run = text === '' ? '' : `<w:r><w:t xml:space="preserve">${escapeXmlText(text)}</w:t></w:r>`

  return `<w:tc>${tcPr}<w:p>${pPr}${run}</w:p></w:tc>`
}

function escapeXmlText(value) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/**
 * The text matrix of a table, cell by cell.
 * @param {string} tblXml
 * @returns {string[][]}
 */
export function tableText(tblXml) {
  return tableRows(tblXml).map(row => rowCells(row.outerXml).map(cell => extractTextFromXml(cell.outerXml).trim()))
}

/**
 * Normalized description of a table (geometry plus per-cell formatting).
 * @param {string} tblXml
 * @param {object} [context]
 * @returns {object}
 */
export function describeTable(tblXml, context = {}) {
  return {
    ...extractTableFormatting(tblXml, context),
    data: tableText(tblXml)
  }
}

/**
 * Merge a rectangular block of cells.
 *
 * `col` and `cols` count `<w:tc>` elements, so a merge can start on a cell
 * that already spans columns.
 *
 * @param {string} tblXml
 * @param {object} options
 * @param {number} options.row - top-left cell's row
 * @param {number} options.col - top-left cell's `<w:tc>` index
 * @param {number} [options.rows=1] - how many rows the merge covers
 * @param {number} [options.cols=1] - how many cells the merge covers per row
 * @returns {string}
 */
export function mergeCells(tblXml, options) {
  const rowStart = Number(options.row ?? 0)
  const col = Number(options.col ?? 0)
  const rowSpan = Math.max(1, Number(options.rows ?? 1))
  const colSpan = Math.max(1, Number(options.cols ?? 1))

  let rows = tableRows(tblXml)
  if (rowStart < 0 || rowStart >= rows.length) {
    throw new Error(`Row index out of range: ${rowStart} (table has ${rows.length} rows)`)
  }
  if (rowStart + rowSpan > rows.length) {
    throw new Error(`Merge covers rows ${rowStart}..${rowStart + rowSpan - 1} but the table has ${rows.length} rows`)
  }
  if (rowSpan === 1 && colSpan === 1) return tblXml

  for (let r = rowStart; r < rowStart + rowSpan; r++) {
    const count = rowCells(rows[r].outerXml).length
    if (col + colSpan > count) {
      throw new Error(`Merge covers columns ${col}..${col + colSpan - 1} but row ${r} has ${count} cells`)
    }
  }

  let xml = tblXml
  for (let r = rowStart; r < rowStart + rowSpan; r++) {
    const cells = rowCells(rows[r].outerXml)
    const changes = {}
    if (colSpan > 1) changes['w:gridSpan'] = `<w:gridSpan w:val="${colSpan}"/>`
    if (rowSpan > 1) changes['w:vMerge'] = r === rowStart ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>'
    const merged = upsertProperties(cells[col].outerXml, changes, TCPR_ORDER, 'w:tcPr')

    const newCells = cells
      .map(cell => cell.outerXml)
      .filter((_, index) => index < col || index >= col + colSpan)
    newCells.splice(col, 0, merged)

    xml = replaceRow(xml, r, rebuildRow(rows[r].outerXml, newCells))
    rows = tableRows(xml)
  }

  return xml
}

/**
 * Undo a horizontal and/or vertical merge starting at a cell.
 *
 * Removing a `gridSpan` also restores the covered cells (so the row regains
 * its column count); removing a `vMerge` restart also clears the continuation
 * cells it covered.
 *
 * @param {string} tblXml
 * @param {object} options
 * @param {number} options.row
 * @param {number} options.col
 * @param {'all'|'horizontal'|'vertical'} [options.axis='all']
 * @returns {string}
 */
export function unmergeCells(tblXml, options) {
  const row = Number(options.row ?? 0)
  const col = Number(options.col ?? 0)
  const axis = options.axis || 'all'

  let rows = tableRows(tblXml)
  if (row < 0 || row >= rows.length) throw new Error(`Row index out of range: ${row} (table has ${rows.length} rows)`)
  let cells = rowCells(rows[row].outerXml)
  if (col < 0 || col >= cells.length) throw new Error(`Column index out of range: ${col} (row has ${cells.length} cells)`)

  let xml = tblXml
  const gridSpan = Number(getAttribute(extractElements(cells[col].outerXml, 'w:gridSpan')[0]?.outerXml ?? '', 'w:val') || 1)

  if (axis !== 'vertical' && gridSpan > 1) {
    const cleared = removeChildElements(cells[col].outerXml, 'w:gridSpan')
    const newCells = cells.map(cell => cell.outerXml)
    newCells[col] = cleared
    newCells.splice(col + 1, 0, ...Array.from({ length: gridSpan - 1 }, () => EMPTY_CELL))
    xml = replaceRow(xml, row, rebuildRow(rows[row].outerXml, newCells))
    rows = tableRows(xml)
    cells = rowCells(rows[row].outerXml)
  }

  if (axis === 'horizontal') return xml

  const vMergeTag = extractElements(cells[col].outerXml, 'w:vMerge')[0]?.outerXml ?? null
  if (!vMergeTag) return xml
  const isRestart = (getAttribute(vMergeTag, 'w:val') || '').toLowerCase() === 'restart'

  if (!isRestart) {
    const cleared = removeChildElements(cells[col].outerXml, 'w:vMerge')
    xml = replaceRow(xml, row, rebuildRow(rows[row].outerXml, cells.map((cell, index) => (
      index === col ? cleared : cell.outerXml
    ))))
    return xml
  }

  for (let r = row; r < rows.length; r++) {
    const rowCellsNow = rowCells(rows[r].outerXml)
    if (col >= rowCellsNow.length) break
    const current = rowCellsNow[col].outerXml
    const mergeTag = extractElements(current, 'w:vMerge')[0]?.outerXml ?? null
    if (!mergeTag) break
    if (r > row && (getAttribute(mergeTag, 'w:val') || '').toLowerCase() === 'restart') break

    const cleared = removeChildElements(current, 'w:vMerge')
    xml = replaceRow(xml, r, rebuildRow(rows[r].outerXml, rowCellsNow.map((cell, index) => (
      index === col ? cleared : cell.outerXml
    ))))
    rows = tableRows(xml)
  }

  return xml
}

/**
 * Set the table's column widths, updating `<w:tblGrid>` and every cell's
 * `<w:tcW>` so the two agree.
 *
 * @param {string} tblXml
 * @param {object} options
 * @param {number[]} options.widthsTwips
 * @param {'dxa'|'pct'|'auto'} [options.mode='dxa']
 * @param {boolean} [options.fixedLayout=true]
 * @returns {string}
 */
export function setColumnWidths(tblXml, options) {
  const widths = options.widthsTwips
  if (!Array.isArray(widths) || widths.length === 0) {
    throw new Error('setColumnWidths requires a non-empty widthsTwips array')
  }
  if (widths.some(width => !Number.isFinite(Number(width)))) {
    throw new Error('setColumnWidths requires numeric widthsTwips values')
  }

  let xml = tblXml
  const grid = extractElements(xml, 'w:tblGrid')[0]
  const gridXml = `<w:tblGrid>${widths.map(width => `<w:gridCol w:w="${Math.round(Number(width))}"/>`).join('')}</w:tblGrid>`
  if (grid) {
    xml = xml.slice(0, grid.index) + gridXml + xml.slice(grid.index + grid.outerXml.length)
  } else {
    const tblPr = extractElements(xml, 'w:tblPr')[0]
    const at = tblPr ? tblPr.index + tblPr.outerXml.length : (/<w:tbl\b[^>]*>/.exec(xml)?.[0].length ?? 0)
    xml = xml.slice(0, at) + gridXml + xml.slice(at)
  }

  if (options.fixedLayout !== false) {
    xml = upsertProperties(xml, { 'w:tblLayout': '<w:tblLayout w:type="fixed"/>' }, TBLEPR_ORDER, 'w:tblPr')
  }

  // Every cell follows the grid: a spanned cell gets the sum of its columns.
  const rows = tableRows(xml)
  for (let r = 0; r < rows.length; r++) {
    const cells = rowCells(rows[r].outerXml)
    let column = 0
    const updatedCells = cells.map(cell => {
      const span = Number(getAttribute(extractElements(cell.outerXml, 'w:gridSpan')[0]?.outerXml ?? '', 'w:val') || 1)
      let width = 0
      for (let i = 0; i < span; i++) width += Number(widths[column + i] ?? widths[widths.length - 1])
      column += span
      return upsertProperties(
        cell.outerXml,
        { 'w:tcW': `<w:tcW w:w="${Math.round(width)}" w:type="${options.mode || 'dxa'}"/>` },
        TCPR_ORDER,
        'w:tcPr'
      )
    })
    xml = replaceRow(xml, r, rebuildRow(rows[r].outerXml, updatedCells))
  }

  return xml
}

/**
 * Apply cell formatting to one cell, optionally replacing its text.
 *
 * @param {string} tblXml
 * @param {object} options
 * @param {number} options.row
 * @param {number} options.col
 * @param {string} [options.text]
 * @param {string} [options.shading]
 * @param {object} [options.borders]
 * @param {boolean} [options.clearBorders]
 * @param {string} [options.verticalAlign]
 * @param {string} [options.alignment]
 * @param {number} [options.widthTwips]
 * @param {number} [options.widthCm]
 * @param {object} [options.margins] - points per edge
 * @returns {string}
 */
export function formatCell(tblXml, options) {
  const rows = tableRows(tblXml)
  const row = Number(options.row ?? 0)
  const col = Number(options.col ?? 0)
  if (row < 0 || row >= rows.length) throw new Error(`Row index out of range: ${row} (table has ${rows.length} rows)`)
  const cells = rowCells(rows[row].outerXml)
  if (col < 0 || col >= cells.length) throw new Error(`Column index out of range: ${col} (row has ${cells.length} cells)`)

  const changes = {}
  if (options.clearBorders) changes['w:tcBorders'] = null
  if (options.borders && typeof options.borders === 'object') {
    const edges = CELL_BORDER_EDGES.filter(edge => options.borders[edge] !== undefined && options.borders[edge] !== null)
    if (edges.length > 0) {
      changes['w:tcBorders'] = `<w:tcBorders>${edges
        .map(edge => buildBorderEdge(edge, options.borders[edge]))
        .join('')}</w:tcBorders>`
    }
  }
  if (options.shading !== undefined && options.shading !== null) {
    changes['w:shd'] = (options.shading === 'none' || options.shading === '')
      ? '<w:shd w:val="clear" w:color="auto" w:fill="auto"/>'
      : `<w:shd w:val="clear" w:color="auto" w:fill="${String(options.shading).replace(/^#/, '')}"/>`
  }
  if (options.verticalAlign !== undefined && options.verticalAlign !== null) {
    changes['w:vAlign'] = options.verticalAlign ? `<w:vAlign w:val="${options.verticalAlign}"/>` : null
  }
  if (options.widthTwips !== undefined && options.widthTwips !== null) {
    changes['w:tcW'] = `<w:tcW w:w="${Math.round(Number(options.widthTwips))}" w:type="dxa"/>`
  } else if (options.widthCm !== undefined && options.widthCm !== null) {
    changes['w:tcW'] = `<w:tcW w:w="${cmToTwips(options.widthCm)}" w:type="dxa"/>`
  }
  if (options.margins && typeof options.margins === 'object') {
    const edges = ['top', 'left', 'bottom', 'right']
      .filter(edge => options.margins[edge] !== undefined && options.margins[edge] !== null)
    if (edges.length > 0) {
      changes['w:tcMar'] = `<w:tcMar>${edges
        .map(edge => `<w:${edge} w:w="${pointsToTwips(options.margins[edge])}" w:type="dxa"/>`)
        .join('')}</w:tcMar>`
    }
  }

  let cell = upsertProperties(cells[col].outerXml, changes, TCPR_ORDER, 'w:tcPr')
  if (options.text !== undefined && options.text !== null) {
    cell = setCellText(cell, options.text, { paragraphIndex: 0 })
  }
  if (options.alignment !== undefined) {
    cell = applyCellParagraphAlignment(cell, options.alignment)
  }

  return replaceRow(tblXml, row, rebuildRow(rows[row].outerXml, cells.map((entry, index) => (
    index === col ? cell : entry.outerXml
  ))))
}

/** Apply (or clear, with a falsy alignment) `<w:jc>` on every paragraph of a cell. */
function applyCellParagraphAlignment(cellXml, alignment) {
  let result = cellXml
  const paragraphs = extractElements(result, 'w:p')
  for (let i = paragraphs.length - 1; i >= 0; i--) {
    const paragraph = paragraphs[i]
    const paragraphXml = expandSelfClosing(paragraph.outerXml)
    let updated
    if (alignment) {
      const withoutJc = removeChildElements(paragraphXml, 'w:jc')
      updated = /<w:pPr(?=[\s/>])/.test(withoutJc)
        ? withoutJc.replace(/(<w:pPr(?=[\s/>])[^>]*>)/, `$1<w:jc w:val="${alignment}"/>`)
        : withoutJc.replace(/^(<w:p(?=[\s/>])[^>]*>)/, `$1<w:pPr><w:jc w:val="${alignment}"/></w:pPr>`)
    } else {
      updated = removeChildElements(paragraphXml, 'w:jc')
    }
    result = result.slice(0, paragraph.index) + updated + result.slice(paragraph.index + paragraph.outerXml.length)
  }
  return result
}

/**
 * Replace a cell's text, keeping the cell properties.
 * @param {string} cellXml
 * @param {string} text
 * @param {object} [options]
 * @param {number} [options.paragraphIndex=0]
 * @returns {string}
 */
export function setCellText(cellXml, text, options = {}) {
  const paragraphIndex = options.paragraphIndex ?? 0
  const paragraphs = extractElements(cellXml, 'w:p')
  if (paragraphs.length === 0) {
    const anchor = extractElements(cellXml, 'w:tcPr')[0]
    const at = anchor ? anchor.index + anchor.outerXml.length : (/<w:tc\b[^>]*>/.exec(cellXml)?.[0].length ?? 0)
    const paragraph = `<w:p><w:r><w:t xml:space="preserve">${escapeXmlText(text)}</w:t></w:r></w:p>`
    return cellXml.slice(0, at) + paragraph + cellXml.slice(at)
  }

  const target = paragraphs[Math.min(paragraphIndex, paragraphs.length - 1)]
  const pPr = extractElements(target.outerXml, 'w:pPr')[0]
  const paragraph = `<w:p>${pPr ? pPr.outerXml : ''}<w:r><w:t xml:space="preserve">${escapeXmlText(text)}</w:t></w:r></w:p>`
  let result = cellXml.slice(0, target.index) + paragraph + cellXml.slice(target.index + target.outerXml.length)

  // Extra paragraphs in a cell are rare; their text is cleared so the cell
  // reads exactly as the caller asked instead of carrying a stale second line.
  const remaining = extractElements(result, 'w:p')
  for (let i = remaining.length - 1; i >= paragraphIndex + 1; i--) {
    const pPrXml = extractElements(remaining[i].outerXml, 'w:pPr')[0]?.outerXml ?? ''
    const empty = `<w:p>${pPrXml}</w:p>`
    result = result.slice(0, remaining[i].index) + empty + result.slice(remaining[i].index + remaining[i].outerXml.length)
  }
  return result
}

/**
 * Insert a row.
 * @param {string} tblXml
 * @param {object} [options]
 * @param {string[]} [options.values] - cell texts
 * @param {number} [options.index] - 0-based insertion index (default: append)
 * @param {number[]} [options.widthsTwips] - per-cell widths
 * @param {string} [options.shading]
 * @param {string} [options.verticalAlign]
 * @param {string} [options.alignment]
 * @returns {string}
 */
export function addRow(tblXml, options = {}) {
  const rows = tableRows(tblXml)
  const values = options.values || []
  const gridWidths = options.widthsTwips
    || extractElements(extractElements(tblXml, 'w:tblGrid')[0]?.innerXml ?? '', 'w:gridCol')
      .map(column => Number(getAttribute(column.outerXml, 'w:w')))

  const cellCount = values.length || gridWidths.length || (rows.length > 0 ? rowCells(rows[0].outerXml).length : 0)
  const cells = []
  for (let i = 0; i < cellCount; i++) {
    cells.push(buildCellXml({
      text: values[i] ?? '',
      widthTwips: gridWidths[i],
      shading: options.shading,
      verticalAlign: options.verticalAlign,
      alignment: options.alignment
    }))
  }

  const index = options.index === undefined || options.index === null ? rows.length : Number(options.index)
  if (index < 0 || index > rows.length) {
    throw new Error(`Row insertion index out of range: ${index} (table has ${rows.length} rows)`)
  }

  const open = /<w:tr\b[^>]*>/.exec(rows[0]?.outerXml ?? '')?.[0] ?? '<w:tr>'
  const allRows = rows.map(row => row.outerXml)
  allRows.splice(index, 0, `${open}${cells.join('')}</w:tr>`)
  return rebuildTable(tblXml, allRows)
}

/**
 * Remove a row.
 * @param {string} tblXml
 * @param {number} rowIndex
 * @returns {string}
 */
export function removeRow(tblXml, rowIndex) {
  const rows = tableRows(tblXml)
  if (rowIndex < 0 || rowIndex >= rows.length) {
    throw new Error(`Row index out of range: ${rowIndex} (table has ${rows.length} rows)`)
  }
  return rebuildTable(tblXml, rows.filter((_, index) => index !== rowIndex).map(row => row.outerXml))
}

/**
 * Insert a column, adding a `<w:gridCol>` and a cell in every row.
 * @param {string} tblXml
 * @param {object} [options]
 * @param {number} [options.index=0] - 0-based column index
 * @param {string[]} [options.values] - cell text per row
 * @param {number} [options.widthTwips]
 * @param {string} [options.verticalAlign]
 * @param {string} [options.alignment]
 * @returns {string}
 */
export function addColumn(tblXml, options = {}) {
  const rows = tableRows(tblXml)
  const index = Number(options.index ?? 0)
  const values = options.values || []
  const maxCols = rows.reduce((max, row) => Math.max(max, rowCells(row.outerXml).length), 0)
  if (index < 0 || index > maxCols) {
    throw new Error(`Column insertion index out of range: ${index} (table has ${maxCols} columns)`)
  }

  const grid = extractElements(tblXml, 'w:tblGrid')[0]
  const gridCols = extractElements(grid?.innerXml ?? '', 'w:gridCol')
    .map(column => Number(getAttribute(column.outerXml, 'w:w')))
  const width = options.widthTwips ?? 1440
  gridCols.splice(index, 0, width)
  const gridXml = `<w:tblGrid>${gridCols.map(value => `<w:gridCol w:w="${Math.round(value)}"/>`).join('')}</w:tblGrid>`

  let xml = tblXml
  if (grid) {
    xml = xml.slice(0, grid.index) + gridXml + xml.slice(grid.index + grid.outerXml.length)
  } else {
    const tblPr = extractElements(xml, 'w:tblPr')[0]
    const at = tblPr ? tblPr.index + tblPr.outerXml.length : (/<w:tbl\b[^>]*>/.exec(xml)?.[0].length ?? 0)
    xml = xml.slice(0, at) + gridXml + xml.slice(at)
  }

  const updatedRows = tableRows(xml)
  for (let r = 0; r < updatedRows.length; r++) {
    const cells = rowCells(updatedRows[r].outerXml)
    const insertAt = Math.min(index, cells.length)
    const newCells = cells.map(cell => cell.outerXml)
    newCells.splice(insertAt, 0, buildCellXml({
      text: values[r] ?? '',
      widthTwips: width,
      shading: options.shading,
      verticalAlign: options.verticalAlign,
      alignment: options.alignment
    }))
    xml = replaceRow(xml, r, rebuildRow(updatedRows[r].outerXml, newCells))
  }

  return xml
}

/**
 * Remove a column.
 * @param {string} tblXml
 * @param {number} colIndex
 * @returns {string}
 */
export function removeColumn(tblXml, colIndex) {
  const rows = tableRows(tblXml)
  const maxCols = rows.reduce((max, row) => Math.max(max, rowCells(row.outerXml).length), 0)
  if (colIndex < 0 || colIndex >= maxCols) {
    throw new Error(`Column index out of range: ${colIndex} (table has ${maxCols} columns)`)
  }

  let xml = tblXml
  const grid = extractElements(xml, 'w:tblGrid')[0]
  if (grid) {
    const gridCols = extractElements(grid.innerXml, 'w:gridCol')
      .map(column => Number(getAttribute(column.outerXml, 'w:w')))
    gridCols.splice(colIndex, 1)
    const gridXml = `<w:tblGrid>${gridCols.map(value => `<w:gridCol w:w="${Math.round(value)}"/>`).join('')}</w:tblGrid>`
    xml = xml.slice(0, grid.index) + gridXml + xml.slice(grid.index + grid.outerXml.length)
  }

  const updatedRows = tableRows(xml)
  for (let r = 0; r < updatedRows.length; r++) {
    const cells = rowCells(updatedRows[r].outerXml)
    if (colIndex >= cells.length) continue
    const newCells = cells.filter((_, index) => index !== colIndex).map(cell => cell.outerXml)
    if (newCells.length === 0) newCells.push(EMPTY_CELL)
    xml = replaceRow(xml, r, rebuildRow(updatedRows[r].outerXml, newCells))
  }

  return xml
}
