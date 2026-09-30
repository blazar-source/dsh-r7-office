/**
 * File end-to-end workflow test.
 *
 * Exercises the complete agent-facing scenario for every supported format
 * through the MCP tool surface, then proves R7-Office itself accepts the
 * result by re-opening each modified document with the native x2t engine.
 */

import { test, describe, before, after } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import { R7McpServer } from '../../src/mcp/server.js'
import { R7Adapter } from '../../src/r7/adapter.js'
import { ZipArchive } from '../../src/shared/zip.js'

const tmpDir = path.join(os.tmpdir(), `dsh_r7_file_e2e_${Date.now()}`)

/** Minimal PDF signature check: "%PDF-" within the first kilobyte. */
function looksLikePdf(filePath) {
  const head = fs.readFileSync(filePath).subarray(0, 1024)
  return head.includes(Buffer.from('%PDF-'))
}

/** Minimal OOXML signature check: a ZIP local header. */
function looksLikeZip(filePath) {
  const head = fs.readFileSync(filePath).subarray(0, 4)
  return head.readUInt32LE(0) === 0x04034b50
}

describe('File E2E: full document workflow', () => {
  let server
  let adapter
  let r7Available = false

  /** Invoke an MCP tool by name and return its parsed JSON payload. */
  async function call(name, args) {
    const res = await server.handleMessage({
      jsonrpc: '2.0',
      id: 1,
      method: 'tools/call',
      params: { name, arguments: args }
    })
    assert.ok(res.result, `${name} returned a result`)
    const text = res.result.content[0].text
    assert.ok(!res.result.isError, `${name} failed: ${text}`)
    try {
      return JSON.parse(text)
    } catch {
      return text
    }
  }

  before(async () => {
    fs.mkdirSync(tmpDir, { recursive: true })
    server = new R7McpServer()
    adapter = new R7Adapter()
    const info = await adapter.detect()
    r7Available = info.installed
  })

  after(() => {
    try {
      fs.rmSync(tmpDir, { recursive: true, force: true })
    } catch {
      /* ignore */
    }
  })

  describe('DOCX: report update scenario', () => {
    const report = path.join(tmpDir, 'Отчет.docx')
    const reportV2 = path.join(tmpDir, 'Отчет_v2.docx')
    const reportPdf = path.join(tmpDir, 'Отчет_v2.pdf')

    test('creates the source report', async () => {
      const created = await call('r7_create', {
        filePath: report,
        title: 'Годовой отчёт компании',
        paragraphs: [
          { text: 'Раздел 1. Введение', style: 'Heading1', bold: true },
          'Вводный текст отчёта за отчётный период.',
          { text: 'Раздел 2. Показатели', style: 'Heading1', bold: true },
          'Операционная прибыль выросла на 25 процентов.',
          { text: 'Раздел 3. Планы развития', style: 'Heading1', bold: true },
          'Черновой текст раздела три, подлежащий обновлению.'
        ]
      })
      assert.equal(created.success, true)
      assert.ok(looksLikeZip(report))
    })

    test('inspects the report outline', async () => {
      const info = await call('r7_inspect', { filePath: report })
      assert.equal(info.type, 'docx')
      assert.ok(info.paragraphsCount >= 5, `paragraphs: ${info.paragraphsCount}`)
      assert.ok(info.headingsCount >= 3, `headings: ${info.headingsCount}`)
      assert.ok(info.headings.some(h => h.text.includes('Раздел 3')))
    })

    test('reads the document as markdown', async () => {
      const md = await call('r7_read', { filePath: report, format: 'markdown' })
      assert.ok(md.content.includes('Годовой отчёт компании'))
      assert.ok(md.content.includes('Черновой текст раздела три'))
    })

    test('updates section 3 preserving formatting, into a new version', async () => {
      const result = await call('r7_replace', {
        filePath: report,
        search: 'Черновой текст раздела три, подлежащий обновлению.',
        replace: 'Утверждён план расширения и автоматизации документооборота на 2026-2028 годы.',
        outputPath: reportV2
      })
      assert.ok(result.matchesCount >= 1, 'the replacement matched')
      assert.ok(fs.existsSync(reportV2), 'a new version was written')
      assert.ok(fs.existsSync(report), 'the original file is preserved')
    })

    test('the replacement preserved the source run formatting', async () => {
      const before = await ZipArchive.fromFile(report)
      const after = await ZipArchive.fromFile(reportV2)

      const beforeDoc = before.getText('word/document.xml')
      const afterDoc = after.getText('word/document.xml')

      // The paragraph style must survive the edit.
      assert.ok(afterDoc.includes('Heading1'), 'heading styles are preserved')

      // Only the document part may differ; styles, fonts and theme must not.
      for (const name of before.list()) {
        if (name === 'word/document.xml') continue
        assert.ok(
          Buffer.compare(before.entries.get(name).raw, after.entries.get(name).raw) === 0,
          `"${name}" must be untouched by the text replacement`
        )
      }
      assert.notEqual(beforeDoc, afterDoc, 'the document part changed')
    })

    test('adds a summary table', async () => {
      await call('r7_table', {
        filePath: reportV2,
        action: 'create',
        rows: [
          ['Направление', 'Срок', 'Ответственный'],
          ['Внедрение Р7-Офис', 'Q2 2026', 'ИТ-департамент'],
          ['Интеграция ИИ-агентов', 'Q3 2026', 'Команда DSH']
        ],
        outputPath: reportV2
      })

      const info = await call('r7_inspect', { filePath: reportV2 })
      assert.equal(info.tablesCount, 1)
    })

    test('inspecting the table returns its cells', async () => {
      const table = await call('r7_table', {
        filePath: reportV2,
        action: 'inspect',
        tableIndex: 0
      })
      assert.equal(table.rowCount, 3)
      assert.equal(table.colCount, 3)
      assert.equal(table.data[1][0], 'Внедрение Р7-Офис')
    })

    test('validates the modified document', async () => {
      const result = await call('r7_validate', { filePath: reportV2 })
      assert.equal(result.valid, true, JSON.stringify(result.errors))
    })

    test('R7 re-opens and renders the modified document to PDF', async (t) => {
      if (!r7Available) {
        t.skip('R7-Office installation not available on this host')
        return
      }
      // A successful x2t render is proof that the R7 engine parsed the edited
      // package: a damaged document would fail here instead.
      const converted = await call('r7_convert', {
        sourcePath: reportV2,
        targetPath: reportPdf
      })
      assert.equal(converted.success, true)
      assert.ok(fs.existsSync(reportPdf))
      assert.ok(looksLikePdf(reportPdf), 'output must be a PDF')
      assert.ok(fs.statSync(reportPdf).size > 500)
    })
  })

  describe('XLSX: budget workflow', () => {
    const budget = path.join(tmpDir, 'Бюджет.xlsx')
    const budgetPdf = path.join(tmpDir, 'Бюджет.pdf')

    test('creates a spreadsheet with data', async () => {
      const created = await call('r7_create', {
        filePath: budget,
        sheets: [{
          name: 'Бюджет',
          data: [
            ['Категория', 'План', 'Факт'],
            ['Маркетинг', 50000, 48000],
            ['Разработка', 120000, 115000],
            ['Офис', 30000, 32000]
          ]
        }]
      })
      assert.equal(created.success, true)
      assert.ok(looksLikeZip(budget))
    })

    test('inspects the workbook', async () => {
      const info = await call('r7_inspect', { filePath: budget })
      assert.equal(info.type, 'xlsx')
      assert.ok(info.sheetsCount >= 1)
    })

    test('reads a cell range', async () => {
      const data = await call('r7_sheet_read', { filePath: budget, range: 'A1:C4' })
      assert.equal(data.data.length, 4)
      assert.equal(data.data[0][0], 'Категория')
      assert.equal(data.data[1][0], 'Маркетинг')
      assert.equal(data.data[1][1], 50000)
    })

    test('writes formulas into the difference column', async () => {
      const result = await call('r7_sheet_write', {
        filePath: budget,
        cells: [
          { ref: 'D2', formula: '=B2-C2' },
          { ref: 'D3', formula: '=B3-C3' },
          { ref: 'D4', formula: '=B4-C4' }
        ]
      })
      assert.equal(result.updatedCells, 3)
    })

    test('inserts a formula through the dedicated tool', async () => {
      const result = await call('r7_sheet_formula', {
        filePath: budget,
        cell: 'D1',
        formula: '=SUM(D2:D4)'
      })
      assert.ok(result.updatedCells >= 1)
    })

    test('the formulas are present in the sheet XML', async () => {
      const zip = await ZipArchive.fromFile(budget)
      const sheetXml = zip.getText('xl/worksheets/sheet1.xml')
      assert.ok(sheetXml.includes('<f>B2-C2</f>'), 'formula D2 is stored')
      assert.ok(sheetXml.includes('<f>SUM(D2:D4)</f>'), 'formula D1 is stored')
    })

    test('validates the workbook', async () => {
      const result = await call('r7_validate', { filePath: budget })
      assert.equal(result.valid, true, JSON.stringify(result.errors))
    })

    test('R7 re-opens and renders the workbook to PDF', async (t) => {
      if (!r7Available) {
        t.skip('R7-Office installation not available on this host')
        return
      }
      const converted = await call('r7_convert', {
        sourcePath: budget,
        targetPath: budgetPdf
      })
      assert.equal(converted.success, true)
      assert.ok(looksLikePdf(budgetPdf))
    })
  })

  describe('PPTX: presentation workflow', () => {
    const deck = path.join(tmpDir, 'Презентация.pptx')
    const deckPdf = path.join(tmpDir, 'Презентация.pdf')

    test('creates a presentation', async () => {
      const created = await call('r7_create', {
        filePath: deck,
        title: 'Презентация проекта DSH R7'
      })
      assert.equal(created.success, true)
      assert.ok(looksLikeZip(deck))
    })

    test('inspects the slides', async () => {
      const info = await call('r7_inspect', { filePath: deck })
      assert.equal(info.type, 'pptx')
      assert.ok(info.slidesCount >= 1)
    })

    test('edits the slide title', async () => {
      const result = await call('r7_slide_edit', {
        filePath: deck,
        slideIndex: 0,
        title: 'Обновлённый заголовок презентации'
      })
      assert.equal(result.success, true)

      const info = await call('r7_inspect', { filePath: deck })
      assert.equal(info.slides[0].title, 'Обновлённый заголовок презентации')
    })

    test('validates the presentation', async () => {
      const result = await call('r7_validate', { filePath: deck })
      assert.equal(result.valid, true, JSON.stringify(result.errors))
    })

    test('R7 re-opens and renders the presentation to PDF', async (t) => {
      if (!r7Available) {
        t.skip('R7-Office installation not available on this host')
        return
      }
      const converted = await call('r7_convert', {
        sourcePath: deck,
        targetPath: deckPdf
      })
      assert.equal(converted.success, true)
      assert.ok(looksLikePdf(deckPdf))
    })
  })

  describe('error handling', () => {
    test('a missing file produces a tool error, not a crash', async () => {
      const res = await server.handleMessage({
        jsonrpc: '2.0',
        id: 9,
        method: 'tools/call',
        params: { name: 'r7_inspect', arguments: { filePath: path.join(tmpDir, 'nope.docx') } }
      })
      assert.equal(res.result.isError, true)
      assert.ok(res.result.content[0].text.length > 0)
    })

    test('an unsupported extension is refused', async () => {
      const res = await server.handleMessage({
        jsonrpc: '2.0',
        id: 10,
        method: 'tools/call',
        params: { name: 'r7_inspect', arguments: { filePath: path.join(tmpDir, 'file.txt') } }
      })
      assert.equal(res.result.isError, true)
      assert.match(res.result.content[0].text, /Unsupported document extension/)
    })

    test('an out-of-range paragraph index is refused', async () => {
      const target = path.join(tmpDir, 'range.docx')
      await call('r7_create', { filePath: target, paragraphs: ['only one'] })
      const res = await server.handleMessage({
        jsonrpc: '2.0',
        id: 11,
        method: 'tools/call',
        params: {
          name: 'r7_edit',
          arguments: { filePath: target, paragraphIndex: 999, newText: 'x' }
        }
      })
      assert.equal(res.result.isError, true)
      assert.match(res.result.content[0].text, /out of range/)
    })
  })
})
