import { test, describe } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import { XlsxEngine } from '../../src/r7/xlsx.js'
import { R7Adapter } from '../../src/r7/adapter.js'

describe('XlsxEngine Integration', () => {
  const tmpDir = path.join(os.tmpdir(), `dsh_r7_xlsx_test_${Date.now()}`)
  const testXlsx = path.join(tmpDir, 'Финансы.xlsx')
  const outPdf = path.join(tmpDir, 'Финансы.pdf')

  test('setup test directory', () => {
    fs.mkdirSync(tmpDir, { recursive: true })
  })

  test('should create a new XLSX spreadsheet', async () => {
    const engine = new XlsxEngine()
    const result = await engine.create(testXlsx, {
      sheets: [
        {
          name: 'Бюджет',
          data: [
            ['Категория', 'План', 'Факт', 'Разница'],
            ['Маркетинг', 50000, 48000, ''],
            ['Разработка', 120000, 115000, ''],
            ['Офис', 30000, 32000, '']
          ]
        }
      ]
    })

    assert.ok(result.success)
    assert.ok(fs.existsSync(testXlsx))
  })

  test('should validate XLSX file integrity', async () => {
    const engine = new XlsxEngine()
    const validation = await engine.validate(testXlsx)
    assert.equal(validation.valid, true)
    assert.equal(validation.errors.length, 0)
  })

  test('should inspect XLSX sheets', async () => {
    const engine = new XlsxEngine()
    const inspection = await engine.inspect(testXlsx)
    assert.equal(inspection.type, 'xlsx')
    assert.ok(inspection.sheetsCount >= 1)
  })

  test('should read range from XLSX', async () => {
    const engine = new XlsxEngine()
    const result = await engine.read(testXlsx, { range: 'A1:C4' })
    assert.equal(result.data.length, 4)
    assert.equal(result.data[0][0], 'Категория')
    assert.equal(result.data[1][0], 'Маркетинг')
  })

  test('should write formula into XLSX cell', async () => {
    const engine = new XlsxEngine()
    const writeRes = await engine.write(testXlsx, {
      cells: [
        { ref: 'D2', formula: '=B2-C2' },
        { ref: 'D3', formula: '=B3-C3' },
        { ref: 'D4', formula: '=B4-C4' }
      ]
    })

    assert.ok(writeRes.success)
    assert.equal(writeRes.updatedCells, 3)
  })

  test('should convert XLSX to PDF using R7 x2t', async () => {
    const adapter = new R7Adapter()
    const info = await adapter.detect()
    if (info.installed) {
      const convRes = await adapter.convert(testXlsx, outPdf)
      assert.ok(convRes.success)
      assert.ok(fs.existsSync(outPdf))
      assert.ok(fs.statSync(outPdf).size > 500)
    }
  })

  test('cleanup', () => {
    try {
      fs.rmSync(tmpDir, { recursive: true, force: true })
    } catch {
      // ignore
    }
  })
})
