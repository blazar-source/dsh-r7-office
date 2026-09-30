import { test, describe } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import { DocxEngine } from '../../src/r7/docx.js'
import { R7Adapter } from '../../src/r7/adapter.js'

describe('DocxEngine E2E & Integration', () => {
  const tmpDir = path.join(os.tmpdir(), `dsh_r7_test_${Date.now()}`)
  const testDocx = path.join(tmpDir, 'Отчет.docx')
  const modifiedDocx = path.join(tmpDir, 'Отчет_обновленный.docx')
  const outPdf = path.join(tmpDir, 'Отчет.pdf')

  test('setup test directory', () => {
    fs.mkdirSync(tmpDir, { recursive: true })
  })

  test('should create a rich DOCX document from template', async () => {
    const engine = new DocxEngine()
    const result = await engine.create(testDocx, {
      title: 'Годовой отчет компании',
      paragraphs: [
        { text: 'Раздел 1. Введение', style: 'Heading1', bold: true },
        'Это вступительный текст отчета компании за 2026 год.',
        { text: 'Раздел 2. Финансовые показатели', style: 'Heading1', bold: true },
        'Выручка компании выросла на 25% по сравнению с прошлым периодом.',
        { text: 'Раздел 3. Планы развития', style: 'Heading1', bold: true },
        'Старый текст планов развития, который необходимо обновить.'
      ],
      tables: [
        {
          rows: [
            ['Показатель', '2025', '2026'],
            ['Выручка, млн руб', '120.5', '150.8'],
            ['Прибыль, млн руб', '34.2', '45.1']
          ]
        }
      ]
    })

    assert.ok(result.success)
    assert.ok(fs.existsSync(testDocx))
  })

  test('should validate the newly created DOCX document', async () => {
    const engine = new DocxEngine()
    const validation = await engine.validate(testDocx)
    assert.equal(validation.valid, true)
    assert.equal(validation.errors.length, 0)
  })

  test('should inspect the document structure', async () => {
    const engine = new DocxEngine()
    const inspection = await engine.inspect(testDocx)

    assert.equal(inspection.type, 'docx')
    assert.ok(inspection.paragraphsCount >= 6)
    assert.ok(inspection.headingsCount >= 3)
    assert.equal(inspection.tablesCount, 1)
  })

  test('should read document in structured and markdown formats', async () => {
    const engine = new DocxEngine()
    const structured = await engine.read(testDocx, { format: 'structured' })
    assert.ok(structured.paragraphs.length > 0)

    const markdown = await engine.read(testDocx, { format: 'markdown' })
    assert.ok(markdown.content.includes('Годовой отчет'))
    assert.ok(markdown.content.includes('Раздел 1'))
  })

  test('should perform style-preserving text replacement', async () => {
    const engine = new DocxEngine()
    const replaceRes = await engine.replaceText(
      testDocx,
      'Старый текст планов развития, который необходимо обновить.',
      'Новый утвержденный стратегический план развития на 2026-2028 годы с интеграцией ИИ.',
      { outputPath: modifiedDocx }
    )

    assert.ok(replaceRes.success)
    assert.ok(replaceRes.matchesCount > 0)
    assert.ok(fs.existsSync(modifiedDocx))

    // Verify replacement in modified file
    const readMod = await engine.read(modifiedDocx, { format: 'markdown' })
    assert.ok(readMod.content.includes('Новый утвержденный стратегический план'))
    assert.ok(!readMod.content.includes('Старый текст планов развития'))
  })

  test('should insert table into modified document', async () => {
    const engine = new DocxEngine()
    const tableRes = await engine.table(modifiedDocx, {
      action: 'create',
      rows: [
        ['Квартал', 'Цель', 'Статус'],
        ['Q1', 'Внедрение Р7', 'Выполнено'],
        ['Q2', 'Интеграция с DSH', 'В работе']
      ],
      outputPath: modifiedDocx
    })

    assert.ok(tableRes.success)

    const inspection = await engine.inspect(modifiedDocx)
    assert.equal(inspection.tablesCount, 2)
  })

  test('should convert document to PDF via R7 x2t adapter if installed', async () => {
    const adapter = new R7Adapter()
    const info = await adapter.detect()
    if (info.installed) {
      const convRes = await adapter.convert(modifiedDocx, outPdf)
      assert.ok(convRes.success)
      assert.ok(fs.existsSync(outPdf))
      assert.ok(fs.statSync(outPdf).size > 1000)
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
