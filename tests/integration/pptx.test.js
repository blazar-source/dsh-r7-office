import { test, describe } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import { PptxEngine } from '../../src/r7/pptx.js'
import { R7Adapter } from '../../src/r7/adapter.js'

describe('PptxEngine Integration', () => {
  const tmpDir = path.join(os.tmpdir(), `dsh_r7_pptx_test_${Date.now()}`)
  const testPptx = path.join(tmpDir, 'Презентация.pptx')
  const outPdf = path.join(tmpDir, 'Презентация.pdf')

  test('setup test directory', () => {
    fs.mkdirSync(tmpDir, { recursive: true })
  })

  test('should create a new PPTX presentation', async () => {
    const engine = new PptxEngine()
    const result = await engine.create(testPptx, {
      title: 'Презентация проекта DSH R7'
    })

    assert.ok(result.success)
    assert.ok(fs.existsSync(testPptx))
  })

  test('should validate PPTX file integrity', async () => {
    const engine = new PptxEngine()
    const validation = await engine.validate(testPptx)
    assert.equal(validation.valid, true)
    assert.equal(validation.errors.length, 0)
  })

  test('should inspect PPTX slides', async () => {
    const engine = new PptxEngine()
    const inspection = await engine.inspect(testPptx)
    assert.equal(inspection.type, 'pptx')
    assert.ok(inspection.slidesCount >= 1)
  })

  test('should read slide content', async () => {
    const engine = new PptxEngine()
    const result = await engine.read(testPptx)
    assert.ok(result.slides.length > 0)
  })

  test('should edit slide text', async () => {
    const engine = new PptxEngine()
    const editRes = await engine.editSlide(testPptx, {
      title: 'Обновленный заголовок презентации'
    })
    assert.ok(editRes.success)

    const inspect = await engine.inspect(testPptx)
    assert.equal(inspect.slides[0].title, 'Обновленный заголовок презентации')
  })

  test('should convert PPTX to PDF using R7 x2t', async () => {
    const adapter = new R7Adapter()
    const info = await adapter.detect()
    if (info.installed) {
      const convRes = await adapter.convert(testPptx, outPdf)
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
