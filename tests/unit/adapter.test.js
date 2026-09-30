import { test, describe } from 'node:test'
import assert from 'node:assert/strict'
import { R7Adapter } from '../../src/r7/adapter.js'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'

describe('R7Adapter', () => {
  test('should detect R7 installation on host', async () => {
    const adapter = new R7Adapter()
    const info = await adapter.detect()
    assert.ok(info)
    assert.ok(typeof info.installed === 'boolean')
    if (info.installed) {
      assert.ok(info.x2tPath)
      assert.ok(fs.existsSync(info.x2tPath))
    }
  })

  test('should find native templates if installed', async () => {
    const adapter = new R7Adapter()
    const info = await adapter.detect()
    if (info.installed && info.templatesPath) {
      const docxTpl = await adapter.getTemplatePath('docx')
      assert.ok(docxTpl)
      assert.ok(fs.existsSync(docxTpl))
    }
  })

  test('should convert docx to pdf if x2t is available', async () => {
    const adapter = new R7Adapter()
    const info = await adapter.detect()
    if (info.installed) {
      const docxTpl = await adapter.getTemplatePath('docx')
      if (docxTpl) {
        const outPdf = path.join(os.tmpdir(), `test_adapter_${Date.now()}.pdf`)
        const result = await adapter.convert(docxTpl, outPdf)
        assert.ok(result.success)
        assert.ok(fs.existsSync(outPdf))
        fs.unlinkSync(outPdf)
      }
    }
  })
})
