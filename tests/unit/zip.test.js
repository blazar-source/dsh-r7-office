import { test, describe } from 'node:test'
import assert from 'node:assert/strict'
import { ZipArchive } from '../../src/shared/zip.js'
import fs from 'node:fs'
import path from 'node:path'

describe('ZipArchive pure JS implementation', () => {
  test('should create and read back files in a ZIP archive', () => {
    const zip = new ZipArchive()
    zip.setText('hello.txt', 'Hello World from R7')
    zip.setText('folder/sub.xml', '<root><item id="1">Тест</item></root>')

    const buffer = zip.toBuffer()
    assert.ok(buffer.length > 0)

    const readZip = ZipArchive.fromBuffer(buffer)
    assert.equal(readZip.getText('hello.txt'), 'Hello World from R7')
    assert.equal(readZip.getText('folder/sub.xml'), '<root><item id="1">Тест</item></root>')
    assert.ok(readZip.has('hello.txt'))
    assert.ok(readZip.has('folder/sub.xml'))
    assert.ok(!readZip.has('missing.txt'))
  })
})
