import { test } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

import {
  buildX2tParamsXml,
  escapeXmlText,
  selectAllFontsJs,
  defaultAllFontsCandidates
} from '../../src/r7/adapter.js'

function tmpDir() {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'adapter-pdf-test-'))
}

test('escapeXmlText escapes the five XML entities', () => {
  assert.equal(escapeXmlText('a&b<c>d"e\'f'), 'a&amp;b&lt;c&gt;d&quot;e&apos;f')
})

test('buildX2tParamsXml writes the x2t params document with the font list', () => {
  const xml = buildX2tParamsXml({
    sourcePath: 'C:\\in\\a&b.docx',
    targetPath: 'C:\\out\\a.pdf',
    allFontsPath: 'C:\\fonts\\AllFonts.js'
  })
  assert.match(xml, /^<\?xml version="1\.0" encoding="utf-8"\?>/)
  assert.match(xml, /<m_sFileFrom>C:\\in\\a&amp;b\.docx<\/m_sFileFrom>/)
  assert.match(xml, /<m_sFileTo>C:\\out\\a\.pdf<\/m_sFileTo>/)
  assert.match(xml, /<m_sAllFontsPath>C:\\fonts\\AllFonts\.js<\/m_sAllFontsPath>/)
})

test('buildX2tParamsXml omits m_sAllFontsPath when there is no font list', () => {
  const xml = buildX2tParamsXml({ sourcePath: '/in/a.docx', targetPath: '/out/a.pdf' })
  assert.doesNotMatch(xml, /m_sAllFontsPath/)
  assert.match(xml, /<m_sFileFrom>\/in\/a\.docx<\/m_sFileFrom>/)
})

test('selectAllFontsJs ignores missing files and R7\'s 0-byte stub', () => {
  const dir = tmpDir()
  const stub = path.join(dir, 'AllFonts.js')
  fs.writeFileSync(stub, '')
  const tiny = path.join(dir, 'tiny.js')
  fs.writeFileSync(tiny, 'x'.repeat(100))
  assert.equal(selectAllFontsJs([stub, tiny, path.join(dir, 'nope.js')]), null)
})

test('selectAllFontsJs picks the largest usable font list', () => {
  const dir = tmpDir()
  const small = path.join(dir, 'small-AllFonts.js')
  const large = path.join(dir, 'large-AllFonts.js')
  const stub = path.join(dir, 'stub-AllFonts.js')
  fs.writeFileSync(small, 'a'.repeat(2048))
  fs.writeFileSync(large, 'b'.repeat(9000))
  fs.writeFileSync(stub, '')
  assert.equal(selectAllFontsJs([stub, small, large]), large)
})

test('selectAllFontsJs tolerates a directory named AllFonts.js', () => {
  const dir = tmpDir()
  const asDir = path.join(dir, 'AllFonts.js')
  fs.mkdirSync(asDir)
  assert.equal(selectAllFontsJs([asDir]), null)
})

test('defaultAllFontsCandidates honours R7_ALL_FONTS_JS first', () => {
  const candidates = defaultAllFontsCandidates({ R7_ALL_FONTS_JS: 'D:\\custom\\AllFonts.js' }, 'win32', null)
  assert.equal(candidates[0], 'D:\\custom\\AllFonts.js')
})

test('defaultAllFontsCandidates looks in the Windows user profile and the install', () => {
  const candidates = defaultAllFontsCandidates(
    { LOCALAPPDATA: 'C:\\Users\\u\\AppData\\Local', APPDATA: 'C:\\Users\\u\\AppData\\Roaming' },
    'win32',
    'C:\\Program Files\\R7-Office\\Editors-1'
  )
  assert.ok(candidates.includes('C:\\Users\\u\\AppData\\Local\\R7-Office\\Editors\\data\\fonts\\AllFonts.js'))
  assert.ok(candidates.includes('C:\\Users\\u\\AppData\\Roaming\\R7-Office\\Editors\\data\\fonts\\AllFonts.js'))
  assert.ok(candidates.includes('C:\\Program Files\\R7-Office\\Editors-1\\data\\fonts\\AllFonts.js'))
  assert.ok(candidates.includes('C:\\Program Files\\R7-Office\\Editors-1\\editors\\sdkjs\\common\\AllFonts.js'))
})

test('defaultAllFontsCandidates uses XDG-style paths on Linux', () => {
  // The candidate builder uses the HOST path module, so compare with path.join to
  // stay separator-agnostic when this test itself runs on Windows.
  const candidates = defaultAllFontsCandidates({ HOME: '/home/u' }, 'linux', '/opt/r7-office')
  assert.ok(candidates.includes(path.join('/home/u', '.local', 'share', 'R7-Office', 'Editors', 'data', 'fonts', 'AllFonts.js')))
  assert.ok(candidates.includes(path.join('/opt/r7-office', 'data', 'fonts', 'AllFonts.js')))
  assert.ok(
    !candidates.some((candidate) => candidate.includes('Library')),
    'macOS user paths must not appear on Linux'
  )
})

test('defaultAllFontsCandidates finds a font list nested in the install tree', () => {
  const dir = tmpDir()
  const nested = path.join(dir, 'editors', 'sdkjs', 'common')
  fs.mkdirSync(nested, { recursive: true })
  const file = path.join(nested, 'AllFonts.js')
  fs.writeFileSync(file, 'x'.repeat(2000))
  const candidates = defaultAllFontsCandidates({}, 'win32', dir)
  assert.ok(candidates.includes(file))
})
