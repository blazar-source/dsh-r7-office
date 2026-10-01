import { test } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

import {
  SUBSTITUTE_TARGETS,
  fontBasename,
  parseFontListPaths,
  planFontSubstitutions,
  sanitizeFontList,
  writeSanitizedFontList,
  indexFontDirs
} from '../../src/r7/font-substitutes.js'

function tmpDir() {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-font-list-'))
}

/** A miniature stand-in for R7's generated AllFonts.js. */
function fontListText(entries) {
  return [
    'window["__all_fonts_js_version__"] = 2;',
    '',
    'window["__fonts_files"] = [',
    ...entries.map((p) => `"${p}",`),
    '""',
    '];',
    '',
    'window["g_fonts_selection_bin"] = "QUJD";'
  ].join('\n')
}

test('fontBasename handles both slash styles', () => {
  assert.equal(fontBasename('C:/Program Files/R7-Office/fonts/ext/LiberationSans-Regular.ttf'), 'LiberationSans-Regular.ttf')
  assert.equal(fontBasename('C:\\WINDOWS\\Fonts\\ARIAL.TTF'), 'ARIAL.TTF')
  assert.equal(fontBasename('arial.ttf'), 'arial.ttf')
})

test('parseFontListPaths returns only quoted font file paths', () => {
  const text = fontListText([
    'C:/Program Files/R7-Office/Editors/fonts/ext/LiberationSans-Regular.ttf',
    'C:/WINDOWS/Fonts/ARIAL.TTF',
    'C:/Program Files/R7-Office/Editors/fonts/ASANA.TTC'
  ]) + '\nwindow["__fonts_ranges"] = [0, 1, 2];'
  assert.deepEqual(parseFontListPaths(text), [
    'C:/Program Files/R7-Office/Editors/fonts/ext/LiberationSans-Regular.ttf',
    'C:/WINDOWS/Fonts/ARIAL.TTF',
    'C:/Program Files/R7-Office/Editors/fonts/ASANA.TTC'
  ])
})

test('every supported substitute maps to a real font file', () => {
  for (const [from, to] of SUBSTITUTE_TARGETS) {
    assert.match(from, /\.(ttf|otf)$/i)
    assert.match(to, /\.(ttf|otf|ttc)$/i)
  }
  assert.equal(SUBSTITUTE_TARGETS.get('liberationsans-regular.ttf'), 'arial.ttf')
  assert.equal(SUBSTITUTE_TARGETS.get('liberationserif-bold.ttf'), 'timesbd.ttf')
  assert.equal(SUBSTITUTE_TARGETS.get('liberationmono-regular.ttf'), 'cour.ttf')
  assert.equal(SUBSTITUTE_TARGETS.get('carlito-regular.ttf'), 'calibri.ttf')
})

test('planFontSubstitutions rewrites only substitutes whose real font is present', () => {
  const text = fontListText([
    'C:/Program Files/R7-Office/Editors/fonts/ext/LiberationSans-Regular.ttf',
    'C:/Program Files/R7-Office/Editors/fonts/ext/LiberationSerif-Bold.ttf',
    'C:/Program Files/R7-Office/Editors/fonts/ASANA.TTC',
    'C:/WINDOWS/Fonts/ARIAL.TTF'
  ])
  const fontIndex = new Map([
    ['arial.ttf', 'C:\\WINDOWS\\Fonts\\arial.ttf']
    // timesbd.ttf deliberately absent: the host has no such real font.
  ])
  const plan = planFontSubstitutions(text, { fontIndex })
  assert.deepEqual(plan.rewrites, [
    { from: 'C:/Program Files/R7-Office/Editors/fonts/ext/LiberationSans-Regular.ttf', to: 'C:\\WINDOWS\\Fonts\\arial.ttf' }
  ])
  assert.deepEqual(plan.unchanged, ['C:/Program Files/R7-Office/Editors/fonts/ext/LiberationSerif-Bold.ttf'])
  assert.equal(plan.fontPaths.length, 4)
})

test('sanitizeFontList rewrites in place and leaves every other entry alone', () => {
  const entries = [
    'C:/Program Files/R7-Office/Editors/fonts/ext/LiberationSans-Regular.ttf',
    'C:/Program Files/R7-Office/Editors/fonts/ASANA.TTC',
    'C:/WINDOWS/Fonts/ARIAL.TTF'
  ]
  const text = fontListText(entries)
  const fontIndex = new Map([['arial.ttf', 'C:\\WINDOWS\\Fonts\\arial.ttf']])
  const result = sanitizeFontList(text, { fontIndex })

  assert.equal(result.changed, true)
  assert.deepEqual(result.fontPaths, [
    'C:\\WINDOWS\\Fonts\\arial.ttf',
    'C:/Program Files/R7-Office/Editors/fonts/ASANA.TTC',
    'C:/WINDOWS/Fonts/ARIAL.TTF'
  ])
  assert.equal(parseFontListPaths(result.text).length, entries.length, 'the array keeps its shape')
  assert.match(result.text, /window\["__all_fonts_js_version__"\] = 2;/)
  assert.match(result.text, /window\["g_fonts_selection_bin"\] = "QUJD";/, 'unrelated keys survive verbatim')
  assert.doesNotMatch(result.text, /LiberationSans-Regular\.ttf/)
})

test('sanitizeFontList is a no-op when the real fonts are not installed', () => {
  const text = fontListText([
    'C:/Program Files/R7-Office/Editors/fonts/ext/LiberationSans-Regular.ttf',
    'C:/Program Files/R7-Office/Editors/fonts/ASANA.TTC'
  ])
  const result = sanitizeFontList(text, { fontIndex: new Map() })
  assert.equal(result.changed, false)
  assert.equal(result.text, text)
  assert.deepEqual(result.rewrites, [])
  assert.deepEqual(result.fontPaths, [
    'C:/Program Files/R7-Office/Editors/fonts/ext/LiberationSans-Regular.ttf',
    'C:/Program Files/R7-Office/Editors/fonts/ASANA.TTC'
  ])
})

test('sanitizeFontList is a no-op on a list with no substitutes at all', () => {
  const text = fontListText(['C:/WINDOWS/Fonts/ARIAL.TTF', 'C:/WINDOWS/Fonts/times.ttf'])
  const result = sanitizeFontList(text, { fontIndex: new Map([['arial.ttf', 'C:/WINDOWS/Fonts/arial.ttf']]) })
  assert.equal(result.changed, false)
})

test('writeSanitizedFontList caches one file per rewritten list', () => {
  const dir = tmpDir()
  const original = path.join(dir, 'AllFonts.js')
  fs.writeFileSync(original, fontListText([
    'C:/Program Files/R7-Office/Editors/fonts/ext/LiberationSans-Bold.ttf',
    'C:/Program Files/R7-Office/Editors/fonts/ext/LiberationMono-Regular.ttf'
  ]), 'utf8')
  const cacheDir = path.join(dir, 'cache')
  const fontIndex = new Map([
    ['arialbd.ttf', 'C:\\WINDOWS\\Fonts\\arialbd.ttf'],
    ['cour.ttf', 'C:\\WINDOWS\\Fonts\\cour.ttf']
  ])

  const first = writeSanitizedFontList(original, { fontIndex, cacheDir })
  assert.ok(first, 'a list with substitutes must produce a rewritten copy')
  assert.equal(first.rewrites.length, 2)
  assert.ok(fs.existsSync(first.path))
  const written = fs.readFileSync(first.path, 'utf8')
  assert.match(written, /C:\\\\WINDOWS\\\\Fonts\\\\arialbd\.ttf|C:\\WINDOWS\\Fonts\\arialbd\.ttf/)
  assert.equal(parseFontListPaths(written).length, 2)

  const second = writeSanitizedFontList(original, { fontIndex, cacheDir })
  assert.equal(second.path, first.path, 'the same content must reuse the same cache file')

  // The original file must never be modified in place.
  assert.match(fs.readFileSync(original, 'utf8'), /LiberationSans-Bold\.ttf/)
})

test('writeSanitizedFontList returns null when there is nothing to rewrite', () => {
  const dir = tmpDir()
  const original = path.join(dir, 'AllFonts.js')
  fs.writeFileSync(original, fontListText(['C:/WINDOWS/Fonts/ARIAL.TTF']), 'utf8')
  assert.equal(writeSanitizedFontList(original, { fontIndex: new Map(), cacheDir: path.join(dir, 'c') }), null)
  assert.equal(writeSanitizedFontList(path.join(dir, 'missing.js'), { fontIndex: new Map() }), null)
})

test('indexFontDirs indexes font files by lower-cased basename', () => {
  const index = indexFontDirs(['/fonts'], {
    readdir: (dir) => {
      if (dir === '/fonts') {
        return [
          { name: 'ARIAL.TTF', isDirectory: () => false, isFile: () => true },
          { name: 'ext', isDirectory: () => true, isFile: () => false }
        ]
      }
      return [{ name: 'LiberationSans-Regular.ttf', isDirectory: () => false, isFile: () => true }]
    }
  })
  assert.equal(index.get('arial.ttf'), path.join('/fonts', 'ARIAL.TTF'))
  assert.equal(index.get('liberationsans-regular.ttf'), path.join('/fonts', 'ext', 'LiberationSans-Regular.ttf'))
})
