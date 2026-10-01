/**
 * font-substitutes.js — repair R7's font list when a metric-compatible stand-in
 * shadows the real font.
 *
 * Why this exists
 * ---------------
 * R7's generated `AllFonts.js` lists the *substitute* families it ships (the
 * Liberation set, Carlito, ...) next to the real fonts they stand in for. x2t
 * then resolves a character's GLYPH INDEX through the substitute's cmap but
 * EMBEDS the real font and takes `/W` from it. Whenever the two fonts number
 * their glyphs differently the glyph is drawn from the wrong slot:
 *
 *     Arial (real, GID space)          Liberation Sans (substitute, GID space)
 *     U+041E CYRILLIC O  -> GID  584   U+041E CYRILLIC O        -> GID 442
 *     GID 442 in Arial is LATIN SMALL LETTER L WITH CEDILLA
 *
 * so Cyrillic body text renders as Latin-Extended gibberish while the text layer
 * (ToUnicode, written from the real characters) still extracts perfectly.
 *
 * The repair is to hand x2t a font list whose substitute entries resolve to the
 * real font FILE. Every pair below is metric-compatible, so nothing about layout
 * changes; the substitute entry simply stops pointing at a different glyph space.
 * Entries whose real font is not installed are left untouched, so the rewrite is
 * a no-op on hosts that only have the substitutes (where the mismatch cannot
 * arise in the first place).
 *
 * Nothing here is R7-specific beyond the file names: the rewrite is a textual
 * replacement of quoted font paths inside the `__fonts_files` array, and it
 * reports every change it makes so the caller can log or reject them.
 */

import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import { createHash } from 'node:crypto'

/**
 * Substitute font file -> the real font file it stands in for.
 * Keys and values are basenames, matched case-insensitively.
 */
export const SUBSTITUTE_TARGETS = new Map([
  // Liberation Sans is metric-compatible with Arial.
  ['liberationsans-regular.ttf', 'arial.ttf'],
  ['liberationsans-bold.ttf', 'arialbd.ttf'],
  ['liberationsans-italic.ttf', 'ariali.ttf'],
  ['liberationsans-bolditalic.ttf', 'arialbi.ttf'],
  // Liberation Sans Narrow <-> Arial Narrow.
  ['liberationsansnarrow-regular.ttf', 'arialn.ttf'],
  ['liberationsansnarrow-bold.ttf', 'arialnb.ttf'],
  ['liberationsansnarrow-italic.ttf', 'arialni.ttf'],
  ['liberationsansnarrow-bolditalic.ttf', 'arialnbi.ttf'],
  // Liberation Serif is metric-compatible with Times New Roman.
  ['liberationserif-regular.ttf', 'times.ttf'],
  ['liberationserif-bold.ttf', 'timesbd.ttf'],
  ['liberationserif-italic.ttf', 'timesi.ttf'],
  ['liberationserif-bolditalic.ttf', 'timesbi.ttf'],
  // Liberation Mono is metric-compatible with Courier New.
  ['liberationmono-regular.ttf', 'cour.ttf'],
  ['liberationmono-bold.ttf', 'courbd.ttf'],
  ['liberationmono-italic.ttf', 'couri.ttf'],
  ['liberationmono-bolditalic.ttf', 'courbi.ttf'],
  // Carlito is metric-compatible with Calibri.
  ['carlito-regular.ttf', 'calibri.ttf'],
  ['carlito-bold.ttf', 'calibrib.ttf'],
  ['carlito-italic.ttf', 'calibrii.ttf'],
  ['carlito-bolditalic.ttf', 'calibriz.ttf']
])

const FONT_EXTENSIONS = /\.(ttf|otf|ttc|otc)$/i

/** Normalise a font path to a platform-independent basename. */
export function fontBasename(p) {
  return String(p).split(/[\\/]/).pop()
}

/** Directory part of a font path, in its original separator style. */
function fontDirname(p) {
  const s = String(p)
  const i = Math.max(s.lastIndexOf('/'), s.lastIndexOf('\\'))
  return i === -1 ? '' : s.slice(0, i)
}

/**
 * Every quoted font path inside a `__fonts_files` style array.
 * Only string literals that look like font files are returned.
 * @returns {string[]}
 */
export function parseFontListPaths(text) {
  const out = []
  const seen = new Set()
  for (const m of String(text).matchAll(/"((?:[^"\\]|\\.)*)"/g)) {
    const value = m[1]
    if (!FONT_EXTENSIONS.test(value)) continue
    if (seen.has(value)) continue
    seen.add(value)
    out.push(value)
  }
  return out
}

/** Directories worth indexing when looking for the real font. */
export function defaultFontDirs(env = process.env, platform = os.platform()) {
  const dirs = []
  if (platform === 'win32') {
    const windir = env.WINDIR || env.SystemRoot || 'C:\\Windows'
    dirs.push(path.join(windir, 'Fonts'))
    if (env.LOCALAPPDATA) dirs.push(path.join(env.LOCALAPPDATA, 'Microsoft', 'Windows', 'Fonts'))
  } else if (platform === 'darwin') {
    dirs.push('/Library/Fonts', '/System/Library/Fonts', '/System/Library/Fonts/Supplemental')
    if (env.HOME) dirs.push(path.join(env.HOME, 'Library', 'Fonts'))
  } else {
    dirs.push('/usr/share/fonts', '/usr/local/share/fonts', '/usr/share/fonts/truetype')
    if (env.HOME) dirs.push(path.join(env.HOME, '.fonts'), path.join(env.HOME, '.local', 'share', 'fonts'))
  }
  return dirs
}

/**
 * Index font files by lower-cased basename across the given directories
 * (recursively, but shallow enough for system font trees).
 * @returns {Map<string, string>}
 */
export function indexFontDirs(dirs, { readdir = fs.readdirSync, maxDepth = 3, maxEntries = 40000 } = {}) {
  const index = new Map()
  let entries = 0
  const walk = (dir, depth) => {
    if (depth > maxDepth || entries > maxEntries) return
    let items
    try { items = readdir(dir, { withFileTypes: true }) } catch { return }
    for (const item of items) {
      entries++
      if (entries > maxEntries) return
      const full = path.join(dir, item.name)
      if (item.isDirectory()) { walk(full, depth + 1); continue }
      if (!FONT_EXTENSIONS.test(item.name)) continue
      const key = item.name.toLowerCase()
      if (!index.has(key)) index.set(key, full)
    }
  }
  for (const dir of new Set(dirs)) walk(dir, 0)
  return index
}

/**
 * Plan the font-list rewrite. Pure except for the directory index it is handed.
 *
 * @param {string} text AllFonts.js contents
 * @param {object} options
 * @param {Map<string,string>} options.fontIndex lower-cased basename -> path
 * @returns {{rewrites: Array<{from: string, to: string}>, fontPaths: string[], unchanged: string[]}}
 */
export function planFontSubstitutions(text, { fontIndex } = {}) {
  const fontPaths = parseFontListPaths(text)
  const rewrites = []
  const unchanged = []
  for (const p of fontPaths) {
    const targetName = SUBSTITUTE_TARGETS.get(fontBasename(p).toLowerCase())
    if (!targetName) continue
    const replacement = fontIndex ? fontIndex.get(targetName.toLowerCase()) : null
    if (!replacement) { unchanged.push(p); continue }
    if (replacement === p) continue
    rewrites.push({ from: p, to: replacement })
  }
  return { rewrites, fontPaths, unchanged }
}

/**
 * Rewrite a font list so each substitute entry points at the real font it stands
 * in for. Returns the original text unchanged when there is nothing to do.
 *
 * @param {string} text
 * @param {object} [options]
 * @param {Map<string,string>} [options.fontIndex]
 * @param {string[]} [options.fontDirs] extra directories to index
 * @returns {{text: string, changed: boolean, rewrites: Array<{from:string,to:string}>, fontPaths: string[], unresolved: string[]}}
 */
export function sanitizeFontList(text, options = {}) {
  const source = String(text)
  const dirs = [...(options.fontDirs || [])]
  for (const p of parseFontListPaths(source)) {
    const dir = fontDirname(p)
    if (dir) dirs.push(dir)
  }
  dirs.push(...defaultFontDirs())
  const fontIndex = options.fontIndex || indexFontDirs(dirs)
  const plan = planFontSubstitutions(source, { fontIndex })

  let out = source
  for (const { from, to } of plan.rewrites) {
    out = out.split(`"${from}"`).join(`"${to}"`)
  }
  const changed = out !== source
  return {
    text: out,
    changed,
    rewrites: plan.rewrites,
    fontPaths: changed ? parseFontListPaths(out) : plan.fontPaths,
    unresolved: plan.unchanged
  }
}

/**
 * Write a sanitized font list next to the OS temp dir, keyed by content so
 * repeated conversions reuse one file. Returns null when there is nothing to
 * rewrite, so callers can keep using the original list.
 *
 * @param {string} originalPath
 * @param {object} [options]
 * @param {string} [options.cacheDir]
 * @returns {{path: string, rewrites: Array<{from:string,to:string}>, fontPaths: string[]}|null}
 */
export function writeSanitizedFontList(originalPath, options = {}) {
  let text
  try { text = fs.readFileSync(originalPath, 'utf8') } catch { return null }
  const sanitized = sanitizeFontList(text, options)
  if (!sanitized.changed) return null
  const cacheDir = options.cacheDir || path.join(os.tmpdir(), 'dsh-r7-fontlist')
  fs.mkdirSync(cacheDir, { recursive: true })
  const digest = createHash('sha1').update(sanitized.text).digest('hex').slice(0, 16)
  const target = path.join(cacheDir, `AllFonts-${digest}.js`)
  if (!fs.existsSync(target)) {
    const tmp = `${target}.${process.pid}.tmp`
    fs.writeFileSync(tmp, sanitized.text, 'utf8')
    fs.renameSync(tmp, target)
  }
  return { path: target, rewrites: sanitized.rewrites, fontPaths: sanitized.fontPaths }
}
