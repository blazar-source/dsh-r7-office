import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import { execFile, execFileSync } from 'node:child_process'
import { promisify } from 'node:util'
import { inspectPdf, repairToUnicodeCMaps } from './pdf-inspect.js'
import { parseFontListPaths, writeSanitizedFontList } from './font-substitutes.js'

const execFileAsync = promisify(execFile)

/** Minimum size that counts as a usable font list; R7 ships a 0-byte stub. */
const MIN_ALL_FONTS_BYTES = 1024

const XML_ESCAPES = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }

/** `1`, `true`, `yes` or `on` — anything else means "not set". */
function isTruthyEnv(value) {
  if (value === undefined || value === null) return false
  return ['1', 'true', 'yes', 'on'].includes(String(value).trim().toLowerCase())
}

/** Escape a path for use as XML text content. */
export function escapeXmlText(value) {
  return String(value).replace(/[&<>"']/g, (ch) => XML_ESCAPES[ch])
}

/**
 * Build the single-argument "params XML" that `x2t` accepts
 * (`x2t "path_to_params_xml"`, per its own usage text).
 *
 * `m_sAllFontsPath` is the field that decides whether DOCX/PPTX/HTML text is
 * rendered at all: without it (or with R7's 0-byte stub) the renderer has no font
 * list and emits one EMPTY fill per glyph, so the page is blank and no text object
 * is written. This is not a Cyrillic problem — Latin text is lost the same way.
 */
export function buildX2tParamsXml({ sourcePath, targetPath, allFontsPath }) {
  const parts = [
    '<?xml version="1.0" encoding="utf-8"?>',
    '<Settings>',
    `<m_sFileFrom>${escapeXmlText(sourcePath)}</m_sFileFrom>`,
    `<m_sFileTo>${escapeXmlText(targetPath)}</m_sFileTo>`
  ]
  if (allFontsPath) parts.push(`<m_sAllFontsPath>${escapeXmlText(allFontsPath)}</m_sAllFontsPath>`)
  parts.push('</Settings>')
  return parts.join('\n')
}

/**
 * Pick the best usable `AllFonts.js` out of candidate paths.
 * R7's install ships a 0-byte placeholder, so size is the discriminator.
 * @param {string[]} candidates
 * @returns {string|null}
 */
export function selectAllFontsJs(candidates) {
  let best = null
  let bestSize = 0
  for (const candidate of candidates) {
    if (!candidate) continue
    try {
      const stat = fs.statSync(candidate)
      if (!stat.isFile() || stat.size < MIN_ALL_FONTS_BYTES) continue
      if (stat.size > bestSize) { best = candidate; bestSize = stat.size }
    } catch {
      // not present — keep looking
    }
  }
  return best
}

/**
 * Candidate locations of the R7 font list. The real file is written into the user
 * profile by R7-Office Desktop's own `CApplicationFontsWorker`; the copy inside the
 * installation is a stub.
 */
export function defaultAllFontsCandidates(env = process.env, platform = os.platform(), installPath = null) {
  const candidates = []
  if (env.R7_ALL_FONTS_JS) candidates.push(env.R7_ALL_FONTS_JS)

  const userRoots = []
  if (platform === 'win32') {
    if (env.LOCALAPPDATA) userRoots.push(path.join(env.LOCALAPPDATA, 'R7-Office', 'Editors'))
    if (env.APPDATA) userRoots.push(path.join(env.APPDATA, 'R7-Office', 'Editors'))
  } else if (platform === 'darwin') {
    if (env.HOME) userRoots.push(path.join(env.HOME, 'Library', 'Application Support', 'R7-Office', 'Editors'))
  } else {
    if (env.HOME) {
      userRoots.push(path.join(env.HOME, '.local', 'share', 'R7-Office', 'Editors'))
      userRoots.push(path.join(env.HOME, '.config', 'R7-Office', 'Editors'))
    }
  }

  for (const root of userRoots) {
    candidates.push(path.join(root, 'data', 'fonts', 'AllFonts.js'))
    candidates.push(...boundedFind(root, 'AllFonts.js', 5))
  }
  if (installPath) {
    candidates.push(path.join(installPath, 'data', 'fonts', 'AllFonts.js'))
    candidates.push(path.join(installPath, 'editors', 'sdkjs', 'common', 'AllFonts.js'))
    candidates.push(...boundedFind(installPath, 'AllFonts.js', 4))
  }
  return candidates
}

/** Depth- and count-capped recursive file search (dependency-free, no glob). */
function boundedFind(root, fileName, maxDepth) {
  const out = []
  const walk = (dir, depth) => {
    if (depth > maxDepth || out.length > 50) return
    let entries
    try { entries = fs.readdirSync(dir, { withFileTypes: true }) } catch { return }
    for (const entry of entries) {
      const full = path.join(dir, entry.name)
      if (entry.isDirectory()) walk(full, depth + 1)
      else if (entry.name === fileName) out.push(full)
    }
  }
  walk(root, 0)
  return out
}

/** Throw unless x2t actually wrote a non-empty file. */
function assertProduced(targetPath) {
  if (!fs.existsSync(targetPath) || fs.statSync(targetPath).size === 0) {
    throw new Error('conversion completed but no output file was written')
  }
}

/** Actionable explanation for a PDF whose glyphs were dropped instead of drawn. */
function buildTextLossMessage(quality, allFontsPath) {
  const remedy =
    'Run R7-Office Desktop once so it generates %LOCALAPPDATA%\\R7-Office\\Editors\\data\\fonts\\AllFonts.js, ' +
    'or point the R7_ALL_FONTS_JS environment variable at a valid AllFonts.js. '

  let cause
  if (!allFontsPath) {
    cause = 'no R7 font list (AllFonts.js) was found, so the DOCX/PPTX/HTML renderer had no fonts to draw with'
  } else if (!fs.existsSync(allFontsPath)) {
    cause = `the font list ${allFontsPath} does not exist`
  } else {
    const size = fs.statSync(allFontsPath).size
    cause = size < MIN_ALL_FONTS_BYTES
      ? `the font list ${allFontsPath} is only ${size} bytes (R7's installer ships a 0-byte stub)`
      : `the font list ${allFontsPath} was supplied but the renderer still produced no glyph outlines`
  }

  return (
    `R7 x2t produced a PDF with no extractable text: ${quality.textGlyphs} text glyphs versus ` +
    `${quality.emptyFillOperators} glyph fills that painted nothing (verdict "${quality.verdict}"). Cause: ${cause}. ` +
    remedy +
    'Pass allowOutlinedPdf:true to keep the unusable PDF and inspect pdfTextQuality for details. ' +
    `Output: ${quality.filePath}`
  )
}

/**
 * Actionable explanation for a PDF whose embedded fonts draw the wrong glyphs.
 *
 * This is the metric-compatible-substitute defect: x2t resolves a character's
 * glyph id through the substitute font R7 ships (the Liberation set) while
 * embedding the REAL font, so every character whose index differs between the
 * two is drawn from the wrong slot. The text layer is unaffected, which is why
 * text extraction and a text-only check both pass.
 */
function buildGlyphMismatchMessage(quality) {
  const chain = quality.glyphChain || { fonts: [], mismatches: 0, cidsChecked: 0 }
  const offenders = chain.fonts.filter((font) => font.mismatches > 0)
  const detail = offenders
    .slice(0, 4)
    .map((font) => {
      const example = font.examples && font.examples[0]
      const sample = example
        ? ` — CID ${example.cid} "${example.character}" draws glyph ${example.drawnGid}, ` +
          `but that character is glyph ${example.expectedGid} in ${font.sourceFont}`
        : ''
      return `${font.baseFont || `object ${font.object}`}: ${font.mismatches}/${font.cidsChecked} characters${sample}`
    })
    .join('; ')

  return (
    `R7 x2t produced a PDF whose embedded fonts map text to the wrong glyphs: ` +
    `${chain.mismatches} of ${chain.cidsChecked} checked characters draw a different outline than the ` +
    `character the text layer names (${detail}). The document's text extracts correctly, so only ` +
    `rendering shows it. Cause: x2t numbers glyphs with the metric-compatible substitute in the font ` +
    `list (Liberation Sans for Arial and so on) but embeds the real font, so non-Latin text comes out ` +
    `as unrelated Latin/Greek shapes. Remedy: the converted font list already redirects each substitute ` +
    `entry at the real font; this PDF still came back inconsistent, so the substitute could not be ` +
    `redirected on this host (the real font is probably missing). ` +
    'Pass allowGlyphMismatchPdf:true to keep the PDF and inspect pdfTextQuality.glyphChain for details. ' +
    `Output: ${quality.filePath}`
  )
}

/**
 * R7 Office Environment Adapter.
 * Detects installed R7-Office components, x2t converter, and native templates.
 */
export class R7Adapter {
  constructor(customPath = null) {
    this.customPath = customPath
    this._info = null
  }

  /**
   * Discover R7-Office installation and capabilities on the host.
   * @returns {Promise<{installed: boolean, version: string|null, installPath: string|null, x2tPath: string|null, templatesPath: string|null, pluginsPath: string|null}>}
   */
  async detect() {
    if (this._info) return this._info

    // Testing seam, and a real one for users: R7_OFFICE_DISABLED makes the
    // adapter report no installation even where one exists, so the fallback
    // paths can be exercised on a developer machine and in CI parity runs.
    // It changes no detection logic — it only short-circuits it.
    if (isTruthyEnv(process.env.R7_OFFICE_DISABLED)) {
      this._info = {
        installed: false,
        version: null,
        installPath: null,
        x2tPath: null,
        templatesPath: null,
        pluginsPath: null,
        disabled: true
      }
      return this._info
    }

    const platform = os.platform()
    let installPath = this.customPath || process.env.R7_OFFICE_PATH || null
    let x2tPath = null
    let templatesPath = null
    let pluginsPath = null
    let version = null

    // 1. Try finding on Windows
    if (platform === 'win32') {
      const candidates = []
      if (installPath) candidates.push(installPath)

      // Check Program Files for R7-Office directories
      const programFiles = [
        process.env.ProgramFiles,
        process.env['ProgramFiles(x86)'],
        path.join(process.env.LOCALAPPDATA || '', 'Programs')
      ].filter(Boolean)

      for (const pf of programFiles) {
        const r7Base = path.join(pf, 'R7-Office')
        if (fs.existsSync(r7Base)) {
          try {
            const entries = fs.readdirSync(r7Base)
            for (const entry of entries) {
              if (entry.startsWith('Editors')) {
                candidates.push(path.join(r7Base, entry))
              }
            }
          } catch {
            // ignore
          }
          candidates.push(r7Base)
        }
      }

      // Version-agnostic fallbacks for a default installation layout. The
      // directory scan above already covers every "Editors-<version>" folder,
      // so nothing here may pin a specific R7 release.
      candidates.push(path.join(process.env.ProgramFiles || 'C:\\Program Files', 'R7-Office', 'Editors'))

      for (const cand of candidates) {
        const testX2t = path.join(cand, 'converter', 'x2t.exe')
        if (fs.existsSync(testX2t)) {
          installPath = cand
          x2tPath = testX2t
          const testTpl = path.join(cand, 'converter', 'empty')
          if (fs.existsSync(testTpl)) {
            templatesPath = testTpl
          }
          break
        }
      }

      const localAppData = process.env.LOCALAPPDATA
      if (localAppData) {
        const userPlugins = path.join(localAppData, 'R7-Office', 'Editors', 'data', 'sdkjs-plugins')
        if (fs.existsSync(userPlugins)) {
          pluginsPath = userPlugins
        }
      }
    } else if (platform === 'linux') {
      const linuxCandidates = [
        '/opt/r7-office/desktopeditors',
        '/usr/bin/r7-office-desktopeditors'
      ]
      for (const cand of linuxCandidates) {
        const testX2t = path.join(cand, 'converter', 'x2t')
        if (fs.existsSync(testX2t)) {
          installPath = cand
          x2tPath = testX2t
          templatesPath = path.join(cand, 'converter', 'empty')
          break
        }
      }
    } else if (platform === 'darwin') {
      const macCandidate = '/Applications/R7-Office.app/Contents/Resources'
      const testX2t = path.join(macCandidate, 'converter', 'x2t')
      if (fs.existsSync(testX2t)) {
        installPath = macCandidate
        x2tPath = testX2t
        templatesPath = path.join(macCandidate, 'converter', 'empty')
      }
    }

    // Try reading version from directory name or package
    if (installPath) {
      const match = installPath.match(/Editors-([0-9.]+)/i)
      if (match) {
        version = match[1]
      } else {
        version = 'detected'
      }
    }

    this._info = {
      installed: Boolean(x2tPath),
      version,
      installPath,
      x2tPath,
      templatesPath,
      pluginsPath,
      platform,
      // The font list the DOCX/PPTX/HTML->PDF renderer needs. R7 ships a 0-byte
      // stub inside the installation; the real one is generated into the user
      // profile the first time R7-Office Desktop runs.
      allFontsPath: selectAllFontsJs(defaultAllFontsCandidates(process.env, platform, installPath))
    }

    return this._info
  }

  /**
   * Locate the R7 font list (`AllFonts.js`) that the DOCX/PPTX/HTML renderer needs.
   * Returns null when the host has never run R7-Office Desktop (which is what
   * generates the file); callers must then treat the conversion as degraded.
   * @returns {string|null}
   */
  findAllFontsJs() {
    const info = this._info
    return selectAllFontsJs(
      defaultAllFontsCandidates(process.env, os.platform(), info ? info.installPath : this.customPath)
    )
  }

  /**
   * Convert a document. For PDF targets the result carries `pdfTextQuality` (see
   * pdf-inspect.js) and `textMapRepairs`, and a PDF whose glyphs were dropped is a
   * hard error unless `allowOutlinedPdf` is set — a silently text-less PDF is worse
   * than a failed conversion.
   *
   * @param {string} sourcePath - Absolute path to source document
   * @param {string} targetPath - Absolute path to target output document
   * @param {object} [options]
   * @param {string} [options.allFontsPath] - explicit AllFonts.js override
   * @param {boolean} [options.repairTextMaps=true] - repair malformed ToUnicode CMaps
   * @param {boolean} [options.allowOutlinedPdf=false] - accept a PDF with no text
   * @param {boolean} [options.allowGlyphMismatchPdf=false] - accept a PDF whose
   *   embedded fonts provably draw glyphs other than the ones the text names
   * @param {number} [options.timeoutMs=120000]
   * @returns {Promise<object>}
   */
  async convert(sourcePath, targetPath, options = {}) {
    const info = await this.detect()
    if (!info.x2tPath) {
      throw new Error('R7 x2t converter is not available on this system')
    }

    if (!fs.existsSync(sourcePath)) {
      throw new Error(`Source file does not exist: ${sourcePath}`)
    }

    const target = path.resolve(targetPath)
    fs.mkdirSync(path.dirname(target), { recursive: true })

    const timeoutMs = options.timeoutMs ?? 120000
    const allFontsPath = options.allFontsPath || this.findAllFontsJs()
    const fontList = this.prepareFontList(allFontsPath, { sanitize: options.sanitizeFontList !== false })
    const x2tFontListPath = fontList && fontList.path ? fontList.path : allFontsPath
    const startTime = Date.now()
    const failures = []
    let mode = null

    // Preferred path: the params-XML form with an explicit font list. The plain
    // two-argument form silently falls back to the install's 0-byte AllFonts.js
    // stub and loses every glyph outline. The list is first rewritten so that a
    // metric-compatible substitute cannot shadow the real font (see
    // font-substitutes.js) — that shadowing is what makes x2t draw Cyrillic with
    // the wrong glyphs.
    if (x2tFontListPath) {
      try {
        await this._runParamsXml(info.x2tPath, sourcePath, target, x2tFontListPath, timeoutMs)
        mode = 'params-xml'
      } catch (err) {
        failures.push(`params-xml: ${err.message}`)
        fs.rmSync(target, { force: true })
      }
    }

    if (mode === null) {
      try {
        await this._runCliArgs(info.x2tPath, sourcePath, target, timeoutMs)
        mode = 'cli-args'
      } catch (err) {
        const detail = failures.length ? ` (after ${failures.join('; ')})` : ''
        throw new Error(`R7 x2t conversion failed${detail}: ${err.message}`)
      }
    }

    const timeMs = Date.now() - startTime
    const result = {
      success: true,
      source: sourcePath,
      target,
      timeMs,
      mode,
      allFontsPath: allFontsPath || null
    }
    if (fontList && fontList.sanitized) {
      result.fontList = {
        original: allFontsPath || null,
        used: fontList.path,
        rewrites: fontList.rewrites
      }
    }

    if (path.extname(target).toLowerCase() === '.pdf') {
      let repairs = []
      if (options.repairTextMaps !== false) {
        try {
          repairs = repairToUnicodeCMaps(target).repairs
        } catch {
          repairs = []
        }
      }
      result.textMapRepairs = repairs

      const quality = this.pdfTextQuality(target, { fontPaths: fontList ? fontList.fontPaths : [] })
      result.pdfTextQuality = quality

      if (quality.textLossSuspected && !options.allowOutlinedPdf) {
        const error = new Error(buildTextLossMessage(quality, allFontsPath))
        error.pdfTextQuality = quality
        error.target = target
        throw error
      }

      // Extracting correctly is not rendering correctly: fail loudly rather than
      // hand back a PDF whose glyphs are provably not the ones the text names.
      if (quality.glyphMapInconsistent && options.allowGlyphMismatchPdf !== true) {
        const error = new Error(buildGlyphMismatchMessage(quality))
        error.pdfTextQuality = quality
        error.target = target
        throw error
      }
    }

    return result
  }

  /**
   * Read the font list x2t will be given and plan the substitute rewrite.
   * @param {string|null} allFontsPath
   * @param {object} [options]
   * @param {boolean} [options.sanitize=true] rewrite metric-compatible substitutes
   *   so they cannot shadow the real font. Disabling it is an escape hatch: the
   *   glyph-chain check still refuses to return a PDF it can prove is wrong.
   * @returns {{original: string, path: string|null, fontPaths: string[], rewrites: Array<{from:string,to:string}>, sanitized: boolean}|null}
   */
  prepareFontList(allFontsPath, options = {}) {
    if (!allFontsPath) return null
    let fontPaths
    try {
      fontPaths = parseFontListPaths(fs.readFileSync(allFontsPath, 'utf8'))
    } catch {
      return null
    }
    let sanitized = null
    if (options.sanitize !== false) {
      try {
        sanitized = writeSanitizedFontList(allFontsPath)
      } catch {
        sanitized = null
      }
    }
    if (!sanitized) {
      return { original: allFontsPath, path: null, fontPaths, rewrites: [], sanitized: false }
    }
    return {
      original: allFontsPath,
      path: sanitized.path,
      fontPaths: sanitized.fontPaths.length ? sanitized.fontPaths : fontPaths,
      rewrites: sanitized.rewrites,
      sanitized: true
    }
  }

  /**
   * Measure the text fidelity of a produced PDF (see pdf-inspect.js).
   * @param {string} pdfPath
   * @param {object} [options]
   * @param {string[]} [options.fontPaths] candidate source fonts, enabling the
   *   glyph-chain check that catches "extracts fine, renders garbage".
   * @returns {ReturnType<typeof inspectPdf>}
   */
  pdfTextQuality(pdfPath, options = {}) {
    return inspectPdf(pdfPath, options)
  }

  async _runCliArgs(x2tPath, sourcePath, targetPath, timeoutMs) {
    try {
      await execFileAsync(x2tPath, [sourcePath, targetPath], { timeout: timeoutMs, windowsHide: true })
    } catch (err) {
      throw new Error(err.message)
    }
    assertProduced(targetPath)
  }

  async _runParamsXml(x2tPath, sourcePath, targetPath, allFontsPath, timeoutMs) {
    const scratch = fs.mkdtempSync(path.join(os.tmpdir(), 'dsh-r7-x2t-'))
    const paramsPath = path.join(scratch, 'params.xml')
    try {
      fs.writeFileSync(paramsPath, buildX2tParamsXml({ sourcePath, targetPath, allFontsPath }), 'utf8')
      try {
        await execFileAsync(x2tPath, [paramsPath], { timeout: timeoutMs, windowsHide: true, cwd: scratch })
      } catch (err) {
        throw new Error(err.message)
      }
      assertProduced(targetPath)
    } finally {
      fs.rmSync(scratch, { recursive: true, force: true })
    }
  }

  /**
   * Get path to native R7 empty template or null.
   * @param {'docx'|'xlsx'|'pptx'} type
   * @param {string} [locale='ru-RU']
   * @returns {Promise<string|null>}
   */
  async getTemplatePath(type, locale = 'ru-RU') {
    const info = await this.detect()
    if (!info.templatesPath) return null

    const candidates = [
      path.join(info.templatesPath, locale, `new.${type}`),
      path.join(info.templatesPath, 'en-US', `new.${type}`),
      path.join(info.templatesPath, `new.${type}`)
    ]

    for (const cand of candidates) {
      if (fs.existsSync(cand)) {
        return cand
      }
    }
    return null
  }
}
