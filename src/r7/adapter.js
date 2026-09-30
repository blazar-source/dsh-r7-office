import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import { execFile, execFileSync } from 'node:child_process'
import { promisify } from 'node:util'

const execFileAsync = promisify(execFile)

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
      platform
    }

    return this._info
  }

  /**
   * Convert document from one format to another using x2t.exe converter.
   * Supported formats: docx, xlsx, pptx, pdf, html, txt, csv, etc.
   * @param {string} sourcePath - Absolute path to source document
   * @param {string} targetPath - Absolute path to target output document
   * @returns {Promise<{success: boolean, source: string, target: string, timeMs: number, error?: string}>}
   */
  async convert(sourcePath, targetPath) {
    const info = await this.detect()
    if (!info.x2tPath) {
      throw new Error('R7 x2t converter is not available on this system')
    }

    if (!fs.existsSync(sourcePath)) {
      throw new Error(`Source file does not exist: ${sourcePath}`)
    }

    const startTime = Date.now()
    try {
      await execFileAsync(info.x2tPath, [sourcePath, targetPath], {
        timeout: 30000,
        windowsHide: true
      })
      const timeMs = Date.now() - startTime

      if (!fs.existsSync(targetPath)) {
        throw new Error('Conversion completed but target file was not created')
      }

      return {
        success: true,
        source: sourcePath,
        target: targetPath,
        timeMs
      }
    } catch (err) {
      throw new Error(`R7 x2t conversion failed: ${err.message}`)
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
