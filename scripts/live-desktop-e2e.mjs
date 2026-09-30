/**
 * Live R7-Office Desktop bridge end-to-end test.
 *
 * This exercises the real integration chain against an installed R7-Office:
 *
 *   R7 Desktop editor  ->  bridge plugin (WebSocket)  ->  DSH MCP bridge
 *                      ->  document modification     ->  save on disk
 *                      ->  reopen + validate + render via the R7 x2t engine
 *
 * R7 Desktop is a CEF application, so it is launched with
 * `--remote-debugging-port` and driven through the DevTools protocol. The
 * editor plugin is started programmatically through the editor's own
 * `pluginsManager`, not by clicking through the UI.
 *
 * The test skips (exit 0) when R7-Office or a live editor is unavailable, so
 * it never fails a machine that cannot host the integration.
 */

import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import { fileURLToPath } from 'node:url'
import { spawn } from 'node:child_process'
import { DesktopBridge } from '../src/mcp/desktop-bridge.js'
import { R7Adapter } from '../src/r7/adapter.js'
import { DocxEngine } from '../src/r7/docx.js'
import { ZipArchive } from '../src/shared/zip.js'
import { waitForCdp, listTargets, CdpSession, delay } from './cdp.mjs'

const here = path.dirname(fileURLToPath(import.meta.url))
const repoRoot = path.resolve(here, '..')

const BRIDGE_PORT = Number(process.env.DSH_R7_BRIDGE_PORT || 7888)
const CDP_PORT = Number(process.env.R7_CDP_PORT || 9222)
const PLUGIN_GUID = 'asc.{D5B29457-194D-4E9A-A37F-02D739818FE1}'

const MARKER = 'DSH-BRIDGE-LIVE-MARKER'

const results = []
function check(name, ok, detail = '') {
  results.push({ name, ok, detail })
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? '  — ' + detail : ''}`)
}

function skip(reason) {
  console.log(`SKIP  ${reason}`)
  console.log('\nRESULT: SKIPPED')
  process.exit(0)
}

/** Copy the bridge plugin into the R7 user plugins directory. */
function installPlugin(pluginsPath) {
  const target = path.join(pluginsPath, PLUGIN_GUID.replace('asc.', ''))
  fs.mkdirSync(target, { recursive: true })
  const src = path.join(repoRoot, 'desktop-bridge')
  for (const name of fs.readdirSync(src)) {
    const from = path.join(src, name)
    if (fs.statSync(from).isFile()) fs.copyFileSync(from, path.join(target, name))
  }
  return target
}

/** Launch R7 Desktop with a document and the DevTools protocol enabled. */
function launchR7(exePath, installPath, docPath) {
  const child = spawn(exePath, [
    `--remote-debugging-port=${CDP_PORT}`,
    '--force-use-tab',
    docPath
  ], {
    detached: true,
    stdio: 'ignore',
    cwd: installPath
  })
  child.unref()
  return child
}

/**
 * Terminate every R7 Desktop process.
 *
 * `DesktopEditors.exe` is only the launcher: the window and the document
 * handles belong to the child `editors.exe` process, so killing only the
 * launcher leaves the document file locked and any pending save unwritten.
 */
async function killR7() {
  const images = ['editors.exe', 'editors_helper.exe', 'DesktopEditors.exe', 'updatesvc.exe']
  for (const image of images) {
    await new Promise((resolve) => {
      const killer = spawn('taskkill', ['/IM', image, '/F', '/T'], { stdio: 'ignore' })
      killer.on('exit', resolve)
      killer.on('error', resolve)
    })
  }
  // Give the OS a moment to release file handles.
  await delay(2500)
}

/** True while any R7 editor process still holds documents open. */
async function r7Running() {
  return new Promise((resolve) => {
    const p = spawn('tasklist', ['/FI', 'IMAGENAME eq editors.exe', '/NH'], { stdio: ['ignore', 'pipe', 'ignore'] })
    let out = ''
    p.stdout.on('data', (d) => { out += d.toString() })
    p.on('exit', () => resolve(/editors\.exe/i.test(out)))
    p.on('error', () => resolve(false))
  })
}

/** Attach to the live document editor target. */
async function attachEditor() {
  const target = await waitForTargetInternal(
    t => t.type === 'page' && t.url.includes('/apps/api/documents/'),
    60000
  )
  const session = await CdpSession.connect(target.webSocketDebuggerUrl)
  await delay(2500)
  return session
}

async function waitForTargetInternal(predicate, timeoutMs) {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    try {
      const targets = await listTargets(CDP_PORT)
      const match = targets.find(predicate)
      if (match) return match
    } catch { /* endpoint not ready yet */ }
    await delay(600)
  }
  throw new Error('editor target did not appear')
}

/**
 * Read `word/document.xml` from a document that a running editor may hold
 * locked. Returns null instead of throwing so callers can poll.
 */
async function tryReadDocument(filePath) {
  try {
    const zip = await ZipArchive.fromFile(filePath)
    return zip.getText('word/document.xml')
  } catch {
    return null
  }
}

/**
 * Expression that resolves the editor window which owns the pluginsManager,
 * walking the frame tree of the documents API page.
 */
const EDITOR_LOOKUP = `
  (function () {
    var found = null;
    (function walk(win, depth) {
      if (!win || depth > 5 || found) return;
      try {
        if (win.editor && win.editor.pluginsManager && win.g_asc_plugins !== undefined) { found = win; return; }
      } catch (e) { /* cross-origin frame */ }
      for (var i = 0; i < win.frames.length; i++) walk(win.frames[i], depth + 1);
    })(window, 0);
    return found;
  })()
`

async function main() {
  console.log('=== Live R7 Desktop bridge E2E ===\n')

  const adapter = new R7Adapter()
  const info = await adapter.detect()
  if (!info.installed) skip('R7-Office is not installed on this host')
  if (!info.pluginsPath) skip('R7-Office plugin directory not found on this host')

  const docx = new DocxEngine(adapter)
  const workDir = path.join(os.tmpdir(), `dsh_r7_live_${Date.now()}`)
  fs.mkdirSync(workDir, { recursive: true })
  const docPath = path.join(workDir, 'Live_документ.docx')

  // 1. Author the source document with the file engine.
  await docx.create(docPath, {
    title: 'Живой документ для E2E',
    paragraphs: [
      { text: 'Раздел 1. Проверка', style: 'Heading1', bold: true },
      'Исходный текст, который мост должен изменить.',
      { text: 'Раздел 2. Итоги', style: 'Heading1', bold: true },
      'Финальный абзац документа.'
    ]
  })
  check('source document created', fs.existsSync(docPath))

  // 2. Install the bridge plugin into the R7 user plugins directory.
  const pluginDir = installPlugin(info.pluginsPath)
  check('bridge plugin installed into R7', fs.existsSync(path.join(pluginDir, 'config.json')), pluginDir)

  // 3. Relaunch R7 so it scans the plugin directory fresh.
  await killR7()
  launchR7(path.join(info.installPath, 'DesktopEditors.exe'), info.installPath, docPath)

  // 4. Start the DSH-side bridge before the editor tries to connect.
  const bridge = new DesktopBridge({ port: BRIDGE_PORT })
  const port = await bridge.start()
  if (port !== BRIDGE_PORT) {
    console.log(`NOTE: bridge fell back to port ${port} because ${BRIDGE_PORT} was busy`)
  }

  let session = null
  try {
    await waitForCdp(CDP_PORT, 60000)
    session = await attachEditor()
    check('attached to live R7 editor over CDP', true)

    // 5. The native host must have discovered the plugin on disk.
    const discovered = await session.evaluate(`(() => {
      var ed = null;
      (function walk(w, d) {
        if (!w || d > 5 || ed) return;
        try { if (w.AscDesktopEditor) ed = w; } catch (e) {}
        if (ed) return;
        for (var i = 0; i < w.frames.length; i++) walk(w.frames[i], d + 1);
      })(window, 0);
      if (!ed) return null;
      try {
        var roots = JSON.parse(ed.AscDesktopEditor.GetInstallPlugins());
        for (var i = 0; i < roots.length; i++) {
          var list = roots[i].pluginsData || [];
          for (var j = 0; j < list.length; j++) {
            if (list[j].guid === '${PLUGIN_GUID}') return { root: roots[i].url, name: list[j].name };
          }
        }
      } catch (e) { return { error: String(e) }; }
      return null;
    })()`)
    check('R7 native host discovered the bridge plugin',
      Boolean(discovered && discovered.name),
      discovered ? JSON.stringify(discovered) : 'not found')

    // 6. The editor must have registered it with its plugin manager.
    const registered = await session.evaluate(`(() => {
      var ed = ${EDITOR_LOOKUP};
      if (!ed) return null;
      var pm = ed.editor.pluginsManager;
      var list = (pm.plugins || []).map(function (p) { return p.guid; });
      return { count: list.length, hasOurs: list.indexOf('${PLUGIN_GUID}') !== -1 };
    })()`)
    check('editor registered the bridge plugin', Boolean(registered && registered.hasOurs),
      registered ? `${registered.count} plugins registered` : 'no pluginsManager')

    // 7. Start the plugin through the editor's own plugin manager.
    const started = await session.evaluate(`(() => {
      var ed = ${EDITOR_LOOKUP};
      if (!ed) return { ok: false, error: 'editor window not found' };
      try {
        ed.editor.pluginsManager.run('${PLUGIN_GUID}', 0, '');
        return { ok: true };
      } catch (e) { return { ok: false, error: String(e) }; }
    })()`)
    check('editor started the bridge plugin', Boolean(started && started.ok),
      started && started.error ? started.error : '')

    // 8. The plugin must connect back over the real WebSocket.
    const deadline = Date.now() + 25000
    while (Date.now() < deadline && !bridge.getStatus().connected) await delay(400)
    const status = bridge.getStatus()
    check('bridge plugin connected over WebSocket', status.connected,
      `clients=${status.clientCount} port=${status.port}`)

    if (!status.connected) {
      throw new Error('the bridge plugin never connected; cannot continue')
    }

    // 9. Read the live document through the bridge.
    const read = await bridge.execute('safeCommand', { command: 'getDocumentText', args: {} }, 8000)
    const liveText = read?.text || ''
    check('read live document text through the bridge',
      liveText.includes('Исходный текст, который мост должен изменить.'),
      `${liveText.length} chars`)

    // 10. Selection path: round-trip through the editor.
    const selection = await bridge.execute('getSelection', {}, 8000)
    check('selection round-trip through the bridge', selection && selection.success === true,
      JSON.stringify(selection?.text ?? '').slice(0, 60))

    // 11. Modify the open document through the bridge.
    const added = await bridge.execute('safeCommand', {
      command: 'addParagraph',
      args: { text: `${MARKER}: изменение через MCP.` }
    }, 10000)
    check('modified the open document through the bridge', Boolean(added && added.success))

    // 12. Confirm the edit landed in the editor's live model, not just that the
    //     command returned success.
    await delay(800)
    const afterEdit = await bridge.execute('safeCommand', { command: 'getDocumentText', args: {} }, 8000)
    const editedText = afterEdit?.text || ''
    check('the edit is present in the live editor model', editedText.includes(MARKER),
      `${editedText.length} chars`)

    // 13. Save the live document through the bridge.
    const saved = await bridge.execute('safeCommand', { command: 'saveDocument', args: {} }, 15000)
    check('save command accepted by the editor', Boolean(saved && saved.success))

    // 14. Wait for the write to reach disk.
    let onDisk = ''
    const diskDeadline = Date.now() + 20000
    while (Date.now() < diskDeadline) {
      const text = await tryReadDocument(docPath)
      if (text !== null) {
        onDisk = text
        if (onDisk.includes(MARKER)) break
      }
      await delay(600)
    }

    // 15. Close R7, which releases the file handle and flushes any pending save.
    await killR7()
    check('R7 desktop processes stopped', !(await r7Running()))

    // 16. Re-read the artifact once the editor has released it.
    if (!onDisk.includes(MARKER)) {
      const afterClose = await tryReadDocument(docPath)
      if (afterClose !== null) onDisk = afterClose
    }
    check('the saved file on disk contains the bridge edit', onDisk.includes(MARKER),
      onDisk ? `document.xml ${onDisk.length} chars` : 'file unreadable')

    // 17. Re-open the saved document with the file engine.
    const inspection = await docx.inspect(docPath)
    check('reopened and inspected the saved document', inspection.paragraphsCount > 0,
      `${inspection.paragraphsCount} paragraphs`)
    check('the reopened document carries the bridge edit',
      (await docx.read(docPath, { format: 'markdown' })).content.includes(MARKER))

    // 18. Validate the saved document.
    const validation = await docx.validate(docPath)
    check('validated the saved document', validation.valid,
      validation.errors.join('; '))

    // 19. Prove R7 itself still accepts the file by rendering it.
    const pdfPath = path.join(workDir, 'Live_документ.pdf')
    try {
      await adapter.convert(docPath, pdfPath)
      const head = fs.readFileSync(pdfPath).subarray(0, 1024)
      check('R7 rendered the saved document to PDF',
        head.includes(Buffer.from('%PDF-')) && fs.statSync(pdfPath).size > 500,
        `${fs.statSync(pdfPath).size} bytes`)
    } catch (err) {
      check('R7 rendered the saved document to PDF', false,
        (err.stderr || err.message || '').toString().slice(0, 300))
    }

    // 20. A second validation cycle proves the artifact is stable on reopen.
    const secondValidation = await docx.validate(docPath)
    check('the saved document validates again on a second reopen', secondValidation.valid)
  } finally {
    try { session?.close() } catch { /* ignore */ }
    await bridge.stop()
    await killR7()
    if (process.env.DSH_R7_KEEP_ARTIFACTS === '1') {
      console.log(`\nartifacts kept in: ${workDir}`)
    } else {
      try { fs.rmSync(workDir, { recursive: true, force: true }) } catch { /* ignore */ }
    }
  }

  const failed = results.filter(r => !r.ok)
  console.log(`\n=== ${results.length - failed.length}/${results.length} checks passed ===`)
  if (failed.length > 0) {
    for (const f of failed) console.log(`  FAILED: ${f.name} — ${f.detail}`)
    console.log('\nRESULT: FAILED')
    process.exit(1)
  }
  console.log('\nRESULT: PASSED')
}

main().catch((err) => {
  console.error('\nLIVE E2E ERROR:', err.message)
  console.log('\nRESULT: FAILED')
  process.exit(1)
})
