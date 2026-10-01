/**
 * Tests against a workbook that R7-Office itself produced.
 *
 * Everything else in the suite authors its own input, which is exactly why a
 * real R7 artifact is worth a separate file: R7's writer orders attributes,
 * numbers its number formats, introduces a shared-string table and drops the
 * parts it does not model, so a reader that only ever saw this engine's output
 * can still be wrong about a genuine R7 file (that is how the `customWidth`
 * misread of `width` was found).
 *
 * The fixture cannot be committed (it is a binary document), so this suite
 * skips when it is absent. Produce it on a host with R7-Office Desktop:
 *
 *     node tests/integration/xlsx-format-r7-authored.test.js --author
 *
 * The authoring run builds a formatted workbook with the engine, opens it in
 * R7-Office Desktop, adds a worksheet through R7's own editor API (which marks
 * the document modified) and saves it, so R7's writer produces the file. It
 * writes the result to
 *
 *     <os.tmpdir()>/r7-acceptance/R7_MCP_XLSX_R7Authored.xlsx
 *
 * or to `R7_XLSX_AUTHORED` when that variable is set.
 */

import { test, describe } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { spawn } from 'node:child_process'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { XlsxEngine } from '../../src/r7/xlsx.js'
import { R7Adapter } from '../../src/r7/adapter.js'
import { ZipArchive } from '../../src/shared/zip.js'
import { serialToDate } from '../../src/r7/xlsx-styles.js'

const here = path.dirname(fileURLToPath(import.meta.url))
const repoRoot = path.resolve(here, '..', '..')

const ACCEPTANCE_DIR = path.join(os.tmpdir(), 'r7-acceptance')
const FIXTURE = process.env.R7_XLSX_AUTHORED
  || path.join(ACCEPTANCE_DIR, 'R7_MCP_XLSX_R7Authored.xlsx')

/** The worksheet the authoring run adds through R7's own editor API. */
const R7_ADDED_SHEET = 'R7_Проверка'
const CDP_PORT = Number(process.env.R7_CDP_PORT || 9222)

// ---------------------------------------------------------------- authoring

/** Press the OK button of any visible native dialog (the R7 trial notice). */
function dismissNativeDialogs(scriptPath) {
  return new Promise((resolve) => {
    const child = spawn(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', scriptPath],
      { stdio: ['ignore', 'pipe', 'ignore'] }
    )
    let out = ''
    child.stdout.on('data', (d) => { out += d.toString() })
    child.on('exit', () => resolve(Number(out.trim()) || 0))
    child.on('error', () => resolve(0))
  })
}

const DISMISS_SCRIPT = `
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class DshR7Dismiss {
  public delegate bool Proc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(Proc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, Proc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  public static string Cls(IntPtr h){ var s=new StringBuilder(200); GetClassName(h,s,200); return s.ToString(); }
}
"@
# R7 shows a modal licensing notice before the editor initialises; while it is
# up the page's JavaScript thread is blocked, so DevTools cannot see anything.
$dialogs = New-Object System.Collections.ArrayList
$cb = [DshR7Dismiss+Proc]{ param($h,$l)
  if ([DshR7Dismiss]::IsWindowVisible($h) -and [DshR7Dismiss]::Cls($h) -eq '#32770') { [void]$dialogs.Add($h) }
  return $true
}
[DshR7Dismiss]::EnumWindows($cb, [IntPtr]::Zero) | Out-Null
$clicked = 0
foreach ($d in $dialogs) {
  $btns = New-Object System.Collections.ArrayList
  $ccb = [DshR7Dismiss+Proc]{ param($h,$l) if ([DshR7Dismiss]::Cls($h) -eq 'Button') { [void]$btns.Add($h) }; return $true }
  [DshR7Dismiss]::EnumChildWindows($d, $ccb, [IntPtr]::Zero) | Out-Null
  foreach ($b in $btns) { [DshR7Dismiss]::SendMessage($b, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null; $clicked = $clicked + 1 }
}
Write-Output $clicked
`

/**
 * Locate the editor window that owns a loaded spreadsheet document.
 *
 * `asc_addWorksheet` only exists on the spreadsheet editor, so requiring it
 * keeps this from silently attaching to a word or presentation window.
 */
const FIND_EDITOR = `(function () {
  var found = null;
  (function walk(win, depth) {
    if (!win || depth > 8 || found) return;
    try {
      if (win.Asc && win.Asc.editor && win.Asc.editor.pluginsManager
          && win.Asc.editor.isDocumentLoadComplete
          && typeof win.Asc.editor.asc_addWorksheet === 'function') { found = win; return; }
    } catch (e) { /* cross-origin frame */ }
    for (var i = 0; i < win.frames.length; i++) walk(win.frames[i], depth + 1);
  })(window, 0);
  return found;
})()`

async function killR7() {
  for (const image of ['editors.exe', 'editors_helper.exe', 'DesktopEditors.exe', 'updatesvc.exe']) {
    await new Promise((resolve) => {
      const killer = spawn('taskkill', ['/IM', image, '/F', '/T'], { stdio: 'ignore' })
      killer.on('exit', resolve)
      killer.on('error', resolve)
    })
  }
  await new Promise((r) => setTimeout(r, 2500))
}

/**
 * Build the fixture: engine-authored formatting, saved by R7's own writer.
 *
 * R7 refuses to reopen a file whose previous session was killed, so every run
 * uses a fresh path and copies the result to the requested destination.
 */
async function authorFixture(destination) {
  if (process.platform !== 'win32') {
    throw new Error('authoring needs R7-Office Desktop on Windows')
  }
  const adapter = new R7Adapter()
  const info = await adapter.detect()
  if (!info.installed || !info.installPath) throw new Error('R7-Office is not installed on this host')

  const cdp = await import(pathToFileURL(path.join(repoRoot, 'scripts', 'cdp.mjs')).href)
  const engine = new XlsxEngine(adapter)

  const workDir = path.join(ACCEPTANCE_DIR, 'authoring')
  fs.mkdirSync(workDir, { recursive: true })
  const stamp = Date.now()
  const source = path.join(workDir, `source-${stamp}.xlsx`)
  const opened = path.join(workDir, `opened-${stamp}.xlsx`)

  console.log('1. building a formatted workbook with the file engine')
  await engine.create(source, {
    overwrite: true,
    sheets: [{ name: 'Доходы', data: [] }, { name: 'Расходы', data: [] }]
  })
  await engine.write(source, {
    sheetName: 'Доходы',
    cells: [
      { ref: 'A1', value: 'Доходы за I квартал' },
      { ref: 'A2', value: 'Дата' }, { ref: 'B2', value: 'Категория' }, { ref: 'C2', value: 'Сумма, ₽' },
      { ref: 'A3', value: '2026-01-15', date: true }, { ref: 'B3', value: 'Лицензии' },
      { ref: 'C3', value: 850000, numberFormat: { type: 'currency', symbol: '₽' } },
      { ref: 'A4', value: '2026-02-15', date: true }, { ref: 'B4', value: 'Внедрение' },
      { ref: 'C4', value: 420000, numberFormat: { type: 'currency', symbol: '₽' } },
      { ref: 'A5', value: 'Итого' },
      { ref: 'C5', formula: 'SUM(C3:C4)', numberFormat: { type: 'currency', symbol: '₽' } }
    ]
  })
  await engine.format(source, {
    sheetName: 'Доходы', range: 'A1:C1', merge: true,
    font: { bold: true, size: 14, color: '#FFFFFF' }, fill: { color: '#2F5597' },
    alignment: { horizontal: 'center' }, rowHeight: 26
  })
  await engine.format(source, {
    sheetName: 'Доходы', range: 'A2:C2', font: { bold: true }, fill: { color: '#D9E2F3' },
    border: { all: 'thin' }, alignment: { horizontal: 'center', wrapText: true }
  })
  await engine.format(source, { sheetName: 'Доходы', range: 'A3:A4', numberFormat: { type: 'date' } })
  await engine.format(source, { sheetName: 'Доходы', range: 'A1:A5', columnWidth: { width: 14 } })
  await engine.format(source, { sheetName: 'Доходы', range: 'B1:B5', columnWidth: { width: 22 } })
  await engine.format(source, { sheetName: 'Доходы', range: 'C1:C5', columnWidth: { width: 16 } })
  await engine.write(source, {
    sheetName: 'Расходы',
    cells: [
      { ref: 'A1', value: 'Расходы' }, { ref: 'A2', value: 'Статья' }, { ref: 'B2', value: 'Сумма, ₽' },
      { ref: 'A3', value: 'ФОТ' }, { ref: 'B3', value: 740000, numberFormat: { type: 'currency', symbol: '₽' } },
      { ref: 'A4', value: 'Аренда' }, { ref: 'B4', value: 180000, numberFormat: { type: 'currency', symbol: '₽' } }
    ]
  })
  await engine.format(source, {
    sheetName: 'Расходы', range: 'A1:B1', merge: true,
    font: { bold: true, size: 14, color: '#FFFFFF' }, fill: { color: '#833C0C' },
    alignment: { horizontal: 'center' }
  })
  await engine.format(source, {
    sheetName: 'Расходы', range: 'A2:B2', font: { bold: true },
    fill: { color: '#FBE5D6' }, border: { all: 'thin' }
  })
  await engine.format(source, { sheetName: 'Расходы', range: 'A1:A4', columnWidth: { width: 20 } })
  await engine.format(source, { sheetName: 'Расходы', range: 'B1:B4', columnWidth: { width: 18 } })

  const before = await ZipArchive.fromFile(source)
  fs.copyFileSync(source, opened)

  console.log('2. opening it in R7-Office Desktop and driving the editor over CDP')
  const dismissScript = path.join(workDir, 'dismiss-trial-dialog.ps1')
  fs.writeFileSync(dismissScript, DISMISS_SCRIPT)

  await killR7()
  const launcher = spawn(path.join(info.installPath, 'DesktopEditors.exe'), [
    `--remote-debugging-port=${CDP_PORT}`, '--force-use-tab', opened
  ], { detached: true, stdio: 'ignore', cwd: info.installPath })
  launcher.unref()

  let target = null
  let session = null
  try {
    await cdp.waitForCdp(CDP_PORT, 90000)
    // R7 is a single-instance application: when another session already has a
    // window open, a fresh launch can hand the file to *that* process, whose
    // document is not ours. The window title in the target URL is what proves
    // which file this connection is looking at.
    const wanted = `title=${path.basename(opened)}`
    for (let i = 0; i < 45 && !target; i++) {
      const clicked = await dismissNativeDialogs(dismissScript)
      if (clicked > 0) console.log(`   dismissed ${clicked} native dialog button(s)`)
      const targets = await cdp.listTargets(CDP_PORT).catch(() => [])
      target = targets.find((t) => t.type === 'page'
        && t.url.includes('/apps/api/documents/')
        && t.url.includes(wanted))
      if (!target) await cdp.delay(2000)
    }
    if (!target) {
      const seen = (await cdp.listTargets(CDP_PORT).catch(() => []))
        .filter((t) => t.url.includes('/apps/api/documents/'))
        .map((t) => (t.url.match(/title=([^&]*)/) || [])[1] || t.url)
      throw new Error(`the R7 window for ${path.basename(opened)} never appeared over CDP `
        + `(CDP saw: ${seen.length > 0 ? seen.join(', ') : 'no document window'})`)
    }

    session = await cdp.CdpSession.connect(target.webSocketDebuggerUrl)
    let ready = false
    for (let i = 0; i < 30 && !ready; i++) {
      await dismissNativeDialogs(dismissScript)
      try {
        const probe = await session.send('Runtime.evaluate', {
          expression: `(() => { const f = ${FIND_EDITOR}; return !!f; })()`,
          returnByValue: true
        }, 6000)
        ready = Boolean(probe.result && probe.result.value)
      } catch { /* the page is still blocked by the modal */ }
      if (!ready) await cdp.delay(2000)
    }
    if (!ready) throw new Error('the R7 spreadsheet editor never finished loading')

    console.log(`3. adding a worksheet ("${R7_ADDED_SHEET}") through R7 and asking R7 to save`)
    const edit = await session.evaluate(`(() => {
      const f = ${FIND_EDITOR};
      const editor = f.Asc.editor;
      const result = { sheetsBefore: editor.asc_getWorksheetsCount() };
      editor.asc_addWorksheet(${JSON.stringify(R7_ADDED_SHEET)});
      result.sheetsAfter = editor.asc_getWorksheetsCount();
      result.save = editor.asc_Save(false) ? 'called' : 'no-op';
      return result;
    })()`)
    console.log('   R7 edit:', JSON.stringify(edit))
    if (edit.sheetsAfter !== edit.sheetsBefore + 1) {
      throw new Error('R7 did not accept the new worksheet, so it would not rewrite the file')
    }
    // asc_Save is asynchronous; give R7 time to flush the package.
    await cdp.delay(8000)
  } finally {
    try { session?.close() } catch { /* ignore */ }
    await killR7()
  }

  const after = await ZipArchive.fromFile(opened)
  let changed = 0
  for (const name of before.list()) {
    const a = before.getBuffer(name)
    const b = after.getBuffer(name)
    if (!b || Buffer.compare(a, b) !== 0) changed++
  }
  if (changed === 0) {
    throw new Error('R7 did not rewrite the workbook (no member changed), so it is not R7-authored')
  }

  const inspection = await engine.inspect(opened)
  const names = inspection.sheets.map((s) => s.name)
  if (!names.includes(R7_ADDED_SHEET)) {
    throw new Error(`the saved workbook does not contain the worksheet R7 added: ${names.join(', ')}`)
  }

  fs.mkdirSync(path.dirname(destination), { recursive: true })
  fs.copyFileSync(opened, destination)
  console.log(`4. R7 rewrote ${changed} package member(s); fixture written to ${destination}`)
  console.log(`   sheets: ${names.join(', ')}`)
  return destination
}

// ------------------------------------------------------------ the test suite

const fixtureAvailable = fs.existsSync(FIXTURE)
const skipReason = fixtureAvailable
  ? false
  : `no R7-authored fixture at ${FIXTURE} — run "node tests/integration/xlsx-format-r7-authored.test.js --author" `
    + 'on a host with R7-Office Desktop, or point R7_XLSX_AUTHORED at one'

/** `--author` turns this file into the tool that produces the fixture. */
const AUTHORING = process.argv.includes('--author')

async function readSheet(filePath, options) {
  const engine = new XlsxEngine()
  return await engine.read(filePath, options)
}

if (AUTHORING) {
  const target = await authorFixture(FIXTURE)
  console.log(`\nfixture: ${target} (${fs.statSync(target).size} bytes)`)
} else {
  describe('a workbook written by R7-Office itself', { skip: skipReason }, () => {
  const engine = new XlsxEngine()
  const tmpDir = path.join(os.tmpdir(), `dsh_r7_authored_${Date.now()}`)
  let copies = 0

  /** A fresh copy per test, so no test can influence another. */
  function copy() {
    fs.mkdirSync(tmpDir, { recursive: true })
    const target = path.join(tmpDir, `copy-${++copies}.xlsx`)
    fs.copyFileSync(FIXTURE, target)
    return target
  }

  test('the package is R7-authored, not this engine\'s output', async () => {
    const zip = await ZipArchive.fromFile(FIXTURE)
    const workbook = zip.getText('xl/workbook.xml')
    const names = [...workbook.matchAll(/<sheet name="([^"]*)"/g)].map((m) => m[1])
    assert.ok(names.includes(R7_ADDED_SHEET),
      `the worksheet R7's editor API added is present: ${names.join(', ')}`)
    assert.ok(names.includes('Доходы') && names.includes('Расходы'))
    assert.ok(zip.has('xl/sharedStrings.xml'), 'R7 wrote a shared-string table')
  })

  test('the engine reads R7\'s values, formulas and cached results', async () => {
    const result = await engine.read(FIXTURE, {
      sheetName: 'Доходы', range: 'A1:C5', includeFormulas: true, includeStyles: true
    })

    assert.equal(result.data[0][0], 'Доходы за I квартал')
    assert.equal(result.data[1][0], 'Дата')
    assert.equal(result.data[1][2], 'Сумма, ₽')
    assert.equal(result.data[2][0], 46037, 'the date is a serial number')
    assert.equal(serialToDate(result.data[2][0]).date, '2026-01-15')
    assert.equal(result.data[2][2], 850000)
    assert.equal(result.formulas[4][2], 'SUM(C3:C4)')
    assert.equal(result.data[4][2], 1270000, 'R7 cached the calculated result in the file')
    assert.deepEqual(result.merged, ['A1:C1'])
    assert.equal(result.rowHeights['1'], 26)
    assert.deepEqual(result.columnWidths, [
      { min: 1, max: 1, width: 14 },
      { min: 2, max: 2, width: 22 },
      { min: 3, max: 3, width: 16 }
    ], 'the widths R7 stored are read back, not the customWidth flag')
  })

  test('the engine reads R7\'s own formatting', async () => {
    const result = await engine.read(FIXTURE, {
      sheetName: 'Доходы', range: 'A1:C5', includeStyles: true
    })

    const title = result.styles[0][0]
    assert.equal(title.font.bold, true)
    assert.equal(title.font.size, 14)
    assert.equal(title.fill.color, 'FF2F5597')
    assert.equal(title.alignment.horizontal, 'center')

    const header = result.styles[1][0]
    assert.equal(header.font.bold, true)
    assert.equal(header.fill.color, 'FFD9E2F3')
    for (const edge of ['left', 'right', 'top', 'bottom']) {
      assert.equal(header.border[edge].style, 'thin', `${edge} border survived R7's rewrite`)
    }
    assert.equal(header.alignment.wrapText, true)

    assert.equal(result.styles[2][0].numberFormatCode, 'DD.MM.YYYY')
    assert.equal(result.styles[2][2].numberFormatCode, '0.00 "₽"')
  })

  test('formatting a copy leaves R7\'s other worksheets and parts intact', async () => {
    const book = copy()
    const before = await ZipArchive.fromFile(book)
    const beforeRead = await engine.read(book, {
      sheetName: 'Доходы', range: 'A1:C5', includeFormulas: true, includeStyles: true
    })

    const result = await engine.format(book, {
      sheetName: 'Доходы',
      range: 'B3:B4',
      fill: { color: '#FFF2CC' },
      alignment: { horizontal: 'left' },
      columnWidth: { width: 26 }
    })
    assert.equal(result.success, true)
    assert.equal(result.cellsFormatted, 2)

    const afterRead = await engine.read(book, {
      sheetName: 'Доходы', range: 'A1:C5', includeFormulas: true, includeStyles: true
    })

    // The values, the formulas and R7's cached result are all still there.
    assert.deepEqual(afterRead.data, beforeRead.data)
    assert.deepEqual(afterRead.formulas, beforeRead.formulas)
    assert.deepEqual(afterRead.merged, beforeRead.merged)
    assert.equal(afterRead.rowHeights['1'], beforeRead.rowHeights['1'])

    // Only the cells that were addressed changed.
    assert.equal(afterRead.styles[2][1].fill.color, 'FFFFF2CC')
    assert.equal(afterRead.styles[3][1].fill.color, 'FFFFF2CC')
    assert.equal(afterRead.styles[2][0].numberFormatCode, 'DD.MM.YYYY', 'the date column is untouched')
    assert.equal(afterRead.styles[2][2].numberFormatCode, '0.00 "₽"', 'the money column is untouched')
    assert.deepEqual(afterRead.styles[0][0], beforeRead.styles[0][0], 'the title keeps R7\'s format')
    assert.deepEqual(afterRead.styles[1][0], beforeRead.styles[1][0], 'the header keeps R7\'s format')

    // Every part but the edited worksheet and the stylesheet is byte-identical.
    const after = await ZipArchive.fromFile(book)
    const edited = new Set(['xl/worksheets/sheet1.xml', 'xl/styles.xml'])
    assert.deepEqual(after.list(), before.list(), 'no member was added or removed')
    for (const name of before.list()) {
      if (edited.has(name)) continue
      assert.equal(Buffer.compare(before.getBuffer(name), after.getBuffer(name)), 0,
        `${name} is byte-identical after the edit`)
    }
    assert.ok(after.has('xl/sharedStrings.xml'), 'R7\'s string table is still there')
  })

  test('a real date written into R7\'s workbook is a serial number', async () => {
    const book = copy()
    await engine.write(book, {
      sheetName: 'Расходы',
      cells: [{ ref: 'A3', value: '2026-03-31', date: true }]
    })

    const result = await engine.read(book, { sheetName: 'Расходы', range: 'A1:B4', includeStyles: true })
    assert.equal(result.data[2][0], 46112, '2026-03-31 as an Excel serial day number')
    assert.equal(serialToDate(result.data[2][0]).date, '2026-03-31')
    assert.equal(result.styles[2][0].numberFormatCode, 'DD.MM.YYYY')
    const expensesPath = (await engine.inspect(book)).sheets.find((s) => s.name === 'Расходы').sheetPath
    assert.ok(!(await ZipArchive.fromFile(book)).getText(expensesPath).includes('2026-03-31'),
      'the date is stored as a serial, never as text')
  })

  test('the formatted copy is still valid and R7 still renders it', async (t) => {
    const adapter = new R7Adapter()
    if (!(await adapter.detect()).installed) { t.skip('R7-Office is not installed'); return }

    const book = copy()
    await engine.format(book, {
      sheetName: 'Доходы',
      range: 'A2:C2',
      font: { bold: true, italic: true, color: '#1F4E79' },
      border: { all: { style: 'thin', color: '#1F4E79' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      rowHeight: 32
    })
    await engine.format(book, { sheetName: 'Доходы', range: 'A1:C1', merge: true, columnWidth: { auto: true } })

    const validation = await engine.validate(book)
    assert.equal(validation.valid, true, JSON.stringify(validation.errors))

    const pdf = path.join(tmpDir, 'authored.pdf')
    await adapter.convert(book, pdf)
    const head = fs.readFileSync(pdf).subarray(0, 1024)
    assert.ok(head.includes(Buffer.from('%PDF-')), 'R7 rendered the edited R7 document')
    assert.ok(fs.statSync(pdf).size > 500)
  })

  test('the whole sheet picture is still readable after the edits', async () => {
    const book = copy()
    await engine.format(book, {
      sheetName: 'Расходы', range: 'A1:B1', merge: true, font: { bold: true }, rowHeight: 24
    })
    const result = await readSheet(book, { sheetName: 'Расходы', range: 'A1:B4', includeStyles: true })
    assert.deepEqual(result.merged, ['A1:B1'])
    assert.equal(result.rowHeights['1'], 24)
    assert.equal(result.data[0][0], 'Расходы')
    assert.equal(result.data[3][1], 180000)
  })
  })
}
