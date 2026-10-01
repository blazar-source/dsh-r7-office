/**
 * DeepSeek Harness integration smoke test.
 *
 * Boots a fresh DeepSeek Harness process for a named profile on an ephemeral
 * port and asserts that the R7-Office plugin reached the Harness tool
 * registry. A fresh process is deliberate: a long-running Harness caches
 * plugin modules, so only a new boot proves the published code activates.
 *
 * Usage:
 *   node scripts/harness-smoke.mjs [--profile web] [--dsh dsh] [--timeout 90]
 *
 * Exit codes: 0 = verified, 1 = failed, 2 = skipped (plugin not installed in
 * the profile, or the dsh launcher is unavailable).
 */

import { spawn, spawnSync } from 'node:child_process'
import process from 'node:process'
import path from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

function argValue(flag, fallback) {
  const i = process.argv.indexOf(flag)
  return i !== -1 && process.argv[i + 1] ? process.argv[i + 1] : fallback
}

const PROFILE = argValue('--profile', process.env.DSH_PROFILE || 'web')
const DSH = argValue('--dsh', process.platform === 'win32' ? 'dsh.cmd' : 'dsh')
const TIMEOUT_MS = Number(argValue('--timeout', '90')) * 1000

// The expected count comes from the source itself, so this check verifies that
// the tools the plugin actually registers in a live Harness match the tools the
// code declares, instead of going stale whenever one is added.
const here = path.dirname(fileURLToPath(import.meta.url))
let EXPECTED_TOOL_COUNT = 0
let EXPECTED_TOOL_NAMES = []
try {
  const { buildR7Tools } = await import(pathToFileURL(path.join(here, '..', 'src', 'mcp', 'tools.js')).href)
  const built = buildR7Tools({ enableDesktopBridge: false })
  EXPECTED_TOOL_COUNT = built.length
  EXPECTED_TOOL_NAMES = built.map((t) => t.name)
} catch (err) {
  console.error(`could not read the declared tool set: ${err.message}`)
  process.exit(1)
}

const ACTIVATION_PATTERN = /\[r7-office\]\s*зарегистрировано инструментов:\s*(\d+)/u
const FAILURE_PATTERN = /r7-office.*did not activate|tool "r7_\w+" must declare output/u

console.log(`booting a fresh DeepSeek Harness (profile "${PROFILE}") to verify plugin activation...`)

// The launcher is a shell script on Windows (`dsh.cmd`) and a plain binary
// elsewhere. Building one command string keeps Windows working without the
// "args + shell: true" combination Node deprecates.
const isWindows = process.platform === 'win32'
const commandLine = [DSH, '--profile', PROFILE, '--port', '0', '--no-open']
  .map((part) => (part.includes(' ') ? `"${part}"` : part))
  .join(' ')

const child = spawn(commandLine, {
  stdio: ['ignore', 'pipe', 'pipe'],
  shell: isWindows,
  env: { ...process.env, NO_COLOR: '1' }
})

let combined = ''
let settled = false

/**
 * Terminate the whole process tree.
 *
 * The launcher runs under a shell on Windows, so killing the child only kills
 * the shell: the actual `dsh` node process survives, keeps serving on an
 * ephemeral port and — because the plugin starts the desktop bridge at mount
 * time — keeps holding the bridge port. Leaked instances then make the next
 * run connect to a stale bridge.
 */
function killTree() {
  if (isWindows) {
    try {
      spawnSync('taskkill', ['/PID', String(child.pid), '/T', '/F'], { stdio: 'ignore' })
    } catch { /* already gone */ }
  }
  try { child.kill('SIGKILL') } catch { /* already gone */ }
}

function finish(code, message) {
  if (settled) return
  settled = true
  clearTimeout(timer)
  console.log(message)
  killTree()
  console.log(`\nRESULT: ${code === 0 ? 'PASSED' : code === 2 ? 'SKIPPED' : 'FAILED'}`)
  // Give the tree kill a moment to reap the grandchildren before exiting.
  setTimeout(() => process.exit(code), 800)
}

const timer = setTimeout(() => {
  const tail = combined.trim().split('\n').slice(-25).join('\n')
  finish(1, `timed out after ${TIMEOUT_MS / 1000}s waiting for the activation line`
    + (tail ? `\n--- captured output ---\n${tail}` : '\n(no output was captured at all)'))
}, TIMEOUT_MS)

function onChunk(buf, source) {
  const text = buf.toString()
  combined += text
  if (source === 'stderr' && FAILURE_PATTERN.test(text)) {
    finish(1, `activation error detected:\n${text.trim()}`)
    return
  }
  const match = combined.match(ACTIVATION_PATTERN)
  if (match) {
    const count = Number(match[1])
    if (count !== EXPECTED_TOOL_COUNT) {
      finish(1, `plugin registered ${count} tools, but the source declares ${EXPECTED_TOOL_COUNT}`)
      return
    }
    // Check the registered names, not just the count: a count alone would pass
    // even if a tool were silently replaced by another.
    const declared = (combined.match(/\(([^)]*r7_[^)]*)\)/) || [])[1] || ''
    const registered = declared.split(',').map((s) => s.trim()).filter(Boolean)
    const missing = EXPECTED_TOOL_NAMES.filter((n) => !registered.includes(n))
    const extra = registered.filter((n) => !EXPECTED_TOOL_NAMES.includes(n))
    if (missing.length || extra.length) {
      finish(1, `registered tools differ from the declared set`
        + (missing.length ? `\n  missing: ${missing.join(', ')}` : '')
        + (extra.length ? `\n  unexpected: ${extra.join(', ')}` : ''))
      return
    }
    const devModeOff = /developerMode=false/.test(combined)
    finish(
      0,
      `plugin activated in DeepSeek Harness: ${count} tools registered, `
      + `matching the declared set\n`
      + `  tools: ${registered.join(', ')}\n`
      + `  arbitrary DocScript disabled by default: ${devModeOff}`
    )
  }
}

child.stdout.on('data', (b) => onChunk(b, 'stdout'))
child.stderr.on('data', (b) => onChunk(b, 'stderr'))

child.on('exit', (code) => {
  if (settled) return
  if (/r7-office/.test(combined) && /not installed|Cannot find package|ERR_MODULE_NOT_FOUND/.test(combined)) {
    finish(2, `the R7-Office plugin is not installed in profile "${PROFILE}"`)
    return
  }
  finish(1, `Harness exited with code ${code} before the plugin reported activation\n${combined.slice(-2000)}`)
})

child.on('error', (err) => {
  if (err.code === 'ENOENT') {
    finish(2, `the dsh launcher ("${DSH}") is not on PATH`)
    return
  }
  finish(1, `failed to launch dsh: ${err.message}`)
})

