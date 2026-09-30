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

import { spawn } from 'node:child_process'
import process from 'node:process'

function argValue(flag, fallback) {
  const i = process.argv.indexOf(flag)
  return i !== -1 && process.argv[i + 1] ? process.argv[i + 1] : fallback
}

const PROFILE = argValue('--profile', process.env.DSH_PROFILE || 'web')
const DSH = argValue('--dsh', process.platform === 'win32' ? 'dsh.cmd' : 'dsh')
const TIMEOUT_MS = Number(argValue('--timeout', '90')) * 1000

const EXPECTED_TOOL_COUNT = 18
const ACTIVATION_PATTERN = /\[r7-office\]\s*Р·Р°СЂРµРіРёСЃС‚СЂРёСЂРѕРІР°РЅРѕ РёРЅСЃС‚СЂСѓРјРµРЅС‚РѕРІ:\s*(\d+)/u
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

function finish(code, message) {
  if (settled) return
  settled = true
  clearTimeout(timer)
  console.log(message)
  try { child.kill('SIGKILL') } catch { /* already gone */ }
  // A detached dsh may leave a node child behind; the port is ephemeral and the
  // process exits on its own once stdin closes, so no process sweep is needed.
  console.log(`\nRESULT: ${code === 0 ? 'PASSED' : code === 2 ? 'SKIPPED' : 'FAILED'}`)
  process.exit(code)
}

const timer = setTimeout(() => {
  finish(1, `timed out after ${TIMEOUT_MS / 1000}s waiting for the activation line`)
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
      finish(1, `plugin registered ${count} tools, expected ${EXPECTED_TOOL_COUNT}`)
      return
    }
    const toolNames = (combined.match(/\(([^)]*r7_inspect[^)]*)\)/) || [])[1] || ''
    const devModeOff = /developerMode=false/.test(combined)
    finish(
      0,
      `plugin activated in DeepSeek Harness: ${count} tools registered\n` +
      `  tools: ${toolNames}\n` +
      `  arbitrary DocScript disabled by default: ${devModeOff}`
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

