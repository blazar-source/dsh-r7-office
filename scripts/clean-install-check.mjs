/**
 * Clean-install verification.
 *
 * Reproduces exactly what a new user gets: clone the committed repository into
 * an empty directory, install, and run the suite there. Nothing from the
 * working tree is reused, so a file that is needed but never committed вЂ” or a
 * test that only passes against local leftovers вЂ” fails this check.
 *
 * Usage:
 *   node scripts/clean-install-check.mjs [--keep] [--skip-tests]
 *
 * Exit codes: 0 = verified, 1 = failed.
 */

import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import { fileURLToPath } from 'node:url'
import { spawn } from 'node:child_process'

const here = path.dirname(fileURLToPath(import.meta.url))
const repoRoot = path.resolve(here, '..')
const keep = process.argv.includes('--keep')
const skipTests = process.argv.includes('--skip-tests')

const steps = []
function record(name, ok, detail = '') {
  steps.push({ name, ok, detail })
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? '  вЂ” ' + detail : ''}`)
}

/** Run a command, streaming output, and resolve with the exit code. */
function run(command, args, cwd, { capture = false } = {}) {
  return new Promise((resolve) => {
    const child = spawn(command, args, {
      cwd,
      shell: process.platform === 'win32',
      stdio: capture ? ['ignore', 'pipe', 'pipe'] : ['ignore', 'inherit', 'inherit'],
      env: { ...process.env, NO_COLOR: '1', npm_config_audit: 'false', npm_config_fund: 'false' }
    })
    let out = ''
    if (capture) {
      child.stdout.on('data', (d) => { out += d.toString() })
      child.stderr.on('data', (d) => { out += d.toString() })
    }
    child.on('error', (err) => resolve({ code: 1, out: err.message }))
    child.on('exit', (code) => resolve({ code: code ?? 1, out }))
  })
}

const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'dsh-r7-clean-'))
const cloneDir = path.join(workDir, 'dsh-r7-office')

console.log(`clean-install check вЂ” work dir: ${workDir}\n`)

try {
  // 1. Confirm the working tree has no uncommitted changes: otherwise this test
  //    would validate something other than what a user receives.
  const status = await run('git', ['status', '--porcelain'], repoRoot, { capture: true })
  const dirty = status.out.trim()
  record('the repository is clean before packaging', dirty === '',
    dirty ? `${dirty.split('\n').length} uncommitted change(s)` : 'nothing uncommitted')

  // 2. Clone the committed HEAD into the empty directory.
  const clone = await run('git', ['clone', '--quiet', '--local', repoRoot, cloneDir], workDir, { capture: true })
  record('git clone into an empty directory', clone.code === 0, clone.out.slice(0, 200))
  if (clone.code !== 0) throw new Error('clone failed')

  // 3. The clone must not carry the working tree's ignored artifacts.
  const cloneHasNodeModules = fs.existsSync(path.join(cloneDir, 'node_modules'))
  record('the clone starts without node_modules', !cloneHasNodeModules)

  // 4. Install.
  const install = await run('npm', ['install'], cloneDir)
  record('npm install in the fresh clone', install.code === 0)
  if (install.code !== 0) throw new Error('install failed')

  // 5. The MCP server must answer a raw JSON-RPC call with no test harness.
  const direct = await new Promise((resolve) => {
    const child = spawn(process.execPath, ['src/mcp/cli.js'], {
      cwd: cloneDir,
      stdio: ['pipe', 'pipe', 'pipe']
    })
    let out = ''
    child.stdout.on('data', (d) => { out += d.toString() })
    child.stdin.write('{"jsonrpc":"2.0","id":1,"method":"tools/list"}\n')
    child.stdin.end()
    child.on('exit', () => resolve(out))
  })
  let toolCount = 0
  try {
    toolCount = JSON.parse(direct.trim()).result.tools.length
  } catch { /* fall through to the assertion below */ }
  record('the MCP server answers tools/list in the clean clone', toolCount === 18,
    `${toolCount} tools`)

  // 6. The full suite, in the clone.
  if (skipTests) {
    console.log('SKIP  npm test (--skip-tests)')
  } else {
    const test = await run('npm', ['test'], cloneDir)
    record('npm test in the fresh clone', test.code === 0)
  }

  // 7. Packaging sanity: the published tarball must contain the runtime and
  //    nothing that leaks local state.
  const pack = await run('npm', ['pack', '--dry-run', '--json'], cloneDir, { capture: true })
  let packed = []
  try {
    packed = JSON.parse(pack.out)[0].files.map((f) => f.path)
  } catch { /* reported by the assertion below */ }
  const hasRuntime = packed.some((p) => p.startsWith('src/'))
  const hasBinary = packed.some((p) => /\.(exe|dll|so|dylib|pak|dat)$/i.test(p))
  const hasTests = packed.some((p) => p.startsWith('tests/'))
  record('npm pack includes the runtime and no R7 binaries', hasRuntime && !hasBinary,
    `${packed.length} files, tests included: ${hasTests}`)
} catch (err) {
  record('clean-install check completed', false, err.message)
} finally {
  if (keep) {
    console.log(`\nartifacts kept in: ${workDir}`)
  } else {
    try { fs.rmSync(workDir, { recursive: true, force: true }) } catch { /* ignore */ }
  }
}

const failed = steps.filter((s) => !s.ok)
console.log(`\n${steps.length - failed.length}/${steps.length} checks passed`)
if (failed.length > 0) {
  for (const f of failed) console.log(`  FAILED: ${f.name} вЂ” ${f.detail}`)
  console.log('\nRESULT: FAILED')
  process.exit(1)
}
console.log('\nRESULT: PASSED')

