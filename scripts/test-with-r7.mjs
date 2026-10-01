#!/usr/bin/env node
/**
 * The R7-dependent run: native templates, `x2t` conversion and the live editor.
 *
 * `npm test` is the portable suite. It must pass on a machine that has never
 * seen R7-Office, skipping the tests whose subject is the installed product.
 * This command is the other half: it REFUSES to run without an installation,
 * because silently skipping everything would look like success while proving
 * nothing.
 *
 * Usage:
 *   npm run test:r7              # suite + live desktop bridge
 *   npm run test:r7 -- --suite   # suite only (no GUI)
 *   npm run test:r7 -- --live    # live desktop bridge only
 */

import { spawnSync } from 'node:child_process'
import process from 'node:process'
import { R7Adapter } from '../src/r7/adapter.js'

const args = process.argv.slice(2)
const suiteOnly = args.includes('--suite')
const liveOnly = args.includes('--live')

function fail(message) {
  console.error(`\n[r7] ${message}\n`)
  process.exit(1)
}

const info = await new R7Adapter().detect()

if (!info.installed) {
  fail(
    'R7-Office is not installed, so there is nothing to verify.\n'
    + '      This command exists for machines that HAVE it; without one, every\n'
    + '      R7-dependent test would skip and the run would prove nothing.\n'
    + '      Use `npm test` for the portable suite, or install R7-Office.'
  )
}

if (process.env.R7_OFFICE_DISABLED) {
  fail(
    'R7_OFFICE_DISABLED is set, which makes detection report no installation.\n'
    + '      Unset it (Remove-Item Env:\\R7_OFFICE_DISABLED) to run the R7-dependent suite.'
  )
}

console.log(`[r7] R7-Office ${info.version || '(version unknown)'} at ${info.installPath}`)
console.log(`[r7] converter: ${info.x2tPath || 'NOT FOUND'}`)
console.log(`[r7] templates: ${info.templatesPath || 'NOT FOUND'}`)
if (!info.x2tPath) fail('no x2t converter was found; PDF conversion cannot be verified')
if (!info.templatesPath) fail('no native templates were found; document creation cannot be verified')

const steps = []
if (!liveOnly) {
  steps.push({
    name: 'portable suite (R7 present, so nothing skips)',
    command: process.execPath,
    args: ['--test', 'tests/**/*.test.js'],
    shell: false
  })
}
if (!suiteOnly) {
  steps.push({
    name: 'live R7-Office Desktop bridge (CDP + WebSocket + save)',
    command: process.execPath,
    args: ['scripts/live-desktop-e2e.mjs'],
    shell: false
  })
}

let failed = 0
for (const step of steps) {
  console.log(`\n[r7] === ${step.name} ===`)
  const result = spawnSync(step.command, step.args, {
    stdio: 'inherit',
    shell: step.shell,
    cwd: process.cwd()
  })
  if (result.status !== 0) {
    failed++
    console.error(`[r7] FAILED: ${step.name} (exit ${result.status})`)
  } else {
    console.log(`[r7] PASSED: ${step.name}`)
  }
}

if (failed > 0) fail(`${failed} of ${steps.length} R7-dependent step(s) failed`)
console.log('\n[r7] RESULT: PASSED — every R7-dependent check ran and passed')
