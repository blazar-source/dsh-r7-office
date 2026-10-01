/**
 * The single place that decides whether a test may run on this machine.
 *
 * Some of this project's tests need a real R7-Office installation: they assert
 * values that come from R7's own templates (page size, margins, style ids, the
 * slide-layout graph) or they drive `x2t` and the desktop editor. On a machine
 * without R7 those tests must SKIP with a stated reason rather than fail, which
 * is what the README promises and what CI relies on.
 *
 * The rule for a test author:
 *
 *   - A test that can run against the synthetic package the engine builds when
 *     no R7 template is available MUST run, on every machine, and pass.
 *   - A test that genuinely needs the installed product must call
 *     `requiresR7(t)` first, so its absence is reported, never guessed at.
 *
 * Do not let an R7 template become a silent precondition of an ordinary test:
 * if such a test fails without R7, it is either in the wrong group or it is
 * asserting something that belongs to the R7-specific suite.
 */

import { R7Adapter } from '../../src/r7/adapter.js'

/** Whether a usable R7-Office installation was found on this host. */
export const R7_AVAILABLE = (await new R7Adapter().detect()).installed

/** The reason quoted in every skip, so a CI log explains itself. */
export const R7_SKIP_REASON =
  'needs an installed R7-Office (native templates, x2t or the desktop editor); '
  + 'run `npm run test:r7` on a machine that has it'

/**
 * Skip the current test unless R7 is installed.
 *
 * Usage — the guard has to come first, before any engine call:
 *
 *   test('renders to PDF through x2t', async (t) => {
 *     if (requiresR7(t)) return
 *     ...
 *   })
 *
 * @param {import('node:test').TestContext} t
 * @returns {boolean} true when the caller should return immediately
 */
export function requiresR7(t) {
  if (R7_AVAILABLE) return false
  t.skip(R7_SKIP_REASON)
  return true
}

/**
 * Does the engine have a real R7 template behind it right now?
 *
 * A few tests are meaningful in both worlds but assert different values. Those
 * branch on this instead of being skipped, so the synthetic path keeps its
 * coverage. Prefer this over a filename check.
 */
export const HAS_R7_TEMPLATES = R7_AVAILABLE
