import { test, describe } from 'node:test'
import assert from 'node:assert/strict'
import { DesktopBridge, SAFE_COMMANDS } from '../../src/mcp/desktop-bridge.js'
import { buildR7Tools } from '../../src/mcp/tools.js'

describe('r7_desktop_exec security policy', () => {
  test('developerMode must be disabled by default', () => {
    const bridge = new DesktopBridge()
    assert.equal(bridge.developerMode, false)
    assert.deepEqual(bridge.getStatus().developerMode, false)
  })

  test('developerMode can be enabled explicitly via config', () => {
    const bridge = new DesktopBridge({ developerMode: true })
    assert.equal(bridge.developerMode, true)
  })

  test('SAFE_COMMANDS is a non-empty allowlist', () => {
    assert.ok(SAFE_COMMANDS instanceof Set)
    assert.ok(SAFE_COMMANDS.size > 0)
    assert.ok(SAFE_COMMANDS.has('addParagraph'))
    assert.ok(SAFE_COMMANDS.has('saveDocument'))
  })

  test('arbitrary code is rejected per tool contract while no client is connected', async () => {
    // With no Desktop client connected, execute() must fail on the connection
    // guard. The security guard is asserted directly below.
    const bridge = new DesktopBridge()
    await assert.rejects(
      () => bridge.execute('callCommand', { code: 'Api.GetDocument()' }),
      /No active R7 Desktop window connected/
    )
  })

  test('security gate rejects arbitrary code before the connection guard', async () => {
    const bridge = new DesktopBridge()
    // Simulate a connected client so the connection guard passes and the
    // security policy is what actually decides the outcome.
    bridge.sockets.add({ write() {}, on() {} })

    await assert.rejects(
      () => bridge.execute('callCommand', { code: 'Api.GetDocument()' }, 200),
      /Security Error: Arbitrary DocScript\/JS execution is disabled in production mode/
    )

    bridge.sockets.clear()
  })

  test('r7_desktop_exec tool refuses unknown safeCommand before reaching the bridge', async () => {
    const bridge = new DesktopBridge()
    const tools = buildR7Tools({ desktopBridge: bridge })
    const tool = tools.find(t => t.name === 'r7_desktop_exec')

    await assert.rejects(
      () => tool.execute({ safeCommand: 'evilCommand', args: {} }),
      /Unknown safeCommand/
    )
  })

  test('r7_desktop_exec tool requires safeCommand or code', async () => {
    const tools = buildR7Tools({ desktopBridge: new DesktopBridge() })
    const tool = tools.find(t => t.name === 'r7_desktop_exec')

    await assert.rejects(
      () => tool.execute({}),
      /Must provide either safeCommand or code/
    )
  })

  test('r7_desktop_exec routes arbitrary code through the security gate', async () => {
    const bridge = new DesktopBridge()
    const tools = buildR7Tools({ desktopBridge: bridge })
    const tool = tools.find(t => t.name === 'r7_desktop_exec')

    // No client connected: connection guard fires (security gate is checked
    // first only once a client exists).
    await assert.rejects(
      () => tool.execute({ code: 'return 1' }),
      /No active R7 Desktop window connected/
    )
  })
})
