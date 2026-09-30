/**
 * Desktop Bridge end-to-end test.
 *
 * Boots the real bridge server and drives it through a real WebSocket client,
 * standing in for the R7-Office Desktop editor plugin. This exercises the
 * actual wire protocol (RFC 6455 upgrade + masked client frames) rather than
 * an in-process stub.
 */

import { test, describe, before, after } from 'node:test'
import assert from 'node:assert/strict'
import { DesktopBridge } from '../../src/mcp/desktop-bridge.js'
import { buildR7Tools } from '../../src/mcp/tools.js'

const PORT_BASE = 17888

/** Resolve once the socket is open, or reject on error. */
function open(url) {
  return new Promise((resolve, reject) => {
    const ws = new WebSocket(url)
    ws.addEventListener('open', () => resolve(ws), { once: true })
    ws.addEventListener('error', (e) => reject(new Error(`WebSocket error: ${e.message || 'unknown'}`)), { once: true })
  })
}

/** Wait for the next message and parse it as JSON. */
function nextMessage(ws) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('timed out waiting for a bridge command')), 4000)
    ws.addEventListener('message', (event) => {
      clearTimeout(timer)
      resolve(JSON.parse(event.data))
    }, { once: true })
  })
}

describe('Desktop Bridge E2E', () => {
  describe('production mode (developerMode off)', () => {
    let bridge
    let port
    let ws

    before(async () => {
      bridge = new DesktopBridge({ port: PORT_BASE })
      port = await bridge.start()
      ws = await open(`ws://127.0.0.1:${port}/r7-bridge`)
      ws.send(JSON.stringify({ type: 'register', editorType: 'word' }))
      // Give the server a tick to record the socket.
      await new Promise(r => setTimeout(r, 60))
    })

    after(async () => {
      try { ws?.close() } catch { /* ignore */ }
      await bridge.stop()
    })

    test('reports a connected client', () => {
      const status = bridge.getStatus()
      assert.equal(status.connected, true)
      assert.equal(status.clientCount, 1)
      assert.equal(status.developerMode, false)
    })

    test('performs a selection round-trip over the wire', async () => {
      const pending = bridge.execute('getSelection')
      const command = await nextMessage(ws)
      assert.equal(command.action, 'getSelection')
      assert.ok(command.id, 'the command carries a correlation id')

      ws.send(JSON.stringify({ id: command.id, success: true, text: 'выделенный фрагмент' }))
      const result = await pending
      assert.equal(result.success, true)
      assert.equal(result.text, 'выделенный фрагмент')
    })

    test('performs a selection replacement over the wire', async () => {
      const pending = bridge.execute('replaceSelection', { text: 'новый текст' })
      const command = await nextMessage(ws)
      assert.equal(command.action, 'replaceSelection')
      assert.equal(command.payload.text, 'новый текст')

      ws.send(JSON.stringify({ id: command.id, success: true }))
      const result = await pending
      assert.equal(result.success, true)
    })

    test('refuses arbitrary DocScript while connected', async () => {
      await assert.rejects(
        () => bridge.execute('callCommand', { code: 'Api.GetDocument()' }),
        /Security Error: Arbitrary DocScript\/JS execution is disabled in production mode/
      )
    })

    test('allows a safe command through the wire', async () => {
      const pending = bridge.execute('safeCommand', { command: 'addParagraph', args: { text: 'Абзац' } })
      const command = await nextMessage(ws)
      assert.equal(command.action, 'safeCommand')
      assert.equal(command.payload.command, 'addParagraph')

      ws.send(JSON.stringify({ id: command.id, success: true }))
      const result = await pending
      assert.equal(result.success, true)
    })

    test('reports a timeout instead of hanging when the client never answers', async () => {
      await assert.rejects(
        () => bridge.execute('getSelection', {}, 150),
        /timed out/
      )
    })
  })

  describe('developer mode (explicitly enabled)', () => {
    let bridge
    let port
    let ws

    before(async () => {
      bridge = new DesktopBridge({ port: PORT_BASE + 1, developerMode: true })
      port = await bridge.start()
      ws = await open(`ws://127.0.0.1:${port}/r7-bridge`)
      await new Promise(r => setTimeout(r, 60))
    })

    after(async () => {
      try { ws?.close() } catch { /* ignore */ }
      await bridge.stop()
    })

    test('advertises developer mode in its status', () => {
      assert.equal(bridge.getStatus().developerMode, true)
    })

    test('permits arbitrary DocScript once explicitly enabled', async () => {
      const pending = bridge.execute('callCommand', { code: 'return 1 + 1' })
      const command = await nextMessage(ws)
      assert.equal(command.action, 'callCommand')
      assert.equal(command.payload.code, 'return 1 + 1')

      ws.send(JSON.stringify({ id: command.id, success: true, result: 2 }))
      const result = await pending
      assert.equal(result.result, 2)
    })
  })

  describe('HTTP health endpoint', () => {
    let bridge
    let port

    before(async () => {
      bridge = new DesktopBridge({ port: PORT_BASE + 2 })
      port = await bridge.start()
    })

    after(async () => {
      await bridge.stop()
    })

    test('reports status and the security posture', async () => {
      const res = await fetch(`http://127.0.0.1:${port}/health`)
      assert.equal(res.status, 200)
      const body = await res.json()
      assert.equal(body.status, 'ok')
      assert.equal(body.connectedClients, 0)
      assert.equal(body.developerMode, false)
    })

    test('rejects unknown paths', async () => {
      const res = await fetch(`http://127.0.0.1:${port}/nope`)
      assert.equal(res.status, 404)
    })
  })

  describe('MCP tool surface against a live bridge', () => {
    let bridge
    let port
    let ws
    let tools

    before(async () => {
      bridge = new DesktopBridge({ port: PORT_BASE + 3 })
      port = await bridge.start()
      tools = buildR7Tools({ desktopBridge: bridge })
      ws = await open(`ws://127.0.0.1:${port}/r7-bridge`)
      await new Promise(r => setTimeout(r, 60))
    })

    after(async () => {
      try { ws?.close() } catch { /* ignore */ }
      await bridge.stop()
    })

    test('r7_desktop_status reflects the live connection', async () => {
      const tool = tools.find(t => t.name === 'r7_desktop_status')
      const status = await tool.execute({})
      assert.equal(status.connected, true)
      assert.equal(status.developerMode, false)
    })

    test('r7_desktop_selection reads the live selection', async () => {
      const tool = tools.find(t => t.name === 'r7_desktop_selection')
      const pending = tool.execute({ action: 'get' })
      const command = await nextMessage(ws)
      assert.equal(command.action, 'getSelection')
      ws.send(JSON.stringify({ id: command.id, success: true, text: 'live' }))
      const result = await pending
      assert.equal(result.text, 'live')
    })

    test('r7_desktop_exec rejects arbitrary code in production mode', async () => {
      const tool = tools.find(t => t.name === 'r7_desktop_exec')
      await assert.rejects(
        () => tool.execute({ code: 'Api.GetDocument()' }),
        /Security Error/
      )
    })

    test('r7_desktop_exec runs an allowlisted safe command', async () => {
      const tool = tools.find(t => t.name === 'r7_desktop_exec')
      const pending = tool.execute({ safeCommand: 'saveDocument' })
      const command = await nextMessage(ws)
      assert.equal(command.action, 'safeCommand')
      assert.equal(command.payload.command, 'saveDocument')
      ws.send(JSON.stringify({ id: command.id, success: true }))
      const result = await pending
      assert.equal(result.success, true)
    })
  })
})
