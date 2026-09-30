/**
 * Minimal Chrome DevTools Protocol client used by the live R7 Desktop tests.
 *
 * R7-Office Desktop is a CEF application, so launching it with
 * `--remote-debugging-port` exposes the DevTools protocol. That lets a test
 * inspect and drive the real, running editor programmatically instead of
 * emulating mouse and keyboard input.
 *
 * This module has no external dependencies: it speaks the protocol over the
 * WebSocket implementation built into Node.js.
 */

/** Fetch the CDP target list from a running CEF application. */
export async function listTargets(port) {
  const res = await fetch(`http://127.0.0.1:${port}/json/list`)
  if (!res.ok) throw new Error(`CDP target list failed: HTTP ${res.status}`)
  return res.json()
}

/** Fetch the browser-level CDP endpoint description. */
export async function browserVersion(port) {
  const res = await fetch(`http://127.0.0.1:${port}/json/version`)
  if (!res.ok) throw new Error(`CDP version probe failed: HTTP ${res.status}`)
  return res.json()
}

/** Wait until a CDP endpoint answers, or throw after the timeout. */
export async function waitForCdp(port, timeoutMs = 30000) {
  const deadline = Date.now() + timeoutMs
  let lastError = null
  while (Date.now() < deadline) {
    try {
      return await browserVersion(port)
    } catch (err) {
      lastError = err
      await delay(400)
    }
  }
  throw new Error(`CDP endpoint on port ${port} never became available: ${lastError?.message}`)
}

/** Wait for a target whose URL matches a predicate. */
export async function waitForTarget(port, predicate, timeoutMs = 60000) {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    const targets = await listTargets(port)
    const match = targets.find(predicate)
    if (match) return match
    await delay(500)
  }
  throw new Error('no matching CDP target appeared before the timeout')
}

export function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms))
}

/**
 * A single CDP session bound to one target's WebSocket endpoint.
 */
export class CdpSession {
  constructor(ws) {
    this.ws = ws
    this.nextId = 1
    this.pending = new Map()
    this.listeners = new Set()

    ws.addEventListener('message', (event) => {
      let msg
      try {
        msg = JSON.parse(event.data)
      } catch {
        return
      }
      if (msg.id && this.pending.has(msg.id)) {
        const { resolve, reject, timer } = this.pending.get(msg.id)
        clearTimeout(timer)
        this.pending.delete(msg.id)
        if (msg.error) reject(new Error(`${msg.error.message} (${msg.error.code})`))
        else resolve(msg.result)
        return
      }
      for (const listener of this.listeners) listener(msg)
    })
  }

  /** Connect to a target's WebSocket debugger URL. */
  static async connect(webSocketDebuggerUrl, timeoutMs = 15000) {
    return new Promise((resolve, reject) => {
      const ws = new WebSocket(webSocketDebuggerUrl)
      const timer = setTimeout(() => {
        reject(new Error('CDP WebSocket connection timed out'))
      }, timeoutMs)

      ws.addEventListener('open', () => {
        clearTimeout(timer)
        resolve(new CdpSession(ws))
      }, { once: true })

      ws.addEventListener('error', () => {
        clearTimeout(timer)
        reject(new Error('CDP WebSocket connection failed'))
      }, { once: true })
    })
  }

  /** Send a CDP command and await its result. */
  send(method, params = {}, timeoutMs = 20000) {
    const id = this.nextId++
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id)
        reject(new Error(`CDP command ${method} timed out`))
      }, timeoutMs)
      this.pending.set(id, { resolve, reject, timer })
      this.ws.send(JSON.stringify({ id, method, params }))
    })
  }

  /** Subscribe to CDP events. */
  onEvent(listener) {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  /** Evaluate an expression in the target and return its value. */
  async evaluate(expression, { awaitPromise = true, returnByValue = true } = {}) {
    const result = await this.send('Runtime.evaluate', {
      expression,
      awaitPromise,
      returnByValue,
      userGesture: true
    })

    if (result.exceptionDetails) {
      const detail = result.exceptionDetails.exception?.description
        || result.exceptionDetails.text
        || 'unknown evaluation error'
      throw new Error(`Evaluation failed: ${detail}`)
    }
    return result.result?.value
  }

  close() {
    try { this.ws.close() } catch { /* ignore */ }
  }
}
