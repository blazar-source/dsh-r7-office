import http from 'node:http'
import fs from 'node:fs'
import path from 'node:path'
import crypto from 'node:crypto'
import { R7Adapter } from '../r7/adapter.js'

/**
 * Lightweight, zero-dependency WebSocket Bridge Server for R7 Desktop Editor connection.
 */
export class DesktopBridge {
  constructor(port = 7888) {
    this.port = port
    this.server = null
    this.sockets = new Set()
    this.pendingRequests = new Map()
    this.r7Adapter = new R7Adapter()
    this.requestId = 1
  }

  /**
   * Start bridge server and listen for connections.
   * @returns {Promise<number>}
   */
  async start() {
    if (this.server) return this.port

    return new Promise((resolve, reject) => {
      this.server = http.createServer((req, res) => {
        if (req.url === '/health') {
          res.writeHead(200, { 'Content-Type': 'application/json' })
          res.end(JSON.stringify({ status: 'ok', connectedClients: this.sockets.size }))
          return
        }
        res.writeHead(404)
        res.end()
      })

      this.server.on('upgrade', (req, socket, head) => {
        if (req.url === '/r7-bridge') {
          this._handleUpgrade(req, socket, head)
        } else {
          socket.destroy()
        }
      })

      this.server.listen(this.port, '127.0.0.1', () => {
        resolve(this.port)
      })

      this.server.on('error', (err) => {
        if (err.code === 'EADDRINUSE') {
          this.port++
          this.server.listen(this.port, '127.0.0.1')
        } else {
          reject(err)
        }
      })
    })
  }

  /**
   * Stop the bridge server.
   */
  async stop() {
    if (!this.server) return
    for (const socket of this.sockets) {
      try { socket.destroy() } catch {}
    }
    this.sockets.clear()
    return new Promise((resolve) => {
      this.server.close(() => {
        this.server = null
        resolve()
      })
    })
  }

  /**
   * Auto-install desktop bridge plugin into R7 Desktop user plugins folder.
   * @returns {Promise<{installed: boolean, targetDir: string|null}>}
   */
  async installPlugin() {
    const info = await this.r7Adapter.detect()
    if (!info.pluginsPath) {
      return { installed: false, targetDir: null, reason: 'R7 plugins directory not found on host' }
    }

    const pluginDirName = '{D5B29457-194D-4E9A-A37F-02D739818FE1}'
    const targetDir = path.join(info.pluginsPath, pluginDirName)

    if (!fs.existsSync(targetDir)) {
      fs.mkdirSync(targetDir, { recursive: true })
    }

    const srcDir = path.resolve('desktop-bridge')
    if (fs.existsSync(srcDir)) {
      const files = fs.readdirSync(srcDir)
      for (const f of files) {
        const srcFile = path.join(srcDir, f)
        const dstFile = path.join(targetDir, f)
        if (fs.statSync(srcFile).isFile()) {
          fs.copyFileSync(srcFile, dstFile)
        }
      }
      return { installed: true, targetDir }
    }

    return { installed: false, targetDir, reason: 'Source desktop-bridge folder not found' }
  }

  /**
   * Check connection status to R7 Desktop.
   * @returns {{connected: boolean, clientCount: number, port: number}}
   */
  getStatus() {
    return {
      connected: this.sockets.size > 0,
      clientCount: this.sockets.size,
      port: this.port
    }
  }

  /**
   * Execute command on active R7 Desktop window.
   * @param {string} action - 'getSelection' | 'replaceSelection' | 'callCommand'
   * @param {object} payload
   * @param {number} [timeoutMs=5000]
   * @returns {Promise<any>}
   */
  async execute(action, payload = {}, timeoutMs = 5000) {
    if (this.sockets.size === 0) {
      throw new Error('No active R7 Desktop window connected to bridge. Please ensure R7 Desktop is open with DSH Bridge plugin active.')
    }

    const id = `req_${this.requestId++}`
    const clientSocket = this.sockets.values().next().value

    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pendingRequests.delete(id)
        reject(new Error(`R7 Desktop bridge request timed out (${timeoutMs}ms)`))
      }, timeoutMs)

      this.pendingRequests.set(id, { resolve, reject, timer })

      const message = JSON.stringify({ id, action, payload })
      this._sendWsFrame(clientSocket, message)
    })
  }

  _handleUpgrade(req, socket, head) {
    const key = req.headers['sec-websocket-key']
    if (!key) {
      socket.destroy()
      return
    }

    const hash = crypto.createHash('sha1')
      .update(key + '258EAFA5-E914-47DA-95CA-C5AB0DC85B11')
      .digest('base64')

    const headers = [
      'HTTP/1.1 101 Switching Protocols',
      'Upgrade: websocket',
      'Connection: Upgrade',
      `Sec-WebSocket-Accept: ${hash}`
    ]

    socket.write(headers.join('\r\n') + '\r\n\r\n')
    this.sockets.add(socket)

    socket.on('data', (buffer) => {
      this._handleWsData(socket, buffer)
    })

    socket.on('close', () => {
      this.sockets.delete(socket)
    })

    socket.on('error', () => {
      this.sockets.delete(socket)
    })
  }

  _handleWsData(socket, buffer) {
    if (buffer.length < 2) return
    const secondByte = buffer[1]
    const isMasked = Boolean(secondByte & 0x80)
    let payloadLen = secondByte & 0x7f
    let currentOffset = 2

    if (payloadLen === 126) {
      payloadLen = buffer.readUInt16BE(2)
      currentOffset += 2
    } else if (payloadLen === 127) {
      payloadLen = Number(buffer.readBigUInt64BE(2))
      currentOffset += 8
    }

    let mask = null
    if (isMasked) {
      mask = buffer.subarray(currentOffset, currentOffset + 4)
      currentOffset += 4
    }

    const payload = buffer.subarray(currentOffset, currentOffset + payloadLen)
    if (isMasked && mask) {
      for (let i = 0; i < payload.length; i++) {
        payload[i] ^= mask[i % 4]
      }
    }

    const text = payload.toString('utf8')
    try {
      const data = JSON.parse(text)
      if (data.id && this.pendingRequests.has(data.id)) {
        const { resolve, timer } = this.pendingRequests.get(data.id)
        clearTimeout(timer)
        this.pendingRequests.delete(data.id)
        resolve(data)
      }
    } catch {
      // ignore
    }
  }

  _sendWsFrame(socket, text) {
    const payload = Buffer.from(text, 'utf8')
    const len = payload.length

    let header
    if (len < 126) {
      header = Buffer.from([0x81, len])
    } else if (len < 65536) {
      header = Buffer.alloc(4)
      header[0] = 0x81
      header[1] = 126
      header.writeUInt16BE(len, 2)
    } else {
      header = Buffer.alloc(10)
      header[0] = 0x81
      header[1] = 127
      header.writeBigUInt64BE(BigInt(len), 2)
    }

    socket.write(Buffer.concat([header, payload]))
  }
}
