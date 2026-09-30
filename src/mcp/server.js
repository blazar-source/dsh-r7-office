import { buildR7Tools } from './tools.js'
import readline from 'node:readline'

/**
 * Standard Model Context Protocol (MCP) Server for R7-Office.
 * Implements JSON-RPC 2.0 protocol over stdio.
 */
export class R7McpServer {
  constructor(options = {}) {
    this.tools = buildR7Tools(options)
    this.serverName = options.name || 'dsh-r7-office'
    this.serverVersion = options.version || '0.1.0'
  }

  /**
   * Start listening for JSON-RPC messages on standard input.
   */
  startStdio() {
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout,
      terminal: false
    })

    rl.on('line', async (line) => {
      const trimmed = line.trim()
      if (!trimmed) return

      try {
        const request = JSON.parse(trimmed)
        const response = await this.handleMessage(request)
        if (response) {
          process.stdout.write(JSON.stringify(response) + '\n')
        }
      } catch (err) {
        process.stdout.write(JSON.stringify({
          jsonrpc: '2.0',
          id: null,
          error: { code: -32700, message: `Parse error: ${err.message}` }
        }) + '\n')
      }
    })
  }

  /**
   * Handle incoming JSON-RPC MCP message.
   * @param {object} request
   * @returns {Promise<object|null>}
   */
  async handleMessage(request) {
    const { id, method, params } = request

    if (method === 'initialize') {
      return {
        jsonrpc: '2.0',
        id,
        result: {
          protocolVersion: '2024-11-05',
          capabilities: {
            tools: {}
          },
          serverInfo: {
            name: this.serverName,
            version: this.serverVersion
          }
        }
      }
    }

    if (method === 'notifications/initialized') {
      return null
    }

    if (method === 'ping') {
      return { jsonrpc: '2.0', id, result: {} }
    }

    if (method === 'tools/list') {
      const list = this.tools.map(t => ({
        name: t.name,
        description: t.description,
        inputSchema: t.parameters
      }))
      return {
        jsonrpc: '2.0',
        id,
        result: { tools: list }
      }
    }

    if (method === 'tools/call') {
      const toolName = params?.name
      const toolArgs = params?.arguments || {}

      const tool = this.tools.find(t => t.name === toolName)
      if (!tool) {
        return {
          jsonrpc: '2.0',
          id,
          error: { code: -32601, message: `Tool not found: ${toolName}` }
        }
      }

      try {
        const result = await tool.execute(toolArgs)
        return {
          jsonrpc: '2.0',
          id,
          result: {
            content: [
              {
                type: 'text',
                text: typeof result === 'string' ? result : JSON.stringify(result, null, 2)
              }
            ]
          }
        }
      } catch (err) {
        return {
          jsonrpc: '2.0',
          id,
          result: {
            isError: true,
            content: [
              {
                type: 'text',
                text: `Error executing ${toolName}: ${err.message}`
              }
            ]
          }
        }
      }
    }

    return {
      jsonrpc: '2.0',
      id,
      error: { code: -32601, message: `Method not found: ${method}` }
    }
  }
}
