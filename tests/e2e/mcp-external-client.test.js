/**
 * External MCP client smoke test.
 *
 * Speaks to the shipped MCP server over real stdio JSON-RPC using the official
 * Model Context Protocol client SDK (`@modelcontextprotocol/client`), which is
 * an independent implementation maintained by Anthropic. This is the
 * interoperability check: the server must satisfy a client it knows nothing
 * about, not merely its own in-process handler.
 *
 * The suite skips cleanly when the optional dev dependency is not installed,
 * so a clone without `npm install` still runs the rest of the test suite.
 */

import { test, describe, before, after } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import { fileURLToPath } from 'node:url'

const here = path.dirname(fileURLToPath(import.meta.url))
const repoRoot = path.resolve(here, '..', '..')
const serverEntry = path.join(repoRoot, 'src', 'mcp', 'cli.js')

let Client = null
let StdioClientTransport = null
let sdkAvailable = true
let sdkError = null

// Plain assignments rather than destructuring assignment: the latter trips the
// Node test runner's module handling for an optional dependency loaded in a
// try/catch.
try {
  const clientModule = await import('@modelcontextprotocol/client')
  const stdioModule = await import('@modelcontextprotocol/client/stdio')
  Client = clientModule.Client
  StdioClientTransport = stdioModule.StdioClientTransport
  if (typeof Client !== 'function' || typeof StdioClientTransport !== 'function') {
    throw new Error('MCP client SDK did not expose Client/StdioClientTransport')
  }
} catch (err) {
  sdkAvailable = false
  sdkError = err
}

const tmpDir = path.join(os.tmpdir(), `dsh_r7_external_${Date.now()}`)

describe('External MCP client smoke test', { skip: sdkAvailable ? false : 'MCP client SDK not installed (run npm install)' }, () => {
  let client
  let connected = false
  const docPath = path.join(tmpDir, 'external.docx')
  const pdfPath = path.join(tmpDir, 'external.pdf')

  before(async () => {
    fs.mkdirSync(tmpDir, { recursive: true })
    const transport = new StdioClientTransport({
      command: process.execPath,
      args: [serverEntry],
      cwd: repoRoot,
      stderr: 'pipe'
    })
    client = new Client(
      { name: 'dsh-r7-office-smoke', version: '1.0.0' },
      { capabilities: {} }
    )
    await client.connect(transport)
    connected = true
  })

  after(async () => {
    try { if (client) await client.close() } catch { /* already closed */ }
    try { fs.rmSync(tmpDir, { recursive: true, force: true }) } catch { /* ignore */ }
  })

  test('completes the MCP handshake with an independent client', () => {
    assert.equal(connected, true)
    const info = client.getServerVersion()
    assert.ok(info, 'the client received serverInfo')
    assert.equal(info.name, 'dsh-r7-office')
    assert.ok(info.version)
  })

  test('advertises exactly the tool set the source declares', async () => {
    const { tools } = await client.listTools()
    // Derived from the source rather than listed here: a hand-maintained copy
    // went stale on every tool added, and the failure it produced said nothing
    // about the tool set actually being wrong. That the documented set matches
    // is asserted separately, against the READMEs and docs/architecture.md.
    const { buildR7Tools } = await import('../../src/mcp/tools.js')
    const declared = buildR7Tools({}).map(t => t.name).sort()

    assert.equal(tools.length, declared.length)
    assert.deepEqual(tools.map(t => t.name).sort(), declared)
  })

  test('every declared tool appears in the documentation', async () => {
    const { buildR7Tools } = await import('../../src/mcp/tools.js')
    const { readFileSync } = await import('node:fs')
    const { fileURLToPath } = await import('node:url')
    const { dirname, join } = await import('node:path')

    const root = join(dirname(fileURLToPath(import.meta.url)), '..', '..')
    const documented = [
      'README.md',
      'README.ru.md',
      join('docs', 'architecture.md')
    ]

    for (const rel of documented) {
      const text = readFileSync(join(root, rel), 'utf8')
      // A tool is documented by a table row `| \`r7_name\` | ...`
      const listed = new Set([...text.matchAll(/^\| [`](r7_[a-z_]+)[`]/gm)].map(m => m[1]))
      const missing = buildR7Tools({}).map(t => t.name).filter(n => !listed.has(n))
      assert.deepEqual(missing, [], `${rel} does not document: ${missing.join(', ')}`)
    }
  })

  test('every tool publishes a usable input schema', async () => {
    const { tools } = await client.listTools()
    for (const tool of tools) {
      assert.ok(tool.description, `${tool.name} has a description`)
      assert.ok(tool.inputSchema, `${tool.name} has an inputSchema`)
      assert.equal(tool.inputSchema.type, 'object', `${tool.name} schema is an object`)
    }
  })

  test('creates a document through the external client', async () => {
    const res = await client.callTool({
      name: 'r7_create',
      arguments: {
        filePath: docPath,
        title: 'Внешний MCP клиент',
        paragraphs: [
          'Первый абзац, созданный сторонним MCP-клиентом.',
          'Второй абзац для последующей замены.'
        ]
      }
    })
    assert.ok(!res.isError, JSON.stringify(res.content))
    assert.ok(fs.existsSync(docPath))
  })

  test('inspects the created document through the external client', async () => {
    const res = await client.callTool({
      name: 'r7_inspect',
      arguments: { filePath: docPath }
    })
    assert.ok(!res.isError)
    const data = JSON.parse(res.content[0].text)
    assert.equal(data.type, 'docx')
    assert.ok(data.paragraphsCount >= 2)
  })

  test('replaces text through the external client', async () => {
    const res = await client.callTool({
      name: 'r7_replace',
      arguments: {
        filePath: docPath,
        search: 'Второй абзац для последующей замены.',
        replace: 'Абзац, заменённый внешним клиентом.'
      }
    })
    assert.ok(!res.isError)

    const read = await client.callTool({
      name: 'r7_read',
      arguments: { filePath: docPath, format: 'markdown' }
    })
    const data = JSON.parse(read.content[0].text)
    assert.ok(data.content.includes('заменённый внешним клиентом'))
  })

  test('validates through the external client', async () => {
    const res = await client.callTool({
      name: 'r7_validate',
      arguments: { filePath: docPath }
    })
    const data = JSON.parse(res.content[0].text)
    assert.equal(data.valid, true, JSON.stringify(data.errors))
  })

  test('reports a tool error as a result, not a protocol failure', async () => {
    const res = await client.callTool({
      name: 'r7_inspect',
      arguments: { filePath: path.join(tmpDir, 'missing.docx') }
    })
    assert.equal(res.isError, true)
    assert.ok(res.content[0].text.length > 0)
  })

  test('rejects an unknown tool with a protocol error', async () => {
    await assert.rejects(
      () => client.callTool({ name: 'not_a_real_tool', arguments: {} }),
      (err) => {
        assert.match(String(err.message || err), /not_a_real_tool|Method not found|not found/i)
        return true
      }
    )
  })

  test('converts to PDF through the external client', async () => {
    const probe = await client.callTool({
      name: 'r7_convert',
      arguments: { sourcePath: docPath, targetPath: pdfPath }
    })
    // Whether R7 is installed decides the outcome; both shapes must be clean.
    if (probe.isError) {
      assert.match(probe.content[0].text, /x2t|R7/i)
    } else {
      const data = JSON.parse(probe.content[0].text)
      assert.equal(data.success, true)
      assert.ok(fs.existsSync(pdfPath))
    }
  })

  test('the server stays usable for later calls', async () => {
    const res = await client.callTool({
      name: 'r7_desktop_status',
      arguments: {}
    })
    const data = JSON.parse(res.content[0].text)
    assert.equal(typeof data.connected, 'boolean')
    assert.equal(data.developerMode, false, 'production mode keeps arbitrary code disabled')
  })
})

