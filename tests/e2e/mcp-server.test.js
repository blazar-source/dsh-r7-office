import { test, describe } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import { R7McpServer } from '../../src/mcp/server.js'

describe('R7McpServer End-to-End', () => {
  const tmpDir = path.join(os.tmpdir(), `dsh_mcp_e2e_${Date.now()}`)
  const testDocx = path.join(tmpDir, 'Тестовый_документ.docx')
  const outPdf = path.join(tmpDir, 'Тестовый_документ.pdf')
  let server

  test('setup test directory and server', () => {
    fs.mkdirSync(tmpDir, { recursive: true })
    server = new R7McpServer()
  })

  test('should handle initialize request', async () => {
    const res = await server.handleMessage({
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: {}
    })
    assert.equal(res.id, 1)
    assert.equal(res.result.serverInfo.name, 'dsh-r7-office')
  })

  test('should list all MCP tools', async () => {
    const res = await server.handleMessage({
      jsonrpc: '2.0',
      id: 2,
      method: 'tools/list'
    })
    assert.equal(res.id, 2)
    const toolNames = res.result.tools.map(t => t.name)
    assert.ok(toolNames.includes('r7_inspect'))
    assert.ok(toolNames.includes('r7_read'))
    assert.ok(toolNames.includes('r7_create'))
    assert.ok(toolNames.includes('r7_edit'))
    assert.ok(toolNames.includes('r7_replace'))
    assert.ok(toolNames.includes('r7_insert'))
    assert.ok(toolNames.includes('r7_table'))
    assert.ok(toolNames.includes('r7_sheet_read'))
    assert.ok(toolNames.includes('r7_sheet_write'))
    assert.ok(toolNames.includes('r7_sheet_add'))
    assert.ok(toolNames.includes('r7_sheet_format'))
    assert.ok(toolNames.includes('r7_sheet_formula'))
    assert.ok(toolNames.includes('r7_docx_formatting'))
    assert.ok(toolNames.includes('r7_docx_sections'))
    assert.ok(toolNames.includes('r7_docx_header_footer'))
    assert.ok(toolNames.includes('r7_docx_image'))
    assert.ok(toolNames.includes('r7_docx_hyperlink'))
    assert.ok(toolNames.includes('r7_slide_read'))
    assert.ok(toolNames.includes('r7_slide_create'))
    assert.ok(toolNames.includes('r7_slide_format'))
    assert.ok(toolNames.includes('r7_slide_edit'))
    assert.ok(toolNames.includes('r7_slide_object'))
    assert.ok(toolNames.includes('r7_convert'))
    assert.ok(toolNames.includes('r7_validate'))
    assert.ok(toolNames.includes('r7_desktop_status'))
    assert.ok(toolNames.includes('r7_desktop_selection'))
    assert.ok(toolNames.includes('r7_desktop_exec'))
    assert.equal(toolNames.length, 27)
  })

  test('should call r7_create via MCP', async () => {
    const res = await server.handleMessage({
      jsonrpc: '2.0',
      id: 3,
      method: 'tools/call',
      params: {
        name: 'r7_create',
        arguments: {
          filePath: testDocx,
          title: 'Документ через MCP',
          paragraphs: [
            'Первый абзац, созданный через вызов MCP-инструмента.',
            'Второй абзац с исходным текстом для замены.'
          ]
        }
      }
    })
    assert.equal(res.id, 3)
    assert.ok(!res.result.isError)
    assert.ok(fs.existsSync(testDocx))
  })

  test('should call r7_inspect via MCP', async () => {
    const res = await server.handleMessage({
      jsonrpc: '2.0',
      id: 4,
      method: 'tools/call',
      params: {
        name: 'r7_inspect',
        arguments: { filePath: testDocx }
      }
    })
    assert.equal(res.id, 4)
    assert.ok(!res.result.isError)
    const data = JSON.parse(res.result.content[0].text)
    assert.equal(data.type, 'docx')
    assert.ok(data.paragraphsCount >= 2)
  })

  test('should call r7_replace via MCP', async () => {
    const res = await server.handleMessage({
      jsonrpc: '2.0',
      id: 5,
      method: 'tools/call',
      params: {
        name: 'r7_replace',
        arguments: {
          filePath: testDocx,
          search: 'Второй абзац с исходным текстом для замены.',
          replace: 'Обновленный и проверенный абзац через MCP.'
        }
      }
    })
    assert.equal(res.id, 5)
    assert.ok(!res.result.isError)

    const readRes = await server.handleMessage({
      jsonrpc: '2.0',
      id: 6,
      method: 'tools/call',
      params: {
        name: 'r7_read',
        arguments: { filePath: testDocx, format: 'markdown' }
      }
    })
    const readData = JSON.parse(readRes.result.content[0].text)
    assert.ok(readData.content.includes('Обновленный и проверенный абзац'))
  })

  test('should call r7_validate via MCP', async () => {
    const res = await server.handleMessage({
      jsonrpc: '2.0',
      id: 7,
      method: 'tools/call',
      params: {
        name: 'r7_validate',
        arguments: { filePath: testDocx }
      }
    })
    assert.equal(res.id, 7)
    const valData = JSON.parse(res.result.content[0].text)
    assert.equal(valData.valid, true)
  })

  test('should convert via r7_convert MCP tool', async () => {
    const res = await server.handleMessage({
      jsonrpc: '2.0',
      id: 8,
      method: 'tools/call',
      params: {
        name: 'r7_convert',
        arguments: {
          sourcePath: testDocx,
          targetPath: outPdf
        }
      }
    })
    assert.equal(res.id, 8)
    assert.ok(fs.existsSync(outPdf))
  })

  test('should handle tool not found gracefully', async () => {
    const res = await server.handleMessage({
      jsonrpc: '2.0',
      id: 9,
      method: 'tools/call',
      params: {
        name: 'unknown_tool',
        arguments: {}
      }
    })
    assert.equal(res.id, 9)
    assert.ok(res.error)
    assert.equal(res.error.code, -32601)
  })

  test('cleanup', () => {
    try {
      fs.rmSync(tmpDir, { recursive: true, force: true })
    } catch {
      // ignore
    }
  })
})

