import { test, describe } from 'node:test'
import assert from 'node:assert/strict'
import { apply, name, inject, R7_GUIDANCE } from '../../src/plugin/index.js'
import { buildR7Tools } from '../../src/mcp/tools.js'

describe('DSH R7 Office Plugin', () => {
  test('should export standard Cordis plugin metadata', () => {
    assert.equal(name, 'r7-office')
    assert.deepEqual(inject, ['tools'])
    assert.ok(typeof apply === 'function')
    assert.ok(R7_GUIDANCE.includes('r7_inspect'))
    assert.ok(R7_GUIDANCE.includes('r7_replace'))
  })

  test('should mount tools and system prompt on Harness context', () => {
    const registeredTools = []
    let promptSection = null

    const mockCtx = {
      tools: {
        register(tool) {
          registeredTools.push(tool)
          return () => {}
        }
      },
      get(service) {
        if (service === 'systemPrompt') {
          return {
            section(sec) {
              promptSection = sec
              return () => {}
            }
          }
        }
        return null
      }
    }

    apply(mockCtx, { enableDesktopBridge: false })

    // Compared against the source, so adding a tool cannot break this test and
    // the assertion still fails if the plugin mounts a different set.
    const declared = buildR7Tools({ enableDesktopBridge: false }).map(t => t.name).sort()
    assert.deepEqual(registeredTools.map(t => t.name).sort(), declared)
    assert.ok(registeredTools.some(t => t.name === 'r7_inspect'))
    assert.ok(registeredTools.some(t => t.name === 'r7_convert'))
    assert.ok(registeredTools.some(t => t.name === 'r7_validate'))

    assert.ok(promptSection)
    assert.equal(promptSection.name, 'plugin:dsh-r7-office')
    assert.equal(promptSection.text, R7_GUIDANCE)
  })
})

