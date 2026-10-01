/**
 * The shared tool output contract.
 *
 * Every tool must declare `output { schema, render }`; this factory keeps that
 * declaration identical everywhere so a new tool cannot silently drift from
 * the harness contract.
 */

/**
 * @returns {{schema: object, render: (args: unknown, value: unknown) => Array<object>}}
 */
export function defaultOutput() {
  return {
    schema: {
      type: 'object',
      additionalProperties: true,
      properties: {
        ok: { type: 'boolean' },
        result: { type: 'string' }
      }
    },
    render: (_args, val) => [{
      type: 'text',
      text: typeof val === 'string' ? val : JSON.stringify(val, null, 2)
    }]
  }
}
