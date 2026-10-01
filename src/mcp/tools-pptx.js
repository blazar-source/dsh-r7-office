import fs from 'node:fs'
import { defaultOutput } from './tool-output.js'

/**
 * PPTX tool definitions.
 *
 * @param {import('../r7/pptx.js').PptxEngine} pptxEngine
 * @returns {Array<object>}
 */
export function buildPptxTools(pptxEngine) {
  return [
    {
      name: 'r7_slide_create',
      description: 'Create a new PPTX deck, or append a slide to an existing one. Appending never modifies the slides already present.',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to a PPTX file. A missing file is created; an existing one gets a new slide.' },
          title: { type: 'string', description: 'Title text for the new slide (or for slide 1 when creating a deck).' },
          baseSlideIndex: { type: 'integer', description: 'When appending: 0-based slide to clone the layout from (default 0).' },
          outputPath: { type: 'string', description: 'Optional target path.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        // Appending to an existing deck must never rebuild it from a template:
        // that is data loss, not creation.
        if (fs.existsSync(args.filePath)) {
          return await pptxEngine.addSlide(args.filePath, args)
        }
        return await pptxEngine.create(args.filePath, { title: args.title })
      }
    },

    {
      name: 'r7_slide_edit',
      description: 'Edit the title and shape text of a PPTX slide.',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to PPTX file.' },
          slideIndex: { type: 'integer', description: '0-based slide index.' },
          title: { type: 'string', description: 'New slide title.' },
          search: { type: 'string', description: 'Text to find on the slide.' },
          replace: { type: 'string', description: 'Replacement text.' },
          outputPath: { type: 'string', description: 'Optional target path.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        return await pptxEngine.editSlide(args.filePath, args)
      }
    }
  ]
}
