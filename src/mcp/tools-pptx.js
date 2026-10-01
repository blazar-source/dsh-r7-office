import fs from 'node:fs'
import { defaultOutput } from './tool-output.js'

/** Whether a file exists, tolerant of a path that cannot be stat'ed. */
async function pathExists(filePath) {
  try {
    return fs.existsSync(filePath)
  } catch {
    return false
  }
}

/**
 * PPTX tool definitions.
 *
 * Five tools cover the whole surface, because a model that has to choose
 * between a dozen near-synonyms will pick the wrong one:
 *
 *  - `r7_slide_read`   — the normalized structure of a slide (geometry, text,
 *                        font, fill, stroke, alignment, paragraphs).
 *  - `r7_slide_create` — create a deck, or append a slide built on one of the
 *                        deck's own layouts.
 *  - `r7_slide_format` — restyle or reposition one existing object, in place.
 *  - `r7_slide_edit`   — deck structure: duplicate, move, reorder, delete.
 *  - `r7_slide_object` — add or remove an object: shape, text box, image.
 *
 * The engine API remains finer-grained; these are the entry points an agent
 * is expected to use.
 *
 * @param {import('../r7/pptx.js').PptxEngine} pptxEngine
 * @returns {Array<object>}
 */
export function buildPptxTools(pptxEngine) {
  /** Font options, shared verbatim by the format and object tools. */
  const fontProperties = {
    family: { type: 'string', description: 'Font name, e.g. "Arial" or "Georgia". Applied to Latin, Cyrillic and complex scripts alike, so Cyrillic does not silently keep the theme font.' },
    size: { type: 'number', description: 'Font size in points, e.g. 24.' },
    bold: { type: 'boolean' },
    italic: { type: 'boolean' },
    underline: { type: 'boolean', description: 'true for a single underline, false to remove it, or a style: double, heavy, dotted, dash, wavy.' },
    color: { type: 'string', description: 'Font colour, e.g. "#C00000".' },
    transparency: { type: 'number', description: 'Font colour transparency 0..1 (0 = opaque).' },
    strike: { type: 'boolean' },
    caps: { type: 'string', enum: ['none', 'small', 'all'], description: 'Capitalisation.' },
    spacing: { type: 'number', description: 'Character spacing in points (negative tightens).' },
    highlight: { type: 'string', description: 'Text highlight colour, e.g. "#FFFF00".' }
  }

  const paragraphProperties = {
    alignment: { type: 'string', enum: ['left', 'center', 'right', 'justify', 'distributed'], description: 'Horizontal paragraph alignment.' },
    verticalAnchor: { type: 'string', enum: ['top', 'center', 'bottom'], description: 'Where the text sits inside its box.' },
    level: { type: 'integer', description: 'Outline level 0..8.' },
    bullet: { type: 'boolean', description: 'true adds a bullet, false removes it. Combine with numbered.' },
    numbered: { type: 'boolean', description: 'true makes the list numbered instead of bulleted.' },
    bulletCharacter: { type: 'string', description: 'Custom bullet glyph, e.g. "–".' },
    lineSpacing: { type: 'number', description: 'Proportional line spacing, 1 = single, 1.5 = one and a half.' },
    spaceBefore: { type: 'number', description: 'Space before the paragraph, in points.' },
    spaceAfter: { type: 'number', description: 'Space after the paragraph, in points.' },
    marginLeft: { type: 'string', description: 'Left indent, e.g. "24pt" or EMU as a number.' },
    indent: { type: 'string', description: 'First-line indent, e.g. "-18pt".' },
    paragraphIndex: { type: 'integer', description: 'Apply the paragraph options to this single paragraph (0-based) instead of all of them.' }
  }

  const geometryProperties = {
    x: { type: 'string', description: 'Left edge. A number is EMU; "2cm", "1in", "30px" and "24pt" also work.' },
    y: { type: 'string', description: 'Top edge, same units as x.' },
    width: { type: 'string', description: 'Width, same units as x.' },
    height: { type: 'string', description: 'Height, same units as x.' },
    rotation: { type: 'number', description: 'Clockwise rotation in degrees.' },
    zOrder: { type: 'integer', description: 'Position in the object list; 0 is furthest back. Omit to place the object on top.' }
  }

  const paintProperties = {
    fill: { type: 'string', description: 'Solid fill colour, e.g. "#1F6FEB". Use "none" for no fill.' },
    fillTransparency: { type: 'number', description: 'Fill transparency 0..1 (0 = opaque).' },
    noFill: { type: 'boolean', description: 'Remove the fill entirely.' },
    line: { type: 'string', description: 'Border colour, e.g. "#0B3D91". Use "none" for no border.' },
    lineWidth: { type: 'number', description: 'Border width in points, e.g. 1.5.' },
    lineTransparency: { type: 'number', description: 'Border transparency 0..1.' },
    lineStyle: { type: 'string', description: 'Border dash style: solid, dash, dot, lgDash, sysDot.' },
    noLine: { type: 'boolean', description: 'Remove the border entirely.' },
    arrows: {
      description: 'Arrow heads for a line: true for one arrow at the end, { head, tail } with triangle/stealth/arrow/diamond/oval.',
      oneOf: [{ type: 'boolean' }, { type: 'object' }]
    }
  }

  const targetProperties = {
    slideIndex: { type: 'integer', description: '0-based position of the slide in the deck.' },
    objectId: { type: 'integer', description: 'Shape id, exactly as r7_slide_read reported it. The most precise way to address an object.' },
    objectName: { type: 'string', description: 'Shape name, as an alternative to objectId.' },
    objectIndex: { type: 'integer', description: '0-based index among the objects that live on the slide.' },
    placeholderType: { type: 'string', description: 'Address a placeholder by role: title, ctrTitle, subTitle, body, pic, tbl.' }
  }

  return [
    {
      name: 'r7_slide_read',
      description: 'Read a PPTX slide as a normalized structure: every object with its id, type, current geometry (x, y, width, height, rotation, z-order), text, font, fill, stroke, alignment and paragraphs. Placeholder geometry and typography are resolved from the slide layout and master, so inherited values are reported rather than left null. Also returns the deck\'s layouts, so the next slide can be built on a real one.',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to the PPTX file.' },
          slideIndex: { type: 'integer', description: '0-based slide position (default 0).' },
          includeInherited: { type: 'boolean', description: 'Also report layout placeholders the slide does not carry itself (default true).' },
          includeRaw: { type: 'boolean', description: 'Attach each object\'s raw OOXML, for debugging only.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        return await pptxEngine.readSlide(args.filePath, args.slideIndex ?? 0, args)
      }
    },

    {
      name: 'r7_slide_create',
      description: 'Create a new PPTX deck, or append a slide to an existing one built on one of the deck\'s own layouts. Appending never modifies the slides already present, and the layout, master and theme the deck already has are reused rather than replaced. Pass baseSlideIndex to clone an existing slide instead of using a layout.',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to the PPTX file. A missing file is created as a new deck; an existing one gets a new slide.' },
          layoutIndex: { type: 'integer', description: '0-based layout to build from, as listed by r7_slide_read. Omit for a sensible default ("heading and content").' },
          layoutName: { type: 'string', description: 'Layout name, e.g. "Титульный слайд". Matched exactly first, then as a substring.' },
          layoutType: { type: 'string', description: 'PowerPoint layout type: title, obj, twoObj, blank, titleOnly, picTx, secHead.' },
          baseSlideIndex: { type: 'integer', description: 'Clone this existing slide instead of building from a layout.' },
          title: { type: 'string', description: 'Text for the new slide\'s title placeholder (or for slide 1 when creating a deck).' },
          subtitle: { type: 'string', description: 'Text for the subtitle placeholder, when the layout has one.' },
          paragraphs: {
            type: 'array',
            description: 'Body text, one entry per paragraph. An entry may be a string, or { text, runs?, bullet?, numbered?, alignment?, size?, color? }.',
            items: {}
          },
          outputPath: { type: 'string', description: 'Write to a copy instead of editing in place.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        const { filePath } = args
        const exists = await pathExists(filePath)
        if (!exists) {
          // Creating the deck is the only moment a template may be used: doing
          // it on an existing file would discard the author's slides, layouts
          // and theme.
          return await pptxEngine.create(filePath, { title: args.title })
        }
        return await pptxEngine.addSlide(filePath, args)
      }
    },

    {
      name: 'r7_slide_format',
      description: 'Format or reposition one object on a PPTX slide, in place. Sets font (family, size, bold, italic, underline, colour), geometry (x, y, width, height, rotation, z-order), fill and border, paragraph alignment, bullet and numbered lists, line and paragraph spacing, and can replace the object\'s text. Only the properties you name change: gradients, shadows, hyperlinks and everything else the author set are left untouched. Falls back to the slide\'s first text object when no target is named.',
      parameters: {
        type: 'object',
        required: ['filePath'],
        properties: {
          filePath: { type: 'string', description: 'Path to the PPTX file.' },
          ...targetProperties,
          text: { type: 'string', description: 'Replace the object\'s entire text with this one paragraph.' },
          paragraphs: {
            type: 'array',
            description: 'Replace the text body with these paragraphs (a string each, or { text, runs?, ...options }).',
            items: {}
          },
          ...geometryProperties,
          ...paintProperties,
          ...fontProperties,
          ...paragraphProperties,
          outputPath: { type: 'string', description: 'Write to a copy instead of editing in place.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        return await pptxEngine.formatObject(args.filePath, args)
      }
    },

    {
      name: 'r7_slide_edit',
      description: 'Change the structure of a PPTX deck: duplicate a slide, move a slide, reorder all slides, or delete a slide. Slides are addressed by their position in the deck, and no slide part is rewritten by a reorder, so nothing on any slide can be affected.',
      parameters: {
        type: 'object',
        required: ['filePath', 'action'],
        properties: {
          filePath: { type: 'string', description: 'Path to the PPTX file.' },
          action: {
            type: 'string',
            enum: ['duplicate', 'move', 'reorder', 'delete'],
            description: 'duplicate one slide, move one, reorder all, or delete one.'
          },
          slideIndex: { type: 'integer', description: 'Target slide for duplicate and delete (0-based).' },
          fromIndex: { type: 'integer', description: 'For move: the slide to move.' },
          toIndex: { type: 'integer', description: 'For move: its new position.' },
          order: {
            type: 'array',
            items: { type: 'integer' },
            description: 'For reorder: the complete new order, a permutation of 0..slideCount-1.'
          },
          position: { type: 'integer', description: 'For duplicate: where the copy goes (default: right after the original).' },
          outputPath: { type: 'string', description: 'Write to a copy instead of editing in place.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        const { action, filePath } = args
        switch (action) {
          case 'delete': {
            if (args.slideIndex === undefined) throw new Error('action "delete" requires slideIndex')
            return await pptxEngine.deleteSlide(filePath, args.slideIndex, args)
          }
          case 'duplicate': {
            if (args.slideIndex === undefined) throw new Error('action "duplicate" requires slideIndex')
            return await pptxEngine.duplicateSlide(filePath, args.slideIndex, args)
          }
          case 'move': {
            if (args.fromIndex === undefined || args.toIndex === undefined) {
              throw new Error('action "move" requires fromIndex and toIndex')
            }
            return await pptxEngine.moveSlide(filePath, args.fromIndex, args.toIndex, args)
          }
          case 'reorder': {
            if (!Array.isArray(args.order)) throw new Error('action "reorder" requires order[]')
            return await pptxEngine.reorderSlides(filePath, args)
          }
          default:
            throw new Error(`Unknown action "${action}". Use duplicate, move, reorder or delete.`)
        }
      }
    },

    {
      name: 'r7_slide_object',
      description: 'Add or remove objects on a PPTX slide: shapes (rectangle, rounded-rectangle, ellipse/circle, line, arrow, triangle, star, chevron, plus more), text boxes and PNG/JPEG images. An object gets its geometry, fill, border, text and font in the same call. Existing media and relationships are never damaged. Set action "replaceImage" with an existing image object\'s id to swap its picture while keeping its geometry and placement.',
      parameters: {
        type: 'object',
        required: ['filePath', 'action'],
        properties: {
          filePath: { type: 'string', description: 'Path to the PPTX file.' },
          action: {
            type: 'string',
            enum: ['addShape', 'addTextBox', 'addImage', 'remove', 'replaceImage'],
            description: 'What to do.'
          },
          slideIndex: { type: 'integer', description: '0-based slide position (default 0).' },
          shape: { type: 'string', description: 'For addShape: rectangle, rounded-rectangle, ellipse, circle, line, arrow, triangle, diamond, pentagon, hexagon, star, chevron, plus, cloud, heart, cylinder, cube, donut, pie, or any DrawingML preset name.' },
          imagePath: { type: 'string', description: 'For addImage/replaceImage: path to a PNG, JPEG, GIF, BMP, TIFF or WebP file. PNG and JPEG sizes are read from the file itself.' },
          objectId: { type: 'integer', description: 'For remove/replaceImage: the shape id reported by r7_slide_read.' },
          objectName: { type: 'string', description: 'Alternative to objectId.' },
          text: { type: 'string', description: 'Text to put in the new object (single paragraph).' },
          paragraphs: {
            type: 'array',
            description: 'Text for the new object, one entry per paragraph; each may be a string or { text, runs?, bullet?, numbered?, alignment? }.',
            items: {}
          },
          name: { type: 'string', description: 'Shape name, so a later call can address it by name.' },
          ...geometryProperties,
          ...paintProperties,
          ...fontProperties,
          ...paragraphProperties,
          lockAspectRatio: { type: 'boolean', description: 'For images: keep the natural aspect ratio when only one of width/height is given (default true).' },
          scale: { type: 'number', description: 'For images with no width or height: render at this multiple of the natural pixel size.' },
          outputPath: { type: 'string', description: 'Write to a copy instead of editing in place.' }
        }
      },
      output: defaultOutput(),
      async execute(args) {
        switch (args.action) {
          case 'addShape':
            return await pptxEngine.addShape(args.filePath, { shape: 'rectangle', ...args })
          case 'addTextBox':
            return await pptxEngine.addTextBox(args.filePath, args)
          case 'addImage':
            if (!args.imagePath) throw new Error('action "addImage" requires imagePath')
            return await pptxEngine.addImage(args.filePath, args)
          case 'remove':
            if (args.objectId === undefined && args.objectName === undefined) {
              throw new Error('action "remove" requires objectId or objectName')
            }
            return await pptxEngine.removeObject(args.filePath, args)
          case 'replaceImage':
            if (!args.imagePath) throw new Error('action "replaceImage" requires imagePath')
            if (args.objectId === undefined && args.objectName === undefined) {
              throw new Error('action "replaceImage" requires objectId or objectName')
            }
            return await pptxEngine.formatObject(args.filePath, args)
          default:
            throw new Error(
              `Unknown action "${args.action}". Use addShape, addTextBox, addImage, remove or replaceImage.`
            )
        }
      }
    }
  ]
}
