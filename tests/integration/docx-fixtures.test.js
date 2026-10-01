import { test, describe } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import zlib from 'node:zlib'
import { probeImage } from '../../src/r7/docx-media.js'
import { ZipArchive } from '../../src/shared/zip.js'

/**
 * Fixtures and helpers shared by the DOCX suites.
 *
 * The project ships no binary documents, so every picture a test needs is
 * encoded here, at run time, with the standard library only. The file is named
 * `docx-*.test.js` because it lives in the integration suite and carries its
 * own tests for the encoders the other suites trust.
 */

/** A temporary directory the caller is expected to remove. */
export function tempDir(label) {
  const dir = path.join(os.tmpdir(), `dsh_r7_${label}_${Date.now()}_${Math.floor(Math.random() * 1e6)}`)
  fs.mkdirSync(dir, { recursive: true })
  return dir
}

/** Remove a directory tree, ignoring failures. */
export function cleanup(dir) {
  try {
    fs.rmSync(dir, { recursive: true, force: true })
  } catch {
    // best effort
  }
}

const CRC_TABLE = (() => {
  const table = new Int32Array(256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    table[n] = c
  }
  return table
})()

/** CRC-32 as PNG requires it. */
export function crc32(buffer) {
  let crc = -1
  for (let i = 0; i < buffer.length; i++) crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ buffer[i]) & 0xff]
  return (crc ^ -1) >>> 0
}

function pngChunk(type, data) {
  const length = Buffer.alloc(4)
  length.writeUInt32BE(data.length)
  const typeBuffer = Buffer.from(type, 'latin1')
  const body = Buffer.concat([typeBuffer, data])
  const crc = Buffer.alloc(4)
  crc.writeUInt32BE(crc32(body))
  return Buffer.concat([length, body, crc])
}

/**
 * Encode a real, valid 8-bit truecolour PNG.
 *
 * @param {number} width
 * @param {number} height
 * @param {[number, number, number]} [rgb]
 * @returns {Buffer}
 */
export function makePng(width, height, rgb = [200, 40, 40]) {
  const raw = Buffer.alloc(height * (1 + width * 3))
  for (let y = 0; y < height; y++) {
    const rowStart = y * (1 + width * 3)
    raw[rowStart] = 0 // filter: none
    for (let x = 0; x < width; x++) {
      const at = rowStart + 1 + x * 3
      raw[at] = (rgb[0] + x * 7) % 256
      raw[at + 1] = (rgb[1] + y * 5) % 256
      raw[at + 2] = rgb[2]
    }
  }

  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(width, 0)
  ihdr.writeUInt32BE(height, 4)
  ihdr[8] = 8 // bit depth
  ihdr[9] = 2 // colour type: truecolour
  ihdr[10] = 0
  ihdr[11] = 0
  ihdr[12] = 0

  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', zlib.deflateSync(raw)),
    pngChunk('IEND', Buffer.alloc(0))
  ])
}

/**
 * A minimal JPEG header: SOI, APP0/JFIF, SOF0 with the requested dimensions,
 * and EOI. Enough for a dimension probe; not a decodable photograph.
 *
 * @param {number} width
 * @param {number} height
 * @returns {Buffer}
 */
export function makeJpegHeader(width, height) {
  const app0 = Buffer.concat([
    Buffer.from([0xff, 0xe0]),
    Buffer.from([0x00, 0x10]),
    Buffer.from('JFIF\0', 'latin1'),
    Buffer.from([0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00])
  ])
  const sof0 = Buffer.concat([
    Buffer.from([0xff, 0xc0, 0x00, 0x11, 0x08]),
    Buffer.from([(height >> 8) & 0xff, height & 0xff]),
    Buffer.from([(width >> 8) & 0xff, width & 0xff]),
    Buffer.from([0x03, 0x01, 0x11, 0x00, 0x02, 0x11, 0x01, 0x03, 0x11, 0x01])
  ])
  return Buffer.concat([Buffer.from([0xff, 0xd8]), app0, sof0, Buffer.from([0xff, 0xd9])])
}

/** True when two buffers are byte-identical. */
export function sameBytes(a, b) {
  return Buffer.isBuffer(a) && Buffer.isBuffer(b) && Buffer.compare(a, b) === 0
}

/**
 * The bytes of one archive member as a comparison should see them.
 *
 * For an archive loaded from disk this is the original compressed stream, which
 * is the strongest available statement ("this member was not re-encoded"). For
 * an archive built in memory there is no replayed stream yet, so the
 * decompressed payload is compared instead.
 */
export function rawMember(zip, name) {
  const entry = zip.entries.get(name)
  if (!entry) return null
  return entry.raw ?? entry.data
}

/**
 * Compare every member of two archives except the ones the caller expects to
 * have changed, returning the list of unexpected differences.
 *
 * @param {ZipArchive} before
 * @param {ZipArchive} after
 * @param {string[]} [allowed]
 * @returns {{changed: string[], added: string[], removed: string[]}}
 */
export function diffMembers(before, after, allowed = []) {
  const allow = new Set(allowed)
  const changed = []
  const added = []
  const removed = []
  const beforeNames = new Set(before.list())
  const afterNames = new Set(after.list())

  for (const name of beforeNames) {
    if (!afterNames.has(name)) {
      removed.push(name)
      continue
    }
    if (allow.has(name)) continue
    if (!sameBytes(rawMember(before, name), rawMember(after, name))) changed.push(name)
  }
  for (const name of afterNames) {
    if (!beforeNames.has(name)) added.push(name)
  }

  return { changed, added, removed }
}

/** Load an archive from a file for a before/after comparison. */
export async function archiveOf(filePath) {
  return ZipArchive.fromFile(filePath)
}

describe('DOCX fixtures', () => {
  test('the generated PNG is a valid image with the requested dimensions', () => {
    const png = makePng(4, 2)
    assert.ok(png.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])))
    assert.equal(png.toString('ascii', 12, 16), 'IHDR')
    // The IHDR chunk CRC written by the encoder must verify.
    const ihdrBody = png.subarray(12, 12 + 4 + 13)
    assert.equal(png.readUInt32BE(12 + 4 + 13), crc32(ihdrBody))

    const probed = probeImage(png)
    assert.equal(probed.format, 'png')
    assert.equal(probed.pixelWidth, 4)
    assert.equal(probed.pixelHeight, 2)
    assert.equal(probed.contentType, 'image/png')
  })

  test('the JPEG probe reads the start-of-frame dimensions', () => {
    const probed = probeImage(makeJpegHeader(320, 200))
    assert.equal(probed.format, 'jpeg')
    assert.equal(probed.pixelWidth, 320)
    assert.equal(probed.pixelHeight, 200)
  })

  test('an unknown payload is reported as unknown rather than guessed', () => {
    const probed = probeImage(Buffer.from('this is not an image at all'))
    assert.equal(probed.format, null)
    assert.equal(probed.pixelWidth, null)
  })

  test('diffMembers reports only the parts that actually changed', () => {
    const before = new ZipArchive()
    before.setText('a.txt', 'a')
    before.setText('b.txt', 'b')

    const after = ZipArchive.fromBuffer(before.toBuffer())
    after.setText('a.txt', 'changed')

    const diff = diffMembers(before, after)
    assert.deepEqual(diff.changed, ['a.txt'])
    assert.deepEqual(diff, { changed: ['a.txt'], added: [], removed: [] })

    const allowed = diffMembers(before, after, ['a.txt'])
    assert.deepEqual(allowed.changed, [])
  })
})
