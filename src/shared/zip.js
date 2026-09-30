import zlib from 'node:zlib'
import fs from 'node:fs'

/** Signature of a Local File Header. */
const SIG_LOCAL = 0x04034b50
/** Signature of a Central Directory File Header. */
const SIG_CENTRAL = 0x02014b50
/** Signature of the End Of Central Directory record. */
const SIG_EOCD = 0x06054b50
/** Signature of the Zip64 End Of Central Directory locator. */
const SIG_ZIP64_LOCATOR = 0x07064b50

/** Maximum length of a ZIP end-of-central-directory comment. */
const MAX_COMMENT = 0xffff

/**
 * A single archive member.
 *
 * `data` is always the decompressed payload. `raw` keeps the bytes exactly as
 * they were read, so an archive that is opened and written back without
 * touching a member reproduces that member byte-for-byte (same compression
 * method, same compressed stream, same CRC and DOS timestamps).
 */
class ZipEntry {
  constructor(name) {
    this.name = name
    this.data = Buffer.alloc(0)
    /** Original compressed bytes, or null for a newly added member. */
    this.raw = null
    /** Original compression method: 0 = stored, 8 = deflate. */
    this.method = 8
    this.crc = 0
    this.dosTime = 0
    this.dosDate = 0
    this.versionMadeBy = 20
    this.externalAttributes = 0
    /** General-purpose bit flag as read. */
    this.flags = 0
    /** Extra fields as read, replayed verbatim for untouched members. */
    this.localExtra = Buffer.alloc(0)
    this.centralExtra = Buffer.alloc(0)
    /** True once the payload is replaced, forcing re-compression on write. */
    this.dirty = false
  }

  /** A directory member carries no payload and a trailing slash. */
  get isDirectory() {
    return this.name.endsWith('/')
  }

  setData(buffer, { method = null } = {}) {
    this.data = buffer
    this.dirty = true
    this.raw = null
    if (method !== null) this.method = method
    else if (this.method !== 0) this.method = 8
    return this
  }
}

/**
 * Pure Node.js ZIP reader/writer with no external dependencies.
 *
 * Uses the built-in zlib for deflate/inflate and deliberately keeps the
 * original compressed stream of every untouched member, which makes
 * read-modify-write cycles on OOXML packages non-destructive for the parts
 * the caller never asked to change.
 */
export class ZipArchive {
  constructor() {
    /** @type {Map<string, ZipEntry>} */
    this.entries = new Map()
  }

  /**
   * Load an archive from a file.
   * @param {string} filePath
   * @returns {Promise<ZipArchive>}
   */
  static async fromFile(filePath) {
    const buffer = await fs.promises.readFile(filePath)
    return ZipArchive.fromBuffer(buffer)
  }

  /**
   * Synchronously load an archive from a file.
   * @param {string} filePath
   * @returns {ZipArchive}
   */
  static fromFileSync(filePath) {
    return ZipArchive.fromBuffer(fs.readFileSync(filePath))
  }

  /**
   * Parse an archive from a Buffer.
   * @param {Buffer} buffer
   * @returns {ZipArchive}
   */
  static fromBuffer(buffer) {
    if (!Buffer.isBuffer(buffer)) {
      throw new Error('Invalid ZIP: expected a Buffer')
    }
    if (buffer.length < 22) {
      throw new Error('Invalid ZIP: truncated archive (shorter than an end-of-central-directory record)')
    }

    const eocdOffset = ZipArchive._findEocd(buffer)
    if (eocdOffset === -1) {
      throw new Error('Invalid ZIP: end-of-central-directory record not found')
    }

    const diskNumber = buffer.readUInt16LE(eocdOffset + 4)
    const cdDisk = buffer.readUInt16LE(eocdOffset + 6)
    if (diskNumber !== 0 || cdDisk !== 0) {
      throw new Error('Invalid ZIP: multi-volume archives are not supported')
    }

    const entryCount = buffer.readUInt16LE(eocdOffset + 10)
    const cdSize = buffer.readUInt32LE(eocdOffset + 12)
    const cdOffset = buffer.readUInt32LE(eocdOffset + 16)

    // Zip64 is signalled by an all-ones field plus a locator record.
    if (entryCount === 0xffff || cdSize === 0xffffffff || cdOffset === 0xffffffff) {
      if (ZipArchive._findZip64Locator(buffer, eocdOffset) !== -1) {
        throw new Error('Invalid ZIP: Zip64 archives are not supported')
      }
    }

    if (cdOffset + cdSize > buffer.length) {
      throw new Error('Invalid ZIP: central directory extends past the end of the file')
    }

    const zip = new ZipArchive()
    let cursor = cdOffset

    for (let i = 0; i < entryCount; i++) {
      if (cursor + 46 > buffer.length) {
        throw new Error('Invalid ZIP: truncated central directory header')
      }
      if (buffer.readUInt32LE(cursor) !== SIG_CENTRAL) {
        throw new Error(`Invalid ZIP: bad central directory signature at entry ${i}`)
      }

      const versionMadeBy = buffer.readUInt16LE(cursor + 4)
      const flags = buffer.readUInt16LE(cursor + 8)
      const method = buffer.readUInt16LE(cursor + 10)
      const dosTime = buffer.readUInt16LE(cursor + 12)
      const dosDate = buffer.readUInt16LE(cursor + 14)
      const crc = buffer.readUInt32LE(cursor + 16)
      const compressedSize = buffer.readUInt32LE(cursor + 20)
      const uncompressedSize = buffer.readUInt32LE(cursor + 24)
      const nameLength = buffer.readUInt16LE(cursor + 28)
      const extraLength = buffer.readUInt16LE(cursor + 30)
      const commentLength = buffer.readUInt16LE(cursor + 32)
      const externalAttributes = buffer.readUInt32LE(cursor + 38)
      const localOffset = buffer.readUInt32LE(cursor + 42)

      const nameStart = cursor + 46
      const nameEnd = nameStart + nameLength
      if (nameEnd > buffer.length) {
        throw new Error('Invalid ZIP: truncated entry name')
      }
      // Bit 11 marks a UTF-8 name; other names are read as latin1-safe UTF-8.
      const name = buffer.toString('utf8', nameStart, nameEnd)
      const centralExtra = Buffer.from(buffer.subarray(nameEnd, nameEnd + extraLength))

      cursor = nameEnd + extraLength + commentLength

      if (flags & 0x1) {
        throw new Error(`Invalid ZIP: encrypted entry is not supported ("${name}")`)
      }
      if (method !== 0 && method !== 8) {
        throw new Error(`Invalid ZIP: unsupported compression method ${method} for "${name}"`)
      }
      if (localOffset + 30 > buffer.length) {
        throw new Error(`Invalid ZIP: entry "${name}" points outside the file`)
      }
      if (buffer.readUInt32LE(localOffset) !== SIG_LOCAL) {
        throw new Error(`Invalid ZIP: bad local header signature for "${name}"`)
      }

      // The local header carries its own (possibly identical) name/extra
      // lengths, which is what locates the payload. Sizes may legitimately be
      // zero here when streaming used a data descriptor, so they are taken
      // from the central directory instead.
      const localMethod = buffer.readUInt16LE(localOffset + 8)
      if (localMethod !== method) {
        throw new Error(
          `Invalid ZIP: compression method mismatch for "${name}" (local ${localMethod}, central ${method})`
        )
      }

      const localNameLength = buffer.readUInt16LE(localOffset + 26)
      const localExtraLength = buffer.readUInt16LE(localOffset + 28)
      const dataStart = localOffset + 30 + localNameLength + localExtraLength
      const dataEnd = dataStart + compressedSize

      if (dataEnd > buffer.length) {
        throw new Error(`Invalid ZIP: truncated payload for "${name}"`)
      }

      const rawCompressed = buffer.subarray(dataStart, dataEnd)
      const entry = new ZipEntry(name)
      entry.method = method
      entry.crc = crc
      entry.dosTime = dosTime
      entry.dosDate = dosDate
      entry.versionMadeBy = versionMadeBy
      entry.externalAttributes = externalAttributes
      entry.flags = flags
      entry.raw = Buffer.from(rawCompressed)
      entry.localExtra = Buffer.from(
        buffer.subarray(localOffset + 30 + localNameLength, dataStart)
      )
      entry.centralExtra = centralExtra

      if (entry.isDirectory || compressedSize === 0) {
        entry.data = Buffer.alloc(0)
      } else if (method === 0) {
        entry.data = Buffer.from(rawCompressed)
      } else {
        try {
          entry.data = zlib.inflateRawSync(rawCompressed)
        } catch (err) {
          throw new Error(`Invalid ZIP: cannot inflate "${name}": ${err.message}`)
        }
      }

      // A mismatch is a corruption signal worth surfacing, but some producers
      // write 0 in the central directory when a data descriptor was used, so
      // only a positive mismatch is treated as an error.
      if (uncompressedSize > 0 && entry.data.length !== uncompressedSize) {
        throw new Error(
          `Invalid ZIP: size mismatch for "${name}" (declared ${uncompressedSize}, inflated ${entry.data.length})`
        )
      }

      zip.entries.set(name, entry)
    }

    return zip
  }

  /** Scan backwards for the end-of-central-directory record. */
  static _findEocd(buffer) {
    const lowest = Math.max(0, buffer.length - MAX_COMMENT - 22)
    for (let i = buffer.length - 22; i >= lowest; i--) {
      if (buffer.readUInt32LE(i) === SIG_EOCD) {
        const commentLength = buffer.readUInt16LE(i + 20)
        // A trailing comment must account for every remaining byte.
        if (i + 22 + commentLength === buffer.length) return i
      }
    }
    return -1
  }

  static _findZip64Locator(buffer, eocdOffset) {
    const locatorOffset = eocdOffset - 20
    if (locatorOffset < 0) return -1
    return buffer.readUInt32LE(locatorOffset) === SIG_ZIP64_LOCATOR ? locatorOffset : -1
  }

  /** Backwards-compatible view of member name to decompressed payload. */
  get files() {
    const view = new Map()
    for (const [name, entry] of this.entries) view.set(name, entry.data)
    return view
  }

  /**
   * Read a member as UTF-8 text.
   * @param {string} name
   * @returns {string|null}
   */
  getText(name) {
    const entry = this.entries.get(name)
    return entry ? entry.data.toString('utf8') : null
  }

  /**
   * Replace a member with UTF-8 text.
   * @param {string} name
   * @param {string} text
   */
  setText(name, text) {
    ZipArchive._upsert(this, name).setData(Buffer.from(text, 'utf8'))
  }

  /**
   * Read a member as a Buffer.
   * @param {string} name
   * @returns {Buffer|null}
   */
  getBuffer(name) {
    const entry = this.entries.get(name)
    return entry ? entry.data : null
  }

  /**
   * Replace a member with raw bytes.
   * @param {string} name
   * @param {Buffer} buffer
   */
  setBuffer(name, buffer) {
    ZipArchive._upsert(this, name).setData(Buffer.from(buffer))
  }

  /**
   * Add a member without compression (stored).
   * @param {string} name
   * @param {Buffer} buffer
   */
  setStored(name, buffer) {
    ZipArchive._upsert(this, name).setData(Buffer.from(buffer), { method: 0 })
  }

  /**
   * Remove a member.
   * @param {string} name
   * @returns {boolean} whether the member existed
   */
  remove(name) {
    return this.entries.delete(name)
  }

  /**
   * Whether the archive contains a member.
   * @param {string} name
   * @returns {boolean}
   */
  has(name) {
    return this.entries.has(name)
  }

  /**
   * All member names, in insertion order.
   * @returns {string[]}
   */
  list() {
    return Array.from(this.entries.keys())
  }

  /**
   * Compression method recorded for a member.
   * @param {string} name
   * @returns {number|null}
   */
  getMethod(name) {
    const entry = this.entries.get(name)
    return entry ? entry.method : null
  }

  /**
   * Serialize the archive back to bytes.
   * @returns {Buffer}
   */
  toBuffer() {
    const parts = []
    const central = []
    let offset = 0

    for (const entry of this.entries.values()) {
      const nameBuffer = Buffer.from(entry.name, 'utf8')
      const crc = entry.dirty || entry.raw === null ? crc32(entry.data) : entry.crc

      let payload
      let method = entry.method

      if (!entry.dirty && entry.raw !== null) {
        // Untouched member: reproduce the original stream verbatim.
        payload = entry.raw
        method = entry.method
      } else if (entry.isDirectory || entry.data.length === 0) {
        payload = Buffer.alloc(0)
        method = 0
      } else if (method === 0) {
        payload = entry.data
      } else {
        const deflated = zlib.deflateRawSync(entry.data, { level: 6 })
        // Storing is cheaper than deflating for already-compressed payloads.
        if (deflated.length >= entry.data.length) {
          payload = entry.data
          method = 0
        } else {
          payload = deflated
        }
      }

      // Untouched members replay their recorded extra fields; a rewritten
      // member drops them, because a stale Zip64 or padding record would no
      // longer describe the new payload.
      const replay = !entry.dirty && entry.raw !== null
      const localExtra = replay ? entry.localExtra : Buffer.alloc(0)
      const centralExtra = replay ? entry.centralExtra : Buffer.alloc(0)

      // Untouched members keep their original flag word (minus the streaming
      // bit, which no longer applies); new members advertise UTF-8 names only
      // when the name actually needs it, so plain ASCII stays maximally
      // compatible with older zip readers.
      let flags
      if (replay) {
        flags = entry.flags & ~0x0008
      } else {
        flags = nameBuffer.length === entry.name.length ? 0x0000 : 0x0800
      }

      const localHeader = Buffer.alloc(30 + nameBuffer.length + localExtra.length)
      localHeader.writeUInt32LE(SIG_LOCAL, 0)
      localHeader.writeUInt16LE(20, 4)
      localHeader.writeUInt16LE(flags, 6)
      localHeader.writeUInt16LE(method, 8)
      localHeader.writeUInt16LE(entry.dosTime, 10)
      localHeader.writeUInt16LE(entry.dosDate, 12)
      localHeader.writeUInt32LE(crc, 14)
      localHeader.writeUInt32LE(payload.length, 18)
      localHeader.writeUInt32LE(entry.data.length, 22)
      localHeader.writeUInt16LE(nameBuffer.length, 26)
      localHeader.writeUInt16LE(localExtra.length, 28)
      nameBuffer.copy(localHeader, 30)
      if (localExtra.length > 0) localExtra.copy(localHeader, 30 + nameBuffer.length)

      parts.push(localHeader, payload)

      const centralHeader = Buffer.alloc(46 + nameBuffer.length + centralExtra.length)
      centralHeader.writeUInt32LE(SIG_CENTRAL, 0)
      centralHeader.writeUInt16LE(entry.versionMadeBy, 4)
      centralHeader.writeUInt16LE(20, 6)
      centralHeader.writeUInt16LE(flags, 8)
      centralHeader.writeUInt16LE(method, 10)
      centralHeader.writeUInt16LE(entry.dosTime, 12)
      centralHeader.writeUInt16LE(entry.dosDate, 14)
      centralHeader.writeUInt32LE(crc, 16)
      centralHeader.writeUInt32LE(payload.length, 20)
      centralHeader.writeUInt32LE(entry.data.length, 24)
      centralHeader.writeUInt16LE(nameBuffer.length, 28)
      centralHeader.writeUInt16LE(centralExtra.length, 30)
      centralHeader.writeUInt16LE(0, 32)
      centralHeader.writeUInt16LE(0, 34)
      centralHeader.writeUInt16LE(0, 36)
      centralHeader.writeUInt32LE(entry.externalAttributes, 38)
      centralHeader.writeUInt32LE(offset, 42)
      nameBuffer.copy(centralHeader, 46)
      if (centralExtra.length > 0) centralExtra.copy(centralHeader, 46 + nameBuffer.length)

      central.push(centralHeader)
      offset += localHeader.length + payload.length
    }

    const centralSize = central.reduce((sum, buf) => sum + buf.length, 0)

    if (this.entries.size > 0xffff) {
      throw new Error('Cannot write ZIP: too many entries for a non-Zip64 archive')
    }
    if (offset > 0xffffffff || centralSize > 0xffffffff) {
      throw new Error('Cannot write ZIP: archive exceeds the 4 GiB Zip64 limit')
    }

    const eocd = Buffer.alloc(22)
    eocd.writeUInt32LE(SIG_EOCD, 0)
    eocd.writeUInt16LE(0, 4)
    eocd.writeUInt16LE(0, 6)
    eocd.writeUInt16LE(this.entries.size, 8)
    eocd.writeUInt16LE(this.entries.size, 10)
    eocd.writeUInt32LE(centralSize, 12)
    eocd.writeUInt32LE(offset, 16)
    eocd.writeUInt16LE(0, 20)

    return Buffer.concat([...parts, ...central, eocd])
  }

  /**
   * Write the archive to disk.
   * @param {string} filePath
   */
  async save(filePath) {
    await fs.promises.writeFile(filePath, this.toBuffer())
  }

  /**
   * Synchronously write the archive to disk.
   * @param {string} filePath
   */
  saveSync(filePath) {
    fs.writeFileSync(filePath, this.toBuffer())
  }

  static _upsert(zip, name) {
    let entry = zip.entries.get(name)
    if (!entry) {
      entry = new ZipEntry(name)
      zip.entries.set(name, entry)
    }
    return entry
  }
}

/** Precomputed CRC-32 lookup table (polynomial 0xEDB88320, reflected). */
const CRC_TABLE = new Uint32Array(256)
for (let n = 0; n < 256; n++) {
  let c = n
  for (let k = 0; k < 8; k++) {
    c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
  }
  CRC_TABLE[n] = c >>> 0
}

/**
 * Standard CRC-32 used by ZIP.
 * @param {Buffer} buf
 * @returns {number}
 */
export function crc32(buf) {
  let crc = 0xffffffff
  for (let i = 0; i < buf.length; i++) {
    crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ buf[i]) & 0xff]
  }
  return (crc ^ 0xffffffff) >>> 0
}
