import zlib from 'node:zlib'
import fs from 'node:fs'

/**
 * Pure Node.js ZIP archive reader and writer without external dependencies.
 * Uses built-in zlib for deflate/inflate.
 */
export class ZipArchive {
  constructor() {
    /** @type {Map<string, Buffer>} */
    this.files = new Map()
  }

  /**
   * Load a ZIP archive from a file path.
   * @param {string} filePath
   * @returns {Promise<ZipArchive>}
   */
  static async fromFile(filePath) {
    const buffer = await fs.promises.readFile(filePath)
    return ZipArchive.fromBuffer(buffer)
  }

  /**
   * Synchronously load a ZIP archive from a file path.
   * @param {string} filePath
   * @returns {ZipArchive}
   */
  static fromFileSync(filePath) {
    const buffer = fs.readFileSync(filePath)
    return ZipArchive.fromBuffer(buffer)
  }

  /**
   * Parse a ZIP archive from a Buffer.
   * @param {Buffer} buffer
   * @returns {ZipArchive}
   */
  static fromBuffer(buffer) {
    const zip = new ZipArchive()
    if (buffer.length < 22) {
      throw new Error('Invalid ZIP file: too small')
    }

    // Find End of Central Directory record (EOCD) by scanning backwards from the end
    let eocdOffset = -1
    for (let i = buffer.length - 22; i >= 0; i--) {
      if (buffer.readUInt32LE(i) === 0x06054b50) {
        eocdOffset = i
        break
      }
    }

    if (eocdOffset === -1) {
      throw new Error('Invalid ZIP file: End of Central Directory record not found')
    }

    const totalEntries = buffer.readUInt16LE(eocdOffset + 10)
    const cdOffset = buffer.readUInt32LE(eocdOffset + 16)

    let currentOffset = cdOffset
    for (let entry = 0; entry < totalEntries; entry++) {
      if (buffer.readUInt32LE(currentOffset) !== 0x02014b50) {
        break
      }

      const compressionMethod = buffer.readUInt16LE(currentOffset + 10)
      const compressedSize = buffer.readUInt32LE(currentOffset + 20)
      const uncompressedSize = buffer.readUInt32LE(currentOffset + 24)
      const fileNameLength = buffer.readUInt16LE(currentOffset + 28)
      const extraFieldLength = buffer.readUInt16LE(currentOffset + 30)
      const fileCommentLength = buffer.readUInt16LE(currentOffset + 32)
      const localHeaderOffset = buffer.readUInt32LE(currentOffset + 42)

      const fileName = buffer.toString('utf8', currentOffset + 46, currentOffset + 46 + fileNameLength)
      currentOffset += 46 + fileNameLength + extraFieldLength + fileCommentLength

      // Read local file header to find data start
      if (buffer.readUInt32LE(localHeaderOffset) !== 0x04034b50) {
        continue
      }
      const localFileNameLen = buffer.readUInt16LE(localHeaderOffset + 26)
      const localExtraLen = buffer.readUInt16LE(localHeaderOffset + 28)
      const dataStart = localHeaderOffset + 30 + localFileNameLen + localExtraLen

      const compressedData = buffer.subarray(dataStart, dataStart + compressedSize)
      let fileData

      if (compressionMethod === 0) {
        fileData = Buffer.from(compressedData)
      } else if (compressionMethod === 8) {
        fileData = zlib.inflateRawSync(compressedData)
      } else {
        throw new Error(`Unsupported ZIP compression method: ${compressionMethod} in ${fileName}`)
      }

      zip.files.set(fileName, fileData)
    }

    return zip
  }

  /**
   * Get file content as string.
   * @param {string} name
   * @returns {string|null}
   */
  getText(name) {
    const buf = this.files.get(name)
    return buf ? buf.toString('utf8') : null
  }

  /**
   * Set file content as string.
   * @param {string} name
   * @param {string} text
   */
  setText(name, text) {
    this.files.set(name, Buffer.from(text, 'utf8'))
  }

  /**
   * Get file content as Buffer.
   * @param {string} name
   * @returns {Buffer|null}
   */
  getBuffer(name) {
    return this.files.get(name) || null
  }

  /**
   * Set file content as Buffer.
   * @param {string} name
   * @param {Buffer} buffer
   */
  setBuffer(name, buffer) {
    this.files.set(name, buffer)
  }

  /**
   * Check if file exists in archive.
   * @param {string} name
   * @returns {boolean}
   */
  has(name) {
    return this.files.has(name)
  }

  /**
   * List all file paths in the archive.
   * @returns {string[]}
   */
  list() {
    return Array.from(this.files.keys())
  }

  /**
   * Generate ZIP binary buffer.
   * @returns {Buffer}
   */
  toBuffer() {
    const localHeaders = []
    const cdHeaders = []
    let offset = 0

    const entries = Array.from(this.files.entries())

    for (const [name, uncompressedData] of entries) {
      const fileNameBuf = Buffer.from(name, 'utf8')
      const compressedData = zlib.deflateRawSync(uncompressedData, { level: 6 })
      const crc = crc32(uncompressedData)
      const useDeflate = compressedData.length < uncompressedData.length
      const compMethod = useDeflate ? 8 : 0
      const compSize = useDeflate ? compressedData.length : uncompressedData.length
      const dataToStore = useDeflate ? compressedData : uncompressedData

      // Local File Header (30 bytes + name)
      const localHeader = Buffer.alloc(30 + fileNameBuf.length)
      localHeader.writeUInt32LE(0x04034b50, 0) // Signature
      localHeader.writeUInt16LE(20, 4)         // Version needed
      localHeader.writeUInt16LE(0x0800, 6)     // Flags (UTF-8)
      localHeader.writeUInt16LE(compMethod, 8) // Compression
      localHeader.writeUInt16LE(0, 10)         // Mod time
      localHeader.writeUInt16LE(0, 12)         // Mod date
      localHeader.writeUInt32LE(crc, 14)       // CRC-32
      localHeader.writeUInt32LE(compSize, 18)  // Compressed size
      localHeader.writeUInt32LE(uncompressedData.length, 22) // Uncompressed size
      localHeader.writeUInt16LE(fileNameBuf.length, 26)      // File name length
      localHeader.writeUInt16LE(0, 28)         // Extra field length
      fileNameBuf.copy(localHeader, 30)

      localHeaders.push(localHeader, dataToStore)

      // Central Directory Header (46 bytes + name)
      const cdHeader = Buffer.alloc(46 + fileNameBuf.length)
      cdHeader.writeUInt32LE(0x02014b50, 0) // Signature
      cdHeader.writeUInt16LE(20, 4)         // Version made by
      cdHeader.writeUInt16LE(20, 6)         // Version needed
      cdHeader.writeUInt16LE(0x0800, 8)     // Flags (UTF-8)
      cdHeader.writeUInt16LE(compMethod, 10)// Compression
      cdHeader.writeUInt16LE(0, 12)         // Mod time
      cdHeader.writeUInt16LE(0, 14)         // Mod date
      cdHeader.writeUInt32LE(crc, 16)       // CRC-32
      cdHeader.writeUInt32LE(compSize, 20)  // Compressed size
      cdHeader.writeUInt32LE(uncompressedData.length, 24) // Uncompressed size
      cdHeader.writeUInt16LE(fileNameBuf.length, 28)      // File name length
      cdHeader.writeUInt16LE(0, 30)         // Extra field length
      cdHeader.writeUInt16LE(0, 32)         // File comment length
      cdHeader.writeUInt16LE(0, 34)         // Disk number start
      cdHeader.writeUInt16LE(0, 36)         // Internal file attributes
      cdHeader.writeUInt32LE(0, 38)         // External file attributes
      cdHeader.writeUInt32LE(offset, 42)    // Relative offset of local header
      fileNameBuf.copy(cdHeader, 46)

      cdHeaders.push(cdHeader)

      offset += localHeader.length + compSize
    }

    const cdOffset = offset
    const cdSize = cdHeaders.reduce((sum, h) => sum + h.length, 0)

    // End of Central Directory Record (22 bytes)
    const eocd = Buffer.alloc(22)
    eocd.writeUInt32LE(0x06054b50, 0) // Signature
    eocd.writeUInt16LE(0, 4)          // Disk number
    eocd.writeUInt16LE(0, 6)          // Disk with CD
    eocd.writeUInt16LE(entries.length, 8)  // Entries on this disk
    eocd.writeUInt16LE(entries.length, 10) // Total entries
    eocd.writeUInt32LE(cdSize, 12)    // CD size
    eocd.writeUInt32LE(cdOffset, 16)  // CD offset
    eocd.writeUInt16LE(0, 20)         // Comment length

    return Buffer.concat([...localHeaders, ...cdHeaders, eocd])
  }

  /**
   * Save ZIP archive to disk.
   * @param {string} filePath
   */
  async save(filePath) {
    const buf = this.toBuffer()
    await fs.promises.writeFile(filePath, buf)
  }

  /**
   * Synchronously save ZIP archive to disk.
   * @param {string} filePath
   */
  saveSync(filePath) {
    const buf = this.toBuffer()
    fs.writeFileSync(filePath, buf)
  }
}

/**
 * Standard CRC32 table calculation.
 */
const CRC_TABLE = new Uint32Array(256)
for (let n = 0; n < 256; n++) {
  let c = n
  for (let k = 0; k < 8; k++) {
    c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
  }
  CRC_TABLE[n] = c >>> 0
}

function crc32(buf) {
  let crc = 0xffffffff
  for (let i = 0; i < buf.length; i++) {
    crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ buf[i]) & 0xff]
  }
  return (crc ^ 0xffffffff) >>> 0
}
