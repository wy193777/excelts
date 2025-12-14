/**
 * Pure Uint8Array-based ZIP parser
 * Works in both Node.js and browser environments
 * No dependency on Node.js stream module
 */

import { decompress, decompressSync } from "../zip/compress.js";

// ZIP file signatures
const LOCAL_FILE_HEADER_SIG = 0x04034b50;
const CENTRAL_DIR_HEADER_SIG = 0x02014b50;
const END_OF_CENTRAL_DIR_SIG = 0x06054b50;
const ZIP64_END_OF_CENTRAL_DIR_SIG = 0x06064b50;
const ZIP64_END_OF_CENTRAL_DIR_LOCATOR_SIG = 0x07064b50;

// Compression methods
const COMPRESSION_STORED = 0;
const COMPRESSION_DEFLATE = 8;

/**
 * Parse DOS date/time format to JavaScript Date
 * Dates in zip file entries are stored as DosDateTime
 * Spec: https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-dosdatetimetofiletime
 */
function parseDateTime(date: number, time?: number): Date {
  const day = date & 0x1f;
  const month = (date >> 5) & 0x0f;
  const year = ((date >> 9) & 0x7f) + 1980;
  const seconds = time ? (time & 0x1f) * 2 : 0;
  const minutes = time ? (time >> 5) & 0x3f : 0;
  const hours = time ? time >> 11 : 0;

  return new Date(Date.UTC(year, month - 1, day, hours, minutes, seconds));
}

/**
 * Parse ZIP64 extra field
 */
function parseZip64ExtraField(
  extraField: Uint8Array,
  compressedSize: number,
  uncompressedSize: number,
  localHeaderOffset: number
): { compressedSize: number; uncompressedSize: number; localHeaderOffset: number } {
  const view = new DataView(extraField.buffer, extraField.byteOffset, extraField.byteLength);
  let offset = 0;

  while (offset + 4 <= extraField.length) {
    const signature = view.getUint16(offset, true);
    const partSize = view.getUint16(offset + 2, true);

    if (signature === 0x0001) {
      // ZIP64 extended information
      let fieldOffset = offset + 4;

      if (uncompressedSize === 0xffffffff && fieldOffset + 8 <= offset + 4 + partSize) {
        uncompressedSize = Number(view.getBigUint64(fieldOffset, true));
        fieldOffset += 8;
      }
      if (compressedSize === 0xffffffff && fieldOffset + 8 <= offset + 4 + partSize) {
        compressedSize = Number(view.getBigUint64(fieldOffset, true));
        fieldOffset += 8;
      }
      if (localHeaderOffset === 0xffffffff && fieldOffset + 8 <= offset + 4 + partSize) {
        localHeaderOffset = Number(view.getBigUint64(fieldOffset, true));
      }
      break;
    }

    offset += 4 + partSize;
  }

  return { compressedSize, uncompressedSize, localHeaderOffset };
}

/**
 * Parsed ZIP entry
 */
export interface ZipEntryInfo {
  /** File path within the ZIP */
  path: string;
  /** Whether this is a directory */
  isDirectory: boolean;
  /** Compressed size */
  compressedSize: number;
  /** Uncompressed size */
  uncompressedSize: number;
  /** Compression method (0 = stored, 8 = deflate) */
  compressionMethod: number;
  /** CRC-32 checksum */
  crc32: number;
  /** Last modified date */
  lastModified: Date;
  /** Offset to local file header */
  localHeaderOffset: number;
  /** File comment */
  comment: string;
  /** External file attributes */
  externalAttributes: number;
  /** Is encrypted */
  isEncrypted: boolean;
}

/**
 * ZIP parsing options
 */
export interface ZipParseOptions {
  /** Whether to decode file names as UTF-8 (default: true) */
  decodeStrings?: boolean;
}

/**
 * DataView helper for reading little-endian values
 */
class BinaryReader {
  private view: DataView;
  private offset: number;
  private data: Uint8Array;

  constructor(data: Uint8Array, offset = 0) {
    this.data = data;
    this.view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    this.offset = offset;
  }

  get position(): number {
    return this.offset;
  }

  set position(value: number) {
    this.offset = value;
  }

  get remaining(): number {
    return this.data.length - this.offset;
  }

  readUint8(): number {
    const value = this.view.getUint8(this.offset);
    this.offset += 1;
    return value;
  }

  readUint16(): number {
    const value = this.view.getUint16(this.offset, true);
    this.offset += 2;
    return value;
  }

  readUint32(): number {
    const value = this.view.getUint32(this.offset, true);
    this.offset += 4;
    return value;
  }

  readBigUint64(): bigint {
    const value = this.view.getBigUint64(this.offset, true);
    this.offset += 8;
    return value;
  }

  readBytes(length: number): Uint8Array {
    const bytes = this.data.subarray(this.offset, this.offset + length);
    this.offset += length;
    return bytes;
  }

  readString(length: number, utf8 = true): string {
    const bytes = this.readBytes(length);
    if (utf8) {
      return new TextDecoder("utf-8").decode(bytes);
    }
    // Fallback to ASCII/Latin-1
    return String.fromCharCode(...bytes);
  }

  skip(length: number): void {
    this.offset += length;
  }

  slice(start: number, end: number): Uint8Array {
    return this.data.subarray(start, end);
  }

  peekUint32(offset: number): number {
    return this.view.getUint32(offset, true);
  }
}

/**
 * Find the End of Central Directory record
 * Searches backwards from the end of the file
 */
function findEndOfCentralDir(data: Uint8Array): number {
  // EOCD is at least 22 bytes, search backwards
  // Comment can be up to 65535 bytes
  const minOffset = Math.max(0, data.length - 65557);
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);

  for (let i = data.length - 22; i >= minOffset; i--) {
    if (view.getUint32(i, true) === END_OF_CENTRAL_DIR_SIG) {
      return i;
    }
  }

  return -1;
}

/**
 * Find ZIP64 End of Central Directory Locator
 */
function findZip64EOCDLocator(data: Uint8Array, eocdOffset: number): number {
  // ZIP64 EOCD Locator is 20 bytes and appears right before EOCD
  const locatorOffset = eocdOffset - 20;
  if (locatorOffset < 0) {
    return -1;
  }

  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  if (view.getUint32(locatorOffset, true) === ZIP64_END_OF_CENTRAL_DIR_LOCATOR_SIG) {
    return locatorOffset;
  }

  return -1;
}

/**
 * Parse ZIP file entries from Central Directory
 */
export function parseZipEntries(data: Uint8Array, options: ZipParseOptions = {}): ZipEntryInfo[] {
  const { decodeStrings = true } = options;
  const entries: ZipEntryInfo[] = [];

  // Find End of Central Directory
  const eocdOffset = findEndOfCentralDir(data);
  if (eocdOffset === -1) {
    throw new Error("Invalid ZIP file: End of Central Directory not found");
  }

  const reader = new BinaryReader(data, eocdOffset);

  // Read EOCD
  // Offset  Size  Description
  // 0       4     EOCD signature (0x06054b50)
  // 4       2     Number of this disk
  // 6       2     Disk where central directory starts
  // 8       2     Number of central directory records on this disk
  // 10      2     Total number of central directory records
  // 12      4     Size of central directory (bytes)
  // 16      4     Offset of start of central directory
  // 20      2     Comment length
  reader.skip(4); // signature
  reader.skip(2); // disk number
  reader.skip(2); // disk where central dir starts
  reader.skip(2); // entries on this disk
  let totalEntries = reader.readUint16(); // total entries
  reader.skip(4); // central directory size (unused)
  let centralDirOffset = reader.readUint32();

  // Check for ZIP64
  const zip64LocatorOffset = findZip64EOCDLocator(data, eocdOffset);
  if (zip64LocatorOffset !== -1) {
    const locatorReader = new BinaryReader(data, zip64LocatorOffset);
    locatorReader.skip(4); // signature
    locatorReader.skip(4); // disk number with ZIP64 EOCD
    const zip64EOCDOffset = Number(locatorReader.readBigUint64());

    // Read ZIP64 EOCD
    const zip64Reader = new BinaryReader(data, zip64EOCDOffset);
    const zip64Sig = zip64Reader.readUint32();
    if (zip64Sig === ZIP64_END_OF_CENTRAL_DIR_SIG) {
      zip64Reader.skip(8); // size of ZIP64 EOCD
      zip64Reader.skip(2); // version made by
      zip64Reader.skip(2); // version needed
      zip64Reader.skip(4); // disk number
      zip64Reader.skip(4); // disk with central dir
      const zip64TotalEntries = Number(zip64Reader.readBigUint64());
      zip64Reader.skip(8); // central directory size (unused)
      const zip64CentralDirOffset = Number(zip64Reader.readBigUint64());

      // Use ZIP64 values if standard values are maxed out
      if (totalEntries === 0xffff) {
        totalEntries = zip64TotalEntries;
      }
      if (centralDirOffset === 0xffffffff) {
        centralDirOffset = zip64CentralDirOffset;
      }
    }
  }

  // Read Central Directory entries
  const centralReader = new BinaryReader(data, centralDirOffset);

  for (let i = 0; i < totalEntries; i++) {
    const sig = centralReader.readUint32();
    if (sig !== CENTRAL_DIR_HEADER_SIG) {
      throw new Error(`Invalid Central Directory header signature at entry ${i}`);
    }

    // Central Directory File Header format:
    // Offset  Size  Description
    // 0       4     Central directory file header signature (0x02014b50)
    // 4       2     Version made by
    // 6       2     Version needed to extract
    // 8       2     General purpose bit flag
    // 10      2     Compression method
    // 12      2     File last modification time
    // 14      2     File last modification date
    // 16      4     CRC-32
    // 20      4     Compressed size
    // 24      4     Uncompressed size
    // 28      2     File name length
    // 30      2     Extra field length
    // 32      2     File comment length
    // 34      2     Disk number where file starts
    // 36      2     Internal file attributes
    // 38      4     External file attributes
    // 42      4     Relative offset of local file header
    // 46      n     File name
    // 46+n    m     Extra field
    // 46+n+m  k     File comment

    centralReader.skip(2); // version made by
    centralReader.skip(2); // version needed
    const flags = centralReader.readUint16();
    const compressionMethod = centralReader.readUint16();
    const lastModTime = centralReader.readUint16();
    const lastModDate = centralReader.readUint16();
    const crc32 = centralReader.readUint32();
    let compressedSize = centralReader.readUint32();
    let uncompressedSize = centralReader.readUint32();
    const fileNameLength = centralReader.readUint16();
    const extraFieldLength = centralReader.readUint16();
    const commentLength = centralReader.readUint16();
    centralReader.skip(2); // disk number start
    centralReader.skip(2); // internal attributes
    const externalAttributes = centralReader.readUint32();
    let localHeaderOffset = centralReader.readUint32();

    // Check for UTF-8 flag (bit 11)
    const isUtf8 = (flags & 0x800) !== 0;
    const useUtf8 = decodeStrings && isUtf8;

    const fileName = centralReader.readString(fileNameLength, useUtf8);
    const extraField = centralReader.readBytes(extraFieldLength);
    const comment = centralReader.readString(commentLength, useUtf8);

    // Parse extra field for ZIP64 values
    if (extraFieldLength > 0) {
      const parsed = parseZip64ExtraField(
        extraField,
        compressedSize,
        uncompressedSize,
        localHeaderOffset
      );
      compressedSize = parsed.compressedSize;
      uncompressedSize = parsed.uncompressedSize;
      localHeaderOffset = parsed.localHeaderOffset;
    }

    const isDirectory = fileName.endsWith("/") || (externalAttributes & 0x10) !== 0;
    const isEncrypted = (flags & 0x01) !== 0;

    entries.push({
      path: fileName,
      isDirectory,
      compressedSize,
      uncompressedSize,
      compressionMethod,
      crc32,
      lastModified: parseDateTime(lastModDate, lastModTime),
      localHeaderOffset,
      comment,
      externalAttributes,
      isEncrypted
    });
  }

  return entries;
}

/**
 * Extract file data for a specific entry
 */
export async function extractEntryData(data: Uint8Array, entry: ZipEntryInfo): Promise<Uint8Array> {
  if (entry.isDirectory) {
    return new Uint8Array(0);
  }

  if (entry.isEncrypted) {
    throw new Error(`File "${entry.path}" is encrypted and cannot be extracted`);
  }

  const reader = new BinaryReader(data, entry.localHeaderOffset);

  // Read local file header
  const sig = reader.readUint32();
  if (sig !== LOCAL_FILE_HEADER_SIG) {
    throw new Error(`Invalid local file header signature for "${entry.path}"`);
  }

  reader.skip(2); // version needed
  reader.skip(2); // flags
  reader.skip(2); // compression method
  reader.skip(2); // last mod time
  reader.skip(2); // last mod date
  reader.skip(4); // crc32
  reader.skip(4); // compressed size
  reader.skip(4); // uncompressed size
  const fileNameLength = reader.readUint16();
  const extraFieldLength = reader.readUint16();

  reader.skip(fileNameLength);
  reader.skip(extraFieldLength);

  // Extract compressed data
  const compressedData = reader.readBytes(entry.compressedSize);

  // Decompress if needed
  if (entry.compressionMethod === COMPRESSION_STORED) {
    return compressedData;
  } else if (entry.compressionMethod === COMPRESSION_DEFLATE) {
    return decompress(compressedData);
  } else {
    throw new Error(`Unsupported compression method: ${entry.compressionMethod}`);
  }
}

/**
 * Extract file data synchronously (Node.js only)
 */
export function extractEntryDataSync(data: Uint8Array, entry: ZipEntryInfo): Uint8Array {
  if (entry.isDirectory) {
    return new Uint8Array(0);
  }

  if (entry.isEncrypted) {
    throw new Error(`File "${entry.path}" is encrypted and cannot be extracted`);
  }

  const reader = new BinaryReader(data, entry.localHeaderOffset);

  // Read local file header
  const sig = reader.readUint32();
  if (sig !== LOCAL_FILE_HEADER_SIG) {
    throw new Error(`Invalid local file header signature for "${entry.path}"`);
  }

  reader.skip(2); // version needed
  reader.skip(2); // flags
  reader.skip(2); // compression method
  reader.skip(2); // last mod time
  reader.skip(2); // last mod date
  reader.skip(4); // crc32
  reader.skip(4); // compressed size
  reader.skip(4); // uncompressed size
  const fileNameLength = reader.readUint16();
  const extraFieldLength = reader.readUint16();

  reader.skip(fileNameLength);
  reader.skip(extraFieldLength);

  // Extract compressed data
  const compressedData = reader.readBytes(entry.compressedSize);

  // Decompress if needed
  if (entry.compressionMethod === COMPRESSION_STORED) {
    return compressedData;
  } else if (entry.compressionMethod === COMPRESSION_DEFLATE) {
    return decompressSync(compressedData);
  } else {
    throw new Error(`Unsupported compression method: ${entry.compressionMethod}`);
  }
}

/**
 * High-level ZIP parser class
 */
export class ZipParser {
  private data: Uint8Array;
  private entries: ZipEntryInfo[];
  private entryMap: Map<string, ZipEntryInfo>;

  constructor(data: Uint8Array | ArrayBuffer, options: ZipParseOptions = {}) {
    this.data = data instanceof ArrayBuffer ? new Uint8Array(data) : data;
    this.entries = parseZipEntries(this.data, options);
    this.entryMap = new Map(this.entries.map(e => [e.path, e]));
  }

  /**
   * Get all entries in the ZIP file
   */
  getEntries(): ZipEntryInfo[] {
    return this.entries;
  }

  /**
   * Get entry by path
   */
  getEntry(path: string): ZipEntryInfo | undefined {
    return this.entryMap.get(path);
  }

  /**
   * Check if entry exists
   */
  hasEntry(path: string): boolean {
    return this.entryMap.has(path);
  }

  /**
   * List all file paths
   */
  listFiles(): string[] {
    return this.entries.map(e => e.path);
  }

  /**
   * Extract a single file (async)
   */
  async extract(path: string): Promise<Uint8Array | null> {
    const entry = this.entryMap.get(path);
    if (!entry) {
      return null;
    }
    return extractEntryData(this.data, entry);
  }

  /**
   * Extract a single file (sync, Node.js only)
   */
  extractSync(path: string): Uint8Array | null {
    const entry = this.entryMap.get(path);
    if (!entry) {
      return null;
    }
    return extractEntryDataSync(this.data, entry);
  }

  /**
   * Extract all files (async)
   */
  async extractAll(): Promise<Map<string, Uint8Array>> {
    const result = new Map<string, Uint8Array>();
    for (const entry of this.entries) {
      const data = await extractEntryData(this.data, entry);
      result.set(entry.path, data);
    }
    return result;
  }

  /**
   * Iterate over entries with async callback
   */
  async forEach(
    callback: (entry: ZipEntryInfo, getData: () => Promise<Uint8Array>) => Promise<boolean | void>
  ): Promise<void> {
    for (const entry of this.entries) {
      let dataPromise: Promise<Uint8Array> | null = null;
      const getData = () => {
        if (!dataPromise) {
          dataPromise = extractEntryData(this.data, entry);
        }
        return dataPromise;
      };

      const shouldContinue = await callback(entry, getData);
      if (shouldContinue === false) {
        break;
      }
    }
  }
}
