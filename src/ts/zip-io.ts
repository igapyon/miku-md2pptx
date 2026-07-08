import { inflateRawSync } from "node:zlib";

export interface ZipFileEntry {
  path: string;
  data: Uint8Array | string;
}

const encoder = new TextEncoder();
const decoder = new TextDecoder();

const crcTable = new Uint32Array(256);
for (let i = 0; i < 256; i += 1) {
  let c = i;
  for (let k = 0; k < 8; k += 1) {
    c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  }
  crcTable[i] = c >>> 0;
}

function crc32(data: Uint8Array): number {
  let crc = 0xffffffff;
  for (const value of data) {
    crc = crcTable[(crc ^ value) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function asBytes(data: Uint8Array | string): Uint8Array {
  return typeof data === "string" ? encoder.encode(data) : data;
}

function writeUint16(buffer: Uint8Array, offset: number, value: number): void {
  buffer[offset] = value & 0xff;
  buffer[offset + 1] = (value >>> 8) & 0xff;
}

function writeUint32(buffer: Uint8Array, offset: number, value: number): void {
  buffer[offset] = value & 0xff;
  buffer[offset + 1] = (value >>> 8) & 0xff;
  buffer[offset + 2] = (value >>> 16) & 0xff;
  buffer[offset + 3] = (value >>> 24) & 0xff;
}

function readUint16(buffer: Uint8Array, offset: number): number {
  return buffer[offset] | (buffer[offset + 1] << 8);
}

function readUint32(buffer: Uint8Array, offset: number): number {
  return (buffer[offset] | (buffer[offset + 1] << 8) | (buffer[offset + 2] << 16) | (buffer[offset + 3] << 24)) >>> 0;
}

function concat(parts: Uint8Array[]): Uint8Array {
  const total = parts.reduce((sum, part) => sum + part.length, 0);
  const output = new Uint8Array(total);
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.length;
  }
  return output;
}

export function createZip(entries: ZipFileEntry[]): Uint8Array {
  const localParts: Uint8Array[] = [];
  const centralParts: Uint8Array[] = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBytes = encoder.encode(entry.path);
    const data = asBytes(entry.data);
    const crc = crc32(data);
    const localHeader = new Uint8Array(30 + nameBytes.length);

    writeUint32(localHeader, 0, 0x04034b50);
    writeUint16(localHeader, 4, 20);
    writeUint16(localHeader, 6, 0x0800);
    writeUint16(localHeader, 8, 0);
    writeUint16(localHeader, 10, 0);
    writeUint16(localHeader, 12, 0);
    writeUint32(localHeader, 14, crc);
    writeUint32(localHeader, 18, data.length);
    writeUint32(localHeader, 22, data.length);
    writeUint16(localHeader, 26, nameBytes.length);
    writeUint16(localHeader, 28, 0);
    localHeader.set(nameBytes, 30);

    localParts.push(localHeader, data);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    writeUint32(centralHeader, 0, 0x02014b50);
    writeUint16(centralHeader, 4, 20);
    writeUint16(centralHeader, 6, 20);
    writeUint16(centralHeader, 8, 0x0800);
    writeUint16(centralHeader, 10, 0);
    writeUint16(centralHeader, 12, 0);
    writeUint16(centralHeader, 14, 0);
    writeUint32(centralHeader, 16, crc);
    writeUint32(centralHeader, 20, data.length);
    writeUint32(centralHeader, 24, data.length);
    writeUint16(centralHeader, 28, nameBytes.length);
    writeUint16(centralHeader, 30, 0);
    writeUint16(centralHeader, 32, 0);
    writeUint16(centralHeader, 34, 0);
    writeUint16(centralHeader, 36, 0);
    writeUint32(centralHeader, 38, 0);
    writeUint32(centralHeader, 42, offset);
    centralHeader.set(nameBytes, 46);
    centralParts.push(centralHeader);

    offset += localHeader.length + data.length;
  }

  const centralDirectory = concat(centralParts);
  const end = new Uint8Array(22);
  writeUint32(end, 0, 0x06054b50);
  writeUint16(end, 4, 0);
  writeUint16(end, 6, 0);
  writeUint16(end, 8, entries.length);
  writeUint16(end, 10, entries.length);
  writeUint32(end, 12, centralDirectory.length);
  writeUint32(end, 16, offset);
  writeUint16(end, 20, 0);

  return concat([...localParts, centralDirectory, end]);
}

export function readZipEntries(data: Uint8Array): Map<string, Uint8Array> {
  let eocdOffset = -1;
  for (let offset = data.length - 22; offset >= 0; offset -= 1) {
    if (readUint32(data, offset) === 0x06054b50) {
      eocdOffset = offset;
      break;
    }
  }
  if (eocdOffset < 0) {
    throw new Error("Template PPTX is not a readable ZIP package.");
  }

  const entryCount = readUint16(data, eocdOffset + 10);
  let centralOffset = readUint32(data, eocdOffset + 16);
  const entries = new Map<string, Uint8Array>();

  for (let index = 0; index < entryCount; index += 1) {
    if (readUint32(data, centralOffset) !== 0x02014b50) {
      throw new Error("Template PPTX has an invalid ZIP central directory.");
    }
    const method = readUint16(data, centralOffset + 10);
    const compressedSize = readUint32(data, centralOffset + 20);
    const uncompressedSize = readUint32(data, centralOffset + 24);
    const fileNameLength = readUint16(data, centralOffset + 28);
    const extraLength = readUint16(data, centralOffset + 30);
    const commentLength = readUint16(data, centralOffset + 32);
    const localHeaderOffset = readUint32(data, centralOffset + 42);
    const fileName = decoder.decode(data.slice(centralOffset + 46, centralOffset + 46 + fileNameLength));

    if (readUint32(data, localHeaderOffset) !== 0x04034b50) {
      throw new Error(`Template PPTX has an invalid ZIP local header for ${fileName}.`);
    }
    const localNameLength = readUint16(data, localHeaderOffset + 26);
    const localExtraLength = readUint16(data, localHeaderOffset + 28);
    const payloadStart = localHeaderOffset + 30 + localNameLength + localExtraLength;
    const compressed = data.slice(payloadStart, payloadStart + compressedSize);

    if (!fileName.endsWith("/")) {
      if (method === 0) {
        entries.set(fileName, compressed);
      } else if (method === 8) {
        const inflated = inflateRawSync(compressed);
        if (inflated.byteLength !== uncompressedSize) {
          throw new Error(`Template PPTX entry has an unexpected size after inflate: ${fileName}.`);
        }
        entries.set(fileName, new Uint8Array(inflated));
      } else {
        throw new Error(`Template PPTX uses unsupported ZIP compression method ${method}: ${fileName}.`);
      }
    }

    centralOffset += 46 + fileNameLength + extraLength + commentLength;
  }

  return entries;
}
