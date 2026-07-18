export interface OfficeDiagnostic {
  severity: "info" | "warning" | "error";
  code: string;
  message: string;
  path?: string;
}

export interface ZipEntryInput {
  path: string;
  data: Uint8Array | string;
  compression?: "store" | "deflate";
  modifiedAt?: Date;
}

export interface ZipEntry {
  path: string;
  data: Uint8Array;
  compression: "store" | "deflate";
  compressedSize: number;
  uncompressedSize: number;
  crc32: number;
  modifiedAt: Date;
}

export interface ZipWriteOptions {
  timestamp?: Date;
  compression?: "store" | "deflate";
  order?: "stable" | "input";
  compressionLevel?: number;
}

export interface OpcRelationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

export declare function readZipPackage(data: Uint8Array): {
  entries: ZipEntry[];
  diagnostics: OfficeDiagnostic[];
};

export declare function writeZipPackage(
  entries: ZipEntryInput[],
  options?: ZipWriteOptions
): Uint8Array;

export declare function escapeXmlAttribute(value: string): string;

export declare function buildOpcRelationshipsXml(
  relationships: OpcRelationship[]
): string;
