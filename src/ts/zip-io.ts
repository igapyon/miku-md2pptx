import {
  readZipPackage,
  writeZipPackage,
  type ZipEntryInput
} from "../vendor/miku-ms-office-core-0.6.0.mjs";

export type ZipFileEntry = ZipEntryInput;

export function createZip(entries: ZipFileEntry[]): Uint8Array {
  return writeZipPackage(entries);
}

export function readZipEntries(data: Uint8Array): Map<string, Uint8Array> {
  const result = readZipPackage(data);
  const error = result.diagnostics.find((diagnostic) => diagnostic.severity === "error");
  if (error) {
    throw new Error(`Template PPTX is not a readable ZIP package: ${error.message}`);
  }
  return new Map(result.entries.map((entry) => [entry.path, entry.data]));
}
