import {
  readZipPackage,
  writeZipPackage,
  type ZipEntryInput
} from "../vendor/miku-ms-office-core-0.8.0.mjs";

export type ZipFileEntry = ZipEntryInput;

export async function createZip(entries: ZipFileEntry[]): Promise<Uint8Array> {
  return writeZipPackage(
    entries.map((entry) => ({ ...entry, compression: "deflate" as const }))
  );
}

export async function readZipEntries(data: Uint8Array): Promise<Map<string, Uint8Array>> {
  const result = await readZipPackage(data);
  const error = result.diagnostics.find((diagnostic) => diagnostic.severity === "error");
  if (error) {
    throw new Error(`Template PPTX is not a readable ZIP package: ${error.message}`);
  }
  return new Map(result.entries.map((entry) => [entry.path, entry.data]));
}
