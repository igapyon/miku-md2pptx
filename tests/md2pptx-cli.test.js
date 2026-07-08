import { describe, expect, it } from "vitest";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { spawnSync } from "node:child_process";
import { markdownToPptx } from "../dist/core.js";
import { unzipStoredEntries } from "./helpers/zip.js";
import packageJson from "../package.json" with { type: "json" };

describe("miku-md2pptx CLI", () => {
  it("prints version", () => {
    const result = spawnSync(process.execPath, ["scripts/miku-md2pptx-cli.mjs", "--version"], { encoding: "utf8" });

    expect(result.status).toBe(0);
    expect(result.stdout.trim()).toBe(packageJson.version);
  });

  it("prints agent-readable help", () => {
    const result = spawnSync(process.execPath, ["scripts/miku-md2pptx-cli.mjs", "--help"], { encoding: "utf8" });

    expect(result.status).toBe(0);
    expect(result.stdout).toContain("converts a Markdown file into a PowerPoint .pptx deck");
    expect(result.stdout).toContain("Examples:");
    expect(result.stdout).toContain("Markdown handling notes:");
  });

  it("writes a pptx file", async () => {
    const dir = await mkdtemp(join(tmpdir(), "miku-md2pptx-"));
    const out = join(dir, "out.pptx");
    try {
      const result = spawnSync(process.execPath, [
        "scripts/miku-md2pptx-cli.mjs",
        "tests/fixtures/smoke.md",
        "--out",
        out
      ], { encoding: "utf8" });

      expect(result.status).toBe(0);
      const entries = unzipStoredEntries(await readFile(out));
      expect(entries.get("ppt/slides/slide1.xml")).toContain("Sales memo");
      expect(entries.get("ppt/slides/slide2.xml")).toContain("<a:tbl>");
      expect(entries.get("ppt/slides/slide2.xml")).toContain("<a:t>Apple</a:t>");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("writes a pptx file using a template without copying template slides", async () => {
    const dir = await mkdtemp(join(tmpdir(), "miku-md2pptx-"));
    const input = join(dir, "input.md");
    const template = join(dir, "template.pptx");
    const out = join(dir, "out.pptx");
    try {
      await writeFile(input, "# Generated\n\nGenerated body\n", "utf8");
      await writeFile(template, markdownToPptx("# Template\n\nTemplate-only body"));
      const result = spawnSync(process.execPath, [
        "scripts/miku-md2pptx-cli.mjs",
        input,
        "--out",
        out,
        "--template",
        template
      ], { encoding: "utf8" });

      expect(result.status).toBe(0);
      expect(result.stderr).toContain("info: template-layout-selected:");
      const entries = unzipStoredEntries(await readFile(out));
      expect(entries.get("ppt/slides/slide1.xml")).toContain("Generated body");
      expect(entries.get("ppt/slides/slide1.xml")).not.toContain("Template-only body");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });

  it("prints conversion diagnostics to stderr", async () => {
    const dir = await mkdtemp(join(tmpdir(), "miku-md2pptx-"));
    const input = join(dir, "input.md");
    const out = join(dir, "out.pptx");
    try {
      await writeFile(input, "# Diagnostics\n\n![Missing](assets/missing.png)\n", "utf8");
      const result = spawnSync(process.execPath, [
        "scripts/miku-md2pptx-cli.mjs",
        input,
        "--out",
        out
      ], { encoding: "utf8" });

      expect(result.status).toBe(0);
      expect(result.stderr).toContain("warning: skipped-image: Markdown image was not embedded: assets/missing.png");
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
