import { describe, expect, it } from "vitest";
import { access, mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { constants } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { spawnSync } from "node:child_process";

const onePixelPng = new Uint8Array([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
  0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 31, 21, 196,
  137, 0, 0, 0, 13, 73, 68, 65, 84, 120, 156, 99, 248, 255, 255,
  63, 0, 5, 254, 2, 254, 167, 53, 129, 132, 0, 0, 0, 0,
  73, 69, 78, 68, 174, 66, 96, 130
]);

async function pathExists(path) {
  try {
    await access(path, constants.R_OK);
    return true;
  } catch {
    return false;
  }
}

describe("miku-pptx2md compatibility", () => {
  it("round-trips the first structural scope through the reverse converter", async () => {
    const pptx2mdDir = resolve("..", "miku-pptx2md");
    const pptx2mdCli = join(pptx2mdDir, "scripts/miku-pptx2md-cli.mjs");
    if (!(await pathExists(pptx2mdCli))) {
      console.warn("Skipping miku-pptx2md compatibility test because ../miku-pptx2md is not available.");
      return;
    }

    const dir = await mkdtemp(join(tmpdir(), "miku-md2pptx-compat-"));
    const inputPath = join(dir, "input.md");
    const assetsDir = join(dir, "assets");
    const imagePath = join(assetsDir, "chart.png");
    const pptxPath = join(dir, "out.pptx");
    const mdPath = join(dir, "roundtrip.md");
    const summaryPath = join(dir, "summary.json");
    try {
      await mkdir(assetsDir, { recursive: true });
      await writeFile(imagePath, onePixelPng);
      await writeFile(inputPath, `# Sales memo

Intro paragraph for the deck.

See [project site](https://example.com/project).

![Chart preview](assets/chart.png)

<!-- speaker-notes: Mention the chart source during the talk. -->

## Numbers

| Item | Quantity |
| --- | --- |
| Apple | 3 |
| Orange | 5 |

- Confirm quantities
- Share with the team

\`\`\`text
plain code
\`\`\`
`, "utf8");

      const generate = spawnSync(process.execPath, [
        "scripts/miku-md2pptx-cli.mjs",
        inputPath,
        "--out",
        pptxPath
      ], { encoding: "utf8" });
      expect(generate.status, generate.stderr).toBe(0);

      const reverse = spawnSync(process.execPath, [
        pptx2mdCli,
        pptxPath,
        "--out",
        mdPath,
        "--summary-json-out",
        summaryPath
      ], {
        cwd: pptx2mdDir,
        encoding: "utf8"
      });
      expect(reverse.status, reverse.stderr).toBe(0);

      const markdown = await readFile(mdPath, "utf8");
      const summary = JSON.parse(await readFile(summaryPath, "utf8"));

      expect(markdown).toContain("# Sales memo");
      expect(markdown).toContain("## Slide 1: Sales memo");
      expect(markdown).toContain("Intro paragraph for the deck.");
      expect(markdown).toContain("[project site](https://example.com/project)");
      expect(markdown).toContain("[Image: Chart preview]");
      expect(markdown).toContain("### Speaker Notes");
      expect(markdown).toContain("Mention the chart source during the talk.");
      expect(markdown).toContain("## Slide 2: Numbers");
      expect(markdown).toContain("| Item | Quantity |");
      expect(markdown).toContain("| Apple | 3 |");
      expect(markdown).toContain("- Confirm quantities");
      expect(markdown).not.toContain("- - Confirm quantities");
      expect(summary.summary.slides).toBe(2);
      expect(summary.summary.slidesWithTitles).toBe(2);
      expect(summary.summary.tables).toBe(1);
      expect(summary.summary.listItems).toBe(2);
      expect(summary.summary.hyperlinks).toBe(1);
      expect(summary.summary.imageAssets).toBe(1);
      expect(summary.summary.notesSlides).toBe(1);
      expect(summary.assets[0].altText).toBe("Chart preview");
      expect(summary.assets[0].sourcePath).toBe("ppt/media/image1.png");
      expect(summary.summary.errors).toBe(0);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});
