#!/usr/bin/env node
import { CliUsageError, main } from "./lib/cli-support.mjs";

try {
  await main(process.argv.slice(2));
} catch (error) {
  process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
  process.exitCode = error instanceof CliUsageError ? 2 : 1;
}
