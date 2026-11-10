import { defineConfig, devices } from "@playwright/test";

export default defineConfig({
  testDir: "./tests/browser",
  testMatch: "**/*.spec.mjs",
  fullyParallel: false,
  forbidOnly: Boolean(process.env.CI),
  workers: 1,
  timeout: 60_000,
  globalTimeout: 20 * 60_000,
  reporter: "list",
  outputDir: "workplace/browser-results",
  use: { headless: true },
  projects: [
    { name: "chromium", use: { ...devices["Desktop Chrome"] } },
    { name: "firefox", use: { ...devices["Desktop Firefox"] } },
    { name: "webkit", use: { ...devices["Desktop Safari"] } }
  ],
  webServer: { command: "node scripts/browser-test-server.mjs", url: "http://127.0.0.1:4176/health", reuseExistingServer: false, timeout: 30_000 }
});
