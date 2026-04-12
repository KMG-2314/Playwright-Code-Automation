const { defineConfig } = require('@playwright/test');

module.exports = defineConfig({
  testDir: './Tests',
  timeout: 120000,
  use: {
    headless: false,
    viewport: null,
    screenshot: 'only-on-failure',
    actionTimeout: 20000,  // 20 sec for clicks
    navigationTimeout: 30000,
    force: true
  }
});
