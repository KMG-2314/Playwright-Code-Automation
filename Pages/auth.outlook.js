const { chromium } = require('@playwright/test');

(async () => {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext();
  const page = await context.newPage();

  await page.goto('https://outlook.office.com/mail/');

  console.log('👉 Login to Outlook manually (email + password + MFA)');

  // Wait for inbox
  await page.waitForSelector('div[role="option"]', { timeout: 300000 });

  await context.storageState({ path: 'auth-outlook.json' });

  console.log('✅ Outlook session saved!');

  await browser.close();
})();