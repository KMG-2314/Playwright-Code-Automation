const { chromium } = require('@playwright/test');

/**
 * Fetches OTP from Outlook.
 * @param {boolean} headless - Whether to run the browser in headless mode.
 */
async function getOtpFromOutlook(headless = true) {
  console.log(`🚀 Launching Outlook browser (Headless: ${headless})...`);
  const browser = await chromium.launch({ 
    headless: headless,
    args: ['--start-maximized'] 
  });

  const context = await browser.newContext({
    storageState: 'auth-outlook.json',
    viewport: null
  });

  const page = await context.newPage();

  try {
    // 🔥 FIX: networkidle is too strict for Outlook, use 'load' or just wait for selector
    await page.goto('https://outlook.office.com/mail/', { waitUntil: 'load', timeout: 60000 });

    // Wait for inbox to load
    await page.waitForSelector('div[role="option"], [aria-label="Message list"]', { timeout: 60000 });

    let otpMatch = null;

    // 🔁 Retry logic
    for (let i = 0; i < 10; i++) {
      console.log(`🔄 Checking for OTP email... Attempt ${i + 1}`);

      // 🔥 FIX: Click Inbox first to ensure we are in the right place
      try {
        await page.click('div[title="Inbox"], [aria-label="Inbox"]', { timeout: 5000 }).catch(() => {});
        await page.waitForTimeout(2000);

        // 🔥 FIX: Click "Other" tab if it exists
        const otherTab = page.locator('button[role="tab"]:has-text("Other"), span:has-text("Other")').first();
        if (await otherTab.isVisible()) {
          console.log("📂 Switching to 'Other' tab...");
          await otherTab.click();
          await page.waitForTimeout(3000);
        }
      } catch (e) {
        console.log("ℹ️ Could not switch tabs, proceeding with current view...");
      }

      // Wait a bit for list to stabilize
      await page.waitForTimeout(2000);

      // Find all email items
      const mailItems = page.locator('div[role="option"]');
      const count = await mailItems.count();

      for (let j = 0; j < Math.min(count, 5); j++) {
        const mailItem = mailItems.nth(j);
        const mailText = await mailItem.innerText();
        const firstLine = mailText.split('\n')[0].substring(0, 50);

        // Check if this looks like a verification email
        if (mailText.includes('verification code') || mailText.includes('Microsoft') || mailText.includes('STAI')) {
          console.log(`🖱️ Clicking matching email: "${firstLine}..."`);
          await mailItem.click();

          // Wait for email body to load/update
          await page.waitForTimeout(3000);
          await page.waitForSelector('div[role="document"]', { timeout: 10000 });

          const bodyText = await page.locator('div[role="document"]').innerText();
          console.log("📩 Checking Email Body Content...");

          // Extract 6-8 digit OTP
          otpMatch = bodyText.match(/Account verification code:\s*(\d{6,8})/i) || 
                     bodyText.match(/code[:\s]+(\d{6,8})/i) ||
                     bodyText.match(/verification code[^0-9]+(\d{6,8})/i);

          if (otpMatch) {
            console.log("✅ OTP Found:", otpMatch[1]);
            break;
          } else {
            console.log("⚠️ No OTP pattern found in this email's body.");
          }
        } else {
          console.log(`⏩ Skipping unrelated email: "${firstLine}..."`);
        }
      }

      if (otpMatch) break;

      console.log("⚠️ OTP not found in first few emails. Refreshing...");
      await page.reload();
      await page.waitForSelector('div[role="option"]', { timeout: 30000 });
      await page.waitForTimeout(5000);
    }

    if (!otpMatch) {
      throw new Error("❌ OTP not found in email after multiple attempts");
    }

    return otpMatch[1];
  } finally {
    // Only close if we found it or failed after all retries
    // If headless is false, we might want to keep it open for debugging, 
    // but usually in automation we should close it.
    await browser.close();
  }
}

module.exports = { getOtpFromOutlook };