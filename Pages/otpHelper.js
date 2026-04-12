const { chromium } = require('@playwright/test');

async function getOtpFromOutlook() {
  const browser = await chromium.launch({ headless: true });

  const context = await browser.newContext({
    storageState: 'auth-outlook.json'
  });

  const page = await context.newPage();

  await page.goto('https://outlook.office.com/mail/');

  // Wait for inbox to load
  await page.waitForSelector('div[role="option"]', { timeout: 60000 });

  let otpMatch = null;

  // 🔁 Retry logic (important because OTP mail may take few seconds)
  for (let i = 0; i < 5; i++) {
    console.log(`🔄 Checking for OTP email... Attempt ${i + 1}`);

    // Click latest email
    const firstMail = page.locator('div[role="option"]').first();
    await firstMail.click();

    // Wait for email body
    await page.waitForSelector('div[role="document"]', { timeout: 10000 });

    const bodyText = await page.locator('div[role="document"]').innerText();

    console.log("📩 Email Body:", bodyText);

    // ✅ BEST: label-based extraction (your email format)
    otpMatch = bodyText.match(/Account verification code:\s*(\d{6,8})/i);

    if (otpMatch) {
      console.log("✅ OTP Found:", otpMatch[1]);
      break;
    }

    // Wait before retry
    await page.waitForTimeout(5000);

    // Refresh inbox
    await page.reload();
  }

  await browser.close();

  if (!otpMatch) {
    throw new Error("❌ OTP not found in email after multiple attempts");
  }

  return otpMatch[1];
}

module.exports = { getOtpFromOutlook };