// const { readConfig } = require('../Utils/excelReader');
// const { getOtpFromOutlook } = require('./otpHelper');

// class LoginPage {
//   constructor(page) {
//     this.page = page;
//   }

//   async login(url, email) {
//     const config = await readConfig();
//     const password = config.password;

//     await this.page.goto(url);

//     // Click Microsoft login
//     await this.page.waitForSelector('button:has-text("Login with Microsoft")');
//     await this.page.click('button:has-text("Login with Microsoft")');


//     // Enter email
//     await this.page.waitForSelector('input[type="email"]', { timeout: 15000 });
//     await this.page.fill('input[type="email"]', email);
//     await this.page.click('input[type="submit"]');

//     console.log('✅ Email entered');

//     // Enter password
//     await this.page.waitForSelector('input[type="password"]', { timeout: 30000 });
//     // await this.page.fill('input[type="password"]', password);
//     // await this.page.click('input[type="submit"]');
//     let isOtpVisible = false;

// try {
//   await this.page.waitForSelector('input[aria-label*="code"], input[placeholder*="code"]', { timeout: 10000 });
//   isOtpVisible = true;
// } catch (e) {
//   isOtpVisible = false;
// }

// if (isOtpVisible) {
//   console.log("🔐 OTP screen detected → using old flow");

//   const otp = await getOtpFromOutlook();
//   console.log("🔢 OTP:", otp);

//   await this.submitOtp(otp);

// } else {
//   console.log("✅ No OTP → continuing normal flow");

//   await this.page.waitForURL(/dashboard|home|app/, { timeout: 30000 });
// }

//     console.log('✅ Password entered');

//     // 🔥 HANDLE "Stay signed in?" (THIS WAS MISSING)
//     try {
//       await this.page.waitForSelector('text="Stay signed in?"', { timeout: 10000 });

//       console.log('⚠️ Stay signed in screen detected');

//       // Click YES
//       await this.page.click('input[value="Yes"]');

//       console.log('✅ Clicked YES on Stay signed in');
//     } catch (e) {
//       console.log('ℹ️ No Stay Signed In prompt');
//     }

//     // Wait for next step (OTP or direct login)
//     await Promise.race([
//       this.page.waitForSelector('input[aria-label*="code"], input[placeholder*="code"]', { timeout: 30000 }).catch(() => {}),
//       this.page.waitForURL(/dashboard|home|app/, { timeout: 30000 }).catch(() => {})
//     ]);

//     console.log('✅ Login flow progressed (OTP or redirect)');
//   }

//   async submitOtp(otp) {
//     const otpField = this.page.locator(
//       'input[aria-label*="code"], input[placeholder*="code"], input[type="text"]:visible'
//     );

//     await otpField.first().waitFor({ timeout: 60000 });
//     await otpField.first().fill(otp);

//     console.log(`🔑 OTP filled: ${otp}`);

//     try {
//       await this.page.getByRole('button', { name: /sign in|verify/i }).click();
//     } catch {
//       await this.page.locator('button:has-text("Sign in")').click({ force: true });
//     }

//     await this.page.waitForURL(/dashboard|home|app/, { timeout: 30000 });

//     console.log('✅ Login successful');
//   }
// }

// module.exports = { LoginPage };


const { readConfig } = require('../Utils/excelReader');
const { getOtpFromOutlook } = require('./otpHelper');

class LoginPage {
  constructor(page) {
    this.page = page;
  }

  async login(url, email) {
    const config = await readConfig();
    const password = config.password;

    await this.page.goto(url);

    // Click Microsoft login
    await this.page.waitForSelector('button:has-text("Login with Microsoft")');
    await this.page.click('button:has-text("Login with Microsoft")');

    // Enter email
    await this.page.waitForSelector('input[type="email"]', { timeout: 15000 });
    await this.page.fill('input[type="email"]', email);
    await this.page.click('input[type="submit"]');

    console.log('✅ Email entered');

    // Small buffer for next screen
    await this.page.waitForTimeout(3000);

    const passwordLocator = this.page.locator('input[type="password"], input[name="passwd"]');
    const otpLocator = this.page.locator('input[name="otc"], input[aria-label*="code"]');

    let isPasswordVisible = await passwordLocator.isVisible().catch(() => false);
    let isOtpVisible = await otpLocator.isVisible().catch(() => false);

    // ─────────────────────────────
    // PASSWORD FLOW
    // ─────────────────────────────
    if (isPasswordVisible) {
      console.log("🔑 Password screen detected");

      if (await passwordLocator.isVisible()) {
        await passwordLocator.fill(password);
        console.log('✅ Password filled');

        await this.page.click('input[type="submit"]');
        console.log('✅ Password submitted');
      } else {
        console.log('⚠️ Password not visible, skipping...');
      }

      // Wait and re-check OTP after password
      await this.page.waitForTimeout(3000);
      isOtpVisible = await otpLocator.isVisible().catch(() => false);
    }

    // ─────────────────────────────
    // OTP FLOW
    // ─────────────────────────────
    if (isOtpVisible) {
      console.log("🔐 OTP screen detected");

      const otp = await getOtpFromOutlook();
      console.log("🔢 OTP:", otp);

      await this.submitOtp(otp);
    }

    // ─────────────────────────────
    // HANDLE "Stay signed in?"
    // ─────────────────────────────
    try {
      await this.page.waitForSelector('text="Stay signed in?"', { timeout: 8000 });

      console.log('⚠️ Stay signed in screen detected');

      await this.page.click('input[value="Yes"]');

      console.log('✅ Clicked YES on Stay signed in');
    } catch {
      console.log('ℹ️ No Stay Signed In prompt');
    }

    // ─────────────────────────────
    // FINAL SAFE WAIT
    // ─────────────────────────────
    console.log("⏳ Waiting for login completion...");

    await this.page.waitForLoadState('networkidle');

    console.log('🎉 Login completed successfully');
  }

  async submitOtp(otp) {
    const otpField = this.page.locator(
      'input[name="otc"], input[aria-label*="code"], input[type="text"]:visible'
    );

    await otpField.first().waitFor({ timeout: 60000 });
    await otpField.first().fill(otp);

    console.log(`🔑 OTP filled: ${otp}`);

    try {
      await this.page.getByRole('button', { name: /sign in|verify/i }).click();
    } catch {
      await this.page.locator('button:has-text("Sign in")').click({ force: true });
    }

    await this.page.waitForLoadState('networkidle');

    console.log('✅ OTP login successful');
  }
}

module.exports = { LoginPage };