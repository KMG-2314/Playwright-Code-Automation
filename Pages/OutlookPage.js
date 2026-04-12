// class OutlookPage {
//   constructor(page) {
//     this.page = page;
//   }

//   async getAuth(email, password) {

//     await this.page.goto('https://outlook.office.com/mail/');
//     await this.page.waitForTimeout(5000);

//     // ================= EMAIL =================
//     console.log('Entering email...');
//     await this.page.locator('#i0116').fill(email);
//     await this.page.locator('#idSIButton9').click();

//     // ================= PASSWORD =================
//     await this.page.locator('#i0118').waitFor({ timeout: 30000 });
//     await this.page.locator('#i0118').fill(password);
//     await this.page.locator('#idSIButton9').click();


//       // ================= VERIFY YOUR IDENTITY (If this Screen comes) =================
//   try {
//     const verifyTitle = this.page.locator('text="Verify your identity"');

//     if (await verifyTitle.isVisible({ timeout: 8000 })) {
//       console.log('🛂 Verify identity screen detected');

//       // Click FIRST option: "Approve a request on my Microsoft Authenticator app"
//       const firstOption = this.page.locator(
//         'text="Approve a request on my Microsoft Authenticator app"'
//       );

//       await firstOption.click({ timeout: 5000 });
//       console.log('✅ Clicked Authenticator approval option');
//     }
//   } catch {
//     console.log('ℹ Verify identity screen not shown, continuing normal flow');
//   }


//     // ================= MFA WAIT =================
//     console.log('⏳ Waiting for MFA / Stay signed in / Inbox...');

//     await Promise.race([
//       // Stay signed in popup
//       this.page.waitForSelector('text="Stay signed in?"', { timeout: 180000 }),

//       // Inbox directly
//       this.page.waitForSelector('div[role="option"]', { timeout: 180000 })
//     ]);

//     console.log('✅ MFA completed');

//     // ================= STAY SIGNED IN =================
//     try {
//       const yesBtn = this.page.locator('#idSIButton9');
//       if (await yesBtn.isVisible({ timeout: 5000 })) {
//         console.log('👉 Clicking YES on Stay signed in');
//         await yesBtn.click();
//       }
//     } catch {}

//     // ================= WAIT FOR INBOX =================
//     await this.page.waitForSelector('div[role="option"]', { timeout: 60000 });
//     console.log('📬 Inbox loaded');

//     // ================= OTP LOGIC =================
// // ================= OTP LOGIC =================
// console.log('🔍 Searching STAI OTP mail (last 5 min only)...');

// const startTime = Date.now();
// const MAX_WAIT = 180000; // 3 min wait loop

// while (Date.now() - startTime < MAX_WAIT) {

//   // Soft refresh inbox
//   await this.page.keyboard.press('F5');
//   await this.page.waitForTimeout(4000);

//   const mails = this.page.locator('div[role="option"]');
//   const count = await mails.count();

//   console.log(`📨 Total mails visible: ${count}`);

//   for (let i = 0; i < count; i++) {

//     const mail = mails.nth(i);
//     const text = await mail.innerText();

//     // ✅ Exact sender match
//     if (!text.includes('STAI (via Microsoft)')) continue;

//     console.log('✅ Valid STAI mail found');
//     await mail.click();
//     await this.page.waitForTimeout(3000);
//     await this.page.waitForSelector('[role="document"]', { timeout: 20000 });
//     console.log('✅ Mail document ready');

// // ================= READ OTP (LAZY LOAD FIX) =================
// // *** SPECIAL COMMENT #2: Multiple fallback selectors - works 100% ***

// // Wait for ANY mail body indicator (most reliable)
// await Promise.race([
//   this.page.waitForSelector('[data-test-id="mailMessageBodyContainer"]', { timeout: 20000 }),
//   this.page.waitForSelector('[role="document"][aria-labelledby*="UniqueMessageBody"]', { timeout: 20000 }),
//   this.page.waitForSelector('div.XbIp4[id*="UniqueMessageBody"]', { timeout: 20000 }),
//   this.page.waitForSelector('td[id="x_i5"]', { timeout: 20000 })  // OTP cell DIRECT!
// ]);

// console.log('✅ Mail body loaded (one of 4 selectors)');

// // *** SPECIAL COMMENT #3: OTP DIRECT FROM CELL (Screenshot se proven) ***
// const otpCell = this.page.locator('td[id="x_i5"], td:has-text("Account verification code:") + td');
// await otpCell.waitFor({ timeout: 5000 });

// const otpText = await otpCell.innerText();
// console.log('📄 OTP Cell text:', otpText);

// const otp = otpText.match(/\b\d{6,8}\b/)?.[0];
// if (!otp) {
//   // Fallback: full page scan
//   const fullText = await this.page.innerText();
//   const fallbackOtp = fullText.match(/\b\d{6,8}\b/)?.[0];
//   if (fallbackOtp) {
//     console.log('✅ Fallback OTP from full page:', fallbackOtp);
//     return fallbackOtp;
//   }
//   throw new Error(`❌ No OTP found. OTP cell: "${otpText}"`);
// }

// console.log('✅ OTP Extracted DIRECT from cell:', otp);
// return otp;


//   }

//   console.log('⏳ No valid STAI OTP mail yet...');
//   await this.page.waitForTimeout(10000);
// }

// throw new Error('❌ OTP mail not received in time');

//   }
// }

// module.exports = { OutlookPage };


class OutlookPage {
  constructor(page) {
    this.page = page;
  }

  async getAuth() {

    // Already logged in via auth-outlook.json
    await this.page.goto('https://outlook.office.com/mail/');

    await this.page.waitForSelector('div[role="option"]', { timeout: 60000 });
    console.log('📬 Outlook inbox ready');

    console.log('🔍 Searching STAI OTP mail...');

    const startTime = Date.now();
    const MAX_WAIT = 180000;

    while (Date.now() - startTime < MAX_WAIT) {

      await this.page.keyboard.press('F5');
      await this.page.waitForTimeout(4000);

      const mails = this.page.locator('div[role="option"]');
      const count = await mails.count();

      console.log(`📨 Mails: ${count}`);

      for (let i = 0; i < count; i++) {

        const mail = mails.nth(i);
        const text = await mail.innerText();

        if (!text.includes('STAI (via Microsoft)')) continue;

        console.log('✅ OTP mail found');
        await mail.click();

        await this.page.waitForSelector('[role="document"]', { timeout: 20000 });

        const otpCell = this.page.locator(
          'td[id="x_i5"], td:has-text("Account verification code:") + td'
        );

        await otpCell.waitFor({ timeout: 5000 });

        const otpText = await otpCell.innerText();
        const otp = otpText.match(/\b\d{6,8}\b/)?.[0];

        if (otp) {
          console.log('✅ OTP:', otp);
          return otp;
        }
      }

      console.log('⏳ Waiting for OTP mail...');
      await this.page.waitForTimeout(10000);
    }

    throw new Error('❌ OTP not found');
  }
}

module.exports = { OutlookPage };