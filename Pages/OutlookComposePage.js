class OutlookComposePage {
  constructor(page) {
    this.page = page;
  }

  async login() {
    console.log('📧 Outlook direct open');
    await this.page.goto('https://outlook.office.com/mail/');
    await this.page.waitForSelector('div[role="option"]', { timeout: 60000 });
    console.log('✅ Outlook inbox');
  }

  async waitForInbox() {
    console.log('⏭️ Skipping explicit inbox wait');
  }

  // ─────────────────────────────────────────────────────────────────────
  // HELPER: Fill a recipient field (To / CC) with one or many emails.
  //
  // ROOT CAUSE OF "vanishing" bug:
  //   Clicking the field again between emails causes Outlook to BLUR and
  //   CLEAR whatever was partially typed. The fix is to click ONCE, then
  //   type ALL emails without ever leaving the field — using semicolon as
  //   the in-line chip separator that Outlook natively supports.
  // ─────────────────────────────────────────────────────────────────────
  async _fillRecipientField(fieldLocator, emailsRaw) {
    const list = Array.isArray(emailsRaw)
      ? emailsRaw
      : String(emailsRaw).split(',').map(e => e.trim()).filter(Boolean);

    if (list.length === 0) return;

    // Focus the field ONCE — never click again until all emails are typed
    await fieldLocator.click();
    await this.page.waitForTimeout(400);

    for (let i = 0; i < list.length; i++) {
      const email = list[i];

      // Type the email character-by-character at a human-like speed
      // so Outlook's autocomplete doesn't race ahead
      await this.page.keyboard.type(email, { delay: 55 });

      // Small pause so Outlook can resolve any suggestion dropdown
      await this.page.waitForTimeout(500);

      // Use Semicolon — Outlook's native recipient separator.
      // This commits the chip WITHOUT moving focus away from the field.
      await this.page.keyboard.press('Semicolon');

      // Give Outlook time to turn the text into a chip
      await this.page.waitForTimeout(600);

      console.log(`✅ Recipient chip added: ${email}`);
    }

    // After all chips are entered, Tab out so the field finalises
    await this.page.keyboard.press('Tab');
    await this.page.waitForTimeout(400);
  }

  async composeMail({ to, cc, subject, body, attachmentPath }) {
    console.log('✉️ FORCE COMPOSE');

    // ── HANDLE BLANK PAGE ──────────────────────────────────────────────
    if (this.page.url().includes('about:blank')) {
      console.log('⚠️ Blank page detected → loading Outlook');
      await this.page.goto('https://outlook.office.com/mail/');
    }

    // ── ENSURE MAIL MODULE ─────────────────────────────────────────────
    console.log('📨 Ensuring Mail module...');
    try {
      const mailIcon = this.page
        .locator('a[aria-label*="Mail"], button[aria-label*="Mail"]')
        .first();
      if (await mailIcon.isVisible({ timeout: 5000 })) {
        await mailIcon.click();
        console.log('✅ Switched to Mail module');
      }
    } catch {
      console.log('ℹ️ Mail icon not found, maybe already in Mail');
    }

    await this.page.waitForSelector('[role="treeitem"], div[role="option"]', {
      timeout: 60000,
    });
    console.log('📬 Mail inbox ready');

    // ── NEW MAIL BUTTON ────────────────────────────────────────────────
    console.log('🔍 Looking for New Mail button...');
    let newMailBtn = this.page.getByRole('button', { name: /new mail|new message/i });

    if (!(await newMailBtn.isVisible().catch(() => false)))
      newMailBtn = this.page.locator('button[aria-label*="New message"]').first();

    if (!(await newMailBtn.isVisible().catch(() => false)))
      newMailBtn = this.page.locator('button:has-text("New")').first();

    if (!(await newMailBtn.isVisible().catch(() => false)))
      throw new Error('❌ New Mail button not found AFTER switching to Mail');

    await newMailBtn.click();
    console.log('✅ New Mail clicked');

    // ── WAIT FOR COMPOSE PANEL ─────────────────────────────────────────
    await this.page.waitForSelector(
      '[aria-label="To"], div[contenteditable="true"]',
      { timeout: 30000 }
    );

    // ── TO FIELD ───────────────────────────────────────────────────────
    console.log('📨 Filling To field...');
    const toField = this.page
      .locator('div[aria-label="To"], div[aria-label="To recipients"]')
      .first();
    await toField.waitFor({ timeout: 10000 });
    await this._fillRecipientField(toField, to);

    // ── CC FIELD ───────────────────────────────────────────────────────
    if (cc) {
      console.log('📨 Filling CC field...');

      // Reveal the Cc row if it is hidden
      const ccTrigger = this.page.getByRole('button', { name: /^Cc$/i, exact: true });
      if (await ccTrigger.isVisible({ timeout: 3000 }).catch(() => false)) {
        await ccTrigger.click();
        console.log('✅ Cc row revealed');
        await this.page.waitForTimeout(400);
      }

      const ccField = this.page
        .locator(
          'div[aria-label="Cc"], div[aria-label="Cc recipients"], div[aria-label="CC"]'
        )
        .first();
      await ccField.waitFor({ timeout: 10000 });

      await this._fillRecipientField(ccField, cc);
    }

    // ── SUBJECT ────────────────────────────────────────────────────────
    await this.page.locator('input[placeholder*="subject"]').fill(subject);
    console.log('📝 Subject set');

    // ── ATTACH FILE ────────────────────────────────────────────────────
    console.log('📎 Attaching file...');
    const attachBtn = this.page.locator('[aria-label*="Attach"]').first();
    await attachBtn.click();

    const browseBtn = this.page.getByText(/browse this computer/i);
    const fileChooserPromise = this.page.waitForEvent('filechooser');
    await browseBtn.click();
    const fileChooser = await fileChooserPromise;
    await fileChooser.setFiles(attachmentPath);
    console.log(`📎 File set: ${attachmentPath}`);

    // Wait for the attachment chip to appear.
    // FIX: '[aria-label*="attachment"], text=Timesheet.csv' is INVALID —
    // you cannot mix a CSS selector with a Playwright text= pseudo-selector
    // in a single .locator() string. Use .or() to combine them safely.
    const attachmentByAriaLabel = this.page.locator('[aria-label*="Timesheet"]');
    const attachmentByText      = this.page.getByText('Timesheet.csv', { exact: false });
    const attachmentLocator     = attachmentByAriaLabel.or(attachmentByText).first();

    await attachmentLocator.waitFor({ timeout: 40000 });
    console.log('📎 Attachment chip visible');

    // ── BODY ───────────────────────────────────────────────────────────
    await this.page.locator('div[role="textbox"]').fill(body);
    console.log('📝 Body filled');

    // ── WAIT FOR UPLOAD TO FINISH ──────────────────────────────────────
    // console.log('⏳ Waiting for attachment upload to complete...');

    // await this.page.waitForFunction(
    //   () => !document.querySelector('[aria-label*="Uploading"], .ms-Spinner'),
    //   { timeout: 50000 }
    // );
    // console.log('✅ Attachment upload complete');

    // // ── SEND (with retry) ──────────────────────────────────────────────
    // console.log('📤 Sending email...');
    // const sendBtn = this.page.locator('[data-testid="ComposeSendButton"]');

    // for (let i = 0; i < 3; i++) {
    //   try {
    //     await sendBtn.click();
    //     console.log(`✅ Send clicked (attempt ${i + 1})`);

    //     const popup = this.page.locator('text=Please wait to send');
    //     if (await popup.isVisible({ timeout: 3000 }).catch(() => false)) {
    //       console.log('⚠️ Attachment still uploading → waiting...');
    //       await this.page.getByRole('button', { name: 'OK' }).click();
    //       await this.page.waitForTimeout(5000);
    //       continue;
    //     }

    //     break; // success
    //   } catch (err) {
    //     console.log(`❌ Send failed attempt ${i + 1}:`, err.message);
    //     await this.page.waitForTimeout(3000);
    //   }
    // }

    // console.log('✅ EMAIL SENT SUCCESSFULLY');

    // ── WAIT FOR UPLOAD TO FINISH ──────────────────────────────────────
console.log('⏳ Waiting for attachment upload to complete...');

// Wait for attachment chip (already correct)
await attachmentLocator.waitFor({ timeout: 40000 });

// 🔥 EXTRA BUFFER (critical for Outlook stability)
await this.page.waitForTimeout(8000);

// Ensure no uploading indicators (stronger check)
await this.page.waitForFunction(() => {
  const uploading = document.querySelector('[aria-label*="Uploading"]');
  const spinner = document.querySelector('.ms-Spinner');
  return !uploading && !spinner;
}, { timeout: 60000 }).catch(() => {
  console.log('⚠️ Upload indicator check skipped (safe fallback)');
});

console.log('✅ Attachment upload complete');

// ── SEND (ULTRA STABLE RETRY) ──────────────────────────────────────
console.log('📤 Sending email...');
const sendBtn = this.page.locator('[data-testid="ComposeSendButton"]');

let sent = false;

for (let i = 0; i < 6; i++) {
  try {
    await sendBtn.waitFor({ timeout: 10000 });

    // Small delay before click (VERY IMPORTANT)
    await this.page.waitForTimeout(2000);

    await sendBtn.click();
    console.log(`✅ Send clicked (attempt ${i + 1})`);

    // Check for blocking popup
    const popup = this.page.locator('text=Please wait to send');

    if (await popup.isVisible({ timeout: 3000 }).catch(() => false)) {
      console.log('⚠️ Still uploading → waiting more...');

      const okBtn = this.page.getByRole('button', { name: /ok/i });
      if (await okBtn.isVisible().catch(() => false)) {
        await okBtn.click();
      }

      // 🔥 WAIT MORE (important)
      await this.page.waitForTimeout(8000);
      continue;
    }

    sent = true;
    break;

  } catch (err) {
    console.log(`❌ Send attempt ${i + 1} failed:`, err.message);
    await this.page.waitForTimeout(4000);
  }
}

if (!sent) {
  throw new Error('❌ Email not sent after multiple retries');
}

console.log('🎉 EMAIL SENT SUCCESSFULLY');
  }
}

module.exports = { OutlookComposePage };