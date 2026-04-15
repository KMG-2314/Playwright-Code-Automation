/**
 * OutlookComposePage — UPDATED to support multiple attachments.
 * 
 * CHANGE: composeMail() now accepts `attachments` (string[]) instead of
 *         single `attachmentPath`, so both the Excel report AND Timesheet.csv
 *         are attached in one email.
 * Backward compatible: `attachmentPath` still works (wrapped into array).
 */

class OutlookComposePage {
  constructor(page) { this.page = page; }

  async login() {
    await this.page.goto('https://outlook.office.com/mail/');
    await this.page.waitForSelector('div[role="option"]', { timeout: 60000 });
  }

  async waitForInbox() {}

  async _fillRecipientField(fieldLocator, emailsRaw) {
    const list = Array.isArray(emailsRaw)
      ? emailsRaw
      : String(emailsRaw).split(',').map(e => e.trim()).filter(Boolean);
    if (!list.length) return;
    await fieldLocator.click();
    await this.page.waitForTimeout(400);
    for (const email of list) {
      await this.page.keyboard.type(email, { delay: 55 });
      await this.page.waitForTimeout(500);
      await this.page.keyboard.press('Semicolon');
      await this.page.waitForTimeout(600);
      console.log(`✅ Recipient: ${email}`);
    }
    await this.page.keyboard.press('Tab');
    await this.page.waitForTimeout(400);
  }

  async _attachFile(filePath) {
    const path     = require('path');
    const fileName = path.basename(filePath);
    console.log(`📎 Attaching: ${fileName}`);

    const attachBtn = this.page.locator('[aria-label*="Attach"]').first();
    await attachBtn.click();
    const browseBtn          = this.page.getByText(/browse this computer/i);
    const fileChooserPromise = this.page.waitForEvent('filechooser');
    await browseBtn.click();
    const fileChooser = await fileChooserPromise;
    await fileChooser.setFiles(filePath);

    const nameBase   = fileName.replace(/\.[^.]+$/, '');
    const chipLocator = this.page
      .locator(`[aria-label*="${nameBase}"]`)
      .or(this.page.getByText(fileName, { exact: false }))
      .first();
    await chipLocator.waitFor({ timeout: 40000 });
    console.log(`✅ Chip visible: ${fileName}`);
    await this.page.waitForTimeout(2000);
  }

  async composeMail({ to, cc, subject, body, attachments, attachmentPath }) {
    console.log('✉️ Composing email...');

    // Normalise to array
    const filePaths = attachments
      ? (Array.isArray(attachments) ? attachments : [attachments])
      : (attachmentPath ? [attachmentPath] : []);

    if (this.page.url().includes('about:blank')) {
      await this.page.goto('https://outlook.office.com/mail/');
    }

    // Switch to Mail
    try {
      const mailIcon = this.page.locator('a[aria-label*="Mail"], button[aria-label*="Mail"]').first();
      if (await mailIcon.isVisible({ timeout: 5000 })) await mailIcon.click();
    } catch {}

    await this.page.waitForSelector('[role="treeitem"], div[role="option"]', { timeout: 60000 });

    // New mail
    let btn = this.page.getByRole('button', { name: /new mail|new message/i });
    if (!await btn.isVisible().catch(() => false))
      btn = this.page.locator('button[aria-label*="New message"]').first();
    if (!await btn.isVisible().catch(() => false))
      btn = this.page.locator('button:has-text("New")').first();
    await btn.click();

    await this.page.waitForSelector('[aria-label="To"], div[contenteditable="true"]', { timeout: 30000 });

    // To
    const toField = this.page.locator('div[aria-label="To"], div[aria-label="To recipients"]').first();
    await toField.waitFor({ timeout: 10000 });
    await this._fillRecipientField(toField, to);

    // CC
    if (cc) {
      const ccTrigger = this.page.getByRole('button', { name: /^Cc$/i, exact: true });
      if (await ccTrigger.isVisible({ timeout: 3000 }).catch(() => false)) {
        await ccTrigger.click();
        await this.page.waitForTimeout(400);
      }
      const ccField = this.page.locator('div[aria-label="Cc"], div[aria-label="Cc recipients"], div[aria-label="CC"]').first();
      await ccField.waitFor({ timeout: 10000 });
      await this._fillRecipientField(ccField, cc);
    }

    // Subject
    await this.page.locator('input[placeholder*="subject"]').fill(subject);
    console.log('📝 Subject set');

    // Attach each file
    for (const fp of filePaths) {
      await this._attachFile(fp);
    }

    // Body
    await this.page.locator('div[role="textbox"]').fill(body);
    console.log('📝 Body filled');

    // Wait for uploads
    await this.page.waitForTimeout(8000);
    await this.page.waitForFunction(() => {
      return !document.querySelector('[aria-label*="Uploading"]') &&
             !document.querySelector('.ms-Spinner');
    }, { timeout: 60000 }).catch(() => console.log('⚠️ Upload check skipped'));
    console.log('✅ Uploads complete');

    // Send with retry
    const sendBtn = this.page.locator('[data-testid="ComposeSendButton"]');
    let sent = false;
    for (let i = 0; i < 6; i++) {
      try {
        await sendBtn.waitFor({ timeout: 10000 });
        await this.page.waitForTimeout(2000);
        await sendBtn.click();
        console.log(`✅ Send attempt ${i + 1}`);
        const popup = this.page.locator('text=Please wait to send');
        if (await popup.isVisible({ timeout: 3000 }).catch(() => false)) {
          const ok = this.page.getByRole('button', { name: /ok/i });
          if (await ok.isVisible().catch(() => false)) await ok.click();
          await this.page.waitForTimeout(8000);
          continue;
        }
        sent = true;
        break;
      } catch (err) {
        console.log(`❌ Attempt ${i + 1}:`, err.message);
        await this.page.waitForTimeout(4000);
      }
    }
    if (!sent) throw new Error('❌ Email send failed');
    console.log('🎉 EMAIL SENT');
  }
}

module.exports = { OutlookComposePage };
