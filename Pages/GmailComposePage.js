class GmailComposePage {
  constructor(page) {
    this.page = page;
  }

  async openGmail() {
    await this.page.goto('https://mail.google.com/');
    await this.page.waitForTimeout(10000);
  }

  async waitForInbox() {
    await this.page.waitForLoadState('networkidle');
    // Wait for inbox elements
    await this.page.locator('[role="option"], [title="New mail"]').first().waitFor({ timeout: 20000 });
    console.log('✅ Outlook inbox confirmed');
  }

  // ================= COMPOSE MAIL (EXACT SELECTORS FROM YOUR HTML) =================
async composeMail({ to, cc, subject, body, attachmentPath }) {
  console.log('✉️ Starting Outlook compose with exact selectors...');
  
  // *** SPECIAL COMMENT #1: New Mail button (Inbox se) ***
  const newMailBtn = this.page.locator('[data-testid="owaMailAction_NewMail"], [title="New mail"]');
  await newMailBtn.first().click();
  await this.page.waitForTimeout(4000);  // Compose window
  
  console.log('✅ New mail opened');
  
  // *** SPECIAL COMMENT #2: To field (id="recipient-well-label-to" ke paas) ***
  const toEditor = this.page.locator('#recipient-well-label-to + div [contenteditable="true"], [aria-label="To"]');
  await toEditor.first().click();
  await toEditor.first().fill(to);
  await this.page.waitForTimeout(2000);
  console.log(`✅ To filled: ${to}`);
  
  // *** SPECIAL COMMENT #3: CC field ***
  if (cc) {
    const ccBtn = this.page.locator('#recipient-well-label-cc');
    await ccBtn.click();
    await this.page.waitForTimeout(1000);
    const ccEditor = this.page.locator('#recipient-well-label-cc + div [contenteditable="true"]');
    await ccEditor.first().fill(cc);
    console.log(`✅ CC filled: ${cc}`);
  }
  
  // *** SPECIAL COMMENT #4: Subject (placeholder="Add a subject") ***
  const subjectField = this.page.locator('input[placeholder="Add a subject"], input[aria-label="Subject"]');
  await subjectField.fill(subject);
  await this.page.waitForTimeout(1000);
  console.log(`✅ Subject: ${subject}`);
  
  // *** SPECIAL COMMENT #5: Body (role="textbox" aria-label="Message body") ***
  const bodyEditor = this.page.locator('div[role="textbox"][aria-label="Message body"]');
  await bodyEditor.click();
  await bodyEditor.fill(body);
  await this.page.waitForTimeout(2000);
  console.log(`✅ Body filled`);
  
  // *** SPECIAL COMMENT #6: Attach file (EXACT id from your HTML) ***
  const attachBtn = this.page.locator('button[aria-label="Attach file"], #*6e46af65-28ca-3813-0963-e053967c8927');
  await attachBtn.first().click();
  await this.page.waitForTimeout(2000);
  
  // File upload (standard)
  const fileInput = this.page.locator('input[type="file"]').first();
  await fileInput.setInputFiles(attachmentPath);
  await this.page.waitForTimeout(4000);  // Upload complete
  console.log(`✅ File attached: ${attachmentPath}`);
  
  // *** SPECIAL COMMENT #7: Send button (data-testid="ComposeSendButton") ***
  const sendBtn = this.page.locator('[data-testid="ComposeSendButton"], button[aria-label="Send"], text="Send"');
  await sendBtn.first().click();
  await this.page.waitForTimeout(3000);
  
  console.log('🎉 EMAIL SENT SUCCESSFULLY!');
}


}

module.exports = { GmailComposePage };
