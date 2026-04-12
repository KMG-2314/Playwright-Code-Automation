class GmailPage {
  constructor(page) {
    this.page = page;
  }

  async getOtp(email, password) {
    await this.page.goto('https://mail.google.com');

    if (await this.page.locator('input[type="email"]').isVisible()) {
      await this.page.fill('input[type="email"]', email);
      await this.page.click('text=Next');
      await this.page.fill('input[type="password"]', password);
      await this.page.click('text=Next');
      await this.page.waitForTimeout(8000);
    }

    await this.page.fill('input[name="q"]', 'from:(microsoft) newer_than:10m');
    await this.page.waitForTimeout(5000);
    await this.page.press('input[name="q"]', 'Enter');

    await this.page.locator('tr.zA').first().click();

    const body = await this.page.locator('div.a3s').last().innerText();
    const otp = body.match(/\b\d{4,8}\b/)[0];

    return otp;
  }
}

module.exports = { GmailPage };
