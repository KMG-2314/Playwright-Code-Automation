// const { test, expect } = require('@playwright/test');
// const { readConfig } = require('../Utils/excelReader');
// const { moveLatestDownload } = require('../Utils/downloadHelper');

// const { LoginPage } = require('../Pages/LoginPage');
// const { OutlookPage } = require('../Pages/OutlookPage');
// const { TimesheetPage } = require('../Pages/TimesheetPage');
// const { OutlookComposePage } = require('../Pages/OutlookComposePage');
// const path = require('path');

// test('KMG TimeTrack Automation', async ({ browser }) => {
//   test.setTimeout(240000);
//   const config = await readConfig();

//   const context = await browser.newContext({ acceptDownloads: true, viewport: null });
//   const page = await context.newPage();

//   const login = new LoginPage(page);
//   await login.login(config.kmgUrl, config.timetrackEmail);

//   await page.waitForSelector('input[aria-label*="code"], input[placeholder*="code"]');


//   // const outlookAuthPage = await context.newPage();
//   // const outlookAuth = new OutlookPage(outlookAuthPage);
//   // const otp = await outlookAuth.getAuth(config.outlookEmail, config.password);  // Same config vars
//   // 🔥 NEW OUTLOOK CONTEXT WITH SAVED SESSION
// const outlookContext = await browser.newContext({
//   storageState: 'auth-outlook.json'
// });

// const outlookAuthPage = await outlookContext.newPage();

// const outlookAuth = new OutlookPage(outlookAuthPage);
// const otp = await outlookAuth.getAuth();

// await outlookContext.close();

//   await outlookAuthPage.close();
//   await page.waitForSelector('input[aria-label*="code"], input[placeholder*="code"], input[title*="code"], input:focus', { timeout: 30000 });
//   console.log('✅ OTP input field ready');

//   await login.submitOtp(otp);

// await page.waitForURL(/\/home$/);

// // Sidebar items print karo
// const sidebarTexts = await page.locator('nav a').allTextContents();
// console.log('🔍 Sidebar items:', sidebarTexts);

//   const timesheet = new TimesheetPage(page);
//   await timesheet.open();
//   await page.waitForSelector('.ag-root-wrapper');
//   //await page.pause();

//   await timesheet.applyAgGridDateFilter(
//   'Date',
//   config.dateFilter,   // e.g. Between
//   config.startDate,    // 01-12-2025
//   config.endDate       // 09-12-2025
// );

// await page.waitForTimeout(8000);

//   await timesheet.applyAgGridSetFilter('Client', config.client);
//   await page.waitForTimeout(5000);
//   await timesheet.applyAgGridSetFilter('Billing Status', config.billingStatus);
//   await page.waitForTimeout(5000);



//   //await timesheet.exportData();
//   const download = await timesheet.exportData();
//   await page.waitForTimeout(10000);
//   await moveLatestDownload(download);
//   //await page.waitForTimeout(8000);

//   const fullCsvPath = path.resolve('Output/Reports/Timesheet.csv');
//   console.log(`FULL PATH: ${fullCsvPath}`);

//   const outlookComposePage = await context.newPage();
//   const outlookCompose = new OutlookComposePage(outlookComposePage);
//   await outlookCompose.login();
//   //await outlookCompose.waitForInbox();

//   await outlookCompose.composeMail({
//     to: config.ToMail,
//     cc: config.CcMail || '',
//     subject: config.SubjectMail,
//     body: config.TextBoxMail,
//     attachmentPath: fullCsvPath
//   });
//   console.log('COMPLETE SUCCESS: Report generated + Email sent!');
//   //await page.pause();
//   await outlookComposePage.close();
// });

const { test, expect } = require('@playwright/test');
const { readConfig } = require('../Utils/excelReader');
const { moveLatestDownload } = require('../Utils/downloadHelper');

const { LoginPage } = require('../Pages/LoginPage');
const { TimesheetPage } = require('../Pages/TimesheetPage');
const { OutlookComposePage } = require('../Pages/OutlookComposePage');
const path = require('path');

test('KMG TimeTrack Automation', async ({ browser }) => {
  test.setTimeout(240000);

  const config = await readConfig();

  // ================= TIMETRACK CONTEXT =================
  const context = await browser.newContext({
    acceptDownloads: true,
    viewport: null
  });

  const page = await context.newPage();

  // ================= LOGIN =================
  const login = new LoginPage(page);
  await login.login(config.kmgUrl, config.timetrackEmail);

  // Wait for dashboard
  await page.waitForURL(/home|dashboard/, { timeout: 30000 });
  console.log('✅ Logged into TimeTrack');

  // ================= TIMESHEET =================
  const timesheet = new TimesheetPage(page);

  await timesheet.open();
  await page.waitForSelector('.ag-root-wrapper');
  console.log('📊 Timesheet grid loaded');

  // ================= FILTERS =================
  await timesheet.applyAgGridDateFilter(
    'Date',
    config.dateFilter,
    config.startDate,
    config.endDate
  );

  await page.waitForTimeout(8000);

  await timesheet.applyAgGridSetFilter('Client', config.client);
  await page.waitForTimeout(5000);

  await timesheet.applyAgGridSetFilter('Billing Status', config.billingStatus);
  await page.waitForTimeout(5000);

  console.log('✅ Filters applied');

  // ================= EXPORT =================
  const download = await timesheet.exportData();

  await page.waitForTimeout(10000);
  await moveLatestDownload(download);

  const fullCsvPath = path.resolve('Output/Reports/Timesheet.csv');
  console.log(`📁 File ready: ${fullCsvPath}`);

  // ================= OUTLOOK CONTEXT (AUTH.JSON) =================
  const outlookContext = await browser.newContext({
    storageState: 'auth-outlook.json'
  });

  const outlookComposePage = await outlookContext.newPage();

  // 🔥 ADD THESE 2 LINES (CRITICAL FIX)
await outlookComposePage.goto('https://outlook.office.com/mail/');
await outlookComposePage.waitForSelector('div[role="option"]', { timeout: 60000 });

  const outlookCompose = new OutlookComposePage(outlookComposePage);

  console.log('📧 Outlook opened (using saved session)');

  // 🔥 NO LOGIN CALL HERE
  await outlookCompose.composeMail({
    to: config.ToMail,
    cc: config.CcMail || '',
    subject: config.SubjectMail,
    body: config.TextBoxMail,
    attachmentPath: fullCsvPath
  });

  console.log('✅ Email sent successfully');

  await outlookContext.close();

  console.log('🎉 COMPLETE SUCCESS: TimeTrack → Export → Email');
});