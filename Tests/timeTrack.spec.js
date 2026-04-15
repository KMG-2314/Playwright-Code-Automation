/**
 * KMG TimeTrack Automation — Playwright Test
 *
 * CHANGES:
 *  1. Email now attaches BOTH files:
 *       - Latest Resource_Effort_*.xlsx  (projection report)
 *       - Output/Reports/Timesheet.csv   (raw timesheet export)
 *  2. inputPath in Config_Data drives which template is used (no hardcoding).
 *  3. Python engine called with python3 (cross-platform fallback).
 */

const { test, expect } = require('@playwright/test');
const { execSync }     = require('child_process');
const fs               = require('fs');
const path             = require('path');

const { readConfig }         = require('../Utils/excelReader');
const { moveLatestDownload } = require('../Utils/downloadHelper');
const { LoginPage }          = require('../Pages/LoginPage');
const { TimesheetPage }      = require('../Pages/TimesheetPage');
const { OutlookComposePage } = require('../Pages/OutlookComposePage');

test('KMG TimeTrack Automation', async ({ browser }) => {
  test.setTimeout(300000);  // 5 min

  const config = await readConfig();

  // ═══════════════════════════════════════════════
  // 1. TIMETRACK LOGIN & EXPORT
  // ═══════════════════════════════════════════════
  const context = await browser.newContext({ acceptDownloads: true, viewport: null });
  const page    = await context.newPage();

  const login = new LoginPage(page);
  await login.login(config.kmgUrl, config.timetrackEmail);
  await page.waitForURL(/home|dashboard/, { timeout: 30000 });
  console.log('✅ Logged into TimeTrack');

  const timesheet = new TimesheetPage(page);
  await timesheet.open();
  await page.waitForSelector('.ag-root-wrapper');
  console.log('📊 Timesheet grid loaded');

  await timesheet.applyAgGridDateFilter('Date', config.dateFilter, config.startDate, config.endDate);
  await page.waitForTimeout(8000);

  await timesheet.applyAgGridSetFilter('Client', config.client);
  await page.waitForTimeout(5000);

  await timesheet.applyAgGridSetFilter('Billing Status', config.billingStatus);
  await page.waitForTimeout(5000);

  console.log('✅ Filters applied');

  const download = await timesheet.exportData();
  await page.waitForTimeout(10000);
  await moveLatestDownload(download);

  const fullCsvPath = path.resolve('Output/Reports/Timesheet.csv');
  console.log(`📁 CSV exported: ${fullCsvPath}`);

  // ═══════════════════════════════════════════════
  // 2. RESOURCE EFFORT PROJECTION ENGINE
  // ═══════════════════════════════════════════════
  console.log('🚀 Running Resource Effort Projection Engine...');
  try {
    // Try python3 first (Linux/Mac), fallback to python (Windows)
    const pythonBin = process.platform === 'win32' ? 'python' : 'python3';
    const cmd = `${pythonBin} -m Scripts.projection_engine --csv "Output/Reports/Timesheet.csv" --config "Data/Config_Data.xlsx"`;
    execSync(cmd, { stdio: 'inherit', env: { ...process.env, PYTHONIOENCODING: 'utf-8' } });
    console.log('✅ Projection Engine completed');
  } catch (err) {
    console.error('❌ Projection engine error:', err.message);
  }

  // ═══════════════════════════════════════════════
  // 3. FIND LATEST EXCEL REPORT
  // ═══════════════════════════════════════════════
  const reportDir = path.resolve('Output/Reports');
  const files     = fs.readdirSync(reportDir);

  const latestExcel = files
    .filter(f => f.startsWith('Resource_Effort') && f.endsWith('.xlsx'))
    .map(f => ({ name: f, time: fs.statSync(path.join(reportDir, f)).mtime.getTime() }))
    .sort((a, b) => b.time - a.time)[0];

  const excelPath = latestExcel
    ? path.join(reportDir, latestExcel.name)
    : null;

  if (!excelPath) {
    console.warn('⚠️ No Excel report found — will attach CSV only');
  } else {
    console.log(`📎 Excel report: ${excelPath}`);
  }

  // ═══════════════════════════════════════════════
  // 4. OUTLOOK — SEND EMAIL WITH BOTH ATTACHMENTS
  // ═══════════════════════════════════════════════
  const outlookContext  = await browser.newContext({ storageState: 'auth-outlook.json' });
  const outlookPage     = await outlookContext.newPage();

  await outlookPage.goto('https://outlook.office.com/mail/');
  await outlookPage.waitForSelector('div[role="option"]', { timeout: 60000 });

  const outlookCompose = new OutlookComposePage(outlookPage);
  console.log('📧 Outlook opened (saved session)');

  // Build attachment list: [Excel, CSV]
  const attachments = [];
  if (excelPath && fs.existsSync(excelPath)) attachments.push(excelPath);
  if (fs.existsSync(fullCsvPath))             attachments.push(fullCsvPath);

  console.log('📎 Attachments:', attachments);

  await outlookCompose.composeMail({
    to:          config.ToMail,
    cc:          config.CcMail || '',
    subject:     config.SubjectMail,
    body:        config.TextBoxMail,
    attachments: attachments,   // ← NEW: array of files
  });

  console.log('✅ Email sent successfully');
  await outlookContext.close();
  console.log('🎉 COMPLETE: TimeTrack → Export → Projection → Email');
});
