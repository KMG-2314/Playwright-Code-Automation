const fs = require('fs');
const path = require('path');
const os = require('os');

async function moveLatestDownload(download) {
  //const downloadDir = path.join(os.homedir(), 'Downloads');
  const reportDir = path.join(__dirname, '..', 'Output', 'Reports');

  if (!fs.existsSync(reportDir)) {
    fs.mkdirSync(reportDir, { recursive: true });
  }

  const fileName = download.suggestedFilename(); // Timesheet.csv
  const filePath = path.join(reportDir, fileName);

  await download.saveAs(filePath);

  console.log('📁 File saved to Reports:', fileName);
}

module.exports = { moveLatestDownload };
