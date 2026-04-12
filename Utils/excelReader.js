const ExcelJS = require('exceljs');
const path = require('path');

async function readConfig() {
  const workbook = new ExcelJS.Workbook();
  const filePath = path.join(__dirname, '..', 'Data', 'Config_Data.xlsx');

  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.getWorksheet('Details');

  const config = {};
  const arrayKeys = [
    'users',
    'projects',
    'months',
    'client',
    'billingStatus'
  ];

  sheet.eachRow((row, index) => {
    if (index === 1) return;

    const key = String(row.getCell(1).value).trim();
    let value = row.getCell(2).value;

    if (value?.text) value = value.text;
    value = value ? String(value).trim() : '';

    if (arrayKeys.includes(key)) {
      config[key] =
        !value || value.toLowerCase() === 'select all'
          ? []
          : value.split(',').map(v => v.trim());
    } else {
      config[key] = value;
    }
  });

  return config;
}

module.exports = { readConfig };
