const XLSX = require('xlsx');

// Excel日期基准：1900-01-01 = 1
function excelDateToString(excelDate) {
  if (typeof excelDate !== 'number' || isNaN(excelDate)) {
    return excelDate;
  }

  const excelEpoch = new Date(1900, 0, 1);
  const daysToAdd = excelDate - 2; // Excel有一个1900年2月29日的bug
  const date = new Date(excelEpoch.getTime() + daysToAdd * 24 * 60 * 60 * 1000);

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');

  return `${year}-${month}-${day}`;
}

function isExcelDate(value) {
  if (typeof value !== 'number' || isNaN(value)) {
    return false;
  }

  // 1 = 1900-01-01, 60000 约等于 2064年
  if (value < 1 || value > 60000) {
    return false;
  }

  // 排除年份和常见小整数
  if (Number.isInteger(value)) {
    if (value >= 1900 && value <= 2100) return false;
    if (value < 100) return false;
  }

  return true;
}

function convertExcelValue(value) {
  return isExcelDate(value) ? excelDateToString(value) : value;
}

console.log('=== 测试 Excel 日期序列号转换 ===');

// 测试一些常见的日期序列号
const testDates = [44713, 45107, 45199, 45473, 45565, 45838, 45930];

testDates.forEach(dateNum => {
  const converted = convertExcelValue(dateNum);
  console.log(`${dateNum} -> ${converted} (isExcelDate: ${isExcelDate(dateNum)})`);
});

console.log('\n=== 读取指标测试.xlsx ===');
const wb1 = XLSX.readFile('/tmp/指标测试.xlsx');
const ws1 = wb1.Sheets['季末监管指标'];
const rawData1 = XLSX.utils.sheet_to_json(ws1, { header: 1, defval: null, raw: true });

console.log('原始表头 (Row 1):', JSON.stringify(rawData1[1]));
console.log('\n转换后表头 (Row 1):', JSON.stringify(rawData1[1].map(cell => convertExcelValue(cell))));

console.log('\n=== 读取 1110(1)(1).xlsx ===');
const wb2 = XLSX.readFile('/tmp/1110(1)(1).xlsx');
const ws2 = wb2.Sheets['Sheet1'];
const rawData2 = XLSX.utils.sheet_to_json(ws2, { header: 1, defval: null, raw: true });

console.log('原始 Row 0:', JSON.stringify(rawData2[0]));
console.log('\n转换后 Row 0:', JSON.stringify(rawData2[0].map(cell => convertExcelValue(cell))));
