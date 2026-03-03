const XLSX = require('xlsx');

// 直接定义日期转换函数
function excelDateToString(excelDate) {
  if (typeof excelDate !== 'number' || isNaN(excelDate)) {
    return null;
  }
  
  // Excel日期基准：1900-01-01 = 1
  const excelEpoch = new Date(1900, 0, 1);
  const daysToAdd = excelDate - 2; // Excel有一个1900年2月29日的bug
  
  const date = new Date(excelEpoch.getTime() + daysToAdd * 24 * 60 * 60 * 1000);
  
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  
  return `${year}-${month}-${day}`;
}

// 读取Excel文件
const workbook = XLSX.readFile('/workspace/projects/temp/testB.xlsx');
const worksheet = workbook.Sheets['Sheet1'];

// 读取原始数据
const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null, raw: false });

console.log('原始数据第1行:');
const firstRow = rawData[0];
for (let i = 0; i < firstRow.length; i++) {
  const colLetter = String.fromCharCode(65 + (i % 26));
  const cell = firstRow[i];
  console.log(`${colLetter}1:`, cell, `(类型: ${typeof cell})`);
  
  // 测试日期转换
  if (typeof cell === 'number') {
    const dateStr = excelDateToString(cell);
    console.log(`  -> 转换为日期: ${dateStr}`);
  }
}

// 测试日期转换
const testDates = [44713, 45107, 45199, 45473, 45565, 45838, 45930];
console.log('\n测试Excel日期序列号转换:');
testDates.forEach(dateNum => {
  const dateStr = excelDateToString(dateNum);
  console.log(`${dateNum} -> ${dateStr}`);
});
