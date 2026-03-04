const XLSX = require('xlsx');
const path = require('path');

console.log('=== 分析 B文件（授信2026.xlsx）===\n');
const workbookB = XLSX.readFile('/tmp/授信2026.xlsx');
console.log('工作表列表:', workbookB.SheetNames);
console.log('工作表数量:', workbookB.SheetNames.length);

const sheetB = workbookB.Sheets[workbookB.SheetNames[0]];
const dataB = XLSX.utils.sheet_to_json(sheetB, { header: 1 });

console.log('\n前15行数据:');
dataB.slice(0, 15).forEach((row, idx) => {
  console.log(`行${idx + 1}:`, row.map((cell, i) => `[列${String.fromCharCode(65 + i)}]:${cell}`).join('  '));
});

console.log('\n=== 分析 A文件（A类授信调整.xlsx）===\n');
const workbookA = XLSX.readFile('/tmp/A类授信调整.xlsx');
console.log('工作表列表:', workbookA.SheetNames);
console.log('工作表数量:', workbookA.SheetNames.length);

// 分析每个工作表
workbookA.SheetNames.forEach((sheetName, idx) => {
  console.log(`\n--- 工作表 ${idx + 1}: ${sheetName} ---`);
  const sheet = workbookA.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  console.log(`总行数: ${data.length}`);
  console.log(`前10行数据:`);
  data.slice(0, 10).forEach((row, rowIdx) => {
    console.log(`  行${rowIdx + 1}:`, row.map((cell, i) => `[${String.fromCharCode(65 + i)}]:${cell}`).join('  '));
  });

  // 检查B列的机构名称（从B4开始）
  if (data.length > 3) {
    console.log(`\nB列的机构名称（从B4开始，前10个）:`);
    data.slice(3, 13).forEach((row, rowIdx) => {
      const orgName = row[1]; // B列索引为1
      if (orgName) {
        const rowLabel = rowIdx + 4;
        console.log(`  B${rowLabel}: ${orgName}  [C列:${row[2]}, D列:${row[3]}]`);
      }
    });
  }
});
