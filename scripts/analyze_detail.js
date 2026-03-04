const XLSX = require('xlsx');

console.log('=== 详细分析 A文件（A类授信调整.xlsx）===\n');
const workbookA = XLSX.readFile('/tmp/A类授信调整.xlsx');
const sheet = workbookA.Sheets['1月批量调整 (2)'];
const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

console.log('完整数据（所有行）:');
data.forEach((row, idx) => {
  console.log(`行${idx + 1}:`, row.map((cell, i) => {
    const colLabel = String.fromCharCode(65 + i);
    return `[${colLabel}]:${cell === undefined || cell === null ? '' : cell}`;
  }).join('  '));
});

console.log('\n=== 详细分析 B文件（授信2026.xlsx）"单体"工作表 ===\n');
const workbookB = XLSX.readFile('/tmp/授信2026.xlsx');
const sheetB = workbookB.Sheets['单体'];
const dataB = XLSX.utils.sheet_to_json(sheetB, { header: 1 });

console.log('完整数据（前20行）:');
dataB.slice(0, 20).forEach((row, idx) => {
  console.log(`行${idx + 1}:`, row.map((cell, i) => {
    const colLabel = String.fromCharCode(65 + i);
    return `[${colLabel}]:${cell === undefined || cell === null ? '' : cell}`;
  }).join('  '));
});

console.log('\n=== 对比机构名称 ===\n');
console.log('A文件的机构名称（从B6开始）:');
data.slice(5, 15).forEach((row, idx) => {
  const orgName = row[1];
  if (orgName) {
    console.log(`  B${idx + 6}: ${orgName}`);
  }
});

console.log('\nB文件"单体"工作表的机构名称（从B4开始）:');
dataB.slice(3, 15).forEach((row, idx) => {
  const orgName = row[1];
  if (orgName) {
    console.log(`  B${idx + 4}: ${orgName}  [C列:${row[2]}, D列:${row[3]}]`);
  }
});
