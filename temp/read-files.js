const XLSX = require('xlsx');
const path = require('path');

console.log('=== 指标测试.xlsx ===');
const wb1 = XLSX.readFile('/tmp/指标测试.xlsx');
console.log('SheetNames:', wb1.SheetNames);
const ws1 = wb1.Sheets[wb1.SheetNames[0]];
const data1 = XLSX.utils.sheet_to_json(ws1, { header: 1, defval: null, raw: true });
console.log('总行数:', data1.length);
console.log('\n前20行数据:');
data1.slice(0, 20).forEach((row, i) => {
  console.log('Row', i, ':', JSON.stringify(row));
});

console.log('\n\n=== 1110(1)(1).xlsx ===');
const wb2 = XLSX.readFile('/tmp/1110(1)(1).xlsx');
console.log('SheetNames:', wb2.SheetNames);
const ws2 = wb2.Sheets[wb2.SheetNames[0]];
const data2 = XLSX.utils.sheet_to_json(ws2, { header: 1, defval: null, raw: true });
console.log('总行数:', data2.length);
console.log('\n前20行数据:');
data2.slice(0, 20).forEach((row, i) => {
  console.log('Row', i, ':', JSON.stringify(row));
});
