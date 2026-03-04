const ExcelJS = require('exceljs');

async function checkD6Again() {
  console.log('=== 检查D6和N6 ===\n');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/workspace/projects/temp/1772599371466_A类授信调整_v2.xlsx');
  const sheet = workbook.worksheets[0];

  const cellD6 = sheet.getCell(6, 4);
  const cellN6 = sheet.getCell(6, 14);

  console.log(`D6值: ${JSON.stringify(cellD6.value)} (type: ${cellD6.type})`);
  console.log(`D6格式: ${cellD6.numFmt || '无'}`);
  console.log(`N6值: ${JSON.stringify(cellN6.value)} (type: ${cellN6.type})`);
  console.log(`N6格式: ${cellN6.numFmt || '无'}`);

  // 检查是否是日期序列号被错误解释
  if (cellD6.type === 4) {
    console.log(`\nD6被错误地认为是日期！`);
    console.log(`如果这是数字9，不应该有type 4`);
  }
}

checkD6Again().catch(console.error);
