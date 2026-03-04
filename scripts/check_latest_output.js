const ExcelJS = require('exceljs');

async function checkLatestOutput() {
  console.log('=== 检查最新输出文件 ===\n');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/workspace/projects/temp/1772599463809_A类授信调整_v2.xlsx');
  const sheet = workbook.worksheets[0];

  console.log('=== B6 ===');
  const cellC = sheet.getCell(6, 3);
  const cellD = sheet.getCell(6, 4);
  const cellN = sheet.getCell(6, 14);

  console.log(`C6: ${cellC.value} (type: ${cellC.type}, fmt: "${cellC.numFmt || '无'}")`);
  console.log(`D6: ${cellD.value} (type: ${cellD.type}, fmt: "${cellD.numFmt || '无'}")`);
  console.log(`N6: ${cellN.value} (type: ${cellN.type}, fmt: "${cellN.numFmt || '无'}")`);
}

checkLatestOutput().catch(console.error);
