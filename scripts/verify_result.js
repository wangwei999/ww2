const ExcelJS = require('exceljs');

async function checkDate() {
  console.log('=== 检查源文件行98的日期 ===\n');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/tmp/授信2026.xlsx');

  const sheet = workbook.getWorksheet('单体');
  
  console.log(`行98 B列: ${sheet.getCell(98, 2).value}`);
  console.log(`行98 C列: ${sheet.getCell(98, 3).value} (type: ${sheet.getCell(98, 3).type})`);
  console.log(`行98 D列: ${sheet.getCell(98, 4).value} (type: ${sheet.getCell(98, 4).type})`);

  // 检查输出文件
  console.log('\n=== 检查输出文件B6 ===\n');
  const outputWorkbook = new ExcelJS.Workbook();
  await outputWorkbook.xlsx.readFile('/workspace/projects/temp/1772598867632_A类授信调整_v2.xlsx');
  const outputSheet = outputWorkbook.worksheets[0];

  console.log(`B6: ${outputSheet.getCell(6, 2).value}`);
  console.log(`C6: ${outputSheet.getCell(6, 3).value} (type: ${outputSheet.getCell(6, 3).type})`);
  console.log(`D6: ${outputSheet.getCell(6, 4).value} (type: ${outputSheet.getCell(6, 4).type})`);
  console.log(`N6: ${outputSheet.getCell(6, 14).value} (type: ${outputSheet.getCell(6, 14).type})`);
}

checkDate().catch(console.error);
