const ExcelJS = require('exceljs');

async function checkTargetOriginal() {
  console.log('=== 检查目标文件原始格式 ===\n');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/tmp/A类授信调整_v2.xlsx');
  const sheet = workbook.worksheets[0];

  console.log('=== B6-C6-D6-N6 ===');
  console.log(`B6格式: ${sheet.getCell(6, 2).numFmt || '无'}`);
  console.log(`C6格式: ${sheet.getCell(6, 3).numFmt || '无'}`);
  console.log(`D6格式: ${sheet.getCell(6, 4).numFmt || '无'}`);
  console.log(`N6格式: ${sheet.getCell(6, 14).numFmt || '无'}`);

  console.log('\n=== 表头行（第4-5行）格式 ===');
  for (let row = 4; row <= 5; row++) {
    console.log(`行${row}:`);
    console.log(`  C格式: ${sheet.getCell(row, 3).numFmt || '无'}`);
    console.log(`  D格式: ${sheet.getCell(row, 4).numFmt || '无'}`);
    console.log(`  N格式: ${sheet.getCell(row, 14).numFmt || '无'}`);
  }
}

checkTargetOriginal().catch(console.error);
