const ExcelJS = require('exceljs');

async function checkSourceFormats() {
  console.log('=== 检查源文件格式 ===\n');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/tmp/授信2026.xlsx');
  const sheet = workbook.getWorksheet('单体');

  console.log('=== 行98（浙江萧山）===');
  const cellC = sheet.getCell(98, 3);
  const cellD = sheet.getCell(98, 4);
  
  console.log(`C98 numFmt: "${cellC.numFmt}"`);
  console.log(`D98 numFmt: "${cellD.numFmt}"`);

  console.log('\n=== 表头行（第3-4行）格式 ===');
  for (let row = 3; row <= 4; row++) {
    const cCell = sheet.getCell(row, 3);
    const dCell = sheet.getCell(row, 4);
    console.log(`行${row}: C="${cCell.value}" (fmt: "${cCell.numFmt || '无'}"), D="${dCell.value}" (fmt: "${dCell.numFmt || '无'}")`);
  }

  console.log('\n=== 检查列级格式 ===');
  console.log(`C列宽度: ${sheet.getColumn(3).width}`);
  console.log(`D列宽度: ${sheet.getColumn(4).width}`);
  console.log(`C列numFmt: ${sheet.getColumn(3).numFmt || '无'}`);
  console.log(`D列numFmt: ${sheet.getColumn(4).numFmt || '无'}`);
}

checkSourceFormats().catch(console.error);
