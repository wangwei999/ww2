const ExcelJS = require('exceljs');

async function checkDColumn() {
  console.log('=== 检查D列格式 ===\n');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/tmp/授信2026.xlsx');
  const sheet = workbook.getWorksheet('单体');

  console.log('=== 源文件行98 ===');
  const cellC = sheet.getCell(98, 3);
  const cellD = sheet.getCell(98, 4);
  
  console.log(`C98值: ${cellC.value} (type: ${cellC.type})`);
  console.log(`C98格式: ${cellC.numFmt || '无'}`);
  console.log(`D98值: ${JSON.stringify(cellD.value)}`);
  console.log(`D98type: ${cellD.type}`);
  console.log(`D98格式: ${cellD.numFmt || '无'}`);
  
  if (typeof cellD.value === 'object') {
    console.log(`D98是对象，内容:`, cellD.value);
    if ('result' in cellD.value) {
      console.log(`D98.result: ${cellD.value.result}`);
    }
  }
}

checkDColumn().catch(console.error);
