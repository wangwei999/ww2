const ExcelJS = require('exceljs');

async function checkSourceFile() {
  console.log('=== 检查A文件（授信2026.xlsx）"单体"工作表内容 ===\n');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/tmp/授信2026.xlsx');

  const sheet = workbook.getWorksheet('单体');
  
  console.log('工作表名称:', sheet.name);
  console.log('工作表行数:', sheet.rowCount);

  console.log('\n=== 查找"浙江萧山农村商业银行股份有限公司" ===\n');

  for (let row = 4; row <= sheet.rowCount; row++) {
    const cellB = sheet.getCell(row, 2);
    const orgName = String(cellB.value || '').trim();
    
    if (orgName.includes('浙江')) {
      console.log(`行${row}: ${orgName}`);
    }
  }

  console.log('\n=== 前30行数据 ===\n');
  for (let row = 4; row <= Math.min(30, sheet.rowCount); row++) {
    const cellB = sheet.getCell(row, 2);
    const cellC = sheet.getCell(row, 3);
    const cellD = sheet.getCell(row, 4);
    
    console.log(`行${row}: B=${cellB.value}, C=${cellC.value}, D=${cellD.value}`);
  }
}

checkSourceFile().catch(console.error);
