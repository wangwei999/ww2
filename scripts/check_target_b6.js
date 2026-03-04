const ExcelJS = require('exceljs');

async function checkTargetFile() {
  console.log('=== 检查B文件（A类授信调整.xlsx）实际内容 ===\n');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/tmp/A类授信调整.xlsx');

  const sheet = workbook.getWorksheet('1月批量调整 (2)');
  
  console.log('工作表名称:', sheet.name);
  console.log('工作表行数:', sheet.rowCount);

  console.log('\n=== 检查前20行数据（重点看B列）===\n');

  for (let row = 1; row <= Math.min(20, sheet.rowCount); row++) {
    const cellB = sheet.getCell(row, 2);
    const orgName = cellB.value;

    console.log(`行${row}: B列="${orgName}" (${typeof orgName})`);
  }
}

checkTargetFile().catch(console.error);
