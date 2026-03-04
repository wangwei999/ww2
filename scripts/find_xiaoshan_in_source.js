const ExcelJS = require('exceljs');

async function findInSource() {
  console.log('=== 在源文件中查找"浙江萧山农商银行股份有限公司" ===\n');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/tmp/授信2026.xlsx');

  const sheet = workbook.getWorksheet('单体');
  
  // 查找完全匹配
  for (let row = 4; row <= sheet.rowCount; row++) {
    const cellB = sheet.getCell(row, 2);
    const orgName = String(cellB.value || '').trim();
    
    if (orgName === '浙江萧山农商银行股份有限公司') {
      console.log(`找到精确匹配在行${row}`);
      console.log(`  C列值: ${sheet.getCell(row, 3).value}`);
      console.log(`  D列值: ${sheet.getCell(row, 4).value}`);
      return;
    }
  }

  console.log('未找到精确匹配，查找包含"萧山"的机构：\n');
  
  for (let row = 4; row <= sheet.rowCount; row++) {
    const cellB = sheet.getCell(row, 2);
    const orgName = String(cellB.value || '').trim();
    
    if (orgName.includes('萧山')) {
      console.log(`行${row}: ${orgName}`);
      console.log(`  C列值: ${sheet.getCell(row, 3).value}`);
      console.log(`  D列值: ${sheet.getCell(row, 4).value}`);
    }
  }
}

findInSource().catch(console.error);
