const ExcelJS = require('exceljs');

async function findInTargetFile() {
  console.log('=== 在B文件中查找"浙江萧山农村商业银行股份有限公司" ===\n');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/tmp/A类授信调整.xlsx');

  const sheet = workbook.getWorksheet('1月批量调整 (2)');
  
  console.log('遍历所有行...\n');

  for (let row = 1; row <= sheet.rowCount; row++) {
    const cellB = sheet.getCell(row, 2);
    const orgName = String(cellB.value || '').trim();
    
    if (orgName === '浙江萧山农村商业银行股份有限公司') {
      console.log(`找到"浙江萧山农村商业银行股份有限公司"在行${row}`);
      console.log(`  C列值: ${sheet.getCell(row, 3).value}`);
      console.log(`  D列值: ${sheet.getCell(row, 4).value}`);
      console.log(`  N列值: ${sheet.getCell(row, 14).value}`);
    }
  }

  console.log('\n=== 所有B列数据 ===\n');
  for (let row = 1; row <= sheet.rowCount; row++) {
    const cellB = sheet.getCell(row, 2);
    const orgName = String(cellB.value || '').trim();
    
    if (orgName && orgName !== '机构名称' && orgName !== '   ') {
      console.log(`行${row}: ${orgName}`);
    }
  }
}

findInTargetFile().catch(console.error);
