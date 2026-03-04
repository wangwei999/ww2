const ExcelJS = require('exceljs');

async function checkLatest() {
  console.log('=== 检查最新输出文件 ===\n');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/workspace/projects/temp/1772598956756_A类授信调整_v2.xlsx');

  const sheet = workbook.worksheets[0];

  console.log(`B6: ${sheet.getCell(6, 2).value}`);
  console.log(`C6: ${sheet.getCell(6, 3).value} (type: ${sheet.getCell(6, 3).type})`);
  console.log(`D6: ${sheet.getCell(6, 4).value} (type: ${sheet.getCell(6, 4).type})`);
  console.log(`N6: ${sheet.getCell(6, 14).value} (type: ${sheet.getCell(6, 14).type})`);

  console.log('\n=== 所有数据行 ===\n');
  for (let row = 6; row <= 20; row++) {
    const cellB = sheet.getCell(row, 2);
    const orgName = cellB.value;
    if (orgName && orgName !== '机构名称' && orgName !== '   ') {
      console.log(`行${row}: B="${orgName}", C=${sheet.getCell(row, 3).value}, D=${sheet.getCell(row, 4).value}, N=${sheet.getCell(row, 14).value}`);
    }
  }
}

checkLatest().catch(console.error);
