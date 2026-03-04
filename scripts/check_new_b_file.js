const ExcelJS = require('exceljs');

async function checkNewFile() {
  console.log('=== 检查新B文件实际内容 ===\n');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/tmp/A类授信调整_v2.xlsx');

  console.log('工作表列表:', workbook.worksheets.map(ws => ws.name));

  // 检查第一个工作表
  const sheet = workbook.worksheets[0];
  console.log('\n工作表名称:', sheet.name);
  console.log('工作表行数:', sheet.rowCount);

  console.log('\n=== 前20行数据（重点看B列）===\n');

  for (let row = 1; row <= Math.min(20, sheet.rowCount); row++) {
    const cellB = sheet.getCell(row, 2);
    const orgName = cellB.value;

    console.log(`行${row}: B列="${orgName}"`);
  }

  console.log('\n=== 特别检查B6、B7、B8 ===\n');
  console.log(`B6: ${sheet.getCell(6, 2).value}`);
  console.log(`B7: ${sheet.getCell(7, 2).value}`);
  console.log(`B8: ${sheet.getCell(8, 2).value}`);
  console.log(`B9: ${sheet.getCell(9, 2).value}`);

  console.log('\n=== 检查C6、D6、N6的值 ===\n');
  console.log(`C6: ${sheet.getCell(6, 3).value} (type: ${sheet.getCell(6, 3).type})`);
  console.log(`D6: ${sheet.getCell(6, 4).value} (type: ${sheet.getCell(6, 4).type})`);
  console.log(`N6: ${sheet.getCell(6, 14).value} (type: ${sheet.getCell(6, 14).type})`);
}

checkNewFile().catch(console.error);
