const ExcelJS = require('exceljs');

async function checkOutputV2() {
  console.log('=== 检查输出文件（v2）===\n');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/workspace/projects/temp/1772598867632_A类授信调整_v2.xlsx');

  const sheet = workbook.worksheets[0];
  console.log('工作表名称:', sheet.name);

  console.log('\n=== 检查B6、C6、D6、N6 ===\n');
  console.log(`B6: ${sheet.getCell(6, 2).value}`);
  console.log(`C6: ${sheet.getCell(6, 3).value} (type: ${sheet.getCell(6, 3).type})`);
  console.log(`D6: ${sheet.getCell(6, 4).value} (type: ${sheet.getCell(6, 4).type})`);
  console.log(`N6: ${sheet.getCell(6, 14).value} (type: ${sheet.getCell(6, 14).type})`);

  console.log('\n=== 检查所有数据行 ===\n');
  for (let row = 6; row <= 20; row++) {
    const cellB = sheet.getCell(row, 2);
    const cellC = sheet.getCell(row, 3);
    const cellD = sheet.getCell(row, 4);
    const cellN = sheet.getCell(row, 14);

    const orgName = cellB.value;
    if (orgName && orgName !== '机构名称' && orgName !== '   ') {
      console.log(`行${row}: B="${orgName}", C=${cellC.value}, D=${cellD.value}, N=${cellN.value}`);
      if (cellD.value !== null && cellD.value !== undefined) {
        console.log(`  ✓ 已填充数据!`);
      }
    }
  }
}

checkOutputV2().catch(console.error);
