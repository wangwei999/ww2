const ExcelJS = require('exceljs');

async function checkOutputFile() {
  console.log('=== 检查输出文件 ===\n');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/tmp/output_latest.xlsx');

  const sheet = workbook.getWorksheet('1月批量调整 (2)');
  
  console.log('工作表名称:', sheet.name);
  console.log('工作表行数:', sheet.rowCount);

  console.log('\n=== 检查前15行数据 ===\n');

  for (let row = 1; row <= Math.min(15, sheet.rowCount); row++) {
    const cellB = sheet.getCell(row, 2);
    const cellC = sheet.getCell(row, 3);
    const cellD = sheet.getCell(row, 4);
    const cellN = sheet.getCell(row, 14);

    const orgName = cellB.value;
    const valueC = cellC.value;
    const valueD = cellD.value;
    const valueN = cellN.value;

    console.log(`行${row}:`);
    console.log(`  B列(机构): ${orgName}`);
    console.log(`  C列(原授信时间): ${valueC} (${typeof valueC})`);
    console.log(`  D列(原授信): ${valueD} (${typeof valueD})`);
    console.log(`  N列(审批额度): ${valueN} (${typeof valueN})`);
    
    if (row >= 6 && valueD !== null && valueD !== undefined) {
      console.log(`  ✓ 已填充数据!`);
    }
  }
}

checkOutputFile().catch(console.error);
