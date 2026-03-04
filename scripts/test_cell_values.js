const ExcelJS = require('exceljs');
const path = require('path');

async function testCellValues() {
  console.log('=== 测试单元格值提取 ===\n');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/tmp/授信2026.xlsx');

  const sheet = workbook.getWorksheet('单体');
  
  // 检查第24行的各个列（华夏银行所在行）
  console.log('=== 第24行（华夏银行）各列详情 ===');
  for (let col = 1; col <= 29; col++) {
    const cell = sheet.getCell(24, col);
    const address = cell.address;
    const value = cell.value;
    
    console.log(`\n${address}:`);
    console.log(`  type: ${cell.type}`);
    console.log(`  value: ${value}`);
    console.log(`  valueType: ${cell.valueType}`);
    
    if (typeof value === 'object' && value !== null) {
      console.log(`  value details:`, JSON.stringify(value, null, 2));
    }
  }

  console.log('\n=== 尝试提取D列的正确值 ===');
  
  for (let row = 4; row <= 30; row++) {
    const cellD = sheet.getCell(row, 4);
    const cellN = sheet.getCell(row, 14);
    const orgName = sheet.getCell(row, 2).value;
    
    console.log(`\n行${row}: ${orgName}`);
    console.log(`  D列: type=${cellD.type}, value=${cellD.value}`);
    console.log(`  N列: type=${cellN.type}, value=${cellN.value}`);
    
    // 尝试不同的取值方式
    if (typeof cellD.value === 'object' && cellD.value !== null) {
      console.log(`  D列对象:`, JSON.stringify(cellD.value));
      if (cellD.value.result !== undefined) {
        console.log(`  D列.result: ${cellD.value.result}`);
      }
      if (cellD.value.value !== undefined) {
        console.log(`  D列.value: ${cellD.value.value}`);
      }
    }
  }
}

testCellValues().catch(console.error);
