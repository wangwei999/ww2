const ExcelJS = require('exceljs');
const path = require('path');

async function testExcelJSReading() {
  console.log('=== 测试ExcelJS读取文件 ===\n');

  const workbookA = new ExcelJS.Workbook();
  await workbookA.xlsx.readFile('/tmp/授信2026.xlsx');

  console.log('源文件工作表:', workbookA.worksheets.map(ws => ws.name));

  const sourceSheet = workbookA.getWorksheet('单体');
  console.log('\n=== 源文件"单体"工作表 ===');
  console.log('工作表名称:', sourceSheet.name);
  console.log('工作表行数:', sourceSheet.rowCount);
  console.log('工作表列数:', sourceSheet.columnCount);

  console.log('\n前10行数据（行1-10）:');
  for (let row = 1; row <= Math.min(10, sourceSheet.rowCount); row++) {
    const cellB = sourceSheet.getCell(row, 2); // B列
    const cellC = sourceSheet.getCell(row, 3); // C列
    const cellD = sourceSheet.getCell(row, 4); // D列

    console.log(`行${row}: B=${cellB.value}, C=${cellC.value}, D=${cellD.value}`);
  }

  console.log('\n=== 目标文件 ===');
  const workbookB = new ExcelJS.Workbook();
  await workbookB.xlsx.readFile('/tmp/A类授信调整.xlsx');

  console.log('目标文件工作表:', workbookB.worksheets.map(ws => ws.name));

  const targetSheet = workbookB.getWorksheet('1月批量调整 (2)');
  console.log('\n目标文件工作表信息:');
  console.log('工作表名称:', targetSheet.name);
  console.log('工作表行数:', targetSheet.rowCount);
  console.log('工作表列数:', targetSheet.columnCount);

  console.log('\n前10行数据（行1-10）:');
  for (let row = 1; row <= Math.min(10, targetSheet.rowCount); row++) {
    const cellB = targetSheet.getCell(row, 2); // B列
    const cellC = targetSheet.getCell(row, 3); // C列
    const cellD = targetSheet.getCell(row, 4); // D列

    console.log(`行${row}: B=${cellB.value}, C=${cellC.value}, D=${cellD.value}`);
  }

  console.log('\n=== 测试机构名称索引构建 ===');
  
  // 构建源文件的机构名称索引（从第4行开始）
  const sourceOrgMap = new Map();
  for (let row = 4; row <= sourceSheet.rowCount; row++) {
    const cell = sourceSheet.getCell(row, 2); // B列
    const orgName = String(cell.value || '').trim();
    
    if (orgName && orgName !== '机构名称') {
      sourceOrgMap.set(orgName, row);
      if (sourceOrgMap.size <= 5) {
        console.log(`索引添加: ${orgName} -> 行${row}`);
      }
    }
  }

  console.log(`\n源文件机构索引大小: ${sourceOrgMap.size}`);

  // 遍历目标文件的机构（从第6行开始）
  console.log('\n=== 测试目标文件匹配 ===');
  for (let row = 6; row <= Math.min(20, targetSheet.rowCount); row++) {
    const cell = targetSheet.getCell(row, 2); // B列
    const orgName = String(cell.value || '').trim();

    if (orgName && orgName !== '机构名称') {
      const sourceRowIndex = sourceOrgMap.get(orgName);
      
      console.log(`目标行${row}: ${orgName}`);
      if (sourceRowIndex !== undefined) {
        console.log(`  ✓ 匹配成功! 源行${sourceRowIndex}`);
        
        const cellC = sourceSheet.getCell(sourceRowIndex, 3);
        const cellD = sourceSheet.getCell(sourceRowIndex, 4);
        console.log(`  C列值: ${cellC.value}, D列值: ${cellD.value}`);
      } else {
        console.log(`  ✗ 匹配失败`);
      }
    }
  }
}

testExcelJSReading().catch(console.error);
