const ExcelJS = require('exceljs');
const path = require('path');

async function testCreditMatcher() {
  console.log('=== 测试授信模式匹配器 ===\n');

  // 读取文件
  const workbookA = new ExcelJS.Workbook();
  const workbookB = new ExcelJS.Workbook();
  
  await workbookA.xlsx.readFile('/tmp/授信2026.xlsx');
  await workbookB.xlsx.readFile('/tmp/A类授信调整.xlsx');

  console.log('源文件工作表:', workbookA.worksheets.map(ws => ws.name));
  console.log('目标文件工作表:', workbookB.worksheets.map(ws => ws.name));

  const sourceSheet = workbookA.getWorksheet('单体');
  const targetSheet = workbookB.getWorksheet('1月批量调整 (2)');

  // 构建源文件的机构名称索引（B列，从第4行开始）
  const sourceOrgMap = new Map();
  
  for (let row = 4; row <= sourceSheet.rowCount; row++) {
    const cell = sourceSheet.getCell(row, 2);
    const orgName = String(cell.value || '').trim();
    
    if (orgName && orgName !== '机构名称') {
      sourceOrgMap.set(orgName, row);
    }
  }

  console.log('源文件机构索引大小:', sourceOrgMap.size);

  // 辅助函数：解析单元格值
  function parseCellValue(value) {
    if (value === null || value === undefined || value === '') {
      return null;
    }

    if (typeof value === 'number') {
      return value;
    }

    // 如果是公式对象，提取result
    if (typeof value === 'object' && value !== null && 'result' in value) {
      const result = value.result;
      return typeof result === 'number' ? result : null;
    }

    // 日期对象（保留原始值）
    if (value instanceof Date) {
      return value;
    }

    return null;
  }

  // 遍历目标文件的前10个机构
  console.log('\n=== 开始匹配前10个机构 ===');
  
  for (let row = 6; row <= Math.min(15, targetSheet.rowCount); row++) {
    const cell = targetSheet.getCell(row, 2);
    const orgName = String(cell.value || '').trim();

    if (orgName && orgName !== '机构名称') {
      const sourceRowIndex = sourceOrgMap.get(orgName);

      console.log(`\n行${row}: ${orgName}`);
      
      if (sourceRowIndex !== undefined) {
        console.log(`  ✓ 匹配成功! 源行${sourceRowIndex}`);

        // 获取源单元格的原始值
        const cellC = sourceSheet.getCell(sourceRowIndex, 3);
        const cellD = sourceSheet.getCell(sourceRowIndex, 4);
        const cellN = sourceSheet.getCell(sourceRowIndex, 14);

        console.log(`  源C列: type=${cellC.type}, value=${cellC.value}`);
        console.log(`  源D列: type=${cellD.type}, value=${typeof cellD.value === 'object' ? JSON.stringify(cellD.value) : cellD.value}`);
        console.log(`  源N列: type=${cellN.type}, value=${cellN.value}`);

        // 解析值
        const valueC = parseCellValue(cellC.value);
        const valueD = parseCellValue(cellD.value);
        const valueN = valueD;

        console.log(`  解析后: C=${valueC}, D=${valueD}, N=${valueN}`);

        // 填充到目标文件
        const targetCellC = targetSheet.getCell(row, 3);
        const targetCellD = targetSheet.getCell(row, 4);
        const targetCellN = targetSheet.getCell(row, 14);

        // C列保留原始日期对象
        targetCellC.value = cellC.value;
        // D列使用解析后的数值
        if (valueD !== null) {
          targetCellD.value = valueD;
        }
        // N列使用解析后的数值
        if (valueN !== null) {
          targetCellN.value = valueN;
        }

        console.log(`  已填充: C=${targetCellC.value}, D=${targetCellD.value}, N=${targetCellN.value}`);
      } else {
        console.log(`  ✗ 匹配失败`);
      }
    }
  }

  // 保存修改后的文件
  const outputPath = '/tmp/A类授信调整_修复版.xlsx';
  await workbookB.xlsx.writeFile(outputPath);
  
  console.log(`\n=== 保存完成 ===`);
  console.log(`输出文件: ${outputPath}`);
}

testCreditMatcher().catch(console.error);
