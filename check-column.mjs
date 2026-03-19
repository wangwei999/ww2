import ExcelJS from 'exceljs';

async function checkFile(filePath, label) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  
  console.log(`\n=== ${label} ===`);
  
  workbook.eachSheet((worksheet, sheetId) => {
    const sheetName = worksheet.name.trim();
    if (sheetName.includes('集团')) {
      console.log(`\n工作表: ${worksheet.name}`);
      
      // 检查C列（第3列）的内容
      console.log('\nC列内容（前20行）:');
      for (let row = 1; row <= 20; row++) {
        const cell = worksheet.getCell(row, 3); // C列
        const model = cell.model || {};
        const hasFormula = model.formula !== undefined || model.sharedFormula !== undefined;
        const formula = model.formula || model.sharedFormula || '';
        
        if (cell.value !== null || hasFormula) {
          console.log(`  行${row}: 值=${cell.value}, 有公式=${hasFormula}, 公式=${formula}`);
        }
      }
      
      // 检查D列（第4列）和E列（第5列）
      console.log('\nD列和E列内容（前10行）:');
      for (let row = 1; row <= 10; row++) {
        const cellD = worksheet.getCell(row, 4);
        const cellE = worksheet.getCell(row, 5);
        const modelD = cellD.model || {};
        const modelE = cellE.model || {};
        
        console.log(`  行${row}: D=${cellD.value} (公式:${modelD.formula || modelD.sharedFormula || '无'}), E=${cellE.value} (公式:${modelE.formula || modelE.sharedFormula || '无'})`);
      }
    }
  });
}

await checkFile('/tmp/original.xlsx', '原B文件');
await checkFile('/tmp/processed.xlsx', '处理后文件');
