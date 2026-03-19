import ExcelJS from 'exceljs';

async function checkFile(filePath, label) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  
  console.log(`\n=== ${label} ===`);
  
  workbook.eachSheet((worksheet, sheetId) => {
    const sheetName = worksheet.name.trim();
    if (sheetName.includes('单体')) {
      console.log(`\n工作表: ${worksheet.name}`);
      
      // 查看第3行（表头）
      console.log('\n第3行表头（列5-20）:');
      for (let col = 5; col <= 20; col++) {
        const cell = worksheet.getCell(3, col);
        if (cell.value) {
          console.log(`  列${col}(${String.fromCharCode(64+col)}): ${cell.value}`);
        }
      }
      
      // 查看第4行数据
      console.log('\n第4行数据（列4-20）:');
      for (let col = 4; col <= 20; col++) {
        const cell = worksheet.getCell(4, col);
        const model = cell.model || {};
        const hasFormula = model.formula !== undefined || model.sharedFormula !== undefined;
        
        if (cell.value !== null || hasFormula) {
          console.log(`  列${col}(${String.fromCharCode(64+col)}): 值=${cell.value}, 公式=${model.formula || '无'}`);
        }
      }
    }
  });
}

await checkFile('/tmp/original.xlsx', '原B文件');
await checkFile('/tmp/processed.xlsx', '处理后文件');
