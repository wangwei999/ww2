import ExcelJS from 'exceljs';

async function checkFile(filePath, label) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  
  console.log(`\n=== ${label} ===`);
  
  workbook.eachSheet((worksheet, sheetId) => {
    const sheetName = worksheet.name.trim();
    console.log(`\n工作表: ${worksheet.name}`);
    
    if (sheetName.includes('单体')) {
      // 查看R列（第18列）的数据
      console.log('\nR列（第18列）数据（前10行）:');
      for (let row = 1; row <= 10; row++) {
        const cell = worksheet.getCell(row, 18);
        const model = cell.model || {};
        const hasFormula = model.formula !== undefined || model.sharedFormula !== undefined;
        
        if (cell.value !== null && cell.value !== undefined) {
          console.log(`  行${row}: 值=${cell.value}, 公式=${model.formula || '无'}`);
        }
      }
      
      // 查看D列公式范围
      console.log('\nD列公式（前10行）:');
      for (let row = 1; row <= 10; row++) {
        const cell = worksheet.getCell(row, 4);
        const model = cell.model || {};
        
        if (model.formula) {
          console.log(`  行${row}: 公式=${model.formula}`);
        }
      }
    }
  });
}

await checkFile('/tmp/original.xlsx', '原B文件');
