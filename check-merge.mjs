import ExcelJS from 'exceljs';

async function checkFile(filePath, label) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  
  console.log(`\n=== ${label} ===`);
  
  workbook.eachSheet((worksheet, sheetId) => {
    const sheetName = worksheet.name.trim();
    if (sheetName.includes('集团')) {
      console.log(`\n工作表: ${worksheet.name}`);
      
      // 查看合并单元格
      console.log('\n合并单元格:');
      const merges = worksheet._merges || {};
      for (const key of Object.keys(merges)) {
        console.log(`  ${key}`);
      }
      
      // 更详细地查看合并单元格
      console.log('\n详细合并信息:');
      const model = worksheet.model;
      if (model && model.merges) {
        model.merges.forEach(m => {
          console.log(`  ${JSON.stringify(m)}`);
        });
      }
      
      // 查看C列合并情况
      console.log('\nC列结构（带合并标记）:');
      for (let row = 3; row <= 15; row++) {
        const cell = worksheet.getCell(row, 3); // C列
        const cellE = worksheet.getCell(row, 5); // E列
        const isMerged = cell.isMerged;
        const master = cell.master;
        
        // 检查E列的值
        let eValue = cellE.value;
        if (typeof eValue === 'object' && eValue !== null) {
          eValue = `[object]`;
        }
        
        console.log(`  行${row}: C=${cell.value}, isMerged=${isMerged}, master行=${master?.row}, E=${eValue}`);
      }
    }
  });
}

await checkFile('/tmp/original.xlsx', '原B文件');
