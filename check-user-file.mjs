import ExcelJS from 'exceljs';

async function checkFile(filePath, label) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  
  console.log(`\n=== ${label} ===`);
  
  workbook.eachSheet((worksheet, sheetId) => {
    const sheetName = worksheet.name.trim();
    console.log(`\n工作表: ${worksheet.name}`);
    
    if (sheetName.includes('单体')) {
      // 查看第3行（表头）
      console.log('\n第3行表头（列5-20）:');
      for (let col = 5; col <= 20; col++) {
        const cell = worksheet.getCell(3, col);
        if (cell.value) {
          console.log(`  列${col}(${String.fromCharCode(64+col)}): ${cell.value}`);
        }
      }
      
      // 查看第4-6行数据（列4-20）
      for (let row = 4; row <= 6; row++) {
        console.log(`\n第${row}行数据（列4-20）:`);
        for (let col = 4; col <= 20; col++) {
          const cell = worksheet.getCell(row, col);
          const model = cell.model || {};
          
          if (cell.value !== null && cell.value !== undefined) {
            console.log(`  列${col}(${String.fromCharCode(64+col)}): 值=${cell.value}`);
          }
        }
      }
    }
  });
}

await checkFile('/tmp/user_processed.xlsx', '用户提供的处理后文件');
