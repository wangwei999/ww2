import ExcelJS from 'exceljs';

async function checkFile(filePath, label) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  
  console.log(`\n=== ${label} ===`);
  
  workbook.eachSheet((worksheet, sheetId) => {
    const sheetName = worksheet.name.trim();
    console.log(`\n工作表: ${worksheet.name}`);
    
    if (sheetName.includes('单体')) {
      // 查看D列（第4列）的数据和公式
      console.log('\nD列（汇总列）数据（前10行）:');
      for (let row = 1; row <= 10; row++) {
        const cell = worksheet.getCell(row, 4);
        const model = cell.model || {};
        
        if (cell.value !== null && cell.value !== undefined) {
          console.log(`  行${row}: 值=${cell.value}, 公式=${model.formula || '无'}`);
        }
      }
      
      // 查看所有有数据的列（第4行）
      console.log('\n第4行所有有数据的列:');
      for (let col = 1; col <= 30; col++) {
        const cell = worksheet.getCell(4, col);
        if (cell.value !== null && cell.value !== undefined) {
          const colLetter = String.fromCharCode(64 + col);
          console.log(`  列${col}(${colLetter}): ${cell.value}`);
        }
      }
    }
  });
}

await checkFile('/tmp/user_processed.xlsx', '用户提供的处理后文件');
