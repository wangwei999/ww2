import ExcelJS from 'exceljs';

async function checkFile(filePath, label) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  
  console.log(`\n=== ${label} ===`);
  
  workbook.eachSheet((worksheet, sheetId) => {
    const sheetName = worksheet.name.trim();
    console.log(`\n工作表: ${worksheet.name}`);
    
    if (sheetName.includes('单体')) {
      // 查看第4行所有有数据的列
      console.log('\n第4行所有有数据的列:');
      for (let col = 1; col <= 30; col++) {
        const cell = worksheet.getCell(4, col);
        const model = cell.model || {};
        if (cell.value !== null && cell.value !== undefined) {
          const colLetter = col <= 26 ? String.fromCharCode(64 + col) : 'A' + String.fromCharCode(64 + col - 26);
          console.log(`  列${col}(${colLetter}): 值=${cell.value}, 公式=${model.formula || '无'}`);
        }
      }
    }
  });
}

await checkFile('/tmp/original.xlsx', '原B文件');
