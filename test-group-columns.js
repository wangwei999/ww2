const ExcelJS = require('exceljs');

async function testGroupColumns() {
  const workbook = new ExcelJS.Workbook();
  
  try {
    await workbook.xlsx.readFile('/workspace/projects/assets/授信2026.xlsx');
    
    const groupSheet = workbook.getWorksheet('集团') || workbook.getWorksheet('集团 ');
    if (!groupSheet) {
      console.log('未找到集团表');
      return;
    }
    
    console.log('=== 集团表第3行（字段名称）===');
    console.log('从A列到AB列:');
    for (let col = 1; col <= 28; col++) {
      const cell = groupSheet.getCell(3, col);
      const value = String(cell.value || '').trim();
      if (value) {
        console.log(`  列${col}(${String.fromCharCode(64 + col)}): ${value}`);
      }
    }
    
    console.log('\n=== 集团表第4行（北京银行股份有限公司）数值分布 ===');
    console.log('从E列到AB列:');
    for (let col = 5; col <= 28; col++) {
      const cell = groupSheet.getCell(4, col);
      const value = cell.value;
      if (value !== null && value !== undefined && value !== 0) {
        const fieldNameCell = groupSheet.getCell(3, col);
        const fieldName = String(fieldNameCell.value || '').trim();
        console.log(`  列${col}(${String.fromCharCode(64 + col)}): ${fieldName} = ${value}`);
      }
    }
    
    // 对比单体表
    const singleSheet = workbook.getWorksheet('单体');
    console.log('\n=== 单体表第3行（字段名称）===');
    console.log('从E列到AB列:');
    for (let col = 5; col <= 28; col++) {
      const cell = singleSheet.getCell(3, col);
      const value = String(cell.value || '').trim();
      if (value) {
        console.log(`  列${col}(${String.fromCharCode(64 + col)}): ${value}`);
      }
    }
    
  } catch (error) {
    console.error('错误:', error);
  }
}

testGroupColumns();
