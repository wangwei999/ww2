const XLSX = require('xlsx');

// 读取Excel文件
const workbook = XLSX.readFile('/workspace/projects/temp/testB.xlsx');

console.log('工作表数量:', workbook.SheetNames.length);
console.log('工作表名称:', workbook.SheetNames);

// 遍历所有工作表
for (const sheetName of workbook.SheetNames) {
  console.log('\n========================================');
  console.log('工作表:', sheetName);
  console.log('========================================');
  
  const worksheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
  
  console.log('总行数:', data.length);
  
  // 打印前20行的详细信息
  console.log('\n前20行数据:');
  for (let i = 0; i < Math.min(20, data.length); i++) {
    const row = data[i];
    console.log(`\n第 ${i + 1} 行:`);
    
    if (Array.isArray(row)) {
      row.forEach((cell, colIndex) => {
        const colLetter = String.fromCharCode(65 + (colIndex % 26));
        console.log(`  ${colLetter}${i + 1}:`, cell, `(类型: ${typeof cell})`);
      });
    }
  }
  
  // 特别检查第1行的H-N列（索引7-13）
  console.log('\n第1行H-N列的详细信息:');
  const firstRow = data[0];
  if (Array.isArray(firstRow)) {
    for (let i = 7; i <= 13 && i < firstRow.length; i++) {
      const colLetter = String.fromCharCode(65 + (i % 26));
      console.log(`${colLetter}1:`, firstRow[i], `(类型: ${typeof firstRow[i]})`);
    }
  }
  
  // 检查B列（索引1）
  console.log('\nB列（字段名）的详细信息:');
  for (let i = 0; i < Math.min(10, data.length); i++) {
    const row = data[i];
    if (Array.isArray(row) && row.length > 1) {
      console.log(`第 ${i + 1} 行B列:`, row[1], `(类型: ${typeof row[1]})`);
    }
  }
}
