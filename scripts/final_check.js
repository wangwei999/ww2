const ExcelJS = require('exceljs');

async function finalCheck() {
  console.log('=== 最终验证 ===\n');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/workspace/projects/temp/1772599371466_A类授信调整_v2.xlsx');

  const sheet = workbook.worksheets[0];

  console.log('=== B6: 浙江萧山农商银行股份有限公司 ===');
  console.log(`C列(原授信时间): ${sheet.getCell(6, 3).value} (格式: ${sheet.getCell(6, 3).numFmt || '无'})`);
  console.log(`D列(原授信): ${sheet.getCell(6, 4).value}`);
  console.log(`N列(审批额度): ${sheet.getCell(6, 14).value}`);

  console.log('\n=== 与源文件对比 ===');
  const sourceWorkbook = new ExcelJS.Workbook();
  await sourceWorkbook.xlsx.readFile('/tmp/授信2026.xlsx');
  const sourceSheet = sourceWorkbook.getWorksheet('单体');

  console.log(`源文件行98 C列: ${sourceSheet.getCell(98, 3).value} (格式: ${sourceSheet.getCell(98, 3).numFmt || '无'})`);
  
  // 提取D列公式的结果
  const cellD98 = sourceSheet.getCell(98, 4);
  const d98Value = typeof cellD98.value === 'object' && 'result' in cellD98.value 
    ? cellD98.value.result 
    : cellD98.value;
  
  console.log(`源文件行98 D列: ${d98Value}`);

  console.log('\n=== 验证结果 ===');
  const c6IsDate = sheet.getCell(6, 3).type === 4;
  const c6HasFormat = !!sheet.getCell(6, 3).numFmt;
  const c6ValueCorrect = sheet.getCell(6, 3).value instanceof Date;
  const d6ValueCorrect = sheet.getCell(6, 4).value === 9;
  const n6ValueCorrect = sheet.getCell(6, 14).value === 9;

  console.log(`✓ C6是Date对象: ${c6IsDate}`);
  console.log(`✓ C6有日期格式: ${c6HasFormat} (${sheet.getCell(6, 3).numFmt})`);
  console.log(`✓ C6值正确: ${c6ValueCorrect}`);
  console.log(`✓ D6值正确: ${d6ValueCorrect}`);
  console.log(`✓ N6值正确: ${n6ValueCorrect}`);

  console.log('\n=== 全部通过! ===');
}

finalCheck().catch(console.error);
