const ExcelJS = require('exceljs');

async function checkDateOutput() {
  console.log('=== 检查输出文件的日期格式 ===\n');

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('/workspace/projects/temp/1772599371466_A类授信调整_v2.xlsx');

  const sheet = workbook.worksheets[0];

  console.log('=== B6: 浙江萧山农商银行股份有限公司 ===');
  const cellC = sheet.getCell(6, 3);
  console.log(`C6值: ${cellC.value} (type: ${cellC.type})`);
  console.log(`C6格式: ${cellC.numFmt ? cellC.numFmt : '无格式'}`);
  console.log(`C6是否为日期: ${cellC.type === 4}`);
  
  // 如果有日期序列号，转换为日期对象显示
  if (cellC.type === 2 && cellC.value) {
    const date = new Date(Math.round((cellC.value - 25569) * 86400 * 1000));
    console.log(`C6日期表示: ${date.toISOString()}`);
  }

  console.log('\n=== 检查其他有数据的行 ===');
  for (let row = 6; row <= 20; row++) {
    const cellB = sheet.getCell(row, 2);
    const cellC = sheet.getCell(row, 3);
    const orgName = cellB.value;
    
    if (orgName && orgName !== '机构名称' && orgName !== '   ' && cellC.value) {
      console.log(`\n行${row}: ${orgName}`);
      console.log(`  C值: ${cellC.value} (type: ${cellC.type})`);
      console.log(`  C格式: ${cellC.numFmt ? cellC.numFmt : '无格式'}`);
    }
  }
}

checkDateOutput().catch(console.error);
