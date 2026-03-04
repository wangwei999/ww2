const ExcelJS = require('exceljs');

async function checkDateFormats() {
  console.log('=== 检查日期格式 ===\n');

  // 检查源文件
  const sourceWorkbook = new ExcelJS.Workbook();
  await sourceWorkbook.xlsx.readFile('/tmp/授信2026.xlsx');
  const sourceSheet = sourceWorkbook.getWorksheet('单体');

  console.log('=== 源文件行98（浙江萧山农村商业银行股份有限公司）===');
  const cellC = sourceSheet.getCell(98, 3);
  console.log(`C98值: ${cellC.value} (type: ${cellC.type})`);
  console.log(`C98格式: ${cellC.numFmt ? cellC.numFmt : '无格式'}`);
  console.log(`C98是否为日期: ${cellC.type === 4}`);

  // 检查目标文件原始格式
  const targetWorkbook = new ExcelJS.Workbook();
  await targetWorkbook.xlsx.readFile('/tmp/A类授信调整_v2.xlsx');
  const targetSheet = targetWorkbook.worksheets[0];

  console.log('\n=== 目标文件B6（浙江萧山农商银行股份有限公司）===');
  const cellB = targetSheet.getCell(6, 2);
  const cellC_target = targetSheet.getCell(6, 3);
  console.log(`B6值: ${cellB.value}`);
  console.log(`C6原始值: ${cellC_target.value} (type: ${cellC_target.type})`);
  console.log(`C6原始格式: ${cellC_target.numFmt ? cellC_target.numFmt : '无格式'}`);

  // 检查输出文件
  const outputWorkbook = new ExcelJS.Workbook();
  await outputWorkbook.xlsx.readFile('/workspace/projects/temp/1772598956756_A类授信调整_v2.xlsx');
  const outputSheet = outputWorkbook.worksheets[0];

  console.log('\n=== 输出文件B6 ===');
  const cellC_output = outputSheet.getCell(6, 3);
  console.log(`C6输出值: ${cellC_output.value} (type: ${cellC_output.type})`);
  console.log(`C6输出格式: ${cellC_output.numFmt ? cellC_output.numFmt : '无格式'}`);
  console.log(`C6是否为日期: ${cellC_output.type === 4}`);
}

checkDateFormats().catch(console.error);
