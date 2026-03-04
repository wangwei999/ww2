const ExcelJS = require('exceljs');

async function testSynonym() {
  const { normalizeOrganizationName } = require('./src/lib/data-utils.ts');
  
  const testCases = [
    '浙江萧山农商银行股份有限公司',
    '浙江萧山农村商业银行股份有限公司',
  ];
  
  console.log('=== 测试同义词规范化 ===\n');
  for (const name of testCases) {
    const normalized = normalizeOrganizationName(name);
    console.log(`${name} → ${normalized}`);
    console.log(`  相等? ${normalized === '浙江萧山农村商业银行股份有限公司'}`);
  }
}

testSynonym().catch(console.error);
