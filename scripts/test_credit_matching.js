const XLSX = require('xlsx');
const path = require('path');

// 模拟CreditMatcher的核心逻辑
function testCreditMatching() {
  console.log('=== 测试授信模式匹配 ===\n');

  // 读取源文件（授信2026.xlsx）
  console.log('读取源文件：/tmp/授信2026.xlsx');
  const sourceWorkbook = XLSX.readFile('/tmp/授信2026.xlsx');
  const sourceSheet = sourceWorkbook.Sheets['单体'];
  const sourceData = XLSX.utils.sheet_to_json(sourceSheet, { header: 1 });

  console.log('源文件工作表:', sourceWorkbook.SheetNames);
  console.log('源文件数据行数:', sourceData.length);

  // 读取目标文件（A类授信调整.xlsx）
  console.log('\n读取目标文件：/tmp/A类授信调整.xlsx');
  const targetWorkbook = XLSX.readFile('/tmp/A类授信调整.xlsx');
  const targetSheet = targetWorkbook.Sheets['1月批量调整 (2)'];
  const targetData = XLSX.utils.sheet_to_json(targetSheet, { header: 1 });

  console.log('目标文件工作表:', targetWorkbook.SheetNames);
  console.log('目标文件数据行数:', targetData.length);

  // 构建源文件的机构名称索引（B列，从第4行开始，索引3）
  const sourceOrgMap = new Map();
  sourceData.slice(3).forEach((row, idx) => {
    const orgName = String(row[1] || '').trim();
    if (orgName && orgName !== '机构名称') {
      const actualRowIndex = idx + 3; // 实际行号（从0开始）
      sourceOrgMap.set(orgName, actualRowIndex);
      console.log(`源文件索引: ${orgName} -> 行${actualRowIndex + 1} (B${actualRowIndex + 4})`);
    }
  });

  console.log(`\n源文件机构索引大小: ${sourceOrgMap.size}`);

  // 遍历目标文件的机构（B列，从第6行开始，索引5）
  console.log('\n=== 开始匹配 ===');
  let matchedCount = 0;
  let unmatchedCount = 0;

  targetData.slice(5).forEach((row, idx) => {
    const orgName = String(row[1] || '').trim();
    if (orgName && orgName !== '机构名称') {
      const actualRowIndex = idx + 5; // 实际行号（从0开始）
      const sourceRowIndex = sourceOrgMap.get(orgName);

      if (sourceRowIndex !== undefined) {
        matchedCount++;
        const sourceRow = sourceData[sourceRowIndex];
        const valueC = sourceRow[2];
        const valueD = sourceRow[3];
        const valueN = valueD;

        console.log(`✓ 匹配成功: ${orgName}`);
        console.log(`  目标: B${actualRowIndex + 1}, 源: B${sourceRowIndex + 4}`);
        console.log(`  C列: ${valueC}, D列: ${valueD}, N列: ${valueN}`);
      } else {
        unmatchedCount++;
        console.log(`✗ 匹配失败: ${orgName}`);
      }
    }
  });

  console.log(`\n=== 匹配统计 ===`);
  console.log(`总机构数: ${matchedCount + unmatchedCount}`);
  console.log(`匹配成功: ${matchedCount}`);
  console.log(`匹配失败: ${unmatchedCount}`);
  console.log(`匹配率: ${((matchedCount / (matchedCount + unmatchedCount)) * 100).toFixed(2)}%`);
}

testCreditMatching();
