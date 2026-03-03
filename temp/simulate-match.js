const XLSX = require('xlsx');

// 模拟完整的匹配流程

function normalizeDate(dateStr) {
  const patterns = [
    { pattern: /(\d{4})[\/\-\.](\d{1,2})/, format: 'YM' },
    { pattern: /(\d{4})年(\d{1,2})月/, format: 'YM_CN' },
    { pattern: /(\d{4})年底/, format: 'YEAR_END' },
    { pattern: /(\d{4})年(\d{1,2})月底/, format: 'MONTH_END_CN' },
    { pattern: /(\d{4})年末/, format: 'YEAR_END' },
  ];

  const trimmed = String(dateStr).trim();

  for (const { pattern, format } of patterns) {
    const match = trimmed.match(pattern);
    if (match) {
      let year, month;

      if (format === 'YM' || format === 'YM_CN' || format === 'MONTH_END_CN') {
        year = match[1];
        month = match[2].padStart(2, '0');
      } else if (format === 'YEAR_END') {
        year = match[1];
        month = '12';
      }

      return `${year}-${month}`;
    }
  }

  return trimmed;
}

function isExcelDate(value) {
  if (typeof value !== 'number' || isNaN(value)) {
    return false;
  }

  if (value < 1 || value > 60000) {
    return false;
  }

  if (Number.isInteger(value)) {
    if (value >= 1900 && value <= 2100) return false;
    if (value < 100) return false;
  }

  return true;
}

function excelDateToString(excelDate) {
  if (typeof excelDate !== 'number' || isNaN(excelDate)) {
    return null;
  }

  const excelEpoch = new Date(1900, 0, 1);
  const daysToAdd = excelDate - 2;
  const date = new Date(excelEpoch.getTime() + daysToAdd * 24 * 60 * 60 * 1000);

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');

  return `${year}-${month}-${day}`;
}

function extractTimePointsFromRow(headers) {
  const timePoints = new Map();

  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    let normalized = null;

    if (typeof header === 'number' && isExcelDate(header)) {
      const fullDate = excelDateToString(header);
      if (fullDate) {
        normalized = fullDate.substring(0, 7);
      }
    } else if (header) {
      normalized = normalizeDate(String(header));
    }

    if (normalized && header && normalized !== String(header)) {
      timePoints.set(i, normalized);
    }
  }

  return timePoints;
}

console.log('=== 读取数据源文件 ===');
const wb1 = XLSX.readFile('/tmp/指标测试.xlsx');
const ws1 = wb1.Sheets['季末监管指标'];
const rawData1 = XLSX.utils.sheet_to_json(ws1, { header: 1, defval: null, raw: true });

console.log('表头:', rawData1[1]);

// 模拟纵向表格的数据索引构建
console.log('\n=== 构建源数据索引（纵向表格）===');
const sourceDataMap = new Map();
const sourceTimeColumnIndex = 0; // 第一列是时间

for (const row of rawData1.slice(2)) { // 跳过前两行
  const timeCell = row[0];
  if (!timeCell) continue;

  const normalizedTime = normalizeDate(String(timeCell));
  console.log(`时间: ${timeCell} -> ${normalizedTime}`);

  if (!sourceDataMap.has(normalizedTime)) {
    sourceDataMap.set(normalizedTime, new Map());
  }

  const timeMap = sourceDataMap.get(normalizedTime);

  // 遍历所有列
  for (let col = 0; col < rawData1[1].length; col++) {
    if (col === sourceTimeColumnIndex) continue;

    const header = rawData1[1][col];
    const value = row[col];

    if (header && value !== null && value !== undefined && !isNaN(value)) {
      timeMap.set(header.trim(), value);
      console.log(`  ${header}: ${value}`);
    }
  }
}

console.log('\n源数据索引时间点:', Array.from(sourceDataMap.keys()));

console.log('\n=== 读取目标文件 ===');
const wb2 = XLSX.readFile('/tmp/1110(1)(1).xlsx');
const ws2 = wb2.Sheets['Sheet1'];
const rawData2 = XLSX.utils.sheet_to_json(ws2, { header: 1, defval: null, raw: true });

console.log('表头:', rawData2[0]);
const timePointsInRow = extractTimePointsFromRow(rawData2[0]);
console.log('目标文件时间点:', Array.from(timePointsInRow.entries()));

console.log('\n=== 模拟填充 ===');
let filledCount = 0;

// 遍历目标文件的每一行
for (let rowIndex = 0; rowIndex < rawData2.length; rowIndex++) {
  const row = rawData2[rowIndex];
  const fieldName = row[1]; // 第二列是字段名

  if (!fieldName || typeof fieldName !== 'string') continue;

  console.log(`处理 Row ${rowIndex}: ${fieldName}`);

  // 遍历每一列（每列代表一个时间点）
  for (const [colIndex, normalizedTime] of timePointsInRow.entries()) {
    const targetCell = row[colIndex];

    // 只填充空单元格，跳过前两列
    if (colIndex < 2 || targetCell !== null) continue;

    const sourceTimeMap = sourceDataMap.get(normalizedTime);
    if (!sourceTimeMap) {
      console.log(`  列${colIndex} 时间${normalizedTime}: 源数据中不存在`);
      continue;
    }

    // 查找匹配的源数据
    let matchedValue = null;
    if (sourceTimeMap.has(fieldName)) {
      matchedValue = sourceTimeMap.get(fieldName);
    }

    if (matchedValue !== null) {
      filledCount++;
      console.log(`  ✓ 列${colIndex} 时间${normalizedTime}: 填充 ${matchedValue}`);
    } else {
      console.log(`  ✗ 列${colIndex} 时间${normalizedTime}: 未找到匹配数据`);
    }
  }
}

console.log(`\n总共填充: ${filledCount} 个单元格`);
