const XLSX = require('xlsx');

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

function extractTimePointsFromRow(headers) {
  const timePoints = new Map();

  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    const normalized = normalizeDate(header);
    if (header && normalized !== String(header)) {
      timePoints.set(i, normalized);
    }
  }

  return timePoints;
}

console.log('=== 正确处理数据源文件（横向表格）===');
const wb1 = XLSX.readFile('/tmp/指标测试.xlsx');
const ws1 = wb1.Sheets['季末监管指标'];
const rawData1 = XLSX.utils.sheet_to_json(ws1, { header: 1, defval: null, raw: true });

console.log('表头:', rawData1[1]);
const sourceTimePoints = extractTimePointsFromRow(rawData1[1]);
console.log('源文件时间点:', Array.from(sourceTimePoints.entries()));

// 构建源数据索引（横向表格）
console.log('\n=== 构建源数据索引 ===');
const sourceDataMap = new Map();

// 遍历每一行（每行代表一个字段）
for (let rowIndex = 2; rowIndex < rawData1.length; rowIndex++) {
  const row = rawData1[rowIndex];

  // 第一列是序号，第二列是字段名
  let fieldName = null;
  if (row[1] && typeof row[1] === 'string') {
    fieldName = row[1];
  }

  if (!fieldName) continue;

  console.log(`处理 Row ${rowIndex}: ${fieldName}`);

  // 遍历每一列（每列代表一个时间点）
  for (const [colIndex, normalizedTime] of sourceTimePoints.entries()) {
    const value = row[colIndex];

    if (value !== null && value !== undefined && !isNaN(value)) {
      if (!sourceDataMap.has(normalizedTime)) {
        sourceDataMap.set(normalizedTime, new Map());
      }

      const timeMap = sourceDataMap.get(normalizedTime);
      timeMap.set(fieldName.trim(), value);

      console.log(`  ${normalizedTime}: ${value}`);
    }
  }
}

console.log('\n源数据索引时间点:', Array.from(sourceDataMap.keys()));
console.log('每个时间点的字段数量:');
for (const [time, fieldMap] of sourceDataMap.entries()) {
  console.log(`  ${time}: ${fieldMap.size} 个字段`);
  if (fieldMap.size > 0) {
    console.log(`    示例: ${Array.from(fieldMap.keys()).slice(0, 3).join(', ')}`);
  }
}

console.log('\n=== 读取目标文件 ===');
const wb2 = XLSX.readFile('/tmp/1110(1)(1).xlsx');
const ws2 = wb2.Sheets['Sheet1'];
const rawData2 = XLSX.utils.sheet_to_json(ws2, { header: 1, defval: null, raw: true });

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

function extractTimePointsFromRowWithExcel(headers) {
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

console.log('目标文件表头:', rawData2[0]);
const targetTimePoints = extractTimePointsFromRowWithExcel(rawData2[0]);
console.log('目标文件时间点:', Array.from(targetTimePoints.entries()));

console.log('\n=== 模拟填充 ===');
let filledCount = 0;

for (let rowIndex = 1; rowIndex < rawData2.length; rowIndex++) {
  const row = rawData2[rowIndex];
  const fieldName = row[1];

  if (!fieldName || typeof fieldName !== 'string') continue;

  console.log(`\n处理 Row ${rowIndex}: ${fieldName}`);

  for (const [colIndex, normalizedTime] of targetTimePoints.entries()) {
    const targetCell = row[colIndex];

    if (colIndex < 2 || targetCell !== null) continue;

    const sourceTimeMap = sourceDataMap.get(normalizedTime);
    if (!sourceTimeMap) {
      console.log(`  列${colIndex} 时间${normalizedTime}: 源数据中不存在`);
      continue;
    }

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
