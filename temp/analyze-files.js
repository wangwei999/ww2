const XLSX = require('xlsx');

// 日期标准化函数
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

// Excel日期转换
function excelDateToString(excelDate) {
  if (typeof excelDate !== 'number' || isNaN(excelDate)) return excelDate;

  const excelEpoch = new Date(1900, 0, 1);
  const daysToAdd = excelDate - 2;
  const date = new Date(excelEpoch.getTime() + daysToAdd * 24 * 60 * 60 * 1000);

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');

  return `${year}-${month}-${day}`;
}

// 同义词检查
const SYNONYMS = {
  '总资产': ['资产总额', '资产合计'],
  '贷款总额': ['贷款额', '贷款金额', '贷款'],
  '资本充足率': ['资本充足率（%）', '核心资本充足率', '核心一级资本充足率'],
  '不良贷款率': ['不良贷款率（%）', '不良资产率'],
  '杠杆率': ['杠杆率（%）'],
  '流动性': ['流动性比例', '流动性覆盖率'],
};

function fieldMatches(fieldA, fieldB) {
  if (!fieldA || !fieldB) return false;

  let normalizedA = String(fieldA).toLowerCase().trim();
  let normalizedB = String(fieldB).toLowerCase().trim();

  // 清理字段名
  const clean = (s) => s.replace(/[（\(][^\)）]*[）\)]/g, '').replace(/%$/g, '').trim();

  const cleanedA = clean(normalizedA);
  const cleanedB = clean(normalizedB);

  // 精确匹配
  if (cleanedA === cleanedB) return true;

  // 包含匹配
  if (cleanedA.includes(cleanedB) || cleanedB.includes(cleanedA)) return true;

  // 同义词匹配
  for (const [key, synonyms] of Object.entries(SYNONYMS)) {
    const group = [key, ...synonyms].map(s => clean(s).toLowerCase());
    if (group.includes(cleanedA) && group.includes(cleanedB)) {
      return true;
    }
  }

  return false;
}

// 分析文件
console.log('=== 分析数据源文件（指标测试.xlsx）===');
const wb1 = XLSX.readFile('/tmp/指标测试.xlsx');
const ws1 = wb1.Sheets['季末监管指标'];
const data1 = XLSX.utils.sheet_to_json(ws1, { header: 1, defval: null, raw: true });

const header1 = data1[1];
console.log('表头:', header1.map((h, i) => {
  const normalized = normalizeDate(h);
  return i > 1 ? `${h} → ${normalized}` : h;
}));

// 提取有数据的字段
const sourceFields = new Set();
const sourceDataByTime = {};

// 从第8行开始（索引7）有实际数据
for (let i = 8; i < data1.length; i++) {
  const row = data1[i];
  if (row[1] && typeof row[1] === 'string') {
    sourceFields.add(row[1]);
  }
}

console.log('\n数据源字段:', Array.from(sourceFields));

// 检查每个时间点的数据
for (let col = 2; col < header1.length; col++) {
  const timePoint = normalizeDate(header1[col]);
  sourceDataByTime[timePoint] = [];

  for (let i = 8; i < data1.length; i++) {
    const row = data1[i];
    if (row[1] && row[col] !== null && row[col] !== undefined) {
      sourceDataByTime[timePoint].push({
        field: row[1],
        value: row[col]
      });
    }
  }
}

console.log('\n各时间点的数据:');
Object.entries(sourceDataByTime).forEach(([time, items]) => {
  console.log(`${time}: ${items.length} 个字段有数据`);
  if (items.length > 0) {
    items.slice(0, 3).forEach(item => {
      console.log(`  - ${item.field}: ${item.value}`);
    });
  }
});

console.log('\n=== 分析目标文件（1110(1)(1).xlsx）===');
const wb2 = XLSX.readFile('/tmp/1110(1)(1).xlsx');
const ws2 = wb2.Sheets['Sheet1'];
const data2 = XLSX.utils.sheet_to_json(ws2, { header: 1, defval: null, raw: true });

const header2 = data2[0];
console.log('Row 0 (表头):', header2.map((h, i) => {
  if (typeof h === 'number') {
    const dateStr = excelDateToString(h);
    const normalized = normalizeDate(dateStr);
    return `${h} → ${dateStr} → ${normalized}`;
  }
  const normalized = normalizeDate(h);
  return normalized !== h ? `${h} → ${normalized}` : h;
}));

// 提取目标文件的字段
const targetFields = [];
for (let i = 1; i < data2.length; i++) {
  const row = data2[i];
  if (row[1] && typeof row[1] === 'string') {
    targetFields.push(row[1]);
  }
}

console.log('\n目标文件字段:', targetFields);

console.log('\n=== 字段匹配检查 ===');
targetFields.forEach(targetField => {
  const matched = Array.from(sourceFields).some(sourceField =>
    fieldMatches(targetField, sourceField)
  );
  console.log(`${targetField}: ${matched ? '✓ 匹配' : '✗ 不匹配'}`);
  if (matched) {
    const matches = Array.from(sourceFields).filter(sourceField =>
      fieldMatches(targetField, sourceField)
    );
    console.log(`  -> 匹配到: ${matches.join(', ')}`);
  }
});
