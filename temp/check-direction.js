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

// 检测表格方向
function detectTableDirection(data) {
  let firstColumnDateCount = 0;
  for (const row of data) {
    if (row.length > 0 && row[0]) {
      if (normalizeDate(String(row[0])) !== String(row[0])) {
        firstColumnDateCount++;
      }
    }
  }

  let firstRowDateCount = 0;
  if (data.length > 0) {
    for (let i = 0; i < data[0].length; i++) {
      const header = data[0][i];
      if (header && normalizeDate(String(header)) !== String(header)) {
        firstRowDateCount++;
      }
    }
  }

  const firstColumnRatio = data.length > 0 ? firstColumnDateCount / data.length : 0;
  const firstRowRatio = data[0].length > 0 ? firstRowDateCount / data[0].length : 0;

  console.log(`第一列日期占比: ${firstColumnRatio}, 第一行日期占比: ${firstRowRatio}`);
  return firstColumnRatio >= firstRowRatio;
}

console.log('=== 检查目标文件（1110(1)(1).xlsx）的表格方向 ===');
const wb = XLSX.readFile('/tmp/1110(1)(1).xlsx');
const ws = wb.Sheets['Sheet1'];
const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });

console.log('总行数:', data.length);
console.log('Row 0:', JSON.stringify(data[0]));
console.log('Row 1:', JSON.stringify(data[1]));

const isVertical = detectTableDirection(data);
console.log('表格方向:', isVertical ? '纵向（第一列是时间）' : '横向（第一行是时间）');

if (!isVertical) {
  console.log('\n✓ 正确识别为横向表格');
  console.log('\n提取表头时间点:');
  const timePoints = [];
  for (let i = 0; i < data[0].length; i++) {
    const header = data[0][i];
    const normalized = normalizeDate(header);
    if (header && normalized !== String(header)) {
      timePoints.push({ col: i, original: header, normalized });
    }
  }
  console.log(JSON.stringify(timePoints, null, 2));

  console.log('\n前5行的字段名:');
  for (let i = 1; i < Math.min(6, data.length); i++) {
    const row = data[i];
    const field = row[1]; // 第二列是字段名
    console.log(`Row ${i}: ${row[0] || '(空)'} | ${field || '(空)'}`);
  }
} else {
  console.log('\n✗ 错误识别为纵向表格');
}
