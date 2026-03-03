// 测试修改后的代码
function isExcelDate(value) {
  if (typeof value !== 'number' || isNaN(value)) {
    return false;
  }

  // 1 = 1900-01-01, 60000 约等于 2064年
  if (value < 1 || value > 60000) {
    return false;
  }

  // 排除年份和常见小整数
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

function normalizeDate(dateStr) {
  const patterns = [
    { pattern: /(\d{4})[\/\-\.](\d{1,2})/, format: 'YM' },
    { pattern: /(\d{4})年(\d{1,2})月/, format: 'YM_CN' },
    { pattern: /(\d{4})年底/, format: 'YEAR_END' },
    { pattern: /(\d{4})年(\d{1,2})月底/, format: 'MONTH_END_CN' },
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

// 模拟 extractTimePointsFromRow 函数
function extractTimePointsFromRow(headers) {
  const timePoints = new Map();

  console.log('开始提取时间点，表头列数:', headers.length);
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    let normalized = null;

    // 1. 检查是否是 Excel 日期序列号
    if (typeof header === 'number' && isExcelDate(header)) {
      normalized = excelDateToString(header);
      console.log(`列${i} 检测到 Excel 日期序列号 ${header} -> ${normalized}`);
    } else if (header) {
      // 2. 检查是否是文本日期
      normalized = normalizeDate(String(header));
      if (normalized !== String(header)) {
        console.log(`列${i} 检测到文本日期 ${header} -> ${normalized}`);
      }
    }

    if (normalized && header && normalized !== String(header)) {
      timePoints.set(i, normalized);
    }
  }

  console.log('提取到的时间点数量:', timePoints.size);
  console.log('时间点详情:', Array.from(timePoints.entries()));
  return timePoints;
}

const testHeaders = [
  null,
  null,
  "2018年底",
  "2019年底",
  "2020年6月底",
  "2020年底",
  "2021年6月底",
  44713,
  45107,
  45199,
  45473,
  45565,
  45838,
  45930
];

console.log('=== 测试时间点提取 ===');
extractTimePointsFromRow(testHeaders);
