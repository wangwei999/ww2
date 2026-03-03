// 测试日期格式识别
const DATE_PATTERNS = [
  { pattern: /(\d{4})[\/\-\.](\d{1,2})/, format: 'YM' },
  { pattern: /(\d{4})年(\d{1,2})月/, format: 'YM_CN' },
  { pattern: /(\d{4})年底/, format: 'YEAR_END' },
  { pattern: /(\d{4})年(\d{1,2})月底/, format: 'MONTH_END_CN' },
  { pattern: /(\d{4})年末/, format: 'YEAR_END' },
];

function normalizeDate(dateStr) {
  if (!dateStr) return '';

  const trimmed = dateStr.trim();

  for (const { pattern, format } of DATE_PATTERNS) {
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

const testCases = [
  '2024.3',
  '2024.6',
  '2024.9',
  '2024.12',
  '2018年底',
  '2020年6月底',
  '2020年底',
];

console.log('日期格式识别测试:');
testCases.forEach(test => {
  const result = normalizeDate(test);
  const changed = result !== test;
  console.log(`${test.padEnd(20)} -> ${result.padEnd(10)} ${changed ? '✓' : '✗'}`);
});
