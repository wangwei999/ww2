/**
 * 将Excel日期序列号转换为日期字符串
 * 支持多种格式：
 * - YYYY-MM-DD: 2025-09-30
 * - YYYY/M/D: 2025/9/30 (用户期望格式)
 * @param excelDate Excel日期序列号
 * @param format 输出格式
 * @param adjustToMonthEnd 是否自动调整为月底日期（默认false）
 */
export function excelDateToString(
  excelDate: number,
  format: 'YYYY-MM-DD' | 'YYYY/M/D' = 'YYYY-MM-DD',
  adjustToMonthEnd: boolean = false
): string | null {
  if (typeof excelDate !== 'number' || isNaN(excelDate)) {
    return null;
  }
  
  // Excel日期基准：1900-01-01 = 1
  const excelEpoch = new Date(1900, 0, 1);
  const daysToAdd = excelDate - 2; // Excel有一个1900年2月29日的bug
  
  const date = new Date(excelEpoch.getTime() + daysToAdd * 24 * 60 * 60 * 1000);
  
  let year = date.getFullYear();
  let month = date.getMonth() + 1;
  let day = date.getDate();
  
  // 如果需要调整为月底日期
  if (adjustToMonthEnd) {
    const lastDay = new Date(year, month, 0).getDate(); // 获取该月的最后一天
    day = lastDay;
  }
  
  if (format === 'YYYY/M/D') {
    return `${year}/${month}/${day}`;
  } else {
    return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }
}

/**
 * 判断一个值是否是Excel日期序列号
 */
export function isExcelDate(value: any): boolean {
  if (typeof value !== 'number' || isNaN(value)) {
    return false;
  }
  
  // Excel日期序列号通常在1-100000之间
  return value > 0 && value < 100000;
}

/**
 * 转换单元格值，如果是Excel日期序列号则转换为日期字符串
 */
export function convertExcelValue(
  value: any,
  format: 'YYYY-MM-DD' | 'YYYY/M/D' = 'YYYY-MM-DD',
  adjustToMonthEnd: boolean = false
): any {
  if (isExcelDate(value)) {
    return excelDateToString(value, format, adjustToMonthEnd);
  }
  return value;
}
