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
 * 优化：更精确的判断逻辑，避免将普通数值误判为日期
 */
export function isExcelDate(value: any): boolean {
  if (typeof value !== 'number' || isNaN(value)) {
    return false;
  }
  
  // Excel日期序列号通常在合理范围内
  // 1 = 1900-01-01, 55192 = 2050-12-31
  if (value < 1 || value > 60000) {
    return false;
  }
  
  // 检查是否是常见的财务数值（通常包含小数且在特定范围内）
  // 财务数据如 321.42、44955 等通常是小数
  // Excel 日期序列号通常是整数，小数部分代表时间
  // 如果是小数且小数部分很精确，很可能不是日期
  if (!Number.isInteger(value)) {
    // 如果小数部分包含多位小数，很可能是财务数据而非日期
    const decimalPlaces = value.toString().split('.')[1]?.length || 0;
    if (decimalPlaces >= 2) {
      return false;
    }
    
    // 如果小数部分不是简单的0.5或0.25等时间单位，很可能是财务数据
    const decimalPart = value % 1;
    const timeUnits = [0, 0.25, 0.5, 0.75]; // 常见的时间单位
    if (!timeUnits.some(unit => Math.abs(decimalPart - unit) < 0.01)) {
      return false;
    }
  }
  
  // 检查是否是常见的特殊数值
  // 例如：整数 1900、2000 等很可能是年份而非日期序列号
  if (Number.isInteger(value)) {
    // 1900-2100 之间的整数很可能是年份
    if (value >= 1900 && value <= 2100) {
      return false;
    }
    
    // 如果是常见的小整数（<100），很可能是计数或索引
    if (value < 100) {
      return false;
    }
  }
  
  // 通过所有检查，才认为是日期序列号
  return true;
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
