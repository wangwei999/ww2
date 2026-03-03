/**
 * 将Excel日期序列号转换为日期字符串
 */
export function excelDateToString(excelDate: number): string | null {
  if (typeof excelDate !== 'number' || isNaN(excelDate)) {
    return null;
  }
  
  // Excel日期基准：1900-01-01 = 1
  const excelEpoch = new Date(1900, 0, 1);
  const daysToAdd = excelDate - 2; // Excel有一个1900年2月29日的bug
  
  const date = new Date(excelEpoch.getTime() + daysToAdd * 24 * 60 * 60 * 1000);
  
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  
  return `${year}-${month}-${day}`;
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
export function convertExcelValue(value: any): any {
  if (isExcelDate(value)) {
    return excelDateToString(value);
  }
  return value;
}
