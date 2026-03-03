import { TableCell, NormalizedDate, UnitInfo } from './types';

// 日期格式正则表达式（更全面的匹配）
const DATE_PATTERNS = [
  { pattern: /(\d{4})\.(\d{1,2})/, format: 'YMDOT' },                // 2025.9, 2024.3 (支持点号分隔)
  { pattern: /(\d{4})[\/\-](\d{1,2})/, format: 'YM' },                // 2025/9, 2025-9, 2025.9
  { pattern: /(\d{4})年(\d{1,2})月/, format: 'YM_CN' },               // 2025年9月
  { pattern: /(\d{4})年底/, format: 'YEAR_END' },                      // 2025年底
  { pattern: /(\d{4})年末/, format: 'YEAR_END' },                      // 2025年末
  { pattern: /(\d{4})年(\d{1,2})月底/, format: 'MONTH_END_CN' },       // 2025年9月底
  { pattern: /(\d{4})年(\d{1,2})月末/, format: 'MONTH_END_CN' },       // 2025年9月末
  { pattern: /(\d{4})年/, format: 'YEAR_ONLY' },                       // 2025年, 2021年
  { pattern: /(\d{4})(\d{2})/, format: 'YM_COMPACT' },                 // 202509
  { pattern: /(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/, format: 'YMD' }, // 2025/9/15
  { pattern: /(\d{4})年(\d{1,2})月(\d{1,2})日/, format: 'YMD_CN' },    // 2025年9月15日
  // 英文月份格式
  { pattern: /(\d{1,2})\/(\d{1,2})\/(\d{2,4})/, format: 'MDY' },        // 6/1/22, 6/1/2022
  { pattern: /([A-Za-z]{3})-(\d{2})/, format: 'MON_YY' },              // Jun-23
];

/**
 * 将日期调整为月底日期
 * 例如：2023-06-29 -> 2023-06-30（6月的最后一天）
 */
export function adjustToMonthEnd(dateStr: string): string {
  if (!dateStr || !/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    return dateStr;
  }
  
  const [year, month] = dateStr.split('-').map(Number);
  const lastDay = new Date(year, month, 0).getDate(); // 获取该月的最后一天
  
  return `${year}-${String(month).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
}

// 扩展的同义词映射
const SYNONYMS: Record<string, string[]> = {
  '总资产': ['资产总额', '资产合计', '总资', '资产总计', '总资产额'],
  '净资产': ['所有者权益', '股东权益', '净资产额', '权益合计'],
  '净利润': ['净利', '净利润额', '纯利润', '净收益'],
  '营业收入': ['营收', '营业收入额', '收入', '销售收入', '主营业务收入'],
  '营业成本': ['成本', '营业成本额', '主营业务成本'],
  '营业利润': ['利润', '营业利润额', '经营利润'],
  '利润总额': ['总利润', '利润总额额'],
  '资产负债率': ['负债率', '资产负债比率', '债务比率'],
  '流动资产': ['流动资产合计'],
  '流动负债': ['流动负债合计'],
  '固定资产': ['固定资产原值'],
  '应收账款': ['应收账款净额'],
  '存货': ['存货净额'],
  '现金': ['货币资金', '现金及现金等价物'],
  '营业收入增长率': ['营收增长率', '收入增长率'],
  '净利润增长率': ['净利增长率'],
  '总资产周转率': ['资产周转率'],
  '贷款总额': ['贷款额', '贷款金额', '贷款'],
  '客户存款': ['存款', '客户资金'],
  '不良贷款率': ['不良资产率', '不良资产占比'],
  '流动性覆盖率': ['流动性覆盖率!'], // 去除感叹号
  '净稳定资金比例': ['净稳定资金比例!'], // 去除感叹号
  '优质流动性资产充足率': ['优质流动性资产充足率!'], // 去除感叹号
  '一级资本充足率': ['一级资本充足率视同'], // 同义词匹配
  '存贷款比例': ['存贷款比例视同'], // 同义词匹配
  // 可以继续添加更多同义词
};

// 单位转换映射（数值越大单位越小）
const UNIT_VALUES: Record<string, number> = {
  '亿元': 100000000,
  '万元': 10000,
  '元': 1,
  '百元': 100,
  '千元': 1000,
};

/**
 * 规范化日期字符串
 */
export function normalizeDate(dateStr: string): string {
  if (!dateStr) return '';
  
  const trimmed = dateStr.trim();
  
  for (const { pattern, format } of DATE_PATTERNS) {
    const match = trimmed.match(pattern);
    if (match) {
      let year: string;
      let month: string;
      
      if (format === 'YM' || format === 'YM_CN' || format === 'YMD') {
        year = match[1];
        month = match[2].padStart(2, '0');
      } else if (format === 'YMDOT') {
        // 点号分隔格式：2024.3 -> 2024-03
        year = match[1];
        month = match[2].padStart(2, '0');
      } else if (format === 'YEAR_END') {
        year = match[1];
        month = '12';
      } else if (format === 'MONTH_END_CN') {
        year = match[1];
        month = match[2].padStart(2, '0');
      } else if (format === 'YEAR_ONLY') {
        // 只包含年份，如"2021年"
        year = match[1];
        month = '12'; // 默认年底
      } else if (format === 'YM_COMPACT') {
        year = match[1];
        month = match[2];
      } else if (format === 'YMD_CN') {
        year = match[1];
        month = match[2].padStart(2, '0');
      } else if (format === 'MDY') {
        // 月/日/年 或 月/日/年的变体
        const part1 = parseInt(match[1], 10);
        const part2 = parseInt(match[2], 10);
        const part3 = parseInt(match[3], 10);
        
        // 判断哪部分是年份（通常是最大的或者4位数的）
        if (part3 >= 100 || (part3 >= 50 && part3 <= 99)) {
          // 格式：月/日/年 (6/1/22)
          year = part3 < 100 ? (part3 >= 50 ? 1900 + part3 : 2000 + part3).toString() : part3.toString();
          month = part1.toString().padStart(2, '0');
        } else if (part1 >= 100 || (part1 >= 50 && part1 <= 99)) {
          // 格式：年/月/日 (2022/6/1)
          year = part1 < 100 ? (part1 >= 50 ? 1900 + part1 : 2000 + part1).toString() : part1.toString();
          month = part2.toString().padStart(2, '0');
        } else {
          // 默认按月/日/年处理
          year = part3 < 100 ? (part3 >= 50 ? 1900 + part3 : 2000 + part3).toString() : part3.toString();
          month = part1.toString().padStart(2, '0');
        }
      } else if (format === 'MON_YY') {
        // 英文月份-年份 (Jun-23)
        const monthNames: Record<string, string> = {
          'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04',
          'May': '05', 'Jun': '06', 'Jul': '07', 'Aug': '08',
          'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'
        };
        const monthAbbr = match[1].charAt(0).toUpperCase() + match[1].slice(1).toLowerCase();
        const yearNum = parseInt(match[2], 10);
        month = monthNames[monthAbbr] || '01';
        year = (yearNum < 100 ? (yearNum >= 50 ? 1900 + yearNum : 2000 + yearNum) : yearNum).toString();
      } else {
        year = match[1];
        month = match[2].padStart(2, '0');
      }
      
      return `${year}-${month}`;
    }
  }
  
  return trimmed;
}

/**
 * 解析日期并返回详细信息
 */
export function parseDate(dateStr: string): NormalizedDate | null {
  if (!dateStr) return null;
  
  const trimmed = dateStr.trim();
  
  for (const { pattern, format } of DATE_PATTERNS) {
    const match = trimmed.match(pattern);
    if (match) {
      let year: number;
      let month: number;
      
      if (format === 'YM' || format === 'YM_CN' || format === 'YMD') {
        year = parseInt(match[1], 10);
        month = parseInt(match[2], 10);
      } else if (format === 'YEAR_END') {
        year = parseInt(match[1], 10);
        month = 12;
      } else if (format === 'MONTH_END_CN') {
        year = parseInt(match[1], 10);
        month = parseInt(match[2], 10);
      } else if (format === 'YM_COMPACT') {
        year = parseInt(match[1], 10);
        month = parseInt(match[2], 10);
      } else if (format === 'YMD_CN') {
        year = parseInt(match[1], 10);
        month = parseInt(match[2], 10);
      } else if (format === 'MDY') {
        // 月/日/年 或 月/日/年的变体
        const part1 = parseInt(match[1], 10);
        const part2 = parseInt(match[2], 10);
        const part3 = parseInt(match[3], 10);
        
        // 判断哪部分是年份（通常是最大的或者4位数的）
        if (part3 >= 100 || (part3 >= 50 && part3 <= 99)) {
          year = part3 < 100 ? (part3 >= 50 ? 1900 + part3 : 2000 + part3) : part3;
          month = part1;
        } else if (part1 >= 100 || (part1 >= 50 && part1 <= 99)) {
          year = part1 < 100 ? (part1 >= 50 ? 1900 + part1 : 2000 + part1) : part1;
          month = part2;
        } else {
          year = part3 < 100 ? (part3 >= 50 ? 1900 + part3 : 2000 + part3) : part3;
          month = part1;
        }
      } else if (format === 'MON_YY') {
        // 英文月份-年份 (Jun-23)
        const monthNames: Record<string, number> = {
          'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4,
          'May': 5, 'Jun': 6, 'Jul': 7, 'Aug': 8,
          'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12
        };
        const monthAbbr = match[1].charAt(0).toUpperCase() + match[1].slice(1).toLowerCase();
        const yearNum = parseInt(match[2], 10);
        month = monthNames[monthAbbr] || 1;
        year = yearNum < 100 ? (yearNum >= 50 ? 1900 + yearNum : 2000 + yearNum) : yearNum;
      } else {
        year = parseInt(match[1], 10);
        month = parseInt(match[2], 10);
      }
      
      return {
        original: trimmed,
        normalized: `${year}-${String(month).padStart(2, '0')}`,
        year,
        month,
      };
    }
  }
  
  return null;
}

/**
 * 清理字段名，移除单位、括号等后缀
 */
export function cleanFieldName(fieldName: string): string {
  let cleaned = fieldName.trim();

  // 移除单位后缀（如"（亿元）"、"（%）"、"(%)"等）
  cleaned = cleaned.replace(/[（\(][^\)）]*[）\)]/g, '');

  // 移除末尾的单位（如"亿元"、"万元"、"元"等）
  cleaned = cleaned.replace(/(?:亿元|万元|千元|百元|元|元)%?$/g, '');

  // 移除末尾的百分号
  cleaned = cleaned.replace(/%$/g, '');

  return cleaned.trim();
}

/**
 * 检查两个字段是否匹配（支持同义词和前缀忽略）
 */
export function fieldMatches(fieldA: string, fieldB: string): boolean {
  if (!fieldA || !fieldB) return false;
  
  // 移除常见的前缀（如"其中："、"其中"）
  let normalizedA = fieldA.trim();
  let normalizedB = fieldB.trim();
  
  const prefixes = ['其中：', '其中', '包括：', '包括', '含：', '含'];
  for (const prefix of prefixes) {
    if (normalizedA.startsWith(prefix)) {
      normalizedA = normalizedA.substring(prefix.length).trim();
    }
    if (normalizedB.startsWith(prefix)) {
      normalizedB = normalizedB.substring(prefix.length).trim();
    }
  }
  
  // 清理字段名（移除单位后缀）- 使用移除了前缀后的名称
  const cleanedA = cleanFieldName(normalizedA).toLowerCase();
  const cleanedB = cleanFieldName(normalizedB).toLowerCase();
  
  // 完全匹配（使用清理后的名称）
  if (cleanedA === cleanedB) return true;
  
  // 移除包含关系的匹配，避免不同指标被错误匹配
  // 例如："资本充足率"、"一级资本充足率"、"核心资本充足率"是不同指标，不应匹配
  
  // 检查同义词（使用清理后的名称）
  for (const [key, synonyms] of Object.entries(SYNONYMS)) {
    const group = [key, ...synonyms].map(s => cleanFieldName(s).toLowerCase());
    if (group.includes(cleanedA) && group.includes(cleanedB)) {
      return true;
    }
  }
  
  return false;
}

/**
 * 计算字段相似度（用于模糊匹配）
 */
export function calculateFieldSimilarity(fieldA: string, fieldB: string): number {
  if (!fieldA || !fieldB) return 0;
  
  if (fieldMatches(fieldA, fieldB)) return 1;
  
  // 简单的相似度计算（基于编辑距离）
  const distance = levenshteinDistance(fieldA.toLowerCase(), fieldB.toLowerCase());
  const maxLen = Math.max(fieldA.length, fieldB.length);
  return 1 - distance / maxLen;
}

/**
 * 编辑距离算法
 */
function levenshteinDistance(str1: string, str2: string): number {
  const m = str1.length;
  const n = str2.length;
  const dp: number[][] = [];
  
  for (let i = 0; i <= m; i++) {
    dp[i] = [i];
  }
  
  for (let j = 0; j <= n; j++) {
    dp[0][j] = j;
  }
  
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      if (str1[i - 1] === str2[j - 1]) {
        dp[i][j] = dp[i - 1][j - 1];
      } else {
        dp[i][j] = Math.min(
          dp[i - 1][j] + 1,
          dp[i][j - 1] + 1,
          dp[i - 1][j - 1] + 1
        );
      }
    }
  }
  
  return dp[m][n];
}

/**
 * 从文本中提取单位信息
 */
export function extractUnitInfo(text: string): UnitInfo | null {
  // 匹配单位：单位：xxx 或 单位 xxx
  const unitPatterns = [
    /单位[:：]\s*([^%]+?)\s*(%)?/,
    /单位\s*[:：]\s*([^%]+?)\s*(%)?/,
  ];
  
  for (const pattern of unitPatterns) {
    const match = text.match(pattern);
    if (match) {
      const unit = match[1].trim();
      const hasPercentage = !!match[2];
      
      // 验证单位是否有效
      if (Object.keys(UNIT_VALUES).includes(unit)) {
        return {
          unit,
          value: UNIT_VALUES[unit],
          hasPercentage,
        };
      }
    }
  }
  
  return null;
}

/**
 * 转换单位
 */
export function convertUnit(
  value: number,
  fromUnit: string,
  toUnit: string,
  hasPercentageFrom?: boolean,
  hasPercentageTo?: boolean
): number {
  let result = value;
  
  // 单位换算
  const fromValue = UNIT_VALUES[fromUnit] || 1;
  const toValue = UNIT_VALUES[toUnit] || 1;
  result = (value * fromValue) / toValue;
  
  // 百分比换算
  if (hasPercentageFrom && !hasPercentageTo) {
    result = result / 100;
  } else if (!hasPercentageFrom && hasPercentageTo) {
    result = result * 100;
  }
  
  return result;
}

/**
 * 检查单元格是否为空
 */
export function isCellEmpty(cell: TableCell): boolean {
  if (cell === null || cell === undefined) return true;
  if (typeof cell === 'string' && cell.trim() === '') return true;
  if (cell === '') return true;
  return false;
}

/**
 * 解析数字单元格
 */
export function parseNumberCell(cell: TableCell): number | null {
  if (isCellEmpty(cell)) return null;
  
  const num = parseFloat(String(cell).replace(/[%，,]/g, ''));
  return isNaN(num) ? null : num;
}

/**
 * 格式化数字
 */
export function formatNumber(num: number, decimals: number = 2): string {
  return num.toFixed(decimals);
}

/**
 * 添加同义词映射
 */
export function addSynonym(field: string, synonym: string): void {
  if (!SYNONYMS[field]) {
    SYNONYMS[field] = [];
  }
  
  if (!SYNONYMS[field].includes(synonym)) {
    SYNONYMS[field].push(synonym);
  }
}

/**
 * 获取所有同义词组
 */
export function getAllSynonyms(): Record<string, string[]> {
  return { ...SYNONYMS };
}

/**
 * 添加单位换算规则
 */
export function addUnitRule(unit: string, value: number): void {
  UNIT_VALUES[unit] = value;
}

/**
 * 获取单位值
 */
export function getUnitValue(unit: string): number {
  return UNIT_VALUES[unit] || 1;
}
