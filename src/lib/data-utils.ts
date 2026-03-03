import { TableCell, NormalizedDate, UnitInfo } from './types';

// 日期格式正则表达式（更全面的匹配）
const DATE_PATTERNS = [
  { pattern: /(\d{4})[\/\-\.](\d{1,2})/, format: 'YM' },              // 2025/9, 2025-9, 2025.9
  { pattern: /(\d{4})年(\d{1,2})月/, format: 'YM_CN' },                 // 2025年9月
  { pattern: /(\d{4})(\d{2})/, format: 'YM_COMPACT' },                   // 202509
  { pattern: /(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/, format: 'YMD' }, // 2025/9/15
  { pattern: /(\d{4})年(\d{1,2})月(\d{1,2})日/, format: 'YMD_CN' },        // 2025年9月15日
];

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
  
  for (const { pattern } of DATE_PATTERNS) {
    const match = trimmed.match(pattern);
    if (match) {
      const year = match[1];
      const month = match[2].padStart(2, '0');
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
  
  for (const { pattern } of DATE_PATTERNS) {
    const match = trimmed.match(pattern);
    if (match) {
      const year = parseInt(match[1], 10);
      const month = parseInt(match[2], 10);
      
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
 * 检查两个字段是否匹配（支持同义词）
 */
export function fieldMatches(fieldA: string, fieldB: string): boolean {
  if (!fieldA || !fieldB) return false;
  
  const normalizedA = fieldA.trim().toLowerCase();
  const normalizedB = fieldB.trim().toLowerCase();
  
  // 完全匹配
  if (normalizedA === normalizedB) return true;
  
  // 包含关系
  if (normalizedA.includes(normalizedB) || normalizedB.includes(normalizedA)) {
    return true;
  }
  
  // 检查同义词
  for (const [key, synonyms] of Object.entries(SYNONYMS)) {
    const group = [key, ...synonyms].map(s => s.toLowerCase());
    if (group.includes(normalizedA) && group.includes(normalizedB)) {
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
