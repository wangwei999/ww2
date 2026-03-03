// 表格单元格数据类型
export type TableCell = string | number | null | undefined;

// 表格数据结构
export interface TableData {
  headers: string[];
  rows: TableCell[][];
  unit?: string;
  hasPercentage?: boolean;
}

// 文件解析结果
export interface ParseResult {
  tables: TableData[];
  unit?: string;
  filename: string;
}

// 字段匹配配置
export interface FieldMatchConfig {
  fieldA: string;
  fieldB: string;
  matchScore: number;
}

// 日期标准化结果
export interface NormalizedDate {
  original: string;
  normalized: string;
  year: number;
  month: number;
}

// 单位信息
export interface UnitInfo {
  unit: string;
  value: number;
  hasPercentage: boolean;
}

// 数据匹配结果
export interface MatchResult {
  rowIndex: number;
  colIndex: number;
  value: number;
  converted?: boolean;
  unitFrom?: string;
  unitTo?: string;
}
