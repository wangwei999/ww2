import { TableData, TableCell, MatchResult } from './types';
import {
  fieldMatches,
  calculateFieldSimilarity,
  normalizeDate,
  convertUnit,
  parseNumberCell,
  isCellEmpty,
} from './data-utils';

/**
 * 智能数据匹配器
 */
export class DataMatcher {
  private sourceTable: TableData;
  private targetTable: TableData;
  private sourceUnit?: string;
  private targetUnit?: string;
  private sourceHasPercentage?: boolean;
  private targetHasPercentage?: boolean;
  
  private sourceTimeColumnIndex: number = 0;
  private targetTimeColumnIndex: number = 0;
  
  private sourceDataMap: Map<string, Map<string, number>> = new Map();
  
  constructor(
    sourceTable: TableData,
    targetTable: TableData,
    sourceUnit?: string,
    targetUnit?: string,
    sourceHasPercentage?: boolean,
    targetHasPercentage?: boolean
  ) {
    this.sourceTable = sourceTable;
    this.targetTable = targetTable;
    this.sourceUnit = sourceUnit;
    this.targetUnit = targetUnit;
    this.sourceHasPercentage = sourceHasPercentage;
    this.targetHasPercentage = targetHasPercentage;
    
    this.initialize();
  }
  
  /**
   * 初始化匹配器
   */
  private initialize(): void {
    // 识别时间列
    this.sourceTimeColumnIndex = this.identifyTimeColumn(this.sourceTable);
    this.targetTimeColumnIndex = this.identifyTimeColumn(this.targetTable);
    
    // 构建源数据索引
    this.buildSourceIndex();
  }
  
  /**
   * 识别时间列
   */
  private identifyTimeColumn(table: TableData): number {
    let bestScore = -1;
    let bestColumn = 0;
    
    // 检查前5列
    for (let col = 0; col < Math.min(5, table.headers.length); col++) {
      let dateCount = 0;
      let totalRows = table.rows.length;
      
      for (const row of table.rows) {
        const cell = row[col];
        if (cell && normalizeDate(String(cell)) !== String(cell)) {
          dateCount++;
        }
      }
      
      const score = totalRows > 0 ? dateCount / totalRows : 0;
      if (score > bestScore) {
        bestScore = score;
        bestColumn = col;
      }
    }
    
    return bestColumn;
  }
  
  /**
   * 构建源数据索引
   */
  private buildSourceIndex(): void {
    for (const row of this.sourceTable.rows) {
      const timeCell = row[this.sourceTimeColumnIndex];
      if (!timeCell) continue;
      
      const normalizedTime = normalizeDate(String(timeCell));
      
      if (!this.sourceDataMap.has(normalizedTime)) {
        this.sourceDataMap.set(normalizedTime, new Map());
      }
      
      const timeMap = this.sourceDataMap.get(normalizedTime)!;
      
      // 遍历所有列
      for (let col = 0; col < this.sourceTable.headers.length; col++) {
        if (col === this.sourceTimeColumnIndex) continue;
        
        const header = this.sourceTable.headers[col];
        const value = parseNumberCell(row[col]);
        
        if (header && value !== null) {
          timeMap.set(header.trim(), value);
        }
      }
    }
  }
  
  /**
   * 匹配并填充数据
   */
  public matchAndFill(): {
    filledTable: TableData;
    matchResults: MatchResult[];
  } {
    const matchResults: MatchResult[] = [];
    
    const filledRows = this.targetTable.rows.map((row, rowIndex) => {
      const timeCell = row[this.targetTimeColumnIndex];
      if (!timeCell) return row;
      
      const normalizedTime = normalizeDate(String(timeCell));
      const sourceTimeMap = this.sourceDataMap.get(normalizedTime);
      
      if (!sourceTimeMap) return row;
      
      const newRow = [...row];
      
      // 遍历目标表的所有列
      for (let col = 0; col < this.targetTable.headers.length; col++) {
        if (col === this.targetTimeColumnIndex) continue;
        
        const targetCell = row[col];
        
        // 只填充空单元格
        if (isCellEmpty(targetCell)) {
          const targetHeader = this.targetTable.headers[col];
          
          // 查找匹配的源数据
          const matchedValue = this.findMatchedValue(targetHeader, sourceTimeMap);
          
          if (matchedValue !== null) {
            const convertedValue = this.convertValue(matchedValue);
            newRow[col] = convertedValue;
            
            matchResults.push({
              rowIndex,
              colIndex: col,
              value: convertedValue,
              converted: this.needsConversion(),
              unitFrom: this.sourceUnit,
              unitTo: this.targetUnit,
            });
          }
        }
      }
      
      return newRow;
    });
    
    return {
      filledTable: {
        ...this.targetTable,
        rows: filledRows,
      },
      matchResults,
    };
  }
  
  /**
   * 查找匹配的值
   */
  private findMatchedValue(
    targetHeader: string,
    sourceTimeMap: Map<string, number>
  ): number | null {
    // 1. 精确匹配
    if (sourceTimeMap.has(targetHeader)) {
      return sourceTimeMap.get(targetHeader)!;
    }
    
    // 2. 同义词匹配
    for (const [sourceHeader, value] of sourceTimeMap.entries()) {
      if (fieldMatches(targetHeader, sourceHeader)) {
        return value;
      }
    }
    
    // 3. 模糊匹配（相似度 > 0.7）
    let bestMatch: string | null = null;
    let bestSimilarity = 0.7;
    
    for (const [sourceHeader, value] of sourceTimeMap.entries()) {
      const similarity = calculateFieldSimilarity(targetHeader, sourceHeader);
      if (similarity > bestSimilarity) {
        bestSimilarity = similarity;
        bestMatch = sourceHeader;
      }
    }
    
    if (bestMatch) {
      return sourceTimeMap.get(bestMatch)!;
    }
    
    return null;
  }
  
  /**
   * 转换值（单位换算）
   */
  private convertValue(value: number): number {
    if (!this.needsConversion()) {
      return value;
    }
    
    return convertUnit(
      value,
      this.sourceUnit || '元',
      this.targetUnit || '元',
      this.sourceHasPercentage,
      this.targetHasPercentage
    );
  }
  
  /**
   * 是否需要转换
   */
  private needsConversion(): boolean {
    return (
      (this.sourceUnit && this.targetUnit && this.sourceUnit !== this.targetUnit) ||
      (this.sourceHasPercentage !== undefined &&
        this.targetHasPercentage !== undefined &&
        this.sourceHasPercentage !== this.targetHasPercentage)
    );
  }
  
  /**
   * 获取匹配统计信息
   */
  public getMatchStatistics(matchResults: MatchResult[]): {
    totalFilled: number;
    convertedCount: number;
    fillRate: number;
  } {
    const totalCells = this.targetTable.rows.length * (this.targetTable.headers.length - 1);
    const totalFilled = matchResults.length;
    const convertedCount = matchResults.filter(r => r.converted).length;
    const fillRate = totalCells > 0 ? (totalFilled / totalCells) * 100 : 0;
    
    return {
      totalFilled,
      convertedCount,
      fillRate,
    };
  }
}

/**
 * 批量匹配多个表格
 */
export class BatchDataMatcher {
  private sourceTables: TableData[];
  private targetTables: TableData[];
  private sourceUnit?: string;
  private targetUnit?: string;
  private sourceHasPercentage?: boolean;
  private targetHasPercentage?: boolean;
  
  constructor(
    sourceTables: TableData[],
    targetTables: TableData[],
    sourceUnit?: string,
    targetUnit?: string,
    sourceHasPercentage?: boolean,
    targetHasPercentage?: boolean
  ) {
    this.sourceTables = sourceTables;
    this.targetTables = targetTables;
    this.sourceUnit = sourceUnit;
    this.targetUnit = targetUnit;
    this.sourceHasPercentage = sourceHasPercentage;
    this.targetHasPercentage = targetHasPercentage;
  }
  
  /**
   * 批量匹配
   */
  public matchAll(): {
    results: Array<{
      filledTable: TableData;
      matchResults: MatchResult[];
      statistics: {
        totalFilled: number;
        convertedCount: number;
        fillRate: number;
      };
    }>;
  } {
    const results = [];
    
    // 为每个目标表格匹配源数据
    for (const targetTable of this.targetTables) {
      // 选择最佳匹配的源表格
      const bestSourceTable = this.findBestSourceTable(targetTable);
      
      if (bestSourceTable) {
        const matcher = new DataMatcher(
          bestSourceTable,
          targetTable,
          this.sourceUnit,
          this.targetUnit,
          this.sourceHasPercentage,
          this.targetHasPercentage
        );
        
        const { filledTable, matchResults } = matcher.matchAndFill();
        const statistics = matcher.getMatchStatistics(matchResults);
        
        results.push({
          filledTable,
          matchResults,
          statistics,
        });
      } else {
        results.push({
          filledTable: targetTable,
          matchResults: [],
          statistics: {
            totalFilled: 0,
            convertedCount: 0,
            fillRate: 0,
          },
        });
      }
    }
    
    return { results };
  }
  
  /**
   * 查找最佳匹配的源表格
   */
  private findBestSourceTable(targetTable: TableData): TableData | null {
    if (this.sourceTables.length === 0) return null;
    if (this.sourceTables.length === 1) return this.sourceTables[0];
    
    let bestScore = -1;
    let bestTable: TableData | null = null;
    
    for (const sourceTable of this.sourceTables) {
      const score = this.calculateTableSimilarity(sourceTable, targetTable);
      if (score > bestScore) {
        bestScore = score;
        bestTable = sourceTable;
      }
    }
    
    return bestTable;
  }
  
  /**
   * 计算表格相似度
   */
  private calculateTableSimilarity(source: TableData, target: TableData): number {
    let score = 0;
    
    // 比较表头相似度
    const sourceHeaders = new Set(source.headers.map(h => h.toLowerCase()));
    const targetHeaders = target.headers.map(h => h.toLowerCase());
    
    let headerMatches = 0;
    for (const header of targetHeaders) {
      if (sourceHeaders.has(header)) {
        headerMatches++;
      }
    }
    
    if (target.headers.length > 0) {
      score += (headerMatches / target.headers.length) * 50;
    }
    
    // 比较行数相似度
    const rowDiff = Math.abs(source.rows.length - target.rows.length);
    const maxRows = Math.max(source.rows.length, target.rows.length);
    if (maxRows > 0) {
      score += (1 - rowDiff / maxRows) * 50;
    }
    
    return score;
  }
}
