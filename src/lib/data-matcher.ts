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
  
  // 表格方向：true=纵向（第一列是时间），false=横向（第一行是时间）
  private isVertical: boolean = true;
  
  // 如果是横向表格，存储第一行的时间点映射（列索引 -> 时间点）
  private timePointsInRow: Map<number, string> = new Map();
  
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
    console.log('DataMatcher - 开始初始化...');
    // 检测表格方向
    this.isVertical = this.detectTableDirection();
    console.log('DataMatcher - 表格方向:', this.isVertical ? '纵向' : '横向');
    
    if (this.isVertical) {
      // 纵向表格：第一列是时间
      this.sourceTimeColumnIndex = this.identifyTimeColumn(this.sourceTable);
      this.targetTimeColumnIndex = this.identifyTimeColumn(this.targetTable);
    } else {
      // 横向表格：第一行是时间
      console.log('DataMatcher - 提取目标表格时间点...');
      this.timePointsInRow = this.extractTimePointsFromRow(this.targetTable.headers);
      console.log('检测到横向表格，时间点:', Array.from(this.timePointsInRow.entries()));
    }
    
    // 构建源数据索引
    console.log('DataMatcher - 构建源数据索引...');
    this.buildSourceIndex();
  }
  
  /**
   * 检测表格方向
   */
  private detectTableDirection(): boolean {
    // 检查第一列是否包含时间
    let firstColumnDateCount = 0;
    for (const row of this.targetTable.rows) {
      const cell = row[0];
      if (cell && normalizeDate(String(cell)) !== String(cell)) {
        firstColumnDateCount++;
      }
    }
    
    // 检查第一行是否包含时间
    let firstRowDateCount = 0;
    for (let i = 0; i < this.targetTable.headers.length; i++) {
      const header = this.targetTable.headers[i];
      if (header && normalizeDate(String(header)) !== String(header)) {
        firstRowDateCount++;
      }
    }
    
    // 如果第一行的日期占比更高，认为是横向表格
    const firstColumnRatio = this.targetTable.rows.length > 0 
      ? firstColumnDateCount / this.targetTable.rows.length 
      : 0;
    const firstRowRatio = this.targetTable.headers.length > 0 
      ? firstRowDateCount / this.targetTable.headers.length 
      : 0;
    
    console.log(`第一列日期占比: ${firstColumnRatio}, 第一行日期占比: ${firstRowRatio}`);
    return firstColumnRatio >= firstRowRatio;
  }
  
  /**
   * 从表头中提取时间点
   */
  private extractTimePointsFromRow(headers: string[]): Map<number, string> {
    const timePoints = new Map<number, string>();
    
    console.log('开始提取时间点，表头列数:', headers.length);
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      const normalized = normalizeDate(String(header));
      if (header && normalized !== String(header)) {
        timePoints.set(i, normalized);
      }
    }
    
    console.log('提取到的时间点数量:', timePoints.size);
    return timePoints;
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
    if (this.isVertical) {
      // 纵向表格：第一列是时间，其他列是字段
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
    } else {
      // 横向表格：第一行是时间，第一列是分类，第二列是字段名
      console.log('构建横向表格的源数据索引...');
      console.log('源表表头列数:', this.sourceTable.headers.length);
      console.log('源表数据行数:', this.sourceTable.rows.length);
      
      // 从表头提取时间点
      const sourceTimePoints = this.extractTimePointsFromRow(this.sourceTable.headers);
      console.log('源表时间点数量:', sourceTimePoints.size);
      
      // 遍历每一行（每行代表一个字段）
      for (let rowIndex = 0; rowIndex < this.sourceTable.rows.length; rowIndex++) {
        const row = this.sourceTable.rows[rowIndex];
        
        // 第一列是分类名，第二列是字段名
        let fieldName: string | null = null;
        
        // 如果第二列有值，使用第二列
        if (row[1] && typeof row[1] === 'string') {
          fieldName = row[1];
        }
        // 否则，使用第一列
        else if (row[0] && typeof row[0] === 'string') {
          fieldName = row[0];
        }
        
        if (!fieldName) continue;
        
        console.log(`处理源表第${rowIndex}行，字段名: ${fieldName}`);
        
        // 遍历每一列（每列代表一个时间点）
        for (const [colIndex, normalizedTime] of sourceTimePoints.entries()) {
          const value = parseNumberCell(row[colIndex]);
          console.log(`  列${colIndex} 时间${normalizedTime} 值${value}`);
          
          if (value !== null) {
            if (!this.sourceDataMap.has(normalizedTime)) {
              this.sourceDataMap.set(normalizedTime, new Map());
            }
            
            const timeMap = this.sourceDataMap.get(normalizedTime)!;
            timeMap.set(fieldName.trim(), value);
          }
        }
      }
      
      console.log('源数据索引构建完成，时间点数量:', this.sourceDataMap.size);
      console.error('[DEBUG] 源数据索引内容:');
      for (const [time, fieldMap] of this.sourceDataMap.entries()) {
        console.error(`[DEBUG]   时间 ${time}: ${fieldMap.size} 个字段`);
        if (fieldMap.size > 0) {
          console.error(`[DEBUG]     字段示例: ${Array.from(fieldMap.keys()).slice(0, 3).join(', ')}`);
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
    
    if (this.isVertical) {
      // 纵向表格填充逻辑
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
    } else {
      // 横向表格填充逻辑
      console.log('开始横向表格数据匹配...');
      console.log('目标表格时间点:', Array.from(this.timePointsInRow.entries()));
      
      const filledRows = this.targetTable.rows.map((row, rowIndex) => {
        const newRow = [...row];
        
        // 第一列是字段名（或者分类名），第二列是真正的字段名
        let fieldName: string | null = null;
        
        // 如果第一列有值，使用第一列
        if (row[0] && typeof row[0] === 'string') {
          fieldName = row[0];
        }
        // 否则，使用第二列
        else if (row[1] && typeof row[1] === 'string') {
          fieldName = row[1];
        }
        
        if (!fieldName) return row;
        
        console.log(`处理第 ${rowIndex + 1} 行，字段名: ${fieldName}`);
        
        // 遍历每一列（每列代表一个时间点）
        for (const [colIndex, normalizedTime] of this.timePointsInRow.entries()) {
          const targetCell = row[colIndex];
          
          // 只填充空单元格，跳过前两列（分类名和字段名）
          if (colIndex < 2 || !isCellEmpty(targetCell)) continue;
          
          const sourceTimeMap = this.sourceDataMap.get(normalizedTime);
          if (!sourceTimeMap) continue;
          
          // 查找匹配的源数据
          const matchedValue = this.findMatchedValue(fieldName, sourceTimeMap);
          
          if (matchedValue !== null) {
            const convertedValue = this.convertValue(matchedValue);
            newRow[colIndex] = convertedValue;
            
            matchResults.push({
              rowIndex,
              colIndex: colIndex,
              value: convertedValue,
              converted: this.needsConversion(),
              unitFrom: this.sourceUnit,
              unitTo: this.targetUnit,
            });
            
            console.log(`  填充: 列${colIndex} 时间${normalizedTime} 值${matchedValue} → ${convertedValue}`);
          }
        }
        
        return newRow;
      });
      
      console.log('横向表格匹配完成，填充数量:', matchResults.length);
      
      return {
        filledTable: {
          ...this.targetTable,
          rows: filledRows,
        },
        matchResults,
      };
    }
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
