import * as XLSX from 'xlsx';
import mammoth from 'mammoth';
import Papa from 'papaparse';
import { TableCell, TableData, ParseResult } from './types';
import {
  extractUnitInfo,
  isCellEmpty,
  parseNumberCell,
  normalizeDate,
} from './data-utils';

/**
 * 文件解析器
 */
export class FileParser {
  /**
   * 解析文件
   */
  static async parseFile(file: File): Promise<ParseResult> {
    const bytes = await file.arrayBuffer();
    const buffer = Buffer.from(bytes);
    const ext = this.getFileExtension(file.name);
    
    // 从文件内容中提取单位信息
    const textSample = buffer.toString('utf-8', 0, Math.min(3000, buffer.length));
    const unitInfo = extractUnitInfo(textSample);
    
    let tables: TableData[] = [];
    
    switch (ext) {
      case '.xlsx':
      case '.xls':
        tables = this.parseExcel(buffer);
        break;
      case '.csv':
        tables = this.parseCSV(buffer);
        break;
      case '.docx':
        tables = await this.parseWord(buffer);
        break;
      case '.txt':
        tables = this.parseText(buffer);
        break;
      default:
        throw new Error(`不支持的文件格式: ${ext}`);
    }
    
    // 为每个表格添加单位信息
    tables = tables.map(table => ({
      ...table,
      unit: table.unit || unitInfo?.unit,
      hasPercentage: table.hasPercentage || unitInfo?.hasPercentage,
    }));
    
    return {
      tables,
      unit: unitInfo?.unit,
      filename: file.name,
    };
  }
  
  /**
   * 获取文件扩展名
   */
  private static getFileExtension(filename: string): string {
    return filename.toLowerCase().slice(filename.lastIndexOf('.'));
  }
  
  /**
   * 解析 Excel 文件
   */
  private static parseExcel(buffer: Buffer): TableData[] {
    const workbook = XLSX.read(buffer);
    const tables: TableData[] = [];
    
    // 遍历所有工作表
    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      const rawData = XLSX.utils.sheet_to_json<TableCell[]>(worksheet, { 
        header: 1,
        defval: null 
      });
      
      if (rawData.length > 0) {
        const tablesInSheet = this.extractTablesFromRawData(rawData);
        tables.push(...tablesInSheet);
      }
    }
    
    return tables;
  }
  
  /**
   * 解析 CSV 文件
   */
  private static parseCSV(buffer: Buffer): TableData[] {
    const text = buffer.toString('utf-8');
    const result = Papa.parse<TableCell[]>(text, {
      header: false,
      skipEmptyLines: false,
    });
    
    if (result.data.length > 0) {
      return this.extractTablesFromRawData(result.data);
    }
    
    return [];
  }
  
  /**
   * 解析 Word 文件
   */
  private static async parseWord(buffer: Buffer): Promise<TableData[]> {
    const result = await mammoth.extractRawText({ buffer });
    const text = result.value;
    
    // Word 文件中的表格需要特殊处理
    // 这里先简单实现，提取可能包含表格的段落
    const tables: TableData[] = [];
    const lines = text.split('\n').filter(line => line.trim());
    
    // 尝试检测表格行（包含制表符或多个空格分隔）
    const potentialTableRows = lines.filter(line => {
      return line.includes('\t') || line.split(/\s{2,}/).length > 2;
    });
    
    if (potentialTableRows.length > 0) {
      const tableData = potentialTableRows.map(line => {
        return line.split('\t').map(cell => cell.trim());
      });
      
      tables.push(this.createTableFromRawData(tableData));
    }
    
    return tables;
  }
  
  /**
   * 解析文本文件
   */
  private static parseText(buffer: Buffer): TableData[] {
    const text = buffer.toString('utf-8');
    const lines = text.split('\n').filter(line => line.trim());
    
    // 尝试检测表格结构
    const tables: TableData[] = [];
    let currentTable: TableCell[][] = [];
    let tableStartLine = -1;
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const cells = line.split(/[\t,;|]/).map(c => c.trim());
      
      // 判断是否是表格行（包含多个列）
      if (cells.length > 1 && cells.some(c => c !== '')) {
        if (tableStartLine === -1) {
          tableStartLine = i;
        }
        currentTable.push(cells);
      } else {
        // 非表格行，保存当前表格
        if (currentTable.length > 1) {
          tables.push(this.createTableFromRawData(currentTable));
        }
        currentTable = [];
        tableStartLine = -1;
      }
    }
    
    // 保存最后一个表格
    if (currentTable.length > 1) {
      tables.push(this.createTableFromRawData(currentTable));
    }
    
    return tables;
  }
  
  /**
   * 从原始数据中提取表格
   */
  private static extractTablesFromRawData(rawData: TableCell[][]): TableData[] {
    const tables: TableData[] = [];
    let currentTable: TableCell[][] = [];
    let consecutiveEmptyRows = 0;
    const maxEmptyRows = 2; // 连续空行数超过这个值则认为表格结束
    
    for (const row of rawData) {
      const hasData = row.some(cell => !isCellEmpty(cell) && cell !== null);
      
      if (hasData) {
        consecutiveEmptyRows = 0;
        currentTable.push(row);
      } else if (currentTable.length > 0) {
        consecutiveEmptyRows++;
        currentTable.push(row);
        
        // 如果连续空行过多，结束当前表格
        if (consecutiveEmptyRows >= maxEmptyRows) {
          if (currentTable.length > 0) {
            tables.push(this.createTableFromRawData(currentTable));
          }
          currentTable = [];
          consecutiveEmptyRows = 0;
        }
      }
    }
    
    // 添加最后一个表格
    if (currentTable.length > 0) {
      tables.push(this.createTableFromRawData(currentTable));
    }
    
    return tables;
  }
  
  /**
   * 从原始数据创建表格对象
   */
  private static createTableFromRawData(data: TableCell[][]): TableData {
    if (data.length === 0) {
      return {
        headers: [],
        rows: [],
      };
    }
    
    // 过滤掉单位行和非数据行
    const cleanedData = data.filter(row => {
      if (!row || row.length === 0) return false;
      
      // 检查是否是单位行（如"单位：万元"）
      const rowText = row.join(' ');
      if (/单位[:：]/.test(rowText)) {
        return false;
      }
      
      // 检查是否有实际数据（至少有一个非空单元格）
      return row.some(cell => !isCellEmpty(cell));
    });
    
    if (cleanedData.length === 0) {
      return {
        headers: [],
        rows: [],
      };
    }
    
    // 第一行作为表头
    const headers = cleanedData[0].map(h => String(h || ''));
    const rows = cleanedData.slice(1);
    
    return {
      headers,
      rows,
    };
  }
  
  /**
   * 智能识别表格结构
   */
  static identifyTableStructure(data: TableCell[][]): {
    hasHeader: boolean;
    headerRowIndex: number;
    timeColumnIndex: number;
  } {
    let bestScore = -1;
    let hasHeader = true;
    let headerRowIndex = 0;
    let timeColumnIndex = 0;
    
    // 尝试每一行作为表头
    for (let i = 0; i < Math.min(3, data.length); i++) {
      const row = data[i];
      if (!row || row.length === 0) continue;
      
      // 尝试每一列作为时间列
      for (let j = 0; j < Math.min(5, row.length); j++) {
        const score = this.scoreStructure(data, i, j);
        
        if (score > bestScore) {
          bestScore = score;
          headerRowIndex = i;
          timeColumnIndex = j;
          hasHeader = i === 0;
        }
      }
    }
    
    return {
      hasHeader,
      headerRowIndex,
      timeColumnIndex,
    };
  }
  
  /**
   * 评分表格结构
   */
  private static scoreStructure(
    data: TableCell[][],
    headerRowIndex: number,
    timeColumnIndex: number
  ): number {
    let score = 0;
    
    // 检查表头行是否有合理的字段名
    const headerRow = data[headerRowIndex];
    if (headerRow && headerRow.length > 1) {
      score += 5; // 有多列
    }
    
    // 检查时间列是否有日期
    let dateCount = 0;
    let totalRows = 0;
    
    for (let i = headerRowIndex + 1; i < data.length; i++) {
      const row = data[i];
      if (row && row.length > timeColumnIndex) {
        totalRows++;
        const cell = row[timeColumnIndex];
        if (cell && normalizeDate(String(cell)) !== String(cell)) {
          dateCount++;
        }
      }
    }
    
    if (totalRows > 0) {
      score += (dateCount / totalRows) * 10; // 日期占比
    }
    
    // 检查数据列是否有数值
    let numberCount = 0;
    let dataCells = 0;
    
    for (let i = headerRowIndex + 1; i < data.length; i++) {
      const row = data[i];
      if (row) {
        for (let j = 0; j < row.length; j++) {
          if (j !== timeColumnIndex) {
            dataCells++;
            if (parseNumberCell(row[j]) !== null) {
              numberCount++;
            }
          }
        }
      }
    }
    
    if (dataCells > 0) {
      score += (numberCount / dataCells) * 10; // 数值占比
    }
    
    return score;
  }
}
