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
import { convertExcelValue } from './excel-date-utils';

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
    
    // 从文件内容中提取单位信息，增加采样范围以检测文件后面的标记
    const textSample = buffer.toString('utf-8', 0, Math.min(15000, buffer.length));
    const unitInfo = extractUnitInfo(textSample);
    
    // 检测是否包含"单位：万元 %"字样，如果是，则保持原始格式（不进行单位转换和百分比格式化）
    // 支持多种格式："单位：万元 %"、"单位: 万元 %"、"单位：万元%"、"单位: 万元%"
    let keepOriginalFormat = textSample.includes('单位：万元 %') || 
                           textSample.includes('单位: 万元 %') ||
                           textSample.includes('单位：万元%') ||
                           textSample.includes('单位: 万元%');
    
    console.log('=== 文件解析检测 ===');
    console.log('文件名:', file.name);
    console.log('文本样本长度:', textSample.length);
    console.log('包含"单位：万元 %"?', textSample.includes('单位：万元 %'));
    console.log('包含"单位: 万元 %"?', textSample.includes('单位: 万元 %'));
    console.log('keepOriginalFormat:', keepOriginalFormat);
    
    let tables: TableData[] = [];
    
    switch (ext) {
      case '.xlsx':
      case '.xls':
      case '.et':  // WPS 表格格式，使用 Excel 解析
        tables = this.parseExcel(buffer);
        break;
      case '.csv':
        tables = this.parseCSV(buffer);
        break;
      case '.docx':
      case '.wps':  // WPS 文档格式，尝试使用 Word 解析
        tables = await this.parseWord(buffer);
        break;
      case '.txt':
        tables = this.parseText(buffer);
        break;
      default:
        throw new Error(`不支持的文件格式: ${ext}。支持的格式包括：.xlsx, .xls, .et, .docx, .wps, .csv, .txt。如果是 .wps 文件，请使用 WPS Office 打开后另存为 .docx 格式再试。`);
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
      keepOriginalFormat,
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
        defval: null,
        raw: true,  // 使用原始值（日期序列号），不自动转换
      });
      
      if (rawData.length > 0) {
        // 转换Excel日期序列号
        const convertedData = rawData.map(row => 
          row.map(cell => convertExcelValue(cell))
        );
        
        const tablesInSheet = this.extractTablesFromRawData(convertedData);
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
    let hasFoundDataAfterHeader = false; // 是否已经找到表头后的数据行
    
    for (const row of rawData) {
      const hasData = row.some(cell => !isCellEmpty(cell) && cell !== null);
      
      if (hasData) {
        consecutiveEmptyRows = 0;
        
        // 如果已经找到了表头后的数据行，继续添加行
        // 否则，如果当前表为空，开始新表格
        if (hasFoundDataAfterHeader || currentTable.length === 0) {
          currentTable.push(row);
          if (currentTable.length > 1) {
            hasFoundDataAfterHeader = true;
          }
        } else {
          // 当前表只有表头，这是第一条数据行，添加到当前表
          currentTable.push(row);
          hasFoundDataAfterHeader = true;
        }
      } else if (currentTable.length > 0) {
        consecutiveEmptyRows++;
        currentTable.push(row);
        
        // 如果连续空行过多，且已经找到数据行，结束当前表格
        if (consecutiveEmptyRows >= maxEmptyRows && hasFoundDataAfterHeader) {
          if (currentTable.length > 0) {
            tables.push(this.createTableFromRawData(currentTable));
          }
          currentTable = [];
          consecutiveEmptyRows = 0;
          hasFoundDataAfterHeader = false;
        }
      }
    }
    
    // 添加最后一个表格（即使没有找到数据行，也要保存）
    if (currentTable.length > 0) {
      // 如果只有表头没有数据行，也要保存
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
      
      // 检查是否是单位行（如"单位：万元 %"）
      // 只有当整行只有单位信息时才过滤，如果一行中包含其他数据（如时间点），则保留
      const nonEmptyCells = row.filter(cell => !isCellEmpty(cell));
      const rowText = nonEmptyCells.join(' ');
      
      // 如果这一行只有单位信息（没有其他数据），则过滤掉
      if (nonEmptyCells.length === 1 && /单位[:：]/.test(rowText)) {
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
    
    // 检测并跳过标题行
    // 如果第一行只有1列，而第二行有更多列，第一行可能是标题行
    let headerRowIndex = 0;
    if (cleanedData.length > 1) {
      const firstRowColCount = cleanedData[0].filter(c => !isCellEmpty(c)).length;
      const secondRowColCount = cleanedData[1].filter(c => !isCellEmpty(c)).length;
      
      console.log('createTableFromRawData - 第一行列数:', firstRowColCount);
      console.log('createTableFromRawData - 第二行列数:', secondRowColCount);
      
      if (firstRowColCount === 1 && secondRowColCount > 2) {
        console.log('createTableFromRawData - 检测到标题行，跳过第一行');
        headerRowIndex = 1;
      }
    }
    
    // 使用检测到的表头行
    let headers = cleanedData[headerRowIndex].map(h => String(h || ''));
    
    // 找出单位列的索引（表头中包含"单位"关键字的列）
    const unitColIndices: number[] = [];
    headers.forEach((h, idx) => {
      if (/单位[:：]/.test(h)) {
        unitColIndices.push(idx);
      }
    });
    
    // 过滤掉单位列
    const filteredHeaders = headers.filter((_, idx) => !unitColIndices.includes(idx));
    
    const rows = cleanedData.slice(headerRowIndex + 1);
    
    // 确保每行的列数与过滤后的表头一致
    const filteredRows = rows.map(row => {
      // 跳过单位列，只保留数据列
      return row.filter((_, idx) => !unitColIndices.includes(idx));
    });
    
    console.log('createTableFromRawData - 原始数据行数:', data.length);
    console.log('createTableFromRawData - 过滤后行数:', cleanedData.length);
    console.log('createTableFromRawData - 表头行索引:', headerRowIndex);
    console.log('createTableFromRawData - 原始表头列数:', headers.length);
    console.log('createTableFromRawData - 过滤后表头列数:', filteredHeaders.length);
    console.log('createTableFromRawData - 单位列索引:', unitColIndices);
    console.log('createTableFromRawData - 数据行数:', filteredRows.length);
    console.log('createTableFromRawData - 表头内容:', filteredHeaders.slice(0, 10));
    
    return {
      headers: filteredHeaders,
      rows: filteredRows,
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
