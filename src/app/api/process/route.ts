import { NextRequest, NextResponse } from 'next/server';
import { writeFile, mkdir } from 'fs/promises';
import { existsSync, mkdirSync, writeFileSync } from 'fs';
import path from 'path';
import * as XLSX from 'xlsx';
import { FileParser } from '@/lib/file-parser';
import { BatchDataMatcher } from '@/lib/data-matcher';
import { excelDateToString, isExcelDate } from '@/lib/excel-date-utils';
import { adjustToMonthEnd } from '@/lib/data-utils';

// 临时文件存储目录
const TEMP_DIR = path.join(process.cwd(), 'temp');

// 确保临时目录存在
async function ensureTempDir() {
  if (!existsSync(TEMP_DIR)) {
    await mkdir(TEMP_DIR, { recursive: true });
  }
}

/**
 * 将表格数据转换为 Excel 格式
 */
function tableToExcel(table: any, matchResults: any[] = []): XLSX.WorkSheet {
  console.log('转换表格到 Excel 格式...');
  console.log('表格 headers:', table.headers);
  console.log('表格 rows 数量:', table.rows?.length);
  
  const { headers, rows } = table;
  
  // 验证数据
  if (!Array.isArray(headers)) {
    console.error('headers 不是数组:', headers);
    throw new Error('表格 headers 必须是数组');
  }
  
  if (!Array.isArray(rows)) {
    console.error('rows 不是数组:', rows);
    throw new Error('表格 rows 必须是数组');
  }
  
  // 创建百分比映射（用于快速查找）
  const percentageMap = new Map<string, boolean>();
  matchResults.forEach((mr: any) => {
    const key = `${mr.rowIndex}-${mr.colIndex}`;
    percentageMap.set(key, mr.isPercentage || false);
  });
  
  // 转换表头中的日期为用户期望的格式 (YYYY/M/D)，并自动调整为月底日期
  const convertedHeaders = headers.map((header: any) => {
    // 检查是否是日期序列号
    if (isExcelDate(header)) {
      const dateString = excelDateToString(header, 'YYYY/M/D', true); // 自动调整为月底
      console.log(`转换表头日期序列号: ${header} -> ${dateString}`);
      return dateString;
    }
    
    // 检查是否是 YYYY-MM-DD 格式的日期字符串
    if (typeof header === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(header)) {
      // 先调整为月底日期，再转换为 YYYY/M/D 格式
      const adjustedDate = adjustToMonthEnd(header);
      const [year, month, day] = adjustedDate.split('-');
      const formattedDate = `${year}/${parseInt(month)}/${parseInt(day)}`;
      console.log(`转换表头日期格式: ${header} -> ${adjustedDate} -> ${formattedDate}`);
      return formattedDate;
    }
    
    return header;
  });
  
  // 转换数据行，格式化百分比和保留两位小数
  const convertedRows = rows.map((row: any[], rowIndex: number) => {
    return row.map((cell: any, colIndex: number) => {
      const key = `${rowIndex}-${colIndex}`;
      const isPercentage = percentageMap.get(key);
      
      // 如果是数值且被标记为百分比
      if (typeof cell === 'number' && isPercentage) {
        // 保留两位小数，添加 % 符号
        return `${cell.toFixed(2)}%`;
      }
      
      // 如果是数值（非百分比），保留两位小数
      if (typeof cell === 'number' && !isPercentage) {
        return parseFloat(cell.toFixed(2));
      }
      
      return cell;
    });
  });
  
  const data = [convertedHeaders, ...convertedRows];
  console.log('转换后的数据行数:', data.length);
  console.log('第一行数据:', data[0]);
  
  const worksheet = XLSX.utils.aoa_to_sheet(data);
  
  // 确保数值单元格的数据类型是数字（防止被误认为是日期）
  Object.keys(worksheet).forEach(cellRef => {
    if (cellRef.startsWith('!')) return; // 跳过元数据
    const cell = worksheet[cellRef];
    if (cell && typeof cell.v === 'number') {
      cell.t = 'n'; // 设置为数字类型
    }
  });
  
  return worksheet;
}

/**
 * 清理文件名，移除特殊字符
 */
function sanitizeFilename(filename: string): string {
  return filename
    .replace(/[\\/:*?"<>|()]/g, '_')  // 替换特殊字符
    .replace(/\s+/g, '_')              // 替换空格
    .substring(0, 100);                // 限制长度
}

/**
 * 保存为 Excel 文件
 */
function saveAsExcel(tables: any[], filename: string, matchResults: any[] = []): string {
  console.log('开始保存 Excel 文件...');
  console.log('原始文件名:', filename);
  console.log('TEMP_DIR:', TEMP_DIR);
  console.log('TEMP_DIR 存在:', existsSync(TEMP_DIR));
  console.log('要保存的表格数量:', tables.length);
  console.log('匹配结果数量:', matchResults.length);
  
  const workbook = XLSX.utils.book_new();
  
  tables.forEach((table, index) => {
    console.log(`处理表格 ${index + 1}...`);
    const worksheet = tableToExcel(table, matchResults);
    XLSX.utils.book_append_sheet(workbook, worksheet, `Sheet${index + 1}`);
  });
  
  // 清理文件名
  const safeFilename = sanitizeFilename(filename);
  const filePath = path.join(TEMP_DIR, safeFilename);
  
  console.log('安全文件名:', safeFilename);
  console.log('完整文件路径:', filePath);
  
  try {
    // 确保目录存在
    if (!existsSync(TEMP_DIR)) {
      console.log('创建 temp 目录...');
      mkdirSync(TEMP_DIR, { recursive: true });
    }
    
    // 使用 XLSX.write 生成 buffer，然后用 writeFileSync 写入
    console.log('生成 Excel buffer...');
    const buffer = XLSX.write(workbook, { 
      type: 'buffer', 
      bookType: 'xlsx' 
    });
    
    console.log('写入文件...');
    writeFileSync(filePath, buffer);
    
    console.log('文件保存成功');
    return safeFilename;
  } catch (error) {
    console.error('保存文件失败，错误详情:', error);
    
    // 如果失败，使用时间戳作为文件名
    const fallbackFilename = `${Date.now()}.xlsx`;
    const fallbackPath = path.join(TEMP_DIR, fallbackFilename);
    
    console.log('尝试使用备用文件名:', fallbackFilename);
    
    try {
      const buffer = XLSX.write(workbook, { 
        type: 'buffer', 
        bookType: 'xlsx' 
      });
      writeFileSync(fallbackPath, buffer);
      console.log('备用文件保存成功');
      return fallbackFilename;
    } catch (fallbackError) {
      console.error('备用文件保存也失败:', fallbackError);
      throw new Error(`无法保存文件: ${fallbackError instanceof Error ? fallbackError.message : String(fallbackError)}`);
    }
  }
}

export async function POST(request: NextRequest) {
  try {
    await ensureTempDir();
    
    const formData = await request.formData();
    const fileA = formData.get('fileA') as File;
    const fileB = formData.get('fileB') as File;
    
    if (!fileA || !fileB) {
      return NextResponse.json({ error: '请上传两个文件' }, { status: 400 });
    }
    
    console.log('开始解析文件...');
    console.log('文件A:', fileA.name);
    console.log('文件B:', fileB.name);
    
    // 解析文件
    const parseResultA = await FileParser.parseFile(fileA);
    const parseResultB = await FileParser.parseFile(fileB);
    
    console.log('文件A解析完成，找到表格数:', parseResultA.tables.length);
    console.log('文件B解析完成，找到表格数:', parseResultB.tables.length);
    console.log('文件A单位:', parseResultA.unit);
    console.log('文件B单位:', parseResultB.unit);
    
    // 打印表格结构信息
    if (parseResultA.tables.length > 0) {
      console.log('文件A第一个表格表头:', parseResultA.tables[0].headers);
      console.log('文件A第一个表格行数:', parseResultA.tables[0].rows.length);
    }
    if (parseResultB.tables.length > 0) {
      console.log('文件B第一个表格表头:', parseResultB.tables[0].headers);
      console.log('文件B第一个表格行数:', parseResultB.tables[0].rows.length);
    }
    
    if (parseResultA.tables.length === 0) {
      return NextResponse.json({ error: '文件A中未找到表格数据' }, { status: 400 });
    }
    
    if (parseResultB.tables.length === 0) {
      return NextResponse.json({ error: '文件B中未找到表格数据' }, { status: 400 });
    }
    
    // 批量匹配
    // 根据需求：源文件默认单位为万元，输出文件转化为亿元
    const matcher = new BatchDataMatcher(
      parseResultA.tables,
      parseResultB.tables,
      parseResultA.unit || '万元',  // 源文件默认为万元
      '亿元',  // 输出文件强制为亿元
      parseResultA.tables[0]?.hasPercentage,
      parseResultB.tables[0]?.hasPercentage
    );
    
    console.log('开始数据匹配...');
    const { results } = matcher.matchAll();
    
    // 统计信息
    let totalFilled = 0;
    let totalConverted = 0;
    
    results.forEach((result, index) => {
      console.log(`表格 ${index + 1} 匹配完成:`, result.statistics);
      totalFilled += result.statistics.totalFilled;
      totalConverted += result.statistics.convertedCount;
    });
    
    // 提取填充后的表格
    const filledTables = results.map(r => r.filledTable);
    
    console.log('准备保存表格，数量:', filledTables.length);
    
    // 打印第一个表格的详细信息
    if (filledTables.length > 0) {
      console.log('第一个表格 headers:', filledTables[0].headers);
      console.log('第一个表格 rows 数量:', filledTables[0].rows?.length);
      if (filledTables[0].rows && filledTables[0].rows.length > 0) {
        console.log('第一个表格第一行数据:', filledTables[0].rows[0]);
      }
    }
    
    // 验证表格数据
    if (filledTables.length === 0) {
      throw new Error('没有可保存的表格数据');
    }
    
    // 合并所有表格为一个（如果有多个表格）
    let finalTable = filledTables[0];
    let allMatchResults: any[] = [];
    
    if (filledTables.length > 1) {
      console.log('检测到多个表格，将合并为一个表格');
      // 使用第一个表格的结构，合并所有行的数据
      const mergedRows = [...finalTable.rows];
      for (let i = 1; i < filledTables.length; i++) {
        const table = filledTables[i];
        if (table.rows && table.rows.length > 0) {
          mergedRows.push(...table.rows);
        }
        // 合并 matchResults
        if (results[i]?.matchResults) {
          allMatchResults.push(...results[i].matchResults);
        }
      }
      if (results[0]?.matchResults) {
        allMatchResults.push(...results[0].matchResults);
      }
      finalTable = {
        ...finalTable,
        rows: mergedRows,
      };
      console.log('合并后的表格行数:', finalTable.rows.length);
    } else {
      // 只有一个表格，直接使用其 matchResults
      if (results[0]?.matchResults) {
        allMatchResults = results[0].matchResults;
      }
    }
    
    // 保存结果（只保存一个表格）
    const originalFilename = fileB.name.replace(/\.[^/.]+$/, ""); // 移除扩展名
    const fileId = `${Date.now()}_${originalFilename}.xlsx`;
    console.log('生成的文件ID:', fileId);
    
    const savedFilename = saveAsExcel([finalTable], fileId, allMatchResults);
    
    console.log('文件保存成功:', savedFilename);
    
    return NextResponse.json({ 
      success: true, 
      fileId: savedFilename,
      message: '处理完成',
      statistics: {
        totalFilled,
        totalConverted,
        tableCount: 1,  // 现在总是保存为一个表格
        mergedTables: filledTables.length > 1 ? filledTables.length : undefined,
      }
    });
  } catch (error) {
    console.error('处理错误:', error);
    return NextResponse.json({ 
      error: error instanceof Error ? error.message : '处理失败',
      details: error instanceof Error ? error.stack : undefined
    }, { status: 500 });
  }
}
