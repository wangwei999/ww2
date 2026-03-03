import { NextRequest, NextResponse } from 'next/server';
import { writeFile, mkdir } from 'fs/promises';
import { existsSync } from 'fs';
import path from 'path';
import * as XLSX from 'xlsx';
import { FileParser } from '@/lib/file-parser';
import { BatchDataMatcher } from '@/lib/data-matcher';

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
function tableToExcel(table: any): XLSX.WorkSheet {
  const { headers, rows } = table;
  const data = [headers, ...rows];
  return XLSX.utils.aoa_to_sheet(data);
}

/**
 * 保存为 Excel 文件
 */
function saveAsExcel(tables: any[], filename: string): string {
  const workbook = XLSX.utils.book_new();
  
  tables.forEach((table, index) => {
    const worksheet = tableToExcel(table);
    XLSX.utils.book_append_sheet(workbook, worksheet, `Sheet${index + 1}`);
  });
  
  const filePath = path.join(TEMP_DIR, filename);
  XLSX.writeFile(workbook, filePath);
  
  return filePath;
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
    
    if (parseResultA.tables.length === 0) {
      return NextResponse.json({ error: '文件A中未找到表格数据' }, { status: 400 });
    }
    
    if (parseResultB.tables.length === 0) {
      return NextResponse.json({ error: '文件B中未找到表格数据' }, { status: 400 });
    }
    
    // 批量匹配
    const matcher = new BatchDataMatcher(
      parseResultA.tables,
      parseResultB.tables,
      parseResultA.unit,
      parseResultB.unit,
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
    
    // 保存结果
    const fileId = `${Date.now()}_${fileB.name}`;
    const outputPath = saveAsExcel(filledTables, fileId);
    
    console.log('文件保存成功:', outputPath);
    
    return NextResponse.json({ 
      success: true, 
      fileId,
      message: '处理完成',
      statistics: {
        totalFilled,
        totalConverted,
        tableCount: filledTables.length,
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
