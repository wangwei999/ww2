import { NextRequest, NextResponse } from 'next/server';
import { PDFMatcher } from '@/lib/pdf-matcher';
import { HeaderUtils } from 'coze-coding-dev-sdk';
import ExcelJS from 'exceljs';

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const pdfFile = formData.get('pdfFile') as File;
    const excelFile = formData.get('excelFile') as File;

    if (!pdfFile || !excelFile) {
      return NextResponse.json(
        { error: '请上传PDF文件和Excel文件' },
        { status: 400 }
      );
    }

    // 验证文件类型
    if (!pdfFile.name.toLowerCase().endsWith('.pdf')) {
      return NextResponse.json(
        { error: '文件A必须是PDF格式' },
        { status: 400 }
      );
    }

    if (!excelFile.name.toLowerCase().match(/\.(xlsx|xls)$/)) {
      return NextResponse.json(
        { error: '文件B必须是Excel格式(.xlsx或.xls)' },
        { status: 400 }
      );
    }

    // 提取请求头
    const customHeaders = HeaderUtils.extractForwardHeaders(request.headers);

    // 创建PDF处理器
    const matcher = new PDFMatcher(pdfFile, excelFile, customHeaders);

    // 处理文件
    const result = await matcher.process();

    // 关键：直接提取数据，不调用writeBuffer
    const workbookData = extractWorkbookData(result.workbook);
    
    // 创建全新的工作簿
    const finalWorkbook = createNewWorkbook(workbookData);

    // 生成输出文件
    const buffer = await finalWorkbook.xlsx.writeBuffer();

    // 返回文件
    return new NextResponse(buffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="processed_${Date.now()}.xlsx"`,
      },
    });
  } catch (error: any) {
    console.error('PDF模式处理失败:', error);
    return NextResponse.json(
      { error: error.message || '处理失败，请重试' },
      { status: 500 }
    );
  }
}

/**
 * 提取工作簿数据（不触发writeBuffer）
 */
interface CellData {
  value: any;
  style?: {
    font?: any;
    fill?: any;
    border?: any;
    alignment?: any;
    numFmt?: string;
  };
}

interface SheetData {
  name: string;
  columns: { width?: number }[];
  rows: {
    height?: number;
    cells: CellData[];
  }[];
  merges: string[];
}

function extractWorkbookData(workbook: ExcelJS.Workbook): SheetData[] {
  const sheets: SheetData[] = [];

  workbook.eachSheet((worksheet) => {
    const sheetData: SheetData = {
      name: worksheet.name,
      columns: [],
      rows: [],
      merges: [],
    };

    // 提取列宽
    try {
      worksheet.columns.forEach((col) => {
        sheetData.columns.push({ width: col.width });
      });
    } catch (e) {}

    const maxRow = worksheet.rowCount || 200;
    const maxCol = worksheet.columnCount || 50;

    // 提取所有单元格数据
    for (let rowNumber = 1; rowNumber <= maxRow; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      const rowData: { height?: number; cells: CellData[] } = {
        height: row.height,
        cells: [],
      };

      for (let colNumber = 1; colNumber <= maxCol; colNumber++) {
        const cell = row.getCell(colNumber);
        const cellData: CellData = { value: null };

        try {
          // 直接访问model，避免触发getter
          const model = (cell as any).model;
          
          if (model) {
            // 获取值（如果有公式，使用公式的值）
            cellData.value = model.value;
            
            // 获取样式
            if (model.style) {
              cellData.style = {
                font: model.style.font,
                fill: model.style.fill,
                border: model.style.border,
                alignment: model.style.alignment,
                numFmt: model.style.numFmt,
              };
            }
          }
        } catch (e) {
          // 备选方案
          try {
            cellData.value = cell.value;
          } catch (e2) {}
        }

        rowData.cells.push(cellData);
      }

      sheetData.rows.push(rowData);
    }

    // 提取合并单元格
    try {
      const merges = (worksheet as any)._merges;
      if (merges) {
        sheetData.merges = Object.keys(merges);
      }
    } catch (e) {}

    sheets.push(sheetData);
  });

  return sheets;
}

/**
 * 创建新工作簿（不包含任何公式引用）
 */
function createNewWorkbook(sheetsData: SheetData[]): ExcelJS.Workbook {
  const workbook = new ExcelJS.Workbook();

  sheetsData.forEach((sheetData) => {
    const worksheet = workbook.addWorksheet(sheetData.name);

    // 设置列宽
    sheetData.columns.forEach((col, index) => {
      if (col.width) {
        worksheet.getColumn(index + 1).width = col.width;
      }
    });

    // 写入数据
    sheetData.rows.forEach((rowData, rowIndex) => {
      const row = worksheet.getRow(rowIndex + 1);
      if (rowData.height) {
        row.height = rowData.height;
      }

      rowData.cells.forEach((cellData, colIndex) => {
        const cell = row.getCell(colIndex + 1);
        
        // 设置值
        cell.value = cellData.value;
        
        // 设置样式
        if (cellData.style) {
          try {
            if (cellData.style.font) {
              cell.font = cellData.style.font;
            }
            if (cellData.style.fill) {
              cell.fill = cellData.style.fill;
            }
            if (cellData.style.border) {
              cell.border = cellData.style.border;
            }
            if (cellData.style.alignment) {
              cell.alignment = cellData.style.alignment;
            }
            if (cellData.style.numFmt) {
              cell.numFmt = cellData.style.numFmt;
            }
          } catch (e) {}
        }
      });
    });

    // 合并单元格
    sheetData.merges.forEach((merge) => {
      try {
        worksheet.mergeCells(merge);
      } catch (e) {}
    });
  });

  return workbook;
}
