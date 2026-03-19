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

    // 关键：创建一个全新的工作簿，完全避免共享公式引用
    const cleanWorkbook = await createFormulaFreeWorkbook(result.workbook);

    // 生成输出文件
    const buffer = await cleanWorkbook.xlsx.writeBuffer();

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
 * 创建一个不含任何公式的工作簿（彻底避免共享公式错误）
 */
async function createFormulaFreeWorkbook(sourceWorkbook: ExcelJS.Workbook): Promise<ExcelJS.Workbook> {
  const newWorkbook = new ExcelJS.Workbook();

  sourceWorkbook.eachSheet((sourceSheet) => {
    const newSheet = newWorkbook.addWorksheet(sourceSheet.name);

    // 复制列宽
    sourceSheet.columns.forEach((col, index) => {
      if (col.width) {
        newSheet.getColumn(index + 1).width = col.width;
      }
    });

    const maxRow = sourceSheet.rowCount || 200;
    const maxCol = sourceSheet.columnCount || 50;

    // 按行按列复制数据
    for (let rowNumber = 1; rowNumber <= maxRow; rowNumber++) {
      const sourceRow = sourceSheet.getRow(rowNumber);
      const newRow = newSheet.getRow(rowNumber);
      newRow.height = sourceRow.height;

      for (let colNumber = 1; colNumber <= maxCol; colNumber++) {
        const sourceCell = sourceRow.getCell(colNumber);
        const newCell = newRow.getCell(colNumber);

        // 获取单元格的值（优先使用结果值）
        let cellValue: any = null;
        try {
          const cellData = sourceCell as any;
          
          // 如果有公式，使用结果值
          if (cellData.sharedFormula || cellData.formula) {
            const result = cellData.result;
            cellValue = result !== undefined && result !== null ? result : null;
          } else {
            // 否则直接使用value
            cellValue = sourceCell.value;
          }
        } catch (e) {
          cellValue = sourceCell.value;
        }

        // 设置值
        newCell.value = cellValue;

        // 复制样式
        try {
          if (sourceCell.font) {
            newCell.font = JSON.parse(JSON.stringify(sourceCell.font));
          }
          if (sourceCell.fill) {
            newCell.fill = JSON.parse(JSON.stringify(sourceCell.fill));
          }
          if (sourceCell.border) {
            newCell.border = JSON.parse(JSON.stringify(sourceCell.border));
          }
          if (sourceCell.alignment) {
            newCell.alignment = JSON.parse(JSON.stringify(sourceCell.alignment));
          }
          if (sourceCell.numFmt) {
            newCell.numFmt = sourceCell.numFmt;
          }
        } catch (e) {
          // 忽略样式复制错误
        }
      }
    }

    // 复制合并单元格
    try {
      const merges = (sourceSheet as any)._merges;
      if (merges) {
        Object.values(merges).forEach((merge: any) => {
          try {
            newSheet.mergeCells(merge);
          } catch (e) {
            // 忽略合并错误
          }
        });
      }
    } catch (e) {
      // 忽略合并单元格复制错误
    }
  });

  return newWorkbook;
}
