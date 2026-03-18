import { NextRequest, NextResponse } from 'next/server';
import { PDFMatcher } from '@/lib/pdf-matcher';
import { HeaderUtils } from 'coze-coding-dev-sdk';

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

    // 生成输出文件
    const buffer = await result.workbook.xlsx.writeBuffer();

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
