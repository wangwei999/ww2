import { NextRequest, NextResponse } from 'next/server';
import { ImageMatcher } from '@/lib/image-matcher';
import { HeaderUtils } from 'coze-coding-dev-sdk';

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    
    // 获取所有图片文件
    const imageFiles: File[] = [];
    let index = 0;
    while (true) {
      const file = formData.get(`imageFile${index}`) as File | null;
      if (!file) break;
      imageFiles.push(file);
      index++;
    }

    const excelFile = formData.get('excelFile') as File;

    if (imageFiles.length === 0 || !excelFile) {
      return NextResponse.json(
        { error: '请上传图片文件和Excel文件' },
        { status: 400 }
      );
    }

    // 验证图片文件类型
    const validImageTypes = ['image/png', 'image/jpeg', 'image/jpg', 'image/gif', 'image/webp'];
    for (const imageFile of imageFiles) {
      if (!validImageTypes.includes(imageFile.type) && 
          !imageFile.name.toLowerCase().match(/\.(png|jpe?g|gif|webp)$/)) {
        return NextResponse.json(
          { error: `文件 ${imageFile.name} 不是有效的图片格式` },
          { status: 400 }
        );
      }
    }

    // 验证Excel文件类型
    if (!excelFile.name.toLowerCase().match(/\.(xlsx|xls)$/)) {
      return NextResponse.json(
        { error: 'Excel文件必须是.xlsx或.xls格式' },
        { status: 400 }
      );
    }

    // 提取请求头
    const customHeaders = HeaderUtils.extractForwardHeaders(request.headers);

    // 创建图片处理器
    const matcher = new ImageMatcher(imageFiles, excelFile, customHeaders);

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
    console.error('图片模式处理失败:', error);
    return NextResponse.json(
      { error: error.message || '处理失败，请重试' },
      { status: 500 }
    );
  }
}
