import { NextRequest, NextResponse } from 'next/server';
import { CouponCopyProcessor } from '@/lib/coupon-copy-processor';

/**
 * 债券副本分类 API
 * 
 * 上传Excel文件后自动生成三个分类工作表：
 * - 国债：排除地方债和政金债
 * - 政金债：包含"国开"、"进出口"、"农发"的债券
 * - 地方债：包含地理名词的债券
 * 
 * 所有分类按可用金额从大到小排序
 */
export async function POST(request: NextRequest) {
  try {
    console.log('=== 开始处理债券副本分类 ===');

    const formData = await request.formData();
    const file = formData.get('file') as File;

    if (!file) {
      return NextResponse.json(
        { error: '请上传文件' },
        { status: 400 }
      );
    }

    // 验证文件格式
    const fileName = file.name.toLowerCase();
    if (!fileName.endsWith('.xls') && !fileName.endsWith('.xlsx')) {
      return NextResponse.json(
        { error: '只支持 .xls 和 .xlsx 格式的文件' },
        { status: 400 }
      );
    }

    console.log('文件名:', file.name);
    console.log('文件大小:', (file.size / 1024).toFixed(2), 'KB');

    // 执行处理
    const processor = new CouponCopyProcessor(file);
    const result = await processor.process();

    // 生成输出文件
    const buffer = await result.workbook.xlsx.writeBuffer();

    // 生成文件名
    const now = new Date();
    const dateStr = `${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
    const outputFileName = `债券分类_${dateStr}.xlsx`;

    console.log('处理完成，输出文件:', outputFileName);
    console.log('统计:', result.statistics);

    // 返回文件
    return new NextResponse(buffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${encodeURIComponent(outputFileName)}"`,
        'X-Statistics': JSON.stringify(result.statistics)
      }
    });

  } catch (error) {
    console.error('处理失败:', error);
    return NextResponse.json(
      { error: error instanceof Error ? error.message : '处理失败' },
      { status: 500 }
    );
  }
}
