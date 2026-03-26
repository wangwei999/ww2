import { NextRequest, NextResponse } from 'next/server';
import { CouponMatcher } from '@/lib/coupon-matcher';

/**
 * 挑券功能API
 * 完全独立于其他功能模块
 */
export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get('file') as File;
    const bondType = formData.get('bondType') as string;
    const amount = formData.get('amount') as string;

    // 参数验证
    if (!file) {
      return NextResponse.json(
        { error: '请上传Excel文件' },
        { status: 400 }
      );
    }

    if (!bondType || !['treasury', 'local'].includes(bondType)) {
      return NextResponse.json(
        { error: '请选择债券类型（国债或地方债）' },
        { status: 400 }
      );
    }

    if (!amount || parseFloat(amount) <= 0) {
      return NextResponse.json(
        { error: '请输入有效的挑券金额' },
        { status: 400 }
      );
    }

    console.log('挑券请求参数:', {
      fileName: file.name,
      bondType,
      amount: parseFloat(amount),
    });

    // 创建挑券处理器
    const matcher = new CouponMatcher(
      file,
      bondType as 'treasury' | 'local',
      parseFloat(amount)
    );

    // 执行处理
    const result = await matcher.process();

    // 生成输出文件
    const buffer = await result.workbook.xlsx.writeBuffer();

    // 生成文件名：债券类型_挑券金额_日期.xlsx
    const bondTypeName = result.statistics.bondType === 'local' ? '地方债' : '国债';
    const now = new Date();
    const dateStr = `${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
    const filename = `${bondTypeName}_${amount}万元_${dateStr}.xlsx`;

    // 返回文件
    return new NextResponse(buffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${encodeURIComponent(filename)}"`,
      },
    });

  } catch (error: any) {
    console.error('挑券处理失败:', error);
    return NextResponse.json(
      { error: error.message || '处理失败，请重试' },
      { status: 500 }
    );
  }
}
