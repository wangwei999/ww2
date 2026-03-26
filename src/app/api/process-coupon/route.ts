import { NextRequest, NextResponse } from 'next/server';

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

    // TODO: 实现挑券处理逻辑
    // 这里后续会添加具体的处理代码

    console.log('挑券请求参数:', {
      fileName: file.name,
      bondType,
      amount: parseFloat(amount),
    });

    // 暂时返回占位响应
    return NextResponse.json(
      { error: '挑券功能正在开发中，请稍后再试' },
      { status: 501 }
    );

  } catch (error: any) {
    console.error('挑券处理失败:', error);
    return NextResponse.json(
      { error: error.message || '处理失败，请重试' },
      { status: 500 }
    );
  }
}
