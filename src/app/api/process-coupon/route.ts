import { NextRequest, NextResponse } from 'next/server';
import { CouponMatcher } from '@/lib/coupon-matcher';

/**
 * 挑券功能API
 * 完全独立于其他功能模块
 * 
 * 支持单金额和多金额模式：
 * - 单金额：传入 amount 参数
 * - 多金额：传入 amounts 参数（逗号分隔）
 * 
 * 支持禁挑券功能：
 * - excludedBonds：禁挑券参数（逗号分隔）
 * - 数字：精确匹配债券代码(B列)
 * - 文字：模糊匹配债券简称(C列)
 */
export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get('file') as File;
    const bondType = formData.get('bondType') as string;
    
    // 支持两种参数格式
    const amountStr = formData.get('amount') as string;
    const amountsStr = formData.get('amounts') as string;
    
    // 禁挑券参数
    const excludedBondsStr = formData.get('excludedBonds') as string;

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

    // 解析金额
    let amounts: number[] = [];
    
    if (amountsStr) {
      // 多金额模式
      amounts = amountsStr.split(',').map(s => parseFloat(s.trim())).filter(n => n > 0);
    } else if (amountStr) {
      // 单金额模式（兼容旧版）
      const amount = parseFloat(amountStr);
      if (amount > 0) {
        amounts = [amount];
      }
    }

    if (amounts.length === 0) {
      return NextResponse.json(
        { error: '请输入有效的挑券金额' },
        { status: 400 }
      );
    }

    // 解析禁挑券列表（支持中英文逗号）
    let excludedBonds: string[] = [];
    if (excludedBondsStr && excludedBondsStr.trim()) {
      excludedBonds = excludedBondsStr.split(/[,，\s\n]+/).map(s => s.trim()).filter(s => s);
    }

    console.log('挑券请求参数:', {
      fileName: file.name,
      bondType,
      amounts,
      mode: amounts.length > 1 ? '多金额' : '单金额',
      excludedBonds: excludedBonds.length > 0 ? excludedBonds : '无',
    });

    // 创建挑券处理器
    const matcher = new CouponMatcher(
      file,
      bondType as 'treasury' | 'local',
      amounts,
      excludedBonds
    );

    // 执行处理
    const result = await matcher.process();

    // 生成输出文件
    const buffer = await result.workbook.xlsx.writeBuffer();

    // 生成文件名：债券类型_挑券金额_日期.xlsx
    const bondTypeName = result.statistics.bondType === 'local' ? '地方债' : '国债政金债';
    const now = new Date();
    const dateStr = `${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
    
    // 多金额时显示多个金额
    const amountDisplay = amounts.length > 1 
      ? `${amounts.map(a => a.toLocaleString()).join('+')}万元`
      : `${amounts[0].toLocaleString()}万元`;
    const filename = `${bondTypeName}_${amountDisplay}_${dateStr}.xlsx`;

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
