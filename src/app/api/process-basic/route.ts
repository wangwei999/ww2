import { NextRequest, NextResponse } from 'next/server';
import * as XLSX from 'xlsx';

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    
    const enterpriseNameFile = formData.get('enterpriseNameFile') as File;
    const qichachaDataFile = formData.get('qichachaDataFile') as File;
    const reportFieldsFile = formData.get('reportFieldsFile') as File;

    if (!enterpriseNameFile || !qichachaDataFile) {
      return NextResponse.json(
        { error: '请上传企业名称文件和企查查数据文件' },
        { status: 400 }
      );
    }

    // 读取企业名称文件
    const enterpriseNameBuffer = await enterpriseNameFile.arrayBuffer();
    const enterpriseNameWorkbook = XLSX.read(Buffer.from(enterpriseNameBuffer), { type: 'buffer' });
    const enterpriseNameSheetName = enterpriseNameWorkbook.SheetNames[0];
    const enterpriseNameSheet = enterpriseNameWorkbook.Sheets[enterpriseNameSheetName];
    const enterpriseNameData = XLSX.utils.sheet_to_json(enterpriseNameSheet, { header: 1 }) as any[][];

    // 读取企查查数据文件
    const qichachaDataBuffer = await qichachaDataFile.arrayBuffer();
    const qichachaDataWorkbook = XLSX.read(Buffer.from(qichachaDataBuffer), { type: 'buffer' });
    const qichachaDataSheetName = qichachaDataWorkbook.SheetNames[0];
    const qichachaDataSheet = qichachaDataWorkbook.Sheets[qichachaDataSheetName];
    const qichachaData = XLSX.utils.sheet_to_json(qichachaDataSheet, { header: 1 }) as any[][];

    // 构建企查查数据映射：A列名称 -> D列数据
    const qichachaMap = new Map<string, any>();
    for (let i = 0; i < qichachaData.length; i++) {
      const row = qichachaData[i];
      if (row && row[0]) {
        const name = String(row[0]).trim();
        const value = row[3]; // D列（索引3）
        if (name) {
          qichachaMap.set(name, value);
        }
      }
    }

    console.log(`企查查数据共 ${qichachaMap.size} 条记录`);

    // 匹配并填充数据
    let matchCount = 0;
    let c01Count = 0;
    let c02Count = 0;
    for (let i = 0; i < enterpriseNameData.length; i++) {
      const row = enterpriseNameData[i];
      if (row && row[0]) {
        const enterpriseName = String(row[0]).trim();
        if (enterpriseName && qichachaMap.has(enterpriseName)) {
          // 在B列（索引1）填入企查查D列数据
          row[1] = qichachaMap.get(enterpriseName);
          matchCount++;
          
          // 在C列（索引2）根据是否包含"公司"填入分类码
          if (enterpriseName.includes('公司')) {
            row[2] = 'C01';
            c01Count++;
          } else {
            row[2] = 'C02';
            c02Count++;
          }
          
          console.log(`匹配成功: ${enterpriseName} -> B列:${qichachaMap.get(enterpriseName)}, C列:${row[2]}`);
        }
      }
    }

    console.log(`共匹配成功 ${matchCount} 条数据，其中C01(含公司) ${c01Count} 条，C02(不含公司) ${c02Count} 条`);

    // 将数据写回工作表
    const newSheet = XLSX.utils.aoa_to_sheet(enterpriseNameData);
    enterpriseNameWorkbook.Sheets[enterpriseNameSheetName] = newSheet;

    // 生成输出文件
    const outputBuffer = XLSX.write(enterpriseNameWorkbook, { type: 'buffer', bookType: 'xlsx' });

    // 生成文件名
    const now = new Date();
    const datePrefix = `${String(now.getFullYear()).slice(2)}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
    const originalName = enterpriseNameFile.name.replace(/\.[^/.]+$/, '');
    const filename = `${datePrefix}基础数据处理结果.xlsx`;

    // 返回文件
    return new NextResponse(outputBuffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`,
      },
    });
  } catch (error) {
    console.error('基础数据处理错误:', error);
    return NextResponse.json(
      { error: error instanceof Error ? error.message : '处理失败' },
      { status: 500 }
    );
  }
}
