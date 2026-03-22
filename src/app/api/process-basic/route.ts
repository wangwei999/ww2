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

    // 构建企查查数据映射：A列名称 -> { D列数据, V列数据 }
    const qichachaMap = new Map<string, { dValue: any; vValue: any }>();
    for (let i = 0; i < qichachaData.length; i++) {
      const row = qichachaData[i];
      if (row && row[0]) {
        const name = String(row[0]).trim();
        const dValue = row[3]; // D列（索引3）
        const vValue = row[21]; // V列（索引21）
        if (name) {
          qichachaMap.set(name, { dValue, vValue });
        }
      }
    }

    console.log(`企查查数据共 ${qichachaMap.size} 条记录`);

    // 读取报表字段文件中的"行业代码"表
    let industryCodeMap = new Map<string, any>();
    if (reportFieldsFile) {
      const reportFieldsBuffer = await reportFieldsFile.arrayBuffer();
      const reportFieldsWorkbook = XLSX.read(Buffer.from(reportFieldsBuffer), { type: 'buffer' });
      
      // 查找"行业代码"工作表
      const industryCodeSheetName = reportFieldsWorkbook.SheetNames.find(
        name => name.includes('行业代码')
      );
      
      if (industryCodeSheetName) {
        const industryCodeSheet = reportFieldsWorkbook.Sheets[industryCodeSheetName];
        const industryCodeData = XLSX.utils.sheet_to_json(industryCodeSheet, { header: 1 }) as any[][];
        
        // 构建"行业代码"映射：A列 -> B列
        for (let i = 0; i < industryCodeData.length; i++) {
          const row = industryCodeData[i];
          if (row && row[0]) {
            const code = String(row[0]).trim();
            const name = row[1]; // B列（索引1）
            if (code) {
              industryCodeMap.set(code, name);
            }
          }
        }
        console.log(`行业代码映射共 ${industryCodeMap.size} 条记录`);
      } else {
        console.log('未找到"行业代码"工作表，可用工作表:', reportFieldsWorkbook.SheetNames);
      }
    }

    // 匹配并填充数据
    let matchCount = 0;
    let c01Count = 0;
    let c02Count = 0;
    let industryMatchCount = 0;
    
    for (let i = 0; i < enterpriseNameData.length; i++) {
      const row = enterpriseNameData[i];
      if (row && row[0]) {
        const enterpriseName = String(row[0]).trim();
        if (enterpriseName && qichachaMap.has(enterpriseName)) {
          const qichachaRow = qichachaMap.get(enterpriseName)!;
          
          // 在B列（索引1）填入企查查D列数据
          row[1] = qichachaRow.dValue;
          matchCount++;
          
          // 在C列（索引2）根据是否包含"公司"填入分类码
          if (enterpriseName.includes('公司')) {
            row[2] = 'C01';
            c01Count++;
          } else {
            row[2] = 'C02';
            c02Count++;
          }
          
          // 在D列（索引3）先填入企查查V列数据
          const vValue = qichachaRow.vValue ? String(qichachaRow.vValue).trim() : '';
          
          // 再用D列值与"行业代码"表A列匹配，如果匹配成功则替换为B列内容
          if (vValue && industryCodeMap.has(vValue)) {
            row[3] = industryCodeMap.get(vValue);
            industryMatchCount++;
            console.log(`匹配成功: ${enterpriseName} -> B:${qichachaRow.dValue}, C:${row[2]}, D:${row[3]}(行业代码转换)`);
          } else {
            row[3] = qichachaRow.vValue;
            console.log(`匹配成功: ${enterpriseName} -> B:${qichachaRow.dValue}, C:${row[2]}, D:${row[3]}(V列原值)`);
          }
        }
      }
    }

    console.log(`共匹配成功 ${matchCount} 条数据`);
    console.log(`其中C01(含公司) ${c01Count} 条，C02(不含公司) ${c02Count} 条`);
    console.log(`行业代码转换成功 ${industryMatchCount} 条`);

    // 将数据写回工作表
    const newSheet = XLSX.utils.aoa_to_sheet(enterpriseNameData);
    enterpriseNameWorkbook.Sheets[enterpriseNameSheetName] = newSheet;

    // 生成输出文件
    const outputBuffer = XLSX.write(enterpriseNameWorkbook, { type: 'buffer', bookType: 'xlsx' });

    // 生成文件名
    const now = new Date();
    const datePrefix = `${String(now.getFullYear()).slice(2)}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
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
