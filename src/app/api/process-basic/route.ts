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

    // 构建企查查数据映射：A列名称 -> { D列数据, E列数据, M列数据, N列数据, T列数据, V列数据 }
    const qichachaMap = new Map<string, { dValue: any; eValue: any; mValue: any; nValue: any; tValue: any; vValue: any }>();
    for (let i = 0; i < qichachaData.length; i++) {
      const row = qichachaData[i];
      if (row && row[0]) {
        const name = String(row[0]).trim();
        const dValue = row[3]; // D列（索引3）
        const eValue = row[4]; // E列（索引4）
        const mValue = row[12]; // M列（索引12）
        const nValue = row[13]; // N列（索引13）
        const tValue = row[19]; // T列（索引19）
        const vValue = row[21]; // V列（索引21）
        if (name) {
          qichachaMap.set(name, { dValue, eValue, mValue, nValue, tValue, vValue });
        }
      }
    }

    console.log(`企查查数据共 ${qichachaMap.size} 条记录`);

    // 读取报表字段文件
    let industryCodeMap = new Map<string, any>();
    let adminDivisionMap = new Map<string, { cValue: any; cValueNextRow: any }>();
    let adminDivisionData: any[][] = [];
    let bankInfoMap = new Map<string, any>();
    
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
        console.log('未找到"行业代码"工作表');
      }

      // 查找"行政区划代码"工作表
      const adminDivisionSheetName = reportFieldsWorkbook.SheetNames.find(
        name => name.includes('行政区划代码')
      );
      
      if (adminDivisionSheetName) {
        const adminDivisionSheet = reportFieldsWorkbook.Sheets[adminDivisionSheetName];
        adminDivisionData = XLSX.utils.sheet_to_json(adminDivisionSheet, { header: 1 }) as any[][];
        
        // 构建"行政区划代码"映射：D列 -> { C列值, C列下一行值 }
        for (let i = 0; i < adminDivisionData.length; i++) {
          const row = adminDivisionData[i];
          if (row && row[3]) { // D列（索引3）
            const dValue = String(row[3]).trim();
            const cValue = row[2]; // C列（索引2）
            // 获取下一行C列的值
            const cValueNextRow = (i + 1 < adminDivisionData.length && adminDivisionData[i + 1]) 
              ? adminDivisionData[i + 1][2] 
              : null;
            if (dValue) {
              adminDivisionMap.set(dValue, { cValue, cValueNextRow });
            }
          }
        }
        console.log(`行政区划代码映射共 ${adminDivisionMap.size} 条记录`);
      } else {
        console.log('未找到"行政区划代码"工作表');
      }

      // 查找"银行信息"工作表
      const bankInfoSheetName = reportFieldsWorkbook.SheetNames.find(
        name => name.includes('银行信息')
      );
      
      if (bankInfoSheetName) {
        const bankInfoSheet = reportFieldsWorkbook.Sheets[bankInfoSheetName];
        const bankInfoData = XLSX.utils.sheet_to_json(bankInfoSheet, { header: 1 }) as any[][];
        
        // 构建"银行信息"映射：A列 -> B列
        for (let i = 0; i < bankInfoData.length; i++) {
          const row = bankInfoData[i];
          if (row && row[0]) {
            const bankName = String(row[0]).trim();
            const bankValue = row[1]; // B列（索引1）
            if (bankName) {
              bankInfoMap.set(bankName, bankValue);
            }
          }
        }
        console.log(`银行信息映射共 ${bankInfoMap.size} 条记录`);
      } else {
        console.log('未找到"银行信息"工作表，可用工作表:', reportFieldsWorkbook.SheetNames);
      }
    }

    // 匹配并填充数据
    let matchCount = 0;
    let c01Count = 0;
    let c02Count = 0;
    let industryMatchCount = 0;
    let adminDivisionMatchCount = 0;
    let bankMatchCount = 0;
    
    for (let i = 0; i < enterpriseNameData.length; i++) {
      const row = enterpriseNameData[i];
      if (row && row[0]) {
        const enterpriseName = String(row[0]).trim();
        if (enterpriseName && qichachaMap.has(enterpriseName)) {
          const qichachaRow = qichachaMap.get(enterpriseName)!;
          
          // 在B列（索引1）填入企查查D列数据，为空则填入"-"
          row[1] = qichachaRow.dValue !== null && qichachaRow.dValue !== undefined && qichachaRow.dValue !== '' 
            ? qichachaRow.dValue : '-';
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
          if (vValue && vValue !== '-' && industryCodeMap.has(vValue)) {
            row[3] = industryCodeMap.get(vValue);
            industryMatchCount++;
          } else if (vValue && vValue !== '-') {
            row[3] = vValue;
          } else {
            row[3] = '-';
          }
          
          // 在E列（索引4）处理N列/M列数据和行政区划代码转换
          // 获取N列内容，如果N列显示"-"则用M列内容
          let eValue = '';
          const nValue = qichachaRow.nValue ? String(qichachaRow.nValue).trim() : '';
          const mValue = qichachaRow.mValue ? String(qichachaRow.mValue).trim() : '';
          
          if (nValue === '-' || nValue === '') {
            eValue = mValue;
          } else {
            eValue = nValue;
          }
          
          // 再用E列值与"行政区划代码"表D列匹配
          if (eValue && eValue !== '-' && adminDivisionMap.has(eValue)) {
            const adminData = adminDivisionMap.get(eValue)!;
            // 如果C列有数据，用C列内容；否则用C列下一行数据
            if (adminData.cValue !== null && adminData.cValue !== undefined && adminData.cValue !== '') {
              row[4] = adminData.cValue;
            } else if (adminData.cValueNextRow !== null && adminData.cValueNextRow !== undefined) {
              row[4] = adminData.cValueNextRow;
            } else {
              row[4] = eValue; // 都没有则保留原值
            }
            adminDivisionMatchCount++;
          } else if (eValue && eValue !== '-') {
            row[4] = eValue;
          } else {
            row[4] = '-';
          }
          
          // 在F列（索引5）填入企查查T列内容，为空则填入"-"
          row[5] = qichachaRow.tValue !== null && qichachaRow.tValue !== undefined && qichachaRow.tValue !== '' 
            ? qichachaRow.tValue : '-';
          
          // 在G列（索引6）填入企查查E列内容，为空则填入"-"
          row[6] = qichachaRow.eValue !== null && qichachaRow.eValue !== undefined && qichachaRow.eValue !== '' 
            ? qichachaRow.eValue : '-';
        }
        
        // 在I列（索引8）处理银行信息匹配（独立匹配，不依赖企查查）
        // 将企业名称表H列（索引7）内容与银行信息表A列匹配
        const hValue = row[7] ? String(row[7]).trim() : '';
        if (hValue && hValue !== '-' && bankInfoMap.has(hValue)) {
          row[8] = bankInfoMap.get(hValue);
          bankMatchCount++;
        } else {
          row[8] = '-';
        }
        
        if (qichachaMap.has(enterpriseName)) {
          console.log(`匹配成功: ${enterpriseName} -> B:${row[1]}, C:${row[2]}, D:${row[3]}, E:${row[4]}, F:${row[5]}, G:${row[6]}, I:${row[8]}`);
        }
      }
    }

    console.log(`共匹配成功 ${matchCount} 条数据`);
    console.log(`其中C01(含公司) ${c01Count} 条，C02(不含公司) ${c02Count} 条`);
    console.log(`行业代码转换成功 ${industryMatchCount} 条`);
    console.log(`行政区划代码转换成功 ${adminDivisionMatchCount} 条`);
    console.log(`银行信息匹配成功 ${bankMatchCount} 条`);

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
