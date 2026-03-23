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
    // 行业代码映射：行业名称 -> 代码（支持D列->E列和A列->B列两种映射）
    let industryCodeMap = new Map<string, any>();
    // 行政区划映射：地区名称 -> 代码（支持A列->B列和G列->D/E/F/H列映射）
    let adminDivisionMapByName = new Map<string, any>();
    let adminDivisionMapByCode = new Map<string, any>();
    // 银行信息映射：银行名称 -> 代码（D列->E列）
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
        
        // 构建"行业代码"映射
        // 方式1：D列行业名称 -> E列代码（主要）
        // 方式2：A列行业名称 -> B列代码（备选）
        for (let i = 0; i < industryCodeData.length; i++) {
          const row = industryCodeData[i];
          if (row) {
            // D列 -> E列（行业名称 -> 代码）
            if (row[3]) {
              const industryName = String(row[3]).trim();
              const industryCode = row[4]; // E列
              if (industryName && industryCode) {
                industryCodeMap.set(industryName, industryCode);
              }
            }
            // A列 -> B列（备选）
            if (row[0]) {
              const industryName = String(row[0]).trim();
              const industryCode = row[1]; // B列
              if (industryName && industryCode) {
                industryCodeMap.set(industryName, industryCode);
              }
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
        const adminDivisionData = XLSX.utils.sheet_to_json(adminDivisionSheet, { header: 1 }) as any[][];
        
        // 构建"行政区划代码"映射
        for (let i = 0; i < adminDivisionData.length; i++) {
          const row = adminDivisionData[i];
          if (row) {
            // A列区县名称 -> B列区县代码
            if (row[0]) {
              const areaName = String(row[0]).trim();
              const areaCode = row[1]; // B列
              if (areaName && areaCode) {
                adminDivisionMapByName.set(areaName, areaCode);
              }
            }
            // G列名称 -> D/E/F列代码（优先取最细粒度的代码）
            if (row[6]) {
              const areaName = String(row[6]).trim();
              // 优先取F列（三级分类），其次E列（二级分类），最后D列（一级分类）
              const areaCode = row[5] || row[4] || row[3]; // F列 > E列 > D列
              if (areaName && areaCode) {
                adminDivisionMapByName.set(areaName, areaCode);
              }
            }
            // D列代码 -> 对应名称（用于代码反向查找）
            if (row[3]) {
              const code = String(row[3]).trim();
              const name = row[6] || row[0]; // G列或A列
              if (code) {
                adminDivisionMapByCode.set(code, name);
              }
            }
          }
        }
        console.log(`行政区划代码映射共 ${adminDivisionMapByName.size} 条记录`);
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
        
        // 构建"银行信息"映射
        // D列银行名称 -> E列银行代码
        for (let i = 0; i < bankInfoData.length; i++) {
          const row = bankInfoData[i];
          if (row) {
            // D列 -> E列
            if (row[3]) {
              const bankName = String(row[3]).trim();
              const bankCode = row[4]; // E列
              if (bankName && bankCode) {
                bankInfoMap.set(bankName, bankCode);
              }
            }
            // A列 -> B列（备选）
            if (row[0]) {
              const bankName = String(row[0]).trim();
              const bankCode = row[1]; // B列
              if (bankName && bankCode) {
                bankInfoMap.set(bankName, bankCode);
              }
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
          
          // 在D列（索引3）处理行业代码
          // V列值是行业名称，匹配行业代码表D列，取E列代码
          const vValue = qichachaRow.vValue ? String(qichachaRow.vValue).trim() : '';
          
          if (vValue && vValue !== '-') {
            // 用行业名称匹配行业代码表
            if (industryCodeMap.has(vValue)) {
              row[3] = industryCodeMap.get(vValue);
              industryMatchCount++;
            } else {
              row[3] = vValue; // 未匹配则保留原值
            }
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
          
          // 用地区名称匹配行政区划代码表
          if (eValue && eValue !== '-') {
            if (adminDivisionMapByName.has(eValue)) {
              row[4] = adminDivisionMapByName.get(eValue);
              adminDivisionMatchCount++;
            } else {
              row[4] = eValue; // 未匹配则保留原值
            }
          } else {
            row[4] = '-';
          }
          
          // 在F列（索引5）填入企查查T列内容，为空则填入"-"
          row[5] = qichachaRow.tValue !== null && qichachaRow.tValue !== undefined && qichachaRow.tValue !== '' 
            ? qichachaRow.tValue : '-';
          
          // 在G列（索引6）根据企查查E列内容转换企业规模代码
          const eColValue = qichachaRow.eValue ? String(qichachaRow.eValue).trim() : '';
          if (eColValue && eColValue !== '-') {
            // 根据企业规模转换代码
            if (eColValue.includes('S') || eColValue === 'S(小型)') {
              row[6] = 'CS03'; // 小型
            } else if (eColValue.includes('M') || eColValue === 'M(中型)') {
              row[6] = 'CS02'; // 中型
            } else if (eColValue.includes('L') || eColValue === 'L(大型)') {
              row[6] = 'CS01'; // 大型
            } else if (eColValue.includes('XS') || eColValue === 'XS(微型)') {
              row[6] = 'CS04'; // 微型
            } else {
              row[6] = eColValue; // 其他情况保留原值
            }
          } else {
            row[6] = '-';
          }
        }
        
        // 在I列（索引8）处理银行信息匹配（独立匹配，不依赖企查查）
        // 将企业名称表H列（索引7）内容与银行信息表D列匹配，取E列代码
        const hValue = row[7] ? String(row[7]).trim() : '';
        if (hValue && hValue !== '-') {
          // 标准化银行名称：把"XX银行XXXXX"或"XX银行"改成"XX银行股份有限公司"
          let normalizedBankName = hValue;
          
          // 如果不是以"股份有限公司"结尾，则进行标准化
          if (!normalizedBankName.endsWith('股份有限公司')) {
            // 匹配"XX银行"开头或以"XX银行"结尾的情况
            const bankMatch = normalizedBankName.match(/^(.+?银行)/);
            if (bankMatch) {
              normalizedBankName = bankMatch[1] + '股份有限公司';
            }
          }
          
          if (bankInfoMap.has(normalizedBankName)) {
            row[8] = bankInfoMap.get(normalizedBankName);
            bankMatchCount++;
          } else {
            row[8] = '-'; // 未匹配
          }
        } else {
          row[8] = '-';
        }
        
        if (qichachaMap.has(enterpriseName)) {
          console.log(`匹配成功: ${enterpriseName} -> B:${row[1]}, C:${row[2]}, D:${row[3]}, E:${row[4]}, F:${row[5]}, G:${row[6]}, I:${row[8]}`);
        }
      }
    }

    // 处理L列：K列数据 + 随机4位数字
    const lColumnValues: string[] = [];
    for (let i = 0; i < enterpriseNameData.length; i++) {
      const row = enterpriseNameData[i];
      if (row) {
        const kValue = row[10] ? String(row[10]).trim() : ''; // K列（索引10）
        if (kValue) {
          // 生成1000-9999之间的随机4位数字
          const randomNum = Math.floor(Math.random() * 9000) + 1000;
          const lValue = kValue + randomNum;
          row[11] = lValue; // L列（索引11）
          lColumnValues.push(lValue);
        }
      }
    }

    // 检查L列是否有重复项
    const lColumnSet = new Set(lColumnValues);
    const hasDuplicates = lColumnSet.size !== lColumnValues.length;
    
    if (hasDuplicates && enterpriseNameData.length > 0) {
      // 在L列第一行输入"有重复项"
      if (!enterpriseNameData[0]) {
        enterpriseNameData[0] = [];
      }
      enterpriseNameData[0][11] = '有重复项';
      console.log('L列存在重复项');
    }

    console.log(`共匹配成功 ${matchCount} 条数据`);
    console.log(`其中C01(含公司) ${c01Count} 条，C02(不含公司) ${c02Count} 条`);
    console.log(`行业代码转换成功 ${industryMatchCount} 条`);
    console.log(`行政区划代码转换成功 ${adminDivisionMatchCount} 条`);
    console.log(`银行信息匹配成功 ${bankMatchCount} 条`);
    console.log(`L列生成 ${lColumnValues.length} 条数据，${hasDuplicates ? '存在重复' : '无重复'}`);

    // 将数据写回工作表
    const newSheet = XLSX.utils.aoa_to_sheet(enterpriseNameData);
    enterpriseNameWorkbook.Sheets[enterpriseNameSheetName] = newSheet;

    // 生成输出文件
    const outputBuffer = XLSX.write(enterpriseNameWorkbook, { type: 'buffer', bookType: 'xlsx' });

    // 生成文件名：日期+PJRZFS（如0322PJRZFS.xlsx）
    const now = new Date();
    const datePrefix = `${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
    const filename = `${datePrefix}PJRZFS.xlsx`;

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
