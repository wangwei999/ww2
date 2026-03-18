import ExcelJS from 'exceljs';
import { LLMClient, Config, HeaderUtils } from 'coze-coding-dev-sdk';
import { normalizeOrganizationName } from './data-utils';

/**
 * PDF模式处理器
 * 用于识别PDF中的表格数据并填充到Excel文件中
 */
export class PDFMatcher {
  private pdfFile: File | Buffer;
  private excelFile: File | Buffer;
  private targetWorkbook: ExcelJS.Workbook;
  private sourceSheet单体: ExcelJS.Worksheet | null = null;
  private sourceSheet集团: ExcelJS.Worksheet | null = null;
  private llmClient: LLMClient;
  private customHeaders: Record<string, string>;

  // PDF识别结果
  private pdfData: Array<{
    orgName: string;
    creditTypes: Array<{
      type: string;
      amount: number;
    }>;
  }> = [];

  // 匹配结果
  private mappings: Array<{
    orgName: string;
    matchedOrgName?: string;
    targetRowIndex?: number;
    sourceSheet?: '单体' | '集团';
    creditTypes: Array<{
      type: string;
      amount: number;
      colIndex?: number;
      filled: boolean;
    }>;
  }> = [];

  constructor(
    pdfFile: File | Buffer,
    excelFile: File | Buffer,
    customHeaders: Record<string, string> = {}
  ) {
    this.pdfFile = pdfFile;
    this.excelFile = excelFile;
    this.targetWorkbook = new ExcelJS.Workbook();
    const config = new Config();
    this.llmClient = new LLMClient(config, customHeaders);
    this.customHeaders = customHeaders;
  }

  /**
   * 主处理方法
   */
  async process(): Promise<{ workbook: ExcelJS.Workbook; statistics: any }> {
    console.log('=== 开始PDF模式处理 ===');

    // 1. 加载Excel文件
    await this.loadExcelFile();

    // 2. 识别PDF表格
    await this.recognizePDFTable();

    // 3. 匹配机构和授信品种
    this.matchOrganizationsAndCreditTypes();

    // 4. 填充金额并标记红色
    this.fillAmountsWithRedMark();

    // 5. 删除多余的授信品种数据
    this.removeExtraCreditTypes();

    // 6. 统计结果
    const statistics = {
      totalOrganizations: this.mappings.length,
      matchedCount: this.mappings.filter(m => m.targetRowIndex).length,
      unmatchedCount: this.mappings.filter(m => !m.targetRowIndex).length,
    };

    console.log('=== PDF模式处理完成 ===');
    console.log('总机构数:', statistics.totalOrganizations);
    console.log('匹配成功:', statistics.matchedCount);
    console.log('匹配失败:', statistics.unmatchedCount);

    return {
      workbook: this.targetWorkbook,
      statistics,
    };
  }

  /**
   * 加载Excel文件
   */
  private async loadExcelFile(): Promise<void> {
    console.log('加载Excel文件...');

    // 获取Buffer
    let buffer: Buffer;
    if (this.excelFile instanceof File) {
      const arrayBuffer = await this.excelFile.arrayBuffer();
      buffer = Buffer.from(arrayBuffer);
    } else {
      buffer = this.excelFile;
    }

    await this.targetWorkbook.xlsx.load(buffer as any);

    // 查找单体表和集团表
    this.targetWorkbook.eachSheet((worksheet, sheetId) => {
      const sheetName = worksheet.name.trim();
      if (sheetName === '单体' || sheetName === '单体表') {
        this.sourceSheet单体 = worksheet;
        console.log(`找到单体表: ${worksheet.name}`);
      } else if (sheetName === '集团' || sheetName === '集团 ' || sheetName === '集团表') {
        this.sourceSheet集团 = worksheet;
        console.log(`找到集团表: ${worksheet.name}`);
      }
    });

    if (!this.sourceSheet单体 && !this.sourceSheet集团) {
      throw new Error('Excel文件中未找到"单体"或"集团"工作表');
    }
  }

  /**
   * 识别PDF表格
   */
  private async recognizePDFTable(): Promise<void> {
    console.log('识别PDF表格...');

    // 将PDF转换为Base64
    let pdfBase64: string;
    let mimeType: string;

    if (this.pdfFile instanceof File) {
      const arrayBuffer = await this.pdfFile.arrayBuffer();
      pdfBase64 = Buffer.from(arrayBuffer).toString('base64');
      mimeType = this.pdfFile.type || 'application/pdf';
    } else {
      pdfBase64 = this.pdfFile.toString('base64');
      mimeType = 'application/pdf';
    }

    const dataUri = `data:${mimeType};base64,${pdfBase64}`;

    // 使用Vision LLM识别表格
    const messages = [
      {
        role: 'user' as const,
        content: [
          {
            type: 'text' as const,
            text: `请识别这个PDF文档中的表格内容。
这是一个扫描版的PDF，内容是一个WORD文档中的表格。

请提取表格中的以下信息：
- 第3列：机构名称
- 第5列：申请授信品种及金额

请以JSON格式返回，格式如下：
{
  "tableData": [
    {
      "orgName": "机构名称",
      "creditTypes": [
        {"type": "授信品种名称", "amount": 金额数字}
      ]
    }
  ]
}

注意：
1. 如果第5列包含多个授信品种和金额，请全部提取
2. 金额请提取数字部分，不要包含单位
3. 如果某个单元格为空或无法识别，请跳过该行
4. 只返回JSON数据，不要有其他说明文字`,
          },
          {
            type: 'image_url' as const,
            image_url: {
              url: dataUri,
              detail: 'high' as const,
            },
          },
        ],
      },
    ];

    const response = await this.llmClient.invoke(messages, {
      model: 'doubao-seed-1-6-vision-250815',
      temperature: 0.1,
    });

    console.log('PDF识别结果:', response.content);

    // 解析JSON结果
    try {
      // 尝试提取JSON部分
      let jsonStr = response.content;
      const jsonMatch = jsonStr.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        jsonStr = jsonMatch[0];
      }

      const result = JSON.parse(jsonStr);
      this.pdfData = result.tableData || [];

      console.log(`成功识别 ${this.pdfData.length} 个机构的数据`);
      this.pdfData.forEach((item, index) => {
        console.log(`  ${index + 1}. ${item.orgName}: ${item.creditTypes.length}个授信品种`);
        item.creditTypes.forEach(ct => {
          console.log(`     - ${ct.type}: ${ct.amount}`);
        });
      });
    } catch (error) {
      console.error('解析PDF识别结果失败:', error);
      throw new Error('PDF表格识别结果解析失败，请检查PDF文件格式');
    }
  }

  /**
   * 匹配机构和授信品种
   */
  private matchOrganizationsAndCreditTypes(): void {
    console.log('\\n=== 开始匹配机构和授信品种 ===');

    // 初始化mapping
    for (const pdfItem of this.pdfData) {
      const mapping = {
        orgName: pdfItem.orgName,
        creditTypes: pdfItem.creditTypes.map(ct => ({
          type: ct.type,
          amount: ct.amount,
          filled: false,
        })),
      };
      this.mappings.push(mapping);
    }

    // 遍历每个机构进行匹配
    for (const mapping of this.mappings) {
      const normalizedOrgName = normalizeOrganizationName(mapping.orgName);
      let found = false;

      // 1. 先在单体表中匹配
      if (this.sourceSheet单体) {
        const result = this.findOrgInSheet(
          this.sourceSheet单体,
          normalizedOrgName,
          '单体',
          'B' // 单体表机构在B列
        );
        if (result) {
          mapping.matchedOrgName = result.orgName;
          mapping.targetRowIndex = result.rowIndex;
          mapping.sourceSheet = '单体';
          found = true;
          console.log(`匹配成功(单体): ${mapping.orgName} -> ${result.orgName} (行${result.rowIndex})`);
        }
      }

      // 2. 如果单体表未匹配，在集团表中匹配
      if (!found && this.sourceSheet集团) {
        const result = this.findOrgInSheet(
          this.sourceSheet集团,
          normalizedOrgName,
          '集团',
          'D' // 集团表机构在D列
        );
        if (result) {
          mapping.matchedOrgName = result.orgName;
          mapping.targetRowIndex = result.rowIndex;
          mapping.sourceSheet = '集团';
          found = true;
          console.log(`匹配成功(集团): ${mapping.orgName} -> ${result.orgName} (行${result.rowIndex})`);
        }
      }

      if (!found) {
        console.log(`匹配失败: ${mapping.orgName}`);
      }

      // 3. 匹配授信品种列
      if (mapping.targetRowIndex) {
        const sheet = mapping.sourceSheet === '集团' ? this.sourceSheet集团 : this.sourceSheet单体;
        if (sheet) {
          this.matchCreditTypeColumns(mapping, sheet);
        }
      }
    }
  }

  /**
   * 在工作表中查找机构
   */
  private findOrgInSheet(
    sheet: ExcelJS.Worksheet,
    normalizedOrgName: string,
    sheetType: '单体' | '集团',
    orgColumn: 'B' | 'D'
  ): { orgName: string; rowIndex: number } | null {
    const orgColIndex = orgColumn === 'B' ? 2 : 4;

    for (let row = 4; row <= sheet.rowCount; row++) {
      const cell = sheet.getCell(row, orgColIndex);
      const cellValue = String(cell.value || '').trim();
      
      if (cellValue) {
        const normalizedCellValue = normalizeOrganizationName(cellValue);
        if (normalizedCellValue === normalizedOrgName || 
            normalizedCellValue.includes(normalizedOrgName) ||
            normalizedOrgName.includes(normalizedCellValue)) {
          return { orgName: cellValue, rowIndex: row };
        }
      }
    }

    return null;
  }

  /**
   * 匹配授信品种列
   */
  private matchCreditTypeColumns(
    mapping: any,
    sheet: ExcelJS.Worksheet
  ): void {
    // 获取第3行的授信品种字段名
    const headerRow = sheet.getRow(3);
    const creditTypeMap = new Map<string, number>();

    // 从E列开始遍历（列索引5）
    for (let col = 5; col <= 50; col++) {
      const cell = headerRow.getCell(col);
      const value = String(cell.value || '').trim();
      if (value) {
        creditTypeMap.set(value, col);
      }
    }

    console.log(`  工作表中的授信品种:`, Array.from(creditTypeMap.keys()));

    // 匹配每个授信品种
    for (const ct of mapping.creditTypes) {
      // 尝试精确匹配
      if (creditTypeMap.has(ct.type)) {
        ct.colIndex = creditTypeMap.get(ct.type);
        console.log(`    授信品种匹配成功: ${ct.type} -> 列${ct.colIndex}`);
      } else {
        // 尝试模糊匹配
        for (const [header, colIndex] of creditTypeMap.entries()) {
          if (header.includes(ct.type) || ct.type.includes(header)) {
            ct.colIndex = colIndex;
            console.log(`    授信品种模糊匹配: ${ct.type} -> ${header} (列${colIndex})`);
            break;
          }
        }
      }

      if (!ct.colIndex) {
        console.log(`    授信品种未匹配: ${ct.type}`);
      }
    }
  }

  /**
   * 填充金额并标记红色
   */
  private fillAmountsWithRedMark(): void {
    console.log('\\n=== 开始填充金额 ===');

    for (const mapping of this.mappings) {
      if (!mapping.targetRowIndex || !mapping.sourceSheet) continue;

      const sheet = mapping.sourceSheet === '集团' ? this.sourceSheet集团 : this.sourceSheet单体;
      if (!sheet) continue;

      console.log(`\\n填充机构 ${mapping.orgName} (行${mapping.targetRowIndex}):`);

      for (const ct of mapping.creditTypes) {
        if (!ct.colIndex) continue;

        const cell = sheet.getCell(mapping.targetRowIndex, ct.colIndex);
        const oldValue = cell.value;
        
        // 填充新金额
        cell.value = ct.amount;
        
        // 设置红色字体
        cell.font = {
          color: { argb: 'FFFF0000' },
          bold: true,
        };

        ct.filled = true;
        console.log(`  ${ct.type} (列${ct.colIndex}): ${oldValue || '(空)'} -> ${ct.amount} [红色标记]`);
      }
    }
  }

  /**
   * 删除多余的授信品种数据
   */
  private removeExtraCreditTypes(): void {
    console.log('\\n=== 开始删除多余的授信品种数据 ===');

    // 收集每个机构应该保留的授信品种列
    const orgCreditTypes = new Map<string, Set<number>>();

    for (const mapping of this.mappings) {
      if (!mapping.targetRowIndex) continue;

      const key = `${mapping.sourceSheet}-${mapping.targetRowIndex}`;
      if (!orgCreditTypes.has(key)) {
        orgCreditTypes.set(key, new Set());
      }

      // 添加该机构在PDF中提到的授信品种列
      for (const ct of mapping.creditTypes) {
        if (ct.colIndex) {
          orgCreditTypes.get(key)!.add(ct.colIndex);
        }
      }
    }

    // 删除多余的授信品种数据
    for (const mapping of this.mappings) {
      if (!mapping.targetRowIndex || !mapping.sourceSheet) continue;

      const sheet = mapping.sourceSheet === '集团' ? this.sourceSheet集团 : this.sourceSheet单体;
      if (!sheet) continue;

      const key = `${mapping.sourceSheet}-${mapping.targetRowIndex}`;
      const allowedCols = orgCreditTypes.get(key);

      // 遍历该行的所有授信品种列（从E列开始）
      for (let col = 5; col <= 50; col++) {
        const cell = sheet.getCell(mapping.targetRowIndex, col);
        
        // 如果该列有值，但不在允许列表中，则删除
        if (cell.value !== null && cell.value !== undefined && 
            allowedCols && !allowedCols.has(col)) {
          const oldValue = cell.value;
          cell.value = null;
          console.log(`  删除多余数据: ${mapping.orgName} 列${col} (${oldValue})`);
        }
      }
    }
  }
}
