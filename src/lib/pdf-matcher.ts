import ExcelJS from 'exceljs';
import { LLMClient, Config, HeaderUtils } from 'coze-coding-dev-sdk';
import { normalizeOrganizationName } from './data-utils';
import { execSync } from 'child_process';
import fs, { mkdirSync, rmSync, existsSync } from 'fs';
import path from 'path';

// 临时目录用于存储PDF转换的图片
const TEMP_DIR = '/tmp/pdf-images';

/**
 * PDF模式处理器
 * 用于识别PDF中的表格数据并填充到Excel文件中
 * 使用 pdftoppm (poppler-utils) 将 PDF 转换为图片
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

    // 1. 加载Excel文件（会自动清理共享公式）
    await this.loadExcelFile();

    // 2. 识别PDF表格
    await this.recognizePDFTable();

    // 3. 匹配机构和授信品种
    this.matchOrganizationsAndCreditTypes();

    // 4. 填充金额
    this.fillAmountsWithRedMark();

    // 5. 删除多余的授信品种数据
    this.removeExtraCreditTypes();

    // 6. 创建新的干净工作簿（避免公式问题）
    const cleanWorkbook = await this.createCleanWorkbook();

    // 7. 统计结果
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
      workbook: cleanWorkbook,
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

    // 关键：加载后立即将所有共享公式转换为值，避免后续操作触发共享公式错误
    this.convertSharedFormulasToValues();
  }

  /**
   * 将所有共享公式转换为值（关键修复）
   * 这是解决"Shared Formula master must exist above and or left of clone"错误的根本方法
   */
  private convertSharedFormulasToValues(): void {
    console.log('\\n=== 将所有共享公式转换为值 ===');

    this.targetWorkbook.eachSheet((worksheet, sheetId) => {
      let convertedCount = 0;
      
      worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
          try {
            const cellData = cell as any;
            
            // 检查是否有共享公式
            if (cellData.sharedFormula) {
              // 获取共享公式的结果值
              const result = cellData.result;
              const value = result !== undefined && result !== null ? result : null;
              
              // 关键：清除共享公式的内部引用，避免后续writeBuffer时触发错误
              // 方法1: 直接操作model
              if (cellData.model) {
                cellData.model.value = value;
                delete cellData.model.formula;
                delete cellData.model.sharedFormula;
              }
              
              // 方法2: 覆盖value
              cell.value = value;
              
              convertedCount++;
            }
            // 也检查普通公式
            else if (cellData.formula) {
              const result = cellData.result;
              const value = result !== undefined && result !== null ? result : null;
              
              // 清除公式的内部引用
              if (cellData.model) {
                cellData.model.value = value;
                delete cellData.model.formula;
              }
              
              cell.value = value;
              convertedCount++;
            }
          } catch (e) {
            // 忽略错误，继续处理下一个单元格
            console.error(`转换单元格(${rowNumber}, ${colNumber})时出错:`, e);
          }
        });
      });
      
      if (convertedCount > 0) {
        console.log(`  工作表 "${worksheet.name}" 转换了 ${convertedCount} 个公式`);
      }
    });

    console.log('公式转换完成');
  }

  /**
   * 识别PDF表格
   */
  private async recognizePDFTable(): Promise<void> {
    console.log('识别PDF表格...');

    // 将PDF转换为Buffer并保存为临时文件
    let pdfBuffer: Buffer;
    if (this.pdfFile instanceof File) {
      const arrayBuffer = await this.pdfFile.arrayBuffer();
      pdfBuffer = Buffer.from(arrayBuffer);
    } else {
      pdfBuffer = this.pdfFile;
    }

    // 创建临时目录
    const sessionId = Date.now();
    const sessionDir = `${TEMP_DIR}/${sessionId}`;
    if (!existsSync(TEMP_DIR)) {
      mkdirSync(TEMP_DIR, { recursive: true });
    }
    mkdirSync(sessionDir, { recursive: true });

    // 保存PDF文件
    const pdfPath = `${sessionDir}/input.pdf`;
    fs.writeFileSync(pdfPath, pdfBuffer);

    console.log('正在将PDF转换为图片...');

    try {
      // 使用 pdftoppm 将 PDF 转换为 PNG 图片
      const outputPrefix = `${sessionDir}/page`;
      execSync(`pdftoppm -png -r 200 "${pdfPath}" "${outputPrefix}"`, {
        timeout: 60000, // 60秒超时
      });

      // 获取生成的图片文件列表
      const imageFiles = fs.readdirSync(sessionDir)
        .filter(f => f.endsWith('.png'))
        .sort()
        .map(f => `${sessionDir}/${f}`);

      console.log(`共转换 ${imageFiles.length} 页PDF`);

      if (imageFiles.length === 0) {
        throw new Error('无法从PDF中提取页面');
      }

      // 对每一页进行OCR识别
      for (let i = 0; i < imageFiles.length; i++) {
        const imagePath = imageFiles[i];
        const pageNum = i + 1;
        console.log(`正在处理第 ${pageNum} 页...`);

        try {
          const imageBuffer = fs.readFileSync(imagePath);
          const imageBase64 = imageBuffer.toString('base64');
          const dataUri = `data:image/png;base64,${imageBase64}`;

          // 使用Vision LLM识别表格
          const messages = [
            {
              role: 'user' as const,
              content: [
                {
                  type: 'text' as const,
                  text: `请识别这个图片中的表格内容。
这是一个扫描版PDF的页面，内容是一个WORD文档中的表格。

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
2. 金额请提取数字部分，不要包含单位（亿元、万元等）
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

          console.log(`第 ${pageNum} 页识别结果:`, response.content.substring(0, 200) + '...');

          // 解析JSON结果
          try {
            let jsonStr = response.content;
            const jsonMatch = jsonStr.match(/\{[\s\S]*\}/);
            if (jsonMatch) {
              jsonStr = jsonMatch[0];
            }

            const result = JSON.parse(jsonStr);
            if (result.tableData && Array.isArray(result.tableData)) {
              this.pdfData.push(...result.tableData);
            }
          } catch (error) {
            console.error(`第 ${pageNum} 页JSON解析失败:`, error);
          }
        } catch (error) {
          console.error(`第 ${pageNum} 页处理失败:`, error);
        }
      }

      // 清理临时文件
      try {
        rmSync(sessionDir, { recursive: true, force: true });
      } catch (e) {
        console.warn('清理临时文件失败:', e);
      }

    } catch (error) {
      // 清理临时文件
      try {
        rmSync(sessionDir, { recursive: true, force: true });
      } catch (e) {
        // 忽略清理错误
      }
      console.error('PDF转换错误:', error);
      throw new Error('PDF文件转换失败，请确保PDF文件格式正确');
    }

    if (this.pdfData.length === 0) {
      throw new Error('无法从PDF中提取任何数据，请检查PDF文件内容');
    }

    console.log(`成功识别 ${this.pdfData.length} 个机构的数据`);
    this.pdfData.forEach((item, index) => {
      console.log(`  ${index + 1}. ${item.orgName}: ${item.creditTypes.length}个授信品种`);
      item.creditTypes.forEach(ct => {
        console.log(`     - ${ct.type}: ${ct.amount}`);
      });
    });
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
          'B'
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
          'D'
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
  private matchCreditTypeColumns(mapping: any, sheet: ExcelJS.Worksheet): void {
    const headerRow = sheet.getRow(3);
    const creditTypeMap = new Map<string, number>();

    for (let col = 5; col <= 50; col++) {
      const cell = headerRow.getCell(col);
      const value = String(cell.value || '').trim();
      if (value) {
        creditTypeMap.set(value, col);
      }
    }

    console.log(`  工作表中的授信品种:`, Array.from(creditTypeMap.keys()));

    for (const ct of mapping.creditTypes) {
      if (creditTypeMap.has(ct.type)) {
        ct.colIndex = creditTypeMap.get(ct.type);
        console.log(`    授信品种匹配成功: ${ct.type} -> 列${ct.colIndex}`);
      } else {
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
   * 填充金额
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
        
        // 设置新值
        cell.value = ct.amount;

        ct.filled = true;
        console.log(`  ${ct.type} (列${ct.colIndex}): ${oldValue ?? '(空)'} -> ${ct.amount}`);
      }
    }
  }

  /**
   * 删除多余的授信品种数据
   */
  private removeExtraCreditTypes(): void {
    console.log('\\n=== 开始删除多余的授信品种数据 ===');

    const orgCreditTypes = new Map<string, Set<number>>();

    for (const mapping of this.mappings) {
      if (!mapping.targetRowIndex) continue;

      const key = `${mapping.sourceSheet}-${mapping.targetRowIndex}`;
      if (!orgCreditTypes.has(key)) {
        orgCreditTypes.set(key, new Set());
      }

      for (const ct of mapping.creditTypes) {
        if (ct.colIndex) {
          orgCreditTypes.get(key)!.add(ct.colIndex);
        }
      }
    }

    for (const mapping of this.mappings) {
      if (!mapping.targetRowIndex || !mapping.sourceSheet) continue;

      const sheet = mapping.sourceSheet === '集团' ? this.sourceSheet集团 : this.sourceSheet单体;
      if (!sheet) continue;

      const key = `${mapping.sourceSheet}-${mapping.targetRowIndex}`;
      const allowedCols = orgCreditTypes.get(key);

      for (let col = 5; col <= 50; col++) {
        const cell = sheet.getCell(mapping.targetRowIndex, col);
        
        if (cell.value !== null && cell.value !== undefined && 
            allowedCols && !allowedCols.has(col)) {
          const oldValue = cell.value;
          cell.value = null;
          console.log(`  删除多余数据: ${mapping.orgName} 列${col} (${oldValue})`);
        }
      }
    }
  }

  /**
   * 创建干净的工作簿（完全重新构建，避免任何公式引用）
   */
  private async createCleanWorkbook(): Promise<ExcelJS.Workbook> {
    console.log('\\n=== 创建干净的工作簿 ===');

    const newWorkbook = new ExcelJS.Workbook();

    this.targetWorkbook.eachSheet((sourceSheet, sheetId) => {
      const newSheet = newWorkbook.addWorksheet(sourceSheet.name);

      // 复制列宽
      sourceSheet.columns.forEach((col, index) => {
        if (col.width) {
          newSheet.getColumn(index + 1).width = col.width;
        }
      });

      const maxRow = sourceSheet.rowCount || 200;
      const maxCol = sourceSheet.columnCount || 50;

      // 按行按列复制数据
      for (let rowNumber = 1; rowNumber <= maxRow; rowNumber++) {
        const sourceRow = sourceSheet.getRow(rowNumber);
        const newRow = newSheet.getRow(rowNumber);
        newRow.height = sourceRow.height;

        for (let colNumber = 1; colNumber <= maxCol; colNumber++) {
          const sourceCell = sourceRow.getCell(colNumber);
          const newCell = newRow.getCell(colNumber);

          // 获取单元格的值（优先使用结果值）
          let cellValue: any = null;
          try {
            const cellData = sourceCell as any;
            
            // 如果有公式，使用结果值
            if (cellData.sharedFormula || cellData.formula) {
              const result = cellData.result;
              cellValue = result !== undefined && result !== null ? result : null;
            } else {
              // 否则直接使用value
              cellValue = sourceCell.value;
            }
          } catch (e) {
            cellValue = sourceCell.value;
          }

          // 设置值
          newCell.value = cellValue;

          // 复制样式
          try {
            if (sourceCell.font) {
              newCell.font = JSON.parse(JSON.stringify(sourceCell.font));
            }
            if (sourceCell.fill) {
              newCell.fill = JSON.parse(JSON.stringify(sourceCell.fill));
            }
            if (sourceCell.border) {
              newCell.border = JSON.parse(JSON.stringify(sourceCell.border));
            } else {
              const firstDataCell = sourceRow.getCell(1);
              if (firstDataCell && firstDataCell.border) {
                newCell.border = JSON.parse(JSON.stringify(firstDataCell.border));
              }
            }
            if (sourceCell.alignment) {
              newCell.alignment = JSON.parse(JSON.stringify(sourceCell.alignment));
            }
            if (sourceCell.numFmt) {
              newCell.numFmt = sourceCell.numFmt;
            }
          } catch (e) {
            // 忽略样式复制错误
          }
        }
      }

      // 复制合并单元格
      try {
        const merges = (sourceSheet as any)._merges;
        if (merges) {
          Object.values(merges).forEach((merge: any) => {
            try {
              newSheet.mergeCells(merge);
            } catch (e) {
              // 忽略合并错误
            }
          });
        }
      } catch (e) {
        // 忽略合并单元格复制错误
      }

      console.log(`  复制工作表: ${sourceSheet.name}`);
    });

    console.log('干净工作簿创建完成');
    return newWorkbook;
  }
}
