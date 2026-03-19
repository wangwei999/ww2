import ExcelJS from 'exceljs';
import { LLMClient, Config, HeaderUtils } from 'coze-coding-dev-sdk';
import { normalizeOrganizationName } from './data-utils';
import path from 'path';
import fs from 'fs';

// 动态导入的类型
type PDFDocument = any;
type PDFPage = any;
type PDFJSLib = {
  getDocument: (options: { data: Uint8Array; useSystemFonts?: boolean }) => { promise: Promise<PDFDocument> };
  GlobalWorkerOptions: { workerSrc: string };
};

// Canvas 相关类型
type CanvasType = any;
type CanvasContext = any;

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

    // 加载后立即清理所有共享公式
    this.cleanupAllSharedFormulas();
  }

  /**
   * 清理所有共享公式（在加载后立即执行）
   */
  private cleanupAllSharedFormulas(): void {
    console.log('\\n=== 清理所有共享公式 ===');

    this.targetWorkbook.eachSheet((worksheet, sheetId) => {
      let cleanedCount = 0;
      worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
          try {
            const cellData = cell as any;
            // 检查是否有共享公式
            if (cellData.sharedFormula) {
              // 获取公式的计算结果
              const result = cellData.result;
              // 清除共享公式引用，直接使用计算结果
              if (result !== undefined && result !== null) {
                cell.value = result;
              } else {
                cell.value = null;
              }
              cleanedCount++;
            }
          } catch (e) {
            // 忽略错误
          }
        });
      });
      if (cleanedCount > 0) {
        console.log(`  工作表 "${worksheet.name}" 清理了 ${cleanedCount} 个共享公式`);
      }
    });

    console.log('共享公式清理完成');
  }

  /**
   * 识别PDF表格
   */
  private async recognizePDFTable(): Promise<void> {
    console.log('识别PDF表格...');

    // 将PDF转换为Buffer
    let pdfBuffer: Buffer;

    if (this.pdfFile instanceof File) {
      const arrayBuffer = await this.pdfFile.arrayBuffer();
      pdfBuffer = Buffer.from(arrayBuffer);
    } else {
      pdfBuffer = this.pdfFile;
    }

    console.log('正在将PDF转换为图片...');

    try {
      // 动态导入 pdfjs-dist 和 canvas（避免 Next.js webpack 兼容性问题）
      const pdfjsLib = await import('pdfjs-dist');
      const canvasModule = await import('canvas');
      const createCanvas = canvasModule.createCanvas;

      // 使用 pdfjs-dist 加载 PDF
      const loadingTask = pdfjsLib.getDocument({
        data: new Uint8Array(pdfBuffer),
        useSystemFonts: true,
      });

      const pdfDocument = await loadingTask.promise;
      const numPages = pdfDocument.numPages;
      console.log(`PDF 共 ${numPages} 页`);

      // 遍历每一页
      for (let pageNum = 1; pageNum <= Math.min(numPages, 20); pageNum++) {
        console.log(`正在处理第 ${pageNum} 页...`);

        try {
          const page = await pdfDocument.getPage(pageNum);
          const scale = 2; // 高分辨率
          const viewport = page.getViewport({ scale });

          // 创建 canvas
          const canvas = createCanvas(viewport.width, viewport.height);
          const context = canvas.getContext('2d');

          // 渲染页面（使用 any 绕过类型检查，因为 node-canvas 与 HTMLCanvasElement 类型不兼容）
          const renderContext: any = {
            canvasContext: context,
            viewport: viewport,
            canvas: canvas,
          };

          await page.render(renderContext).promise;

          // 转换为 PNG Buffer
          const imageBuffer = canvas.toBuffer('image/png');
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
    } catch (error) {
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
        
        // 填充新金额
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
          
          // 检查是否包含公式
          try {
            const cellData = cell as any;
            if (cellData.formula || cellData.sharedFormula) {
              console.log(`  删除多余数据(含公式): ${mapping.orgName} 列${col}`);
            }
          } catch (e) {
            // 忽略错误
          }
          
          // 直接赋值null会自动清除公式
          cell.value = null;
          console.log(`  删除多余数据: ${mapping.orgName} 列${col} (${oldValue})`);
        }
      }
    }
  }

  /**
   * 创建干净的工作簿（保留样式和公式）
   */
  private async createCleanWorkbook(): Promise<ExcelJS.Workbook> {
    console.log('\\n=== 创建干净的工作簿 ===');

    const newWorkbook = new ExcelJS.Workbook();

    // 复制所有工作表
    this.targetWorkbook.eachSheet((sourceSheet, sheetId) => {
      const newSheet = newWorkbook.addWorksheet(sourceSheet.name);

      // 复制列宽
      sourceSheet.columns.forEach((col, index) => {
        if (col.width) {
          newSheet.getColumn(index + 1).width = col.width;
        }
      });

      // 获取实际使用的行列范围
      const maxRow = sourceSheet.rowCount || 200;
      const maxCol = sourceSheet.columnCount || 50;

      // 复制所有单元格（包括空单元格，保留边框）
      for (let rowNumber = 1; rowNumber <= maxRow; rowNumber++) {
        const sourceRow = sourceSheet.getRow(rowNumber);
        const newRow = newSheet.getRow(rowNumber);
        newRow.height = sourceRow.height;

        for (let colNumber = 1; colNumber <= maxCol; colNumber++) {
          const sourceCell = sourceRow.getCell(colNumber);
          const newCell = newRow.getCell(colNumber);

          // 复制值和公式（保留公式，不用结果替换）
          try {
            const cellData = sourceCell as any;
            
            // 检查是否有公式（保留公式）- 使用try-catch保护
            let hasFormula = false;
            let hasSharedFormula = false;
            let formulaValue = null;
            let sharedFormulaValue = null;
            let resultValue = null;

            try {
              if (cellData.formula) {
                hasFormula = true;
                formulaValue = cellData.formula;
              }
            } catch (e) {
              // formula属性访问失败，忽略
            }

            try {
              if (cellData.sharedFormula) {
                hasSharedFormula = true;
                sharedFormulaValue = cellData.sharedFormula;
                resultValue = cellData.result;
              }
            } catch (e) {
              // sharedFormula属性访问失败，忽略
            }

            if (hasFormula && formulaValue) {
              // 有公式，复制公式
              newCell.value = { formula: formulaValue };
            } else if (hasSharedFormula && sharedFormulaValue) {
              // 有共享公式，复制共享公式
              if (resultValue !== undefined && resultValue !== null) {
                newCell.value = { 
                  sharedFormula: sharedFormulaValue, 
                  result: resultValue 
                };
              } else {
                newCell.value = { sharedFormula: sharedFormulaValue };
              }
            } else {
              // 无公式，复制值
              newCell.value = sourceCell.value;
            }
          } catch (e) {
            // 如果公式处理失败，直接复制值
            newCell.value = sourceCell.value;
          }

          // 复制样式（包括边框）
          try {
            if (sourceCell.style) {
              // 使用model来复制样式
              const styleModel = (sourceCell as any).model;
              if (styleModel && styleModel.style) {
                newCell.style = { ...styleModel.style };
              } else {
                newCell.style = JSON.parse(JSON.stringify(sourceCell.style));
              }
            }
            if (sourceCell.font) {
              newCell.font = JSON.parse(JSON.stringify(sourceCell.font));
            }
            if (sourceCell.fill) {
              newCell.fill = JSON.parse(JSON.stringify(sourceCell.fill));
            }
            if (sourceCell.border) {
              newCell.border = JSON.parse(JSON.stringify(sourceCell.border));
            } else {
              // 如果原单元格没有边框，但有数据的单元格有边框，则设置默认边框
              // 检查同行其他单元格是否有边框
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
            // 样式复制失败，忽略
          }
        }
      }

      // 复制合并单元格
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

      console.log(`  复制工作表: ${sourceSheet.name}`);
    });

    console.log('干净工作簿创建完成');
    return newWorkbook;
  }
}
