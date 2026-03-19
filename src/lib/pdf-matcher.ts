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

    // 1. 加载Excel文件
    await this.loadExcelFile();

    // 2. 识别PDF表格
    await this.recognizePDFTable();

    // 3. 匹配机构和授信品种
    this.matchOrganizationsAndCreditTypes();

    // 4. 填充金额
    this.fillAmountsWithRedMark();

    // 5. 删除多余的授信品种数据
    this.removeExtraCreditTypes();

    // 6. 计算汇总值（单体表D列、集团表E列）
    this.calculateSummaryValues();

    // 7. 清理授信品种列和E列共享公式，但保留C列公式
    // 注意：集团表C列有公式（如C4=E4+E5），必须保留
    this.cleanupFormulasSelective();

    // 8. 统计结果
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
    
    // 不再在这里清理公式，改为在最后统一修复共享公式引用
  }

  /**
   * 清理授信品种列的共享公式，保留汇总列的公式
   * 策略：只清理汇总列之后的授信品种列的公式，保留汇总列本身的公式
   */
  private cleanupCreditTypeFormulas(): void {
    console.log('\\n=== 清理授信品种列的共享公式 ===');

    this.targetWorkbook.eachSheet((worksheet, sheetId) => {
      const sheetName = worksheet.name.trim();
      const isDanTi = sheetName === '单体' || sheetName === '单体表';
      const isJiTuan = sheetName === '集团' || sheetName === '集团 ' || sheetName === '集团表';
      
      if (!isDanTi && !isJiTuan) return;

      // 确定汇总列位置
      // 单体表：D列（第4列）
      // 集团表：E列（第5列）
      const summaryCol = isDanTi ? 4 : 5;
      
      let cleanedCount = 0;
      let sharedFormulaRefCount = 0;

      // 先扫描整行，找出所有共享公式的引用关系
      const sharedFormulaMasters = new Map<string, number>(); // 共享公式ID -> 行号
      const sharedFormulaClones: Array<{ row: number; col: number; refId: string }> = [];
      
      worksheet.eachRow((row, rowNumber) => {
        for (let col = 1; col <= 50; col++) {
          const cell = row.getCell(col);
          try {
            const model = (cell as any).model;
            if (!model) continue;

            // 记录共享公式的主公式和克隆
            if (model.sharedFormula !== undefined) {
              if (typeof model.sharedFormula === 'string') {
                // 这是一个克隆，记录引用的ID
                sharedFormulaClones.push({ row: rowNumber, col, refId: model.sharedFormula });
              } else if (typeof model.sharedFormula === 'number') {
                // 这是一个主公式（ref = 数字）
                // 这里我们不知道ID，但可以记录位置
              }
              sharedFormulaRefCount++;
            }
            
            // 记录主公式
            if (model.formula !== undefined && model.sharedFormula === undefined) {
              // 这是一个主公式（有formula但没有sharedFormula）
            }
          } catch (e) {
            // 忽略错误
          }
        }
      });

      console.log(`  工作表 "${worksheet.name}": 发现 ${sharedFormulaRefCount} 个共享公式引用`);

      // 清理授信品种列的公式（汇总列之后）
      worksheet.eachRow((row, rowNumber) => {
        // 从汇总列之后开始清理授信品种列
        for (let col = summaryCol + 1; col <= 50; col++) {
          const cell = row.getCell(col);
          try {
            const model = (cell as any).model;
            if (!model) continue;

            // 检查是否有公式
            if (model.sharedFormula !== undefined || model.formula !== undefined) {
              // 获取公式结果值
              let result = null;
              try {
                result = (cell as any).result;
              } catch (e) {
                try {
                  result = model.result;
                } catch (e2) {
                  result = null;
                }
              }

              // 转换为值
              const value = result !== undefined && result !== null ? result : null;
              
              // 同时清理 model 和 cell.value
              model.value = value;
              delete model.formula;
              delete model.sharedFormula;
              if (model.result !== undefined) delete model.result;
              
              // 重新设置cell的值
              cell.value = value;
              
              cleanedCount++;
              console.log(`    清理: 行${rowNumber} 列${col} 公式已转为值: ${value}`);
            }
          } catch (e) {
            console.warn(`    清理失败: 行${rowNumber} 列${col}`, (e as Error).message);
          }
        }
      });

      if (cleanedCount > 0) {
        console.log(`  工作表 "${worksheet.name}" 清理了 ${cleanedCount} 个授信品种列公式`);
      }
    });

    console.log('授信品种列公式清理完成');
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
      // 检查pdftoppm是否可用
      try {
        execSync('which pdftoppm', { timeout: 5000 });
      } catch (e) {
        throw new Error('系统工具pdftoppm未安装，请联系管理员安装poppler-utils');
      }

      // 使用 pdftoppm 将 PDF 转换为 PNG 图片
      const outputPrefix = `${sessionDir}/page`;
      console.log(`执行命令: pdftoppm -png -r 200 "${pdfPath}" "${outputPrefix}"`);
      
      try {
        const result = execSync(`pdftoppm -png -r 200 "${pdfPath}" "${outputPrefix}"`, {
          timeout: 60000, // 60秒超时
          encoding: 'utf-8',
        });
        console.log('pdftoppm命令执行成功');
      } catch (cmdError: any) {
        console.error('pdftoppm命令执行失败:', cmdError.message);
        console.error('错误详情:', cmdError.stderr || cmdError.stdout);
        throw new Error(`PDF转换失败: ${cmdError.message}`);
      }

      // 获取生成的图片文件列表
      const imageFiles = fs.readdirSync(sessionDir)
        .filter(f => f.endsWith('.png'))
        .sort()
        .map(f => `${sessionDir}/${f}`);

      console.log(`共转换 ${imageFiles.length} 页PDF`);

      if (imageFiles.length === 0) {
        throw new Error('无法从PDF中提取页面，可能是空PDF或格式不支持');
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

    } catch (error: any) {
      // 清理临时文件
      try {
        rmSync(sessionDir, { recursive: true, force: true });
      } catch (e) {
        // 忽略清理错误
      }
      
      console.error('PDF转换错误:', error);
      
      // 提供更详细的错误信息
      let errorMessage = 'PDF文件转换失败';
      
      if (error.message.includes('pdftoppm未安装')) {
        errorMessage = '系统工具pdftoppm未安装，请联系管理员安装poppler-utils';
      } else if (error.message.includes('PDF转换失败')) {
        errorMessage = error.message;
      } else if (error.message.includes('timeout')) {
        errorMessage = 'PDF转换超时，文件可能过大或损坏';
      } else if (error.message.includes('无法从PDF中提取')) {
        errorMessage = error.message;
      } else {
        errorMessage = `PDF文件转换失败: ${error.message}`;
      }
      
      throw new Error(errorMessage);
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
   * 注意：保护汇总列（单体表D列-第4列，集团表E列-第5列）
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

      // 确定汇总列的位置
      // 单体表：D列（第4列）是汇总列
      // 集团表：E列（第5列）是汇总列
      const summaryCol = mapping.sourceSheet === '单体' ? 4 : 5;

      // 从授信品种列开始（跳过汇总列）
      for (let col = 5; col <= 50; col++) {
        // 跳过汇总列
        if (col === summaryCol) {
          continue;
        }

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
   * 计算汇总值
   * 单体表：D列（第4列）= 授信品种列（E列开始）的金额总和
   * 集团表：
   *   - E列（第5列）= 授信品种列（F列开始）的金额总和
   *   - C列合并单元格 = 对应的E列单元格求和
   */
  private calculateSummaryValues(): void {
    console.log('\\n=== 计算汇总值 ===');

    // 处理单体表
    if (this.sourceSheet单体) {
      this.calculateDanTiSummary(this.sourceSheet单体);
    }

    // 处理集团表
    if (this.sourceSheet集团) {
      this.calculateJiTuanSummary(this.sourceSheet集团);
    }
  }

  /**
   * 计算单体表汇总值
   * D列 = E列开始的授信品种金额之和
   */
  private calculateDanTiSummary(sheet: ExcelJS.Worksheet): void {
    console.log('\\n计算单体表汇总值 (D列):');
    let calculatedCount = 0;

    for (const mapping of this.mappings) {
      if (!mapping.targetRowIndex || mapping.sourceSheet !== '单体') continue;

      const row = mapping.targetRowIndex;
      let sum = 0;
      let hasData = false;

      // 从E列（第5列）开始累加授信品种金额
      for (let col = 5; col <= 50; col++) {
        const cell = sheet.getCell(row, col);
        const value = cell.value;
        if (typeof value === 'number' && value !== 0) {
          sum += value;
          hasData = true;
        }
      }

      if (hasData) {
        const summaryCell = sheet.getCell(row, 4); // D列
        summaryCell.value = sum;
        calculatedCount++;
        console.log(`  行${row}: D列 = ${sum}`);
      }
    }

    console.log(`单体表汇总计算完成，共 ${calculatedCount} 行`);
  }

  /**
   * 计算集团表汇总值
   * E列 = F列开始的授信品种金额之和
   * C列合并单元格 = 对应的E列单元格求和
   */
  private calculateJiTuanSummary(sheet: ExcelJS.Worksheet): void {
    console.log('\\n计算集团表汇总值:');

    // 第一步：计算每行E列的值（F列开始的授信品种之和）
    console.log('\\n1. 计算E列值:');
    const eColumnValues = new Map<number, number>(); // 行号 -> E列值

    for (const mapping of this.mappings) {
      if (!mapping.targetRowIndex || mapping.sourceSheet !== '集团') continue;

      const row = mapping.targetRowIndex;
      let sum = 0;
      let hasData = false;

      // 从F列（第6列）开始累加授信品种金额
      for (let col = 6; col <= 50; col++) {
        const cell = sheet.getCell(row, col);
        const value = cell.value;
        if (typeof value === 'number' && value !== 0) {
          sum += value;
          hasData = true;
        }
      }

      if (hasData) {
        eColumnValues.set(row, sum);
        const eCell = sheet.getCell(row, 5); // E列
        eCell.value = sum;
        console.log(`  行${row}: E列 = ${sum}`);
      }
    }

    // 第二步：为C列合并单元格设置公式
    console.log('\\n2. 为C列合并单元格设置公式:');
    
    // 获取合并单元格信息
    const model = sheet.model;
    if (model && model.merges) {
      for (const merge of model.merges) {
        // 解析合并范围，格式如 "C4:C5"
        if (typeof merge === 'string' && merge.startsWith('C')) {
          const parts = merge.split(':');
          if (parts.length === 2) {
            const startMatch = parts[0].match(/C(\d+)/);
            const endMatch = parts[1].match(/C(\d+)/);
            
            if (startMatch && endMatch) {
              const startRow = parseInt(startMatch[1]);
              const endRow = parseInt(endMatch[1]);
              
              // 构建公式：如 =E4+E5
              const formulaParts: string[] = [];
              for (let r = startRow; r <= endRow; r++) {
                formulaParts.push(`E${r}`);
              }
              const formula = formulaParts.join('+');
              
              // 为C列合并单元格设置公式
              const cCell = sheet.getCell(startRow, 3); // C列，用master行
              cCell.value = { formula: formula };
              console.log(`  C${startRow}:C${endRow} 公式 = =${formula}`);
            }
          }
        }
      }
    }

    console.log('集团表汇总计算完成');
  }

  /**
   * 清理共享公式（保留C列的普通公式）
   * 策略：只清理共享公式，保留普通公式
   */
  private cleanupFormulasSelective(): void {
    console.log('\\n=== 清理共享公式 ===');

    this.targetWorkbook.eachSheet((worksheet, sheetId) => {
      const sheetName = worksheet.name.trim();
      const isDanTi = sheetName === '单体' || sheetName === '单体表';
      const isJiTuan = sheetName === '集团' || sheetName === '集团 ' || sheetName === '集团表';
      
      if (!isDanTi && !isJiTuan) return;

      let clearedCount = 0;

      worksheet.eachRow((row, rowNumber) => {
        for (let col = 1; col <= 50; col++) {
          const cell = row.getCell(col);
          try {
            const model = (cell as any).model;
            if (!model) continue;

            // 只清理共享公式，保留普通公式
            const hasSharedFormula = model.sharedFormula !== undefined;
            
            if (hasSharedFormula) {
              // 获取当前值
              let value = null;
              try {
                if (typeof cell.value === 'number') {
                  value = cell.value;
                } else if ((cell as any).result !== undefined) {
                  value = (cell as any).result;
                } else if (model.result !== undefined) {
                  value = model.result;
                }
              } catch (e) {
                // 忽略
              }

              // 清理共享公式，转为值
              model.value = value;
              delete model.formula;
              delete model.sharedFormula;
              if (model.result !== undefined) delete model.result;
              cell.value = value;
              
              clearedCount++;
            }
          } catch (e) {
            // 忽略错误
          }
        }
      });

      if (clearedCount > 0) {
        console.log(`  工作表 "${worksheet.name}" 清理了 ${clearedCount} 个共享公式`);
      }
    });

    console.log('共享公式清理完成');
  }
}
