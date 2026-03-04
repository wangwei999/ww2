import * as ExcelJS from 'exceljs';
import { normalizeOrganizationName } from './data-utils';

/**
 * 机构映射信息
 */
export interface OrganizationMapping {
  targetRowIndex: number;
  sourceRowIndex: number;
  orgName: string;
  matched: boolean;
  valueC: any;
  valueD: number | null;
  valueN: number | null;
}

/**
 * 授信模式匹配器（使用exceljs保留原始格式）
 */
export class CreditMatcher {
  private sourceWorkbook: ExcelJS.Workbook;
  private targetWorkbook: ExcelJS.Workbook;
  private targetSheet: ExcelJS.Worksheet;
  private sourceSheet: ExcelJS.Worksheet;
  private mappings: OrganizationMapping[] = [];

  constructor(
    sourceWorkbook: ExcelJS.Workbook,
    targetWorkbook: ExcelJS.Workbook,
    targetSheetName?: string
  ) {
    this.sourceWorkbook = sourceWorkbook;
    this.targetWorkbook = targetWorkbook;
    
    const sheetName = targetSheetName || targetWorkbook.worksheets[0]?.name;
    if (!sheetName) {
      throw new Error('目标文件中没有工作表');
    }
    
    const targetSheet = targetWorkbook.getWorksheet(sheetName);
    if (!targetSheet) {
      throw new Error(`目标文件中未找到工作表：${sheetName}`);
    }
    this.targetSheet = targetSheet as ExcelJS.Worksheet;

    const sourceSheetName = '单体';
    const sourceSheet = sourceWorkbook.getWorksheet(sourceSheetName);
    if (!sourceSheet) {
      throw new Error(`源文件中未找到"${sourceSheetName}"工作表`);
    }
    this.sourceSheet = sourceSheet as ExcelJS.Worksheet;

    console.log('=== CreditMatcher 初始化（使用exceljs）===');
    console.log('源文件工作表:', sourceWorkbook.worksheets.map(ws => ws.name));
    console.log('目标文件工作表:', targetWorkbook.worksheets.map(ws => ws.name));
    console.log('目标工作表:', this.targetSheet.name);
    console.log('源工作表:', this.sourceSheet.name);
  }

  /**
   * 执行匹配并直接修改目标workbook
   */
  public async matchAndFill(): Promise<{
    workbook: ExcelJS.Workbook;
    mappings: OrganizationMapping[];
    statistics: {
      totalOrganizations: number;
      matchedCount: number;
      unmatchedCount: number;
    };
  }> {
    console.log('=== 开始授信模式匹配（使用exceljs保留格式）===');

    // 构建源文件的机构名称索引（B列，从第4行开始）
    const sourceOrgMap = new Map<string, number>();
    const sourceRowCount = this.sourceSheet.rowCount;
    
    console.log('源文件总行数:', sourceRowCount);
    
    for (let row = 4; row <= sourceRowCount; row++) {
      const cell = this.sourceSheet.getCell(row, 2);
      const orgName = String(cell.value || '').trim();
      
      if (orgName && orgName !== '机构名称') {
        const normalizedName = normalizeOrganizationName(orgName);
        sourceOrgMap.set(normalizedName, row);
        
        if (normalizedName !== orgName) {
          console.log(`源文件规范化: ${orgName} → ${normalizedName} (行${row})`);
        }
      }
    }

    console.log('源文件机构索引大小:', sourceOrgMap.size);

    // 遍历目标文件的机构（B列，从第6行开始）
    const targetRowCount = this.targetSheet.rowCount;
    console.log('目标文件总行数:', targetRowCount);

    for (let row = 6; row <= targetRowCount; row++) {
      const cell = this.targetSheet.getCell(row, 2);
      const orgName = String(cell.value || '').trim();
      
      if (orgName && orgName !== '机构名称') {
        const normalizedName = normalizeOrganizationName(orgName);
        
        console.log(`目标行${row}: ${orgName}`);
        console.log(`  规范化后: ${normalizedName}`);
        console.log(`  是否与原始相同: ${normalizedName === orgName}`);
        
        const sourceRowIndex = sourceOrgMap.get(normalizedName);
        console.log(`  源文件行号: ${sourceRowIndex !== undefined ? sourceRowIndex : '未找到'}`);

        const mapping: OrganizationMapping = {
          targetRowIndex: row,
          sourceRowIndex: -1,
          orgName,
          matched: false,
          valueC: null,
          valueD: null,
          valueN: null,
        };

        if (sourceRowIndex !== undefined) {
          mapping.matched = true;
          mapping.sourceRowIndex = sourceRowIndex;

          const cellC = this.sourceSheet.getCell(sourceRowIndex, 3);
          const cellD = this.sourceSheet.getCell(sourceRowIndex, 4);

          const matchInfo = normalizedName !== orgName 
            ? `(${orgName} → ${normalizedName})` 
            : orgName;
          
          console.log(`匹配成功: ${matchInfo}`);
          console.log(`  源C列type: ${cellC.type}, 源D列type: ${cellD.type}`);
          console.log(`  源C列value: ${JSON.stringify(cellC.value)}, 源D列value: ${JSON.stringify(cellD.value)}`);
          console.log(`  源C列完整对象: ${JSON.stringify({type: cellC.type, value: cellC.value, result: (cellC as any).result, formula: (cellC as any).formula, text: (cellC as any).text, numFmt: cellC.numFmt})}`);
          console.log(`  源D列完整对象: ${JSON.stringify({type: cellD.type, value: cellD.value, result: (cellD as any).result, formula: (cellD as any).formula, text: (cellD as any).text, numFmt: cellD.numFmt})}`);
          
          // 对于第一行匹配，额外输出源文件该行的所有信息
          if (this.mappings.length === 0) {
            console.log(`  === 第一行匹配，输出源文件行${sourceRowIndex}的所有信息 ===`);
            const sourceRow = this.sourceSheet.getRow(sourceRowIndex);
            console.log(`  源行${sourceRowIndex}单元格数: ${sourceRow.cellCount}`);
            for (let col = 1; col <= 15; col++) {
              const cell = sourceRow.getCell(col);
              console.log(`  列${col} (type=${cell.type}): value=${JSON.stringify(cell.value)}, text=${JSON.stringify((cell as any).text)}, numFmt=${JSON.stringify(cell.numFmt)}`);
            }
          }

          mapping.valueC = this.parseCellValue(cellC.value);
          mapping.valueD = this.parseCellValue(cellD.value);
          mapping.valueN = mapping.valueD;

          console.log(`  目标行: ${row} (B${row}), 源行: ${sourceRowIndex} (B${sourceRowIndex})`);
          console.log(`  解析后C列值: ${JSON.stringify(mapping.valueC)}, 解析后D列值: ${JSON.stringify(mapping.valueD)}, N列值: ${JSON.stringify(mapping.valueN)}`);

          // C列：直接使用源单元格的值（保留Date对象），并复制格式
          if (cellC.value !== null && cellC.value !== undefined) {
            const targetCellC = this.targetSheet.getCell(row, 3);
            targetCellC.value = cellC.value;
            // 使用style API设置日期格式
            targetCellC.style = {
              numFmt: cellC.numFmt
            };
            console.log(`  C列格式已设置: ${cellC.numFmt}, 值: ${JSON.stringify(cellC.value)}, 值类型: ${typeof cellC.value}, 是否为Date: ${cellC.value instanceof Date}`);
          }

          // D列：使用解析后的数值
          if (mapping.valueD !== null) {
            const targetCellD = this.targetSheet.getCell(row, 4);
            targetCellD.value = mapping.valueD;
            // 使用style API设置为通用格式
            targetCellD.style = {
              numFmt: 'General'
            };
            console.log(`  D列已填充值: ${mapping.valueD}`);
          }

          // N列：使用解析后的数值
          if (mapping.valueN !== null) {
            const targetCellN = this.targetSheet.getCell(row, 14);
            targetCellN.value = mapping.valueN;
            // 使用style API设置为通用格式
            targetCellN.style = {
              numFmt: 'General'
            };
            console.log(`  N列已填充值: ${mapping.valueN}`);
          }

          console.log(`已填充: C${row}=${mapping.valueC}, D${row}=${mapping.valueD}, N${row}=${mapping.valueN}`);
        } else {
          console.log(`匹配失败: ${orgName}`);
        }

        this.mappings.push(mapping);
      }
    }

    // 统计
    const statistics = {
      totalOrganizations: this.mappings.length,
      matchedCount: this.mappings.filter(m => m.matched).length,
      unmatchedCount: this.mappings.filter(m => !m.matched).length,
    };

    console.log('=== 匹配完成 ===');
    console.log('总机构数:', statistics.totalOrganizations);
    console.log('匹配成功:', statistics.matchedCount);
    console.log('匹配失败:', statistics.unmatchedCount);

    return {
      workbook: this.targetWorkbook,
      mappings: this.mappings,
      statistics,
    };
  }

  /**
   * 解析单元格值（处理公式对象和日期对象）
   */
  private parseCellValue(value: any): any {
    if (value === null || value === undefined || value === '') {
      return null;
    }

    // 如果是Date对象，直接返回（保留日期格式）
    if (value instanceof Date) {
      console.log(`    parseCellValue: 检测到Date对象 - ${value.toISOString()}`);
      return value;
    }

    // 如果是数字，直接返回
    if (typeof value === 'number') {
      return value;
    }

    // 处理公式结果对象
    if (typeof value === 'object' && value !== null && 'result' in value) {
      const result = (value as any).result;
      return result !== undefined && result !== null ? result : null;
    }

    // 尝试转换为数字
    const num = parseFloat(String(value).trim());
    return isNaN(num) ? null : num;
  }
}
