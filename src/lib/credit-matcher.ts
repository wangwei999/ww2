import * as ExcelJS from 'exceljs';

/**
 * 机构映射信息
 */
export interface OrganizationMapping {
  targetRowIndex: number;  // 目标文件中的行索引（从1开始，Excel行号）
  sourceRowIndex: number;  // 源文件中的行索引（从1开始，Excel行号）
  orgName: string;         // 机构名称
  matched: boolean;        // 是否匹配成功
  valueC: any;             // C列的值（原授信时间，可能是Date对象）
  valueD: number | null;   // D列的值（原授信-授信总额）
  valueN: number | null;   // N列的值（拟调整授信-授信总额）
}

/**
 * 授信模式匹配器（使用exceljs保留原始格式）
 *
 * 规则：
 * 1. B文件（目标）：第5行是字段行，B6开始是机构名称
 * 2. A文件（源）："单体"工作表，B4开始是机构名称
 * 3. 匹配成功后：
 *    - A文件C列X行 → B文件C列对应行
 *    - A文件D列X行 → B文件D列对应行
 *    - A文件D列X行 → B文件N列对应行
 *
 * 特点：使用exceljs直接修改B文件，保留原始格式（样式、合并单元格、公式等）
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
    
    // 默认使用第一个工作表，如果指定了工作表名则使用指定的工作表
    const sheetName = targetSheetName || targetWorkbook.worksheets[0]?.name;
    if (!sheetName) {
      throw new Error('目标文件中没有工作表');
    }
    
    const targetSheet = targetWorkbook.getWorksheet(sheetName);
    if (!targetSheet) {
      throw new Error(`目标文件中未找到工作表：${sheetName}`);
    }
    this.targetSheet = targetSheet as ExcelJS.Worksheet;

    // 获取源文件的"单体"工作表
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
      const cell = this.sourceSheet.getCell(row, 2); // B列（列索引为2）
      const orgName = String(cell.value || '').trim();
      
      if (orgName && orgName !== '机构名称') {
        sourceOrgMap.set(orgName, row);
      }
    }

    console.log('源文件机构索引大小:', sourceOrgMap.size);

    // 遍历目标文件的机构（B列，从第6行开始）
    const targetRowCount = this.targetSheet.rowCount;
    console.log('目标文件总行数:', targetRowCount);

    for (let row = 6; row <= targetRowCount; row++) {
      const cell = this.targetSheet.getCell(row, 2); // B列（列索引为2）
      const orgName = String(cell.value || '').trim();

      if (orgName && orgName !== '机构名称') {
        const sourceRowIndex = sourceOrgMap.get(orgName);

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

          // 从源文件获取值（保留原始类型）
          const cellC = this.sourceSheet.getCell(sourceRowIndex, 3); // C列
          const cellD = this.sourceSheet.getCell(sourceRowIndex, 4); // D列

          console.log(`  源C列type: ${cellC.type}, 源D列type: ${cellD.type}`);

          // 处理C列：保留日期对象（用于统计，不用于填充）
          mapping.valueC = this.parseCellValue(cellC.value);
          // 处理D列：可能是公式对象
          mapping.valueD = this.parseCellValue(cellD.value);
          mapping.valueN = mapping.valueD;

          console.log(`匹配成功: ${orgName}`);
          console.log(`  目标行: ${row} (B${row}), 源行: ${sourceRowIndex} (B${sourceRowIndex})`);
          console.log(`  C列值: ${mapping.valueC}, D列值: ${mapping.valueD}, N列值: ${mapping.valueN}`);

          // 直接修改目标workbook的单元格（exceljs会保留格式）
          // C列：保留原始日期对象（cellC.value是Date对象或Excel日期数字）
          const valueC = cellC.value;
          if (valueC !== null && valueC !== undefined) {
            const targetCellC = this.targetSheet.getCell(row, 3);
            targetCellC.value = valueC; // 使用原始值，保留日期类型
          }

          // D列：使用解析后的数值
          if (mapping.valueD !== null) {
            const targetCellD = this.targetSheet.getCell(row, 4);
            targetCellD.value = mapping.valueD;
          }

          // N列：使用解析后的数值
          if (mapping.valueN !== null) {
            const targetCellN = this.targetSheet.getCell(row, 14);
            targetCellN.value = mapping.valueN;
          }

          console.log(`已填充: C${row}=${valueC}, D${row}=${mapping.valueD}, N${row}=${mapping.valueN}`);
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
   * 更新目标workbook的单元格（exceljs保留格式）
   */
  private updateTargetCell(
    rowIndex: number,
    colIndex: number,
    value: any
  ): void {
    if (value === null || value === undefined) return;

    const cell = this.targetSheet.getCell(rowIndex, colIndex);
    cell.value = value;
    // exceljs会自动保留原有的格式、样式等
  }

  /**
   * 解析单元格值（处理公式对象和日期对象）
   */
  private parseCellValue(value: any): number | null {
    if (value === null || value === undefined || value === '') {
      return null;
    }

    // 如果是数字，直接返回
    if (typeof value === 'number') {
      return value;
    }

    // 如果是公式对象（{result: number, sharedFormula: string}），提取result
    if (typeof value === 'object' && value !== null && 'result' in value) {
      const result = (value as any).result;
      return typeof result === 'number' ? result : null;
    }

    // 如果是日期对象，转换为时间戳数字（用于统计）
    if (value instanceof Date) {
      return value.getTime();
    }

    // 尝试解析字符串为数字
    const num = parseFloat(String(value).trim());
    return isNaN(num) ? null : num;
  }
}
