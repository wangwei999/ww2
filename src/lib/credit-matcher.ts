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

    // 填充随机字段名称和数值
    this.fillRandomFields();

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

  /**
   * 填充随机字段名称和数值
   */
  private fillRandomFields(): void {
    console.log('=== 开始填充随机字段名称 ===');

    // 1. 获取A文件第3行（字段名）和第4-173行（数据）
    const headerRow = this.sourceSheet.getRow(3);
    const fieldNames: string[] = [];
    for (let col = 5; col <= 28; col++) { // E列(5)到AB列(28)
      const cell = headerRow.getCell(col);
      const value = String(cell.value || '').trim();
      if (value) {
        fieldNames.push(value);
      }
    }
    console.log(`A文件E3-AB3共找到 ${fieldNames.length} 个字段名`);

    // 2. 查找所有匹配成功的机构及其非零数值字段
    const matchedData: Array<{
      targetRow: number;
      orgName: string;
      sourceRow: number;
      nonZeroFields: Array<{ name: string; value: number; sourceCol: number }>;
    }> = [];

    for (const mapping of this.mappings) {
      if (mapping.matched && mapping.sourceRowIndex > 0) {
        const sourceRow = this.sourceSheet.getRow(mapping.sourceRowIndex);
        const nonZeroFields: Array<{ name: string; value: number; sourceCol: number }> = [];

        // 检查E列到AB列的数值
        for (let col = 5; col <= 28; col++) { // E列(5)到AB列(28)
          const cell = sourceRow.getCell(col);
          const value = this.parseCellValue(cell.value);
          
          if (value !== null && value !== 0) {
            const fieldName = fieldNames[col - 5] || `字段${col}`;
            nonZeroFields.push({
              name: fieldName,
              value: value,
              sourceCol: col
            });
          }
        }

        if (nonZeroFields.length > 0) {
          matchedData.push({
            targetRow: mapping.targetRowIndex,
            orgName: mapping.orgName,
            sourceRow: mapping.sourceRowIndex,
            nonZeroFields
          });
          console.log(`机构: ${mapping.orgName} (目标行${mapping.targetRowIndex}), 非零字段数: ${nonZeroFields.length}`);
        }
      }
    }

    console.log(`共找到 ${matchedData.length} 个有非零数值的机构`);

    // 3. 为每个匹配成功的机构填充随机字段
    for (const data of matchedData) {
      this.fillFieldsForOrganization(data, fieldNames);
    }

    console.log('=== 随机字段填充完成 ===');
  }

  /**
   * 为单个机构填充随机字段名称和数值
   */
  private fillFieldsForOrganization(
    data: {
      targetRow: number;
      orgName: string;
      sourceRow: number;
      nonZeroFields: Array<{ name: string; value: number; sourceCol: number }>;
    },
    fieldNames: string[]
  ): void {
    const { targetRow, nonZeroFields } = data;
    console.log(`\\n=== 为 ${data.orgName} (目标行${targetRow}) 填充随机字段 ===`);

    // B文件E5-M5列（5-13列）
    const targetColumns = [5, 6, 7, 8, 9, 10, 11, 12, 13]; // E-M列
    const totalSlots = targetColumns.length; // 9个单元格

    // 随机打乱列顺序
    const shuffledColumns = [...targetColumns].sort(() => Math.random() - 0.5);

    // 1. 选择非零字段名称（最多9个，但非零字段可能不足）
    const selectedNonZeroFields = nonZeroFields.slice(0, Math.min(nonZeroFields.length, totalSlots));
    console.log(`选择了 ${selectedNonZeroFields.length} 个非零字段`);

    // 2. 随机打乱非零字段
    const shuffledNonZeroFields = selectedNonZeroFields.sort(() => Math.random() - 0.5);

    // 3. 填充非零字段名称和数值
    for (let i = 0; i < shuffledNonZeroFields.length; i++) {
      const col = shuffledColumns[i];
      const field = shuffledNonZeroFields[i];

      // 填充字段名称到第5行（带自动换行和边框）
      const nameCell = this.targetSheet.getCell(5, col);
      nameCell.value = field.name;
      nameCell.style = {
        numFmt: 'General',
        alignment: { wrapText: true },
        border: {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        }
      };

      // 填充数值到第6行（带边框）
      const valueCell = this.targetSheet.getCell(6, col);
      valueCell.value = field.value;
      valueCell.style = {
        numFmt: 'General',
        border: {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        }
      };

      console.log(`  ${String.fromCharCode(64 + col)}5: ${field.name}, ${String.fromCharCode(64 + col)}6: ${field.value}`);
    }

    // 4. 填充剩余的空单元格（从所有字段名称中随机选择，排除已选择的）
    const remainingSlots = shuffledColumns.slice(shuffledNonZeroFields.length);
    const usedFieldNames = new Set(shuffledNonZeroFields.map(f => f.name));
    const availableFieldNames = fieldNames.filter(name => !usedFieldNames.has(name));

    // 随机打乱可用字段名称
    const shuffledAvailableFields = availableFieldNames.sort(() => Math.random() - 0.5);

    for (let i = 0; i < remainingSlots.length; i++) {
      const col = remainingSlots[i];
      const fieldName = shuffledAvailableFields[i] || `字段${col}`;

      // 只填充字段名称到第5行（带自动换行和边框），第6行保持空（带边框）
      const nameCell = this.targetSheet.getCell(5, col);
      nameCell.value = fieldName;
      nameCell.style = {
        numFmt: 'General',
        alignment: { wrapText: true },
        border: {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        }
      };

      // 第6行为空，但也要添加边框
      const emptyCell = this.targetSheet.getCell(6, col);
      emptyCell.style = {
        numFmt: 'General',
        border: {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        }
      };

      console.log(`  ${String.fromCharCode(64 + col)}5: ${fieldName} (无数值)`);
    }

    // 5. 复制E5-M6到O5-W6（偏移10列）
    console.log(`\\n复制E5-M6到O5-W6...`);
    for (let col = 5; col <= 13; col++) { // E-M列
      const targetCol = col + 10; // O-W列（E+10=O=15, M+10=W=23）
      
      // 复制第5行（字段名称，带自动换行和边框）
      const nameCell = this.targetSheet.getCell(5, col);
      const copyNameCell = this.targetSheet.getCell(5, targetCol);
      copyNameCell.value = nameCell.value;
      copyNameCell.style = {
        numFmt: 'General',
        alignment: { wrapText: true },
        border: {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        }
      };
      
      // 复制第6行（数值，带边框）
      const valueCell = this.targetSheet.getCell(6, col);
      const copyValueCell = this.targetSheet.getCell(6, targetCol);
      copyValueCell.value = valueCell.value;
      copyValueCell.style = {
        numFmt: 'General',
        border: {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        }
      };
      
      console.log(`  ${String.fromCharCode(64 + targetCol)}5: ${nameCell.value}, ${String.fromCharCode(64 + targetCol)}6: ${valueCell.value}`);
    }
  }
}
