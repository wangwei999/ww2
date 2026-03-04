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

    // 2. 查找所有匹配成功的机构及其数值（按字段名称统计）
    const matchedData: Array<{
      targetRow: number;
      orgName: string;
      sourceRow: number;
      fieldValueMap: Map<string, number>; // 字段名称 -> 数值
    }> = [];

    // 统计每个字段是否至少有一个机构有数值
    const fieldHasValueSet = new Set<string>();
    // 记录每个字段的值（按机构）
    const fieldValuesByOrg = new Map<string, Map<number, number>>(); // 字段名称 -> Map<源行, 数值>

    for (const mapping of this.mappings) {
      if (mapping.matched && mapping.sourceRowIndex > 0) {
        const sourceRow = this.sourceSheet.getRow(mapping.sourceRowIndex);
        const fieldValueMap = new Map<string, number>();

        // 检查E列到AB列的数值
        for (let col = 5; col <= 28; col++) { // E列(5)到AB列(28)
          const cell = sourceRow.getCell(col);
          const value = this.parseCellValue(cell.value);
          
          const fieldName = fieldNames[col - 5] || `字段${col}`;
          
          if (value !== null && value !== 0) {
            // 记录该机构在此字段的值
            fieldValueMap.set(fieldName, value);
            
            // 标记该字段有数值
            fieldHasValueSet.add(fieldName);
            
            // 记录该字段的值（按机构）
            if (!fieldValuesByOrg.has(fieldName)) {
              fieldValuesByOrg.set(fieldName, new Map());
            }
            fieldValuesByOrg.get(fieldName)!.set(mapping.sourceRowIndex, value);
          }
        }

        matchedData.push({
          targetRow: mapping.targetRowIndex,
          orgName: mapping.orgName,
          sourceRow: mapping.sourceRowIndex,
          fieldValueMap
        });
        console.log(`机构: ${mapping.orgName} (目标行${mapping.targetRowIndex}), 有数值字段数: ${fieldValueMap.size}`);
      }
    }

    console.log(`共找到 ${matchedData.length} 个有数值的机构`);
    console.log(`共有 ${fieldHasValueSet.size} 个字段至少有一个机构有数值`);

    // 3. 确定字段列表
    // 固定字段：至少有一个机构有数值的字段
    const fixedFields = Array.from(fieldHasValueSet);
    console.log(`固定字段数: ${fixedFields.length}`, fixedFields);
    
    // 随机字段：所有机构都没有数值的字段
    const allFieldNamesSet = new Set(fieldNames);
    const randomFieldsCandidates = fieldNames.filter(name => !fieldHasValueSet.has(name));
    console.log(`可用随机字段数: ${randomFieldsCandidates.length}`);
    
    // 计算需要多少随机字段（最多9个）
    const totalSlots = 9; // E5-M5共9个单元格
    const randomFieldsNeeded = Math.max(0, totalSlots - fixedFields.length);
    const selectedRandomFields = randomFieldsCandidates
      .sort(() => Math.random() - 0.5)
      .slice(0, randomFieldsNeeded);
    
    // 合并固定字段和随机字段
    const allFieldsToFill = [...fixedFields, ...selectedRandomFields];
    console.log(`最终填充字段数: ${allFieldsToFill.length}`, allFieldsToFill);

    // 4. 填充第5行（字段名称行，所有机构共享）和每个机构的数值行
    this.fillSharedFieldsAndValues(matchedData, allFieldsToFill, fieldValuesByOrg);

    console.log('=== 随机字段填充完成 ===');
  }

  /**
   * 填充共享字段名称和各机构的数值
   * 第5行：共享的字段名称
   * 每个机构的行：对应的数值
   */
  private fillSharedFieldsAndValues(
    matchedData: Array<{
      targetRow: number;
      orgName: string;
      sourceRow: number;
      fieldValueMap: Map<string, number>;
    }>,
    fieldNames: string[],
    fieldValuesByOrg: Map<string, Map<number, number>>
  ): void {
    console.log('\\n=== 开始填充共享字段和各机构数值 ===');

    // B文件E-M列（5-13列）
    const targetColumns = [5, 6, 7, 8, 9, 10, 11, 12, 13]; // E-M列
    
    // 随机打乱列顺序
    const shuffledColumns = [...targetColumns].sort(() => Math.random() - 0.5);

    // 1. 填充第5行（字段名称行，所有机构共享）
    console.log('\\n填充第5行（字段名称）:');
    for (let i = 0; i < Math.min(fieldNames.length, targetColumns.length); i++) {
      const col = shuffledColumns[i];
      const fieldName = fieldNames[i];
      
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
      
      console.log(`  ${String.fromCharCode(64 + col)}5: ${fieldName}`);
    }

    // 2. 为每个机构填充对应的数值行
    for (const orgData of matchedData) {
      const { targetRow, orgName, sourceRow } = orgData;
      console.log(`\\n填充机构 ${orgName} (目标行${targetRow}) 的数值:`);

      for (let i = 0; i < Math.min(fieldNames.length, targetColumns.length); i++) {
        const col = shuffledColumns[i];
        const fieldName = fieldNames[i];
        
        // 查找该机构在此字段的值
        const valueCell = this.targetSheet.getCell(targetRow, col);
        const fieldValues = fieldValuesByOrg.get(fieldName);
        const value = fieldValues?.get(sourceRow);

        if (value !== undefined) {
          valueCell.value = value;
          console.log(`  ${String.fromCharCode(64 + col)}${targetRow}: ${value}`);
        } else {
          valueCell.value = null; // 该机构在此字段无数值
          console.log(`  ${String.fromCharCode(64 + col)}${targetRow}: (空)`);
        }
        
        valueCell.style = {
          numFmt: 'General',
          border: {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
          }
        };
      }
    }

    // 3. 为所有列添加边框（包括未填充的列）
    console.log('\\n为所有E-M列和O-W列添加边框:');
    const rowsWithBorders = [5, ...matchedData.map(m => m.targetRow)];
    for (const row of rowsWithBorders) {
      for (const col of [...targetColumns, ...targetColumns.map(c => c + 10)]) {
        const cell = this.targetSheet.getCell(row, col);
        // 如果没有设置过样式，则添加边框
        if (!cell.style || !cell.style.border) {
          cell.style = {
            numFmt: 'General',
            alignment: row === 5 ? { wrapText: true } : undefined,
            border: {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            }
          };
        }
      }
    }

    // 4. 复制第5行（字段名称）到O5-W5
    console.log('\\n复制第5行到O5-W5:');
    for (let col = 5; col <= 13; col++) {
      const targetCol = col + 10;
      const nameCell = this.targetSheet.getCell(5, col);
      const copyNameCell = this.targetSheet.getCell(5, targetCol);
      copyNameCell.value = nameCell.value;
      copyNameCell.style = nameCell.style;
      console.log(`  ${String.fromCharCode(64 + targetCol)}5: ${nameCell.value}`);
    }

    // 5. 复制各机构的数值行到O-W区域
    for (const orgData of matchedData) {
      const { targetRow, orgName } = orgData;
      console.log(`\\n复制机构 ${orgName} (目标行${targetRow}) 的数值到O-W区域:`);

      for (let col = 5; col <= 13; col++) {
        const targetCol = col + 10;
        const valueCell = this.targetSheet.getCell(targetRow, col);
        const copyValueCell = this.targetSheet.getCell(targetRow, targetCol);
        copyValueCell.value = valueCell.value;
        copyValueCell.style = valueCell.style;
        console.log(`  ${String.fromCharCode(64 + targetCol)}${targetRow}: ${valueCell.value}`);
      }
    }
  }
}
