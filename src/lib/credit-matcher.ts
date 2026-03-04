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
  sourceSheet?: string; // 标记从哪个表匹配的（'单体' 或 '集团'）
  colC?: number; // C列的列号（单体表3，集团表30）
  colD?: number; // D列的列号（单体表4，集团表5）
}

/**
 * 授信模式匹配器（使用exceljs保留原始格式）
 */
export class CreditMatcher {
  private sourceWorkbook: ExcelJS.Workbook;
  private targetWorkbook: ExcelJS.Workbook;
  private targetSheet: ExcelJS.Worksheet;
  private sourceSheet单体: ExcelJS.Worksheet;
  private sourceSheet集团: ExcelJS.Worksheet | null;
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

    // 尝试获取单体表
    const sourceSheetName = '单体';
    const sourceSheet = sourceWorkbook.getWorksheet(sourceSheetName);
    if (!sourceSheet) {
      throw new Error(`源文件中未找到"${sourceSheetName}"工作表`);
    }
    this.sourceSheet单体 = sourceSheet as ExcelJS.Worksheet;

    // 尝试获取集团表（可选，尝试多种可能的名称）
    const sourceSheet集团 = sourceWorkbook.getWorksheet('集团') || 
                             sourceWorkbook.getWorksheet('集团 ') ||
                             sourceWorkbook.getWorksheet('集团表');
    this.sourceSheet集团 = sourceSheet集团 || null;

    console.log('=== CreditMatcher 初始化（使用exceljs）===');
    console.log('源文件工作表:', sourceWorkbook.worksheets.map(ws => ws.name));
    console.log('目标文件工作表:', targetWorkbook.worksheets.map(ws => ws.name));
    console.log('目标工作表:', this.targetSheet.name);
    console.log('源工作表-单体:', this.sourceSheet单体.name);
    console.log('源工作表-集团:', this.sourceSheet集团?.name || '无');
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

    // 构建单体表的机构名称索引（B列，从第4行开始）
    const sourceOrgMap单体 = new Map<string, number>();
    const sourceRowCount单体 = this.sourceSheet单体.rowCount;
    
    console.log('单体表总行数:', sourceRowCount单体);
    
    for (let row = 4; row <= sourceRowCount单体; row++) {
      const cell = this.sourceSheet单体.getCell(row, 2);
      const orgName = String(cell.value || '').trim();
      
      if (orgName && orgName !== '机构名称') {
        const normalizedName = normalizeOrganizationName(orgName);
        sourceOrgMap单体.set(normalizedName, row);
        
        if (normalizedName !== orgName) {
          console.log(`单体表规范化: ${orgName} → ${normalizedName} (行${row})`);
        }
      }
    }

    // 构建集团表的机构名称索引（D列，从第4行开始）
    const sourceOrgMap集团 = new Map<string, number>();
    if (this.sourceSheet集团) {
      const sourceRowCount集团 = this.sourceSheet集团.rowCount;
      console.log('集团表总行数:', sourceRowCount集团);
      
      for (let row = 4; row <= sourceRowCount集团; row++) {
        const cell = this.sourceSheet集团.getCell(row, 4); // D列
        const orgName = String(cell.value || '').trim();
        
        if (orgName && orgName !== '机构名称') {
          const normalizedName = normalizeOrganizationName(orgName);
          sourceOrgMap集团.set(normalizedName, row);
          
          if (normalizedName !== orgName) {
            console.log(`集团表规范化: ${orgName} → ${normalizedName} (行${row})`);
          }
        }
      }
    }

        console.log('单体表机构索引大小:', sourceOrgMap单体.size);
    console.log('集团表机构索引大小:', sourceOrgMap集团.size);
    console.log('集团表是否存在:', !!this.sourceSheet集团);
    
    // 输出集团表的前5个机构名称用于调试
    if (this.sourceSheet集团) {
      console.log('集团表前5个机构名称:');
      let count = 0;
      for (let row = 4; row <= this.sourceSheet集团.rowCount && count < 5; row++) {
        const cell = this.sourceSheet集团.getCell(row, 4);
        const orgName = String(cell.value || '').trim();
        if (orgName && orgName !== '机构名称') {
          console.log(`  行${row}: ${orgName}`);
          count++;
        }
      }
    }

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
        
        // 先从单体表查找
        let sourceRowIndex = sourceOrgMap单体.get(normalizedName);
        let sourceSheet = '单体';
        let colC = 3; // 单体表C列
        let colD = 4; // 单体表D列
        
        // 如果单体表找不到，再从集团表查找
        if (sourceRowIndex === undefined && this.sourceSheet集团) {
          sourceRowIndex = sourceOrgMap集团.get(normalizedName);
          if (sourceRowIndex !== undefined) {
            sourceSheet = '集团';
            colC = 30; // 集团表AD列（AD=30）
            colD = 5;  // 集团表E列
          }
        }
        
        console.log(`  源文件行号: ${sourceRowIndex !== undefined ? `${sourceSheet}行${sourceRowIndex}` : '未找到'}`);

        const mapping: OrganizationMapping = {
          targetRowIndex: row,
          sourceRowIndex: -1,
          orgName,
          matched: false,
          valueC: null,
          valueD: null,
          valueN: null,
          sourceSheet: undefined,
          colC: undefined,
          colD: undefined,
        };

        if (sourceRowIndex !== undefined) {
          const actualSourceSheet = sourceSheet === '单体' ? this.sourceSheet单体 : this.sourceSheet集团;
          if (!actualSourceSheet) continue;
          
          mapping.matched = true;
          mapping.sourceRowIndex = sourceRowIndex;
          mapping.sourceSheet = sourceSheet;
          mapping.colC = colC;
          mapping.colD = colD;

          const cellC = actualSourceSheet.getCell(sourceRowIndex, colC);
          const cellD = actualSourceSheet.getCell(sourceRowIndex, colD);

          const matchInfo = normalizedName !== orgName 
            ? `(${orgName} → ${normalizedName})` 
            : orgName;
          
          console.log(`匹配成功 (${sourceSheet}): ${matchInfo}`);
          console.log(`  源列${String.fromCharCode(64 + colC)}type: ${cellC.type}, 源列${String.fromCharCode(64 + colD)}type: ${cellD.type}`);
          console.log(`  源列${String.fromCharCode(64 + colC)}value: ${JSON.stringify(cellC.value)}, 源列${String.fromCharCode(64 + colD)}value: ${JSON.stringify(cellD.value)}`);
          console.log(`  源列${String.fromCharCode(64 + colC)}完整对象: ${JSON.stringify({type: cellC.type, value: cellC.value, result: (cellC as any).result, formula: (cellC as any).formula, text: (cellC as any).text, numFmt: cellC.numFmt})}`);
          console.log(`  源列${String.fromCharCode(64 + colD)}完整对象: ${JSON.stringify({type: cellD.type, value: cellD.value, result: (cellD as any).result, formula: (cellD as any).formula, text: (cellD as any).text, numFmt: cellD.numFmt})}`);
          
          // 对于第一行匹配，额外输出源文件该行的所有信息
          if (this.mappings.length === 0) {
            console.log(`  === 第一行匹配，输出源文件行${sourceRowIndex}的所有信息 ===`);
            const sourceRow = actualSourceSheet.getRow(sourceRowIndex);
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

    // 收集所有机构的字段名称和数值
    const allFieldNames = new Set<string>();
    const fieldValuesByOrg = new Map<string, Map<number, { value: number; sourceSheet: string }>>();

    for (const mapping of this.mappings) {
      if (mapping.matched && mapping.sourceRowIndex > 0) {
        // 确定机构匹配的源表
        const sourceSheet = mapping.sourceSheet === '集团' ? this.sourceSheet集团 : this.sourceSheet单体;
        if (!sourceSheet) continue;
        
        const headerRow = sourceSheet.getRow(3);
        const sourceRow = sourceSheet.getRow(mapping.sourceRowIndex);
        
        // 集团表从F列开始（6），单体表从E列开始（5）
        const startCol = mapping.sourceSheet === '集团' ? 6 : 5;
        const endCol = mapping.sourceSheet === '集团' ? 29 : 28; // 集团表多一列

        // 收集字段名称和数值
        for (let col = startCol; col <= endCol; col++) {
          const fieldNameCell = headerRow.getCell(col);
          const fieldName = String(fieldNameCell.value || '').trim();
          
          if (!fieldName) continue;
          
          allFieldNames.add(fieldName);
          
          const valueCell = sourceRow.getCell(col);
          const value = this.parseCellValue(valueCell.value);
          
          if (value !== null && value !== 0) {
            if (!fieldValuesByOrg.has(fieldName)) {
              fieldValuesByOrg.set(fieldName, new Map());
            }
            fieldValuesByOrg.get(fieldName)!.set(mapping.sourceRowIndex, {
              value: value,
              sourceSheet: mapping.sourceSheet || '单体'
            });
          }
        }
      }
    }

    console.log(`共找到 ${allFieldNames.size} 个唯一字段名称`);
    console.log(`字段列表:`, Array.from(allFieldNames));

    // 确定字段列表：固定字段（有数值的）和随机字段
    const fixedFields = Array.from(fieldValuesByOrg.keys());
    const randomFieldsCandidates = Array.from(allFieldNames).filter(name => !fieldValuesByOrg.has(name));
    const randomFieldsNeeded = Math.max(0, 9 - fixedFields.length);
    const selectedRandomFields = randomFieldsCandidates.sort(() => Math.random() - 0.5).slice(0, randomFieldsNeeded);
    
    const allFieldsToFill = [...fixedFields, ...selectedRandomFields];
    console.log(`最终填充字段数: ${allFieldsToFill.length}`);

    // 填充第5行和各机构的数值行
    const matchedData = this.mappings
      .filter(m => m.matched && m.sourceRowIndex > 0)
      .map(m => ({
        targetRow: m.targetRowIndex,
        orgName: m.orgName,
        sourceRow: m.sourceRowIndex,
        sourceSheet: m.sourceSheet || '单体'
      }));

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
      sourceSheet: string;
    }>,
    fieldNames: string[],
    fieldValuesByOrg: Map<string, Map<number, { value: number; sourceSheet: string }>>
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
      const { targetRow, orgName, sourceRow, sourceSheet } = orgData;
      console.log(`\\n填充机构 ${orgName} (目标行${targetRow}, 源表:${sourceSheet}) 的数值:`);

      // 确定源表
      const sourceSheetObj = sourceSheet === '集团' ? this.sourceSheet集团 : this.sourceSheet单体;
      if (!sourceSheetObj) {
        console.log(`  警告: 找不到源表 ${sourceSheet}`);
        continue;
      }

      // 确定源表的列起始位置
      const startCol = sourceSheet === '集团' ? 6 : 5;

      for (let i = 0; i < Math.min(fieldNames.length, targetColumns.length); i++) {
        const col = shuffledColumns[i];
        const fieldName = fieldNames[i];
        
        const valueCell = this.targetSheet.getCell(targetRow, col);
        const fieldValues = fieldValuesByOrg.get(fieldName);
        const valueInfo = fieldValues?.get(sourceRow);

        if (valueInfo && valueInfo.value !== undefined) {
          valueCell.value = valueInfo.value;
          console.log(`  ${String.fromCharCode(64 + col)}${targetRow}: ${valueInfo.value} (从${valueInfo.sourceSheet})`);
        } else {
          valueCell.value = null;
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
