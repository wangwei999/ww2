import { TableData, TableCell } from './types';
import { normalizeDate } from './data-utils';
import * as XLSX from 'xlsx';

/**
 * 机构映射信息
 */
interface OrganizationMapping {
  targetRowIndex: number;  // 目标文件中的行索引（从0开始）
  sourceRowIndex: number;  // 源文件中的行索引（从0开始）
  orgName: string;         // 机构名称
  matched: boolean;        // 是否匹配成功
  valueC: number | null;   // C列的值（原授信时间）
  valueD: number | null;   // D列的值（原授信-授信总额）
  valueN: number | null;   // N列的值（拟调整授信-授信总额）
}

/**
 * 授信模式匹配器
 *
 * 规则：
 * 1. B文件（目标）：第5行是字段行，B6开始是机构名称
 * 2. A文件（源）："单体"工作表，B4开始是机构名称
 * 3. 匹配成功后：
 *    - A文件C列X行 → B文件C列对应行
 *    - A文件D列X行 → B文件D列对应行
 *    - A文件D列X行 → B文件N列对应行
 */
export class CreditMatcher {
  private sourceWorkbook: XLSX.WorkBook;
  private targetWorkbook: XLSX.WorkBook;
  private sourceTable: TableData;
  private targetTable: TableData;
  private mappings: OrganizationMapping[] = [];

  constructor(
    sourceWorkbook: XLSX.WorkBook,
    targetWorkbook: XLSX.WorkBook
  ) {
    this.sourceWorkbook = sourceWorkbook;
    this.targetWorkbook = targetWorkbook;

    // 提取源文件（"单体"工作表）
    this.sourceTable = this.extractSheetTable(sourceWorkbook, '单体');

    // 提取目标文件（第一个工作表）
    this.targetTable = this.extractSheetTable(targetWorkbook, targetWorkbook.SheetNames[0]);

    console.log('=== CreditMatcher 初始化 ===');
    console.log('源文件工作表:', sourceWorkbook.SheetNames);
    console.log('目标文件工作表:', targetWorkbook.SheetNames);
    console.log('源表格（单体）行数:', this.sourceTable.rows.length);
    console.log('目标表格行数:', this.targetTable.rows.length);
  }

  /**
   * 从工作簿中提取表格数据
   */
  private extractSheetTable(workbook: XLSX.WorkBook, sheetName: string): TableData {
    const sheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json<TableCell[]>(sheet, {
      header: 1,
      defval: null,
      raw: true,
    });

    // 使用第一行作为表头（如果存在）
    const headers: string[] = [];
    if (rawData.length > 0) {
      rawData[0].forEach((cell, idx) => {
        headers.push(String(cell || `列${idx + 1}`));
      });
    }

    // 数据行（跳过表头行）
    const rows = rawData.slice(1);

    return {
      headers,
      rows,
    };
  }

  /**
   * 执行匹配
   */
  public matchAndFill(): {
    filledTable: TableData;
    mappings: OrganizationMapping[];
    statistics: {
      totalOrganizations: number;
      matchedCount: number;
      unmatchedCount: number;
    };
  } {
    console.log('=== 开始授信模式匹配 ===');

    // 构建源文件的机构名称索引（B列，从第3行开始）
    const sourceOrgMap = new Map<string, number>();
    this.sourceTable.rows.forEach((row, idx) => {
      const orgName = String(row[1] || '').trim(); // B列
      if (orgName && orgName !== '机构名称') {
        sourceOrgMap.set(orgName, idx); // 存储行索引（从0开始）
      }
    });

    console.log('源文件机构索引大小:', sourceOrgMap.size);

    // 遍历目标文件的机构（B列，从第5行开始）
    this.targetTable.rows.forEach((row, idx) => {
      const orgName = String(row[1] || '').trim(); // B列

      // 只处理有机构名称的行（从第5行开始，即idx >= 5）
      if (orgName && orgName !== '机构名称' && idx >= 5) {
        const sourceRowIndex = sourceOrgMap.get(orgName);

        const mapping: OrganizationMapping = {
          targetRowIndex: idx,
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

          // 从源文件获取数值
          const sourceRow = this.sourceTable.rows[sourceRowIndex];
          mapping.valueC = this.parseNumber(sourceRow[2]); // C列
          mapping.valueD = this.parseNumber(sourceRow[3]); // D列
          // N列的值是D列的值（根据规则）
          mapping.valueN = mapping.valueD;

          console.log(`匹配成功: ${orgName}`);
          console.log(`  目标行: ${idx + 1} (B${idx + 1})`);
          console.log(`  源行: ${sourceRowIndex + 2} (B${sourceRowIndex + 4})`);
          console.log(`  C列值: ${mapping.valueC}`);
          console.log(`  D列值: ${mapping.valueD}`);
          console.log(`  N列值: ${mapping.valueN}`);
        } else {
          console.log(`匹配失败: ${orgName}`);
        }

        this.mappings.push(mapping);
      }
    });

    // 填充目标表格
    this.fillTargetTable();

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
      filledTable: this.targetTable,
      mappings: this.mappings,
      statistics,
    };
  }

  /**
   * 填充目标表格
   */
  private fillTargetTable(): void {
    console.log('=== 开始填充目标表格 ===');

    this.mappings.forEach(mapping => {
      if (mapping.matched) {
        const row = this.targetTable.rows[mapping.targetRowIndex];

        // 填充C列（原授信时间）
        if (mapping.valueC !== null) {
          row[2] = mapping.valueC;
        }

        // 填充D列（原授信-授信总额）
        if (mapping.valueD !== null) {
          row[3] = mapping.valueD;
        }

        // 填充N列（拟调整授信-授信总额）
        if (mapping.valueN !== null) {
          row[13] = mapping.valueN; // N列索引为13
        }

        console.log(`已填充行 ${mapping.targetRowIndex + 1}: C=${mapping.valueC}, D=${mapping.valueD}, N=${mapping.valueN}`);
      }
    });
  }

  /**
   * 解析数值
   */
  private parseNumber(cell: TableCell): number | null {
    if (cell === null || cell === undefined || cell === '') {
      return null;
    }

    if (typeof cell === 'number') {
      return cell;
    }

    const num = parseFloat(String(cell).trim());
    return isNaN(num) ? null : num;
  }
}
