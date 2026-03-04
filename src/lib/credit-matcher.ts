import * as XLSX from 'xlsx';

/**
 * 机构映射信息
 */
export interface OrganizationMapping {
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
 *
 * 特点：直接修改B文件的workbook，保留原始格式（样式、合并单元格、公式等）
 */
export class CreditMatcher {
  private sourceWorkbook: XLSX.WorkBook;
  private targetWorkbook: XLSX.WorkBook;
  private targetSheetName: string;
  private mappings: OrganizationMapping[] = [];

  constructor(
    sourceWorkbook: XLSX.WorkBook,
    targetWorkbook: XLSX.WorkBook,
    targetSheetName?: string
  ) {
    this.sourceWorkbook = sourceWorkbook;
    this.targetWorkbook = targetWorkbook;
    // 默认使用第一个工作表，如果指定了工作表名则使用指定的工作表
    this.targetSheetName = targetSheetName || targetWorkbook.SheetNames[0];

    console.log('=== CreditMatcher 初始化 ===');
    console.log('源文件工作表:', sourceWorkbook.SheetNames);
    console.log('目标文件工作表:', targetWorkbook.SheetNames);
    console.log('目标工作表:', this.targetSheetName);
  }

  /**
   * 执行匹配并直接修改目标workbook
   */
  public matchAndFill(): {
    workbook: XLSX.WorkBook;
    mappings: OrganizationMapping[];
    statistics: {
      totalOrganizations: number;
      matchedCount: number;
      unmatchedCount: number;
    };
  } {
    console.log('=== 开始授信模式匹配（直接修改B文件）===');

    // 获取源文件"单体"工作表的数据
    const sourceSheetName = '单体';
    if (!this.sourceWorkbook.SheetNames.includes(sourceSheetName)) {
      throw new Error(`源文件中未找到"${sourceSheetName}"工作表`);
    }

    const sourceSheet = this.sourceWorkbook.Sheets[sourceSheetName];
    const sourceData = XLSX.utils.sheet_to_json<TableCell[]>(sourceSheet, {
      header: 1,
      defval: null,
      raw: true,
    });

    // 获取目标工作表的数据
    const targetSheet = this.targetWorkbook.Sheets[this.targetSheetName];
    const targetData = XLSX.utils.sheet_to_json<TableCell[]>(targetSheet, {
      header: 1,
      defval: null,
      raw: true,
    });

    console.log('源表格（单体）数据行数:', sourceData.length);
    console.log('目标表格数据行数:', targetData.length);

    // 构建源文件的机构名称索引（B列，从第4行开始，索引3）
    const sourceOrgMap = new Map<string, number>();
    sourceData.slice(3).forEach((row, idx) => {
      const orgName = String(row[1] || '').trim(); // B列
      if (orgName && orgName !== '机构名称') {
        const actualRowIndex = idx + 3; // 实际行索引（从0开始）
        sourceOrgMap.set(orgName, actualRowIndex);
      }
    });

    console.log('源文件机构索引大小:', sourceOrgMap.size);

    // 遍历目标文件的机构（B列，从第6行开始，索引5）
    this.targetSheetName = this.targetWorkbook.SheetNames[0];
    targetData.slice(5).forEach((row, idx) => {
      const orgName = String(row[1] || '').trim(); // B列

      // 只处理有机构名称的行（从第5行开始，即idx >= 5）
      if (orgName && orgName !== '机构名称' && idx >= 5) {
        const targetRowIndex = idx + 5; // 实际行索引（从0开始）
        const sourceRowIndex = sourceOrgMap.get(orgName);

        const mapping: OrganizationMapping = {
          targetRowIndex,
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
          const sourceRow = sourceData[sourceRowIndex];
          mapping.valueC = this.parseNumber(sourceRow[2]); // C列
          mapping.valueD = this.parseNumber(sourceRow[3]); // D列
          // N列的值是D列的值（根据规则）
          mapping.valueN = mapping.valueD;

          console.log(`匹配成功: ${orgName}`);
          console.log(`  目标行: ${targetRowIndex + 1} (B${targetRowIndex + 1})`);
          console.log(`  源行: ${sourceRowIndex + 1} (B${sourceRowIndex + 4})`);
          console.log(`  C列值: ${mapping.valueC}`);
          console.log(`  D列值: ${mapping.valueD}`);
          console.log(`  N列值: ${mapping.valueN}`);

          // 直接修改目标workbook的单元格
          this.updateTargetCell(targetSheet, targetRowIndex, 2, mapping.valueC); // C列
          this.updateTargetCell(targetSheet, targetRowIndex, 3, mapping.valueD); // D列
          this.updateTargetCell(targetSheet, targetRowIndex, 13, mapping.valueN); // N列

          console.log(`已填充: C${targetRowIndex + 1}=${mapping.valueC}, D${targetRowIndex + 1}=${mapping.valueD}, N${targetRowIndex + 1}=${mapping.valueN}`);
        } else {
          console.log(`匹配失败: ${orgName}`);
        }

        this.mappings.push(mapping);
      }
    });

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
   * 更新目标workbook的单元格
   */
  private updateTargetCell(
    sheet: XLSX.WorkSheet,
    rowIndex: number,
    colIndex: number,
    value: number | null
  ): void {
    if (value === null) return;

    // 将行列索引转换为单元格地址（例如：C6, D6, N6）
    const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });

    // 如果单元格存在，只修改值，保留格式
    if (sheet[cellAddress]) {
      sheet[cellAddress].v = value;
      // 确保数据类型是数字
      sheet[cellAddress].t = 'n';
    } else {
      // 如果单元格不存在，创建新单元格
      sheet[cellAddress] = {
        v: value,
        t: 'n', // 数字类型
      };
    }
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

// 类型定义
type TableCell = string | number | null | undefined;
