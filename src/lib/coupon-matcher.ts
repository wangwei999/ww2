import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { containsGeographyKeyword } from './geo-keywords';

/**
 * 债券数据结构
 */
interface BondData {
  rowNumber: number;       // 原始行号
  bondCode: string;        // 债券代码（B列）
  bondName: string;        // 债券名称（C列）
  availableAmount: number; // 可用金额（E列，万元）
  rowData: any[];          // 整行数据
  bondActualType: 'treasury' | 'local'; // 债券实际类型
}

/**
 * 单个债券的分配结果
 */
interface BondAllocation {
  bond: BondData;           // 债券信息
  allocatedAmount: number;  // 本次分配的金额
}

/**
 * 单个金额的处理结果（一个债券集合）
 */
interface AmountGroup {
  targetAmount: number;           // 目标挑券金额
  allocations: BondAllocation[];  // 分配的债券列表
  actualAmount: number;           // 实际挑券金额
}

/**
 * 禁挑券规则
 */
interface ExclusionRule {
  original: string;         // 原始输入
  type: 'code' | 'name';    // 类型：代码精确匹配 / 名称模糊匹配
  groupIndex?: number;      // 针对第几笔金额（从1开始），不设置表示全局禁挑
  keyword: string;          // 匹配关键词
}

/**
 * 挑券模式处理器
 * 完全独立于其他功能模块
 * 
 * 支持单金额和多金额模式：
 * - 单金额：输出第7行F列显示总金额
 * - 多金额：每个集合间空一行，空行F列显示下一个金额
 * 
 * 支持禁挑券功能：
 * - 数字：精确匹配债券代码(B列)
 * - 文字：模糊匹配债券简称(C列)
 * 
 * 支持文件格式：.xls 和 .xlsx
 * 
 * Excel结构说明：
 * - 第1-7行：其他内容（标题、说明等）
 * - 第8行：字段列名（表头）
 * - 第9行开始：债券数据
 * - B列：债券代码
 * - C列：债券名称
 * - E列：可用金额
 */
export class CouponMatcher {
  private file: File | Buffer;
  private bondType: 'treasury' | 'local'; // 用户选择的债券类型
  private amounts: number[]; // 挑券金额数组（万元）
  private excludedBonds: string[]; // 禁挑券列表
  private workbook: ExcelJS.Workbook;
  
  // 原始工作表数据（用于复制格式）
  private sourceSheetData: any[][] = [];
  private isXlsFormat: boolean = false;

  // Excel结构常量
  private readonly HEADER_ROW = 8;      // 表头行号
  private readonly DATA_START_ROW = 9;  // 数据起始行号
  private readonly BOND_CODE_COL = 2;   // B列 - 债券代码
  private readonly BOND_NAME_COL = 3;   // C列 - 债券名称
  private readonly AVAILABLE_COL = 5;   // E列 - 可用金额

  // 债券数据列表
  private bonds: BondData[] = [];

  // 禁挑券规则
  private exclusionRules: ExclusionRule[] = [];

  // 处理结果
  private result: {
    bondType: 'treasury' | 'local';
    totalRows: number;
    totalAvailable: number;
    excludedCount: number;
    typeFilteredCount: number; // 因类型不符被过滤的数量
    groups: AmountGroup[];
  } = {
    bondType: 'treasury',
    totalRows: 0,
    totalAvailable: 0,
    excludedCount: 0,
    typeFilteredCount: 0,
    groups: []
  };

  // 最小占用金额（万元）
  private readonly MIN_OCCUPY_AMOUNT = 100;

  constructor(
    file: File | Buffer,
    bondType: 'treasury' | 'local',
    amounts: number[],
    excludedBonds: string[] = []
  ) {
    this.file = file;
    this.bondType = bondType;
    this.amounts = amounts;
    this.excludedBonds = excludedBonds;
    this.workbook = new ExcelJS.Workbook();
    
    // 解析禁挑券规则
    this.parseExclusionRules();
  }

  /**
   * 解析禁挑券规则
   * 格式：
   * - 纯数字：代码精确匹配（全局）
   * - 文字：名称模糊匹配（全局）
   * - /1国开：第1笔金额禁挑名称包含"国开"的债券
   * - /2250206：第2笔金额禁挑代码为"250206"的债券
   */
  private parseExclusionRules(): void {
    this.exclusionRules = [];
    
    for (const item of this.excludedBonds) {
      const trimmed = item.trim();
      if (!trimmed) continue;
      
      // 检查是否有 /数字 前缀（序号只取1位数字，支持1-9）
      const groupMatch = trimmed.match(/^\/([1-9])(.+)$/);
      
      if (groupMatch) {
        // 带金额序号的禁挑券
        const groupIndex = parseInt(groupMatch[1], 10);
        const keyword = groupMatch[2].trim();
        const isNumeric = /^\d+$/.test(keyword);
        
        this.exclusionRules.push({
          original: trimmed,
          type: isNumeric ? 'code' : 'name',
          groupIndex,
          keyword
        });
      } else {
        // 全局禁挑券
        const isNumeric = /^\d+$/.test(trimmed);
        
        this.exclusionRules.push({
          original: trimmed,
          type: isNumeric ? 'code' : 'name',
          keyword: trimmed
        });
      }
    }
    
    if (this.exclusionRules.length > 0) {
      console.log('禁挑券规则:', this.exclusionRules.map(r => {
        const scope = r.groupIndex ? `第${r.groupIndex}笔` : '全局';
        const matchType = r.type === 'code' ? '代码精确' : '名称模糊';
        return `${r.keyword}(${scope},${matchType})`;
      }).join(', '));
    }
  }

  /**
   * 检查债券是否被禁挑
   * @param bond 债券信息
   * @param currentGroupIndex 当前处理的金额序号（从1开始）
   */
  private isBondExcluded(bond: BondData, currentGroupIndex: number): boolean {
    for (const rule of this.exclusionRules) {
      // 如果规则指定了金额序号，只对对应序号生效
      if (rule.groupIndex !== undefined && rule.groupIndex !== currentGroupIndex) {
        continue;
      }
      
      if (rule.type === 'code') {
        // 代码精确匹配
        if (bond.bondCode === rule.keyword) {
          const scope = rule.groupIndex ? `第${rule.groupIndex}笔` : '全局';
          console.log(`  禁挑: ${bond.bondName}(${bond.bondCode}) - ${scope}代码精确匹配"${rule.keyword}"`);
          return true;
        }
      } else {
        // 名称模糊匹配（包含）
        if (bond.bondName.includes(rule.keyword)) {
          const scope = rule.groupIndex ? `第${rule.groupIndex}笔` : '全局';
          console.log(`  禁挑: ${bond.bondName}(${bond.bondCode}) - ${scope}名称包含"${rule.keyword}"`);
          return true;
        }
      }
    }
    return false;
  }

  /**
   * 主处理方法
   */
  async process(): Promise<{ workbook: ExcelJS.Workbook; statistics: any }> {
    console.log('=== 开始挑券处理 ===');
    console.log('用户选择类型:', this.bondType === 'local' ? '地方债' : '国债');
    console.log('挑券金额:', this.amounts, '万元');
    console.log('模式:', this.amounts.length > 1 ? '多金额' : '单金额');
    console.log('禁挑券数量:', this.exclusionRules.length);

    // 1. 加载Excel文件（支持 .xls 和 .xlsx）
    await this.loadExcelFile();

    // 2. 读取债券数据（从第9行开始），并判断每只债券的类型
    this.readBondData();

    // 3. 根据金额匹配债券（根据用户选择的类型过滤，排除禁挑券）
    this.matchBondsByAmounts();

    // 4. 生成结果Excel
    await this.generateResultWorkbook();

    console.log('=== 挑券处理完成 ===');
    console.log('挑券统计:', {
      总债券数: this.result.totalRows,
      用户选择类型: this.bondType === 'local' ? '地方债' : '国债',
      类型不符过滤: this.result.typeFilteredCount,
      禁挑数量: this.result.excludedCount,
      总可用金额: this.result.totalAvailable,
      总挑券金额: this.amounts.reduce((a, b) => a + b, 0),
      集合数量: this.result.groups.length,
      各集合: this.result.groups.map((g, i) => ({
        序号: i + 1,
        目标金额: g.targetAmount,
        实际金额: g.actualAmount,
        债券数量: g.allocations.length
      }))
    });

    return {
      workbook: this.workbook,
      statistics: {
        bondType: this.bondType,
        totalRows: this.result.totalRows,
        totalAvailable: this.result.totalAvailable,
        excludedCount: this.result.excludedCount,
        typeFilteredCount: this.result.typeFilteredCount,
        groups: this.result.groups.map(g => ({
          targetAmount: g.targetAmount,
          actualAmount: g.actualAmount,
          bondCount: g.allocations.length
        })),
        totalSelected: this.result.groups.reduce((sum, g) => sum + g.actualAmount, 0),
        requestedAmounts: this.amounts
      }
    };
  }

  /**
   * 加载Excel文件
   * 使用 xlsx 库支持 .xls 和 .xlsx 格式
   */
  private async loadExcelFile(): Promise<void> {
    console.log('加载Excel文件...');

    let buffer: Buffer;
    let fileName = '';
    
    if (this.file instanceof File) {
      const arrayBuffer = await this.file.arrayBuffer();
      buffer = Buffer.from(arrayBuffer);
      fileName = this.file.name;
      console.log('文件名:', fileName, '大小:', buffer.length, 'bytes');
    } else {
      buffer = this.file;
      console.log('Buffer大小:', buffer.length, 'bytes');
    }

    // 检测文件格式
    this.isXlsFormat = fileName.toLowerCase().endsWith('.xls') && !fileName.toLowerCase().endsWith('.xlsx');
    console.log('文件格式:', this.isXlsFormat ? '.xls (旧格式)' : '.xlsx (新格式)');

    // 使用 xlsx 库读取文件（支持 .xls 和 .xlsx）
    const xlsxWorkbook = XLSX.read(buffer, { type: 'buffer' });
    
    const sheetNames = xlsxWorkbook.SheetNames;
    console.log('工作表列表:', sheetNames);
    
    if (sheetNames.length === 0) {
      throw new Error('Excel文件中没有工作表，请检查文件内容');
    }

    // 读取第一个工作表的数据
    const firstSheet = xlsxWorkbook.Sheets[sheetNames[0]];
    this.sourceSheetData = XLSX.utils.sheet_to_json(firstSheet, { 
      header: 1,
      defval: null,
      raw: false  // 保持原始格式
    }) as any[][];
    
    console.log('工作表名称:', sheetNames[0]);
    console.log('数据行数:', this.sourceSheetData.length);

    // 如果是 .xlsx 格式，用 exceljs 加载以保持格式
    if (!this.isXlsFormat) {
      try {
        await this.workbook.xlsx.load(buffer as any);
        console.log('ExcelJS 加载成功，保留原格式');
      } catch (e) {
        console.warn('ExcelJS 加载失败，将使用基础格式:', (e as Error).message);
        // 如果 exceljs 加载失败，创建空工作簿
        this.workbook = new ExcelJS.Workbook();
      }
    } else {
      // .xls 格式，创建新工作簿
      this.workbook = new ExcelJS.Workbook();
      console.log('创建新工作簿用于输出 .xlsx 格式');
    }
  }

  /**
   * 读取债券数据
   * 从第9行开始读取（第8行是表头）
   * 同时判断每只债券的类型（地方债/国债）
   */
  private readBondData(): void {
    console.log('读取债券数据（从第9行开始）...');

    if (this.sourceSheetData.length === 0) {
      console.warn('没有数据可读取');
      return;
    }

    // 从第9行开始读取数据（数组索引从0开始，所以是索引8）
    for (let rowIndex = this.DATA_START_ROW - 1; rowIndex < this.sourceSheetData.length; rowIndex++) {
      const row = this.sourceSheetData[rowIndex];
      if (!row) continue;
      
      // 读取B列（索引1）、C列（索引2）和E列（索引4）
      const bondCode = String(row[this.BOND_CODE_COL - 1] || '').trim();
      const bondName = String(row[this.BOND_NAME_COL - 1] || '').trim();
      const availableAmountStr = String(row[this.AVAILABLE_COL - 1] || '0').replace(/,/g, '');
      const availableAmount = parseFloat(availableAmountStr);

      if (bondName && availableAmount > 0) {
        // 判断债券类型：包含地理名词为地方债，否则为国债
        const bondActualType = containsGeographyKeyword(bondName) ? 'local' : 'treasury';
        
        this.bonds.push({
          rowNumber: rowIndex + 1, // 行号从1开始
          bondCode,
          bondName,
          availableAmount,
          rowData: row,
          bondActualType
        });
        
        // 输出前10只债券的类型判断结果
        if (this.bonds.length <= 10) {
          console.log(`  ${bondName} -> ${bondActualType === 'local' ? '地方债' : '国债'}`);
        }
      }
    }

    // 统计地方债和国债数量
    const localCount = this.bonds.filter(b => b.bondActualType === 'local').length;
    const treasuryCount = this.bonds.filter(b => b.bondActualType === 'treasury').length;
    
    console.log(`读取到 ${this.bonds.length} 条债券数据`);
    console.log(`其中：地方债 ${localCount} 条，国债 ${treasuryCount} 条`);
    this.result.totalRows = this.bonds.length;
  }

  /**
   * 多金额匹配债券
   * 根据用户选择的类型过滤债券
   * 排除被禁挑的债券（支持按金额序号指定）
   */
  private matchBondsByAmounts(): void {
    console.log('开始多金额匹配债券...');
    console.log(`用户选择类型: ${this.bondType === 'local' ? '地方债' : '国债'}`);

    // 第一步：根据用户选择的类型过滤债券
    const typeFilteredBonds = this.bonds.filter(bond => bond.bondActualType === this.bondType);
    this.result.typeFilteredCount = this.bonds.length - typeFilteredBonds.length;
    console.log(`类型过滤: 保留 ${typeFilteredBonds.length} 条${this.bondType === 'local' ? '地方债' : '国债'}，过滤 ${this.result.typeFilteredCount} 条其他类型`);

    // 计算总可用金额（类型过滤后）
    this.result.totalAvailable = typeFilteredBonds.reduce((sum, bond) => sum + bond.availableAmount, 0);
    console.log('类型过滤后总可用金额:', this.result.totalAvailable, '万元');

    if (typeFilteredBonds.length === 0) {
      console.warn('警告：没有符合条件的债券可挑选！');
      return;
    }

    // 按可用金额从大到小排序
    const sortedBonds = [...typeFilteredBonds].sort((a, b) => b.availableAmount - a.availableAmount);
    
    // 跟踪每只券的剩余可用金额
    const bondRemaining = new Map<BondData, number>();
    for (const bond of sortedBonds) {
      bondRemaining.set(bond, bond.availableAmount);
    }

    // 对每个金额进行处理
    for (let i = 0; i < this.amounts.length; i++) {
      const groupIndex = i + 1; // 金额序号从1开始
      const targetAmount = this.amounts[i];
      console.log(`\n处理第 ${groupIndex} 个金额: ${targetAmount} 万元`);

      const group: AmountGroup = {
        targetAmount,
        allocations: [],
        actualAmount: 0
      };

      let remainingToAllocate = targetAmount;

      // 按可用金额从大到小遍历债券
      for (const bond of sortedBonds) {
        if (remainingToAllocate <= 0) break;

        const bondRemainingAmount = bondRemaining.get(bond) || 0;
        if (bondRemainingAmount <= 0) continue;

        // 检查该债券是否在当前金额序号下被禁挑
        if (this.isBondExcluded(bond, groupIndex)) {
          continue;
        }

        // 计算可以分配的金额
        let allocateAmount = 0;

        if (bondRemainingAmount <= remainingToAllocate) {
          // 当前债券剩余金额 <= 还需分配的金额，可以全部占用
          allocateAmount = bondRemainingAmount;
        } else {
          // 当前债券剩余金额 > 还需分配的金额，部分占用
          allocateAmount = remainingToAllocate;
        }

        // 检查剩余部分是否满足最小占用金额
        const leftover = bondRemainingAmount - allocateAmount;
        if (leftover > 0 && leftover < this.MIN_OCCUPY_AMOUNT) {
          // 如果剩余部分不足100万，则调整分配金额
          if (remainingToAllocate === targetAmount && bondRemainingAmount >= targetAmount + this.MIN_OCCUPY_AMOUNT) {
            // 如果是第一只券且足够大，可以多分配一些
            allocateAmount = targetAmount;
          } else if (leftover + remainingToAllocate <= bondRemainingAmount) {
            // 尝试多分配一些，让剩余部分满足最小占用
            allocateAmount = remainingToAllocate;
            const newLeftover = bondRemainingAmount - allocateAmount;
            if (newLeftover > 0 && newLeftover < this.MIN_OCCUPY_AMOUNT) {
              // 还是不能满足，跳过这只券
              continue;
            }
          } else {
            continue;
          }
        }

        if (allocateAmount >= this.MIN_OCCUPY_AMOUNT || allocateAmount === remainingToAllocate) {
          group.allocations.push({
            bond,
            allocatedAmount: allocateAmount
          });
          group.actualAmount += allocateAmount;
          remainingToAllocate -= allocateAmount;
          bondRemaining.set(bond, bondRemainingAmount - allocateAmount);

          console.log(`  挑选: ${bond.bondName}(${bond.bondCode}), 剩余${bondRemainingAmount}万, 分配${allocateAmount}万, 还需${remainingToAllocate}万`);
        }
      }

      this.result.groups.push(group);

      if (remainingToAllocate > 0) {
        console.warn(`警告：第 ${i + 1} 个金额未能完全匹配，还剩 ${remainingToAllocate} 万元未挑`);
      }
    }
  }

  /**
   * 生成结果Excel
   */
  private async generateResultWorkbook(): Promise<void> {
    console.log('生成结果Excel...');

    // 创建新工作表
    const resultSheet = this.workbook.addWorksheet('挑券结果');

    // 写入前7行（标题等）
    for (let rowIdx = 0; rowIdx < this.HEADER_ROW - 1; rowIdx++) {
      const sourceRow = this.sourceSheetData[rowIdx] || [];
      const newRow = resultSheet.getRow(rowIdx + 1);
      
      for (let colIdx = 0; colIdx < sourceRow.length; colIdx++) {
        if (colIdx < this.AVAILABLE_COL) {
          // A-E列直接复制
          newRow.getCell(colIdx + 1).value = sourceRow[colIdx];
        } else {
          // F列及后面的列往右移一格
          newRow.getCell(colIdx + 2).value = sourceRow[colIdx];
        }
      }
      
      // 第7行F列显示第一个金额
      if (rowIdx === 6 && this.result.groups.length > 0) {
        newRow.getCell(this.AVAILABLE_COL + 1).value = this.result.groups[0].targetAmount;
        console.log(`F列第7行设置第一个金额: ${this.result.groups[0].targetAmount}万元`);
      }
      
      newRow.commit();
    }

    // 写入表头行（第8行）
    const headerRow = this.sourceSheetData[this.HEADER_ROW - 1] || [];
    const newHeaderRow = resultSheet.getRow(this.HEADER_ROW);
    
    for (let colIdx = 0; colIdx < Math.max(headerRow.length, this.AVAILABLE_COL); colIdx++) {
      if (colIdx < this.AVAILABLE_COL) {
        newHeaderRow.getCell(colIdx + 1).value = headerRow[colIdx] || '';
      } else if (colIdx === this.AVAILABLE_COL - 1) {
        // E列
        newHeaderRow.getCell(colIdx + 1).value = headerRow[colIdx] || '可用金额';
      }
    }
    
    // F列插入"挑券金额（万元）"
    newHeaderRow.getCell(this.AVAILABLE_COL + 1).value = '挑券金额（万元）';
    
    // G列及之后
    for (let colIdx = this.AVAILABLE_COL - 1; colIdx < headerRow.length; colIdx++) {
      newHeaderRow.getCell(colIdx + 3).value = headerRow[colIdx];
    }
    
    newHeaderRow.commit();

    // 写入债券数据
    let currentRowNumber = this.DATA_START_ROW;

    for (let groupIndex = 0; groupIndex < this.result.groups.length; groupIndex++) {
      const group = this.result.groups[groupIndex];
      console.log(`写入第 ${groupIndex + 1} 个债券集合，目标金额: ${group.targetAmount}万元`);

      // 写入该集合的债券数据
      for (const allocation of group.allocations) {
        const newRow = resultSheet.getRow(currentRowNumber);
        const rowData = allocation.bond.rowData;
        
        // A-E列
        for (let colIdx = 0; colIdx < this.AVAILABLE_COL; colIdx++) {
          newRow.getCell(colIdx + 1).value = rowData[colIdx] || '';
        }
        
        // F列：挑券金额
        newRow.getCell(this.AVAILABLE_COL + 1).value = allocation.allocatedAmount;
        
        // G列及之后：原F列及之后的数据
        for (let colIdx = this.AVAILABLE_COL - 1; colIdx < rowData.length; colIdx++) {
          newRow.getCell(colIdx + 3).value = rowData[colIdx];
        }
        
        newRow.commit();
        currentRowNumber++;
      }

      // 如果不是最后一个集合，插入空行并显示下一个金额
      if (groupIndex < this.result.groups.length - 1) {
        const nextGroup = this.result.groups[groupIndex + 1];
        const emptyRow = resultSheet.getRow(currentRowNumber);
        
        // 空行F列显示下一个金额
        emptyRow.getCell(this.AVAILABLE_COL + 1).value = nextGroup.targetAmount;
        emptyRow.commit();
        
        console.log(`插入空行，F列显示下一个金额: ${nextGroup.targetAmount}万元`);
        currentRowNumber++;
      }
    }

    console.log('结果Excel生成完成，总行数:', currentRowNumber - 1);
  }
}
