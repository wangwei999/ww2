import ExcelJS from 'exceljs';
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
  original: string;    // 原始输入
  type: 'code' | 'name'; // 类型：代码精确匹配 / 名称模糊匹配
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
  private worksheet: ExcelJS.Worksheet | null = null;

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
    groups: AmountGroup[];
  } = {
    bondType: 'treasury',
    totalRows: 0,
    totalAvailable: 0,
    excludedCount: 0,
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
   * 数字 -> 代码精确匹配
   * 文字 -> 名称模糊匹配
   */
  private parseExclusionRules(): void {
    this.exclusionRules = [];
    
    for (const item of this.excludedBonds) {
      const trimmed = item.trim();
      if (!trimmed) continue;
      
      // 判断是纯数字还是包含文字
      // 纯数字：精确匹配代码
      // 包含文字：模糊匹配名称
      const isNumeric = /^\d+$/.test(trimmed);
      
      this.exclusionRules.push({
        original: trimmed,
        type: isNumeric ? 'code' : 'name'
      });
    }
    
    if (this.exclusionRules.length > 0) {
      console.log('禁挑券规则:', this.exclusionRules.map(r => 
        `${r.original}(${r.type === 'code' ? '代码精确' : '名称模糊'})`
      ).join(', '));
    }
  }

  /**
   * 检查债券是否被禁挑
   */
  private isBondExcluded(bond: BondData): boolean {
    for (const rule of this.exclusionRules) {
      if (rule.type === 'code') {
        // 代码精确匹配
        if (bond.bondCode === rule.original) {
          console.log(`  禁挑: ${bond.bondName}(${bond.bondCode}) - 代码精确匹配"${rule.original}"`);
          return true;
        }
      } else {
        // 名称模糊匹配（包含）
        if (bond.bondName.includes(rule.original)) {
          console.log(`  禁挑: ${bond.bondName}(${bond.bondCode}) - 名称包含"${rule.original}"`);
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
    console.log('用户选择类型:', this.bondType);
    console.log('挑券金额:', this.amounts, '万元');
    console.log('模式:', this.amounts.length > 1 ? '多金额' : '单金额');
    console.log('禁挑券数量:', this.exclusionRules.length);

    // 1. 加载Excel文件
    await this.loadExcelFile();

    // 2. 读取债券数据（从第9行开始）
    this.readBondData();

    // 3. 判断债券类型（如果用户选择地方债）
    if (this.bondType === 'local') {
      this.determineBondType();
    } else {
      this.result.bondType = 'treasury';
    }
    console.log('实际债券类型:', this.result.bondType);

    // 4. 根据金额匹配债券（支持多金额，排除禁挑券）
    this.matchBondsByAmounts();

    // 5. 生成结果Excel
    await this.generateResultWorkbook();

    console.log('=== 挑券处理完成 ===');
    console.log('挑券统计:', {
      总可用金额: this.result.totalAvailable,
      总挑券金额: this.amounts.reduce((a, b) => a + b, 0),
      禁挑数量: this.result.excludedCount,
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
        bondType: this.result.bondType,
        totalRows: this.result.totalRows,
        totalAvailable: this.result.totalAvailable,
        excludedCount: this.result.excludedCount,
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
   */
  private async loadExcelFile(): Promise<void> {
    console.log('加载Excel文件...');

    let buffer: Buffer;
    if (this.file instanceof File) {
      const arrayBuffer = await this.file.arrayBuffer();
      buffer = Buffer.from(arrayBuffer);
      console.log('文件名:', this.file.name, '大小:', buffer.length, 'bytes');
    } else {
      buffer = this.file;
      console.log('Buffer大小:', buffer.length, 'bytes');
    }

    try {
      await this.workbook.xlsx.load(buffer as any);
    } catch (e: any) {
      console.error('Excel加载失败:', e.message);
      // 可能是 .xls 格式不支持
      throw new Error('Excel文件加载失败，请确保上传的是 .xlsx 格式（不支持旧版 .xls 格式）');
    }

    const worksheetCount = this.workbook.worksheets.length;
    console.log('工作表数量:', worksheetCount);
    
    if (worksheetCount === 0) {
      throw new Error('Excel文件中没有工作表，请检查文件内容是否正确');
    }

    this.worksheet = this.workbook.worksheets[0];
    console.log('工作表名称:', this.worksheet.name);
  }

  /**
   * 读取债券数据
   * 从第9行开始读取（第8行是表头）
   */
  private readBondData(): void {
    console.log('读取债券数据（从第9行开始）...');

    if (!this.worksheet) return;

    const rowCount = this.worksheet.rowCount;
    console.log('工作表总行数:', rowCount);

    // 从第9行开始读取数据
    for (let rowNumber = this.DATA_START_ROW; rowNumber <= rowCount; rowNumber++) {
      const row = this.worksheet.getRow(rowNumber);
      
      // 读取B列（债券代码）、C列（债券名称）和E列（可用金额）
      const bondCode = String(row.getCell(this.BOND_CODE_COL).value || '').trim();
      const bondName = String(row.getCell(this.BOND_NAME_COL).value || '').trim();
      const availableAmountStr = String(row.getCell(this.AVAILABLE_COL).value || '0').replace(/,/g, '');
      const availableAmount = parseFloat(availableAmountStr);

      if (bondName && availableAmount > 0) {
        // 保存整行数据
        const rowData: any[] = [];
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          rowData[colNumber - 1] = cell.value;
        });

        this.bonds.push({
          rowNumber,
          bondCode,
          bondName,
          availableAmount,
          rowData
        });
      }
    }

    console.log(`读取到 ${this.bonds.length} 条债券数据`);
    this.result.totalRows = this.bonds.length;
  }

  /**
   * 判断债券类型
   * 规则：读取C列，如果包含地理名词则为地方债，否则为国债
   */
  private determineBondType(): void {
    console.log('判断债券类型（读取C列）...');

    let hasGeographyKeyword = false;

    for (const bond of this.bonds) {
      if (containsGeographyKeyword(bond.bondName)) {
        hasGeographyKeyword = true;
        console.log(`  "${bond.bondName}" 包含地理名词`);
      }
    }

    this.result.bondType = hasGeographyKeyword ? 'local' : 'treasury';
    console.log(`检查了 ${this.bonds.length} 行，包含地理名词: ${hasGeographyKeyword}`);
  }

  /**
   * 多金额匹配债券
   * 每个金额形成一个债券集合，集合间可以拆分同一只券
   * 排除被禁挑的债券
   */
  private matchBondsByAmounts(): void {
    console.log('开始多金额匹配债券...');

    // 过滤掉被禁挑的债券
    const availableBonds = this.bonds.filter(bond => !this.isBondExcluded(bond));
    this.result.excludedCount = this.bonds.length - availableBonds.length;
    
    console.log(`可用债券: ${availableBonds.length} 条，禁挑: ${this.result.excludedCount} 条`);

    // 计算总可用金额
    this.result.totalAvailable = availableBonds.reduce((sum, bond) => sum + bond.availableAmount, 0);
    console.log('总可用金额:', this.result.totalAvailable, '万元');

    // 按可用金额从大到小排序（全局排序一次）
    const sortedBonds = [...availableBonds].sort((a, b) => b.availableAmount - a.availableAmount);
    
    // 跟踪每只券的剩余可用金额
    const bondRemaining = new Map<BondData, number>();
    for (const bond of sortedBonds) {
      bondRemaining.set(bond, bond.availableAmount);
    }

    // 对每个金额进行处理
    for (let i = 0; i < this.amounts.length; i++) {
      const targetAmount = this.amounts[i];
      console.log(`\n处理第 ${i + 1} 个金额: ${targetAmount} 万元`);

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
   * 
   * 单金额模式：
   * - 第7行F列显示挑券总金额
   * - 数据从第9行开始
   * 
   * 多金额模式：
   * - 第7行F列显示第一个金额
   * - 每个集合间空一行
   * - 空行F列显示下一个金额
   * - 被拆分的券在下一个集合中复制字段信息
   */
  private async generateResultWorkbook(): Promise<void> {
    console.log('生成结果Excel...');

    if (!this.worksheet) return;

    // 创建新的工作簿
    const resultWorkbook = new ExcelJS.Workbook();
    const resultSheet = resultWorkbook.addWorksheet(this.worksheet.name);

    // 复制第1-7行（保持原样，但要插入F列）
    for (let rowNumber = 1; rowNumber < this.HEADER_ROW; rowNumber++) {
      const sourceRow = this.worksheet.getRow(rowNumber);
      const newRow = resultSheet.getRow(rowNumber);
      
      sourceRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (colNumber < this.AVAILABLE_COL + 1) {
          // A-E列直接复制
          newRow.getCell(colNumber).value = cell.value;
          this.copyCellStyle(cell, newRow.getCell(colNumber));
        } else {
          // F列及后面的列往右移一格
          newRow.getCell(colNumber + 1).value = cell.value;
          this.copyCellStyle(cell, newRow.getCell(colNumber + 1));
        }
      });
      
      // 第7行F列显示第一个金额
      if (rowNumber === 7 && this.result.groups.length > 0) {
        newRow.getCell(this.AVAILABLE_COL + 1).value = this.result.groups[0].targetAmount;
        console.log(`F列第7行设置第一个金额: ${this.result.groups[0].targetAmount}万元`);
      }
      
      newRow.commit();
    }

    // 复制第8行（表头行），插入"挑券金额（万元）"列
    const sourceHeaderRow = this.worksheet.getRow(this.HEADER_ROW);
    const newHeaderRow = resultSheet.getRow(this.HEADER_ROW);
    
    sourceHeaderRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      if (colNumber < this.AVAILABLE_COL + 1) {
        newHeaderRow.getCell(colNumber).value = cell.value;
        this.copyCellStyle(cell, newHeaderRow.getCell(colNumber));
      } else {
        newHeaderRow.getCell(colNumber + 1).value = cell.value;
        this.copyCellStyle(cell, newHeaderRow.getCell(colNumber + 1));
      }
    });
    
    // F列插入"挑券金额（万元）"
    newHeaderRow.getCell(this.AVAILABLE_COL + 1).value = '挑券金额（万元）';
    newHeaderRow.commit();

    // 写入债券数据
    let currentRowNumber = this.DATA_START_ROW;

    for (let groupIndex = 0; groupIndex < this.result.groups.length; groupIndex++) {
      const group = this.result.groups[groupIndex];
      console.log(`写入第 ${groupIndex + 1} 个债券集合，目标金额: ${group.targetAmount}万元`);

      // 写入该集合的债券数据
      for (const allocation of group.allocations) {
        const newRow = resultSheet.getRow(currentRowNumber);
        
        // 复制原数据（A-E列）
        for (let colIndex = 0; colIndex < allocation.bond.rowData.length; colIndex++) {
          const colNumber = colIndex + 1;
          if (colNumber < this.AVAILABLE_COL + 1) {
            newRow.getCell(colNumber).value = allocation.bond.rowData[colIndex];
          } else if (colNumber === this.AVAILABLE_COL) {
            // E列：可用金额
            newRow.getCell(colNumber).value = allocation.bond.availableAmount;
          }
        }
        
        // F列：挑券金额
        newRow.getCell(this.AVAILABLE_COL + 1).value = allocation.allocatedAmount;
        
        // G列及之后：原F列及之后的数据
        for (let colIndex = this.AVAILABLE_COL; colIndex < allocation.bond.rowData.length; colIndex++) {
          newRow.getCell(colIndex + 2).value = allocation.bond.rowData[colIndex];
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
        
        // 可选：设置空行的样式（如背景色）以区分
        emptyRow.commit();
        
        console.log(`插入空行，F列显示下一个金额: ${nextGroup.targetAmount}万元`);
        currentRowNumber++;
      }
    }

    this.workbook = resultWorkbook;
  }

  /**
   * 复制单元格样式
   */
  private copyCellStyle(sourceCell: ExcelJS.Cell, targetCell: ExcelJS.Cell): void {
    try {
      if (sourceCell.font) targetCell.font = { ...sourceCell.font };
      if (sourceCell.fill) targetCell.fill = { ...sourceCell.fill };
      if (sourceCell.border) targetCell.border = { ...sourceCell.border };
      if (sourceCell.alignment) targetCell.alignment = { ...sourceCell.alignment };
      if (sourceCell.numFmt) targetCell.numFmt = sourceCell.numFmt;
    } catch (e) {
      // 忽略样式复制错误
    }
  }
}
