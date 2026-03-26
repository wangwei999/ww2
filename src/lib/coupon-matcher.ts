import ExcelJS from 'exceljs';
import { containsGeographyKeyword } from './geo-keywords';

/**
 * 债券数据结构
 */
interface BondData {
  rowNumber: number;       // 行号
  bondName: string;        // 债券名称（C列）
  availableAmount: number; // 可用金额（E列，万元）
  selectedAmount: number;  // 挑券金额（计算得出）
  // 其他列数据
  [key: string]: any;
}

/**
 * 挑券模式处理器
 * 完全独立于其他功能模块
 */
export class CouponMatcher {
  private file: File | Buffer;
  private bondType: 'treasury' | 'local'; // 用户选择的债券类型
  private amount: number; // 挑券金额（万元）
  private workbook: ExcelJS.Workbook;
  private worksheet: ExcelJS.Worksheet | null = null;

  // 债券数据列表
  private bonds: BondData[] = [];

  // 处理结果
  private result: {
    bondType: 'treasury' | 'local'; // 实际判断的债券类型
    totalRows: number;
    filteredRows: number;
    totalAvailable: number;
    totalSelected: number;
    selectedBonds: BondData[];
  } = {
    bondType: 'treasury',
    totalRows: 0,
    filteredRows: 0,
    totalAvailable: 0,
    totalSelected: 0,
    selectedBonds: []
  };

  // 最小占用金额（万元）
  private readonly MIN_OCCUPY_AMOUNT = 100;

  constructor(
    file: File | Buffer,
    bondType: 'treasury' | 'local',
    amount: number
  ) {
    this.file = file;
    this.bondType = bondType;
    this.amount = amount;
    this.workbook = new ExcelJS.Workbook();
  }

  /**
   * 主处理方法
   */
  async process(): Promise<{ workbook: ExcelJS.Workbook; statistics: any }> {
    console.log('=== 开始挑券处理 ===');
    console.log('用户选择类型:', this.bondType);
    console.log('挑券金额:', this.amount, '万元');

    // 1. 加载Excel文件
    await this.loadExcelFile();

    // 2. 读取债券数据
    this.readBondData();

    // 3. 判断债券类型（如果用户选择地方债）
    if (this.bondType === 'local') {
      this.determineBondType();
    } else {
      this.result.bondType = 'treasury';
    }
    console.log('实际债券类型:', this.result.bondType);

    // 4. 根据金额匹配债券
    this.matchBondsByAmount();

    // 5. 生成结果Excel
    await this.generateResultWorkbook();

    console.log('=== 挑券处理完成 ===');
    console.log('挑券统计:', {
      总可用金额: this.result.totalAvailable,
      挑券金额: this.amount,
      实际挑券: this.result.totalSelected,
      挑券数量: this.result.selectedBonds.length
    });

    return {
      workbook: this.workbook,
      statistics: {
        bondType: this.result.bondType,
        totalRows: this.result.totalRows,
        filteredRows: this.result.filteredRows,
        totalAvailable: this.result.totalAvailable,
        totalSelected: this.result.totalSelected,
        selectedCount: this.result.selectedBonds.length,
        requestedAmount: this.amount
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
    } else {
      buffer = this.file;
    }

    await this.workbook.xlsx.load(buffer as any);

    this.worksheet = this.workbook.worksheets[0];
    if (!this.worksheet) {
      throw new Error('Excel文件中没有工作表');
    }

    console.log('工作表名称:', this.worksheet.name);
  }

  /**
   * 读取债券数据
   */
  private readBondData(): void {
    console.log('读取债券数据...');

    if (!this.worksheet) return;

    // 从第2行开始读取（跳过表头）
    this.worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // 跳过表头

      // 读取C列（债券名称）和E列（可用金额）
      const cellC = row.getCell(3);
      const cellE = row.getCell(5);
      
      const bondName = String(cellC.value || '').trim();
      const availableAmount = parseFloat(String(cellE.value || '0').replace(/,/g, ''));

      if (bondName && availableAmount > 0) {
        this.bonds.push({
          rowNumber,
          bondName,
          availableAmount,
          selectedAmount: 0
        });
      }
    });

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
    let checkedRows = 0;

    for (const bond of this.bonds) {
      checkedRows++;
      if (containsGeographyKeyword(bond.bondName)) {
        hasGeographyKeyword = true;
        console.log(`  "${bond.bondName}" 包含地理名词`);
      }
    }

    this.result.bondType = hasGeographyKeyword ? 'local' : 'treasury';
    console.log(`检查了 ${checkedRows} 行，包含地理名词: ${hasGeographyKeyword}`);
  }

  /**
   * 根据金额匹配债券
   * 规则：
   * 1. 按E列可用金额从大到小排序
   * 2. 优先占用金额大的债券
   * 3. 最小占用金额为100万
   */
  private matchBondsByAmount(): void {
    console.log('开始匹配债券，挑券金额:', this.amount, '万元');

    // 按可用金额从大到小排序
    const sortedBonds = [...this.bonds].sort((a, b) => b.availableAmount - a.availableAmount);

    // 计算总可用金额
    this.result.totalAvailable = sortedBonds.reduce((sum, bond) => sum + bond.availableAmount, 0);
    console.log('总可用金额:', this.result.totalAvailable, '万元');

    // 检查是否足够
    if (this.result.totalAvailable < this.amount) {
      console.warn('警告：总可用金额不足！');
    }

    let remainingAmount = this.amount;
    const selectedBonds: BondData[] = [];

    for (let i = 0; i < sortedBonds.length && remainingAmount > 0; i++) {
      const bond = sortedBonds[i];
      
      // 计算剩余债券的总可用金额
      let remainingAvailable = 0;
      for (let j = i; j < sortedBonds.length; j++) {
        remainingAvailable += sortedBonds[j].availableAmount;
      }

      // 计算可以挑的金额
      let selectAmount = 0;

      if (bond.availableAmount <= remainingAmount) {
        // 当前债券可用金额 <= 剩余挑券金额
        // 可以挑全部或部分
        const afterSelect = remainingAvailable - bond.availableAmount;
        
        if (afterSelect >= this.MIN_OCCUPY_AMOUNT || afterSelect === 0) {
          // 挑完后剩余的金额足够满足最小占用要求，或者没有剩余了
          // 可以挑这个债券的全部金额
          selectAmount = bond.availableAmount;
        } else {
          // 需要给后续债券留够最小占用金额
          // 这个债券只能挑：可用金额 - 最小占用金额
          selectAmount = bond.availableAmount - this.MIN_OCCUPY_AMOUNT;
          if (selectAmount < this.MIN_OCCUPY_AMOUNT) {
            // 如果挑完连最小占用都不够，就跳过这个债券
            continue;
          }
        }
      } else {
        // 当前债券可用金额 > 剩余挑券金额
        // 只需要挑剩余的金额
        selectAmount = remainingAmount;
        
        // 检查挑完后剩余金额是否满足最小占用要求
        const leftover = bond.availableAmount - selectAmount;
        if (leftover > 0 && leftover < this.MIN_OCCUPY_AMOUNT) {
          // 剩余金额不足最小占用，需要调整
          // 方案1：多挑一点，让剩余为0
          // 方案2：少挑一点，让剩余>=100万
          if (remainingAmount === this.amount) {
            // 这是第一个债券，可以多挑
            selectAmount = bond.availableAmount;
          } else {
            // 不是第一个债券，需要确保之前的债券已经挑了
            // 这种情况下，跳过当前债券，尝试下一个
            continue;
          }
        }
      }

      // 确保挑券金额不超过剩余需要挑的金额
      if (selectAmount > remainingAmount) {
        selectAmount = remainingAmount;
      }

      // 确保挑券金额 >= 最小占用金额（除非是最后一个债券且金额不足）
      if (selectAmount < this.MIN_OCCUPY_AMOUNT && remainingAmount >= this.MIN_OCCUPY_AMOUNT) {
        continue;
      }

      // 应用选择
      if (selectAmount > 0) {
        bond.selectedAmount = selectAmount;
        selectedBonds.push(bond);
        remainingAmount -= selectAmount;
        
        console.log(`  挑选: ${bond.bondName}, 可用${bond.availableAmount}万, 挑${selectAmount}万, 剩余需挑${remainingAmount}万`);
      }
    }

    this.result.selectedBonds = selectedBonds;
    this.result.totalSelected = this.amount - remainingAmount;
    this.result.filteredRows = selectedBonds.length;

    if (remainingAmount > 0) {
      console.warn(`警告：未能完全匹配，还剩 ${remainingAmount} 万元未挑`);
    }
  }

  /**
   * 生成结果Excel
   */
  private async generateResultWorkbook(): Promise<void> {
    console.log('生成结果Excel...');

    if (!this.worksheet) return;

    // 创建新的工作簿
    const resultWorkbook = new ExcelJS.Workbook();
    const resultSheet = resultWorkbook.addWorksheet('挑券结果');

    // 复制表头
    const headerRow = this.worksheet.getRow(1);
    const newHeaderRow = resultSheet.getRow(1);
    headerRow.eachCell((cell, colNumber) => {
      newHeaderRow.getCell(colNumber).value = cell.value;
    });

    // 添加"挑券金额"列
    const lastCol = headerRow.cellCount + 1;
    newHeaderRow.getCell(lastCol).value = '挑券金额（万元）';
    newHeaderRow.commit();

    // 复制选中的债券数据
    for (const bond of this.result.selectedBonds) {
      const sourceRow = this.worksheet.getRow(bond.rowNumber);
      const newRow = resultSheet.addRow([]);
      
      sourceRow.eachCell((cell, colNumber) => {
        newRow.getCell(colNumber).value = cell.value;
      });
      
      // 添加挑券金额
      newRow.getCell(lastCol).value = bond.selectedAmount;
    }

    // 更新工作簿引用
    this.workbook = resultWorkbook;
  }
}
