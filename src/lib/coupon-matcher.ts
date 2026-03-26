import ExcelJS from 'exceljs';
import { containsGeographyKeyword } from './geo-keywords';

/**
 * 债券数据结构
 */
interface BondData {
  rowNumber: number;       // 原始行号
  bondName: string;        // 债券名称（C列）
  availableAmount: number; // 可用金额（E列，万元）
  selectedAmount: number;  // 挑券金额（计算得出）
  rowData: any[];          // 整行数据
}

/**
 * 挑券模式处理器
 * 完全独立于其他功能模块
 * 
 * Excel结构说明：
 * - 第1-7行：其他内容（标题、说明等）
 * - 第8行：字段列名（表头）
 * - 第9行开始：债券数据
 * - C列：债券名称
 * - E列：可用金额
 */
export class CouponMatcher {
  private file: File | Buffer;
  private bondType: 'treasury' | 'local'; // 用户选择的债券类型
  private amount: number; // 挑券金额（万元）
  private workbook: ExcelJS.Workbook;
  private worksheet: ExcelJS.Worksheet | null = null;

  // Excel结构常量
  private readonly HEADER_ROW = 8;      // 表头行号
  private readonly DATA_START_ROW = 9;  // 数据起始行号
  private readonly BOND_NAME_COL = 3;   // C列 - 债券名称
  private readonly AVAILABLE_COL = 5;   // E列 - 可用金额

  // 债券数据列表
  private bonds: BondData[] = [];

  // 处理结果
  private result: {
    bondType: 'treasury' | 'local';
    totalRows: number;
    totalAvailable: number;
    totalSelected: number;
    selectedBonds: BondData[];
  } = {
    bondType: 'treasury',
    totalRows: 0,
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

    // 2. 读取债券数据（从第9行开始）
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
      
      // 读取C列（债券名称）和E列（可用金额）
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
          bondName,
          availableAmount,
          selectedAmount: 0,
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
        const afterSelect = remainingAvailable - bond.availableAmount;
        
        if (afterSelect >= this.MIN_OCCUPY_AMOUNT || afterSelect === 0) {
          selectAmount = bond.availableAmount;
        } else {
          selectAmount = bond.availableAmount - this.MIN_OCCUPY_AMOUNT;
          if (selectAmount < this.MIN_OCCUPY_AMOUNT) {
            continue;
          }
        }
      } else {
        // 当前债券可用金额 > 剩余挑券金额
        selectAmount = remainingAmount;
        
        const leftover = bond.availableAmount - selectAmount;
        if (leftover > 0 && leftover < this.MIN_OCCUPY_AMOUNT) {
          if (remainingAmount === this.amount) {
            selectAmount = bond.availableAmount;
          } else {
            continue;
          }
        }
      }

      if (selectAmount > remainingAmount) {
        selectAmount = remainingAmount;
      }

      if (selectAmount < this.MIN_OCCUPY_AMOUNT && remainingAmount >= this.MIN_OCCUPY_AMOUNT) {
        continue;
      }

      if (selectAmount > 0) {
        bond.selectedAmount = selectAmount;
        selectedBonds.push(bond);
        remainingAmount -= selectAmount;
        
        console.log(`  挑选: ${bond.bondName}, 可用${bond.availableAmount}万, 挑${selectAmount}万, 剩余需挑${remainingAmount}万`);
      }
    }

    this.result.selectedBonds = selectedBonds;
    this.result.totalSelected = this.amount - remainingAmount;

    if (remainingAmount > 0) {
      console.warn(`警告：未能完全匹配，还剩 ${remainingAmount} 万元未挑`);
    }
  }

  /**
   * 生成结果Excel
   * 1. F列插入"挑券金额（万元）"，原F列及后面的列往右移
   * 2. 债券按E列可用金额从大到小排序
   * 3. F列第7行显示挑券完成后的总金额
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
      
      // F列第7行显示挑券总金额
      if (rowNumber === 7) {
        newRow.getCell(this.AVAILABLE_COL + 1).value = this.result.totalSelected;
        console.log(`F列第7行设置挑券总金额: ${this.result.totalSelected}万元`);
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

    // 按可用金额从大到小排序选中的债券，写入数据行
    const sortedBonds = [...this.result.selectedBonds].sort((a, b) => b.availableAmount - a.availableAmount);
    
    for (let i = 0; i < sortedBonds.length; i++) {
      const bond = sortedBonds[i];
      const newRowNumber = this.DATA_START_ROW + i;
      const newRow = resultSheet.getRow(newRowNumber);
      
      // 复制原数据（A-E列）
      for (let colIndex = 0; colIndex < bond.rowData.length; colIndex++) {
        const colNumber = colIndex + 1;
        if (colNumber < this.AVAILABLE_COL + 1) {
          newRow.getCell(colNumber).value = bond.rowData[colIndex];
        } else if (colNumber === this.AVAILABLE_COL) {
          // E列
          newRow.getCell(colNumber).value = bond.availableAmount;
        }
      }
      
      // F列：挑券金额
      newRow.getCell(this.AVAILABLE_COL + 1).value = bond.selectedAmount;
      
      // G列及之后：原F列及之后的数据
      for (let colIndex = this.AVAILABLE_COL; colIndex < bond.rowData.length; colIndex++) {
        newRow.getCell(colIndex + 2).value = bond.rowData[colIndex];
      }
      
      newRow.commit();
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
