import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { containsGeographyKeyword } from './geo-keywords';

/**
 * 债券分类数据
 */
interface BondData {
  rowNumber: number;       // 原始行号
  bondCode: string;        // 债券代码（B列）
  bondName: string;        // 债券名称（C列）
  availableAmount: number; // 可用金额（E列，万元）
  rowData: any[];          // 整行数据
  category: 'treasury' | 'policy' | 'local' | 'credit'; // 债券分类：国债/政金债/地方债/信用债
}

/**
 * 债券副本处理器
 * 上传文件后自动生成四个分类工作表：国债、政金债、地方债、信用债
 * 
 * 分类规则：
 * - 地方债：C列（债券名称）包含地理名词（省-市-县-区）
 * - 政金债：C列包含"国开"、"进出口"、"农发"
 * - 国债：既不是地方债也不是政金债，且债券简称含"国债"字样
 * - 信用债：既不是地方债也不是政金债，且债券简称不含"国债"字样
 * 
 * 排序规则：按可用金额（E列）从大到小排列
 */
export class CouponCopyProcessor {
  private file: File | Buffer;
  private workbook: ExcelJS.Workbook;
  private sourceSheet: ExcelJS.Worksheet | null = null;
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

  // 处理结果
  private result: {
    totalRows: number;
    treasuryCount: number;
    policyCount: number;
    localCount: number;
    creditCount: number;
  } = {
    totalRows: 0,
    treasuryCount: 0,
    policyCount: 0,
    localCount: 0,
    creditCount: 0
  };

  constructor(file: File | Buffer) {
    this.file = file;
    this.workbook = new ExcelJS.Workbook();
  }

  /**
   * 执行处理
   */
  async process(): Promise<{ workbook: ExcelJS.Workbook; statistics: any }> {
    console.log('=== 开始债券副本分类处理 ===');

    // 1. 加载Excel文件
    await this.loadExcel();

    // 2. 读取债券数据并分类
    this.readAndClassifyBonds();

    // 3. 创建四个分类工作表
    this.createClassificationSheets();

    // 4. 统计信息
    this.result.treasuryCount = this.bonds.filter(b => b.category === 'treasury').length;
    this.result.policyCount = this.bonds.filter(b => b.category === 'policy').length;
    this.result.localCount = this.bonds.filter(b => b.category === 'local').length;
    this.result.creditCount = this.bonds.filter(b => b.category === 'credit').length;

    console.log('处理完成:', {
      总债券数: this.result.totalRows,
      国债数: this.result.treasuryCount,
      政金债数: this.result.policyCount,
      地方债数: this.result.localCount,
      信用债数: this.result.creditCount
    });

    return {
      workbook: this.workbook,
      statistics: this.result
    };
  }

  /**
   * 加载Excel文件
   */
  private async loadExcel(): Promise<void> {
    console.log('加载Excel文件...');
    
    let buffer: Buffer;
    if (this.file instanceof File) {
      const arrayBuffer = await this.file.arrayBuffer();
      buffer = Buffer.from(arrayBuffer);
    } else {
      buffer = this.file;
    }

    // 判断文件格式
    const fileName = this.file instanceof File ? this.file.name.toLowerCase() : '';
    this.isXlsFormat = fileName.endsWith('.xls');

    if (this.isXlsFormat) {
      // .xls 格式：先用 xlsx 解析为 JSON 数据
      console.log('检测到 .xls 格式，使用 xlsx 库解析...');
      const xlsxWorkbook = XLSX.read(buffer as any, { type: 'buffer' });
      const sheetName = xlsxWorkbook.SheetNames[0];
      const worksheet = xlsxWorkbook.Sheets[sheetName];
      this.sourceSheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
      console.log(`解析完成，共 ${this.sourceSheetData.length} 行数据`);
    } else {
      // .xlsx 格式：使用 ExcelJS 解析
      console.log('检测到 .xlsx 格式，使用 ExcelJS 解析...');
      await this.workbook.xlsx.load(buffer as any);
      this.sourceSheet = this.workbook.worksheets[0];
      
      // 读取所有数据到内存
      this.sourceSheetData = [];
      this.sourceSheet.eachRow((row, rowNumber) => {
        const rowData: any[] = [];
        row.eachCell((cell, colNumber) => {
          rowData[colNumber - 1] = cell.value;
        });
        this.sourceSheetData[rowNumber - 1] = rowData;
      });
      console.log(`解析完成，共 ${this.sourceSheetData.length} 行数据`);
    }
  }

  /**
   * 读取债券数据并分类
   */
  private readAndClassifyBonds(): void {
    console.log('读取并分类债券数据...');

    if (this.sourceSheetData.length === 0) {
      console.warn('没有数据可读取');
      return;
    }

    // 从第9行开始读取数据
    for (let rowIndex = this.DATA_START_ROW - 1; rowIndex < this.sourceSheetData.length; rowIndex++) {
      const row = this.sourceSheetData[rowIndex];
      if (!row) continue;

      const bondCode = String(row[this.BOND_CODE_COL - 1] || '').trim();
      const bondName = String(row[this.BOND_NAME_COL - 1] || '').trim();
      const availableAmountStr = String(row[this.AVAILABLE_COL - 1] || '0').replace(/,/g, '');
      const availableAmount = parseFloat(availableAmountStr);

      if (bondName && availableAmount > 0) {
        // 判断债券分类
        const category = this.classifyBond(bondName);

        this.bonds.push({
          rowNumber: rowIndex + 1,
          bondCode,
          bondName,
          availableAmount,
          rowData: row,
          category
        });
      }
    }

    this.result.totalRows = this.bonds.length;
    console.log(`读取到 ${this.bonds.length} 条债券数据`);
  }

  /**
   * 债券分类
   * 优先级：地方债 > 政金债 > 国债/信用债
   */
  private classifyBond(bondName: string): 'treasury' | 'policy' | 'local' | 'credit' {
    // 1. 判断是否为地方债：包含地理名词
    if (containsGeographyKeyword(bondName)) {
      return 'local';
    }

    // 2. 判断是否为政金债：包含"国开"、"进出口"、"农发"
    if (bondName.includes('国开') || bondName.includes('进出口') || bondName.includes('农发')) {
      return 'policy';
    }

    // 3. 剩余债券：判断是否含"国债"字样
    if (bondName.includes('国债')) {
      return 'treasury';
    }

    // 4. 其余为信用债
    return 'credit';
  }

  /**
   * 创建四个分类工作表
   */
  private createClassificationSheets(): void {
    console.log('创建分类工作表...');

    // 获取表头行（第8行）
    const headerRow = this.sourceSheetData[this.HEADER_ROW - 1] || [];

    // 分类并排序
    const treasuryBonds = this.bonds
      .filter(b => b.category === 'treasury')
      .sort((a, b) => b.availableAmount - a.availableAmount);

    const policyBonds = this.bonds
      .filter(b => b.category === 'policy')
      .sort((a, b) => b.availableAmount - a.availableAmount);

    const localBonds = this.bonds
      .filter(b => b.category === 'local')
      .sort((a, b) => b.availableAmount - a.availableAmount);

    const creditBonds = this.bonds
      .filter(b => b.category === 'credit')
      .sort((a, b) => b.availableAmount - a.availableAmount);

    // 如果是 .xls 格式，需要创建新工作簿
    if (this.isXlsFormat) {
      this.workbook = new ExcelJS.Workbook();
    }

    // 删除除第一个工作表外的其他工作表（如果有）
    const sheetsToRemove = this.workbook.worksheets.slice(1);
    sheetsToRemove.forEach(sheet => this.workbook.removeWorksheet(sheet.id));

    // 创建四个分类工作表
    this.createSheet('国债', headerRow, treasuryBonds);
    this.createSheet('政金债', headerRow, policyBonds);
    this.createSheet('地方债', headerRow, localBonds);
    this.createSheet('信用债', headerRow, creditBonds);

    console.log(`创建完成：国债 ${treasuryBonds.length} 条，政金债 ${policyBonds.length} 条，地方债 ${localBonds.length} 条，信用债 ${creditBonds.length} 条`);
  }

  /**
   * 创建单个分类工作表
   */
  private createSheet(sheetName: string, headerRow: any[], bonds: BondData[]): void {
    const sheet = this.workbook.addWorksheet(sheetName);

    // 复制前8行（标题区域）
    for (let i = 0; i < this.HEADER_ROW; i++) {
      const row = this.sourceSheetData[i];
      if (row) {
        const sheetRow = sheet.getRow(i + 1);
        row.forEach((value, colIndex) => {
          sheetRow.getCell(colIndex + 1).value = value;
        });
        sheetRow.commit();
      }
    }

    // 写入表头（第8行）
    const headerSheetRow = sheet.getRow(this.HEADER_ROW);
    headerRow.forEach((value, colIndex) => {
      headerSheetRow.getCell(colIndex + 1).value = value;
    });
    headerSheetRow.commit();

    // 写入数据（从第9行开始）
    bonds.forEach((bond, index) => {
      const sheetRow = sheet.getRow(this.DATA_START_ROW + index);
      bond.rowData.forEach((value, colIndex) => {
        sheetRow.getCell(colIndex + 1).value = value;
      });
      sheetRow.commit();
    });

    // 设置列宽
    sheet.columns.forEach((column, colIndex) => {
      column.width = 15;
    });
  }
}
