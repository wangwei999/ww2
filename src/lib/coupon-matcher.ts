import ExcelJS from 'exceljs';
import { containsGeographyKeyword } from './geo-keywords';

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

  // 处理结果
  private result: {
    bondType: 'treasury' | 'local'; // 实际判断的债券类型
    totalRows: number;
    filteredRows: number;
    data: any[];
  } = {
    bondType: 'treasury',
    totalRows: 0,
    filteredRows: 0,
    data: []
  };

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

    // 2. 判断债券类型（如果用户选择地方债）
    if (this.bondType === 'local') {
      this.determineBondType();
    } else {
      this.result.bondType = 'treasury';
    }

    console.log('实际债券类型:', this.result.bondType);

    // 3. TODO: 根据金额筛选债券（等待后续规则）

    // 4. 生成结果Excel
    await this.generateResultWorkbook();

    console.log('=== 挑券处理完成 ===');

    return {
      workbook: this.workbook,
      statistics: {
        bondType: this.result.bondType,
        totalRows: this.result.totalRows,
        filteredRows: this.result.filteredRows,
        amount: this.amount
      }
    };
  }

  /**
   * 加载Excel文件
   */
  private async loadExcelFile(): Promise<void> {
    console.log('加载Excel文件...');

    // 获取Buffer
    let buffer: Buffer;
    if (this.file instanceof File) {
      const arrayBuffer = await this.file.arrayBuffer();
      buffer = Buffer.from(arrayBuffer);
    } else {
      buffer = this.file;
    }

    await this.workbook.xlsx.load(buffer as any);

    // 获取第一个工作表
    this.worksheet = this.workbook.worksheets[0];
    if (!this.worksheet) {
      throw new Error('Excel文件中没有工作表');
    }

    console.log('工作表名称:', this.worksheet.name);
  }

  /**
   * 判断债券类型
   * 规则：读取C列，如果包含地理名词则为地方债，否则为国债
   * 匹配方式：包含匹配，如"20山西03"包含"山西"、"19柯城国资债"包含"柯城"
   */
  private determineBondType(): void {
    console.log('判断债券类型（读取C列）...');

    if (!this.worksheet) {
      this.result.bondType = 'treasury';
      return;
    }

    let hasGeographyKeyword = false;
    let checkedRows = 0;

    // 从第2行开始检查（跳过表头）
    this.worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // 跳过表头

      // 获取C列（第3列）的值
      const cellC = row.getCell(3);
      const cellValue = String(cellC.value || '').trim();

      if (cellValue) {
        checkedRows++;
        if (containsGeographyKeyword(cellValue)) {
          hasGeographyKeyword = true;
          console.log(`  行${rowNumber} C列: "${cellValue}" 包含地理名词`);
        }
      }
    });

    this.result.bondType = hasGeographyKeyword ? 'local' : 'treasury';
    console.log(`检查了 ${checkedRows} 行，包含地理名词: ${hasGeographyKeyword}`);
  }

  /**
   * 生成结果Excel
   * TODO: 根据后续规则实现具体的筛选逻辑
   */
  private async generateResultWorkbook(): Promise<void> {
    console.log('生成结果Excel...');

    if (!this.worksheet) return;

    // 统计总行数（不含表头）
    let totalRows = 0;
    this.worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) totalRows++;
    });

    this.result.totalRows = totalRows;
    this.result.filteredRows = totalRows; // 暂时保留所有行

    // TODO: 后续根据金额筛选
  }
}
