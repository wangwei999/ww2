import ExcelJS from 'exceljs';

/**
 * 中国地理名词列表（用于判断地方债）
 * 包含省、市、区县等地理名称
 */
const GEOGRAPHY_KEYWORDS = [
  // 直辖市
  '北京', '上海', '天津', '重庆',
  // 省会及地级市
  '杭州', '宁波', '温州', '绍兴', '金华', '台州', '嘉兴', '湖州', '衢州', '丽水', '舟山',
  '南京', '苏州', '无锡', '常州', '南通', '扬州', '徐州', '盐城', '镇江', '泰州', '淮安', '连云港', '宿迁',
  '广州', '深圳', '珠海', '汕头', '佛山', '东莞', '中山', '惠州', '江门', '湛江', '肇庆', '清远', '揭阳', '梅州', '潮州', '河源', '汕尾', '阳江', '云浮', '韶关',
  '武汉', '宜昌', '襄阳', '荆州', '黄冈', '孝感', '十堰', '咸宁', '黄石', '鄂州', '荆门', '随州', '恩施',
  '成都', '绵阳', '德阳', '宜宾', '南充', '泸州', '达州', '乐山', '自贡', '内江', '遂宁', '广元', '眉山', '广安', '资阳', '攀枝花', '雅安', '巴中',
  '西安', '咸阳', '宝鸡', '渭南', '汉中', '榆林', '延安', '安康', '商洛', '铜川',
  '郑州', '洛阳', '开封', '南阳', '安阳', '新乡', '平顶山', '焦作', '商丘', '许昌', '信阳', '周口', '驻马店', '濮阳', '漯河', '三门峡', '鹤壁', '济源',
  '长沙', '株洲', '湘潭', '衡阳', '岳阳', '常德', '邵阳', '益阳', '娄底', '郴州', '永州', '怀化', '张家界', '湘西',
  '济南', '青岛', '烟台', '潍坊', '临沂', '淄博', '济宁', '泰安', '威海', '德州', '聊城', '菏泽', '滨州', '东营', '枣庄', '日照', '莱芜',
  '福州', '厦门', '泉州', '漳州', '龙岩', '莆田', '三明', '南平', '宁德',
  '合肥', '芜湖', '蚌埠', '阜阳', '淮南', '安庆', '宿州', '六安', '马鞍山', '铜陵', '滁州', '宣城', '亳州', '黄山', '池州', '淮北',
  '南昌', '赣州', '九江', '吉安', '上饶', '宜春', '抚州', '新余', '景德镇', '萍乡', '鹰潭',
  '石家庄', '唐山', '保定', '邯郸', '廊坊', '沧州', '秦皇岛', '邢台', '张家口', '承德', '衡水',
  '哈尔滨', '大庆', '齐齐哈尔', '牡丹江', '佳木斯', '鸡西', '鹤岗', '双鸭山', '伊春', '七台河', '黑河', '绥化', '大兴安岭',
  '沈阳', '大连', '鞍山', '抚顺', '本溪', '丹东', '锦州', '营口', '阜新', '辽阳', '盘锦', '铁岭', '朝阳', '葫芦岛',
  '长春', '吉林', '四平', '辽源', '通化', '白山', '松原', '白城', '延边',
  '昆明', '大理', '曲靖', '红河', '玉溪', '丽江', '文山', '楚雄', '西双版纳', '昭通', '保山', '普洱', '临沧', '德宏', '怒江', '迪庆',
  '贵阳', '遵义', '黔东南', '毕节', '铜仁', '黔南', '六盘水', '黔西南', '安顺',
  '兰州', '天水', '庆阳', '平凉', '酒泉', '张掖', '武威', '定西', '金昌', '陇南', '临夏', '嘉峪关', '甘南',
  '南宁', '柳州', '桂林', '梧州', '北海', '玉林', '钦州', '百色', '河池', '贵港', '贺州', '来宾', '崇左', '防城港',
  '海口', '三亚', '三沙', '儋州', '琼海', '文昌', '万宁', '东方',
  '太原', '大同', '临汾', '运城', '长治', '晋中', '晋城', '忻州', '朔州', '吕梁', '阳泉',
  '西宁', '海东', '海北', '黄南', '海南', '果洛', '玉树', '海西',
  '银川', '石嘴山', '吴忠', '固原', '中卫',
  '乌鲁木齐', '克拉玛依', '吐鲁番', '哈密', '昌吉', '博尔塔拉', '巴音郭楞', '阿克苏', '克孜勒苏', '喀什', '和田', '伊犁', '塔城', '阿勒泰',
  '呼和浩特', '包头', '鄂尔多斯', '赤峰', '通辽', '呼伦贝尔', '巴彦淖尔', '乌兰察布', '兴安盟', '锡林郭勒', '阿拉善', '乌海',
  '拉萨', '日喀则', '昌都', '林芝', '山南', '那曲', '阿里',
  // 常见区县名
  '黄岩', '椒江', '路桥', '温岭', '临海', '玉环', '天台', '仙居', '三门',
  '萧山', '余杭', '富阳', '临安', '建德', '桐庐', '淳安',
  '浦东', '浦西', '闵行', '宝山', '嘉定', '松江', '青浦', '奉贤', '金山', '崇明',
  '江宁', '浦口', '六合', '溧水', '高淳', '江阴', '宜兴', '新吴', '锡山', '惠山', '滨湖',
  // 其他常见地名
  '义乌', '东阳', '永康', '武义', '浦江', '磐安', '兰溪',
  '诸暨', '上虞', '嵊州', '新昌', '越城', '柯桥',
  '瑞安', '乐清', '永嘉', '平阳', '苍南', '文成', '泰顺', '洞头', '龙湾', '瓯海', '鹿城',
  '余姚', '慈溪', '象山', '宁海', '奉化', '北仑', '镇海', '鄞州', '海曙', '江北',
];

/**
 * 判断字符串是否包含地理名词
 * @param text 要检查的文本
 * @returns 是否包含地理名词
 */
export function containsGeographyKeyword(text: string): boolean {
  if (!text || typeof text !== 'string') return false;
  
  const upperText = text.toUpperCase();
  return GEOGRAPHY_KEYWORDS.some(keyword => 
    text.includes(keyword) || upperText.includes(keyword.toUpperCase())
  );
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
