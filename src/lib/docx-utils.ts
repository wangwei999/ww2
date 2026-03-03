import { Document, Packer, Paragraph, Table, TableRow, TableCell as DocxTableCell, WidthType, BorderStyle, TextRun, AlignmentType, VerticalAlign } from 'docx';
import { TableData, TableCell } from './types';

/**
 * 将表格数据转换为 DOCX 格式
 * @param tables 表格数据数组
 * @param filename 文件名（用于生成文档标题）
 * @returns Buffer - DOCX 文件的二进制数据
 */
export async function tablesToDocx(tables: TableData[], filename: string): Promise<Buffer> {
  console.log('=== 转换表格到 DOCX 格式 ===');
  console.log('表格数量:', tables.length);
  console.log('文件名:', filename);
  
  // 创建文档内容
  const children: any[] = [];
  
  // 添加文档标题
  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: filename.replace(/\.(xlsx|xls|docx|doc|csv)$/i, ''),
          bold: true,
          size: 32, // 16pt
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: {
        after: 400,
      },
    })
  );
  
  // 遍历每个表格
  tables.forEach((table, tableIndex) => {
    console.log(`处理表格 ${tableIndex + 1}...`);
    console.log(`  表头: ${JSON.stringify(table.headers)}`);
    console.log(`  数据行数: ${table.rows?.length || 0}`);
    
    // 添加表格标题（如果有多个表格）
    if (tables.length > 1) {
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: `表格 ${tableIndex + 1}`,
              bold: true,
              size: 24, // 12pt
            }),
          ],
          spacing: {
            before: 300,
            after: 200,
          },
        })
      );
    }
    
    // 创建表格行
    const tableRows: TableRow[] = [];
    
    // 创建表头行
    if (table.headers && table.headers.length > 0) {
      const headerCells = table.headers.map((header: string | number | null) => {
        return new DocxTableCell({
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: String(header || ''),
                  bold: true,
                  size: 20, // 10pt
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: {
                before: 100,
                after: 100,
              },
            }),
          ],
          width: {
            size: 100 / table.headers.length,
            type: WidthType.PERCENTAGE,
          },
          verticalAlign: VerticalAlign.CENTER,
          shading: {
            fill: 'E8E8E8', // 浅灰色背景
          },
          margins: {
            top: 100,
            bottom: 100,
            left: 100,
            right: 100,
          },
        });
      });
      
      tableRows.push(
        new TableRow({
          children: headerCells,
          tableHeader: true,
        })
      );
    }
    
    // 创建数据行
    if (table.rows && table.rows.length > 0) {
      table.rows.forEach((row: TableCell[]) => {
        console.log(`  处理行: ${JSON.stringify(row)}`);
        
        const rowCells = row.map((cell: TableCell) => {
          const cellValue = cell === null || cell === undefined ? '' : String(cell);
          
          return new DocxTableCell({
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: cellValue,
                    size: 20, // 10pt
                  }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: {
                  before: 80,
                  after: 80,
                },
              }),
            ],
            width: {
              size: 100 / row.length,
              type: WidthType.PERCENTAGE,
            },
            verticalAlign: VerticalAlign.CENTER,
            margins: {
              top: 80,
              bottom: 80,
              left: 100,
              right: 100,
            },
          });
        });
        
        tableRows.push(new TableRow({ children: rowCells }));
      });
    }
    
    // 添加表格到文档
    children.push(
      new Table({
        rows: tableRows,
        width: {
          size: 100,
          type: WidthType.PERCENTAGE,
        },
        borders: {
          top: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
          bottom: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
          left: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
          right: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
          insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
          insideVertical: { style: BorderStyle.SINGLE, size: 1, color: '000000' },
        },
      })
    );
    
    // 表格之间添加空行
    if (tableIndex < tables.length - 1) {
      children.push(
        new Paragraph({
          children: [new TextRun('')],
          spacing: { after: 200 },
        })
      );
    }
  });
  
  // 创建文档
  const doc = new Document({
    sections: [
      {
        properties: {},
        children: children,
      },
    ],
  });
  
  // 生成Buffer
  const buffer = await Packer.toBuffer(doc);
  console.log('DOCX 文件生成成功，大小:', buffer.length);
  
  return buffer;
}
