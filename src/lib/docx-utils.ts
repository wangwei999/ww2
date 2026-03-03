import { Stream } from 'stream';
import * as fs from 'fs';
import { TableData, TableCell } from './types';

/**
 * 将表格数据转换为 DOCX 格式
 * @param tables 表格数据数组
 * @param filename 文件名（用于生成文档标题）
 * @returns Promise<Buffer> - DOCX 文件的二进制数据
 */
export async function tablesToDocx(tables: TableData[], filename: string): Promise<Buffer> {
  console.log('=== 转换表格到 DOCX 格式 ===');
  console.log('表格数量:', tables.length);
  console.log('文件名:', filename);

  // 动态导入 officegen（避免 CommonJS/ESM 问题）
  const officegen = (await import('officegen')).default;

  return new Promise((resolve, reject) => {
    try {
      // 创建 DOCX 文档
      const docx: any = officegen('docx');

      docx.on('finalize', function(written: any) {
        console.log('DOCX 文件生成成功，大小:', written, 'bytes');
      });

      docx.on('error', function(err: any) {
        console.error('DOCX 生成错误:', err);
        reject(err);
      });

      // 添加文档标题
      const titleObj = docx.createP({ align: 'center' });
      titleObj.addText(filename.replace(/\.(xlsx|xls|docx|doc|csv)$/i, ''), {
        bold: true,
        size: 32, // 16pt
      });

      // 添加空行
      docx.createP();

      // 遍历每个表格
      tables.forEach((table, tableIndex) => {
        console.log(`处理表格 ${tableIndex + 1}...`);
        console.log(`  表头: ${JSON.stringify(table.headers)}`);
        console.log(`  数据行数: ${table.rows?.length || 0}`);

        // 添加表格标题（如果有多个表格）
        if (tables.length > 1) {
          const tableTitle = docx.createP();
          tableTitle.addText(`表格 ${tableIndex + 1}`, {
            bold: true,
            size: 24, // 12pt
          });
          docx.createP();
        }

        // 准备表格数据
        const tableData: any[] = [];

        // 添加表头行
        if (table.headers && table.headers.length > 0) {
          const headerRow = table.headers.map((header) => ({
            val: String(header || ''),
            opts: { bold: true, color: '000000' },
          }));
          tableData.push(headerRow);
        }

        // 添加数据行
        if (table.rows && table.rows.length > 0) {
          table.rows.forEach((row: TableCell[]) => {
            console.log(`  处理行: ${JSON.stringify(row)}`);
            const rowData = row.map((cell: TableCell) => ({
              val: cell === null || cell === undefined ? '' : String(cell),
            }));
            tableData.push(rowData);
          });
        }

        // 创建表格
        if (tableData.length > 0) {
          docx.createTable(tableData, {
            table: {
              border: 1,
              align: 'center',
              width: {
                size: 100,
                type: 'pct',
              },
            },
          });
        }

        // 表格之间添加空行
        if (tableIndex < tables.length - 1) {
          docx.createP();
        }
      });

      // 使用可写流收集数据
      const chunks: Buffer[] = [];
      const writable = new Stream.Writable({
        write(chunk: any, encoding: string, callback: Function) {
          chunks.push(chunk);
          callback();
        },
      });

      writable.on('finish', () => {
        const buffer = Buffer.concat(chunks);
        console.log('Buffer 大小:', buffer.length);
        resolve(buffer);
      });

      writable.on('error', (err) => {
        console.error('写入流错误:', err);
        reject(err);
      });

      // 生成文档到流
      docx.generate(writable);
    } catch (error) {
      console.error('生成 DOCX 失败:', error);
      reject(error);
    }
  });
}
