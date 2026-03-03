import * as fs from 'fs';
import * as path from 'path';
import FormData from 'form-data';
import fetch from 'node-fetch';

const TEST_DIR = path.join(process.cwd(), 'temp');

async function testAPI() {
  console.log('=== 测试 API 接口 ===\n');

  const missingFilePath = path.join(TEST_DIR, 'testB.xlsx');
  const sourceFilePath = path.join(TEST_DIR, 'data-source-complete.xlsx');

  if (!fs.existsSync(missingFilePath) || !fs.existsSync(sourceFilePath)) {
    console.error('文件不存在');
    return;
  }

  const formData = new FormData();
  formData.append('fileA', fs.createReadStream(sourceFilePath), 'data-source-complete.xlsx');
  formData.append('fileB', fs.createReadStream(missingFilePath), 'testB.xlsx');

  try {
    const response = await fetch('http://localhost:5000/api/process', {
      method: 'POST',
      body: formData,
    });

    const data = await response.json();
    console.log('API 响应:', JSON.stringify(data, null, 2));

    if (data.success && data.downloadUrl) {
      // 下载生成的文件
      const fileResponse = await fetch(`http://localhost:5000${data.downloadUrl}`);
      const buffer = await fileResponse.buffer();

      const outputPath = path.join(TEST_DIR, `test-final-${Date.now()}.xlsx`);
      fs.writeFileSync(outputPath, buffer);
      console.log('\n文件已保存:', outputPath);

      // 检查数据
      const XLSX = require('xlsx');
      const workbook = XLSX.readFile(outputPath);
      const fileData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
      
      console.log('\n检查填充情况:');
      const expectedFields = ['总资产', '其中：贷款总额', '营业收入（亿元）', '净利润（亿元）'];
      for (let i = 1; i < fileData.length; i++) {
        const fieldName = fileData[i][1];
        if (expectedFields.includes(fieldName)) {
          const lastCols = fileData[i].slice(8);
          const hasData = lastCols.some(v => v !== null && v !== undefined && v !== '' && typeof v === 'number');
          console.log(`${fieldName}: ${hasData ? '✓ 有数据' : '✗ 无数据'}`, lastCols);
        }
      }
    }
  } catch (error) {
    console.error('测试失败:', error);
  }
}

testAPI();
