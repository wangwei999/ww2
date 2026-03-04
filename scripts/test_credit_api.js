const fs = require('fs');
const FormData = require('form-data');

async function testCreditModeAPI() {
  console.log('=== 测试授信模式API ===\n');

  // 准备文件
  const fileAPath = '/tmp/授信2026.xlsx';
  const fileBPath = '/tmp/A类授信调整.xlsx';

  // 检查文件是否存在
  if (!fs.existsSync(fileAPath)) {
    console.error(`文件A不存在: ${fileAPath}`);
    process.exit(1);
  }
  if (!fs.existsSync(fileBPath)) {
    console.error(`文件B不存在: ${fileBPath}`);
    process.exit(1);
  }

  console.log('文件A存在:', fileAPath);
  console.log('文件B存在:', fileBPath);

  // 创建FormData
  const form = new FormData();
  form.append('fileA', fs.createReadStream(fileAPath));
  form.append('fileB', fs.createReadStream(fileBPath));

  console.log('\n发送请求到 /api/process...');

  const response = await fetch('http://localhost:5000/api/process', {
    method: 'POST',
    body: form,
    headers: form.getHeaders(),
  });

  console.log('响应状态:', response.status, response.statusText);

  const data = await response.json();
  console.log('\n响应数据:');
  console.log(JSON.stringify(data, null, 2));

  if (data.success && data.downloadUrl) {
    console.log(`\n下载链接: ${data.downloadUrl}`);
    console.log('统计信息:', data.statistics);
  } else if (data.error) {
    console.error(`\n错误: ${data.error}`);
  }
}

testCreditModeAPI().catch(console.error);
