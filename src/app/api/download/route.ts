import { NextRequest, NextResponse } from 'next/server';
import { readFile, unlink, stat } from 'fs/promises';
import { existsSync } from 'fs';
import path from 'path';

const TEMP_DIR = path.join(process.cwd(), 'temp');

export async function GET(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url);
    const fileId = searchParams.get('fileId');
    
    if (!fileId) {
      return NextResponse.json({ error: '缺少文件ID' }, { status: 400 });
    }
    
    const filePath = path.join(TEMP_DIR, fileId);
    
    console.log('下载请求 - 文件ID:', fileId);
    console.log('下载请求 - 文件路径:', filePath);
    
    if (!existsSync(filePath)) {
      console.log('下载请求 - 文件不存在');
      return NextResponse.json({ error: '文件不存在' }, { status: 404 });
    }
    
    // 获取文件大小
    const stats = await stat(filePath);
    console.log('下载请求 - 文件大小:', stats.size, 'bytes');
    
    const fileBuffer = await readFile(filePath);
    console.log('下载请求 - 读取文件大小:', fileBuffer.length, 'bytes');
    
    // 获取文件名
    const fileName = fileId.replace(/^\d+_/, '');
    console.log('下载请求 - 文件名:', fileName);
    
    // 设置响应头
    const headers = new Headers();
    headers.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    // 使用 RFC 5987 格式的 Content-Disposition
    headers.set('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`);
    headers.set('Content-Length', String(fileBuffer.length));
    
    console.log('下载请求 - 开始发送响应...');
    
    // 在响应完成后再删除文件
    setTimeout(() => {
      unlink(filePath).catch(err => console.error('删除临时文件失败:', err));
      console.log('下载请求 - 已删除临时文件:', fileId);
    }, 1000); // 延迟1秒删除，确保响应完成
    
    return new NextResponse(fileBuffer, { headers });
  } catch (error) {
    console.error('下载错误:', error);
    return NextResponse.json({ 
      error: error instanceof Error ? error.message : '下载失败' 
    }, { status: 500 });
  }
}
