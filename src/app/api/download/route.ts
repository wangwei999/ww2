import { NextRequest, NextResponse } from 'next/server';
import { readFile, unlink } from 'fs/promises';
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
    
    if (!existsSync(filePath)) {
      return NextResponse.json({ error: '文件不存在' }, { status: 404 });
    }
    
    const fileBuffer = await readFile(filePath);
    
    // 获取文件名
    const fileName = fileId.replace(/^\d+_/, '');
    
    // 设置响应头
    const headers = new Headers();
    headers.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    headers.set('Content-Disposition', `attachment; filename="${encodeURIComponent(fileName)}"`);
    
    // 异步删除文件（不影响下载）
    unlink(filePath).catch(err => console.error('删除临时文件失败:', err));
    
    return new NextResponse(fileBuffer, { headers });
  } catch (error) {
    console.error('下载错误:', error);
    return NextResponse.json({ 
      error: error instanceof Error ? error.message : '下载失败' 
    }, { status: 500 });
  }
}
