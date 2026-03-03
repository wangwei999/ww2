'use client';

import { useState, useEffect } from 'react';
import { Button } from '@/components/ui/button';
import { Card } from '@/components/ui/card';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Upload, FileSpreadsheet, Download, Loader2, CheckCircle, AlertCircle, X } from 'lucide-react';
import { toast } from 'sonner';

interface FileUploadProps {
  label: string;
  description: string;
  file: File | null;
  onFileChange: (file: File | null) => void;
  acceptedTypes: string;
}

function FileUpload({ label, description, file, onFileChange, acceptedTypes }: FileUploadProps) {
  const handleDelete = () => {
    onFileChange(null);
  };

  return (
    <Card className="p-6">
      <div className="space-y-4">
        <div>
          <Label className="text-base font-semibold">{label}</Label>
          <p className="text-sm text-muted-foreground mt-1">{description}</p>
        </div>
        
        <div className="flex items-center gap-4">
          <div className="flex-1">
            <Input
              type="file"
              accept={acceptedTypes}
              onChange={(e) => onFileChange(e.target.files?.[0] || null)}
              className="cursor-pointer"
              disabled={!!file}
            />
          </div>
          
          {file && (
            <div className="flex items-center gap-3 text-sm">
              <div className="flex items-center gap-2 text-green-600 bg-green-50 dark:bg-green-950/20 px-3 py-1.5 rounded-md">
                <CheckCircle className="h-4 w-4" />
                <span className="max-w-[200px] truncate">{file.name}</span>
              </div>
              <Button
                type="button"
                variant="ghost"
                size="sm"
                onClick={handleDelete}
                className="h-8 w-8 p-0 hover:bg-red-50 hover:text-red-600 dark:hover:bg-red-950/20"
                title="删除文件"
              >
                <X className="h-4 w-4" />
              </Button>
            </div>
          )}
        </div>
      </div>
    </Card>
  );
}

export default function Home() {
  const [fileA, setFileA] = useState<File | null>(null);
  const [fileB, setFileB] = useState<File | null>(null);
  const [processing, setProcessing] = useState(false);
  const [downloading, setDownloading] = useState(false);
  const [hasProcessedFile, setHasProcessedFile] = useState(false);

  // 客户端加载时检查是否有已处理的文件
  useEffect(() => {
    if (typeof window !== 'undefined') {
      const fileId = sessionStorage.getItem('processedFileId');
      setHasProcessedFile(!!fileId);
    }
  }, []);

  // 清空处理状态
  const clearProcessedState = () => {
    if (typeof window !== 'undefined') {
      sessionStorage.removeItem('processedFileId');
      setHasProcessedFile(false);
    }
  };

  const handleProcess = async () => {
    console.log('开始处理 - fileA:', fileA?.name, 'fileB:', fileB?.name);
    
    if (!fileA || !fileB) {
      toast.error('请上传两个文件');
      return;
    }

    setProcessing(true);
    
    try {
      const formData = new FormData();
      formData.append('fileA', fileA);
      formData.append('fileB', fileB);

      const response = await fetch('/api/process', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.error || '处理失败');
      }

      const data = await response.json();
      console.log('处理结果:', data);
      
      if (!data.success) {
        throw new Error(data.message || '处理失败');
      }
      
      if (!data.fileId) {
        throw new Error('未获取到文件ID');
      }
      
      // 保存文件ID以便后续下载
      if (typeof window !== 'undefined') {
        sessionStorage.setItem('processedFileId', data.fileId);
        setHasProcessedFile(true);
      }
      
      toast.success(`数据处理完成！共填充 ${data.statistics?.totalFilled || 0} 个单元格`);
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : '处理失败，请重试';
      toast.error(errorMessage);
      console.error('处理错误:', error);
    } finally {
      setProcessing(false);
    }
  };

  const handleDownload = async () => {
    if (typeof window === 'undefined') return;
    
    const fileId = sessionStorage.getItem('processedFileId');
    if (!fileId) {
      toast.error('没有可下载的文件');
      return;
    }

    setDownloading(true);
    
    try {
      const response = await fetch(`/api/download?fileId=${fileId}`);
      if (!response.ok) throw new Error('下载失败');

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `填充结果_${fileB?.name || '结果.xlsx'}`;
      link.click();
      window.URL.revokeObjectURL(url);
      
      toast.success('下载成功！');
    } catch (error) {
      toast.error('下载失败，请重试');
      console.error(error);
    } finally {
      setDownloading(false);
    }
  };

  return (
    <div className="min-h-screen bg-background py-12 px-4 sm:px-6 lg:px-8">
      <div className="max-w-4xl mx-auto space-y-8">
        {/* 标题 */}
        <div className="text-center space-y-2">
          <div className="flex justify-center">
            <FileSpreadsheet className="h-12 w-12 text-primary" />
          </div>
          <h1 className="text-3xl font-bold tracking-tight">智能表格数据填充工具</h1>
          <p className="text-muted-foreground">
            上传数据源文件和缺失文件，自动识别表格、匹配字段并填充数据
          </p>
        </div>

        {/* 功能说明 */}
        <Card className="p-6 bg-muted/50">
          <h3 className="font-semibold mb-3">支持的功能</h3>
          <ul className="grid grid-cols-1 md:grid-cols-2 gap-2 text-sm text-muted-foreground">
            <li>✓ 支持 Word、Excel、WPS 格式</li>
            <li>✓ 自动识别文档中的所有表格</li>
            <li>✓ 同义词智能匹配（如"总资产"匹配"资产总额"）</li>
            <li>✓ 多格式日期识别（2025/9、2025-9、2025.9等）</li>
            <li>✓ 自动识别单位并换算（亿元、万元、百分比）</li>
            <li>✓ 表格位置任意，自动定位</li>
          </ul>
        </Card>

        {/* 文件上传区域 */}
        <div className="space-y-6">
          <FileUpload
            label="文件A（数据源文件）"
            description="上传包含完整数据的数据源文件"
            file={fileA}
            onFileChange={(file) => {
              setFileA(file);
              if (!file) clearProcessedState();
            }}
            acceptedTypes=".xlsx,.xls,.docx,.doc"
          />

          <FileUpload
            label="文件B（数据缺失文件）"
            description="上传需要填充数据的缺失文件，横轴为字段，纵轴为时间点"
            file={fileB}
            onFileChange={(file) => {
              setFileB(file);
              if (!file) clearProcessedState();
            }}
            acceptedTypes=".xlsx,.xls,.docx,.doc"
          />
        </div>

        {/* 操作按钮 */}
        <div className="flex gap-4 justify-center">
          <Button
            onClick={handleProcess}
            disabled={processing || !fileA || !fileB}
            size="lg"
            className="min-w-[180px]"
          >
            {processing ? (
              <>
                <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                处理中...
              </>
            ) : (
              <>
                <Upload className="mr-2 h-4 w-4" />
                开始处理
              </>
            )}
          </Button>

          <div className="flex flex-col items-center gap-2">
            <Button
              onClick={handleDownload}
              disabled={downloading || !hasProcessedFile}
              variant="outline"
              size="lg"
              className="min-w-[180px]"
            >
              {downloading ? (
                <>
                  <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                  下载中...
                </>
              ) : hasProcessedFile ? (
                <>
                  <Download className="mr-2 h-4 w-4" />
                  下载结果
                </>
              ) : (
                <>
                  <Download className="mr-2 h-4 w-4" />
                  等待处理
                </>
              )}
            </Button>
            {!hasProcessedFile && (
              <p className="text-xs text-muted-foreground">
                请先上传文件并点击"开始处理"
              </p>
            )}
          </div>
        </div>

        {/* 注意事项 */}
        <Card className="p-6 border-amber-200 bg-amber-50 dark:bg-amber-950/20">
          <div className="flex gap-3">
            <AlertCircle className="h-5 w-5 text-amber-600 flex-shrink-0 mt-0.5" />
            <div className="space-y-1 text-sm">
              <p className="font-medium text-amber-900 dark:text-amber-200">使用说明</p>
              <ul className="text-amber-800 dark:text-amber-300/80 space-y-1 list-disc list-inside">
                <li>确保文件B中的空白单元格可以从文件A中找到匹配数据</li>
                <li>单位会自动识别并进行换算，无需手动调整</li>
                <li>百分比符号会根据表格外标注自动处理</li>
                <li>同义词匹配功能会自动识别相似字段名称</li>
              </ul>
            </div>
          </div>
        </Card>
      </div>
    </div>
  );
}
