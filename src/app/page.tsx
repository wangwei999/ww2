'use client';

import { useState, useEffect, useRef } from 'react';
import { Button } from '@/components/ui/button';
import { Card } from '@/components/ui/card';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group';
import { Upload, FileSpreadsheet, Download, Loader2, CheckCircle, AlertCircle, X, FileText } from 'lucide-react';
import { toast } from 'sonner';

type ProcessMode = 'normal' | 'credit' | 'pdf';

interface FileUploadProps {
  label: string;
  description: string;
  file: File | null;
  onFileChange: (file: File | null) => void;
  acceptedTypes: string;
}

function FileUpload({ label, description, file, onFileChange, acceptedTypes }: FileUploadProps) {
  const inputRef = useRef<HTMLInputElement>(null);
  
  const handleClick = () => {
    inputRef.current?.click();
  };
  
  const handleDelete = () => {
    onFileChange(null);
    if (inputRef.current) {
      inputRef.current.value = '';
    }
  };

  return (
    <Card className="p-6">
      <div className="space-y-4">
        <div>
          <Label className="text-base font-semibold">{label}</Label>
          <p className="text-sm text-muted-foreground mt-1">{description}</p>
        </div>
        
        <div className="flex items-center gap-4">
          {file ? (
            <div className="flex-1 flex items-center gap-3">
              <div className="flex-1 flex items-center gap-2 text-sm text-green-600 bg-green-50 dark:bg-green-950/20 px-3 py-1.5 rounded-md">
                <CheckCircle className="h-4 w-4" />
                <span className="max-w-[300px] truncate">{file.name}</span>
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
          ) : (
            <div className="flex-1">
              <Button
                type="button"
                variant="outline"
                onClick={handleClick}
                className="w-full h-9 justify-start text-left px-3"
              >
                <Upload className="mr-2 h-4 w-4" />
                <span className="text-muted-foreground">选择文件</span>
              </Button>
              <input
                ref={inputRef}
                type="file"
                accept={acceptedTypes}
                onChange={(e) => onFileChange(e.target.files?.[0] || null)}
                className="hidden"
              />
            </div>
          )}
        </div>
      </div>
    </Card>
  );
}

export default function Home() {
  const [mode, setMode] = useState<ProcessMode>('normal');
  const [fileA, setFileA] = useState<File | null>(null);
  const [fileB, setFileB] = useState<File | null>(null);
  const [pdfFile, setPdfFile] = useState<File | null>(null);
  const [excelFile, setExcelFile] = useState<File | null>(null);
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

  const handleFileAChange = (file: File | null) => {
    setFileA(file);
    clearProcessedState();
  };

  const handleFileBChange = (file: File | null) => {
    setFileB(file);
    clearProcessedState();
  };

  const handlePdfFileChange = (file: File | null) => {
    setPdfFile(file);
    clearProcessedState();
  };

  const handleExcelFileChange = (file: File | null) => {
    setExcelFile(file);
    clearProcessedState();
  };

  const handleModeChange = (newMode: ProcessMode) => {
    setMode(newMode);
    clearProcessedState();
  };

  const handleProcess = async () => {
    if (mode === 'pdf') {
      // 授信写入处理
      console.log('授信写入处理 - pdfFile:', pdfFile?.name, 'excelFile:', excelFile?.name);
      
      if (!pdfFile || !excelFile) {
        toast.error('请上传PDF文件和Excel文件');
        return;
      }

      setProcessing(true);
      
      try {
        const formData = new FormData();
        formData.append('pdfFile', pdfFile);
        formData.append('excelFile', excelFile);

        const response = await fetch('/api/process-pdf', {
          method: 'POST',
          body: formData,
        });

        if (!response.ok) {
          const errorData = await response.json().catch(() => ({}));
          throw new Error(errorData.error || '处理失败');
        }

        // 从响应头获取文件名
        const contentDisposition = response.headers.get('Content-Disposition');
        let filename = `PDF处理结果_${Date.now()}.xlsx`;
        if (contentDisposition) {
          const filenameMatch = contentDisposition.match(/filename\*?=['"]?(?:UTF-\d['"]*)?([^;'"]+)/i);
          if (filenameMatch) {
            filename = decodeURIComponent(filenameMatch[1]);
          }
        }

        // 直接下载文件
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        link.click();
        window.URL.revokeObjectURL(url);
        
        toast.success('授信写入完成，文件已下载！');
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : '处理失败，请重试';
        toast.error(errorMessage);
        console.error('处理错误:', error);
      } finally {
        setProcessing(false);
      }
      return;
    }

    // 对外提供数据/A类授信调整处理
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
      formData.append('mode', mode);

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

  // 根据模式获取文件上传组件
  const renderFileUploads = () => {
    if (mode === 'pdf') {
      return (
        <div className="space-y-6">
          <FileUpload
            label="文件A（PDF扫描件）"
            description="上传扫描版PDF文件，包含机构名称和授信品种及金额表格"
            file={pdfFile}
            onFileChange={handlePdfFileChange}
            acceptedTypes=".pdf"
          />

          <FileUpload
            label="文件B（Excel文件）"
            description="上传包含单体表和集团表的Excel文件"
            file={excelFile}
            onFileChange={handleExcelFileChange}
            acceptedTypes=".xlsx,.xls"
          />
        </div>
      );
    }

    return (
      <div className="space-y-6">
        <FileUpload
          label="文件A（数据源文件）"
          description="上传包含完整数据的数据源文件"
          file={fileA}
          onFileChange={handleFileAChange}
          acceptedTypes=".xlsx,.xls,.docx,.doc"
        />

        <FileUpload
          label="文件B（数据缺失文件）"
          description="上传需要填充数据的缺失文件，横轴为字段，纵轴为时间点"
          file={fileB}
          onFileChange={handleFileBChange}
          acceptedTypes=".xlsx,.xls,.docx,.doc"
        />
      </div>
    );
  };

  // 获取处理按钮是否禁用
  const isProcessDisabled = () => {
    if (processing) return true;
    if (mode === 'pdf') {
      return !pdfFile || !excelFile;
    }
    return !fileA || !fileB;
  };

  // 获取下载按钮是否禁用（授信写入模式直接下载，不需要下载按钮）
  const isDownloadDisabled = () => {
    if (mode === 'pdf') return true;
    return downloading || !hasProcessedFile;
  };

  return (
    <div className="min-h-screen bg-background py-12 px-4 sm:px-6 lg:px-8">
      <div className="max-w-4xl mx-auto space-y-8">
        {/* 标题 */}
        <div className="text-center space-y-2">
          <div className="flex justify-center">
            <FileSpreadsheet className="h-12 w-12 text-primary" />
          </div>
          <h1 className="text-3xl font-bold tracking-tight">授信数据填充</h1>
          <p className="text-muted-foreground">
            上传数据源文件和缺失文件，自动识别表格、匹配字段并填充数据
          </p>
        </div>

        {/* 模式选择 */}
        <Card className="p-6">
          <div className="space-y-4">
            <Label className="text-base font-semibold">选择处理模式</Label>
            <RadioGroup
              value={mode}
              onValueChange={(value) => handleModeChange(value as ProcessMode)}
              className="grid grid-cols-1 md:grid-cols-3 gap-4"
            >
              <div className="flex items-center space-x-2 p-4 border rounded-lg hover:bg-muted/50 cursor-pointer">
                <RadioGroupItem value="normal" id="normal" />
                <div className="flex-1">
                  <Label htmlFor="normal" className="font-medium cursor-pointer">对外提供数据</Label>
                  <p className="text-xs text-muted-foreground mt-1">
                    自动识别表格结构，智能匹配字段
                  </p>
                </div>
              </div>
              <div className="flex items-center space-x-2 p-4 border rounded-lg hover:bg-muted/50 cursor-pointer">
                <RadioGroupItem value="credit" id="credit" />
                <div className="flex-1">
                  <Label htmlFor="credit" className="font-medium cursor-pointer">A类授信调整</Label>
                  <p className="text-xs text-muted-foreground mt-1">
                    基于机构名称匹配，支持单体/集团表
                  </p>
                </div>
              </div>
              <div className="flex items-center space-x-2 p-4 border rounded-lg hover:bg-muted/50 cursor-pointer">
                <RadioGroupItem value="pdf" id="pdf" />
                <div className="flex-1">
                  <Label htmlFor="pdf" className="font-medium cursor-pointer">授信写入</Label>
                  <p className="text-xs text-muted-foreground mt-1">
                    识别扫描PDF表格，自动填充金额
                  </p>
                </div>
              </div>
            </RadioGroup>
          </div>
        </Card>

        {/* 功能说明 */}
        <Card className="p-6 bg-muted/50">
          <h3 className="font-semibold mb-3">
            {mode === 'pdf' ? '授信写入功能' : '支持的功能'}
          </h3>
          {mode === 'pdf' ? (
            <ul className="grid grid-cols-1 md:grid-cols-2 gap-2 text-sm text-muted-foreground">
              <li>✓ OCR识别扫描版PDF表格</li>
              <li>✓ 自动提取机构名称和授信品种</li>
              <li>✓ 智能匹配单体表和集团表</li>
              <li>✓ 自动填充金额并标记红色</li>
              <li>✓ 删除多余的授信品种数据</li>
              <li>✓ 修改内容用红色字体显示</li>
            </ul>
          ) : (
            <ul className="grid grid-cols-1 md:grid-cols-2 gap-2 text-sm text-muted-foreground">
              <li>✓ 支持 Word、Excel、WPS 格式</li>
              <li>✓ 自动识别文档中的所有表格</li>
              <li>✓ 同义词智能匹配（如"总资产"匹配"资产总额"）</li>
              <li>✓ 多格式日期识别（2025/9、2025-9、2025.9等）</li>
              <li>✓ 自动识别单位并换算（亿元、万元、百分比）</li>
              <li>✓ 表格位置任意，自动定位</li>
            </ul>
          )}
        </Card>

        {/* 文件上传区域 */}
        {renderFileUploads()}

        {/* 操作按钮 */}
        <div className="flex gap-4 justify-center">
          <Button
            onClick={handleProcess}
            disabled={isProcessDisabled()}
            size="lg"
            className="min-w-[180px]"
          >
            {processing ? (
              <>
                <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                处理中...
              </>
            ) : mode === 'pdf' ? (
              <>
                <FileText className="mr-2 h-4 w-4" />
                识别并处理
              </>
            ) : (
              <>
                <Upload className="mr-2 h-4 w-4" />
                开始处理
              </>
            )}
          </Button>

          {mode !== 'pdf' && (
            <div className="flex flex-col items-center gap-2">
              <Button
                onClick={handleDownload}
                disabled={isDownloadDisabled()}
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
          )}
        </div>

        {/* 注意事项 */}
        <Card className="p-6 border-amber-200 bg-amber-50 dark:bg-amber-950/20">
          <div className="flex gap-3">
            <AlertCircle className="h-5 w-5 text-amber-600 flex-shrink-0 mt-0.5" />
            <div className="space-y-1 text-sm">
              <p className="font-medium text-amber-900 dark:text-amber-200">使用说明</p>
              {mode === 'pdf' ? (
                <ul className="text-amber-800 dark:text-amber-300/80 space-y-1 list-disc list-inside">
                  <li>文件A必须是扫描版PDF，包含机构名称和授信品种表格</li>
                  <li>文件B必须包含"单体"或"集团"工作表</li>
                  <li>单体表机构字段在B列，授信品种在第3行</li>
                  <li>集团表机构字段在D列，授信品种在第3行</li>
                  <li>修改的内容会用红色字体标记</li>
                </ul>
              ) : (
                <ul className="text-amber-800 dark:text-amber-300/80 space-y-1 list-disc list-inside">
                  <li>确保文件B中的空白单元格可以从文件A中找到匹配数据</li>
                  <li>单位会自动识别并进行换算，无需手动调整</li>
                  <li>百分比符号会根据表格外标注自动处理</li>
                  <li>同义词匹配功能会自动识别相似字段名称</li>
                </ul>
              )}
            </div>
          </div>
        </Card>
      </div>
    </div>
  );
}
