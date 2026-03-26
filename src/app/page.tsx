'use client';

import { useState, useEffect, useRef } from 'react';
import { Button } from '@/components/ui/button';
import { Card } from '@/components/ui/card';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group';
import { Upload, FileSpreadsheet, Download, Loader2, CheckCircle, AlertCircle, X, FileText, Ticket } from 'lucide-react';
import { toast } from 'sonner';

type ProcessMode = 'coupon' | 'normal' | 'credit' | 'pdf' | 'basic';
type BondType = 'treasury' | 'local'; // 国债 / 地方债

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
  const [mode, setMode] = useState<ProcessMode>('coupon');
  const [fileA, setFileA] = useState<File | null>(null);
  const [fileB, setFileB] = useState<File | null>(null);
  const [pdfFile, setPdfFile] = useState<File | null>(null);
  const [excelFile, setExcelFile] = useState<File | null>(null);
  // 基础数据模式文件
  const [enterpriseNameFile, setEnterpriseNameFile] = useState<File | null>(null);
  const [qichachaDataFile, setQichachaDataFile] = useState<File | null>(null);
  const [reportFieldsFile, setReportFieldsFile] = useState<File | null>(null);
  // 挑券模式
  const [couponFile, setCouponFile] = useState<File | null>(null);
  const [bondType, setBondType] = useState<BondType>('treasury');
  const [couponAmount, setCouponAmount] = useState<string>('');
  const [excludedBonds, setExcludedBonds] = useState<string>(''); // 禁挑券
  
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

  // 基础数据模式文件处理
  const handleEnterpriseNameFileChange = (file: File | null) => {
    setEnterpriseNameFile(file);
    clearProcessedState();
  };

  const handleQichachaDataFileChange = (file: File | null) => {
    setQichachaDataFile(file);
    clearProcessedState();
  };

  const handleReportFieldsFileChange = (file: File | null) => {
    setReportFieldsFile(file);
    clearProcessedState();
  };

  // 挑券模式文件处理
  const handleCouponFileChange = (file: File | null) => {
    setCouponFile(file);
    clearProcessedState();
  };

  const handleCouponAmountChange = (value: string) => {
    // 支持多金额输入，用逗号、空格或换行分隔
    setCouponAmount(value);
  };

  // 解析多金额输入
  const parseAmounts = (value: string): number[] => {
    if (!value.trim()) return [];
    
    // 按中英文逗号、空格、换行分隔
    const parts = value.split(/[,，\s\n]+/).filter(s => s.trim() !== '');
    
    const amounts: number[] = [];
    for (const part of parts) {
      const num = parseFloat(part.trim());
      if (!isNaN(num) && num > 0) {
        amounts.push(num);
      }
    }
    
    return amounts;
  };

  const handleModeChange = (newMode: ProcessMode) => {
    setMode(newMode);
    clearProcessedState();
  };

  const handleProcess = async () => {
    // 挑券模式处理
    if (mode === 'coupon') {
      if (!couponFile) {
        toast.error('请上传Excel文件');
        return;
      }

      const amounts = parseAmounts(couponAmount);
      if (amounts.length === 0) {
        toast.error('请输入有效的挑券金额');
        return;
      }

      setProcessing(true);

      try {
        const formData = new FormData();
        formData.append('file', couponFile);
        formData.append('bondType', bondType);
        // 发送多个金额，用逗号分隔
        formData.append('amounts', amounts.join(','));
        // 发送禁挑券参数
        formData.append('excludedBonds', excludedBonds.trim());

        const response = await fetch('/api/process-coupon', {
          method: 'POST',
          body: formData,
        });

        if (!response.ok) {
          const errorData = await response.json().catch(() => ({}));
          throw new Error(errorData.error || '处理失败');
        }

        // 从响应头获取文件名
        const contentDisposition = response.headers.get('Content-Disposition');
        let filename = `挑券结果_${Date.now()}.xlsx`;
        if (contentDisposition) {
          const filenameMatch = contentDisposition.match(/filename\*?=['"]?(?:UTF-\d['"]*)?([^;'"]+)/i);
          if (filenameMatch) {
            filename = decodeURIComponent(filenameMatch[1]);
          }
        }

        // 下载文件
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        link.click();
        window.URL.revokeObjectURL(url);

        toast.success('挑券处理完成，文件已下载！');
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : '处理失败，请重试';
        toast.error(errorMessage);
        console.error('处理错误:', error);
      } finally {
        setProcessing(false);
      }
      return;
    }

    if (mode === 'basic') {
      // 基础数据处理
      console.log('基础数据处理 - enterpriseNameFile:', enterpriseNameFile?.name, 
                  'qichachaDataFile:', qichachaDataFile?.name, 
                  'reportFieldsFile:', reportFieldsFile?.name);
      
      if (!enterpriseNameFile || !qichachaDataFile) {
        toast.error('请上传企业名称文件和企查查数据文件');
        return;
      }

      setProcessing(true);
      
      try {
        const formData = new FormData();
        formData.append('enterpriseNameFile', enterpriseNameFile);
        formData.append('qichachaDataFile', qichachaDataFile);
        if (reportFieldsFile) {
          formData.append('reportFieldsFile', reportFieldsFile);
        }

        const response = await fetch('/api/process-basic', {
          method: 'POST',
          body: formData,
        });

        if (!response.ok) {
          const errorData = await response.json().catch(() => ({}));
          throw new Error(errorData.error || '处理失败');
        }

        // 从响应头获取文件名
        const contentDisposition = response.headers.get('Content-Disposition');
        let filename = `基础数据处理结果_${Date.now()}.xlsx`;
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
        
        toast.success('基础数据处理完成，文件已下载！');
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : '处理失败，请重试';
        toast.error(errorMessage);
        console.error('处理错误:', error);
      } finally {
        setProcessing(false);
      }
      return;
    }

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
    // 挑券模式
    if (mode === 'coupon') {
      return (
        <div className="space-y-6">
          {/* 文件上传 */}
          <FileUpload
            label="上传Excel文件"
            description="上传包含债券数据的Excel文件（支持 .xls 和 .xlsx 格式）"
            file={couponFile}
            onFileChange={handleCouponFileChange}
            acceptedTypes=".xlsx,.xls"
          />

          {/* 债券类型选择 */}
          <Card className="p-6">
            <div className="space-y-4">
              <Label className="text-base font-semibold">选择债券类型</Label>
              <RadioGroup
                value={bondType}
                onValueChange={(value) => setBondType(value as BondType)}
                className="grid grid-cols-2 gap-4"
              >
                <div className="flex items-center space-x-2 p-4 border rounded-lg hover:bg-muted/50 cursor-pointer">
                  <RadioGroupItem value="treasury" id="treasury" />
                  <div className="flex-1">
                    <Label htmlFor="treasury" className="font-medium cursor-pointer">国债</Label>
                    <p className="text-xs text-muted-foreground mt-1">
                      筛选国债类型债券
                    </p>
                  </div>
                </div>
                <div className="flex items-center space-x-2 p-4 border rounded-lg hover:bg-muted/50 cursor-pointer">
                  <RadioGroupItem value="local" id="local" />
                  <div className="flex-1">
                    <Label htmlFor="local" className="font-medium cursor-pointer">地方债</Label>
                    <p className="text-xs text-muted-foreground mt-1">
                      筛选地方债类型债券
                    </p>
                  </div>
                </div>
              </RadioGroup>
            </div>
          </Card>

          {/* 挑券金额输入 */}
          <Card className="p-6">
            <div className="space-y-4">
              <div>
                <Label className="text-base font-semibold">挑券金额</Label>
                <p className="text-sm text-muted-foreground mt-1">
                  支持输入多个金额，用中英文逗号、空格或换行分隔（单位：万元）
                </p>
              </div>
              <textarea
                placeholder="请输入金额，多个金额用逗号、空格或换行分隔&#10;例如：5000, 4000 或 5000，4000（中英文逗号均可）"
                value={couponAmount}
                onChange={(e) => handleCouponAmountChange(e.target.value)}
                className="flex min-h-[100px] w-full rounded-md border border-input bg-background px-3 py-2 text-sm ring-offset-background placeholder:text-muted-foreground focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 disabled:cursor-not-allowed disabled:opacity-50"
              />
              {/* 显示解析后的金额列表 */}
              {couponAmount.trim() && (
                <div className="space-y-2">
                  <Label className="text-sm font-medium">已解析的金额列表：</Label>
                  <div className="flex flex-wrap gap-2">
                    {parseAmounts(couponAmount).map((amount, index) => (
                      <div
                        key={index}
                        className="px-3 py-1 bg-primary/10 text-primary rounded-full text-sm font-medium"
                      >
                        第{index + 1}笔: {amount.toLocaleString()}万元
                      </div>
                    ))}
                  </div>
                  {parseAmounts(couponAmount).length > 1 && (
                    <p className="text-xs text-muted-foreground">
                      共 {parseAmounts(couponAmount).length} 笔金额，合计 {parseAmounts(couponAmount).reduce((a, b) => a + b, 0).toLocaleString()} 万元
                    </p>
                  )}
                </div>
              )}
            </div>
          </Card>

          {/* 禁挑券输入 */}
          <Card className="p-6">
            <div className="space-y-4">
              <div>
                <Label className="text-base font-semibold">禁挑券（可选）</Label>
                <p className="text-sm text-muted-foreground mt-1">
                  输入需要排除的债券，支持全局禁挑或指定某笔金额禁挑
                </p>
              </div>
              <textarea
                placeholder="全局禁挑：输入代码或简称，如 250206 或 国开&#10;指定金额禁挑：/序号+关键词，如 /2250206 表示第2笔禁挑250206&#10;多个用逗号、空格或换行分隔（序号只支持1-9）"
                value={excludedBonds}
                onChange={(e) => setExcludedBonds(e.target.value)}
                className="flex min-h-[100px] w-full rounded-md border border-input bg-background px-3 py-2 text-sm ring-offset-background placeholder:text-muted-foreground focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 disabled:cursor-not-allowed disabled:opacity-50"
              />
              {/* 显示解析后的禁挑券列表 */}
              {excludedBonds.trim() && (
                <div className="space-y-2">
                  <Label className="text-sm font-medium">已解析的禁挑券：</Label>
                  <div className="flex flex-wrap gap-2">
                    {excludedBonds.split(/[,，\s\n]+/).filter(s => s.trim()).map((item, index) => {
                      const trimmed = item.trim();
                      // 序号只取1位数字（1-9）
                      const groupMatch = trimmed.match(/^\/([1-9])(.+)$/);
                      const isGlobal = !groupMatch;
                      const keyword = groupMatch ? groupMatch[2] : trimmed;
                      const groupIndex = groupMatch ? groupMatch[1] : null;
                      
                      return (
                        <div
                          key={index}
                          className={`px-3 py-1 rounded-full text-sm font-medium ${
                            isGlobal 
                              ? 'bg-red-100 text-red-700 dark:bg-red-950/30 dark:text-red-400' 
                              : 'bg-orange-100 text-orange-700 dark:bg-orange-950/30 dark:text-orange-400'
                          }`}
                        >
                          {isGlobal ? (
                            <span>{keyword} (全局)</span>
                          ) : (
                            <span>{keyword} (第{groupIndex}笔)</span>
                          )}
                        </div>
                      );
                    })}
                  </div>
                  <p className="text-xs text-muted-foreground">
                    数字精确匹配债券代码(B列)，文字模糊匹配债券简称(C列)；
                    <span className="text-orange-600 dark:text-orange-400">橙色</span>表示指定金额禁挑，<span className="text-red-600 dark:text-red-400">红色</span>表示全局禁挑
                  </p>
                </div>
              )}
            </div>
          </Card>
        </div>
      );
    }

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

    if (mode === 'basic') {
      return (
        <div className="space-y-6">
          <FileUpload
            label="企业名称"
            description="上传包含企业名称数据的文件"
            file={enterpriseNameFile}
            onFileChange={handleEnterpriseNameFileChange}
            acceptedTypes=".xlsx,.xls,.csv"
          />

          <FileUpload
            label="企查查数据"
            description="上传企查查数据文件"
            file={qichachaDataFile}
            onFileChange={handleQichachaDataFileChange}
            acceptedTypes=".xlsx,.xls,.csv"
          />

          <FileUpload
            label="报表字段（可选）"
            description="上传包含【行业代码】【行政区划代码】【银行信息】表的文件"
            file={reportFieldsFile}
            onFileChange={handleReportFieldsFileChange}
            acceptedTypes=".xlsx,.xls,.csv"
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
    if (mode === 'coupon') {
      return !couponFile || parseAmounts(couponAmount).length === 0;
    }
    if (mode === 'pdf') {
      return !pdfFile || !excelFile;
    }
    if (mode === 'basic') {
      // 基础数据模式只需要企业名称和企查查数据文件，报表字段可选
      return !enterpriseNameFile || !qichachaDataFile;
    }
    return !fileA || !fileB;
  };

  // 获取下载按钮是否禁用（挑券模式、授信写入模式和基础数据模式直接下载，不需要下载按钮）
  const isDownloadDisabled = () => {
    if (mode === 'coupon' || mode === 'pdf' || mode === 'basic') return true;
    return downloading || !hasProcessedFile;
  };

  // 获取处理按钮文本
  const getProcessButtonText = () => {
    if (processing) return '处理中...';
    if (mode === 'coupon') return '开始挑券';
    if (mode === 'pdf') return '识别并处理';
    return '开始处理';
  };

  // 获取处理按钮图标
  const getProcessButtonIcon = () => {
    if (processing) return <Loader2 className="mr-2 h-4 w-4 animate-spin" />;
    if (mode === 'coupon') return <Ticket className="mr-2 h-4 w-4" />;
    if (mode === 'pdf') return <FileText className="mr-2 h-4 w-4" />;
    return <Upload className="mr-2 h-4 w-4" />;
  };

  return (
    <div className="min-h-screen bg-background py-12 px-4 sm:px-6 lg:px-8">
      <div className="max-w-4xl mx-auto space-y-8">
        {/* 标题 */}
        <div className="text-center space-y-2">
          <div className="flex justify-center">
            <FileSpreadsheet className="h-12 w-12 text-primary" />
          </div>
          <h1 className="text-3xl font-bold tracking-tight">业务数据填充</h1>
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
              className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-4"
            >
              <div className="flex items-center space-x-2 p-4 border rounded-lg hover:bg-muted/50 cursor-pointer">
                <RadioGroupItem value="coupon" id="coupon" />
                <div className="flex-1">
                  <Label htmlFor="coupon" className="font-medium cursor-pointer">挑券</Label>
                  <p className="text-xs text-muted-foreground mt-1">
                    根据金额筛选国债/地方债
                  </p>
                </div>
              </div>
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
              <div className="flex items-center space-x-2 p-4 border rounded-lg hover:bg-muted/50 cursor-pointer">
                <RadioGroupItem value="basic" id="basic" />
                <div className="flex-1">
                  <Label htmlFor="basic" className="font-medium cursor-pointer">基础数据</Label>
                  <p className="text-xs text-muted-foreground mt-1">
                    处理企业名称、企查查数据、报表字段
                  </p>
                </div>
              </div>
            </RadioGroup>
          </div>
        </Card>

        {/* 功能说明 */}
        <Card className="p-6 bg-muted/50">
          <h3 className="font-semibold mb-3">
            {mode === 'coupon' ? '挑券功能' : mode === 'pdf' ? '授信写入功能' : mode === 'basic' ? '基础数据处理功能' : '支持的功能'}
          </h3>
          {mode === 'coupon' ? (
            <ul className="grid grid-cols-1 md:grid-cols-2 gap-2 text-sm text-muted-foreground">
              <li>✓ 支持国债和地方债筛选</li>
              <li>✓ 根据金额智能匹配债券</li>
              <li>✓ 支持多金额分批挑券</li>
              <li>✓ 债券集合间自动隔离</li>
            </ul>
          ) : mode === 'pdf' ? (
            <ul className="grid grid-cols-1 md:grid-cols-2 gap-2 text-sm text-muted-foreground">
              <li>✓ OCR识别扫描版PDF表格</li>
              <li>✓ 自动提取机构名称和授信品种</li>
              <li>✓ 智能匹配单体表和集团表</li>
              <li>✓ 自动填充金额并标记红色</li>
              <li>✓ 删除多余的授信品种数据</li>
              <li>✓ 修改内容用红色字体显示</li>
            </ul>
          ) : mode === 'basic' ? (
            <ul className="grid grid-cols-1 md:grid-cols-2 gap-2 text-sm text-muted-foreground">
              <li>✓ B列：企查查D列数据</li>
              <li>✓ C列：含"公司"填C01，否则C02</li>
              <li>✓ D列：企查查V列→行业代码转换</li>
              <li>✓ E列：企查查N/M列→行政区划转换</li>
              <li>✓ F列：企查查T列数据</li>
              <li>✓ G列：企业规模转换代码</li>
              <li>✓ I列：H列→银行信息转换</li>
              <li>✓ L列：K列+随机4位数</li>
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
            {getProcessButtonIcon()}
            {getProcessButtonText()}
          </Button>

          {mode !== 'coupon' && mode !== 'pdf' && mode !== 'basic' && (
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
              {mode === 'coupon' ? (
                <ul className="text-amber-800 dark:text-amber-300/80 space-y-1 list-disc list-inside">
                  <li>请确保上传的Excel文件包含债券数据</li>
                  <li>选择正确的债券类型（国债或地方债）</li>
                  <li>支持输入多个挑券金额，用逗号、空格或换行分隔</li>
                  <li>多金额模式下，每个金额对应的债券集合间会空一行隔离</li>
                  <li>空行F列显示下一个债券集合的挑券金额</li>
                </ul>
              ) : mode === 'pdf' ? (
                <ul className="text-amber-800 dark:text-amber-300/80 space-y-1 list-disc list-inside">
                  <li>文件A必须是扫描版PDF，包含机构名称和授信品种表格</li>
                  <li>文件B必须包含"单体"或"集团"工作表</li>
                  <li>单体表机构字段在B列，授信品种在第3行</li>
                  <li>集团表机构字段在D列，授信品种在第3行</li>
                  <li>修改的内容会用红色字体标记</li>
                </ul>
              ) : mode === 'basic' ? (
                <ul className="text-amber-800 dark:text-amber-300/80 space-y-1 list-disc list-inside">
                  <li>B列：填入企查查D列数据</li>
                  <li>C列：含"公司"填C01，不含填C02</li>
                  <li>D列：V列行业名称→行业代码表匹配→填入代码</li>
                  <li>E列：N/M列地区名称→行政区划表匹配→填入代码</li>
                  <li>F列：T列内容，三类有限责任公司统一填B01</li>
                  <li>G列：L→CS01，M→CS02，S→CS03，XS→CS04</li>
                  <li>I列：H列银行名称标准化→银行信息表匹配→填入代码</li>
                  <li>L列：K列内容+随机4位数字，重复则在首行标注</li>
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
