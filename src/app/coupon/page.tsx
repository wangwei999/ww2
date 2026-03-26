'use client';

import { useState, useRef } from 'react';
import { Button } from '@/components/ui/button';
import { Card } from '@/components/ui/card';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group';
import { Upload, FileSpreadsheet, Loader2, CheckCircle, AlertCircle, X, Ticket } from 'lucide-react';
import { toast } from 'sonner';
import Link from 'next/link';
import { ArrowLeft } from 'lucide-react';

type BondType = 'treasury' | 'local'; // 国债 / 地方债

export default function CouponPage() {
  const [bondType, setBondType] = useState<BondType>('treasury');
  const [file, setFile] = useState<File | null>(null);
  const [amount, setAmount] = useState<string>('');
  const [processing, setProcessing] = useState(false);

  const inputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (newFile: File | null) => {
    setFile(newFile);
    if (inputRef.current && !newFile) {
      inputRef.current.value = '';
    }
  };

  const handleDeleteFile = () => {
    setFile(null);
    if (inputRef.current) {
      inputRef.current.value = '';
    }
  };

  const handleAmountChange = (value: string) => {
    // 只允许输入数字和小数点
    const regex = /^[0-9]*\.?[0-9]*$/;
    if (value === '' || regex.test(value)) {
      setAmount(value);
    }
  };

  const handleProcess = async () => {
    // 验证
    if (!file) {
      toast.error('请上传Excel文件');
      return;
    }

    if (!amount || parseFloat(amount) <= 0) {
      toast.error('请输入有效的挑券金额');
      return;
    }

    setProcessing(true);

    try {
      const formData = new FormData();
      formData.append('file', file);
      formData.append('bondType', bondType);
      formData.append('amount', amount);

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
  };

  const isProcessDisabled = () => {
    return processing || !file || !amount || parseFloat(amount) <= 0;
  };

  return (
    <div className="min-h-screen bg-background py-12 px-4 sm:px-6 lg:px-8">
      <div className="max-w-4xl mx-auto space-y-8">
        {/* 标题 */}
        <div className="text-center space-y-2">
          <div className="flex justify-center">
            <Ticket className="h-12 w-12 text-primary" />
          </div>
          <h1 className="text-3xl font-bold tracking-tight">挑券</h1>
          <p className="text-muted-foreground">
            上传Excel文件，选择债券类型，输入挑券金额进行筛选
          </p>
        </div>

        {/* 返回主页链接 */}
        <div className="flex justify-start">
          <Link href="/">
            <Button variant="ghost" size="sm">
              <ArrowLeft className="mr-2 h-4 w-4" />
              返回主页
            </Button>
          </Link>
        </div>

        {/* 文件上传 */}
        <Card className="p-6">
          <div className="space-y-4">
            <div>
              <Label className="text-base font-semibold">上传Excel文件</Label>
              <p className="text-sm text-muted-foreground mt-1">
                上传包含债券数据的Excel文件
              </p>
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
                    onClick={handleDeleteFile}
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
                    onClick={() => inputRef.current?.click()}
                    className="w-full h-9 justify-start text-left px-3"
                  >
                    <Upload className="mr-2 h-4 w-4" />
                    <span className="text-muted-foreground">选择文件</span>
                  </Button>
                  <input
                    ref={inputRef}
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={(e) => handleFileChange(e.target.files?.[0] || null)}
                    className="hidden"
                  />
                </div>
              )}
            </div>
          </div>
        </Card>

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
                输入需要筛选的债券金额（单位：亿元）
              </p>
            </div>
            <div className="flex items-center gap-2">
              <Input
                type="text"
                placeholder="请输入金额"
                value={amount}
                onChange={(e) => handleAmountChange(e.target.value)}
                className="flex-1"
              />
              <span className="text-muted-foreground whitespace-nowrap">亿元</span>
            </div>
          </div>
        </Card>

        {/* 功能说明 */}
        <Card className="p-6 bg-muted/50">
          <h3 className="font-semibold mb-3">挑券功能说明</h3>
          <ul className="grid grid-cols-1 md:grid-cols-2 gap-2 text-sm text-muted-foreground">
            <li>✓ 支持国债和地方债筛选</li>
            <li>✓ 根据金额智能匹配债券</li>
            <li>✓ 输出筛选结果Excel文件</li>
          </ul>
        </Card>

        {/* 操作按钮 */}
        <div className="flex justify-center">
          <Button
            onClick={handleProcess}
            disabled={isProcessDisabled()}
            size="lg"
            className="min-w-[200px]"
          >
            {processing ? (
              <>
                <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                处理中...
              </>
            ) : (
              <>
                <Ticket className="mr-2 h-4 w-4" />
                开始挑券
              </>
            )}
          </Button>
        </div>

        {/* 注意事项 */}
        <Card className="p-6 border-amber-200 bg-amber-50 dark:bg-amber-950/20">
          <div className="flex gap-3">
            <AlertCircle className="h-5 w-5 text-amber-600 flex-shrink-0 mt-0.5" />
            <div className="space-y-1 text-sm">
              <p className="font-medium text-amber-900 dark:text-amber-200">使用说明</p>
              <ul className="text-amber-800 dark:text-amber-300/80 space-y-1 list-disc list-inside">
                <li>请确保上传的Excel文件包含债券数据</li>
                <li>选择正确的债券类型（国债或地方债）</li>
                <li>输入挑券金额后点击"开始挑券"进行处理</li>
              </ul>
            </div>
          </div>
        </Card>
      </div>
    </div>
  );
}
