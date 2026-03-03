#!/bin/bash

# 测试脚本 - 验证文件处理 API

echo "开始测试文件处理 API..."

# 1. 测试健康检查
echo ""
echo "1. 检查服务状态..."
curl -I http://localhost:5000

# 2. 测试 API 端点（需要真实文件）
echo ""
echo "2. 测试文件处理 API（使用示例文件）..."
curl -X POST http://localhost:5000/api/process \
  -F "fileA=@/workspace/projects/temp/示例-数据源A.txt" \
  -F "fileB=@/workspace/projects/temp/示例-缺失B.txt" \
  -v

echo ""
echo "测试完成"
