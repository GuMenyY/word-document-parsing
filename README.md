# Word 文档解析服务

将 Word 文档（`.docx` / `.doc`）转换为富文本 HTML 格式，支持加粗、斜体、颜色、超链接等样式识别。

## 使用方式

### 方式一：直接运行 exe（推荐）

> 直接双击 `word-parser.exe`

### 方式二：开发模式

```bash
npm install
npm run dev
```

## 打包 exe

```bash
npm run pkg
```

生成 `word-parser.exe`。

## API 接口

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/parse` | 同步解析，直接返回结果 |
| POST | `/api/upload` | 异步解析，返回 taskId |
| GET | `/api/result/:id` | 获取任务结果 |
| GET | `/api/districts` | 获取区域列表 |
| POST | `/api/generate` | 生成区域代码 |
| GET | `/api/health` | 健康检查 |

所有上传接口均接受 `multipart/form-data`，字段：
- `file`：Word 文档
- `districtName`（可选）：区域名称，如 `闵行区`

## 注意事项

- 文件大小限制：10MB
- 上传的文件解析后立即删除
- 任务结果保存在内存中，1 小时后自动清理
