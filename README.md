# Word 文档解析服务

将 Word 文档（`.docx` / `.doc`）转换为富文本 HTML 格式，支持加粗、斜体、颜色、超链接等样式识别。

## 使用方式

### 方式一：直接运行 exe（推荐）

双击 `启动服务.bat`，服务启动后访问 http://localhost:3000

> 直接双击 `word-parser.exe` 也可以，但出错时窗口会立即关闭，建议用 `.bat` 启动。

### 方式二：开发模式

```bash
npm install
npm run dev
```

## 打包 exe

```bash
npm run pkg
```

生成 `word-parser.exe`，将 `word-parser.exe` 和 `启动服务.bat` 放在同一目录即可分发。

## API 接口

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/upload` | 上传并解析 Word 文档 |
| GET | `/api/health` | 健康检查 |

`/api/upload` 接受 `multipart/form-data`，字段：
- `file`：Word 文档
- `districtName`（可选）：区域名称，如 `闵行区`

## 注意事项

- 文件大小限制：10MB
- 上传的文件解析后立即删除
