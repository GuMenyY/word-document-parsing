# Word 文档解析服务

将 Word 文档（`.docx` / `.doc`）转换为富文本 HTML 格式，支持加粗、斜体、颜色、超链接等样式识别。

## 使用方式

### 方式一：直接运行 exe

双击 `word-parser.exe`，服务启动后访问 http://localhost:3000

### 方式二：开发模式

```bash
npm install
npm run dev
```

## 打包 exe

```bash
npm run pkg
```

生成 `word-parser.exe`，双击即可运行。

## API 接口

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/upload` | 上传并解析 Word 文档 |
| GET | `/api/health` | 健康检查 |

`/api/upload` 接受 `multipart/form-data`，字段：
- `file`：Word 文档
- `districtName`（可选）：区域名称，如 `闵行区`

## 样式识别说明

解析过程会自动识别以下样式：

- 加粗、斜体、下划线、删除线
- 字体颜色（非纯黑色才会输出 `<span style="color:">`）
- 超链接
- **段落开头的中文数字编号自动加粗**：如果段落以"一、""二、"…"十、""十一、"…等形式开头（即汉字数字+"、"），编号部分会自动用 `<strong>` 包裹，使其加粗。例如 `一、工作目标` 会被渲染成 `<strong>一、</strong>工作目标`

## 注意事项

- 文件大小限制：10MB
- 上传的文件解析后立即删除
