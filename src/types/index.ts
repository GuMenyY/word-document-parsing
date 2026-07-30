/**
 * 解析结果状态
 */
export type ParseStatus = 'pending' | 'processing' | 'completed' | 'failed'

/**
 * 富文本片段
 */
export interface RichTextFragment {
  /** 文本内容 */
  text: string
  /** 是否加粗 */
  bold?: boolean
  /** 是否斜体 */
  italic?: boolean
  /** 是否下划线 */
  underline?: boolean
  /** 字体颜色 (十六进制) */
  color?: string
  /** 字体大小 */
  fontSize?: number
  /** 字体名称 */
  fontName?: string
  /** 是否删除线 */
  strike?: boolean
  /** 超链接地址 */
  link?: string
}

/**
 * 段落信息
 */
export interface Paragraph {
  /** 段落文本片段列表 */
  fragments: RichTextFragment[]
  /** 段落样式 */
  style?: {
    /** 首行缩进 (twips, 1/20 磅) */
    firstLineIndent?: number
    /** 对齐方式 */
    alignment?: 'left' | 'center' | 'right' | 'justify'
    /** 行间距 */
    spacing?: {
      before?: number
      after?: number
      line?: number
    }
  }
  /** 段落纯文本 (便于搜索) */
  plainText?: string
}

/**
 * 解析后的文档结构
 */
export interface ParsedDocument {
  /** 文档标题 */
  title?: string
  /** 段落列表 */
  paragraphs: Paragraph[]
  /** 原始文本内容 (所有文本拼接) */
  fullText: string
  /** 解析元数据 */
  meta?: {
    /** 解析时间 */
    parsedAt: string
    /** 原文件名 */
    fileName: string
    /** 段落数量 */
    paragraphCount: number
  }
}

/**
 * 转换后的富文本内容
 */
export interface RichTextContent {
  /** HTML 富文本内容 */
  htmlContent: string
  /** 纯文本内容 */
  plainText: string
  /** 段落列表 (每个段落独立) */
  paragraphs: string[]
  /** 区域名称 (可选) */
  districtName?: string
  /** 可复制的代码格式 */
  codeFormat?: string
}

/**
 * 解析任务
 */
export interface ParseTask {
  /** 任务ID */
  id: string
  /** 任务状态 */
  status: ParseStatus
  /** 原文件名 */
  fileName?: string
  /** 文件路径 */
  filePath?: string
  /** 解析结果 */
  result?: ParsedDocument
  /** 错误信息 */
  error?: string
  /** 创建时间 */
  createdAt: string
  /** 完成时间 */
  completedAt?: string
}

/**
 * API 响应基础结构
 */
export interface ApiResponse<T = any> {
  /** 是否成功 */
  success: boolean
  /** 消息 */
  message?: string
  /** 数据 */
  data?: T
  /** 错误码 */
  code?: number
}

/**
 * 上传响应
 */
export interface UploadResponse extends ApiResponse {
  data?: {
    /** 任务ID */
    taskId: string
    /** 文件名 */
    fileName: string
  }
}

/**
 * 解析响应
 */
export interface ParseResponse extends ApiResponse {
  data?: {
    /** 任务ID */
    taskId?: string
    /** 解析结果 */
    result: ParsedDocument
    /** 富文本内容 */
    richText: RichTextContent
  }
}

/**
 * 支持的区域列表
 */
export interface District {
  /** 区域编码 */
  code: string
  /** 区域名称 */
  name: string
}
