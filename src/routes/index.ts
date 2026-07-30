import { Router } from 'express'
import multer from 'multer'
import * as path from 'path'
import * as fs from 'fs'
import {
  uploadAndParse,
  parseDocument,
  getResult,
  getDistricts,
  generateCode,
} from '../controllers/parseController'

// 创建路由
const router = Router()

// 上传目录：打包后写到 exe 同级目录，开发时写到项目根目录
function getUploadDir(): string {
  // pkg 打包后 process.pkg 存在
  const isPkg = typeof (process as any).pkg !== 'undefined'
  const base = isPkg ? path.dirname(process.execPath) : process.cwd()
  return path.join(base, 'uploads')
}

// 配置 multer 上传
const storage = multer.diskStorage({
  destination: (_req, _file, cb) => {
    const uploadDir = getUploadDir()
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir, { recursive: true })
    }
    cb(null, uploadDir)
  },
  filename: (_req, file, cb) => {
    const uniqueSuffix = `${Date.now()}-${Math.round(Math.random() * 1e9)}`
    cb(null, `${uniqueSuffix}-${file.originalname}`)
  },
})

const upload = multer({
  storage,
  limits: {
    fileSize: 10 * 1024 * 1024, // 10MB
  },
  fileFilter: (_req, file, cb) => {
    // 只允许 Word 文档
    const allowedMimes = [
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // .docx
      'application/msword', // .doc
    ]
    // 从原始文件名获取扩展名
    const ext = path.extname(file.originalname).toLowerCase()
    const allowedExts = ['.docx', '.doc']

    if (allowedMimes.includes(file.mimetype) || allowedExts.includes(ext)) {
      cb(null, true)
    } else {
      cb(new Error('只支持 Word 文档 (.docx, .doc)'))
    }
  },
})

// 路由定义

/**
 * POST /api/upload
 * 上传并解析 Word 文档 (异步模式，返回任务ID)
 */
router.post('/upload', upload.single('file'), uploadAndParse)

/**
 * POST /api/parse
 * 直接上传并解析 Word 文档 (同步模式，直接返回结果)
 */
router.post('/parse', upload.single('file'), parseDocument)

/**
 * GET /api/result/:id
 * 获取任务解析结果
 */
router.get('/result/:id', getResult)

/**
 * GET /api/districts
 * 获取支持的区域列表
 */
router.get('/districts', getDistricts)

/**
 * POST /api/generate
 * 生成区域代码 (可直接复制到 districtNotification.ts)
 */
router.post('/generate', upload.single('file'), generateCode)

/**
 * GET /api/health
 * 健康检查
 */
router.get('/health', (_req, res) => {
  res.json({
    success: true,
    message: '服务运行正常',
    data: {
      version: '1.0.0',
      uptime: process.uptime(),
    },
  })
})

export default router
