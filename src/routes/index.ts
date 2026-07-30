import { Router } from 'express'
import multer from 'multer'
import * as path from 'path'
import * as fs from 'fs'
import { uploadAndParse } from '../controllers/parseController'

const router = Router()

// 上传目录：打包后写到 exe 同级目录，开发时写到项目根目录
function getUploadDir(): string {
  const isPkg = typeof (process as any).pkg !== 'undefined'
  const base = isPkg ? path.dirname(process.execPath) : process.cwd()
  return path.join(base, 'uploads')
}

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
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase()
    const allowedMimes = [
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'application/msword',
    ]
    if (allowedMimes.includes(file.mimetype) || ['.docx', '.doc'].includes(ext)) {
      cb(null, true)
    } else {
      cb(new Error('只支持 Word 文档 (.docx, .doc)'))
    }
  },
})

router.post('/upload', upload.single('file'), uploadAndParse)

router.get('/health', (_req, res) => {
  res.json({
    success: true,
    message: '服务运行正常',
    data: { version: '1.0.0', uptime: process.uptime() },
  })
})

export default router
