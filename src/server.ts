import express, { Express, Request, Response, NextFunction } from 'express'
import cors from 'cors'
import path from 'path'
import routes from './routes'

const app: Express = express()
const PORT = process.env.PORT || 3000

// 中间件
app.use(cors())
app.use(express.json())
app.use(express.urlencoded({ extended: true }))

// 静态文件服务
// - 打包为 exe 时：优先读 exe 同级的 public 目录
// - Render / 普通 node 运行时：读编译目录旁的 public
const isPkg = typeof (process as any).pkg !== 'undefined'
const publicDir = isPkg
  ? path.join(path.dirname(process.execPath), 'public')
  : path.join(__dirname, '../public')
app.use(express.static(publicDir))

// 请求日志
app.use((req: Request, _res: Response, next: NextFunction) => {
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.path}`)
  next()
})

// 路由
app.use('/api', routes)

// 根路径 → 返回前端页面
app.get('/', (_req: Request, res: Response) => {
  res.sendFile(path.join(publicDir, 'index.html'))
})

// 错误处理
app.use((err: Error, _req: Request, res: Response, _next: NextFunction) => {
  console.error('Server error:', err)
  res.status(500).json({
    success: false,
    message: err.message || '服务器内部错误',
    code: 500,
  })
})

// 启动服务器
const server = app.listen(PORT, () => {
  console.log(`
╔══════════════════════════════════════════════════════════╗
║                                                          ║
║   Word 文档解析服务已启动                                 ║
║                                                          ║
║   本地地址: http://localhost:${PORT}                        ║
║                                                          ║
║   API 接口:                                              ║
║   - POST   /api/upload     上传并解析 (异步)              ║
║   - POST   /api/parse      直接解析 (同步)                ║
║   - GET    /api/result/:id 获取任务结果                   ║
║   - GET    /api/districts  获取区域列表                   ║
║   - POST   /api/generate   生成区域代码                  ║
║   - GET    /api/health     健康检查                      ║
║                                                          ║
╚══════════════════════════════════════════════════════════╝
  `)
})

server.on('error', (err: NodeJS.ErrnoException) => {
  if (err.code === 'EADDRINUSE') {
    console.error(`\n❌ 端口 ${PORT} 已被占用！`)
    console.error(`   请检查是否已有一个服务在运行，或修改端口后重试。`)
    console.error(`   当前已有服务可直接访问: http://localhost:${PORT}\n`)
  } else {
    console.error(`\n❌ 服务启动失败: ${err.message}\n`)
  }
  console.log('按任意键退出...')
  process.stdin.setRawMode?.(true)
  process.stdin.resume()
  process.stdin.once('data', () => process.exit(1))
})

export default app
