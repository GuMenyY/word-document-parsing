import { Request, Response } from 'express'
import * as fs from 'fs'
import * as path from 'path'
import { wordParserService } from '../services/wordParser'
import { taskManager } from '../services/taskManager'
import { ApiResponse, UploadResponse, ParseResponse, District } from '../types'

// 支持的区域列表
const supportedDistricts: District[] = [
  { code: 'mh', name: '闵行区' },
  { code: 'sj', name: '松江区' },
  { code: 'xh', name: '徐汇区' },
  { code: 'pd', name: '浦东新区' },
  { code: 'fx', name: '奉贤区' },
  { code: 'ja', name: '嘉定区' },
  { code: 'pt', name: '普陀区' },
  { code: 'ja2', name: '静安区' },
]

/**
 * 上传并解析 Word 文档
 * POST /api/upload
 */
export async function uploadAndParse(req: Request, res: Response): Promise<void> {
  try {
    if (!req.file) {
      const response: ApiResponse = {
        success: false,
        message: '请上传 Word 文档',
        code: 400,
      }
      res.status(400).json(response)
      return
    }

    const filePath = req.file.path
    const fileName = req.file.originalname
    const districtName = req.body.districtName as string

    // 创建任务
    const task = taskManager.createTask(fileName, filePath)
    taskManager.updateTaskStatus(task.id, 'processing')

    try {
      // 解析文档
      const result = await wordParserService.parseDocument(filePath)

      // 设置结果
      taskManager.setTaskResult(task.id, result)

      // 生成富文本
      const richText = wordParserService.convertToRichText(result, districtName)

      const response: ParseResponse = {
        success: true,
        message: '解析成功',
        data: {
          taskId: task.id,
          result,
          richText,
        },
      }

      // 清理上传的文件
      fs.unlinkSync(filePath)

      res.json(response)
    } catch (parseError) {
      taskManager.updateTaskStatus(task.id, 'failed', String(parseError))
      throw parseError
    }
  } catch (error) {
    console.error('Upload parse error:', error)
    const response: ApiResponse = {
      success: false,
      message: `解析失败: ${error instanceof Error ? error.message : String(error)}`,
      code: 500,
    }
    res.status(500).json(response)
  }
}

/**
 * 直接解析文档 (不创建任务)
 * POST /api/parse
 */
export async function parseDocument(req: Request, res: Response): Promise<void> {
  try {
    if (!req.file) {
      const response: ApiResponse = {
        success: false,
        message: '请上传 Word 文档',
        code: 400,
      }
      res.status(400).json(response)
      return
    }

    const filePath = req.file.path
    const districtName = req.body.districtName as string

    // 解析文档
    const result = await wordParserService.parseDocument(filePath)

    // 生成富文本
    const richText = wordParserService.convertToRichText(result, districtName)

    // 生成区域代码 (如果提供了区域名称)
    let districtCode = ''
    if (districtName) {
      const district = supportedDistricts.find(d => d.name === districtName)
      districtCode = district?.code || districtName
    }

    const response: ParseResponse = {
      success: true,
      message: '解析成功',
      data: {
        result,
        richText,
      },
    }

    // 清理上传的文件
    fs.unlinkSync(filePath)

    res.json(response)
  } catch (error) {
    console.error('Parse error:', error)
    const response: ApiResponse = {
      success: false,
      message: `解析失败: ${error instanceof Error ? error.message : String(error)}`,
      code: 500,
    }
    res.status(500).json(response)
  }
}

/**
 * 获取任务状态和结果
 * GET /api/result/:id
 */
export async function getResult(req: Request, res: Response): Promise<void> {
  try {
    const taskId = req.params.id
    const districtName = req.query.districtName as string | undefined

    const task = taskManager.getTask(taskId)

    if (!task) {
      const response: ApiResponse = {
        success: false,
        message: '任务不存在',
        code: 404,
      }
      res.status(404).json(response)
      return
    }

    if (task.status === 'pending' || task.status === 'processing') {
      const response: ApiResponse = {
        success: true,
        message: '任务处理中',
        data: {
          taskId: task.id,
          status: task.status,
          createdAt: task.createdAt,
        },
      }
      res.json(response)
      return
    }

    if (task.status === 'failed') {
      const response: ApiResponse = {
        success: false,
        message: `任务失败: ${task.error}`,
        code: 500,
      }
      res.status(500).json(response)
      return
    }

    // 生成富文本
    if (!task.result) {
      const response: ApiResponse = {
        success: false,
        message: '任务结果不存在',
        code: 500,
      }
      res.status(500).json(response)
      return
    }

    const richText = wordParserService.convertToRichText(task.result, districtName)

    const response: ParseResponse = {
      success: true,
      message: '获取成功',
      data: {
        taskId: task.id,
        result: task.result,
        richText,
      },
    }

    res.json(response)
  } catch (error) {
    console.error('Get result error:', error)
    const response: ApiResponse = {
      success: false,
      message: `获取结果失败: ${error instanceof Error ? error.message : String(error)}`,
      code: 500,
    }
    res.status(500).json(response)
  }
}

/**
 * 获取支持的区域列表
 * GET /api/districts
 */
export function getDistricts(_req: Request, res: Response): void {
  const response: ApiResponse = {
    success: true,
    message: '获取成功',
    data: supportedDistricts,
  }
  res.json(response)
}

/**
 * 生成区域代码 (用于复制到 districtNotification.ts)
 * POST /api/generate
 */
export async function generateCode(req: Request, res: Response): Promise<void> {
  try {
    if (!req.file) {
      const response: ApiResponse = {
        success: false,
        message: '请上传 Word 文档',
        code: 400,
      }
      res.status(400).json(response)
      return
    }

    const filePath = req.file.path
    const districtName = req.body.districtName as string

    if (!districtName) {
      const response: ApiResponse = {
        success: false,
        message: '请提供区域名称 (districtName)',
        code: 400,
      }
      res.status(400).json(response)
      return
    }

    // 解析文档
    const result = await wordParserService.parseDocument(filePath)

    // 生成区域代码
    const code = wordParserService.generateDistrictNotification(districtName, result)

    // 清理上传的文件
    fs.unlinkSync(filePath)

    const response: ApiResponse = {
      success: true,
      message: '生成成功',
      data: {
        districtName,
        code,
        paragraphCount: result.paragraphs.length,
      },
    }

    res.json(response)
  } catch (error) {
    console.error('Generate code error:', error)
    const response: ApiResponse = {
      success: false,
      message: `生成失败: ${error instanceof Error ? error.message : String(error)}`,
      code: 500,
    }
    res.status(500).json(response)
  }
}
