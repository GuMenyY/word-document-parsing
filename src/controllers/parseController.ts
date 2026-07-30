import { Request, Response } from 'express'
import * as fs from 'fs'
import { wordParserService } from '../services/wordParser'
import { ApiResponse, ParseResponse } from '../types'

export async function uploadAndParse(req: Request, res: Response): Promise<void> {
  try {
    if (!req.file) {
      const response: ApiResponse = { success: false, message: '请上传 Word 文档', code: 400 }
      res.status(400).json(response)
      return
    }

    const filePath = req.file.path
    const districtName = req.body.districtName as string | undefined

    const result = await wordParserService.parseDocument(filePath)
    const richText = wordParserService.convertToRichText(result, districtName)

    fs.unlinkSync(filePath)

    const response: ParseResponse = {
      success: true,
      message: '解析成功',
      data: { result, richText },
    }
    res.json(response)
  } catch (error) {
    // 出错时也尝试清理上传的文件
    if (req.file?.path) {
      try { fs.unlinkSync(req.file.path) } catch {}
    }
    console.error('Upload parse error:', error)
    const response: ApiResponse = {
      success: false,
      message: `解析失败: ${error instanceof Error ? error.message : String(error)}`,
      code: 500,
    }
    res.status(500).json(response)
  }
}
