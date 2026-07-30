import * as fs from 'fs'
import * as path from 'path'
import * as JSZip from 'jszip'
import iconv from 'iconv-lite'
// @ts-expect-error - word-extractor 没有官方类型定义
import WordExtractor from 'word-extractor'
import { DOMParser, Element as XmlElement } from '@xmldom/xmldom'
import {
  ParsedDocument,
  Paragraph as AppParagraph,
  RichTextFragment,
  RichTextContent,
} from '../types'

const wordExtractor = new WordExtractor()

/**
 * 使用 word-extractor 解析 .doc 文件，提取纯文本
 * word-extractor 正确处理 FIB 压缩、CJK 字符宽度、PieceTable 等
 */
async function extractDocText(filePath: string): Promise<string> {
  try {
    const doc = await wordExtractor.extract(filePath)
    // getBody() 直接返回完整文本字符串
    const body = doc.getBody()
    return body || ''
  } catch (error) {
    throw new Error(`word-extractor 解析失败: ${error instanceof Error ? error.message : String(error)}`)
  }
}

/**
 * 兜底：自实现 OLE2 + 多编码识别解析
 * 用于 word-extractor 失败时的备选方案
 */
function extractDocTextByBytes(fileBuffer: Buffer): string {
  try {
    let offset = 0
    if (fileBuffer[0] === 0xD0 && fileBuffer[1] === 0xCF) {
      // OLE2 头
      offset = 0x4800
    }

    const safeStart = offset
    const safeEnd = fileBuffer.length
    if (safeEnd <= safeStart) return ''

    const rawBytes = fileBuffer.slice(safeStart, safeEnd)

    // 检测编码：UTF-16 LE 中 ASCII 字符后必有 0x00
    let asciiCount = 0
    let zeroCount = 0
    for (let i = 0; i < Math.min(rawBytes.length, 2000); i++) {
      const c = rawBytes[i]
      if (c >= 0x20 && c <= 0x7E) asciiCount++
      else if (c === 0) zeroCount++
    }
    const isUtf16 = zeroCount > asciiCount * 0.5

    if (isUtf16) {
      return rawBytes.toString('utf16le').replace(/\u0000/g, '')
    }

    // ANSI 时尝试 GBK
    try {
      return iconv.decode(rawBytes, 'gbk')
    } catch {
      return rawBytes.toString('binary')
    }
  } catch {
    return ''
  }
}

/**
 * 解析 .doc 文件
 * 使用 word-extractor 库正确处理压缩文本和 CJK 字符
 */
async function parseDocFileImpl(filePath: string): Promise<string> {
  // 第一选择：word-extractor
  try {
    const text = await extractDocText(filePath)
    if (text && text.length > 0) {
      // 检查是否乱码：如果出现大量 PUA/CJK 扩展区字符，疑似乱码
      let garbledCount = 0
      const sample = text.slice(0, 5000)
      for (let i = 0; i < sample.length; i++) {
        const c = sample.charCodeAt(i)
        // 真正的乱码通常出现在 0xFFxx 等区域
        if (c >= 0xFF00 && c <= 0xFFEF) garbledCount++
      }
      // 真正的乱码比例通常很高（>20%），给予宽松阈值
      if (garbledCount < sample.length * 0.2) {
        return text
      }
    }
  } catch (error) {
    console.warn(`word-extractor 解析失败: ${error instanceof Error ? error.message : String(error)}`)
  }

  // 兜底：直接读取字节
  const buffer = fs.readFileSync(filePath)
  return extractDocTextByBytes(buffer)
}

export class WordParserService {
  async parseDocument(filePath: string): Promise<ParsedDocument> {
    // 从原始文件名获取扩展名（处理中文路径）
    const originalName = path.basename(filePath)
    const ext = path.extname(originalName).toLowerCase()
    let allParagraphs: AppParagraph[]

    if (ext === '.doc') {
      // 使用 cfb 库解析 .doc 文件
      allParagraphs = await this.parseDocFile(filePath)
    } else {
      // 使用 JSZip 解析 .docx 文件
      const buffer = fs.readFileSync(filePath)
      const zip = await JSZip.loadAsync(buffer)

      const documentXml = await zip.file('word/document.xml')?.async('string')
      if (!documentXml) {
        throw new Error('无法读取文档内容')
      }

      const relationshipsXml = await zip.file('word/_rels/document.xml.rels')?.async('string')
      allParagraphs = this.parseDocumentXml(documentXml, relationshipsXml || '')
    }

    // 只保留"用户告知书"之后的内容
    const paragraphs = this.filterAfterUserNotice(allParagraphs)

    let fullText = paragraphs.map(p => p.plainText).join('\n')
    const fileName = path.basename(filePath)

    return {
      title: this.extractTitle(paragraphs),
      paragraphs,
      fullText: fullText.trim(),
      meta: {
        parsedAt: new Date().toISOString(),
        fileName,
        paragraphCount: paragraphs.length,
      },
    }
  }

  /**
   * 解析 .doc 文件
   * 使用 cfb 库读取 OLE2 格式，提取纯文本
   */
  private async parseDocFile(filePath: string): Promise<AppParagraph[]> {
    try {
      // 读取文件为 Buffer
      const fileBuffer = fs.readFileSync(filePath)

      // 检测文件魔数 (Magic Number) 判断实际文件类型
      const isDocx = fileBuffer[0] === 0x50 && fileBuffer[1] === 0x4B // "PK"
      const isOle2 = fileBuffer[0] === 0xD0 && fileBuffer[1] === 0xCF // OLE2

      // 如果扩展名为 .doc 但内容是 docx 格式（常见于改名保存的文档）
      if (isDocx) {
        const zip = await JSZip.loadAsync(fileBuffer)
        const documentXml = await zip.file('word/document.xml')?.async('string')
        if (documentXml) {
          const relationshipsXml = await zip.file('word/_rels/document.xml.rels')?.async('string')
          return this.parseDocumentXml(documentXml, relationshipsXml || '')
        }
      }

      if (!isOle2) {
        throw new Error('文件格式不支持：文件既不是 .docx 也不是 .doc 格式，或文件已损坏')
      }

      // 使用 word-extractor 解析 OLE2 格式的 .doc 文件
      const rawText = await parseDocFileImpl(filePath)
      return this.textToParagraphs(rawText)
    } catch (error) {
      throw new Error(`解析 .doc 文件失败: ${error instanceof Error ? error.message : String(error)}`)
    }
  }

  /**
   * 将纯文本转换为段落数组
   */
  private textToParagraphs(text: string): AppParagraph[] {
    if (!text) return []

    const paragraphs: AppParagraph[] = []
    const lines = text.split(/\r?\n/)

    for (const line of lines) {
      const trimmedLine = line.trim()
      if (trimmedLine) {
        paragraphs.push({
          fragments: [{
            text: trimmedLine,
            bold: false,
            italic: false,
            underline: false,
          }],
          plainText: trimmedLine,
          style: undefined,
        })
      }
    }

    return paragraphs
  }

  /**
   * 解析 HTML 为段落
   */
  private parseHtmlToParagraphs(html: string): AppParagraph[] {
    const paragraphs: AppParagraph[] = []
    // 匹配段落标签 (p 或 div)
    const paragraphRegex = /<(?:p|div)[^>]*>([\s\S]*?)<\/(?:p|div)>/gi
    let match

    while ((match = paragraphRegex.exec(html)) !== null) {
      const content = match[1]
      const fragments: RichTextFragment[] = []

      // 解析文本片段
      let remaining = content
      while (remaining.length > 0) {
        // 匹配加粗标签 <strong> 或 <b>
        const boldMatch = remaining.match(/<(?:strong|b)[^>]*>([\s\S]*?)<\/(?:strong|b)>/i)
        if (boldMatch && boldMatch.index !== undefined) {
          // 添加之前的普通文本
          const beforeText = remaining.substring(0, boldMatch.index).trim()
          if (beforeText) {
            fragments.push({
              text: this.decodeHtmlEntities(beforeText),
              bold: false,
              italic: false,
              underline: false,
            })
          }
          // 添加加粗文本
          fragments.push({
            text: this.decodeHtmlEntities(boldMatch[1]),
            bold: true,
            italic: false,
            underline: false,
          })
          remaining = remaining.substring(boldMatch.index + boldMatch[0].length)
        } else {
          // 没有更多标签，添加剩余文本
          const text = remaining.trim()
          if (text) {
            fragments.push({
              text: this.decodeHtmlEntities(text),
              bold: false,
              italic: false,
              underline: false,
            })
          }
          break
        }
      }

      // 组装纯文本
      const plainText = fragments.map(f => f.text).join('')

      if (fragments.length > 0 && plainText.trim()) {
        paragraphs.push({
          fragments,
          plainText,
          style: undefined,
        })
      }
    }

    return paragraphs
  }

  /**
   * 解码 HTML 实体
   */
  private decodeHtmlEntities(text: string): string {
    return text
      .replace(/&nbsp;/g, ' ')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&amp;/g, '&')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      .replace(/<[^>]+>/g, '') // 移除剩余的 HTML 标签
  }

  /**
   * 过滤掉"用户告知书"之前的内容
   */
  private filterAfterUserNotice(paragraphs: AppParagraph[]): AppParagraph[] {
    const keyword = '用户告知书'
    let foundIndex = -1

    // 找到"用户告知书"所在的位置
    for (let i = 0; i < paragraphs.length; i++) {
      const text = paragraphs[i].plainText || ''
      if (text.includes(keyword)) {
        foundIndex = i
        break
      }
    }

    // 如果找到关键词，只返回之后的段落（+1 跳过关键词本身）
    if (foundIndex >= 0) {
      return paragraphs.slice(foundIndex + 1)
    }

    // 如果没找到关键词，返回所有段落
    return paragraphs
  }

  private parseDocumentXml(documentXml: string, relationshipsXml: string): AppParagraph[] {
    const paragraphs: AppParagraph[] = []
    const parser = new DOMParser()
    const doc = parser.parseFromString(documentXml, 'text/xml')

    const paraElements = doc.getElementsByTagName('w:p')
    const rels = this.parseRelationships(relationshipsXml)

    for (let i = 0; i < paraElements.length; i++) {
      const paraEl = paraElements[i] as unknown as XmlElement
      const paragraph = this.parseParagraphElement(paraEl, rels)
      paragraphs.push(paragraph)
    }

    return paragraphs
  }

  private parseRelationships(relationshipsXml: string): Map<string, string> {
    const rels = new Map<string, string>()
    if (!relationshipsXml) return rels

    const parser = new DOMParser()
    const doc = parser.parseFromString(relationshipsXml, 'text/xml')
    const linkElements = doc.getElementsByTagName('Relationship')

    for (let i = 0; i < linkElements.length; i++) {
      const el = linkElements[i] as unknown as XmlElement
      const id = el.getAttribute('Id') || ''
      const target = el.getAttribute('Target') || ''
      const type = el.getAttribute('Type') || ''

      if (type.includes('hyperlink')) {
        rels.set(id, target)
      }
    }

    return rels
  }

  private parseParagraphElement(paraEl: XmlElement, rels: Map<string, string>): AppParagraph {
    const fragments: RichTextFragment[] = []

    const runElements = paraEl.getElementsByTagName('w:r')

    for (let i = 0; i < runElements.length; i++) {
      const runEl = runElements[i] as unknown as XmlElement
      const fragment = this.parseRunElement(runEl, rels)
      fragments.push(fragment)
    }

    const hyperlinkElements = paraEl.getElementsByTagName('w:hyperlink')
    for (let i = 0; i < hyperlinkElements.length; i++) {
      const hyperlinkEl = hyperlinkElements[i] as unknown as XmlElement
      const fragment = this.parseHyperlinkElement(hyperlinkEl, rels)
      fragments.push(fragment)
    }

    const paraProps = paraEl.getElementsByTagName('w:pPr')[0] as unknown as XmlElement | undefined
    let firstLineIndent: number | undefined
    let alignment: string | undefined
    let spacingAfter: number | undefined

    if (paraProps) {
      const indentEl = paraProps.getElementsByTagName('w:ind')[0] as unknown as XmlElement | undefined
      if (indentEl) {
        const val = indentEl.getAttribute('w:firstLine')
        if (val) {
          firstLineIndent = parseInt(val, 10)
        }
      }

      const jcEl = paraProps.getElementsByTagName('w:jc')[0] as unknown as XmlElement | undefined
      if (jcEl) {
        alignment = jcEl.getAttribute('w:val') || undefined
      }

      const spacingEl = paraProps.getElementsByTagName('w:spacing')[0] as unknown as XmlElement | undefined
      if (spacingEl) {
        const after = spacingEl.getAttribute('w:after')
        if (after) {
          spacingAfter = parseInt(after, 10)
        }
      }
    }

    const plainText = fragments.map(f => f.text).join('')

    return {
      fragments,
      plainText,
      style: {
        firstLineIndent,
        alignment: alignment as 'left' | 'center' | 'right' | 'justify' | undefined,
        spacing: spacingAfter ? { after: spacingAfter } : undefined,
      },
    }
  }

  private parseRunElement(runEl: XmlElement, _rels: Map<string, string>): RichTextFragment {
    let text = ''
    let bold = false
    let italic = false
    let underline = false
    let strike = false
    let color: string | undefined
    let fontSize: number | undefined
    let fontName: string | undefined

    const tElements = runEl.getElementsByTagName('w:t')
    for (let i = 0; i < tElements.length; i++) {
      text += (tElements[i].textContent || '')
    }

    const rPrEl = runEl.getElementsByTagName('w:rPr')[0] as unknown as XmlElement | undefined
    if (rPrEl) {
      const bEl = rPrEl.getElementsByTagName('w:b')[0] as unknown as XmlElement | undefined
      if (bEl) bold = true

      const iEl = rPrEl.getElementsByTagName('w:i')[0] as unknown as XmlElement | undefined
      if (iEl) italic = true

      const uEl = rPrEl.getElementsByTagName('w:u')[0] as unknown as XmlElement | undefined
      if (uEl) underline = true

      const strikeEl = rPrEl.getElementsByTagName('w:strike')[0] as unknown as XmlElement | undefined
      if (strikeEl) strike = true

      const colorEl = rPrEl.getElementsByTagName('w:color')[0] as unknown as XmlElement | undefined
      if (colorEl) {
        const val = colorEl.getAttribute('w:val')
        if (val) {
          color = this.normalizeColor(val)
        }
      }

      const szEl = rPrEl.getElementsByTagName('w:sz')[0] as unknown as XmlElement | undefined
      if (szEl) {
        const val = szEl.getAttribute('w:val')
        if (val) {
          fontSize = parseInt(val, 10) / 2
        }
      }

      const rFontsEl = rPrEl.getElementsByTagName('w:rFonts')[0] as unknown as XmlElement | undefined
      if (rFontsEl) {
        fontName = rFontsEl.getAttribute('w:eastAsia') ||
                   rFontsEl.getAttribute('w:ascii') ||
                   undefined
      }
    }

    return {
      text,
      bold,
      italic,
      underline,
      strike,
      color,
      fontSize,
      fontName,
    }
  }

  private parseHyperlinkElement(hyperlinkEl: XmlElement, rels: Map<string, string>): RichTextFragment {
    let text = ''
    const rId = hyperlinkEl.getAttribute('r:id') || ''

    let link = ''
    if (rId && rels.has(rId)) {
      link = rels.get(rId) || ''
    }

    const tElements = hyperlinkEl.getElementsByTagName('w:t')
    for (let i = 0; i < tElements.length; i++) {
      text += (tElements[i].textContent || '')
    }

    const runElements = hyperlinkEl.getElementsByTagName('w:r')
    let bold = false
    let italic = false
    let underline = false
    let color: string | undefined

    if (runElements.length > 0) {
      const runEl = runElements[0] as unknown as XmlElement
      const rPrEl = runEl.getElementsByTagName('w:rPr')[0] as unknown as XmlElement | undefined
      if (rPrEl) {
        const bEl = rPrEl.getElementsByTagName('w:b')[0] as unknown as XmlElement | undefined
        if (bEl) bold = true
        const iEl = rPrEl.getElementsByTagName('w:i')[0] as unknown as XmlElement | undefined
        if (iEl) italic = true
        const uEl = rPrEl.getElementsByTagName('w:u')[0] as unknown as XmlElement | undefined
        if (uEl) underline = true

        const colorEl = rPrEl.getElementsByTagName('w:color')[0] as unknown as XmlElement | undefined
        if (colorEl) {
          const val = colorEl.getAttribute('w:val')
          if (val) color = this.normalizeColor(val)
        }
      }
    }

    return {
      text,
      bold,
      italic,
      underline,
      color,
      link,
    }
  }

  private extractTitle(paragraphs: AppParagraph[]): string | undefined {
    const firstNonEmpty = paragraphs.find(p => p.plainText && p.plainText.trim())
    return firstNonEmpty?.plainText?.trim()
  }

  convertToRichText(document: ParsedDocument, districtName?: string): RichTextContent {
    const paragraphs: string[] = []
    const paragraphTexts: string[] = []

    for (const paragraph of document.paragraphs) {
      let paragraphHtml = ''
      let hasContent = false

      for (const fragment of paragraph.fragments) {
        if (!fragment.text) continue

        hasContent = true
        let text = fragment.text
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')

        if (fragment.color) {
          const colorHex = fragment.color.startsWith('#') ? fragment.color.toUpperCase() : ('#' + fragment.color).toUpperCase()
          // 跳过默认黑色，避免产生多余的 <span style="color: #000000;"></span>
          const isDefaultBlack = colorHex === '#000000' || colorHex === '#000'
          if (!isDefaultBlack && text.trim()) {
            text = `<span style="color: ${colorHex};">${text}</span>`
          }
        }

        if (fragment.bold) {
          text = `<strong>${text}</strong>`
        }

        if (fragment.italic) {
          text = `<em>${text}</em>`
        }

        if (fragment.underline) {
          text = `<u>${text}</u>`
        }

        if (fragment.link) {
          text = `<a href="${fragment.link}" style="color: #333;text-decoration: underline;">${text}</a>`
        }

        paragraphHtml += text
      }

      if (hasContent) {
        // 收集段落 HTML 内容（包含样式标签）
        paragraphTexts.push(paragraphHtml)
        // 段落后加两个换行
        if (paragraph.style?.spacing?.after && paragraph.style.spacing.after > 240) {
          paragraphs.push(paragraphHtml + '<br><br>')
        } else {
          paragraphs.push(paragraphHtml + '<br>')
        }
      }
    }

    // 合并多个连续换行
    let htmlContent = paragraphs.join('')
    htmlContent = htmlContent.replace(/(<br>){3,}/g, '<br><br>')

    // 生成可复制的代码格式 - 每个段落独立
    const escapedParagraphs = paragraphTexts.map(p =>
      `'${p.replace(/\\/g, '\\\\').replace(/`/g, '\\`').replace(/\$/g, '\\$')}'`
    )
    const codeFormat = `[\n    ${escapedParagraphs.join(',\n    ')}\n  ]`

    // 纯文本内容
    const plainText = document.fullText

    return {
      htmlContent,
      plainText,
      paragraphs: paragraphTexts,
      districtName,
      codeFormat,
    }
  }

  private escapeHtml(text: string): string {
    const map: Record<string, string> = {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;',
    }
    return text.replace(/[&<>"']/g, char => map[char])
  }

  private normalizeColor(color: string): string {
    color = color.replace(/^#/, '')
    color = color.toUpperCase()
    if (/^[0-9A-F]{6}$/i.test(color)) {
      return '#' + color
    }
    return color
  }

  generateDistrictNotification(
    districtName: string,
    document: ParsedDocument
  ): string {
    const richText = this.convertToRichText(document, districtName)
    const paragraphs = richText.paragraphs

    if (paragraphs.length === 0) {
      return `  ${districtName}: {},`
    }

    // 每个段落末尾加 <br>，最后一个段落不加
    const lines = paragraphs.map((p, index) => {
      const escaped = p.replace(/\\/g, '\\\\').replace(/`/g, '\\`').replace(/\$/g, '\\$')
      if (index === paragraphs.length - 1) {
        return `      '${escaped}'`
      } else {
        return `      '${escaped}<br>' +`
      }
    })

    const code = `  ${districtName}: {
    content:
${lines.join('\n')},
  },`

    return code
  }
}

export const wordParserService = new WordParserService()
