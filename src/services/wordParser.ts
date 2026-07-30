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

async function extractDocText(filePath: string): Promise<string> {
  try {
    const doc = await wordExtractor.extract(filePath)
    return doc.getBody() || ''
  } catch (error) {
    throw new Error(`word-extractor 解析失败: ${error instanceof Error ? error.message : String(error)}`)
  }
}

function extractDocTextByBytes(fileBuffer: Buffer): string {
  try {
    let offset = 0
    if (fileBuffer[0] === 0xD0 && fileBuffer[1] === 0xCF) {
      offset = 0x4800
    }

    const safeEnd = fileBuffer.length
    if (safeEnd <= offset) return ''

    const rawBytes = fileBuffer.slice(offset, safeEnd)

    let asciiCount = 0
    let zeroCount = 0
    for (let i = 0; i < Math.min(rawBytes.length, 2000); i++) {
      const c = rawBytes[i]
      if (c >= 0x20 && c <= 0x7E) asciiCount++
      else if (c === 0) zeroCount++
    }

    if (zeroCount > asciiCount * 0.5) {
      return rawBytes.toString('utf16le').replace(/\u0000/g, '')
    }

    try {
      return iconv.decode(rawBytes, 'gbk')
    } catch {
      return rawBytes.toString('binary')
    }
  } catch {
    return ''
  }
}

async function parseDocFileImpl(filePath: string): Promise<string> {
  try {
    const text = await extractDocText(filePath)
    if (text && text.length > 0) {
      let garbledCount = 0
      const sample = text.slice(0, 5000)
      for (let i = 0; i < sample.length; i++) {
        const c = sample.charCodeAt(i)
        if (c >= 0xFF00 && c <= 0xFFEF) garbledCount++
      }
      if (garbledCount < sample.length * 0.2) {
        return text
      }
    }
  } catch (error) {
    console.warn(`word-extractor 解析失败: ${error instanceof Error ? error.message : String(error)}`)
  }

  const buffer = fs.readFileSync(filePath)
  return extractDocTextByBytes(buffer)
}

export class WordParserService {
  async parseDocument(filePath: string): Promise<ParsedDocument> {
    const originalName = path.basename(filePath)
    const ext = path.extname(originalName).toLowerCase()
    let allParagraphs: AppParagraph[]

    if (ext === '.doc') {
      allParagraphs = await this.parseDocFile(filePath)
    } else {
      const buffer = fs.readFileSync(filePath)
      const zip = await JSZip.loadAsync(buffer)

      const documentXml = await zip.file('word/document.xml')?.async('string')
      if (!documentXml) {
        throw new Error('无法读取文档内容')
      }

      const relationshipsXml = await zip.file('word/_rels/document.xml.rels')?.async('string')
      allParagraphs = this.parseDocumentXml(documentXml, relationshipsXml || '')
    }

    const paragraphs = this.filterAfterUserNotice(allParagraphs)
    const fullText = paragraphs.map(p => p.plainText).join('\n')

    return {
      title: this.extractTitle(paragraphs),
      paragraphs,
      fullText: fullText.trim(),
      meta: {
        parsedAt: new Date().toISOString(),
        fileName: path.basename(filePath),
        paragraphCount: paragraphs.length,
      },
    }
  }

  private async parseDocFile(filePath: string): Promise<AppParagraph[]> {
    try {
      const fileBuffer = fs.readFileSync(filePath)
      const isDocx = fileBuffer[0] === 0x50 && fileBuffer[1] === 0x4B
      const isOle2 = fileBuffer[0] === 0xD0 && fileBuffer[1] === 0xCF

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

      const rawText = await parseDocFileImpl(filePath)
      return this.textToParagraphs(rawText)
    } catch (error) {
      throw new Error(`解析 .doc 文件失败: ${error instanceof Error ? error.message : String(error)}`)
    }
  }

  private textToParagraphs(text: string): AppParagraph[] {
    if (!text) return []

    return text.split(/\r?\n/)
      .map(line => line.trim())
      .filter(line => line.length > 0)
      .map(line => ({
        fragments: [{ text: line, bold: false, italic: false, underline: false }],
        plainText: line,
        style: undefined,
      }))
  }

  private filterAfterUserNotice(paragraphs: AppParagraph[]): AppParagraph[] {
    const keyword = '用户告知书'
    const foundIndex = paragraphs.findIndex(p => (p.plainText || '').includes(keyword))

    if (foundIndex >= 0) {
      return paragraphs.slice(foundIndex + 1)
    }
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
      paragraphs.push(this.parseParagraphElement(paraEl, rels))
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

    for (let i = 0; i < paraEl.getElementsByTagName('w:r').length; i++) {
      const runEl = paraEl.getElementsByTagName('w:r')[i] as unknown as XmlElement
      fragments.push(this.parseRunElement(runEl))
    }

    for (let i = 0; i < paraEl.getElementsByTagName('w:hyperlink').length; i++) {
      const hyperlinkEl = paraEl.getElementsByTagName('w:hyperlink')[i] as unknown as XmlElement
      fragments.push(this.parseHyperlinkElement(hyperlinkEl, rels))
    }

    const paraProps = paraEl.getElementsByTagName('w:pPr')[0] as unknown as XmlElement | undefined
    let firstLineIndent: number | undefined
    let alignment: string | undefined
    let spacingAfter: number | undefined

    if (paraProps) {
      const indentEl = paraProps.getElementsByTagName('w:ind')[0] as unknown as XmlElement | undefined
      const indVal = indentEl?.getAttribute('w:firstLine')
      if (indVal) firstLineIndent = parseInt(indVal, 10)

      const jcEl = paraProps.getElementsByTagName('w:jc')[0] as unknown as XmlElement | undefined
      alignment = jcEl?.getAttribute('w:val') || undefined

      const spacingEl = paraProps.getElementsByTagName('w:spacing')[0] as unknown as XmlElement | undefined
      const afterVal = spacingEl?.getAttribute('w:after')
      if (afterVal) spacingAfter = parseInt(afterVal, 10)
    }

    return {
      fragments,
      plainText: fragments.map(f => f.text).join(''),
      style: {
        firstLineIndent,
        alignment: alignment as 'left' | 'center' | 'right' | 'justify' | undefined,
        spacing: spacingAfter ? { after: spacingAfter } : undefined,
      },
    }
  }

  private parseRunElement(runEl: XmlElement): RichTextFragment {
    let text = ''
    const tElements = runEl.getElementsByTagName('w:t')
    for (let i = 0; i < tElements.length; i++) {
      text += tElements[i].textContent || ''
    }

    let bold = false, italic = false, underline = false, strike = false
    let color: string | undefined
    let fontSize: number | undefined
    let fontName: string | undefined

    const rPrEl = runEl.getElementsByTagName('w:rPr')[0] as unknown as XmlElement | undefined
    if (rPrEl) {
      if (rPrEl.getElementsByTagName('w:b')[0]) bold = true
      if (rPrEl.getElementsByTagName('w:i')[0]) italic = true
      if (rPrEl.getElementsByTagName('w:u')[0]) underline = true
      if (rPrEl.getElementsByTagName('w:strike')[0]) strike = true

      const colorEl = rPrEl.getElementsByTagName('w:color')[0] as unknown as XmlElement | undefined
      const colorVal = colorEl?.getAttribute('w:val')
      if (colorVal) color = this.normalizeColor(colorVal)

      const szEl = rPrEl.getElementsByTagName('w:sz')[0] as unknown as XmlElement | undefined
      const szVal = szEl?.getAttribute('w:val')
      if (szVal) fontSize = parseInt(szVal, 10) / 2

      const rFontsEl = rPrEl.getElementsByTagName('w:rFonts')[0] as unknown as XmlElement | undefined
      if (rFontsEl) {
        fontName = rFontsEl.getAttribute('w:eastAsia') || rFontsEl.getAttribute('w:ascii') || undefined
      }
    }

    return { text, bold, italic, underline, strike, color, fontSize, fontName }
  }

  private parseHyperlinkElement(hyperlinkEl: XmlElement, rels: Map<string, string>): RichTextFragment {
    const rId = hyperlinkEl.getAttribute('r:id') || ''
    const link = rels.get(rId) || ''

    let text = ''
    const tElements = hyperlinkEl.getElementsByTagName('w:t')
    for (let i = 0; i < tElements.length; i++) {
      text += tElements[i].textContent || ''
    }

    let bold = false, italic = false, underline = false
    let color: string | undefined

    const firstRun = hyperlinkEl.getElementsByTagName('w:r')[0] as unknown as XmlElement | undefined
    const rPrEl = firstRun?.getElementsByTagName('w:rPr')[0] as unknown as XmlElement | undefined
    if (rPrEl) {
      if (rPrEl.getElementsByTagName('w:b')[0]) bold = true
      if (rPrEl.getElementsByTagName('w:i')[0]) italic = true
      if (rPrEl.getElementsByTagName('w:u')[0]) underline = true
      const colorEl = rPrEl.getElementsByTagName('w:color')[0] as unknown as XmlElement | undefined
      const colorVal = colorEl?.getAttribute('w:val')
      if (colorVal) color = this.normalizeColor(colorVal)
    }

    return { text, bold, italic, underline, color, link }
  }

  private extractTitle(paragraphs: AppParagraph[]): string | undefined {
    return paragraphs.find(p => p.plainText?.trim())?.plainText?.trim()
  }

  convertToRichText(document: ParsedDocument, districtName?: string): RichTextContent {
    const paragraphTexts: string[] = []
    const paragraphsWithBreaks: string[] = []

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
          const colorHex = fragment.color.startsWith('#')
            ? fragment.color.toUpperCase()
            : ('#' + fragment.color).toUpperCase()
          const isDefaultBlack = colorHex === '#000000' || colorHex === '#000'
          if (!isDefaultBlack && text.trim()) {
            text = `<span style="color: ${colorHex};">${text}</span>`
          }
        }

        if (fragment.bold) text = `<strong>${text}</strong>`
        if (fragment.italic) text = `<em>${text}</em>`
        if (fragment.underline) text = `<u>${text}</u>`
        if (fragment.link) {
          text = `<a href="${fragment.link}" style="color: #333;text-decoration: underline;">${text}</a>`
        }

        paragraphHtml += text
      }

      if (hasContent) {
        paragraphTexts.push(paragraphHtml)
        const br = paragraph.style?.spacing?.after && paragraph.style.spacing.after > 240
          ? '<br><br>'
          : '<br>'
        paragraphsWithBreaks.push(paragraphHtml + br)
      }
    }

    let htmlContent = paragraphsWithBreaks.join('')
    htmlContent = htmlContent.replace(/(<br>){3,}/g, '<br><br>')

    const escapedParagraphs = paragraphTexts.map(p =>
      `'${p.replace(/\\/g, '\\\\').replace(/`/g, '\\`').replace(/\$/g, '\\$')}'`
    )
    const codeFormat = `[\n    ${escapedParagraphs.join(',\n    ')}\n  ]`

    return {
      htmlContent,
      plainText: document.fullText,
      paragraphs: paragraphTexts,
      districtName,
      codeFormat,
    }
  }

  private normalizeColor(color: string): string {
    color = color.replace(/^#/, '').toUpperCase()
    return /^[0-9A-F]{6}$/i.test(color) ? '#' + color : color
  }
}

export const wordParserService = new WordParserService()
