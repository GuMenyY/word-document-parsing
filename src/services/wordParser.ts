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
 * 自动编号层级定义（来自 word/numbering.xml）
 */
interface NumberingLevel {
  /** 起始值（w:start） */
  start: number
  /** 编号格式（decimal / chineseCounting / ...） */
  numFmt: string
  /** 编号模板，如 "%1、" */
  lvlText: string
  /** 编号与正文之间的分隔（nothing / space / tab） */
  suff: string
}

/**
 * 自动编号上下文：定义映射 + 计数器（按文档段落顺序推进）
 */
interface NumberingContext {
  /** numId -> abstractNumId */
  numToAbstract: Map<number, number>
  /** abstractNumId -> (ilvl -> 层级定义) */
  abstracts: Map<number, Map<number, NumberingLevel>>
  /** numId -> (ilvl -> 起始值覆盖) */
  startOverrides: Map<number, Map<number, number>>
  /** `${numId}:${ilvl}` -> 当前计数值 */
  counters: Map<string, number>
}

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
      const numberingXml = await zip.file('word/numbering.xml')?.async('string')
      allParagraphs = this.parseDocumentXml(documentXml, relationshipsXml || '', numberingXml || '')
    }

    const afterNotice = this.filterAfterUserNotice(allParagraphs)
    const { title, paragraphs } = this.extractSpecialSection(afterNotice, originalName)
    const fullText = paragraphs.map(p => p.plainText).join('\n')

    return {
      title,
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
          const numberingXml = await zip.file('word/numbering.xml')?.async('string')
          return this.parseDocumentXml(documentXml, relationshipsXml || '', numberingXml || '')
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

  /**
   * 特殊文档统一入口：承诺书 / 审核流程说明等需要"去标题、去样式、
   * 只保留正文纯文本"的文档走这里，其他文档保持原样。
   */
  private extractSpecialSection(
    paragraphs: AppParagraph[],
    fileName: string
  ): { title: string | undefined; paragraphs: AppParagraph[] } {
    if (this.isCommitmentDoc(paragraphs, fileName)) {
      return this.extractCommitmentSection(paragraphs)
    }
    if (this.isReviewProcessDoc(paragraphs, fileName)) {
      return this.extractReviewProcessSection(paragraphs)
    }
    return { title: this.extractTitle(paragraphs), paragraphs }
  }

  /** 是否为申请承诺书：内容识别为主、文件名识别为辅 */
  private isCommitmentDoc(paragraphs: AppParagraph[], fileName: string): boolean {
    if (fileName.includes('承诺书')) return true
    return paragraphs.some(p => {
      const t = p.plainText || ''
      return t.includes('承诺书') || t.includes('郑重做出以下承诺') || t.includes('郑重承诺')
    })
  }

  /** 是否为审核流程说明：内容识别为主、文件名识别为辅 */
  private isReviewProcessDoc(paragraphs: AppParagraph[], fileName: string): boolean {
    if (fileName.includes('审核流程说明')) return true
    return paragraphs.some(p => (p.plainText || '').includes('审核流程说明'))
  }

  /**
   * 审核流程说明专用处理：丢弃标题，只保留正文，
   * 每段一行纯文本、去除一切文字样式。
   */
  private extractReviewProcessSection(
    paragraphs: AppParagraph[]
  ): { title: undefined; paragraphs: AppParagraph[] } {
    // 找到标题段（"审核流程说明"），从它之后开始截取
    const titleIndex = paragraphs.findIndex(p =>
      (p.plainText || '').trim() === '审核流程说明'
      || (p.plainText || '').includes('审核流程说明')
    )
    const sliced = titleIndex >= 0 ? paragraphs.slice(titleIndex + 1) : paragraphs

    const cleaned: AppParagraph[] = []
    for (const p of sliced) {
      const text = (p.plainText || '').trim()
      // 空段落直接跳过
      if (!text) continue
      // 签名落款段才停止截取（正文如"申请人对审核结果有异议的……"不受影响）
      if (this.isSignatureLine(text)) break
      // 标题残留（如重复出现的说明标题）跳过
      if (text === '审核流程说明') continue
      cleaned.push(this.toPlainParagraph(p, text))
    }

    return { title: undefined, paragraphs: cleaned }
  }

  /**
   * 是否为签名落款段：只有"申请人/承诺人"后紧跟冒号、空白+短姓名、
   * "日期/联系电话/身份证"后紧跟冒号时才算落款。
   * 正文如"申请人对审核结果有异议的……"不受影响。
   */
  private isSignatureLine(text: string): boolean {
    return /^(申请人|承诺人)\s*[:：]/.test(text)
      || /^(申请人|承诺人)\s+\S{1,8}$/.test(text)
      || /^(日期|联系电话|身份证)\s*[:：]/.test(text)
  }

  /**
   * 把段落转为纯文本段落：去除加粗/颜色等一切文字样式，
   * 每段就是一行干净文本。
   */
  private toPlainParagraph(source: AppParagraph, text: string): AppParagraph {
    return {
      ...source,
      fragments: [{ text, bold: false, italic: false, underline: false }],
      plainText: text,
      style: undefined,
    }
  }

  /**
   * 申请承诺书专用处理：从"承诺：/承诺:"之后开始截取，
   * 有编号的保证从1开始连续，无编号的在"本人"前面补上序号，
   * 条款为纯文本、无任何样式。
   */
  private extractCommitmentSection(
    paragraphs: AppParagraph[]
  ): { title: undefined; paragraphs: AppParagraph[] } {
    // 找到"承诺：/承诺:"所在段，从它的下一段开始截取；
    // 找不到时回退到原来的"从1.开始"逻辑
    const promiseIndex = paragraphs.findIndex(p =>
      /承诺\s*[:：]/.test(p.plainText || '')
    )
    let sliced: AppParagraph[]
    if (promiseIndex >= 0) {
      sliced = paragraphs.slice(promiseIndex + 1)
      // "承诺："同段后面若还跟有实质内容，也作为第一条处理
      const tail = (paragraphs[promiseIndex].plainText || '').split(/承诺\s*[:：]/)[1]?.trim()
      if (tail) {
        sliced = [{ ...paragraphs[promiseIndex], plainText: tail, fragments: [{ text: tail }] } as AppParagraph, ...sliced]
      }
    } else {
      const startIndex = paragraphs.findIndex(p =>
        /^\s*1\s*[.．、]/.test(p.plainText || '')
      )
      if (startIndex < 0) {
        return { title: undefined, paragraphs }
      }
      sliced = paragraphs.slice(startIndex)
    }
    // 期望序号计数器：从1开始递增
    let expectedNum = 1
    const renumbered: AppParagraph[] = []

    for (const p of sliced) {
      const text = (p.plainText || '').trim()
      // 空段落直接跳过
      if (!text) continue
      // 签名落款段才停止截取（正文如"申请人对审核结果……"不受影响）
      if (this.isSignatureLine(text)) break

      const numMatch = text.match(/^\s*(\d+)\s*[.．、]\s*/)
      if (numMatch) {
        // 有编号：用期望序号重写，保证从1开始连续
        const newText = `${expectedNum}.${text.slice(numMatch[0].length)}`
        renumbered.push(this.rebuildParagraphWithNumber(p, newText))
        expectedNum++
      } else {
        // 无编号：在"本人"前面补上序号；找不到"本人"时直接加在段首
        const benIndex = text.indexOf('本人')
        const newText = benIndex >= 0
          ? `${text.slice(0, benIndex)}${expectedNum}.${text.slice(benIndex)}`
          : `${expectedNum}.${text}`
        renumbered.push(this.rebuildParagraphWithNumber(p, newText))
        expectedNum++
      }
    }

    return {
      // 承诺书不需要标题头，前面的标题/导语已丢弃，title 置空
      title: undefined,
      paragraphs: renumbered,
    }
  }

  /**
   * 用新的纯文本重建段落：承诺书条款不需要任何文字样式，
   * 直接用无样式的纯文本片段，保证输出就是 "1.xxx" 这样的干净文本。
   */
  private rebuildParagraphWithNumber(source: AppParagraph, newText: string): AppParagraph {
    return {
      ...source,
      fragments: [{ text: newText, bold: false, italic: false, underline: false }],
      plainText: newText,
      style: undefined,
    }
  }

  private parseDocumentXml(documentXml: string, relationshipsXml: string, numberingXml = ''): AppParagraph[] {
    const paragraphs: AppParagraph[] = []
    const parser = new DOMParser()
    const doc = parser.parseFromString(documentXml, 'text/xml')
    const paraElements = doc.getElementsByTagName('w:p')
    const rels = this.parseRelationships(relationshipsXml)
    // 自动编号上下文：按段落顺序推进计数器，把 Word 自动生成的序号还原出来
    const numbering = this.parseNumbering(numberingXml)

    for (let i = 0; i < paraElements.length; i++) {
      const paraEl = paraElements[i] as unknown as XmlElement
      paragraphs.push(this.parseParagraphElement(paraEl, rels, numbering))
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

  private parseParagraphElement(paraEl: XmlElement, rels: Map<string, string>, numbering?: NumberingContext): AppParagraph {
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

      // 自动编号（w:numPr）：正文的 w:t 里没有序号文本，需要从 numbering.xml 还原
      if (numbering && numbering.numToAbstract.size > 0) {
        const numPrEl = paraProps.getElementsByTagName('w:numPr')[0] as unknown as XmlElement | undefined
        if (numPrEl) {
          const ilvlEl = numPrEl.getElementsByTagName('w:ilvl')[0] as unknown as XmlElement | undefined
          const numIdEl = numPrEl.getElementsByTagName('w:numId')[0] as unknown as XmlElement | undefined
          const ilvl = parseInt(ilvlEl?.getAttribute('w:val') || '0', 10)
          const numId = parseInt(numIdEl?.getAttribute('w:val') || '-1', 10)
          const label = this.resolveNumberLabel(numbering, numId, ilvl)
          if (label && !this.hasManualNumberPrefix(fragments.map(f => f.text).join(''))) {
            fragments.unshift({ text: label, bold: false, italic: false, underline: false })
          }
        }
      }
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

  /**
   * 解析 word/numbering.xml，建立 numId -> 编号定义的映射。
   * 没有编号文件时返回空上下文，调用方按无自动编号处理。
   */
  private parseNumbering(numberingXml: string): NumberingContext {
    const ctx: NumberingContext = {
      numToAbstract: new Map(),
      abstracts: new Map(),
      startOverrides: new Map(),
      counters: new Map(),
    };
    if (!numberingXml) return ctx;

    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(numberingXml, 'text/xml');

      const abstractEls = doc.getElementsByTagName('w:abstractNum');
      for (let i = 0; i < abstractEls.length; i++) {
        const absEl = abstractEls[i] as unknown as XmlElement;
        const abstractId = parseInt(absEl.getAttribute('w:abstractNumId') || '-1', 10);
        if (abstractId < 0) continue;
        const levels = new Map<number, NumberingLevel>();
        const lvlEls = absEl.getElementsByTagName('w:lvl');
        for (let j = 0; j < lvlEls.length; j++) {
          const lvlEl = lvlEls[j] as unknown as XmlElement;
          const ilvl = parseInt(lvlEl.getAttribute('w:ilvl') || '0', 10);
          const startEl = lvlEl.getElementsByTagName('w:start')[0] as unknown as XmlElement | undefined;
          const fmtEl = lvlEl.getElementsByTagName('w:numFmt')[0] as unknown as XmlElement | undefined;
          const textEl = lvlEl.getElementsByTagName('w:lvlText')[0] as unknown as XmlElement | undefined;
          const suffEl = lvlEl.getElementsByTagName('w:suff')[0] as unknown as XmlElement | undefined;
          levels.set(ilvl, {
            start: parseInt(startEl?.getAttribute('w:val') || '1', 10),
            numFmt: fmtEl?.getAttribute('w:val') || 'decimal',
            lvlText: textEl?.getAttribute('w:val') || '%1.',
            suff: suffEl?.getAttribute('w:val') || 'tab',
          });
        }
        ctx.abstracts.set(abstractId, levels);
      }

      const numEls = doc.getElementsByTagName('w:num');
      for (let i = 0; i < numEls.length; i++) {
        const numEl = numEls[i] as unknown as XmlElement;
        const numId = parseInt(numEl.getAttribute('w:numId') || '-1', 10);
        const absEl = numEl.getElementsByTagName('w:abstractNumId')[0] as unknown as XmlElement | undefined;
        const abstractId = parseInt(absEl?.getAttribute('w:val') || '-1', 10);
        if (numId < 0 || abstractId < 0) continue;
        ctx.numToAbstract.set(numId, abstractId);
        const overrideEls = numEl.getElementsByTagName('w:lvlOverride');
        for (let j = 0; j < overrideEls.length; j++) {
          const ovEl = overrideEls[j] as unknown as XmlElement;
          const ilvl = parseInt(ovEl.getAttribute('w:ilvl') || '0', 10);
          const startEl = ovEl.getElementsByTagName('w:startOverride')[0] as unknown as XmlElement | undefined;
          const startVal = parseInt(startEl?.getAttribute('w:val') || '', 10);
          if (!Number.isNaN(startVal)) {
            if (!ctx.startOverrides.has(numId)) ctx.startOverrides.set(numId, new Map());
            ctx.startOverrides.get(numId)!.set(ilvl, startVal);
          }
        }
      }
    } catch (error) {
      console.warn('numbering.xml 解析失败，将忽略自动编号: ' + (error instanceof Error ? error.message : String(error)));
    }

    return ctx;
  }

  /**
   * 根据 (numId, ilvl) 计算当前段落的自动编号文本，并推进计数器。
   * 返回如 "二、"、"6、"；无法解析时返回空字符串。
   */
  private resolveNumberLabel(ctx: NumberingContext, numId: number, ilvl: number): string {
    const abstractId = ctx.numToAbstract.get(numId);
    if (abstractId === undefined) return '';
    const levels = ctx.abstracts.get(abstractId);
    if (!levels) return '';
    const level = levels.get(ilvl) || levels.get(0);
    if (!level) return '';

    const key = numId + ':' + ilvl;
    let current = ctx.counters.get(key);
    if (current === undefined) {
      const ov = ctx.startOverrides.get(numId)?.get(ilvl);
      current = ov !== undefined ? ov : level.start;
    }
    ctx.counters.set(key, current + 1);
    for (const k of [...ctx.counters.keys()]) {
      const parts = k.split(':');
      const kNum = Number(parts[0]);
      const kLvl = Number(parts[1]);
      if (kNum === numId && kLvl > ilvl) ctx.counters.delete(k);
    }

    const numText = this.formatNumber(current, level.numFmt);
    let label = level.lvlText.split('%1').join(numText);
    label = label.replace(/%[2-9]/g, '');
    if (level.suff === 'space') label += ' ';
    else if (level.suff === 'tab') label += '\t';
    return label;
  }

  /**
   * 段落正文是否已经带有手打编号前缀（如"一、"、"8."）。
   * 有的话不再补自动编号，避免重复。
   */
  private hasManualNumberPrefix(text: string): boolean {
    return /^[\s\t]*(?:\d+\s*[.．、]|[一二三四五六七八九十百千零〇]+\s*[、.．])/.test(text);
  }

  /** 按 Word 编号格式渲染数字，未知格式回退为阿拉伯数字。 */
  private formatNumber(value: number, numFmt: string): string {
    switch (numFmt) {
      case 'decimal':
      case 'decimalZero':
        return String(value);
      case 'upperLetter':
        return this.toAlphaNumber(value, 'upper');
      case 'lowerLetter':
        return this.toAlphaNumber(value, 'lower');
      case 'upperRoman':
        return this.toRomanNumber(value).toUpperCase();
      case 'lowerRoman':
        return this.toRomanNumber(value).toLowerCase();
      case 'chineseCounting':
      case 'chineseCountingThousand':
      case 'chineseLegalSimplified':
      case 'chineseSimplified':
        return this.toChineseNumber(value);
      case 'none':
        return '';
      case 'bullet':
        return '•';
      default:
        return String(value);
    }
  }

  /** 阿拉伯数字 -> 中文计数（一..九十九…），0 返回"零" */
  private toChineseNumber(value: number): string {
    if (value <= 0) return '零';
    const digits = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
    if (value < 10) return digits[value];
    if (value < 20) return value === 10 ? '十' : '十' + digits[value - 10];
    if (value < 100) {
      const tens = Math.floor(value / 10);
      const ones = value % 10;
      return digits[tens] + '十' + (ones ? digits[ones] : '');
    }
    if (value < 1000) {
      const hundreds = Math.floor(value / 100);
      const rest = value % 100;
      return digits[hundreds] + '百' + (rest ? this.toChineseNumber(rest) : '');
    }
    const thousands = Math.floor(value / 1000);
    const rest = value % 1000;
    return this.toChineseNumber(thousands) + '千' + (rest ? this.toChineseNumber(rest) : '');
  }

  /** 阿拉伯数字 -> 字母编号（1=A, 27=AA…） */
  private toAlphaNumber(value: number, casing: 'upper' | 'lower'): string {
    if (value <= 0) return String(value);
    let result = '';
    let n = value;
    while (n > 0) {
      n--;
      result = String.fromCharCode((n % 26) + 65) + result;
      n = Math.floor(n / 26);
    }
    return casing === 'lower' ? result.toLowerCase() : result;
  }

  /** 阿拉伯数字 -> 罗马数字 */
  private toRomanNumber(value: number): string {
    if (value <= 0) return String(value);
    const table: Array<[number, string]> = [
      [1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'],
      [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'],
      [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i'],
    ];
    let result = '';
    let n = value;
    for (const pair of table) {
      const v = pair[0];
      const s = pair[1];
      while (n >= v) {
        result += s;
        n -= v;
      }
    }
    return result;
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

      let firstContentHandled = false
      for (const fragment of paragraph.fragments) {
        if (!fragment.text) continue
        hasContent = true

        let text = fragment.text
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')

        // 段落开头的编号（如"一、"、"十二、"、"4、"）自动加粗
        // 仅在段落的第一个非空片段上处理一次，避免重复包裹
        if (!firstContentHandled && text.trim().length > 0) {
          firstContentHandled = true
          // 匹配：中文数字（一、…、十二、…、二十一、…）或阿拉伯数字（1. / 1、/ 4、）
          const chineseNumberMatch = text.match(/^([一二三四五六七八九十百千零〇两]{1,6}、|\d{1,3}[.．、])/)
          if (chineseNumberMatch) {
            const marker = chineseNumberMatch[1]
            const rest = text.slice(marker.length)
            let boldPart = `<strong>${marker}</strong>`
            if (fragment.color) {
              const colorHex = fragment.color.startsWith('#')
                ? fragment.color.toUpperCase()
                : ('#' + fragment.color).toUpperCase()
              const isDefaultBlack = colorHex === '#000000' || colorHex === '#000'
              if (!isDefaultBlack) {
                boldPart = `<span style="color: ${colorHex};">${boldPart}</span>`
              }
            }
            text = boldPart + rest
            paragraphHtml += text
            continue
          }
        }

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
