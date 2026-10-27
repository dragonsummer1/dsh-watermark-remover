/**
 * DOCX 结构级水印分析与无损移除。
 *
 * Word 的“页面水印”存在页眉/页脚部件里：
 *  - 文字水印 = VML 艺术字 shape（<v:shape> + <v:textpath string="机密">，Word/WPS 同构）；
 *  - 图片水印 = VML <v:imagedata>（z-index<0 衬底）或 DrawingML <wp:anchor behindDoc="1">。
 * 移除 = 从对应 XML 里摘掉承载水印的 <w:r>（或元素本身）并重新打包，
 * 正文内容零改动，因此无损。.doc 老格式不支持。
 */
import fs from 'node:fs'
import AdmZip from 'adm-zip'

export interface DocxCandidate {
  id: string
  kind: 'vml-textpath' | 'vml-image' | 'behinddoc-image' | 'doc-background'
  label: string
  file: string
  confidence: 'high' | 'medium' | 'low'
  detail?: string
}

export interface DocxAnalyzeResult {
  file: string
  kind: 'docx'
  candidates: DocxCandidate[]
  notes: string[]
}

export interface DocxRemoveResult {
  ok: boolean
  file: string
  output: string
  results: Array<{ id: string; ok: boolean; file?: string; reason?: string }>
  notes: string[]
}

const HEADER_RE = /^word\/(header|footer)\d*\.xml$/
const WM_TEXT_RE = /水印|water.?mark|draft|sample|草稿/i

function escapeRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

/** 找到包住 [innerStart, innerEnd) 的最小 <w:r>…</w:r> 范围（w:r 不嵌套）。 */
function enclosingRunRange(xml: string, innerStart: number, innerEnd: number): [number, number] | null {
  const lastClose = xml.lastIndexOf('</w:r>', innerStart)
  let start = -1
  const re = /<w:r(?=[\s>])/g
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) && m.index < innerStart) {
    if (m.index > lastClose) start = m.index
  }
  if (start === -1) return null
  const close = xml.indexOf('</w:r>', innerEnd)
  if (close === -1) return null
  return [start, close + '</w:r>'.length]
}

/** 摘掉承载元素：优先整个 <w:r>，否则仅删元素本身。 */
function excise(xml: string, elemStart: number, elemEnd: number): string {
  const run = enclosingRunRange(xml, elemStart, elemEnd)
  if (run) return xml.slice(0, run[0]) + xml.slice(run[1])
  return xml.slice(0, elemStart) + xml.slice(elemEnd)
}

function scanHeaderXml(file: string, xml: string, out: DocxCandidate[]): void {
  // 1) VML 形状：文字水印（v:textpath）与衬底/绝对定位图片
  const shapeStarts: number[] = []
  const shapeRe = /<v:shape(?=[\s>])/g
  let m: RegExpExecArray | null
  while ((m = shapeRe.exec(xml))) shapeStarts.push(m.index)
  for (let i = 0; i < shapeStarts.length; i++) {
    const start = shapeStarts[i]
    const close = xml.indexOf('</v:shape>', start)
    const selfClose = xml.indexOf('/>', start)
    const end = close !== -1 ? close + '</v:shape>'.length : selfClose !== -1 ? selfClose + 2 : -1
    if (end === -1) continue
    const block = xml.slice(start, end)
    const id = `vml:${file}:${i}`
    const tp = /<v:textpath\b[^>]*string="([^"]*)"/.exec(block)
    if (tp) {
      out.push({
        id,
        kind: 'vml-textpath',
        label: `艺术字水印「${tp[1]}」`,
        file,
        confidence: 'high',
        detail: 'VML 艺术字（Word/WPS 水印标准形态）',
      })
      continue
    }
    const style = /style="([^"]*)"/.exec(block)?.[1] ?? ''
    if (/<v:imagedata\b/.test(block) && (/z-index\s*:\s*-/i.test(style) || /position\s*:\s*absolute/i.test(style))) {
      out.push({
        id,
        kind: 'vml-image',
        label: 'VML 衬底/绝对定位图片',
        file,
        confidence: 'medium',
        detail: '典型图片水印摆放（z-index<0 或 absolute）',
      })
    }
  }
  // 2) DrawingML 锚定图片：behindDoc="1" 或名称/描述含水印字样
  const anchorStarts: number[] = []
  const anchorRe = /<wp:anchor(?=[\s>])/g
  while ((m = anchorRe.exec(xml))) anchorStarts.push(m.index)
  for (let i = 0; i < anchorStarts.length; i++) {
    const start = anchorStarts[i]
    const close = xml.indexOf('</wp:anchor>', start)
    if (close === -1) continue
    const block = xml.slice(start, close + '</wp:anchor>'.length)
    const behind = /behindDoc="1"/.test(block)
    const name = /<wp:docPr\b[^>]*name="([^"]*)"/.exec(block)?.[1] ?? ''
    const descr = /descr="([^"]*)"/.exec(block)?.[1] ?? ''
    const wmNamed = WM_TEXT_RE.test(name) || WM_TEXT_RE.test(descr)
    if (!behind && !wmNamed) continue
    const extent = /<wp:extent\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(block)
    const size = extent ? `${(Number(extent[1]) / 360000).toFixed(1)}×${(Number(extent[2]) / 360000).toFixed(1)}cm` : ''
    out.push({
      id: `behind:${file}:${i}`,
      kind: 'behinddoc-image',
      label: `锚定图片${name ? `「${name}」` : ''}${size ? ` ${size}` : ''}`,
      file,
      confidence: behind || wmNamed ? (behind && wmNamed ? 'high' : 'medium') : 'low',
      detail: behind ? '衬于文字下方（behindDoc=1），典型水印摆放' : '名称/描述含水印字样',
    })
  }
}

export function analyzeDocx(file: string): DocxAnalyzeResult {
  const zip = new AdmZip(file)
  const candidates: DocxCandidate[] = []
  const notes: string[] = []
  let sawHeader = false
  for (const entry of zip.getEntries()) {
    const name = entry.entryName
    if (HEADER_RE.test(name)) {
      sawHeader = true
      try { scanHeaderXml(name, zip.readAsText(entry), candidates) } catch { notes.push(`部件 ${name} 解析失败，已跳过`) }
    }
  }
  const docXml = zip.getEntry('word/document.xml')
  if (docXml) {
    const xml = zip.readAsText(docXml)
    if (/<w:background\b/.test(xml)) {
      candidates.push({
        id: 'bg:word/document.xml:0',
        kind: 'doc-background',
        label: '文档页面背景（w:background）',
        file: 'word/document.xml',
        confidence: 'low',
        detail: '页面底色/背景，通常不是水印；仅在确认时移除',
      })
    }
  }
  if (!sawHeader) notes.push('文档没有页眉/页脚部件：Word 页面水印通常存在页眉里；若你仍看到水印，它可能已烧进正文图片像素，无法无损去除。')
  notes.push('正文/页眉图片内部的水印像素属于内容级，无法无损去除。')
  return { file, kind: 'docx', candidates, notes }
}

export function removeDocx(file: string, targets: string[] | null, outPath: string): DocxRemoveResult {
  const zip = new AdmZip(file)
  const results: DocxRemoveResult['results'] = []
  let requested: string[] = []
  if (targets === null) {
    const analysis = analyzeDocx(file)
    requested = analysis.candidates.filter((c) => c.confidence === 'high' || c.confidence === 'medium').map((c) => c.id)
    if (!requested.length) results.push({ id: '(auto)', ok: false, reason: '没有中高置信度候选；请先 analyze 人工挑选 id' })
  } else {
    requested = targets
  }
  for (const id of requested) {
    try {
      const parsed = /^([a-z-]+):(.+):(\d+)$/.exec(id)
      if (!parsed) { results.push({ id, ok: false, reason: '无法识别的候选 id' }); continue }
      const scheme = parsed[1]
      const part = parsed[2]
      const idx = Number(parsed[3])
      if (scheme === 'bg') {
        const entry = zip.getEntry('word/document.xml')
        if (!entry) { results.push({ id, ok: false, reason: '找不到 word/document.xml' }); continue }
        const xml = zip.readAsText(entry)
        const st = xml.indexOf('<w:background')
        if (st === -1) { results.push({ id, ok: false, reason: '未找到 w:background（文件可能已变化，请重新 analyze）' }); continue }
        const selfClose = xml.indexOf('/>', st)
        const close = xml.indexOf('</w:background>', st)
        let end: number
        if (close !== -1 && (selfClose === -1 || close < selfClose)) end = close + '</w:background>'.length
        else if (selfClose !== -1) end = selfClose + 2
        else { results.push({ id, ok: false, reason: 'w:background 未闭合' }); continue }
        zip.updateFile(entry, Buffer.from(xml.slice(0, st) + xml.slice(end), 'utf8'))
        results.push({ id, ok: true, file: 'word/document.xml' })
        continue
      }
      const tag = scheme === 'vml' ? '<v:shape' : scheme === 'behind' ? '<wp:anchor' : null
      if (!tag) { results.push({ id, ok: false, reason: '无法识别的候选方案（应为 vml/behind/bg）' }); continue }
      const entry = zip.getEntry(part)
      if (!entry) { results.push({ id, ok: false, reason: `找不到部件 ${part}` }); continue }
      const xml = zip.readAsText(entry)
      const re = new RegExp(escapeRe(tag) + '(?=[\\s>])', 'g')
      let start = -1
      let n = 0
      let mm: RegExpExecArray | null
      while ((mm = re.exec(xml))) { if (n++ === idx) { start = mm.index; break } }
      if (start === -1) { results.push({ id, ok: false, reason: `第 ${idx} 个元素未找到（文件可能已变化，请重新 analyze）` }); continue }
      const closeTag = scheme === 'vml' ? '</v:shape>' : '</wp:anchor>'
      const close = xml.indexOf(closeTag, start)
      const selfClose = xml.indexOf('/>', start)
      let end: number
      if (close !== -1 && (selfClose === -1 || close < selfClose)) end = close + closeTag.length
      else if (selfClose !== -1) end = selfClose + 2
      else { results.push({ id, ok: false, reason: '元素未闭合' }); continue }
      const newXml = excise(xml, start, end)
      zip.updateFile(entry, Buffer.from(newXml, 'utf8'))
      results.push({ id, ok: true, file: part })
    } catch (err) {
      results.push({ id, ok: false, reason: `处理异常：${err instanceof Error ? err.message : String(err)}` })
    }
  }
  const outBuf = zip.toBuffer()
  fs.writeFileSync(outPath, outBuf)
  const anyOk = results.some((r) => r.ok)
  return {
    ok: anyOk,
    file,
    output: outPath,
    results,
    notes: anyOk
      ? ['输出为新文件，原文件未改动；遗留的未引用关系(rels)不影响文档打开。']
      : ['没有候选被移除。'],
  }
}
