/**
 * PDF 结构级水印分析与无损移除。
 *
 * 原理：PDF 页面内容由内容流指令驱动。水印通常表现为
 *  1) 跨页共享的 XObject（Form/Image），以 `/Name Do` 绘制；
 *  2) 每页重复出现的文字绘制（BT..ET 中的 Tj/TJ）。
 * 移除 = 从内容流删掉对应指令（q/Q 配对不受影响）+ 摘除 Resources 引用，
 * 正文像素零改动，因此无损。
 *
 * 局限（会在 analyze 的 notes 里如实告知）：
 *  - CID/子集字体编码的文字水印字节不可读，无法按文本匹配；此类水印多以
 *    Form XObject 承载，会由 XObject 候选覆盖；
 *  - 扫描图/图片内部的水印像素属于内容级，无法无损去除；
 *  - 加密 PDF 拒绝移除。
 */
import fs from 'node:fs'
import zlib from 'node:zlib'
import * as pdfLib from 'pdf-lib'

/* eslint-disable @typescript-eslint/no-explicit-any */
const PDFDocument: any = (pdfLib as any).PDFDocument
const PDFName: any = (pdfLib as any).PDFName
const PDFArray: any = (pdfLib as any).PDFArray
const PDFRawStream: any = (pdfLib as any).PDFRawStream

const WM_NAME_RE = /watermark|water.?mark|stamp|draft|logo|seal|sample|confidential|background|overlay|水印|章/i
/**
 * 水印词表：只有当内容流里直接绘制的文字命中这些词，才允许升到中/高置信度。
 * 原因：CID/子集字体编码的正文解码后是「盩/盯/%^」这类乱码，若仅凭「跨页重复」判定，
 * auto 模式会把整篇正文的 BT 块当成水印删掉。
 */
const WM_WORD_RE = /水印|内部|机密|绝密|保密|内部资料|样张|样本|草稿|仅供参考|严禁|请勿|复制|翻印|作废|试用版|未授权|watermark|confidential|draft|sample|specimen|copy|void|do\s*not\s*copy|unauthorized|trial/i

export interface WatermarkCandidate {
  id: string
  kind: 'xobject-image' | 'xobject-form' | 'text'
  label: string
  pages: number[]
  confidence: 'high' | 'medium' | 'low'
  detail?: string
}

export interface PdfAnalyzeResult {
  file: string
  kind: 'pdf'
  pageCount: number
  encrypted: boolean
  candidates: WatermarkCandidate[]
  notes: string[]
}

export interface RemoveOutcome {
  id: string
  ok: boolean
  pages?: number[]
  reason?: string
}

export interface PdfRemoveResult {
  ok: boolean
  file: string
  output: string
  targetsRequested: string[]
  results: RemoveOutcome[]
  notes: string[]
}

interface DetectedXo {
  key: string
  subtype: string
  names: Set<string>
  pages: Set<number>
  detail: string
}

interface Detection {
  doc: any
  pageCount: number
  pageDo: Set<string>[]
  xos: Map<string, DetectedXo>
  texts: Map<string, Set<number>>
}

// ── 基础工具 ──────────────────────────────────────────────────────────────

function escapeRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function scoreText(s: string): number {
  if (!s.length) return 0
  let good = 0
  for (const ch of s) {
    const c = ch.codePointAt(0) ?? 0
    if ((c >= 0x20 && c <= 0x7e) || (c >= 0x4e00 && c <= 0x9fff) || (c >= 0x3000 && c <= 0x303f) || (c >= 0xff00 && c <= 0xffef) || (c >= 0x2000 && c <= 0x206f)) good++
  }
  return good / s.length
}

/** 按可读性挑最优解码：latin1 / UTF-16BE（含 BOM）/ UTF-8。 */
function decodeTextBytes(b: Buffer): string {
  if (b.length >= 2 && b[0] === 0xfe && b[1] === 0xff) {
    try { return Buffer.from(b.subarray(2)).swap16().toString('utf16le') } catch { /* fallthrough */ }
  }
  let best = b.toString('latin1')
  let bestScore = scoreText(best)
  if (b.length >= 2 && b.length % 2 === 0) {
    try {
      const u16 = Buffer.from(b).swap16().toString('utf16le')
      const sc = scoreText(u16)
      if (sc > bestScore) { best = u16; bestScore = sc }
    } catch { /* ignore */ }
  }
  if (bestScore < 1 && b.length >= 3) {
    try {
      const u8 = b.toString('utf8')
      if (!u8.includes('\uFFFD')) {
        const sc = scoreText(u8)
        if (sc > bestScore) best = u8
      }
    } catch { /* ignore */ }
  }
  return best
}

function unescapeLiteral(s: string): Buffer {
  const out: number[] = []
  for (let i = 0; i < s.length; i++) {
    const c = s[i]
    if (c !== '\\') { out.push(s.charCodeAt(i) & 0xff); continue }
    const n = s[++i]
    if (n === undefined) break
    if (n === 'n') out.push(10)
    else if (n === 'r') out.push(13)
    else if (n === 't') out.push(9)
    else if (n === 'b') out.push(8)
    else if (n === 'f') out.push(12)
    else if (n === '(') out.push(40)
    else if (n === ')') out.push(41)
    else if (n === '\\') out.push(92)
    else if (n >= '0' && n <= '7') {
      let oct = n
      while (oct.length < 3 && i + 1 < s.length && s[i + 1] >= '0' && s[i + 1] <= '7') oct += s[++i]
      out.push(parseInt(oct, 8) & 0xff)
    } else if (n === '\r' || n === '\n') {
      if (n === '\r' && s[i + 1] === '\n') i++
    } else out.push(n.charCodeAt(0) & 0xff)
  }
  return Buffer.from(out)
}

const LITERAL_TJ_RE = /\(((?:\\.|[^\\()])*)\)\s*(?:Tj|TJ|'|")/g
const HEX_TJ_RE = /<([0-9A-Fa-f\s]+)>\s*(?:Tj|TJ)/g
const TJ_ARRAY_RE = /\[((?:[^\]\\]|\\.)*)\]\s*TJ/g

function pushTextInto(out: Set<string>, b: Buffer): void {
  const s = decodeTextBytes(b).trim()
  if (s.length >= 1 && s.length <= 60 && scoreText(s) >= 0.9) out.add(s)
}

/** 抽取内容流中所有“绘制出来的文字”（用于跨页重复判定与文本移除）。 */
function blockTexts(block: string, out: string[]): void {
  const push = (b: Buffer) => out.push(decodeTextBytes(b))
  for (const m of block.matchAll(LITERAL_TJ_RE)) push(unescapeLiteral(m[1]))
  for (const m of block.matchAll(HEX_TJ_RE)) {
    const hex = m[1].replace(/\s+/g, '')
    if (!hex) continue
    const padded = hex.length % 2 ? '0' + hex : hex
    push(Buffer.from(padded, 'hex'))
  }
  for (const m of block.matchAll(TJ_ARRAY_RE)) {
    const inner = m[1]
    for (const li of inner.matchAll(/\(((?:\\.|[^\\()])*)\)/g)) push(unescapeLiteral(li[1]))
    for (const hi of inner.matchAll(/<([0-9A-Fa-f\s]+)>/g)) {
      const hex = hi[1].replace(/\s+/g, '')
      if (!hex) continue
      push(Buffer.from(hex.length % 2 ? '0' + hex : hex, 'hex'))
    }
  }
}

function collectTexts(content: string): Set<string> {
  const out = new Set<string>()
  const raw: string[] = []
  blockTexts(content, raw)
  for (const t of raw) {
    const s = t.trim()
    if (s.length >= 1 && s.length <= 60 && scoreText(s) >= 0.9) out.add(s)
  }
  return out
}

function scanDo(content: string): Set<string> {
  const out = new Set<string>()
  for (const m of content.matchAll(/\/([^\s/<>\[\](){}%]+)\s+Do\b/g)) out.add(m[1])
  return out
}

// ── pdf-lib 对象层工具 ────────────────────────────────────────────────────

function pageResources(leaf: any, ctx: any): any {
  let node = leaf
  for (let depth = 0; node && depth < 32; depth++) {
    const res = node.get(PDFName.of('Resources'))
    if (res) return ctx.lookup(res)
    const parent = node.get(PDFName.of('Parent'))
    node = parent ? ctx.lookup(parent) : undefined
  }
  return undefined
}

function contentStreams(leaf: any, ctx: any): any[] {
  const c = leaf.get(PDFName.of('Contents'))
  const resolved = c ? ctx.lookup(c) : undefined
  if (!resolved) return []
  if (resolved instanceof PDFArray) {
    const out: any[] = []
    for (let i = 0; i < resolved.size(); i++) {
      const s = ctx.lookup(resolved.get(i))
      if (s && s.contents) out.push(s)
    }
    return out
  }
  return resolved && resolved.contents ? [resolved] : []
}

function streamBytes(s: any): Uint8Array {
  const raw: Uint8Array = s.contents
  let filter = ''
  try { filter = String(s.dict?.get?.(PDFName.of('Filter')) ?? '') } catch { /* ignore */ }
  if (!filter.includes('FlateDecode')) return raw
  try {
    const helper = (pdfLib as any).decodePDFRawStream
    if (typeof helper === 'function') return helper(s).decode()
  } catch { /* fallthrough */ }
  return new Uint8Array(zlib.inflateSync(Buffer.from(raw)))
}

function pageContentText(leaf: any, ctx: any): string {
  const parts: string[] = []
  for (const st of contentStreams(leaf, ctx)) {
    try { parts.push(Buffer.from(streamBytes(st)).toString('latin1')) } catch { /* 跳过损坏流 */ }
  }
  return parts.join('\n')
}

/** 用等价 FlateDecode 流替换旧内容流，并修正 Contents（单流或数组内引用）。 */
function replaceStream(leaf: any, ctx: any, old: any, bytes: Buffer): void {
  const dict = ctx.obj({ Length: bytes.length, Filter: 'FlateDecode' })
  const ref = ctx.register(PDFRawStream.of(dict, new Uint8Array(bytes)))
  const c = leaf.get(PDFName.of('Contents'))
  const resolved = c ? ctx.lookup(c) : undefined
  if (resolved instanceof PDFArray) {
    for (let i = 0; i < resolved.size(); i++) {
      if (ctx.lookup(resolved.get(i)) === old) { resolved.set(i, ref); return }
    }
  }
  leaf.set(PDFName.of('Contents'), ref)
}

// ── 检测核心 ─────────────────────────────────────────────────────────────

function detect(doc: any): Detection {
  const ctx = doc.context
  const pages: any[] = doc.getPages()
  const pageDo: Set<string>[] = []
  const xos = new Map<string, DetectedXo>()
  const texts = new Map<string, Set<number>>()
  for (let i = 0; i < pages.length; i++) {
    const pno = i + 1
    const leaf = pages[i].node
    const content = pageContentText(leaf, ctx)
    const dos = scanDo(content)
    pageDo.push(dos)
    const res = pageResources(leaf, ctx)
    const xdict = res ? ctx.lookup(res.get(PDFName.of('XObject'))) : undefined
    if (xdict && typeof xdict.entries === 'function') {
      for (const [k, v] of xdict.entries()) {
        const rawName = String(k).slice(1)
        const objNum: number | null = v && typeof v.objectNumber === 'number' ? v.objectNumber : null
        const key = objNum !== null ? `xobj:${objNum}` : `xobj:inline:${pno}:${rawName}`
        let sub = '?'
        let detail = ''
        const resolved = ctx.lookup(v)
        if (resolved && resolved.dict) {
          const d = resolved.dict
          sub = String(d.get(PDFName.of('Subtype')) ?? '?').replace(/^\//, '')
          if (sub === 'Image') {
            const w = d.get(PDFName.of('Width'))
            const h = d.get(PDFName.of('Height'))
            detail = `图片 ${w ?? '?'}×${h ?? '?'}`
          } else if (sub === 'Form') {
            detail = 'Form XObject'
          }
        }
        let xo = xos.get(key)
        if (!xo) {
          xo = { key, subtype: sub, names: new Set(), pages: new Set(), detail }
          xos.set(key, xo)
        }
        xo.names.add(rawName)
        if (dos.has(rawName)) xo.pages.add(pno)
      }
    }
    for (const s of collectTexts(content)) {
      let set = texts.get(s)
      if (!set) { set = new Set(); texts.set(s, set) }
      set.add(pno)
    }
  }
  return { doc, pageCount: pages.length, pageDo, xos, texts }
}

function rank(c: WatermarkCandidate['confidence']): number {
  return c === 'high' ? 0 : c === 'medium' ? 1 : 2
}

function candidatesFrom(d: Detection): WatermarkCandidate[] {
  const candidates: WatermarkCandidate[] = []
  for (const xo of d.xos.values()) {
    if (xo.pages.size === 0) continue
    const nameHit = [...xo.names].some((n) => WM_NAME_RE.test(n))
    const allPages = xo.pages.size === d.pageCount
    let confidence: WatermarkCandidate['confidence']
    if (d.pageCount >= 2 && allPages && nameHit) confidence = 'high'
    else if (d.pageCount >= 2 && allPages) confidence = 'medium'
    else if (d.pageCount >= 2 && nameHit) confidence = 'medium'
    else confidence = 'low'
    if (!allPages && !nameHit && d.pageCount > 1) continue
    const kind = xo.subtype === 'Image' ? 'xobject-image' as const : 'xobject-form' as const
    candidates.push({
      id: xo.key,
      kind,
      label: `XObject「${[...xo.names].join('/')}」`,
      pages: [...xo.pages].sort((a, b) => a - b),
      confidence,
      detail: `${xo.detail || xo.subtype}，绘制于 ${xo.pages.size}/${d.pageCount} 页`,
    })
  }
  if (d.pageCount >= 2) {
    for (const [s, pages] of d.texts.entries()) {
      if (pages.size !== d.pageCount) continue
      if (/^[\d\s.,\-\/:：.%页第]*$/u.test(s)) continue
      const wordHit = WM_WORD_RE.test(s)
      if (!wordHit) {
        candidates.push({
          id: `text:${s}`,
          kind: 'text',
          label: `每页重复文字「${s}」`,
          pages: [...pages].sort((a, b) => a - b),
          confidence: 'low',
          detail: scoreText(s) >= 0.9
            ? '跨页重复、但未命中水印词表：可能是正文，确认是水印后再显式指定'
            : '疑似受 CID/子集字体编码影响的正文文字（字节不可读），切勿自动移除',
        })
        continue
      }
      candidates.push({
        id: `text:${s}`,
        kind: 'text',
        label: `每页重复文字「${s}」`,
        pages: [...pages].sort((a, b) => a - b),
        confidence: scoreText(s) >= 0.9 ? 'high' : 'medium',
        detail: '内容流直接绘制的文字，命中水印词表；移除以整块 BT..ET 为单位',
      })
    }
  }
  candidates.sort((a, b) => rank(a.confidence) - rank(b.confidence))
  return candidates
}

// ── 对外 API ─────────────────────────────────────────────────────────────

export async function analyzePdf(file: string): Promise<PdfAnalyzeResult> {
  const bytes = fs.readFileSync(file)
  const doc = await PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false })
  const notes: string[] = []
  const encrypted = !!doc.isEncrypted
  if (encrypted) notes.push('PDF 已加密：可分析，但移除前需先解密（本工具不处理加密文档）。')
  const d = detect(doc)
  const candidates = candidatesFrom(d)
  if (d.pageCount === 1) notes.push('单页 PDF：无跨页对照，置信度一律降为 low。若已知水印文字，可在 remove 时用 "text:水印文字" 直接指定。')
  notes.push('文字水印若经 CID/子集字体编码，字节不可读，无法按文本匹配；此类水印多以 Form XObject 承载，会由 XObject 候选覆盖。')
  notes.push('若候选都不能对应你看到的水印，它可能是扫描图/图片内部的水印像素——无法无损去除。')
  return { file, kind: 'pdf', pageCount: d.pageCount, encrypted, candidates, notes }
}

async function removeXobject(d: Detection, id: string): Promise<RemoveOutcome> {
  const xo = d.xos.get(id)
  if (!xo) return { id, ok: false, reason: '候选不存在（文件与 analyze 时可能不一致，请重新 analyze）' }
  const ctx = d.doc.context
  const pages: any[] = d.doc.getPages()
  const touched: number[] = []
  for (let i = 0; i < pages.length; i++) {
    const pno = i + 1
    if (![...xo.names].some((n) => d.pageDo[i].has(n))) continue
    const leaf = pages[i].node
    const res = pageResources(leaf, ctx)
    const xdict = res ? ctx.lookup(res.get(PDFName.of('XObject'))) : undefined
    if (!xdict || typeof xdict.entries !== 'function') continue
    let contentChanged = false
    for (const st of contentStreams(leaf, ctx)) {
      const dec = Buffer.from(streamBytes(st))
      const s = dec.toString('latin1')
      let ns = s
      for (const nm of xo.names) ns = ns.replace(new RegExp('\\/' + escapeRe(nm) + '\\s+Do\\b', 'g'), '')
      if (ns !== s) {
        replaceStream(leaf, ctx, st, zlib.deflateSync(Buffer.from(ns, 'latin1')))
        contentChanged = true
      }
    }
    if (contentChanged) {
      for (const [k, v] of [...xdict.entries()]) {
        const key = v && typeof v.objectNumber === 'number'
          ? `xobj:${v.objectNumber}`
          : `xobj:inline:${pno}:${String(k).slice(1)}`
        if (key === id) xdict.delete(k)
      }
      touched.push(pno)
    }
  }
  if (!touched.length) return { id, ok: false, reason: '未在任何页面内容流中发现该对象的绘制指令' }
  return { id, ok: true, pages: touched }
}

function isRemovableTextBlock(block: string, target: string): boolean {
  const texts: string[] = []
  blockTexts(block, texts)
  const shown = texts.filter((t) => t.trim().length > 0)
  if (!shown.length) return false
  return shown.every((t) => {
    const s = t.trim()
    if (s.includes(target)) return true
    if (/^[\d\s.,\-\/:：.%]*$/u.test(s)) return true
    return false
  })
}

function removeText(d: Detection, target: string): RemoveOutcome {
  const id = `text:${target}`
  if (!target.trim()) return { id, ok: false, reason: '空目标' }
  const ctx = d.doc.context
  const pages: any[] = d.doc.getPages()
  const touched: number[] = []
  for (let i = 0; i < pages.length; i++) {
    const leaf = pages[i].node
    let changed = false
    for (const st of contentStreams(leaf, ctx)) {
      const s = Buffer.from(streamBytes(st)).toString('latin1')
      const ns = s.replace(/(?<![A-Za-z])BT[\s\S]*?ET(?![A-Za-z])/g, (block) => (isRemovableTextBlock(block, target) ? '' : block))
      if (ns !== s) {
        replaceStream(leaf, ctx, st, zlib.deflateSync(Buffer.from(ns, 'latin1')))
        changed = true
      }
    }
    if (changed) touched.push(i + 1)
  }
  if (!touched.length) return { id, ok: false, reason: '内容流中未找到匹配的文字块' }
  return { id, ok: true, pages: touched }
}

/**
 * 删除不可达的间接对象。pdf-lib 的 save() 会把 context 里所有对象都写出来，
 * 因此移除水印后那层图片（含 SMask）仍会以字节形式留在文件里 —— 体积不降反升，
 * 水印像素也没真正消失。这里从 trailer(Root/Info/Encrypt) 出发做可达性标记，其余删掉。
 */
function gcUnreachable(doc: any): number {
  const ctx = doc.context
  const seen = new Set<string>()
  const keyOf = (ref: any) => `${ref.objectNumber}:${ref.generationNumber}`
  const walk = (obj: any): void => {
    if (!obj || typeof obj !== 'object') return
    const t = obj.constructor?.name ?? ''
    if (t === 'PDFRef') {
      const k = keyOf(obj)
      if (seen.has(k)) return
      seen.add(k)
      walk(ctx.lookup(obj))
      return
    }
    if (typeof obj.entries === 'function') { for (const [, v] of obj.entries()) walk(v); return }
    if (typeof obj.asArray === 'function') { for (const v of obj.asArray()) walk(v); return }
    if (obj.dict) walk(obj.dict)
  }
  for (const k of ['Root', 'Info', 'Encrypt']) walk(ctx.trailerInfo?.[k])
  let removed = 0
  for (const [ref] of ctx.enumerateIndirectObjects()) {
    if (!seen.has(keyOf(ref))) { if (ctx.delete(ref)) removed++ }
  }
  return removed
}

export async function removePdf(file: string, targets: string[] | null, outPath: string): Promise<PdfRemoveResult> {
  const bytes = fs.readFileSync(file)
  const doc = await PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false })
  if (doc.isEncrypted) throw new Error('PDF 已加密，拒绝移除操作；请先解密。')
  const d = detect(doc)
  const results: RemoveOutcome[] = []
  let requested: string[] = []
  if (targets === null) {
    if (d.pageCount < 2) {
      results.push({ id: '(auto)', ok: false, reason: '单页 PDF 不做自动移除（跨页证据不足）；请显式给 targets 或用 text:水印文字' })
    } else {
      requested = candidatesFrom(d).filter((c) => c.confidence === 'high' || c.confidence === 'medium').map((c) => c.id)
      if (!requested.length) results.push({ id: '(auto)', ok: false, reason: '没有中高置信度候选；请先 analyze 人工挑选，或用 text:水印文字 指定' })
    }
  } else {
    requested = targets
  }
  for (const id of requested) {
    if (id.startsWith('xobj:')) results.push(await removeXobject(d, id))
    else if (id.startsWith('text:')) results.push(removeText(d, id.slice(5)))
    else results.push({ id, ok: false, reason: '无法识别的候选 id（应为 xobj:N 或 text:文字）' })
  }
  const anyOk = results.some((r) => r.ok)
  const dropped = anyOk ? gcUnreachable(doc) : 0
  const outBytes = await doc.save({ useObjectStreams: true, addDefaultPage: false })
  fs.writeFileSync(outPath, outBytes)
  return {
    ok: anyOk,
    file,
    output: outPath,
    targetsRequested: requested,
    results,
    notes: anyOk
      ? ['输出为新文件，原文件未改动。', `已清理 ${dropped} 个失去引用的间接对象（水印图像像素不再残留在文件里）。`]
      : ['没有任何候选被移除；输出文件与原文件内容等效。'],
  }
}
