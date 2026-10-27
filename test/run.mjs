/**
 * 端到端自测：构造带水印的 PDF（共享 Form XObject + 每页直绘文字）与
 * DOCX（页眉 VML 艺术字水印），跑 analyze → remove → 复验：
 * 水印消失、正文内容原样保留。
 */
import fs from 'node:fs'
import path from 'node:path'
import zlib from 'node:zlib'
import { fileURLToPath } from 'node:url'
import { PDFDocument, PDFName, PDFRawStream } from 'pdf-lib'
import AdmZip from 'adm-zip'
import { analyzePdf, removePdf } from '../lib/pdf-watermark.js'
import { analyzeDocx, removeDocx } from '../lib/docx-watermark.js'

const dir = path.dirname(fileURLToPath(import.meta.url))
const pdfPath = path.join(dir, 'sample-watermarked.pdf')
const docxPath = path.join(dir, 'sample-watermarked.docx')

let pass = 0
let fail = 0
function check(label, cond, extra = '') {
  if (cond) { pass++; console.log(`  PASS  ${label}`) }
  else { fail++; console.error(`  FAIL  ${label} ${extra}`) }
}

// ── 生成样例 PDF ──────────────────────────────────────────────────────────
const doc = await PDFDocument.create()
const ctx = doc.context
const wmBytes = zlib.deflateSync(Buffer.from('BT /F1 48 Tf 1 0 0 1 120 400 Tm (WATERMARK) Tj ET'))
const wmRef = ctx.register(PDFRawStream.of(
  ctx.obj({
    Type: 'XObject', Subtype: 'Form', BBox: [0, 0, 612, 792],
    Resources: ctx.obj({ Font: ctx.obj({ F1: ctx.obj({ Type: 'Font', Subtype: 'Type1', BaseFont: 'Helvetica' }) }) }),
  }),
  new Uint8Array(wmBytes),
))
for (let i = 1; i <= 3; i++) {
  const page = doc.addPage([612, 792])
  page.node.set(PDFName.of('Resources'), ctx.obj({
    Font: ctx.obj({ F1: ctx.obj({ Type: 'Font', Subtype: 'Type1', BaseFont: 'Helvetica' }) }),
    XObject: ctx.obj({ Wm1: wmRef }),
  }))
  const content = `BT /F1 14 Tf 72 700 Td (Hello page ${i}) Tj ET\nq /Wm1 Do Q\nBT /F1 36 Tf 200 300 Td (DRAFT) Tj ET`
  const cb = zlib.deflateSync(Buffer.from(content))
  const cs = ctx.register(PDFRawStream.of(ctx.obj({ Length: cb.length, Filter: 'FlateDecode' }), new Uint8Array(cb)))
  page.node.set(PDFName.of('Contents'), cs)
}
fs.writeFileSync(pdfPath, await doc.save())
console.log('sample PDF written')

// ── PDF：analyze ─────────────────────────────────────────────────────────
console.log('\n[PDF analyze]')
const pa = await analyzePdf(pdfPath)
check('pageCount=3', pa.pageCount === 3)
const xo = pa.candidates.find((c) => c.id.startsWith('xobj:'))
check('发现跨页 XObject 候选', !!xo && xo.pages.join(',') === '1,2,3', JSON.stringify(pa.candidates))
check('发现每页重复文字 DRAFT', pa.candidates.some((c) => c.id === 'text:DRAFT'))
check('正文 Hello page 不在候选里', !pa.candidates.some((c) => c.label.includes('Hello')))

// ── PDF：remove（显式 targets） ───────────────────────────────────────────
console.log('\n[PDF remove explicit]')
const out1 = path.join(dir, 'out-explicit.pdf')
const targets = [xo.id, 'text:DRAFT']
const pr = await removePdf(pdfPath, targets, out1)
check('ok=true', pr.ok === true)
check('两个目标都移除成功', pr.results.every((r) => r.ok), JSON.stringify(pr.results))

const outDoc = await PDFDocument.load(fs.readFileSync(out1))
const pdfLibAny = await import('pdf-lib')
const decode = pdfLibAny.decodePDFRawStream
const leaf = outDoc.getPages()[1].node
const st = outDoc.context.lookup(leaf.get(PDFName.of('Contents')))
const s = Buffer.from(decode(st).decode()).toString('latin1')
check('第2页正文保留', s.includes('Hello page 2'))
check('第2页 /Wm1 Do 已消失', !/\/Wm1\s+Do\b/.test(s))
check('第2页 DRAFT 文字已消失', !s.includes('DRAFT'))

const pa2 = await analyzePdf(out1)
check('输出文件再 analyze 无跨页候选', pa2.candidates.length === 0, JSON.stringify(pa2.candidates))

// ── PDF：remove（auto 模式，副本） ────────────────────────────────────────
console.log('\n[PDF remove auto]')
const copy2 = path.join(dir, 'sample-auto.pdf')
fs.copyFileSync(pdfPath, copy2)
const pr2 = await removePdf(copy2, null, path.join(dir, 'out-auto.pdf'))
check('auto 移除中高置信度候选', pr2.ok && pr2.results.length === 2 && pr2.results.every((r) => r.ok), JSON.stringify(pr2.results))

// ── 生成样例 DOCX ────────────────────────────────────────────────────────
const zip = new AdmZip()
zip.addFile('[Content_Types].xml', Buffer.from(
  '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'))
zip.addFile('_rels/.rels', Buffer.from(
  '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'))
zip.addFile('word/document.xml', Buffer.from(
  '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>正文第一段：合同金额一百万元整。</w:t></w:r></w:p><w:p><w:r><w:t>正文第二段：本合同自双方签字之日起生效。</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>'))
zip.addFile('word/_rels/document.xml.rels', Buffer.from(
  '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>'))
zip.addFile('word/header1.xml', Buffer.from(
  '<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word"><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shape id="PowerPlusWaterMarkObject1" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:400pt;height:150pt;z-index:-251654144;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;\u5B8B\u4F53&quot;;font-size:1pt" string="\u5185\u90E8\u8D44\u6599"/></v:shape></w:pict></w:r></w:p></w:hdr>'))
zip.writeZip(docxPath)
console.log('\nsample DOCX written')

// ── DOCX：analyze ────────────────────────────────────────────────────────
console.log('\n[DOCX analyze]')
const da = analyzeDocx(docxPath)
const vml = da.candidates.find((c) => c.kind === 'vml-textpath')
check('发现页眉艺术字水印「内部资料」', !!vml && vml.confidence === 'high' && vml.label.includes('内部资料'), JSON.stringify(da.candidates))

// ── DOCX：remove（显式 targets） ─────────────────────────────────────────
console.log('\n[DOCX remove explicit]')
const dout = path.join(dir, 'out-explicit.docx')
const dr = removeDocx(docxPath, [vml.id], dout)
check('ok=true', dr.ok === true, JSON.stringify(dr.results))
const outZip = new AdmZip(dout)
const header = outZip.readAsText(outZip.getEntry('word/header1.xml'))
const body = outZip.readAsText(outZip.getEntry('word/document.xml'))
check('页眉 v:shape 已摘除', !header.includes('<v:shape'))
check('正文第一段原样保留', body.includes('正文第一段：合同金额一百万元整。'))
check('正文第二段原样保留', body.includes('正文第二段：本合同自双方签字之日起生效。'))

// ── DOCX：remove（auto 模式，副本） ──────────────────────────────────────
console.log('\n[DOCX remove auto]')
const dcopy = path.join(dir, 'sample-auto.docx')
fs.copyFileSync(docxPath, dcopy)
const dr2 = removeDocx(dcopy, null, path.join(dir, 'out-auto.docx'))
check('auto 移除中高置信度候选', dr2.ok && dr2.results.length === 1 && dr2.results[0].ok, JSON.stringify(dr2.results))

console.log(`\n===== ${pass} passed, ${fail} failed =====`)
process.exit(fail ? 1 : 0)
