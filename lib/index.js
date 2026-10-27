/**
 * dsh-watermark-remover — PDF/DOCX 结构级去水印工具包。
 *
 * 用法：先 analyze（列出候选 id 与置信度）→ remove（按 id 无损移除，输出新文件）。
 * 移除只动结构层（内容流指令 / 页眉水印节点），正文像素零改动，因此无损。
 * 扫描图/图片内部的水印像素属于内容级，analyze 会如实提示无法无损处理。
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { defineTool } from '@deepseek-ai/dsh-tools';
import z from '@deepseek-ai/schemastery';
import { analyzePdf, removePdf } from './pdf-watermark.js';
import { analyzeDocx, removeDocx } from './docx-watermark.js';
export const name = 'dsh-watermark-remover';
export const inject = ['tools'];
export const Config = z.object({
    outSuffix: z.string().default('.no-watermark'),
});
function resolveInput(p) {
    const t = String(p ?? '').trim();
    if (!t)
        throw new Error('path 不能为空');
    const expanded = t.startsWith('~') ? path.join(os.homedir(), t.slice(1)) : t;
    return path.resolve(expanded);
}
function defaultOutput(input, suffix) {
    const dir = path.dirname(input);
    const ext = path.extname(input);
    const base = path.basename(input, ext);
    return path.join(dir, `${base}${suffix}${ext}`);
}
function guardOutput(input, out) {
    const a = path.resolve(out);
    const b = path.resolve(input);
    if (a === b)
        throw new Error('输出路径不能与原文件相同（本工具绝不覆盖原文件）');
    return a;
}
export function apply(ctx, config) {
    ctx.effect(() => ctx.tools.register(defineTool({
        name: '_dsh_watermark_remover_analyze',
        description: '分析 PDF/DOCX 的结构级水印候选，返回候选 id 列表与置信度（先 analyze，再 remove）',
        parameters: {
            path: { type: 'string', required: true, description: '文件路径（.pdf 或 .docx）' },
        },
        output: {
            schema: { type: 'string' },
            render: (_args, value) => [{ type: 'text', text: String(value) }],
        },
        async execute(args) {
            const file = resolveInput(args.path);
            if (!fs.existsSync(file))
                throw new Error(`文件不存在: ${file}`);
            const ext = path.extname(file).toLowerCase();
            if (ext === '.pdf')
                return JSON.stringify(await analyzePdf(file), null, 2);
            if (ext === '.docx')
                return JSON.stringify(analyzeDocx(file), null, 2);
            throw new Error('仅支持 .pdf 与 .docx（.doc 老格式请先用 Word 另存为 .docx）');
        },
    })), 'dsh-doc-watermark-remover: analyze');
    ctx.effect(() => ctx.tools.register(defineTool({
        name: '_dsh_watermark_remover_remove',
        description: '按 analyze 的候选 id 无损移除 PDF/DOCX 水印，输出新文件（不覆盖原件）',
        parameters: {
            path: { type: 'string', required: true, description: '原文件路径（.pdf 或 .docx）' },
            targets: { type: 'array', items: { type: 'string' }, description: '候选 id 列表（来自 analyze）；缺省=自动移除全部中高置信度候选' },
            outPath: { type: 'string', description: '输出路径；缺省=原文件名 + .no-watermark 后缀' },
        },
        output: {
            schema: { type: 'string' },
            render: (_args, value) => [{ type: 'text', text: String(value) }],
        },
        async execute(args) {
            const file = resolveInput(args.path);
            if (!fs.existsSync(file))
                throw new Error(`文件不存在: ${file}`);
            const ext = path.extname(file).toLowerCase();
            if (ext !== '.pdf' && ext !== '.docx')
                throw new Error('仅支持 .pdf 与 .docx');
            const out = guardOutput(file, args.outPath ? resolveInput(args.outPath) : defaultOutput(file, config.outSuffix));
            fs.mkdirSync(path.dirname(out), { recursive: true });
            const targets = args.targets && args.targets.length ? args.targets : null;
            const result = ext === '.pdf'
                ? await removePdf(file, targets, out)
                : removeDocx(file, targets, out);
            return JSON.stringify(result, null, 2);
        },
    })), 'dsh-doc-watermark-remover: remove');
}
//# sourceMappingURL=index.js.map