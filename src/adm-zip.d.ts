/**
 * @types/adm-zip@0.6.0 发布包损坏（官方 tarball 缺 index.d.ts），此处按本插件
 * 实际用到的方法面提供环境声明；后续若安装到完好的类型包，本文件自动退位。
 */
declare module 'adm-zip' {
  export interface AdmZipEntry {
    entryName: string
    name: string
    isDirectory: boolean
    comment: string
    header: Record<string, unknown>
    getData(): Buffer
  }
  export default class AdmZip {
    constructor(file?: string | Buffer)
    getEntries(): AdmZipEntry[]
    getEntry(name: string): AdmZipEntry | null
    readAsText(entry: AdmZipEntry): string
    readAsTextAsync(entry: AdmZipEntry, callback: (data: Buffer, err: Error | null) => void): void
    updateFile(entry: AdmZipEntry, content: Buffer | string): void
    deleteFile(entry: AdmZipEntry | string): void
    toBuffer(): Buffer
  }
}
