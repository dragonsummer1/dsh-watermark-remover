export interface WatermarkCandidate {
    id: string;
    kind: 'xobject-image' | 'xobject-form' | 'text';
    label: string;
    pages: number[];
    confidence: 'high' | 'medium' | 'low';
    detail?: string;
}
export interface PdfAnalyzeResult {
    file: string;
    kind: 'pdf';
    pageCount: number;
    encrypted: boolean;
    candidates: WatermarkCandidate[];
    notes: string[];
}
export interface RemoveOutcome {
    id: string;
    ok: boolean;
    pages?: number[];
    reason?: string;
}
export interface PdfRemoveResult {
    ok: boolean;
    file: string;
    output: string;
    targetsRequested: string[];
    results: RemoveOutcome[];
    notes: string[];
}
export declare function analyzePdf(file: string): Promise<PdfAnalyzeResult>;
export declare function removePdf(file: string, targets: string[] | null, outPath: string): Promise<PdfRemoveResult>;
