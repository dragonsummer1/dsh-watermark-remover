export interface DocxCandidate {
    id: string;
    kind: 'vml-textpath' | 'vml-image' | 'behinddoc-image' | 'doc-background';
    label: string;
    file: string;
    confidence: 'high' | 'medium' | 'low';
    detail?: string;
}
export interface DocxAnalyzeResult {
    file: string;
    kind: 'docx';
    candidates: DocxCandidate[];
    notes: string[];
}
export interface DocxRemoveResult {
    ok: boolean;
    file: string;
    output: string;
    results: Array<{
        id: string;
        ok: boolean;
        file?: string;
        reason?: string;
    }>;
    notes: string[];
}
export declare function analyzeDocx(file: string): DocxAnalyzeResult;
export declare function removeDocx(file: string, targets: string[] | null, outPath: string): DocxRemoveResult;
