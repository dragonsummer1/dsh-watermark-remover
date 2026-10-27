import type { Context } from '@deepseek-ai/cordis';
import z from '@deepseek-ai/schemastery';
export declare const name = "dsh-watermark-remover";
export declare const inject: string[];
export interface Config {
    outSuffix: string;
}
export declare const Config: z<Schemastery.ObjectS<NoInfer<{
    outSuffix: z<string, string, "defined">;
}>>, Schemastery.ObjectT<NoInfer<{
    outSuffix: z<string, string, "defined">;
}>>, "plain">;
export declare function apply(ctx: Context, config: Config): void;
