import { Ctor } from '@univerjs/core';
import { IFunctionNames } from '../basics/function';
import { BaseFunction } from '../functions/base-function';
export declare const ENGINE_FORMULA_PLUGIN_CONFIG_KEY = "engine-formula.config";
export declare const DEFAULT_CYCLE_REFERENCE_COUNT = 1;
export declare const ENGINE_FORMULA_CYCLE_REFERENCE_COUNT = "CYCLE_REFERENCE_COUNT";
export declare const configSymbol: unique symbol;
export interface IUniverEngineFormulaConfig {
    notExecuteFormula?: boolean;
    function?: Array<[Ctor<BaseFunction>, IFunctionNames]>;
    /**
     * The formula calculation quantity interval for waiting for the main thread message in the worker. Each time the formula calculates the `intervalCount` quantity, it will receive a main thread message to support stopping the calculation. Default is 500
     */
    intervalCount?: number;
    /**
     * The maximum number of dependency tree nodes to execute before yielding and
     * clearing transient reference caches. Lower values reduce peak working-set
     * size at the cost of extra scheduler overhead.
     */
    batchExecutionCount?: number;
}
export declare const defaultPluginConfig: IUniverEngineFormulaConfig;
