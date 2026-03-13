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
    /**
     * Replay/testing toggle: when true, disables typed-array-backed fast paths.
     */
    disableTypedArrayOptimization?: boolean;
    /**
     * Replay/testing toggle: when true, keeps intermediate AST values alive.
     */
    disableClearIntermediate?: boolean;
    /**
     * Replay/testing toggle: when true, disables scalar *IFS fast paths.
     */
    disableScalarFastPath?: boolean;
    /**
     * Replay/testing toggle: when true, disables fill-down AST sharing.
     */
    disableFillDownSharing?: boolean;
    /**
     * Replay/testing toggle: when true, disables SUMIF/COUNTIF/COUNTIFS hash caches.
     */
    disableFormulaHashCache?: boolean;
    /**
     * Replay/testing toggle: when true, disables indexed equality lookup in
     * BaseFunction.equalSearch and uses scan-based fallback instead.
     */
    disableEqualSearchIndex?: boolean;
    /**
     * Replay/testing toggle: when true, disables RANK.EQ numeric-data fast path
     * and uses iterator-based extraction for parity checks.
     */
    disableRankEqNumericFastPath?: boolean;
    /**
     * Replay/testing toggle: when true, disables sorted numeric cache lookups
     * used by LARGE and PERCENTRANK.INC.
     */
    disableSortedNumericCache?: boolean;
    /**
     * Replay/testing toggle: when true, disables RANK.EQ index-map lookup and
     * falls back to indexOf over sorted arrays.
     */
    disableRankEqIndexMap?: boolean;
    /**
     * Optional sampling interval for execution tracing. When > 0, records
     * progress/memory samples every N formulas.
     */
    traceIntervalCount?: number;
    /**
     * When traceIntervalCount is enabled, emit per-sample progress markers to stderr.
     */
    traceToStderr?: boolean;
}
export declare const defaultPluginConfig: IUniverEngineFormulaConfig;
