/**
 * Copyright 2023-present DreamNum Co., Ltd.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import type { Ctor } from '@univerjs/core';
import type { IFunctionNames } from '../basics/function';
import type { BaseFunction } from '../functions/base-function';

export const ENGINE_FORMULA_PLUGIN_CONFIG_KEY = 'engine-formula.config';

export const DEFAULT_CYCLE_REFERENCE_COUNT = 1;

export const ENGINE_FORMULA_CYCLE_REFERENCE_COUNT = 'CYCLE_REFERENCE_COUNT';

export const configSymbol = Symbol(ENGINE_FORMULA_PLUGIN_CONFIG_KEY);

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
     * Optional sampling interval for execution tracing. When > 0, records
     * progress/memory samples every N formulas.
     */
    traceIntervalCount?: number;

    /**
     * When traceIntervalCount is enabled, emit per-sample progress markers to stderr.
     */
    traceToStderr?: boolean;

}

export const defaultPluginConfig: IUniverEngineFormulaConfig = {};
