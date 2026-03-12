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

import type { IUniverEngineFormulaConfig } from '@univerjs/engine-formula';

function envFlag(name: string): boolean {
    return process.env[name] === '1';
}

function envPositiveInt(name: string): number | undefined {
    const raw = process.env[name];
    if (!raw) return undefined;
    const n = Number(raw);
    if (!Number.isFinite(n) || n <= 0) return undefined;
    return Math.floor(n);
}

export function getFormulaReplayConfigFromEnv(): Pick<
    IUniverEngineFormulaConfig,
    | 'disableTypedArrayOptimization'
    | 'disableClearIntermediate'
    | 'disableScalarFastPath'
    | 'disableFillDownSharing'
    | 'disableFormulaHashCache'
    | 'traceIntervalCount'
    | 'traceToStderr'
> {
    const traceIntervalCount = envPositiveInt('FORMULA_TRACE_INTERVAL');
    return {
        disableTypedArrayOptimization: envFlag('DISABLE_TYPED_ARRAY'),
        disableClearIntermediate: envFlag('DISABLE_CLEAR_INTERMEDIATE'),
        disableScalarFastPath: envFlag('DISABLE_SCALAR_FAST_PATH'),
        disableFillDownSharing: envFlag('DISABLE_FILL_DOWN_SHARING'),
        disableFormulaHashCache: envFlag('DISABLE_FORMULA_HASH_CACHE'),
        traceIntervalCount,
        traceToStderr: envFlag('FORMULA_TRACE_STDERR') || traceIntervalCount !== undefined,
    };
}
