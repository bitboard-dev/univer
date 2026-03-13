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

import type { Ctor, IWorkbookData, Worksheet } from '@univerjs/core';
import type { BaseFunction, IFunctionNames } from '@univerjs/engine-formula';
import type { FFormula } from '@univerjs/engine-formula/facade';
import { ICommandService, IConfigService } from '@univerjs/core';
import {
    ENGINE_FORMULA_PLUGIN_CONFIG_KEY,
    functionArray,
    functionCompatibility,
    functionDatabase,
    functionDate,
    functionEngineering,
    functionFinancial,
    functionInformation,
    functionLogical,
    functionLookup,
    functionMath,
    functionMeta,
    functionStatistical,
    functionText,
    functionWeb,
    ICalculateFormulaService,
    IFormulaCurrentConfigService,
    IFormulaRuntimeService,
    IFunctionService,
    SetArrayFormulaDataMutation,
    SetFormulaCalculationNotificationMutation,
    SetFormulaCalculationResultMutation,
    SetFormulaCalculationStartMutation,
    SetFormulaCalculationStopMutation,
} from '@univerjs/engine-formula';
import { describe, expect, it } from 'vitest';
import { SetRangeValuesMutation } from '../../../commands/mutations/set-range-values.mutation';
import { buildDataSheet, DATA_SHEET_ID } from './benchmark/data-generator';
import { FORMULA_GROUPS } from './benchmark/formula-groups';
import { createFunctionTestBed } from './create-function-test-bed';
import { profileCalculation } from './formula-profiler';
import { getFormulaReplayConfigFromEnv } from './replay-config';

import '@univerjs/engine-formula/facade';

const unitId = 'test';
const RUN_ENV = process.env.RUN_COMPOUND_PERF === '1' ? 'RUN_COMPOUND_PERF' : 'RUN_FORMULA_BENCHMARK';
const DEFAULT_ROWS = 10000;

const allFunctionRegistrations = [
    ...functionMath,
    ...functionStatistical,
    ...functionLogical,
    ...functionLookup,
    ...functionMeta,
    ...functionDate,
    ...functionText,
    ...functionInformation,
    ...functionFinancial,
    ...functionEngineering,
    ...functionDatabase,
    ...functionCompatibility,
    ...functionArray,
    ...functionWeb,
];

interface IGroupResult {
    id: string;
    name: string;
    formulas: number;
    rows: number;
    wallMs: number;
    peakHeapMB: number;
}

function buildWorkbook(rows: number, groupIds: string[]): IWorkbookData {
    const sheets: IWorkbookData['sheets'] = {};
    const dataSheet = buildDataSheet({ rows });
    sheets[DATA_SHEET_ID] = dataSheet;

    for (const group of FORMULA_GROUPS) {
        if (groupIds.length > 0 && !groupIds.includes(group.id)) continue;
        sheets[group.sheetId] = group.build(rows);
    }

    return {
        id: unitId,
        appVersion: '1.0.0',
        name: 'formula-benchmark',
        locale: 'enUS' as any,
        styles: {},
        sheetOrder: [DATA_SHEET_ID, ...Object.keys(sheets).filter((k) => k !== DATA_SHEET_ID)],
        sheets,
        resources: [],
    };
}

function setupTestBed(workbookData: IWorkbookData) {
    const testBed = createFunctionTestBed(workbookData);
    const get = testBed.get;
    const formulaEngine = testBed.api.getFormula() as FFormula;
    const calcService = get(ICalculateFormulaService);
    const commandService = get(ICommandService);

    commandService.registerCommand(SetFormulaCalculationStartMutation);
    commandService.registerCommand(SetFormulaCalculationStopMutation);
    commandService.registerCommand(SetFormulaCalculationResultMutation);
    commandService.registerCommand(SetFormulaCalculationNotificationMutation);
    commandService.registerCommand(SetArrayFormulaDataMutation);
    commandService.registerCommand(SetRangeValuesMutation);

    const configService = get(IConfigService);
    configService.setConfig(ENGINE_FORMULA_PLUGIN_CONFIG_KEY, {
        ...getFormulaReplayConfigFromEnv(),
    });

    const functionService = get(IFunctionService);
    const formulaCurrentConfigService = get(IFormulaCurrentConfigService);
    const formulaRuntimeService = get(IFormulaRuntimeService);

    formulaCurrentConfigService.load({
        formulaData: {},
        arrayFormulaCellData: {},
        arrayFormulaRange: {},
        forceCalculate: false,
        dirtyRanges: [],
        dirtyNameMap: {},
        dirtyDefinedNameMap: {},
        dirtyUnitFeatureMap: {},
        dirtyUnitOtherFormulaMap: {},
        excludedCell: {},
        allUnitData: {
            [testBed.unitId]: testBed.sheetData,
        },
    });

    const sheetItem = testBed.sheetData[testBed.sheetId];
    formulaRuntimeService.setCurrent(
        0,
        0,
        sheetItem.rowCount,
        sheetItem.columnCount,
        testBed.sheetId,
        testBed.unitId
    );

    const functions = allFunctionRegistrations.map((registerObject) => {
        const Func = registerObject[0] as Ctor<BaseFunction>;
        const name = registerObject[1] as IFunctionNames;
        return new Func(name);
    });
    functionService.registerExecutors(...functions);

    const getCellValue = (sheetId: string, row: number, column: number) => {
        const worksheet = testBed.sheet.getSheetBySheetId(sheetId) as Worksheet;
        return worksheet.getCellRaw(row, column)?.v;
    };

    return { testBed, formulaEngine, calcService, getCellValue };
}

function printSummary(results: IGroupResult[]) {
    const totalFormulas = results.reduce((s, r) => s + r.formulas, 0);
    const totalCells = results.reduce((s, r) => s + r.formulas * r.rows, 0);
    const totalMs = results.reduce((s, r) => s + r.wallMs, 0);
    const peakHeap = Math.max(...results.map((r) => r.peakHeapMB), 0);
    const rows = results[0]?.rows ?? 0;

    console.warn('\n========================================');
    console.warn(`Formula Benchmark Summary (N=${rows})`);
    console.warn('========================================');
    for (const r of results) {
        const pad = (s: string, n: number) => s.padEnd(n);
        const cells = r.formulas * r.rows;
        console.warn(
            `  ${pad(`${r.id}:`, 20)} ${String(r.formulas).padStart(4)} formulas  ${String(cells).padStart(9)} cells  ${String(r.wallMs).padStart(8)}ms  ${String(r.peakHeapMB).padStart(6)}MB`
        );
    }
    console.warn('----------------------------------------');
    console.warn(
        `  ${'TOTAL:'.padEnd(20)} ${String(totalFormulas).padStart(4)} formulas  ${String(totalCells).padStart(9)} cells  ${String(totalMs).padStart(8)}ms  ${String(peakHeap).padStart(6)}MB peak`
    );
    console.warn('========================================\n');
}

// ---------------------------------------------------------------------------
// Correctness: small-N test per group (always runs)
// ---------------------------------------------------------------------------
describe('formula benchmark correctness', () => {
    const rows = 120;

    it.each(FORMULA_GROUPS.map((g) => [g.id, g]))(
        'correctness: %s',
        async (_id, group) => {
            const g = group as typeof FORMULA_GROUPS[number];
            const workbook = buildWorkbook(rows, [g.id]);
            const { testBed, formulaEngine, getCellValue } = setupTestBed(workbook);

            try {
                const endPromise = formulaEngine.onCalculationEnd(60_000);
                formulaEngine.executeCalculation();
                await endPromise;

                const errors: string[] = [];
                for (let r = 1; r <= Math.min(rows, 10); r++) {
                    for (let c = 0; c < g.formulas.length; c++) {
                        const val = getCellValue(g.sheetId, r, c);
                        if (val === '#NAME?' || val === '#REF!') {
                            errors.push(`${g.formulas[c]} row ${r}: ${val}`);
                        }
                    }
                }

                if (errors.length > 0) {
                    console.warn(`[${g.id}] Errors (${errors.length}):`);
                    for (const e of errors) console.warn(`  ${e}`);
                }
                expect(errors.length, `${errors.length} formula errors in ${g.id}`).toBe(0);
            } finally {
                testBed.univer.dispose();
            }
        },
        60000
    );
});

// ---------------------------------------------------------------------------
// Performance: heavy benchmark (gated)
// ---------------------------------------------------------------------------
describe('formula benchmark performance', () => {
    const runHeavy = process.env[RUN_ENV] === '1';
    const heavyIt = runHeavy ? it : it.skip;
    const rows = Number(process.env.BENCHMARK_ROWS ?? DEFAULT_ROWS);
    const requestedGroups = (process.env.BENCHMARK_GROUPS ?? 'all')
        .split(',')
        .map((s) => s.trim())
        .filter((s) => s && s !== 'all');

    const activeGroups = FORMULA_GROUPS.filter(
        (g) => requestedGroups.length === 0 || requestedGroups.includes(g.id)
    );

    const results: IGroupResult[] = [];

    heavyIt.each(activeGroups.map((g) => [g.id, g]))(
        'benchmarks %s',
        async (_id, group) => {
            const g = group as typeof FORMULA_GROUPS[number];
            const workbook = buildWorkbook(rows, [g.id]);
            const { testBed, formulaEngine, calcService, getCellValue } = setupTestBed(workbook);

            try {
                const cpu = !!process.env.FORMULA_CPU_PROF;
                const heap = !!process.env.FORMULA_HEAP_PROF;

                const formulaCells = g.formulas.length * rows;
                console.warn(`\n[${g.id}] ${g.formulas.length} formulas x ${rows} rows = ${formulaCells.toLocaleString()} cells`);

                const heapBefore = process.memoryUsage().heapUsed;
                const t0 = performance.now();

                const result = await profileCalculation(formulaEngine, calcService, {
                    label: `${g.id}-${rows}`,
                    timeoutMs: 600_000,
                    cpu,
                    heap,
                });

                const wallMs = Math.round(performance.now() - t0);
                const heapAfter = process.memoryUsage().heapUsed;
                const peakMB = result.peakHeapMB || Math.round((heapAfter - heapBefore) / 1048576);
                const traceWall = result.wallMs || wallMs;

                console.warn(`  wall: ${wallMs}ms | trace wall: ${traceWall}ms | heap delta: ${Math.round((heapAfter - heapBefore) / 1048576)}MB`);

                results.push({
                    id: g.id,
                    name: g.name,
                    formulas: g.formulas.length,
                    rows,
                    wallMs,
                    peakHeapMB: peakMB,
                });

                const sample = getCellValue(g.sheetId, 1, 0);
                expect(sample !== undefined && sample !== null, `${g.id} first cell should have a value`).toBe(true);
            } finally {
                testBed.univer.dispose();
            }
        },
        600000
    );

    heavyIt('prints summary', () => {
        printSummary(results);
        expect(results.length).toBeGreaterThan(0);
    });
});
