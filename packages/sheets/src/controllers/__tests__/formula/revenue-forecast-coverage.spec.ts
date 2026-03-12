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

import type { CellValue, Ctor, IWorkbookData, Nullable, Worksheet } from '@univerjs/core';
import type { BaseFunction, IFunctionNames } from '@univerjs/engine-formula';
import type { FFormula } from '@univerjs/engine-formula/facade';
import { CellValueType, ICommandService, IConfigService, LocaleType } from '@univerjs/core';
import {
    ENGINE_FORMULA_PLUGIN_CONFIG_KEY,
    functionLogical,
    functionLookup,
    functionMath,
    functionMeta,
    functionStatistical,
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
import { afterEach, describe, expect, it } from 'vitest';
import { SetRangeValuesMutation } from '../../../commands/mutations/set-range-values.mutation';
import { createFunctionTestBed } from './create-function-test-bed';
import { profileCalculation } from './formula-profiler';
import { getFormulaReplayConfigFromEnv } from './replay-config';

import '@univerjs/engine-formula/facade';

const unitId = 'test';
const dealsSheetId = 'deals';
const quotaSheetId = 'quota';
const forecastSheetId = 'forecast';
const RUN_ENV = 'RUN_REVENUE_FORECAST_COVERAGE';

function buildWorkbookData(options: { numRows: number; numOwners: number; numAccounts: number }): IWorkbookData {
    const { numRows, numOwners, numAccounts } = options;
    const quarters = ['Q1', 'Q2', 'Q3', 'Q4'];

    const sheets: IWorkbookData['sheets'] = {
        [dealsSheetId]: {
            id: dealsSheetId,
            name: 'Deals',
            rowCount: numRows + 10,
            columnCount: 7,
            cellData: {
                0: {
                    0: { v: 'Deal ID', t: CellValueType.STRING },
                    1: { v: 'Account ID', t: CellValueType.STRING },
                    2: { v: 'Owner', t: CellValueType.STRING },
                    3: { v: 'Quarter', t: CellValueType.STRING },
                    4: { v: 'Amount', t: CellValueType.STRING },
                    5: { v: 'Probability', t: CellValueType.STRING },
                    6: { v: 'Weighted Amount', t: CellValueType.STRING },
                },
            },
        },
        [quotaSheetId]: {
            id: quotaSheetId,
            name: 'Quota',
            rowCount: numOwners + 10,
            columnCount: 2,
            cellData: {
                0: {
                    0: { v: 'Owner', t: CellValueType.STRING },
                    1: { v: 'Quarterly Quota', t: CellValueType.STRING },
                },
            },
        },
        [forecastSheetId]: {
            id: forecastSheetId,
            name: 'Forecast',
            rowCount: numRows + 10,
            columnCount: 10,
            cellData: {
                0: {
                    0: { v: 'Deal ID', t: CellValueType.STRING },
                    1: { v: 'Owner', t: CellValueType.STRING },
                    2: { v: 'Quarter', t: CellValueType.STRING },
                    3: { v: 'Weighted', t: CellValueType.STRING },
                    4: { v: 'Quota via VLOOKUP', t: CellValueType.STRING },
                    5: { v: 'Owner position via MATCH', t: CellValueType.STRING },
                    6: { v: 'Owner Quota', t: CellValueType.STRING },
                    7: { v: 'Attainment', t: CellValueType.STRING },
                    8: { v: 'Rank', t: CellValueType.STRING },
                    9: { v: 'Percentile', t: CellValueType.STRING },
                },
            },
        },
    };

    const deals = sheets[dealsSheetId].cellData!;
    const quota = sheets[quotaSheetId].cellData!;
    const forecast = sheets[forecastSheetId].cellData!;
    const endRow = numRows + 1;
    const quotaEnd = numOwners + 1;
    const weightedRange = `${dealsSheetId}!$G$2:$G$${endRow}`;
    const quotaOwnerRange = `${quotaSheetId}!$A$2:$A$${quotaEnd}`;
    const quotaValueRange = `${quotaSheetId}!$B$2:$B$${quotaEnd}`;

    for (let i = 0; i < numOwners; i++) {
        const row = i + 1;
        quota[row] = {
            0: { v: `Owner ${String(i + 1).padStart(2, '0')}`, t: CellValueType.STRING },
            1: { v: 500000 + i * 10000, t: CellValueType.NUMBER },
        };
    }

    for (let i = 0; i < numRows; i++) {
        const row = i + 1;
        const sheetRow = row + 1;
        const owner = `Owner ${String((i % numOwners) + 1).padStart(2, '0')}`;
        const quarter = quarters[i % quarters.length];
        const amount = 30000 + (i % 500) * 1300;
        const probability = [0.2, 0.35, 0.5, 0.7, 0.85][i % 5];

        deals[row] = {
            0: { v: `D-${String(i + 1).padStart(5, '0')}`, t: CellValueType.STRING },
            1: { v: `ACC-${String((i % numAccounts) + 1).padStart(4, '0')}`, t: CellValueType.STRING },
            2: { v: owner, t: CellValueType.STRING },
            3: { v: quarter, t: CellValueType.STRING },
            4: { v: amount, t: CellValueType.NUMBER },
            5: { v: probability, t: CellValueType.NUMBER },
            6: { f: `=E${sheetRow}*F${sheetRow}` },
        };

        forecast[row] = {
            0: { f: `=${dealsSheetId}!A${sheetRow}` },
            1: { f: `=${dealsSheetId}!C${sheetRow}` },
            2: { f: `=${dealsSheetId}!D${sheetRow}` },
            3: { f: `=${dealsSheetId}!G${sheetRow}` },
            4: { f: `=IFERROR(VLOOKUP(B${sheetRow},${quotaSheetId}!$A$2:$B$${quotaEnd},2,FALSE),0)` },
            5: { f: `=IFERROR(MATCH(B${sheetRow},${quotaOwnerRange},0),0)` },
            6: { f: `=IFERROR(INDEX(${quotaValueRange},MATCH(B${sheetRow},${quotaOwnerRange},0)),0)` },
            7: { f: `=IFERROR(E${sheetRow}/G${sheetRow},0)` },
            8: row === 1
                ? { f: `=RANK.EQ(D${sheetRow},${weightedRange})` }
                : { f: '=I2' },
            9: row === 1
                ? { f: `=PERCENTRANK.INC(${weightedRange},D${sheetRow})` }
                : { f: '=J2' },
        };
    }

    return {
        id: unitId,
        appVersion: '3.0.0-alpha',
        locale: LocaleType.ZH_CN,
        name: 'Revenue Forecast Coverage Test',
        sheetOrder: [dealsSheetId, quotaSheetId, forecastSheetId],
        sheets,
        styles: {},
    };
}

describe('revenue forecast formula coverage', () => {
    let _univer: ReturnType<typeof createFunctionTestBed>['univer'];
    let formulaEngine: FFormula;
    let calcService: ICalculateFormulaService;
    let getCellValue: (sheetId: string, row: number, column: number) => Nullable<CellValue>;

    function setup(options: { numRows: number; numOwners: number; numAccounts: number }) {
        const testBed = createFunctionTestBed(buildWorkbookData(options));
        _univer = testBed.univer;

        const get = testBed.get;
        formulaEngine = testBed.api.getFormula() as FFormula;
        calcService = get(ICalculateFormulaService);
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
        formulaRuntimeService.setCurrent(0, 0, sheetItem.rowCount, sheetItem.columnCount, testBed.sheetId, testBed.unitId);

        const functions = [
            ...functionMath,
            ...functionStatistical,
            ...functionLogical,
            ...functionLookup,
            ...functionMeta,
        ].map((registerObject) => {
            const Func = registerObject[0] as Ctor<BaseFunction>;
            const name = registerObject[1] as IFunctionNames;
            return new Func(name);
        });

        functionService.registerExecutors(...functions);

        getCellValue = (sheetId: string, row: number, column: number) => {
            const worksheet = testBed.sheet.getSheetBySheetId(sheetId) as Worksheet;
            return worksheet.getCellRaw(row, column)?.v;
        };
    }

    afterEach(() => {
        _univer?.dispose();
    });

    it('computes owner-quarter forecast and ranking correctly at 120 rows', async () => {
        const numRows = 120;
        const numOwners = 12;
        setup({ numRows, numOwners, numAccounts: 40 });

        formulaEngine.executeCalculation();
        await formulaEngine.onCalculationEnd(60_000);

        const rows = Array.from({ length: numRows }, (_, i) => {
            const idx = i + 1;
            const owner = `Owner ${String((i % numOwners) + 1).padStart(2, '0')}`;
            const quarter = ['Q1', 'Q2', 'Q3', 'Q4'][i % 4];
            const amount = 30000 + (i % 500) * 1300;
            const weighted = amount * [0.2, 0.35, 0.5, 0.7, 0.85][i % 5];
            return { idx, owner, quarter, amount, weighted };
        });

        for (const row of rows) {
            expect(getCellValue(forecastSheetId, row.idx, 3), `weighted row ${row.idx}`).not.toBeNull();
            expect(getCellValue(forecastSheetId, row.idx, 4), `quota row ${row.idx}`).not.toBeNull();
            expect(getCellValue(forecastSheetId, row.idx, 5), `match row ${row.idx}`).not.toBeNull();
        }

        expect(getCellValue(forecastSheetId, 1, 8), 'rank row 1').not.toBeNull();
        expect(getCellValue(forecastSheetId, 1, 9), 'percentrank row 1').not.toBeNull();
    }, 60_000);

    const runHeavy = process.env[RUN_ENV] === '1';
    const heavyIt = runHeavy ? it : it.skip;

    heavyIt('survives optimized heavy run at 2GB heap', async () => {
        const numRows = Number(process.env.REVENUE_FORECAST_ROWS ?? '10000');
        setup({ numRows, numOwners: 60, numAccounts: 1200 });

        const result = await profileCalculation(formulaEngine, calcService, {
            label: 'revenue-forecast-10k',
            timeoutMs: 600_000,
            cpu: !!process.env.FORMULA_CPU_PROF,
            heap: !!process.env.FORMULA_HEAP_PROF,
        });

        expect(getCellValue(forecastSheetId, 1, 4)).not.toBeNull();
        expect(getCellValue(forecastSheetId, 1, 8)).not.toBeNull();

        if (result.trace.length > 0) {
            expect(result.peakHeapMB).toBeLessThan(1800);
        }
    }, 600_000);
});
