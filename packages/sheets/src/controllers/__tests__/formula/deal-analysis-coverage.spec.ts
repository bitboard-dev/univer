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
const opportunitiesSheetId = 'opportunities';
const accountsSheetId = 'accounts';
const dealAnalysisSheetId = 'deal-analysis';

const RUN_ENV = process.env.RUN_COMPOUND_PERF === '1' ? 'RUN_COMPOUND_PERF' : 'RUN_DEAL_ANALYSIS_COVERAGE';

const stages = ['Prospecting', 'Qualification', 'Discovery', 'Proposal', 'Negotiation', 'Closed Won', 'Closed Lost'];
const industries = ['Technology', 'Financial Services', 'Healthcare', 'Retail', 'Manufacturing', 'Media', 'Education', 'Government'];

function buildWorkbookData(options: { numOpps: number; numReps: number; numAccounts: number }): IWorkbookData {
    const { numOpps, numReps, numAccounts } = options;

    const sheets: IWorkbookData['sheets'] = {
        [opportunitiesSheetId]: {
            id: opportunitiesSheetId,
            name: 'Opportunities',
            rowCount: numOpps + 10,
            columnCount: 8,
            cellData: {
                0: {
                    0: { v: 'Opp ID', t: CellValueType.STRING },
                    1: { v: 'Account ID', t: CellValueType.STRING },
                    2: { v: 'Rep', t: CellValueType.STRING },
                    3: { v: 'Region', t: CellValueType.STRING },
                    4: { v: 'Stage', t: CellValueType.STRING },
                    5: { v: 'Amount', t: CellValueType.STRING },
                    6: { v: 'Probability', t: CellValueType.STRING },
                },
            },
        },
        [accountsSheetId]: {
            id: accountsSheetId,
            name: 'Accounts',
            rowCount: numAccounts + 10,
            columnCount: 4,
            cellData: {
                0: {
                    0: { v: 'Account ID', t: CellValueType.STRING },
                    1: { v: 'Account Name', t: CellValueType.STRING },
                    2: { v: 'Industry', t: CellValueType.STRING },
                    3: { v: 'Employees', t: CellValueType.STRING },
                },
            },
        },
        [dealAnalysisSheetId]: {
            id: dealAnalysisSheetId,
            name: 'Deal Analysis',
            rowCount: numOpps + 10,
            columnCount: 14,
            cellData: {
                0: {
                    0: { v: 'Opp ID', t: CellValueType.STRING },
                    1: { v: 'Account ID', t: CellValueType.STRING },
                    2: { v: 'Rep', t: CellValueType.STRING },
                    3: { v: 'Amount', t: CellValueType.STRING },
                    4: { v: 'Stage', t: CellValueType.STRING },
                    5: { v: 'Industry (VLOOKUP)', t: CellValueType.STRING },
                    6: { v: 'Employees (INDEX/MATCH)', t: CellValueType.STRING },
                    7: { v: 'Rep Stage Pipeline (SUMIFS)', t: CellValueType.STRING },
                    8: { v: 'Rep Stage Avg Deal (AVERAGEIFS)', t: CellValueType.STRING },
                    9: { v: 'Rep Max Deal (MAXIFS)', t: CellValueType.STRING },
                    10: { v: 'Rep Min Deal (MINIFS)', t: CellValueType.STRING },
                    11: { v: 'Deal Rank (RANK.EQ)', t: CellValueType.STRING },
                    12: { v: 'Top 5 Deal (LARGE)', t: CellValueType.STRING },
                    13: { v: 'Percentile (PERCENTRANK.INC)', t: CellValueType.STRING },
                },
            },
        },
    };

    const opportunities = sheets[opportunitiesSheetId].cellData!;
    const accounts = sheets[accountsSheetId].cellData!;
    const dealAnalysis = sheets[dealAnalysisSheetId].cellData!;

    for (let i = 0; i < numAccounts; i++) {
        const row = i + 1;
        accounts[row] = {
            0: { v: `ACC-${String(i + 1).padStart(4, '0')}`, t: CellValueType.STRING },
            1: { v: `Company ${i + 1}`, t: CellValueType.STRING },
            2: { v: industries[i % industries.length], t: CellValueType.STRING },
            3: { v: 50 * (1 + (i % 20)), t: CellValueType.NUMBER },
        };
    }

    const endRow = numOpps + 1;
    const acctEndRow = numAccounts + 1;
    const _oppAcctId = `Opportunities!$B$2:$B$${endRow}`;
    const oppRep = `Opportunities!$C$2:$C$${endRow}`;
    const oppStage = `Opportunities!$E$2:$E$${endRow}`;
    const oppAmt = `Opportunities!$F$2:$F$${endRow}`;

    const acctRange = `Accounts!$A$2:$D$${acctEndRow}`;
    const acctIdCol = `Accounts!$A$2:$A$${acctEndRow}`;
    const acctEmployeesCol = `Accounts!$D$2:$D$${acctEndRow}`;

    for (let i = 0; i < numOpps; i++) {
        const row = i + 1;
        const sheetRow = row + 1;
        const rep = `Rep ${String((i % numReps) + 1).padStart(2, '0')}`;
        const account = `ACC-${String((i % numAccounts) + 1).padStart(4, '0')}`;
        const stage = stages[i % stages.length];
        const amount = 50000 + (i % 200) * 2500;

        opportunities[row] = {
            0: { v: `OPP-${String(i + 1).padStart(5, '0')}`, t: CellValueType.STRING },
            1: { v: account, t: CellValueType.STRING },
            2: { v: rep, t: CellValueType.STRING },
            3: { v: ['West', 'East', 'Central', 'EMEA', 'APAC'][i % 5], t: CellValueType.STRING },
            4: { v: stage, t: CellValueType.STRING },
            5: { v: amount, t: CellValueType.NUMBER },
            6: { v: 0.5, t: CellValueType.NUMBER },
        };

        dealAnalysis[row] = {
            0: { f: `=Opportunities!A${sheetRow}` },
            1: { f: `=Opportunities!B${sheetRow}` },
            2: { f: `=Opportunities!C${sheetRow}` },
            3: { f: `=Opportunities!F${sheetRow}` },
            4: { f: `=Opportunities!E${sheetRow}` },
            5: { f: `=IFERROR(VLOOKUP(B${sheetRow},${acctRange},3,FALSE),"")` },
            6: { f: `=IFERROR(INDEX(${acctEmployeesCol},MATCH(B${sheetRow},${acctIdCol},0)),0)` },
            7: { f: `=SUMIFS(${oppAmt},${oppRep},C${sheetRow},${oppStage},E${sheetRow})` },
            8: { f: `=IFERROR(AVERAGEIFS(${oppAmt},${oppRep},C${sheetRow},${oppStage},E${sheetRow}),0)` },
            9: { f: `=MAXIFS(${oppAmt},${oppRep},C${sheetRow})` },
            10: { f: `=MINIFS(${oppAmt},${oppRep},C${sheetRow})` },
            11: { f: `=RANK.EQ(D${sheetRow},${oppAmt})` },
            12: { f: `=LARGE(${oppAmt},5)` },
            13: { f: `=PERCENTRANK.INC(${oppAmt},D${sheetRow})` },
        };
    }

    return {
        id: unitId,
        appVersion: '3.0.0-alpha',
        locale: LocaleType.ZH_CN,
        name: 'Deal Analysis Coverage Test',
        sheetOrder: [opportunitiesSheetId, accountsSheetId, dealAnalysisSheetId],
        sheets,
        styles: {},
    };
}

describe('deal analysis formula coverage', () => {
    let _univer: ReturnType<typeof createFunctionTestBed>['univer'];
    let formulaEngine: FFormula;
    let getCellValue: (sheetId: string, row: number, column: number) => Nullable<CellValue>;
    let calcService: ICalculateFormulaService;

    function setup(options: { numOpps: number; numReps: number; numAccounts: number }) {
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
        formulaRuntimeService.setCurrent(
            0,
            0,
            sheetItem.rowCount,
            sheetItem.columnCount,
            testBed.sheetId,
            testBed.unitId
        );

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

    it('computes all formula types correctly at 120 rows', async () => {
        const numOpps = 120;
        const numReps = 12;
        const numAccounts = 40;
        setup({ numOpps, numReps, numAccounts });

        const endPromise = formulaEngine.onCalculationEnd(60_000);
        formulaEngine.executeCalculation();
        await endPromise;

        const opps: { accountId: string; rep: string; stage: string; amount: number }[] = [];
        for (let i = 0; i < numOpps; i++) {
            opps.push({
                accountId: `ACC-${String((i % numAccounts) + 1).padStart(4, '0')}`,
                rep: `Rep ${String((i % numReps) + 1).padStart(2, '0')}`,
                stage: stages[i % stages.length],
                amount: 50000 + (i % 200) * 2500,
            });
        }

        const acctMap = new Map<string, { industry: string; employees: number }>();
        for (let i = 0; i < numAccounts; i++) {
            acctMap.set(`ACC-${String(i + 1).padStart(4, '0')}`, {
                industry: industries[i % industries.length],
                employees: 50 * (1 + (i % 20)),
            });
        }

        const allAmounts = opps.map((o) => o.amount);
        const sortedDesc = [...allAmounts].sort((a, b) => b - a);
        const sortedAsc = [...allAmounts].sort((a, b) => a - b);
        const expectedLarge5 = sortedDesc[4];

        function percentrankInc(arr: number[], x: number): number {
            const n = arr.length;
            for (let j = 0; j < n; j++) {
                if (x === arr[j]) {
                    return Math.floor((j / (n - 1)) * 1000) / 1000;
                }
                if (x > arr[j] && j + 1 < n && x < arr[j + 1]) {
                    const raw = (j + (x - arr[j]) / (arr[j + 1] - arr[j])) / (n - 1);
                    return Math.floor(raw * 1000) / 1000;
                }
            }
            return -1;
        }

        for (let i = 0; i < numOpps; i++) {
            const row = i + 1;
            const opp = opps[i];
            const acct = acctMap.get(opp.accountId)!;
            const label = `row ${row} (opp ${i})`;

            const actualIndustry = getCellValue(dealAnalysisSheetId, row, 5);
            expect(actualIndustry, `VLOOKUP ${label}`).toBe(acct.industry);

            const actualEmployees = getCellValue(dealAnalysisSheetId, row, 6);
            expect(actualEmployees, `INDEX/MATCH ${label}`).toBe(acct.employees);

            const repStageOpps = opps.filter((o) => o.rep === opp.rep && o.stage === opp.stage);
            const expectedSumifs = repStageOpps.reduce((s, o) => s + o.amount, 0);
            const actualSumifs = getCellValue(dealAnalysisSheetId, row, 7);
            expect(actualSumifs, `SUMIFS ${label}`).toBe(expectedSumifs);

            const expectedAvgifs = repStageOpps.length > 0
                ? repStageOpps.reduce((s, o) => s + o.amount, 0) / repStageOpps.length
                : 0;
            const actualAvgifs = getCellValue(dealAnalysisSheetId, row, 8);
            expect(actualAvgifs, `AVERAGEIFS ${label}`).toBeCloseTo(expectedAvgifs, 6);

            const repOpps = opps.filter((o) => o.rep === opp.rep);
            const expectedMaxifs = Math.max(...repOpps.map((o) => o.amount));
            const actualMaxifs = getCellValue(dealAnalysisSheetId, row, 9);
            expect(actualMaxifs, `MAXIFS ${label}`).toBe(expectedMaxifs);

            const expectedMinifs = Math.min(...repOpps.map((o) => o.amount));
            const actualMinifs = getCellValue(dealAnalysisSheetId, row, 10);
            expect(actualMinifs, `MINIFS ${label}`).toBe(expectedMinifs);

            const expectedRank = sortedDesc.indexOf(opp.amount) + 1;
            const actualRank = getCellValue(dealAnalysisSheetId, row, 11);
            expect(actualRank, `RANK.EQ ${label}`).toBe(expectedRank);

            const actualLarge = getCellValue(dealAnalysisSheetId, row, 12);
            expect(actualLarge, `LARGE ${label}`).toBe(expectedLarge5);

            const expectedPctRank = percentrankInc(sortedAsc, opp.amount);
            const actualPctRank = getCellValue(dealAnalysisSheetId, row, 13);
            expect(actualPctRank, `PERCENTRANK.INC ${label}`).toBeCloseTo(expectedPctRank, 3);
        }
    }, 60000);

    const runHeavy = process.env[RUN_ENV] === '1';
    const heavyIt = runHeavy ? it : it.skip;

    heavyIt('survives 10k rows with uncached formula mix at 2GB heap', async () => {
        const numOpps = Number(process.env.DEAL_ANALYSIS_OPPS ?? '10000');
        setup({ numOpps, numReps: 50, numAccounts: 800 });

        const result = await profileCalculation(formulaEngine, calcService, {
            label: 'deal-analysis-10k',
            timeoutMs: 600_000,
            cpu: !!process.env.FORMULA_CPU_PROF,
            heap: !!process.env.FORMULA_HEAP_PROF,
        });

        const vlookup = getCellValue(dealAnalysisSheetId, 1, 5);
        expect(vlookup).toBeTruthy();

        const sumifs = getCellValue(dealAnalysisSheetId, 1, 7);
        expect(typeof sumifs).toBe('number');
        expect(sumifs).toBeGreaterThan(0);

        const rank = getCellValue(dealAnalysisSheetId, 1, 11);
        expect(typeof rank).toBe('number');

        if (result.trace.length > 0) {
            expect(result.peakHeapMB).toBeLessThan(1800);
        }
    }, 600000);
});
