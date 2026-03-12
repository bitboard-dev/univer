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

import type { CellValue, Ctor, Injector, IWorkbookData, Nullable, Worksheet } from '@univerjs/core';
import type { BaseFunction, IFunctionNames } from '@univerjs/engine-formula';
import type { FFormula } from '@univerjs/engine-formula/facade';
import { CellValueType, ICommandService, IConfigService, LocaleType } from '@univerjs/core';
import {
    ENGINE_FORMULA_PLUGIN_CONFIG_KEY,
    functionLogical,
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
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { SetRangeValuesMutation } from '../../../commands/mutations/set-range-values.mutation';
import { createFunctionTestBed } from './create-function-test-bed';
import { profileCalculation } from './formula-profiler';
import { getFormulaReplayConfigFromEnv } from './replay-config';

import '@univerjs/engine-formula/facade';

const unitId = 'test';
const opportunitiesSheetId = 'opportunities';
const pipelineSummarySheetId = 'pipeline-summary';
const dealScoringSheetId = 'deal-scoring';

const RUN_ENV = 'RUN_SALES_PIPELINE_CRASH';
const numOpps = Number(process.env.SALES_PIPELINE_OPPS ?? '10000');
const numReps = Number(process.env.SALES_PIPELINE_REPS ?? '50');
const numAccounts = Number(process.env.SALES_PIPELINE_ACCOUNTS ?? '800');
const stages = ['Prospecting', 'Qualification', 'Discovery', 'Proposal', 'Negotiation', 'Closed Won', 'Closed Lost'];
const stageProbabilities: Record<string, number> = {
    Prospecting: 0.1,
    Qualification: 0.2,
    Discovery: 0.3,
    Proposal: 0.5,
    Negotiation: 0.7,
    'Closed Won': 1,
    'Closed Lost': 0,
};

function buildWorkbookData(options?: { numOpps?: number; numReps?: number; numAccounts?: number }): IWorkbookData {
    const opps = options?.numOpps ?? numOpps;
    const reps = options?.numReps ?? numReps;
    const accounts = options?.numAccounts ?? numAccounts;
    const sheets: IWorkbookData['sheets'] = {
        [opportunitiesSheetId]: {
            id: opportunitiesSheetId,
            name: 'Opportunities',
            rowCount: opps + 10,
            columnCount: 8,
            cellData: {
                0: {
                    0: { v: 'Opp ID', t: CellValueType.STRING },
                    1: { v: 'Account', t: CellValueType.STRING },
                    2: { v: 'Rep', t: CellValueType.STRING },
                    3: { v: 'Region', t: CellValueType.STRING },
                    4: { v: 'Stage', t: CellValueType.STRING },
                    5: { v: 'Amount', t: CellValueType.STRING },
                    6: { v: 'Probability', t: CellValueType.STRING },
                    7: { v: 'Close Date', t: CellValueType.STRING },
                },
            },
        },
        [pipelineSummarySheetId]: {
            id: pipelineSummarySheetId,
            name: 'Pipeline Summary',
            rowCount: stages.length + 12,
            columnCount: 5,
            cellData: {
                0: {
                    0: { v: 'Stage', t: CellValueType.STRING },
                    1: { v: 'Count', t: CellValueType.STRING },
                    2: { v: 'Total Value', t: CellValueType.STRING },
                    3: { v: 'Avg Deal Size', t: CellValueType.STRING },
                    4: { v: 'Weighted Value', t: CellValueType.STRING },
                },
            },
        },
        [dealScoringSheetId]: {
            id: dealScoringSheetId,
            name: 'Deal Scoring',
            rowCount: opps + 10,
            columnCount: 8,
            cellData: {
                0: {
                    0: { v: 'Opp ID', t: CellValueType.STRING },
                    1: { v: 'Account', t: CellValueType.STRING },
                    2: { v: 'Rep', t: CellValueType.STRING },
                    3: { v: 'Amount', t: CellValueType.STRING },
                    4: { v: 'Rep Win Rate', t: CellValueType.STRING },
                    5: { v: 'Rep Avg Deal', t: CellValueType.STRING },
                    6: { v: 'Account Pipeline', t: CellValueType.STRING },
                    7: { v: 'Score', t: CellValueType.STRING },
                },
            },
        },
    };

    const opportunities = sheets[opportunitiesSheetId].cellData;
    const pipelineSummary = sheets[pipelineSummarySheetId].cellData;
    const dealScoring = sheets[dealScoringSheetId].cellData;

    const endRow = opps + 1;
    const accountRange = `Opportunities!$B$2:$B$${endRow}`;
    const repRange = `Opportunities!$C$2:$C$${endRow}`;
    const stageRange = `Opportunities!$E$2:$E$${endRow}`;
    const amountRange = `Opportunities!$F$2:$F$${endRow}`;
    const probabilityRange = `Opportunities!$G$2:$G$${endRow}`;

    for (let index = 0; index < opps; index++) {
        const row = index + 1;
        const sheetRow = row + 1;
        const rep = `Rep ${String((index % reps) + 1).padStart(2, '0')}`;
        const account = `ACC-${String((index % accounts) + 1).padStart(4, '0')}`;
        const region = ['West', 'East', 'Central', 'EMEA', 'APAC'][index % 5];
        const stage = stages[index % stages.length];
        const amount = 50000 + (index % 200) * 2500;
        const probability = stageProbabilities[stage];
        const closeDate = 45567 + (index % 720);

        opportunities[row] = {
            0: { v: `OPP-${String(index + 1).padStart(5, '0')}`, t: CellValueType.STRING },
            1: { v: account, t: CellValueType.STRING },
            2: { v: rep, t: CellValueType.STRING },
            3: { v: region, t: CellValueType.STRING },
            4: { v: stage, t: CellValueType.STRING },
            5: { v: amount, t: CellValueType.NUMBER },
            6: { v: probability, t: CellValueType.NUMBER },
            7: { v: closeDate, t: CellValueType.NUMBER },
        };

        dealScoring[row] = {
            0: { f: `=Opportunities!A${sheetRow}` },
            1: { f: `=Opportunities!B${sheetRow}` },
            2: { f: `=Opportunities!C${sheetRow}` },
            3: { f: `=Opportunities!F${sheetRow}` },
            4: {
                f: `=IF(COUNTIFS(${repRange},Opportunities!C${sheetRow},${stageRange},"Closed Won")+COUNTIFS(${repRange},Opportunities!C${sheetRow},${stageRange},"Closed Lost")>0,COUNTIFS(${repRange},Opportunities!C${sheetRow},${stageRange},"Closed Won")/(COUNTIFS(${repRange},Opportunities!C${sheetRow},${stageRange},"Closed Won")+COUNTIFS(${repRange},Opportunities!C${sheetRow},${stageRange},"Closed Lost")),0)`,
            },
            5: {
                f: `=IF(COUNTIF(${repRange},Opportunities!C${sheetRow})>0,SUMIF(${repRange},Opportunities!C${sheetRow},${amountRange})/COUNTIF(${repRange},Opportunities!C${sheetRow}),0)`,
            },
            6: {
                f: `=SUMPRODUCT((${accountRange}=Opportunities!B${sheetRow})*${amountRange}*${probabilityRange})`,
            },
            7: {
                f: `=E${sheetRow}*0.3+IF(F${sheetRow}>0,D${sheetRow}/F${sheetRow},0)*0.3+IF(G${sheetRow}>0,D${sheetRow}/G${sheetRow},0)*0.4`,
            },
        };
    }

    for (let index = 0; index < stages.length; index++) {
        const row = index + 1;
        const sheetRow = row + 1;
        pipelineSummary[row] = {
            0: { v: stages[index], t: CellValueType.STRING },
            1: { f: `=COUNTIF(${stageRange},A${sheetRow})` },
            2: { f: `=SUMIF(${stageRange},A${sheetRow},${amountRange})` },
            3: { f: `=IF(B${sheetRow}>0,C${sheetRow}/B${sheetRow},0)` },
            4: { f: `=SUMPRODUCT((${stageRange}=A${sheetRow})*${amountRange}*${probabilityRange})` },
        };
    }

    const totalRow = stages.length + 1;
    const totalSheetRow = totalRow + 1;
    pipelineSummary[totalRow] = {
        0: { v: 'Total', t: CellValueType.STRING },
        1: { f: `=SUM(B2:B${totalSheetRow - 1})` },
        2: { f: `=SUM(C2:C${totalSheetRow - 1})` },
        4: { f: `=SUM(E2:E${totalSheetRow - 1})` },
    };

    return {
        id: unitId,
        appVersion: '3.0.0-alpha',
        locale: LocaleType.ZH_CN,
        name: 'Sales Pipeline Crash Reproducer',
        sheetOrder: [opportunitiesSheetId, pipelineSummarySheetId, dealScoringSheetId],
        sheets,
        styles: {},
    };
}

async function runSalesPipelineScenario(options?: {
    numOpps?: number;
    numReps?: number;
    numAccounts?: number;
    batchExecutionCount?: number;
    intervalCount?: number;
}) {
    const testBed = createFunctionTestBed(buildWorkbookData(options));
    const get = testBed.get;
    const formulaEngine = testBed.api.getFormula() as FFormula;
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
        batchExecutionCount: options?.batchExecutionCount,
        intervalCount: options?.intervalCount,
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
        ...functionMeta,
    ]
        .map((registerObject) => {
            const Func = registerObject[0] as Ctor<BaseFunction>;
            const name = registerObject[1] as IFunctionNames;

            return new Func(name);
        });

    functionService.registerExecutors(...functions);

    const getCellValue = (sheetId: string, row: number, column: number) => {
        const worksheet = testBed.sheet.getSheetBySheetId(sheetId) as Worksheet;
        return worksheet.getCellRaw(row, column)?.v;
    };

    formulaEngine.executeCalculation();
    await formulaEngine.onCalculationEnd(300_000);

    return {
        totalCount: getCellValue(pipelineSummarySheetId, stages.length + 1, 1),
        firstAccountPipeline: getCellValue(dealScoringSheetId, 1, 6),
        sampleScore: getCellValue(dealScoringSheetId, Math.min(options?.numOpps ?? numOpps, 100), 7),
    };
}

describe('sales pipeline workload crash reproducer', () => {
    let get: Injector['get'];
    let formulaEngine: FFormula;
    let commandService: ICommandService;
    let calcService: ICalculateFormulaService;
    let getCellValue: (sheetId: string, row: number, column: number) => Nullable<CellValue>;
    let _univer: ReturnType<typeof createFunctionTestBed>['univer'];

    beforeEach(async () => {
        const testBed = createFunctionTestBed(buildWorkbookData());
        _univer = testBed.univer;

        get = testBed.get;
        formulaEngine = testBed.api.getFormula() as FFormula;
        calcService = get(ICalculateFormulaService);
        commandService = get(ICommandService);

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
            ...functionMeta,
        ]
            .map((registerObject) => {
                const Func = registerObject[0] as Ctor<BaseFunction>;
                const name = registerObject[1] as IFunctionNames;

                return new Func(name);
            });

        functionService.registerExecutors(...functions);

        getCellValue = (sheetId: string, row: number, column: number) => {
            const worksheet = testBed.sheet.getSheetBySheetId(sheetId) as Worksheet;
            return worksheet.getCellRaw(row, column)?.v;
        };
    });

    afterEach(() => {
        _univer?.dispose();
    });

    const runHeavy = process.env[RUN_ENV] === '1';
    const heavyIt = runHeavy ? it : it.skip;

    heavyIt(
        `replays the sales-pipeline formula mix across ${numOpps} opportunities`,
        async () => {
            const result = await profileCalculation(formulaEngine, calcService, {
                label: `sales-pipeline-${numOpps}`,
                timeoutMs: 300_000,
                cpu: !!process.env.FORMULA_CPU_PROF,
                heap: !!process.env.FORMULA_HEAP_PROF,
            });

            expect(getCellValue(pipelineSummarySheetId, stages.length + 1, 1)).toBe(numOpps);
            expect(getCellValue(dealScoringSheetId, 1, 6)).not.toBeNull();
            expect(getCellValue(dealScoringSheetId, Math.min(numOpps, 100), 7)).not.toBeNull();

            if (result.trace.length > 0) {
                expect(result.peakHeapMB).toBeLessThan(1800);
            }
        },
        300000
    );

    it('computes the expected result with tiny execution batches', async () => {
        const batched = await runSalesPipelineScenario({
            numOpps: 120,
            numReps: 12,
            numAccounts: 40,
            batchExecutionCount: 7,
        });

        expect(batched.totalCount).toBe(120);
        expect(batched.firstAccountPipeline).not.toBeNull();
        expect(batched.sampleScore).not.toBeNull();
    }, 30000);
});

describe('sales pipeline correctness (120 rows)', () => {
    const testOpps = 120;
    const testReps = 12;
    const testAccounts = 40;

    function oppData(index: number) {
        const rep = `Rep ${String((index % testReps) + 1).padStart(2, '0')}`;
        const account = `ACC-${String((index % testAccounts) + 1).padStart(4, '0')}`;
        const stage = stages[index % stages.length];
        const amount = 50000 + (index % 200) * 2500;
        const probability = stageProbabilities[stage];
        return { rep, account, stage, amount, probability };
    }

    function allOpps() {
        return Array.from({ length: testOpps }, (_, i) => oppData(i));
    }

    it('verifies every formula cell against JS reference computation', async () => {
        const scenario = createFunctionTestBed(buildWorkbookData({
            numOpps: testOpps,
            numReps: testReps,
            numAccounts: testAccounts,
        }));
        const get = scenario.get;
        const formulaEngine = scenario.api.getFormula() as FFormula;
        const commandService = get(ICommandService);

        commandService.registerCommand(SetFormulaCalculationStartMutation);
        commandService.registerCommand(SetFormulaCalculationStopMutation);
        commandService.registerCommand(SetFormulaCalculationResultMutation);
        commandService.registerCommand(SetFormulaCalculationNotificationMutation);
        commandService.registerCommand(SetArrayFormulaDataMutation);
        commandService.registerCommand(SetRangeValuesMutation);

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
                [scenario.unitId]: scenario.sheetData,
            },
        });

        const sheetItem = scenario.sheetData[scenario.sheetId];
        formulaRuntimeService.setCurrent(0, 0, sheetItem.rowCount, sheetItem.columnCount, scenario.sheetId, scenario.unitId);

        const functions = [...functionMath, ...functionStatistical, ...functionLogical, ...functionMeta]
            .map((reg) => new (reg[0] as Ctor<BaseFunction>)(reg[1] as IFunctionNames));
        functionService.registerExecutors(...functions);

        const cell = (sheetId: string, row: number, col: number) => {
            const ws = scenario.sheet.getSheetBySheetId(sheetId) as Worksheet;
            return ws.getCellRaw(row, col)?.v;
        };

        formulaEngine.executeCalculation();
        await formulaEngine.onCalculationEnd(60_000);

        const opps = allOpps();

        // --- Pipeline Summary sheet: 7 stage rows ---
        for (let s = 0; s < stages.length; s++) {
            const stage = stages[s];
            const row = s + 1;

            const matching = opps.filter((o) => o.stage === stage);
            const expectedCount = matching.length;
            const expectedTotal = matching.reduce((sum, o) => sum + o.amount, 0);
            const expectedAvg = expectedCount > 0 ? expectedTotal / expectedCount : 0;
            const expectedWeighted = matching.reduce((sum, o) => sum + o.amount * o.probability, 0);

            expect(cell(pipelineSummarySheetId, row, 1)).toBe(expectedCount);
            expect(cell(pipelineSummarySheetId, row, 2)).toBe(expectedTotal);
            expect(cell(pipelineSummarySheetId, row, 3)).toBeCloseTo(expectedAvg, 5);
            expect(cell(pipelineSummarySheetId, row, 4)).toBeCloseTo(expectedWeighted, 2);
        }

        // Total row
        const totalRow = stages.length + 1;
        expect(cell(pipelineSummarySheetId, totalRow, 1)).toBe(testOpps);
        const totalValue = opps.reduce((sum, o) => sum + o.amount, 0);
        expect(cell(pipelineSummarySheetId, totalRow, 2)).toBe(totalValue);
        const totalWeighted = opps.reduce((sum, o) => sum + o.amount * o.probability, 0);
        expect(cell(pipelineSummarySheetId, totalRow, 4)).toBeCloseTo(totalWeighted, 2);

        // --- Deal Scoring sheet: every row ---
        for (let i = 0; i < testOpps; i++) {
            const row = i + 1;
            const o = opps[i];

            // Col A-D: direct references
            expect(cell(dealScoringSheetId, row, 0)).toBe(`OPP-${String(i + 1).padStart(5, '0')}`);
            expect(cell(dealScoringSheetId, row, 1)).toBe(o.account);
            expect(cell(dealScoringSheetId, row, 2)).toBe(o.rep);
            expect(cell(dealScoringSheetId, row, 3)).toBe(o.amount);

            // Col E (4): Rep Win Rate
            const repOpps = opps.filter((x) => x.rep === o.rep);
            const repWon = repOpps.filter((x) => x.stage === 'Closed Won').length;
            const repLost = repOpps.filter((x) => x.stage === 'Closed Lost').length;
            const repClosed = repWon + repLost;
            const expectedWinRate = repClosed > 0 ? repWon / repClosed : 0;
            expect(cell(dealScoringSheetId, row, 4)).toBeCloseTo(expectedWinRate, 8);

            // Col F (5): Rep Avg Deal
            const repCount = repOpps.length;
            const repTotalAmt = repOpps.reduce((sum, x) => sum + x.amount, 0);
            const expectedAvgDeal = repCount > 0 ? repTotalAmt / repCount : 0;
            expect(cell(dealScoringSheetId, row, 5)).toBeCloseTo(expectedAvgDeal, 5);

            // Col G (6): Account Pipeline = SUMPRODUCT((account=acct)*amount*probability)
            const acctOpps = opps.filter((x) => x.account === o.account);
            const expectedAcctPipeline = acctOpps.reduce((sum, x) => sum + x.amount * x.probability, 0);
            expect(cell(dealScoringSheetId, row, 6)).toBeCloseTo(expectedAcctPipeline, 2);

            // Col H (7): Score = winRate*0.3 + (amount/avgDeal)*0.3 + (amount/acctPipeline)*0.4
            const dealRatio = expectedAvgDeal > 0 ? o.amount / expectedAvgDeal : 0;
            const pipelineRatio = expectedAcctPipeline > 0 ? o.amount / expectedAcctPipeline : 0;
            const expectedScore = expectedWinRate * 0.3 + dealRatio * 0.3 + pipelineRatio * 0.4;
            expect(cell(dealScoringSheetId, row, 7)).toBeCloseTo(expectedScore, 6);
        }

        scenario.univer?.dispose();
    }, 60000);
});
