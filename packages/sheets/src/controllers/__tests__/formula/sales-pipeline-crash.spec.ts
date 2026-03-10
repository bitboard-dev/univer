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
    IFormulaCurrentConfigService,
    IFormulaRuntimeService,
    IFunctionService,
    SetArrayFormulaDataMutation,
    SetFormulaCalculationNotificationMutation,
    SetFormulaCalculationResultMutation,
    SetFormulaCalculationStartMutation,
    SetFormulaCalculationStopMutation,
} from '@univerjs/engine-formula';
import { beforeEach, describe, expect, it } from 'vitest';
import { SetRangeValuesMutation } from '../../../commands/mutations/set-range-values.mutation';
import { createFunctionTestBed } from './create-function-test-bed';

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
    await formulaEngine.onCalculationEnd();

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
    let getCellValue: (sheetId: string, row: number, column: number) => Nullable<CellValue>;

    beforeEach(async () => {
        const testBed = createFunctionTestBed(buildWorkbookData());

        get = testBed.get;
        formulaEngine = testBed.api.getFormula() as FFormula;
        commandService = get(ICommandService);

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

    const runHeavy = process.env[RUN_ENV] === '1';
    const heavyIt = runHeavy ? it : it.skip;

    heavyIt(
        `replays the sales-pipeline formula mix across ${numOpps} opportunities`,
        async () => {
            formulaEngine.executeCalculation();
            await formulaEngine.onCalculationEnd();

            expect(getCellValue(pipelineSummarySheetId, stages.length + 1, 1)).toBe(numOpps);
            expect(getCellValue(dealScoringSheetId, 1, 6)).not.toBeNull();
            expect(getCellValue(dealScoringSheetId, Math.min(numOpps, 100), 7)).not.toBeNull();
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
