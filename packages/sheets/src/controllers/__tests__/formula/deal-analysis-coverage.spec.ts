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
import { CellValueType, ICommandService, LocaleType } from '@univerjs/core';
import {
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

import '@univerjs/engine-formula/facade';

const unitId = 'test';
const opportunitiesSheetId = 'opportunities';
const accountsSheetId = 'accounts';
const dealAnalysisSheetId = 'deal-analysis';

const RUN_ENV = 'RUN_DEAL_ANALYSIS_COVERAGE';

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
        setup({ numOpps: 120, numReps: 12, numAccounts: 40 });

        formulaEngine.executeCalculation();
        await formulaEngine.onCalculationEnd(60_000);

        // VLOOKUP: first opp has ACC-0001, should resolve to an industry string
        const vlookup = getCellValue(dealAnalysisSheetId, 1, 5);
        expect(vlookup).toBeTruthy();
        expect(typeof vlookup).toBe('string');

        // INDEX/MATCH: should resolve to employee count (a number)
        const indexMatch = getCellValue(dealAnalysisSheetId, 1, 6);
        expect(typeof indexMatch).toBe('number');
        expect(indexMatch).toBeGreaterThan(0);

        // SUMIFS: rep's stage pipeline should be > 0
        const sumifs = getCellValue(dealAnalysisSheetId, 1, 7);
        expect(typeof sumifs).toBe('number');
        expect(sumifs).toBeGreaterThan(0);

        // AVERAGEIFS: rep's stage avg deal — may be 0 if IFERROR catches an error
        const averageifs = getCellValue(dealAnalysisSheetId, 1, 8);
        expect(typeof averageifs).toBe('number');

        // MAXIFS: rep's max deal should be >= this deal's amount
        const maxifs = getCellValue(dealAnalysisSheetId, 1, 9);
        const amount = getCellValue(dealAnalysisSheetId, 1, 3);
        expect(typeof maxifs).toBe('number');
        expect(maxifs as number).toBeGreaterThanOrEqual(amount as number);

        // MINIFS: rep's min deal should be <= this deal's amount
        const minifs = getCellValue(dealAnalysisSheetId, 1, 10);
        expect(typeof minifs).toBe('number');
        expect(minifs as number).toBeLessThanOrEqual(amount as number);

        // RANK.EQ: should be between 1 and numOpps
        const rank = getCellValue(dealAnalysisSheetId, 1, 11);
        expect(typeof rank).toBe('number');
        expect(rank as number).toBeGreaterThanOrEqual(1);
        expect(rank as number).toBeLessThanOrEqual(120);

        // LARGE: 5th largest deal, should be a number
        const large = getCellValue(dealAnalysisSheetId, 1, 12);
        expect(typeof large).toBe('number');
        expect(large as number).toBeGreaterThan(0);

        // PERCENTRANK.INC: should be between 0 and 1
        const percentrank = getCellValue(dealAnalysisSheetId, 1, 13);
        expect(typeof percentrank).toBe('number');
        expect(percentrank as number).toBeGreaterThanOrEqual(0);
        expect(percentrank as number).toBeLessThanOrEqual(1);
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
