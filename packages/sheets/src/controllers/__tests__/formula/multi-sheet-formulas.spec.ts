import type { Ctor, IWorkbookData, Worksheet } from '@univerjs/core';
import type { BaseFunction, IFunctionNames } from '@univerjs/engine-formula';
import type { FFormula } from '@univerjs/engine-formula/facade';
import { CellValueType, ICommandService, LocaleType } from '@univerjs/core';
import {
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
import { afterEach, describe, expect, it } from 'vitest';
import { SetRangeValuesMutation } from '../../../commands/mutations/set-range-values.mutation';
import { createFunctionTestBed } from './create-function-test-bed';

import '@univerjs/engine-formula/facade';

const unitId = 'test';

const allFunctions = [
    ...functionMath, ...functionStatistical, ...functionLogical,
    ...functionLookup, ...functionMeta, ...functionDate,
    ...functionText, ...functionInformation, ...functionFinancial,
    ...functionEngineering, ...functionDatabase, ...functionCompatibility,
    ...functionArray, ...functionWeb,
];

function setup(workbookData: IWorkbookData) {
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

    const functionService = get(IFunctionService);
    const configService = get(IFormulaCurrentConfigService);
    const runtimeService = get(IFormulaRuntimeService);

    configService.load({
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

    const firstSheet = testBed.sheetData[testBed.sheetId];
    runtimeService.setCurrent(0, 0, firstSheet.rowCount, firstSheet.columnCount, testBed.sheetId, testBed.unitId);

    const fns = allFunctions.map((reg) => {
        const Func = reg[0] as Ctor<BaseFunction>;
        const name = reg[1] as IFunctionNames;
        return new Func(name);
    });
    functionService.registerExecutors(...fns);

    const getCellValue = (sheetId: string, row: number, col: number) => {
        const ws = testBed.sheet.getSheetBySheetId(sheetId) as Worksheet;
        return ws.getCellRaw(row, col)?.v;
    };

    return { testBed, formulaEngine, calcService, getCellValue };
}

async function runCalc(formulaEngine: FFormula) {
    formulaEngine.executeCalculation();
    await formulaEngine.onCalculationEnd(30_000);
}

// ---------------------------------------------------------------------------
// Workbook: Sales detail on Sheet1, regional summary on Sheet2, dashboard on Sheet3
// ---------------------------------------------------------------------------
function buildSalesWorkbook(): IWorkbookData {
    const regions = ['North', 'South', 'East', 'West'];
    const N = 40;

    const detailCellData: Record<number, Record<number, any>> = {};
    detailCellData[0] = {
        0: { v: 'Region', t: CellValueType.STRING },
        1: { v: 'Q1', t: CellValueType.STRING },
        2: { v: 'Q2', t: CellValueType.STRING },
        3: { v: 'Q3', t: CellValueType.STRING },
        4: { v: 'Q4', t: CellValueType.STRING },
    };

    let seed = 42;
    const rand = () => { seed = (seed * 1664525 + 1013904223) & 0x7fffffff; return seed / 0x7fffffff; };

    for (let i = 0; i < N; i++) {
        const r = i + 1;
        detailCellData[r] = {
            0: { v: regions[i % 4], t: CellValueType.STRING },
            1: { v: Math.round(rand() * 50000), t: CellValueType.NUMBER },
            2: { v: Math.round(rand() * 50000), t: CellValueType.NUMBER },
            3: { v: Math.round(rand() * 50000), t: CellValueType.NUMBER },
            4: { v: Math.round(rand() * 50000), t: CellValueType.NUMBER },
        };
    }

    // Sheet2: Regional summary — cross-sheet SUMIF and AVERAGEIF
    const summaryCellData: Record<number, Record<number, any>> = {};
    summaryCellData[0] = {
        0: { v: 'Region', t: CellValueType.STRING },
        1: { v: 'Total Q1', t: CellValueType.STRING },
        2: { v: 'Total Q2', t: CellValueType.STRING },
        3: { v: 'Total Q3', t: CellValueType.STRING },
        4: { v: 'Total Q4', t: CellValueType.STRING },
        5: { v: 'Annual', t: CellValueType.STRING },
        6: { v: 'Avg Q1', t: CellValueType.STRING },
        7: { v: 'Count', t: CellValueType.STRING },
        8: { v: 'Max Q1', t: CellValueType.STRING },
        9: { v: 'Min Q1', t: CellValueType.STRING },
    };

    for (let i = 0; i < 4; i++) {
        const r = i + 1;
        const region = regions[i];
        summaryCellData[r] = {
            0: { v: region, t: CellValueType.STRING },
            1: { f: `=SUMIF(Detail!$A$2:$A$${N + 1},A${r + 1},Detail!$B$2:$B$${N + 1})` },
            2: { f: `=SUMIF(Detail!$A$2:$A$${N + 1},A${r + 1},Detail!$C$2:$C$${N + 1})` },
            3: { f: `=SUMIF(Detail!$A$2:$A$${N + 1},A${r + 1},Detail!$D$2:$D$${N + 1})` },
            4: { f: `=SUMIF(Detail!$A$2:$A$${N + 1},A${r + 1},Detail!$E$2:$E$${N + 1})` },
            5: { f: `=B${r + 1}+C${r + 1}+D${r + 1}+E${r + 1}` },
            6: { f: `=AVERAGEIF(Detail!$A$2:$A$${N + 1},A${r + 1},Detail!$B$2:$B$${N + 1})` },
            7: { f: `=COUNTIF(Detail!$A$2:$A$${N + 1},A${r + 1})` },
            8: { f: `=MAXIFS(Detail!$B$2:$B$${N + 1},Detail!$A$2:$A$${N + 1},A${r + 1})` },
            9: { f: `=MINIFS(Detail!$B$2:$B$${N + 1},Detail!$A$2:$A$${N + 1},A${r + 1})` },
        };
    }
    // Grand total row referencing same-sheet cells
    summaryCellData[5] = {
        0: { v: 'Grand Total', t: CellValueType.STRING },
        1: { f: '=SUM(B2:B5)' },
        2: { f: '=SUM(C2:C5)' },
        3: { f: '=SUM(D2:D5)' },
        4: { f: '=SUM(E2:E5)' },
        5: { f: '=SUM(F2:F5)' },
        6: { f: '=AVERAGE(G2:G5)' },
        7: { f: '=SUM(H2:H5)' },
    };

    // Sheet3: Dashboard — references Summary sheet (chain depth 2)
    const dashCellData: Record<number, Record<number, any>> = {};
    dashCellData[0] = {
        0: { v: 'Metric', t: CellValueType.STRING },
        1: { v: 'Value', t: CellValueType.STRING },
    };
    dashCellData[1] = {
        0: { v: 'Grand Total Revenue', t: CellValueType.STRING },
        1: { f: '=Summary!F6' },
    };
    dashCellData[2] = {
        0: { v: 'Top Region Q1', t: CellValueType.STRING },
        1: { f: '=INDEX(Summary!A2:A5,MATCH(MAX(Summary!B2:B5),Summary!B2:B5,0))' },
    };
    dashCellData[3] = {
        0: { v: 'Total Reps', t: CellValueType.STRING },
        1: { f: '=Summary!H6' },
    };
    dashCellData[4] = {
        0: { v: 'Avg Q1 Across Regions', t: CellValueType.STRING },
        1: { f: '=AVERAGE(Summary!B2:B5)' },
    };
    dashCellData[5] = {
        0: { v: 'Max Annual', t: CellValueType.STRING },
        1: { f: '=MAX(Summary!F2:F5)' },
    };
    dashCellData[6] = {
        0: { v: 'Q1 Share of Grand Total', t: CellValueType.STRING },
        1: { f: '=Summary!B6/Summary!F6' },
    };
    dashCellData[7] = {
        0: { v: 'North % of Total', t: CellValueType.STRING },
        1: { f: '=Summary!F2/Summary!F6' },
    };
    dashCellData[8] = {
        0: { v: 'Detail Row 1 Q1', t: CellValueType.STRING },
        1: { f: '=Detail!B2' },
    };
    dashCellData[9] = {
        0: { v: 'VLOOKUP from Detail', t: CellValueType.STRING },
        1: { f: `=VLOOKUP("South",Detail!A2:E${N + 1},2,FALSE)` },
    };

    return {
        id: unitId,
        appVersion: '1.0.0',
        name: 'multi-sheet-sales',
        locale: LocaleType.EN_US,
        styles: {},
        sheetOrder: ['detail', 'summary', 'dashboard'],
        sheets: {
            detail: {
                id: 'detail',
                name: 'Detail',
                rowCount: N + 10,
                columnCount: 5,
                cellData: detailCellData,
            },
            summary: {
                id: 'summary',
                name: 'Summary',
                rowCount: 20,
                columnCount: 10,
                cellData: summaryCellData,
            },
            dashboard: {
                id: 'dashboard',
                name: 'Dashboard',
                rowCount: 20,
                columnCount: 2,
                cellData: dashCellData,
            },
        },
        resources: [],
    };
}

// ---------------------------------------------------------------------------
// Workbook: Shared constants sheet + computation sheet (common pattern)
// ---------------------------------------------------------------------------
function buildConstantsWorkbook(): IWorkbookData {
    const configCellData: Record<number, Record<number, any>> = {};
    configCellData[0] = {
        0: { v: 'Parameter', t: CellValueType.STRING },
        1: { v: 'Value', t: CellValueType.STRING },
    };
    configCellData[1] = { 0: { v: 'Tax Rate', t: CellValueType.STRING }, 1: { v: 0.08, t: CellValueType.NUMBER } };
    configCellData[2] = { 0: { v: 'Discount', t: CellValueType.STRING }, 1: { v: 0.15, t: CellValueType.NUMBER } };
    configCellData[3] = { 0: { v: 'Shipping', t: CellValueType.STRING }, 1: { v: 12.5, t: CellValueType.NUMBER } };
    configCellData[4] = { 0: { v: 'Threshold', t: CellValueType.STRING }, 1: { v: 100, t: CellValueType.NUMBER } };

    const ordersCellData: Record<number, Record<number, any>> = {};
    ordersCellData[0] = {
        0: { v: 'Item', t: CellValueType.STRING },
        1: { v: 'Qty', t: CellValueType.STRING },
        2: { v: 'Unit Price', t: CellValueType.STRING },
        3: { v: 'Subtotal', t: CellValueType.STRING },
        4: { v: 'Discount', t: CellValueType.STRING },
        5: { v: 'Tax', t: CellValueType.STRING },
        6: { v: 'Shipping', t: CellValueType.STRING },
        7: { v: 'Total', t: CellValueType.STRING },
    };

    const items = [
        { item: 'Widget A', qty: 5, price: 29.99 },
        { item: 'Widget B', qty: 12, price: 14.50 },
        { item: 'Gadget C', qty: 1, price: 199.00 },
        { item: 'Part D', qty: 50, price: 3.25 },
        { item: 'Service E', qty: 2, price: 75.00 },
    ];

    for (let i = 0; i < items.length; i++) {
        const r = i + 1;
        const { item, qty, price } = items[i];
        ordersCellData[r + 0] = {
            0: { v: item, t: CellValueType.STRING },
            1: { v: qty, t: CellValueType.NUMBER },
            2: { v: price, t: CellValueType.NUMBER },
            3: { f: `=B${r + 1}*C${r + 1}` },
            4: { f: `=D${r + 1}*Config!$B$3` },
            5: { f: `=(D${r + 1}-E${r + 1})*Config!$B$2` },
            6: { f: `=IF(D${r + 1}>=Config!$B$5,0,Config!$B$4)` },
            7: { f: `=D${r + 1}-E${r + 1}+F${r + 1}+G${r + 1}` },
        };
    }

    const totalsRow = items.length + 1;
    ordersCellData[totalsRow] = {
        0: { v: 'TOTALS', t: CellValueType.STRING },
        3: { f: `=SUM(D2:D${totalsRow})` },
        4: { f: `=SUM(E2:E${totalsRow})` },
        5: { f: `=SUM(F2:F${totalsRow})` },
        6: { f: `=SUM(G2:G${totalsRow})` },
        7: { f: `=SUM(H2:H${totalsRow})` },
    };

    return {
        id: unitId,
        appVersion: '1.0.0',
        name: 'multi-sheet-constants',
        locale: LocaleType.EN_US,
        styles: {},
        sheetOrder: ['config', 'orders'],
        sheets: {
            config: {
                id: 'config',
                name: 'Config',
                rowCount: 10,
                columnCount: 2,
                cellData: configCellData,
            },
            orders: {
                id: 'orders',
                name: 'Orders',
                rowCount: 20,
                columnCount: 8,
                cellData: ordersCellData,
            },
        },
        resources: [],
    };
}

// ---------------------------------------------------------------------------
// Workbook: Monthly sheets (Jan..Dec) + Annual summary (common pattern)
// ---------------------------------------------------------------------------
function buildMonthlyWorkbook(): IWorkbookData {
    const months = ['Jan', 'Feb', 'Mar'];
    const sheets: IWorkbookData['sheets'] = {};
    const sheetOrder: string[] = [];

    for (const month of months) {
        const id = month.toLowerCase();
        const cellData: Record<number, Record<number, any>> = {};
        cellData[0] = {
            0: { v: 'Metric', t: CellValueType.STRING },
            1: { v: 'Value', t: CellValueType.STRING },
        };
        cellData[1] = { 0: { v: 'Revenue', t: CellValueType.STRING }, 1: { v: months.indexOf(month) * 1000 + 5000, t: CellValueType.NUMBER } };
        cellData[2] = { 0: { v: 'Expenses', t: CellValueType.STRING }, 1: { v: months.indexOf(month) * 500 + 3000, t: CellValueType.NUMBER } };
        cellData[3] = { 0: { v: 'Profit', t: CellValueType.STRING }, 1: { f: '=B2-B3' } };

        sheets[id] = { id, name: month, rowCount: 10, columnCount: 2, cellData };
        sheetOrder.push(id);
    }

    // Annual summary referencing all monthly sheets
    const annualCellData: Record<number, Record<number, any>> = {};
    annualCellData[0] = {
        0: { v: 'Metric', t: CellValueType.STRING },
        1: { v: 'Jan', t: CellValueType.STRING },
        2: { v: 'Feb', t: CellValueType.STRING },
        3: { v: 'Mar', t: CellValueType.STRING },
        4: { v: 'Total', t: CellValueType.STRING },
        5: { v: 'Average', t: CellValueType.STRING },
        6: { v: 'Max', t: CellValueType.STRING },
        7: { v: 'Min', t: CellValueType.STRING },
    };
    annualCellData[1] = {
        0: { v: 'Revenue', t: CellValueType.STRING },
        1: { f: '=Jan!B2' },
        2: { f: '=Feb!B2' },
        3: { f: '=Mar!B2' },
        4: { f: '=SUM(B2:D2)' },
        5: { f: '=AVERAGE(B2:D2)' },
        6: { f: '=MAX(B2:D2)' },
        7: { f: '=MIN(B2:D2)' },
    };
    annualCellData[2] = {
        0: { v: 'Expenses', t: CellValueType.STRING },
        1: { f: '=Jan!B3' },
        2: { f: '=Feb!B3' },
        3: { f: '=Mar!B3' },
        4: { f: '=SUM(B3:D3)' },
        5: { f: '=AVERAGE(B3:D3)' },
        6: { f: '=MAX(B3:D3)' },
        7: { f: '=MIN(B3:D3)' },
    };
    annualCellData[3] = {
        0: { v: 'Profit', t: CellValueType.STRING },
        1: { f: '=Jan!B4' },
        2: { f: '=Feb!B4' },
        3: { f: '=Mar!B4' },
        4: { f: '=SUM(B4:D4)' },
        5: { f: '=AVERAGE(B4:D4)' },
        6: { f: '=MAX(B4:D4)' },
        7: { f: '=MIN(B4:D4)' },
    };
    // Margin row: profit / revenue (chain depth 3: annual → monthly profit → monthly rev - exp)
    annualCellData[4] = {
        0: { v: 'Margin', t: CellValueType.STRING },
        1: { f: '=IF(B2=0,0,B4/B2)' },
        2: { f: '=IF(C2=0,0,C4/C2)' },
        3: { f: '=IF(D2=0,0,D4/D2)' },
        4: { f: '=IF(E2=0,0,E4/E2)' },
    };

    sheets['annual'] = { id: 'annual', name: 'Annual', rowCount: 10, columnCount: 8, cellData: annualCellData };
    sheetOrder.push('annual');

    return {
        id: unitId,
        appVersion: '1.0.0',
        name: 'multi-sheet-monthly',
        locale: LocaleType.EN_US,
        styles: {},
        sheetOrder,
        sheets,
        resources: [],
    };
}

// ===========================================================================
// Tests
// ===========================================================================

describe('multi-sheet formula tests', () => {
    describe('sales detail → regional summary → dashboard (3-sheet chain)', () => {
        it('computes cross-sheet SUMIF, COUNTIF, MAXIFS, MINIFS correctly', async () => {
            const { testBed, formulaEngine, getCellValue } = setup(buildSalesWorkbook());

            try {
                await runCalc(formulaEngine);

                // Summary!H2 = COUNTIF for "North" — should be 10 (40 rows / 4 regions)
                expect(getCellValue('summary', 1, 7)).toBe(10);
                expect(getCellValue('summary', 2, 7)).toBe(10);
                expect(getCellValue('summary', 3, 7)).toBe(10);
                expect(getCellValue('summary', 4, 7)).toBe(10);

                // Grand total count = 40
                expect(getCellValue('summary', 5, 7)).toBe(40);

                // Summary!B6 (grand total Q1) should equal Detail column B sum
                const grandQ1 = getCellValue('summary', 5, 1) as number;
                expect(grandQ1).toBeGreaterThan(0);

                // Grand total annual = sum of all 4 quarter totals
                const grandAnnual = getCellValue('summary', 5, 5) as number;
                expect(grandAnnual).toBeGreaterThan(grandQ1);

                // Per-region annual = sum of 4 quarter sums
                for (let i = 1; i <= 4; i++) {
                    const q1 = getCellValue('summary', i, 1) as number;
                    const q2 = getCellValue('summary', i, 2) as number;
                    const q3 = getCellValue('summary', i, 3) as number;
                    const q4 = getCellValue('summary', i, 4) as number;
                    const annual = getCellValue('summary', i, 5) as number;
                    expect(annual).toBe(q1 + q2 + q3 + q4);
                }

                // MAXIFS / MINIFS should be sensible
                for (let i = 1; i <= 4; i++) {
                    const maxQ1 = getCellValue('summary', i, 8) as number;
                    const minQ1 = getCellValue('summary', i, 9) as number;
                    const avgQ1 = getCellValue('summary', i, 6) as number;
                    expect(maxQ1).toBeGreaterThanOrEqual(avgQ1);
                    expect(minQ1).toBeLessThanOrEqual(avgQ1);
                }
            } finally {
                testBed.univer.dispose();
            }
        });

        it('dashboard reads from summary (depth-2 chain)', async () => {
            const { testBed, formulaEngine, getCellValue } = setup(buildSalesWorkbook());

            try {
                await runCalc(formulaEngine);

                // Dashboard!B2 = Grand Total Revenue = Summary!F6
                const dashTotal = getCellValue('dashboard', 1, 1) as number;
                const summaryTotal = getCellValue('summary', 5, 5) as number;
                expect(dashTotal).toBe(summaryTotal);
                expect(dashTotal).toBeGreaterThan(0);

                // Dashboard!B4 = Total Reps = Summary!H6 = 40
                expect(getCellValue('dashboard', 3, 1)).toBe(40);

                // Dashboard!B5 = AVERAGE(Summary!B2:B5) — average of regional Q1 totals
                const avgQ1 = getCellValue('dashboard', 4, 1) as number;
                expect(avgQ1).toBeGreaterThan(0);

                // Dashboard!B6 = MAX(Summary!F2:F5) — max annual across regions
                const maxAnnual = getCellValue('dashboard', 5, 1) as number;
                expect(maxAnnual).toBeGreaterThan(0);
                expect(maxAnnual).toBeLessThanOrEqual(summaryTotal);

                // Dashboard!B7 = Q1 share of grand total — should be between 0 and 1
                const q1Share = getCellValue('dashboard', 6, 1) as number;
                expect(q1Share).toBeGreaterThan(0);
                expect(q1Share).toBeLessThan(1);

                // Dashboard!B9 = Detail!B2 — direct cross-sheet single cell
                const detailB2 = getCellValue('detail', 1, 1) as number;
                expect(getCellValue('dashboard', 8, 1)).toBe(detailB2);

                // Dashboard!B3 = INDEX/MATCH for top region — should be one of the 4 regions
                const topRegion = getCellValue('dashboard', 2, 1) as string;
                expect(['North', 'South', 'East', 'West']).toContain(topRegion);
            } finally {
                testBed.univer.dispose();
            }
        });
    });

    describe('config sheet → orders (shared constants pattern)', () => {
        it('propagates config values into order calculations', async () => {
            const { testBed, formulaEngine, getCellValue } = setup(buildConstantsWorkbook());

            try {
                await runCalc(formulaEngine);

                const taxRate = 0.08;
                const discount = 0.15;
                const shipping = 12.5;
                const threshold = 100;

                // Verify config sheet values
                expect(getCellValue('config', 1, 1)).toBe(taxRate);
                expect(getCellValue('config', 2, 1)).toBe(discount);
                expect(getCellValue('config', 3, 1)).toBe(shipping);
                expect(getCellValue('config', 4, 1)).toBe(threshold);

                // Check first order row: Widget A, qty=5, price=29.99
                const subtotal1 = getCellValue('orders', 1, 3) as number;
                expect(subtotal1).toBeCloseTo(5 * 29.99, 2);

                const discount1 = getCellValue('orders', 1, 4) as number;
                expect(discount1).toBeCloseTo(subtotal1 * discount, 2);

                const tax1 = getCellValue('orders', 1, 5) as number;
                expect(tax1).toBeCloseTo((subtotal1 - discount1) * taxRate, 2);

                // Subtotal = 149.95 > threshold=100, so shipping = 0
                const ship1 = getCellValue('orders', 1, 6) as number;
                expect(ship1).toBe(0);

                // Gadget C: qty=1, price=199, subtotal=199 > 100 → free shipping
                const subtotal3 = getCellValue('orders', 3, 3) as number;
                expect(subtotal3).toBe(199);
                expect(getCellValue('orders', 3, 6)).toBe(0);

                // Part D: qty=50, price=3.25, subtotal=162.50 > 100 → free shipping
                const subtotal4 = getCellValue('orders', 4, 3) as number;
                expect(subtotal4).toBeCloseTo(162.50, 2);
                expect(getCellValue('orders', 4, 6)).toBe(0);

                // Totals row
                const totalSubtotal = getCellValue('orders', 6, 3) as number;
                expect(totalSubtotal).toBeGreaterThan(0);

                // Total = sum of individual totals
                let manualSum = 0;
                for (let r = 1; r <= 5; r++) {
                    manualSum += getCellValue('orders', r, 7) as number;
                }
                const totalTotal = getCellValue('orders', 6, 7) as number;
                expect(totalTotal).toBeCloseTo(manualSum, 2);
            } finally {
                testBed.univer.dispose();
            }
        });
    });

    describe('monthly sheets → annual summary (3-depth chain with margins)', () => {
        it('pulls monthly values and computes aggregates', async () => {
            const { testBed, formulaEngine, getCellValue } = setup(buildMonthlyWorkbook());

            try {
                await runCalc(formulaEngine);

                // Monthly profits: Jan=5000-3000=2000, Feb=6000-3500=2500, Mar=7000-4000=3000
                expect(getCellValue('jan', 3, 1)).toBe(2000);
                expect(getCellValue('feb', 3, 1)).toBe(2500);
                expect(getCellValue('mar', 3, 1)).toBe(3000);

                // Annual pulls monthly revenues
                expect(getCellValue('annual', 1, 1)).toBe(5000);  // Jan revenue
                expect(getCellValue('annual', 1, 2)).toBe(6000);  // Feb revenue
                expect(getCellValue('annual', 1, 3)).toBe(7000);  // Mar revenue

                // Annual total revenue
                expect(getCellValue('annual', 1, 4)).toBe(18000);
                expect(getCellValue('annual', 1, 5)).toBe(6000);  // average
                expect(getCellValue('annual', 1, 6)).toBe(7000);  // max
                expect(getCellValue('annual', 1, 7)).toBe(5000);  // min

                // Annual expenses
                expect(getCellValue('annual', 2, 1)).toBe(3000);
                expect(getCellValue('annual', 2, 2)).toBe(3500);
                expect(getCellValue('annual', 2, 3)).toBe(4000);
                expect(getCellValue('annual', 2, 4)).toBe(10500);

                // Annual profits (chain: annual → monthly profit → monthly revenue - expenses)
                expect(getCellValue('annual', 3, 1)).toBe(2000);
                expect(getCellValue('annual', 3, 2)).toBe(2500);
                expect(getCellValue('annual', 3, 3)).toBe(3000);
                expect(getCellValue('annual', 3, 4)).toBe(7500);

                // Margin = profit / revenue
                const janMargin = getCellValue('annual', 4, 1) as number;
                expect(janMargin).toBeCloseTo(2000 / 5000, 6);

                const totalMargin = getCellValue('annual', 4, 4) as number;
                expect(totalMargin).toBeCloseTo(7500 / 18000, 6);
            } finally {
                testBed.univer.dispose();
            }
        });
    });
});
