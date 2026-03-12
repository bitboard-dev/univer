import type { ICellData, IWorkbookData } from '@univerjs/core';
import { CellValueType } from '@univerjs/core';
import { DATA_SHEET_NAME } from './data-generator';

type CellRow = Record<number, ICellData>;

export interface IFormulaGroup {
    id: string;
    name: string;
    sheetId: string;
    formulas: string[];
    build: (rows: number) => IWorkbookData['sheets'][string];
}

function makeSheet(
    sheetId: string,
    name: string,
    rows: number,
    headers: string[],
    formulaRow: (r: number) => CellRow
): IWorkbookData['sheets'][string] {
    const cellData: Record<number, CellRow> = {};
    const headerRow: CellRow = {};
    for (let c = 0; c < headers.length; c++) {
        headerRow[c] = { v: headers[c], t: CellValueType.STRING };
    }
    cellData[0] = headerRow;

    for (let i = 0; i < rows; i++) {
        cellData[i + 1] = formulaRow(i + 2);
    }

    return {
        id: sheetId,
        name,
        rowCount: rows + 10,
        columnCount: headers.length,
        cellData,
    };
}

const D = DATA_SHEET_NAME;

function rangeCol(col: string, rows: number): string {
    return `${D}!$${col}$2:$${col}$${rows + 1}`;
}

function cell(col: string, r: number): string {
    return `${D}!${col}${r}`;
}

function absCell(col: string, r: number): string {
    return `${D}!$${col}$${r}`;
}

function rangeLocal(col: string, r1: number, r2: number): string {
    return `${D}!$${col}$${r1}:$${col}$${r2}`;
}

// ---------------------------------------------------------------------------
// Group 1: Range Aggregation
// ---------------------------------------------------------------------------
const rangeAggFormulas = [
    'SUM', 'AVERAGE', 'MAX', 'MIN', 'COUNT', 'COUNTA', 'COUNTBLANK',
    'PRODUCT', 'SUMSQ', 'AVEDEV', 'DEVSQ', 'GEOMEAN', 'HARMEAN',
    'KURT', 'SKEW', 'MEDIAN', 'MODE.SNGL', 'STDEV.S', 'STDEV.P',
    'VAR.S', 'VAR.P', 'LARGE', 'SMALL',
];

function buildRangeAgg(rows: number): IWorkbookData['sheets'][string] {
    const R = rangeCol('B', rows);
    const headers = rangeAggFormulas;
    return makeSheet('range-agg', 'RangeAgg', rows, headers, (r) => {
        const row: CellRow = {};
        const fns: string[] = [
            `=SUM(${R})`, `=AVERAGE(${R})`, `=MAX(${R})`, `=MIN(${R})`,
            `=COUNT(${R})`, `=COUNTA(${R})`, `=COUNTBLANK(${R})`,
            `=PRODUCT(${rangeLocal('C', 2, Math.min(r, 20))})`,
            `=SUMSQ(${R})`, `=AVEDEV(${R})`, `=DEVSQ(${R})`,
            `=GEOMEAN(${rangeLocal('B', 2, Math.min(r, 50))})`,
            `=HARMEAN(${rangeLocal('B', 2, Math.min(r, 50))})`,
            `=KURT(${R})`, `=SKEW(${R})`, `=MEDIAN(${R})`,
            `=MODE.SNGL(${rangeCol('D', rows)})`,
            `=STDEV.S(${R})`, `=STDEV.P(${R})`,
            `=VAR.S(${R})`, `=VAR.P(${R})`,
            `=LARGE(${R},${r - 1})`, `=SMALL(${R},${r - 1})`,
        ];
        fns.forEach((f, i) => { row[i] = { f }; });
        return row;
    });
}

// ---------------------------------------------------------------------------
// Group 2: Conditional Aggregation
// ---------------------------------------------------------------------------
const condAggFormulas = [
    'SUMIF', 'SUMIFS', 'COUNTIF', 'COUNTIFS',
    'AVERAGEIF', 'AVERAGEIFS', 'MAXIFS', 'MINIFS',
];

function buildCondAgg(rows: number): IWorkbookData['sheets'][string] {
    const amtR = rangeCol('B', rows);
    const catR = rangeCol('E', rows);
    const grpR = rangeCol('F', rows);
    return makeSheet('cond-agg', 'CondAgg', rows, condAggFormulas, (r) => ({
        0: { f: `=SUMIF(${catR},${cell('E', r)},${amtR})` },
        1: { f: `=SUMIFS(${amtR},${catR},${cell('E', r)},${grpR},${cell('F', r)})` },
        2: { f: `=COUNTIF(${catR},${cell('E', r)})` },
        3: { f: `=COUNTIFS(${catR},${cell('E', r)},${grpR},${cell('F', r)})` },
        4: { f: `=AVERAGEIF(${catR},${cell('E', r)},${amtR})` },
        5: { f: `=AVERAGEIFS(${amtR},${catR},${cell('E', r)},${grpR},${cell('F', r)})` },
        6: { f: `=MAXIFS(${amtR},${catR},${cell('E', r)})` },
        7: { f: `=MINIFS(${amtR},${catR},${cell('E', r)})` },
    }));
}

// ---------------------------------------------------------------------------
// Group 3: Lookup
// ---------------------------------------------------------------------------
const lookupFormulas = [
    'VLOOKUP', 'INDEX+MATCH', 'XLOOKUP', 'MATCH', 'LOOKUP',
    'CHOOSE', 'ADDRESS', 'ROW', 'COLUMN', 'ROWS', 'COLUMNS',
];

function buildLookup(rows: number): IWorkbookData['sheets'][string] {
    const tbl = `${D}!$A$2:$B$${rows + 1}`;
    const colA = rangeCol('A', rows);
    const colB = rangeCol('B', rows);
    return makeSheet('lookup', 'Lookup', rows, lookupFormulas, (r) => ({
        0: { f: `=VLOOKUP(${cell('A', r)},${tbl},2,FALSE)` },
        1: { f: `=INDEX(${colB},MATCH(${cell('A', r)},${colA},0))` },
        2: { f: `=XLOOKUP(${cell('A', r)},${colA},${colB})` },
        3: { f: `=MATCH(${cell('A', r)},${colA},0)` },
        4: { f: `=LOOKUP(${cell('A', r)},${colA},${colB})` },
        5: { f: `=CHOOSE(${cell('D', r)},"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")` },
        6: { f: `=ADDRESS(${r},1)` },
        7: { f: `=ROW(${cell('A', r)})` },
        8: { f: `=COLUMN(${cell('B', r)})` },
        9: { f: `=ROWS(${tbl})` },
        10: { f: `=COLUMNS(${tbl})` },
    }));
}

// ---------------------------------------------------------------------------
// Group 4: Statistical Ranking
// ---------------------------------------------------------------------------
const statRankFormulas = [
    'RANK.EQ', 'RANK.AVG', 'LARGE_fixed', 'SMALL_fixed',
    'PERCENTILE.INC', 'PERCENTILE.EXC',
    'PERCENTRANK.INC', 'PERCENTRANK.EXC',
    'QUARTILE.INC', 'QUARTILE.EXC',
];

function buildStatRank(rows: number): IWorkbookData['sheets'][string] {
    const R = rangeCol('B', rows);
    return makeSheet('stat-rank', 'StatRank', rows, statRankFormulas, (r) => ({
        0: { f: `=RANK.EQ(${cell('B', r)},${R})` },
        1: { f: `=RANK.AVG(${cell('B', r)},${R})` },
        2: { f: `=LARGE(${R},5)` },
        3: { f: `=SMALL(${R},5)` },
        4: { f: `=PERCENTILE.INC(${R},0.75)` },
        5: { f: `=PERCENTILE.EXC(${R},0.75)` },
        6: { f: `=PERCENTRANK.INC(${R},${cell('B', r)})` },
        7: { f: `=PERCENTRANK.EXC(${R},${cell('B', r)})` },
        8: { f: `=QUARTILE.INC(${R},2)` },
        9: { f: `=QUARTILE.EXC(${R},2)` },
    }));
}

// ---------------------------------------------------------------------------
// Group 5: Row-Local Math
// ---------------------------------------------------------------------------
const rowMathFormulas = [
    'ABS', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'INT', 'TRUNC',
    'MOD', 'POWER', 'SQRT', 'SQRTPI', 'SIGN', 'EVEN', 'ODD',
    'CEILING', 'FLOOR', 'MROUND',
    'SIN', 'COS', 'TAN', 'ASIN', 'ACOS', 'ATAN',
    'SINH', 'COSH', 'TANH',
    'LN', 'LOG', 'LOG10', 'EXP',
    'FACT', 'COMBIN', 'COMBINA', 'GCD', 'LCM',
    'DEGREES', 'RADIANS', 'PI_mult',
    'ROMAN', 'ARABIC',
    'BASE', 'DECIMAL',
    'RAND', 'RANDBETWEEN',
    'QUOTIENT', 'PRODUCT_pair',
];

function buildRowMath(rows: number): IWorkbookData['sheets'][string] {
    return makeSheet('row-math', 'RowMath', rows, rowMathFormulas, (r) => ({
        0: { f: `=ABS(${cell('B', r)}-50000)` },
        1: { f: `=ROUND(${cell('B', r)},2)` },
        2: { f: `=ROUNDUP(${cell('B', r)},0)` },
        3: { f: `=ROUNDDOWN(${cell('B', r)},0)` },
        4: { f: `=INT(${cell('B', r)})` },
        5: { f: `=TRUNC(${cell('B', r)},1)` },
        6: { f: `=MOD(${cell('A', r)},7)` },
        7: { f: `=POWER(${cell('C', r)},2)` },
        8: { f: `=SQRT(${cell('B', r)})` },
        9: { f: `=SQRTPI(${cell('C', r)})` },
        10: { f: `=SIGN(${cell('B', r)}-50000)` },
        11: { f: `=EVEN(${cell('B', r)})` },
        12: { f: `=ODD(${cell('B', r)})` },
        13: { f: `=CEILING(${cell('B', r)},100)` },
        14: { f: `=FLOOR(${cell('B', r)},100)` },
        15: { f: `=MROUND(${cell('B', r)},50)` },
        16: { f: `=SIN(${cell('C', r)}*PI())` },
        17: { f: `=COS(${cell('C', r)}*PI())` },
        18: { f: `=TAN(${cell('C', r)})` },
        19: { f: `=ASIN(${cell('C', r)})` },
        20: { f: `=ACOS(${cell('C', r)})` },
        21: { f: `=ATAN(${cell('B', r)})` },
        22: { f: `=SINH(${cell('C', r)})` },
        23: { f: `=COSH(${cell('C', r)})` },
        24: { f: `=TANH(${cell('C', r)})` },
        25: { f: `=LN(${cell('B', r)}+1)` },
        26: { f: `=LOG(${cell('B', r)}+1,10)` },
        27: { f: `=LOG10(${cell('B', r)}+1)` },
        28: { f: `=EXP(${cell('C', r)})` },
        29: { f: `=FACT(${cell('D', r)})` },
        30: { f: `=COMBIN(${cell('D', r)}+5,${cell('D', r)})` },
        31: { f: `=COMBINA(${cell('D', r)},3)` },
        32: { f: `=GCD(${cell('A', r)},${cell('J', r)})` },
        33: { f: `=LCM(${cell('A', r)},${cell('J', r)})` },
        34: { f: `=DEGREES(${cell('C', r)})` },
        35: { f: `=RADIANS(${cell('B', r)})` },
        36: { f: `=PI()*${cell('C', r)}` },
        37: { f: `=ROMAN(${cell('J', r)})` },
        38: { f: `=ARABIC("XIV")` },
        39: { f: `=BASE(${cell('A', r)},16)` },
        40: { f: `=DECIMAL(${cell('K', r)},16)` },
        41: { f: `=RAND()` },
        42: { f: `=RANDBETWEEN(1,1000)` },
        43: { f: `=QUOTIENT(${cell('B', r)},${cell('J', r)}+1)` },
        44: { f: `=PRODUCT(${cell('C', r)},${cell('J', r)})` },
    }));
}

// ---------------------------------------------------------------------------
// Group 6: Row-Local Text
// ---------------------------------------------------------------------------
const rowTextFormulas = [
    'LEFT', 'RIGHT', 'MID', 'LEN', 'UPPER', 'LOWER', 'PROPER',
    'TRIM', 'CLEAN', 'CONCATENATE', 'CONCAT', 'TEXTJOIN',
    'SUBSTITUTE', 'REPLACE', 'REPT',
    'FIND', 'SEARCH', 'EXACT',
    'TEXT', 'VALUE', 'DOLLAR', 'FIXED',
    'CHAR', 'CODE', 'UNICODE', 'UNICHAR',
    'T', 'NUMBERVALUE',
];

function buildRowText(rows: number): IWorkbookData['sheets'][string] {
    return makeSheet('row-text', 'RowText', rows, rowTextFormulas, (r) => ({
        0: { f: `=LEFT(${cell('I', r)},5)` },
        1: { f: `=RIGHT(${cell('I', r)},4)` },
        2: { f: `=MID(${cell('I', r)},3,5)` },
        3: { f: `=LEN(${cell('I', r)})` },
        4: { f: `=UPPER(${cell('I', r)})` },
        5: { f: `=LOWER(${cell('I', r)})` },
        6: { f: `=PROPER(${cell('I', r)})` },
        7: { f: `=TRIM(${cell('I', r)})` },
        8: { f: `=CLEAN(${cell('I', r)})` },
        9: { f: `=CONCATENATE(${cell('E', r)}," - ",${cell('F', r)})` },
        10: { f: `=CONCAT(${cell('E', r)},${cell('F', r)})` },
        11: { f: `=TEXTJOIN("-",TRUE,${cell('E', r)},${cell('F', r)})` },
        12: { f: `=SUBSTITUTE(${cell('I', r)}," ","-")` },
        13: { f: `=REPLACE(${cell('I', r)},1,3,"XXX")` },
        14: { f: `=REPT(${cell('E', r)},2)` },
        15: { f: `=FIND(" ",${cell('I', r)})` },
        16: { f: `=SEARCH(" ",${cell('I', r)})` },
        17: { f: `=EXACT(${cell('E', r)},${cell('E', r)})` },
        18: { f: `=TEXT(${cell('B', r)},"#,##0.00")` },
        19: { f: `=VALUE("123")` },
        20: { f: `=DOLLAR(${cell('B', r)},2)` },
        21: { f: `=FIXED(${cell('B', r)},2)` },
        22: { f: `=CHAR(65+MOD(${cell('A', r)},26))` },
        23: { f: `=CODE(${cell('E', r)})` },
        24: { f: `=UNICODE(${cell('E', r)})` },
        25: { f: `=UNICHAR(65+MOD(${cell('A', r)},26))` },
        26: { f: `=T(${cell('E', r)})` },
        27: { f: `=NUMBERVALUE("1,234.56",".",",")` },
    }));
}

// ---------------------------------------------------------------------------
// Group 7: Row-Local Logical
// ---------------------------------------------------------------------------
const rowLogicalFormulas = [
    'IF', 'AND', 'OR', 'NOT', 'XOR',
    'IFS', 'SWITCH', 'IFERROR', 'IFNA',
];

function buildRowLogical(rows: number): IWorkbookData['sheets'][string] {
    return makeSheet('row-logical', 'RowLogical', rows, rowLogicalFormulas, (r) => ({
        0: { f: `=IF(${cell('B', r)}>50000,"high","low")` },
        1: { f: `=AND(${cell('B', r)}>10000,${cell('C', r)}>0.5)` },
        2: { f: `=OR(${cell('D', r)}>10,${cell('G', r)})` },
        3: { f: `=NOT(${cell('G', r)})` },
        4: { f: `=XOR(${cell('G', r)},${cell('B', r)}>50000)` },
        5: { f: `=IFS(${cell('B', r)}>75000,"top",${cell('B', r)}>25000,"mid",TRUE,"low")` },
        6: { f: `=SWITCH(${cell('D', r)},1,"Jan",2,"Feb",3,"Mar","other")` },
        7: { f: `=IFERROR(${cell('B', r)}/(${cell('A', r)}-${cell('A', r)}+1),0)` },
        8: { f: `=IFNA(VLOOKUP(${cell('A', r)},${D}!$A$2:$B$${rows + 1},2,FALSE),"missing")` },
    }));
}

// ---------------------------------------------------------------------------
// Group 8: Row-Local Date
// ---------------------------------------------------------------------------
const rowDateFormulas = [
    'YEAR', 'MONTH', 'DAY', 'DATE', 'DATEDIF',
    'DAYS', 'DAYS360', 'EDATE', 'EOMONTH',
    'WEEKDAY', 'WEEKNUM', 'ISOWEEKNUM',
    'NETWORKDAYS', 'WORKDAY',
    'DATEVALUE', 'HOUR', 'MINUTE', 'SECOND',
];

function buildRowDate(rows: number): IWorkbookData['sheets'][string] {
    return makeSheet('row-date', 'RowDate', rows, rowDateFormulas, (r) => ({
        0: { f: `=YEAR(${cell('H', r)})` },
        1: { f: `=MONTH(${cell('H', r)})` },
        2: { f: `=DAY(${cell('H', r)})` },
        3: { f: `=DATE(YEAR(${cell('H', r)}),MONTH(${cell('H', r)}),1)` },
        4: { f: `=DATEDIF(${absCell('H', 2)},${cell('H', r)},"d")` },
        5: { f: `=DAYS(${cell('H', r)},${absCell('H', 2)})` },
        6: { f: `=DAYS360(${absCell('H', 2)},${cell('H', r)})` },
        7: { f: `=EDATE(${cell('H', r)},3)` },
        8: { f: `=EOMONTH(${cell('H', r)},0)` },
        9: { f: `=WEEKDAY(${cell('H', r)})` },
        10: { f: `=WEEKNUM(${cell('H', r)})` },
        11: { f: `=ISOWEEKNUM(${cell('H', r)})` },
        12: { f: `=NETWORKDAYS(${absCell('H', 2)},${cell('H', r)})` },
        13: { f: `=WORKDAY(${cell('H', r)},10)` },
        14: { f: `=DATEVALUE("2024-06-15")` },
        15: { f: `=HOUR(0.75)` },
        16: { f: `=MINUTE(0.75)` },
        17: { f: `=SECOND(0.7512)` },
    }));
}

// ---------------------------------------------------------------------------
// Group 9: Row-Local Info
// ---------------------------------------------------------------------------
const rowInfoFormulas = [
    'ISNUMBER', 'ISTEXT', 'ISBLANK', 'ISERROR', 'ISLOGICAL',
    'ISEVEN', 'ISODD', 'ISNA', 'ISERR', 'ISNONTEXT',
    'N', 'TYPE',
];

function buildRowInfo(rows: number): IWorkbookData['sheets'][string] {
    return makeSheet('row-info', 'RowInfo', rows, rowInfoFormulas, (r) => ({
        0: { f: `=ISNUMBER(${cell('B', r)})` },
        1: { f: `=ISTEXT(${cell('E', r)})` },
        2: { f: `=ISBLANK(${cell('B', r)})` },
        3: { f: `=ISERROR(1/0)` },
        4: { f: `=ISLOGICAL(${cell('G', r)})` },
        5: { f: `=ISEVEN(${cell('A', r)})` },
        6: { f: `=ISODD(${cell('A', r)})` },
        7: { f: `=ISNA(NA())` },
        8: { f: `=ISERR(1/0)` },
        9: { f: `=ISNONTEXT(${cell('B', r)})` },
        10: { f: `=N(${cell('G', r)})` },
        11: { f: `=TYPE(${cell('B', r)})` },
    }));
}

// ---------------------------------------------------------------------------
// Group 10: Cross-Range
// ---------------------------------------------------------------------------
const crossRangeFormulas = [
    'SUMPRODUCT', 'SUMX2MY2', 'SUMX2PY2', 'SUMXMY2',
];

function buildCrossRange(rows: number): IWorkbookData['sheets'][string] {
    const RB = rangeCol('B', rows);
    const RC = rangeCol('C', rows);
    const RE = rangeCol('E', rows);
    return makeSheet('cross-range', 'CrossRange', rows, crossRangeFormulas, (r) => ({
        0: { f: `=SUMPRODUCT((${RE}=${cell('E', r)})*${RB})` },
        1: { f: `=SUMX2MY2(${rangeLocal('B', 2, 11)},${rangeLocal('C', 2, 11)})` },
        2: { f: `=SUMX2PY2(${rangeLocal('B', 2, 11)},${rangeLocal('C', 2, 11)})` },
        3: { f: `=SUMXMY2(${rangeLocal('B', 2, 11)},${rangeLocal('C', 2, 11)})` },
    }));
}

// ---------------------------------------------------------------------------
// Group 11: Financial
// ---------------------------------------------------------------------------
const financialFormulas = [
    'PMT', 'FV', 'PV', 'NPER', 'RATE',
    'IPMT', 'PPMT', 'CUMIPMT', 'CUMPRINC',
    'NPV', 'IRR', 'MIRR',
    'DB', 'DDB', 'SLN', 'SYD',
    'EFFECT', 'NOMINAL', 'PDURATION', 'RRI',
    'FVSCHEDULE', 'ISPMT', 'DISC',
    'DOLLARDE', 'DOLLARFR', 'DURATION',
];

function buildFinancial(rows: number): IWorkbookData['sheets'][string] {
    return makeSheet('financial', 'Financial', rows, financialFormulas, (r) => ({
        0: { f: `=PMT(0.05/12,${cell('J', r)},-${cell('B', r)})` },
        1: { f: `=FV(0.05/12,${cell('J', r)},-1000)` },
        2: { f: `=PV(0.05/12,${cell('J', r)},0,-${cell('B', r)})` },
        3: { f: `=NPER(0.05/12,-500,${cell('B', r)})` },
        4: { f: `=RATE(${cell('J', r)},-500,${cell('B', r)})` },
        5: { f: `=IPMT(0.05/12,1,${cell('J', r)},-${cell('B', r)})` },
        6: { f: `=PPMT(0.05/12,1,${cell('J', r)},-${cell('B', r)})` },
        7: { f: `=CUMIPMT(0.05/12,${cell('J', r)},-${cell('B', r)},1,12,0)` },
        8: { f: `=CUMPRINC(0.05/12,${cell('J', r)},-${cell('B', r)},1,12,0)` },
        9: { f: `=NPV(0.1,${absCell('B', 2)},${absCell('B', 3)},${absCell('B', 4)},${absCell('B', 5)})` },
        10: { f: `=IRR({-1000,300,300,300,300})` },
        11: { f: `=MIRR({-1000,300,300,300,300},0.1,0.12)` },
        12: { f: `=DB(${cell('B', r)},1000,${cell('J', r)},1)` },
        13: { f: `=DDB(${cell('B', r)},1000,${cell('J', r)},1)` },
        14: { f: `=SLN(${cell('B', r)},1000,${cell('J', r)})` },
        15: { f: `=SYD(${cell('B', r)},1000,${cell('J', r)},1)` },
        16: { f: `=EFFECT(0.1,4)` },
        17: { f: `=NOMINAL(0.1,4)` },
        18: { f: `=PDURATION(0.05,1000,${cell('B', r)})` },
        19: { f: `=RRI(${cell('J', r)},1000,${cell('B', r)})` },
        20: { f: `=FVSCHEDULE(1000,{0.05,0.06,0.07})` },
        21: { f: `=ISPMT(0.05/12,1,${cell('J', r)},${cell('B', r)})` },
        22: { f: `=DISC(44927,45292,95,100)` },
        23: { f: `=DOLLARDE(1.02,16)` },
        24: { f: `=DOLLARFR(1.125,16)` },
        25: { f: `=DURATION(44927,46388,0.05,0.06,2)` },
    }));
}

// ---------------------------------------------------------------------------
// Group 12: Engineering
// ---------------------------------------------------------------------------
const engineeringFormulas = [
    'BIN2DEC', 'BIN2HEX', 'BIN2OCT',
    'DEC2BIN', 'DEC2HEX', 'DEC2OCT',
    'HEX2BIN', 'HEX2DEC', 'HEX2OCT',
    'OCT2BIN', 'OCT2DEC', 'OCT2HEX',
    'DELTA', 'GESTEP',
    'COMPLEX', 'IMABS', 'IMAGINARY', 'IMREAL',
    'IMSUM', 'IMSUB', 'IMPRODUCT', 'IMDIV',
    'IMCONJUGATE', 'IMARGUMENT',
    'BITAND', 'BITOR', 'BITXOR', 'BITLSHIFT', 'BITRSHIFT',
];

function buildEngineering(rows: number): IWorkbookData['sheets'][string] {
    return makeSheet('engineering', 'Engineering', rows, engineeringFormulas, (r) => ({
        0: { f: `=BIN2DEC(${cell('L', r)})` },
        1: { f: `=BIN2HEX(${cell('L', r)})` },
        2: { f: `=BIN2OCT(${cell('L', r)})` },
        3: { f: `=DEC2BIN(MOD(${cell('A', r)},256),8)` },
        4: { f: `=DEC2HEX(${cell('A', r)})` },
        5: { f: `=DEC2OCT(${cell('A', r)})` },
        6: { f: `=HEX2BIN(DEC2HEX(MOD(${cell('A', r)},256)),8)` },
        7: { f: `=HEX2DEC(${cell('K', r)})` },
        8: { f: `=HEX2OCT(${cell('K', r)})` },
        9: { f: `=OCT2BIN(DEC2OCT(MOD(${cell('A', r)},256)),8)` },
        10: { f: `=OCT2DEC(DEC2OCT(${cell('A', r)}))` },
        11: { f: `=OCT2HEX(DEC2OCT(${cell('A', r)}))` },
        12: { f: `=DELTA(${cell('A', r)},${cell('A', r)})` },
        13: { f: `=GESTEP(${cell('B', r)},50000)` },
        14: { f: `=COMPLEX(${cell('C', r)},${cell('C', r)})` },
        15: { f: `=IMABS(COMPLEX(3,4))` },
        16: { f: `=IMAGINARY(COMPLEX(${cell('C', r)},${cell('C', r)}))` },
        17: { f: `=IMREAL(COMPLEX(${cell('C', r)},${cell('C', r)}))` },
        18: { f: `=IMSUM(COMPLEX(1,2),COMPLEX(3,4))` },
        19: { f: `=IMSUB(COMPLEX(3,4),COMPLEX(1,2))` },
        20: { f: `=IMPRODUCT(COMPLEX(1,2),COMPLEX(3,4))` },
        21: { f: `=IMDIV(COMPLEX(3,4),COMPLEX(1,2))` },
        22: { f: `=IMCONJUGATE(COMPLEX(${cell('C', r)},${cell('C', r)}))` },
        23: { f: `=IMARGUMENT(COMPLEX(3,4))` },
        24: { f: `=BITAND(${cell('A', r)},255)` },
        25: { f: `=BITOR(${cell('A', r)},255)` },
        26: { f: `=BITXOR(${cell('A', r)},255)` },
        27: { f: `=BITLSHIFT(MOD(${cell('A', r)},128),1)` },
        28: { f: `=BITRSHIFT(${cell('A', r)},1)` },
    }));
}

// ---------------------------------------------------------------------------
// Group 13: Database
// ---------------------------------------------------------------------------
const databaseFormulas = [
    'DSUM', 'DAVERAGE', 'DCOUNT', 'DCOUNTA',
    'DMAX', 'DMIN', 'DPRODUCT',
    'DSTDEV', 'DVAR',
];

function buildDatabase(rows: number): IWorkbookData['sheets'][string] {
    const db = `${D}!$A$1:$F$${rows + 1}`;
    const headers = databaseFormulas;
    const sheetId = 'database';

    const cellData: Record<number, CellRow> = {};
    const headerRow: CellRow = {};
    for (let c = 0; c < headers.length; c++) {
        headerRow[c] = { v: headers[c], t: CellValueType.STRING };
    }
    cellData[0] = headerRow;

    // Criteria block: rows 1..2 hold {header, value} for each column
    // We use a criteria area at columns headers.length+0..headers.length+1
    const critCol = headers.length;
    cellData[0][critCol] = { v: 'Category', t: CellValueType.STRING };
    cellData[1] = cellData[1] || {};
    cellData[1][critCol] = { v: 'Alpha', t: CellValueType.STRING };

    const critRange = `Database!$${colLetter(critCol)}$1:$${colLetter(critCol)}$2`;

    for (let i = 0; i < rows; i++) {
        const r = i + 1;
        if (!cellData[r]) cellData[r] = {};
        cellData[r][0] = { f: `=DSUM(${db},"Amount",${critRange})` };
        cellData[r][1] = { f: `=DAVERAGE(${db},"Amount",${critRange})` };
        cellData[r][2] = { f: `=DCOUNT(${db},"Amount",${critRange})` };
        cellData[r][3] = { f: `=DCOUNTA(${db},"Category",${critRange})` };
        cellData[r][4] = { f: `=DMAX(${db},"Amount",${critRange})` };
        cellData[r][5] = { f: `=DMIN(${db},"Amount",${critRange})` };
        cellData[r][6] = { f: `=DPRODUCT(${db},"Rate",${critRange})` };
        cellData[r][7] = { f: `=DSTDEV(${db},"Amount",${critRange})` };
        cellData[r][8] = { f: `=DVAR(${db},"Amount",${critRange})` };
    }

    return {
        id: sheetId,
        name: 'Database',
        rowCount: rows + 10,
        columnCount: critCol + 2,
        cellData,
    };
}

// ---------------------------------------------------------------------------
// Group 14: Compatibility (additional stats with old names)
// ---------------------------------------------------------------------------
const compatFormulas = [
    'BETADIST', 'BINOMDIST', 'CHIDIST',
    'EXPONDIST', 'FDIST', 'GAMMADIST',
    'NORMDIST', 'NORMINV', 'NORMSDIST', 'NORMSINV',
    'POISSON', 'STDEV_c', 'STDEVP_c', 'VAR_c', 'VARP_c',
    'WEIBULL', 'CONFIDENCE', 'COVAR',
    'CORREL', 'FISHER', 'FISHERINV',
    'STANDARDIZE', 'TRIMMEAN_c',
];

function buildCompat(rows: number): IWorkbookData['sheets'][string] {
    const R = rangeCol('B', rows);
    const RC = rangeCol('C', rows);
    return makeSheet('compat', 'Compat', rows, compatFormulas, (r) => ({
        0: { f: `=BETADIST(${cell('C', r)},2,5)` },
        1: { f: `=BINOMDIST(3,10,0.5,FALSE)` },
        2: { f: `=CHIDIST(${cell('C', r)}*10,5)` },
        3: { f: `=EXPONDIST(${cell('C', r)},1,TRUE)` },
        4: { f: `=FDIST(${cell('C', r)}*5,5,10)` },
        5: { f: `=GAMMADIST(${cell('C', r)}*10,2,1,TRUE)` },
        6: { f: `=NORMDIST(${cell('B', r)},50000,15000,TRUE)` },
        7: { f: `=NORMINV(${cell('C', r)},50000,15000)` },
        8: { f: `=NORMSDIST(${cell('C', r)})` },
        9: { f: `=NORMSINV(${cell('C', r)}*0.98+0.01)` },
        10: { f: `=POISSON(${cell('D', r)},5,FALSE)` },
        11: { f: `=STDEV(${R})` },
        12: { f: `=STDEVP(${R})` },
        13: { f: `=VAR(${R})` },
        14: { f: `=VARP(${R})` },
        15: { f: `=WEIBULL(${cell('C', r)},2,1,TRUE)` },
        16: { f: `=CONFIDENCE(0.05,15000,100)` },
        17: { f: `=COVAR(${rangeLocal('B', 2, 11)},${rangeLocal('C', 2, 11)})` },
        18: { f: `=CORREL(${rangeLocal('B', 2, 51)},${rangeLocal('C', 2, 51)})` },
        19: { f: `=FISHER(${cell('C', r)}*0.98)` },
        20: { f: `=FISHERINV(${cell('C', r)})` },
        21: { f: `=STANDARDIZE(${cell('B', r)},50000,15000)` },
        22: { f: `=TRIMMEAN(${R},0.1)` },
    }));
}

function colLetter(idx: number): string {
    let s = '';
    let n = idx;
    while (n >= 0) {
        s = String.fromCharCode(65 + (n % 26)) + s;
        n = Math.floor(n / 26) - 1;
    }
    return s;
}

// ---------------------------------------------------------------------------
// Export all groups
// ---------------------------------------------------------------------------
export const FORMULA_GROUPS: IFormulaGroup[] = [
    { id: 'range-agg', name: 'Range Aggregation', sheetId: 'range-agg', formulas: rangeAggFormulas, build: buildRangeAgg },
    { id: 'cond-agg', name: 'Conditional Aggregation', sheetId: 'cond-agg', formulas: condAggFormulas, build: buildCondAgg },
    { id: 'lookup', name: 'Lookup', sheetId: 'lookup', formulas: lookupFormulas, build: buildLookup },
    { id: 'stat-rank', name: 'Statistical Ranking', sheetId: 'stat-rank', formulas: statRankFormulas, build: buildStatRank },
    { id: 'row-math', name: 'Row-Local Math', sheetId: 'row-math', formulas: rowMathFormulas, build: buildRowMath },
    { id: 'row-text', name: 'Row-Local Text', sheetId: 'row-text', formulas: rowTextFormulas, build: buildRowText },
    { id: 'row-logical', name: 'Row-Local Logical', sheetId: 'row-logical', formulas: rowLogicalFormulas, build: buildRowLogical },
    { id: 'row-date', name: 'Row-Local Date', sheetId: 'row-date', formulas: rowDateFormulas, build: buildRowDate },
    { id: 'row-info', name: 'Row-Local Info', sheetId: 'row-info', formulas: rowInfoFormulas, build: buildRowInfo },
    { id: 'cross-range', name: 'Cross-Range', sheetId: 'cross-range', formulas: crossRangeFormulas, build: buildCrossRange },
    { id: 'financial', name: 'Financial', sheetId: 'financial', formulas: financialFormulas, build: buildFinancial },
    { id: 'engineering', name: 'Engineering', sheetId: 'engineering', formulas: engineeringFormulas, build: buildEngineering },
    { id: 'database', name: 'Database', sheetId: 'database', formulas: databaseFormulas, build: buildDatabase },
    { id: 'compat', name: 'Compatibility/Stats', sheetId: 'compat', formulas: compatFormulas, build: buildCompat },
];
