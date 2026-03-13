import type { IWorkbookData } from '@univerjs/core';
import { CellValueType } from '@univerjs/core';

export interface IBenchmarkOptions {
    rows: number;
    unitId?: string;
}

const CATEGORIES_20 = [
    'Alpha', 'Bravo', 'Charlie', 'Delta', 'Echo',
    'Foxtrot', 'Golf', 'Hotel', 'India', 'Juliet',
    'Kilo', 'Lima', 'Mike', 'November', 'Oscar',
    'Papa', 'Quebec', 'Romeo', 'Sierra', 'Tango',
];

const GROUPS_5 = ['GroupA', 'GroupB', 'GroupC', 'GroupD', 'GroupE'];

const WORDS = [
    'apple', 'banana', 'cherry', 'delta', 'eagle',
    'falcon', 'grape', 'honey', 'ivory', 'jade',
    'kite', 'lemon', 'mango', 'nutmeg', 'olive',
];

function seededRandom(seed: number): () => number {
    let s = seed;
    return () => {
        s = (s * 1664525 + 1013904223) & 0x7fffffff;
        return s / 0x7fffffff;
    };
}

function randomText(rand: () => number): string {
    const w1 = WORDS[Math.floor(rand() * WORDS.length)];
    const w2 = WORDS[Math.floor(rand() * WORDS.length)];
    const w3 = WORDS[Math.floor(rand() * WORDS.length)];
    return `${w1} ${w2} ${w3}`;
}

function toBinaryString(n: number): string {
    return (n & 0xff).toString(2).padStart(8, '0');
}

function toHexString(n: number): string {
    return (n & 0xffff).toString(16).toUpperCase().padStart(4, '0');
}

export const DATA_SHEET_ID = 'data';
export const DATA_SHEET_NAME = 'Data';

export function buildDataSheet(opts: IBenchmarkOptions): IWorkbookData['sheets'][string] {
    const { rows } = opts;
    const rand = seededRandom(42);
    const cellData: IWorkbookData['sheets'][string]['cellData'] = {};

    cellData[0] = {
        0: { v: 'ID', t: CellValueType.STRING },
        1: { v: 'Amount', t: CellValueType.STRING },
        2: { v: 'Rate', t: CellValueType.STRING },
        3: { v: 'Month', t: CellValueType.STRING },
        4: { v: 'Category', t: CellValueType.STRING },
        5: { v: 'Group', t: CellValueType.STRING },
        6: { v: 'Flag', t: CellValueType.STRING },
        7: { v: 'DateVal', t: CellValueType.STRING },
        8: { v: 'Words', t: CellValueType.STRING },
        9: { v: 'Periods', t: CellValueType.STRING },
        10: { v: 'Hex', t: CellValueType.STRING },
        11: { v: 'Binary', t: CellValueType.STRING },
    };

    const baseDate = 44927; // 2023-01-01 serial

    for (let i = 0; i < rows; i++) {
        const row = i + 1;
        const amount = Math.round(rand() * 100000 * 100) / 100;
        const rate = Math.round(rand() * 10000) / 10000;
        const month = Math.floor(rand() * 12) + 1;
        const cat = CATEGORIES_20[Math.floor(rand() * CATEGORIES_20.length)];
        const grp = GROUPS_5[Math.floor(rand() * GROUPS_5.length)];
        const flag = rand() > 0.5;
        const dateVal = baseDate + Math.floor(rand() * 1095);
        const words = randomText(rand);
        const periods = Math.floor(rand() * 99) + 2;
        const hexVal = toHexString(Math.floor(rand() * 65536));
        const binVal = toBinaryString(Math.floor(rand() * 256));

        cellData[row] = {
            0: { v: row, t: CellValueType.NUMBER },
            1: { v: amount, t: CellValueType.NUMBER },
            2: { v: rate, t: CellValueType.NUMBER },
            3: { v: month, t: CellValueType.NUMBER },
            4: { v: cat, t: CellValueType.STRING },
            5: { v: grp, t: CellValueType.STRING },
            6: { v: flag, t: CellValueType.BOOLEAN },
            7: { v: dateVal, t: CellValueType.NUMBER },
            8: { v: words, t: CellValueType.STRING },
            9: { v: periods, t: CellValueType.NUMBER },
            10: { v: hexVal, t: CellValueType.STRING },
            11: { v: binVal, t: CellValueType.STRING },
        };
    }

    return {
        id: DATA_SHEET_ID,
        name: DATA_SHEET_NAME,
        rowCount: rows + 10,
        columnCount: 12,
        cellData,
    };
}
