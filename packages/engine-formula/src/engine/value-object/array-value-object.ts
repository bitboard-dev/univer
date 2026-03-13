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

import type { Nullable } from '@univerjs/core';
import type { callbackMapFnType, IArrayValueObject } from './base-value-object';

import { isRealNum } from '@univerjs/core';
import { BooleanValue, isTypedArrayOptimizationEnabled } from '../../basics/common';
import { ERROR_TYPE_SET, ErrorType } from '../../basics/error-type';
import { CELL_INVERTED_INDEX_CACHE } from '../../basics/inverted-index-cache';
import { regexTestArrayValue } from '../../basics/regex';
import { compareToken } from '../../basics/token';
import { ArrayBinarySearchType, ArrayOrderSearchType, getCompare } from '../utils/compare';
import { stringIsNumberPattern } from '../utils/numfmt-kit';
import { BaseValueObject, ErrorValueObject } from './base-value-object';
import { BooleanValueObject, createBooleanValueObjectByRawValue, createNumberValueObjectByRawValue, createStringValueObjectByRawValue, NullValueObject, NumberValueObject, StringValueObject } from './primitive-object';

enum BatchOperatorType {
    MINUS,
    PLUS,
    MULTIPLY,
    DIVIDED,
    MOD,
    COMPARE,
    CONCATENATE_FRONT,
    CONCATENATE_BACK,
    POW,
    ROUND,
    FLOOR,
    CEIL,
    ATAN2,
}

enum ArrayCalculateType {
    PRODUCT,
    ROW,
    COLUMN,
    SINGLE,
}

export function fromObjectToString(array: IArrayValueObject) {
    return '';
}

export function transformToValueObject(array: Array<Array<number | string | boolean | null>> = [], isIgnoreNumberPattern: boolean = false) {
    const arrayValueList: BaseValueObject[][] = [];

    for (let r = 0; r < array.length; r++) {
        const row = array[r];

        if (arrayValueList[r] == null) {
            arrayValueList[r] = [];
        }

        for (let c = 0; c < row.length; c++) {
            const cell = row[c];

            arrayValueList[r][c] = ValueObjectFactory.create(cell, isIgnoreNumberPattern);
        }
    }

    return arrayValueList;
}

export function transformToValue(array: Nullable<BaseValueObject>[][] = []) {
    const arrayValueList: Array<Array<string | number | boolean | null>> = [];

    for (let r = 0; r < array.length; r++) {
        const row = array[r];

        if (arrayValueList[r] == null) {
            arrayValueList[r] = [];
        }

        for (let c = 0; c < row.length; c++) {
            const cell = row[c];

            if (cell == null) {
                arrayValueList[r][c] = null;
            } else if (cell.isError()) {
                arrayValueList[r][c] = (cell as ErrorValueObject).getErrorType();
            } else {
                arrayValueList[r][c] = (cell as BaseValueObject).getValue();
            }
        }
    }

    return arrayValueList;
}

export class ArrayValueObject extends BaseValueObject {
    /**
     * Create an array value object based on the string or IArrayValueObject data.
     * @param rawValue
     * @returns
     */
    static create(rawValue: string | IArrayValueObject) {
        return new ArrayValueObject(rawValue);
    }

    /**
     * Create an array value object based on the array data.
     * @param array
     * @returns
     */
    static createByArray(array: Array<Array<number | string | boolean | null>>) {
        const calculateValueList = transformToValueObject(array);
        const arrayValueObjectData: IArrayValueObject = {
            calculateValueList,
            rowCount: array.length,
            columnCount: array[0].length || 0,
            unitId: '',
            sheetId: '',
            row: -1,
            column: -1,
        };

        return new ArrayValueObject(arrayValueObjectData);
    }

    static createNumberArray(
        data: Float64Array,
        rowCount: number,
        columnCount: number,
        unitId: string = '',
        sheetId: string = '',
        row: number = -1,
        column: number = -1
    ): ArrayValueObject {
        if (!isTypedArrayOptimizationEnabled()) {
            const calculateValueList = ArrayValueObject._typedArrayToValueObjects(data, rowCount, columnCount);
            return new ArrayValueObject({ calculateValueList, rowCount, columnCount, unitId, sheetId, row, column });
        }
        const obj = new ArrayValueObject({ calculateValueList: [], rowCount, columnCount, unitId, sheetId, row, column });
        obj._values = [];
        obj._numericData = data;
        return obj;
    }

    private static _typedArrayToValueObjects(data: Float64Array, rowCount: number, columnCount: number): BaseValueObject[][] {
        const values: BaseValueObject[][] = [];
        for (let r = 0; r < rowCount; r++) {
            const row: BaseValueObject[] = [];
            const offset = r * columnCount;
            for (let c = 0; c < columnCount; c++) {
                row[c] = createNumberValueObjectByRawValue(data[offset + c]);
            }
            values[r] = row;
        }
        return values;
    }

    private _values: Nullable<BaseValueObject>[][] = [];

    private _numericData: Float64Array | null = null;

    private _rowCount: number = -1;

    private _columnCount: number = -1;

    private _unitId: string = '';

    private _sheetId: string = '';

    private _currentRow: number = -1;

    private _currentColumn: number = -1;

    private _sliceCache = new Map<string, ArrayValueObject>();

    private _flattenCache: Nullable<ArrayValueObject>;
    private _sortedNumericAsc: Nullable<number[]>;
    private _sortedNumericDesc: Nullable<number[]>;
    private _equalSearchFirstIndex: Nullable<Map<string | number | boolean, { row: number; column: number }>>;
    private _equalSearchLastIndex: Nullable<Map<string | number | boolean, { row: number; column: number }>>;

    /**
     * The default value of the array, null values in comparison results support setting to false
     */
    private _defaultValue: Nullable<BaseValueObject> = null;

    private _flattenPosition: Nullable<{
        stringArray: BaseValueObject[];
        stringPosition: number[];
        numberArray: BaseValueObject[];
        numberPosition: number[];
    }>;

    constructor(rawValue: string | IArrayValueObject) {
        super(typeof rawValue === 'string' ? rawValue as string : fromObjectToString(rawValue as IArrayValueObject));

        this._values = this._formatValue(rawValue);
    }

    override dispose(): void {
        this._values = [];
        this._numericData = null;
        this._defaultValue = null;
        this._flattenPosition = null;
        this._clearCache();
    }

    isNumericArray(): boolean {
        return this._numericData !== null;
    }

    getNumericData(): Float64Array | null {
        return this._numericData;
    }

    getSortedNumericValues(descending: boolean): number[] | null {
        if (descending) {
            if (this._sortedNumericDesc) return this._sortedNumericDesc;
        } else {
            if (this._sortedNumericAsc) return this._sortedNumericAsc;
        }

        let values: number[];
        if (this._numericData) {
            values = [];
            for (let i = 0; i < this._numericData.length; i++) {
                const v = this._numericData[i];
                if (Number.isFinite(v)) values.push(v);
            }
        } else {
            values = [];
            for (let r = 0; r < this._rowCount; r++) {
                const row = this._values[r];
                if (!row) continue;
                for (let c = 0; c < this._columnCount; c++) {
                    const cell = row[c];
                    if (!cell || cell.isNull()) continue;
                    if (cell.isError()) return null;
                    if (cell.isBoolean() || cell.isString()) continue;
                    const v = cell.getValue() as number;
                    if (Number.isFinite(v)) values.push(v);
                }
            }
        }

        if (values.length === 0) return null;

        const asc = [...values].sort((a, b) => a - b);
        this._sortedNumericAsc = asc;
        this._sortedNumericDesc = [...asc].reverse();

        return descending ? this._sortedNumericDesc : this._sortedNumericAsc;
    }

    getEqualSearchIndex(isFirst: boolean): Map<string | number | boolean, { row: number; column: number }> {
        const cached = isFirst ? this._equalSearchFirstIndex : this._equalSearchLastIndex;
        if (cached) return cached;

        const index = new Map<string | number | boolean, { row: number; column: number }>();
        const rowCount = this._rowCount;
        const colCount = this._columnCount;

        if (isFirst) {
            for (let r = rowCount - 1; r >= 0; r--) {
                for (let c = colCount - 1; c >= 0; c--) {
                    const val = this._getRawValue(r, c);
                    if (val !== null) {
                        const key = typeof val === 'string' ? val.toLocaleLowerCase() : val;
                        index.set(key, { row: r, column: c });
                    }
                }
            }
            this._equalSearchFirstIndex = index;
        } else {
            for (let r = 0; r < rowCount; r++) {
                for (let c = 0; c < colCount; c++) {
                    const val = this._getRawValue(r, c);
                    if (val !== null) {
                        const key = typeof val === 'string' ? val.toLocaleLowerCase() : val;
                        index.set(key, { row: r, column: c });
                    }
                }
            }
            this._equalSearchLastIndex = index;
        }
        return index;
    }

    private _getRawValue(row: number, column: number): string | number | boolean | null {
        if (this._numericData !== null) {
            return this._numericData[row * this._columnCount + column];
        }
        const cell = this._values[row]?.[column];
        if (!cell || cell.isNull() || cell.isError()) return null;
        return cell.getValue() as string | number | boolean;
    }

    getNumberDirect(row: number, column: number): number {
        if (this._numericData !== null) {
            return this._numericData[row * this._columnCount + column];
        }
        const cell = this._values[row]?.[column];
        return cell ? (cell.getValue() as number) : 0;
    }

    rawCompare(row: number, column: number, criteriaRaw: string | number | boolean, operator: compareToken): boolean {
        if (this._numericData !== null) {
            const cellVal = this._numericData[row * this._columnCount + column];
            if (typeof criteriaRaw === 'number') {
                return this._rawCompareNumbers(cellVal, criteriaRaw, operator);
            }
            if ((operator === compareToken.EQUALS || operator === compareToken.NOT_EQUAL) && typeof criteriaRaw === 'string') {
                const criteriaNum = Number(criteriaRaw);
                if (!Number.isNaN(criteriaNum)) {
                    return this._rawCompareNumbers(cellVal, criteriaNum, operator);
                }
            }
            return operator === compareToken.NOT_EQUAL;
        }

        const cell = this._values[row]?.[column];
        if (!cell || cell.isNull() || cell.isError()) return false;

        const cellRaw = cell.getValue();
        if (typeof cellRaw === typeof criteriaRaw) {
            if (typeof cellRaw === 'string' && typeof criteriaRaw === 'string') {
                return this._rawCompareStrings(cellRaw.toLocaleLowerCase(), criteriaRaw.toLocaleLowerCase(), operator);
            }
            if (typeof cellRaw === 'number' && typeof criteriaRaw === 'number') {
                return this._rawCompareNumbers(cellRaw, criteriaRaw, operator);
            }
        }

        if (operator === compareToken.EQUALS || operator === compareToken.NOT_EQUAL) {
            if (typeof cellRaw === 'number' && typeof criteriaRaw === 'string') {
                const criteriaNum = Number(criteriaRaw);
                if (!Number.isNaN(criteriaNum)) {
                    return this._rawCompareNumbers(cellRaw, criteriaNum, operator);
                }
            }
            if (typeof cellRaw === 'string' && typeof criteriaRaw === 'number') {
                const cellNum = Number(cellRaw);
                if (!Number.isNaN(cellNum)) {
                    return this._rawCompareNumbers(cellNum, criteriaRaw, operator);
                }
            }
        }
        return operator === compareToken.NOT_EQUAL;
    }

    private _rawCompareNumbers(a: number, b: number, op: compareToken): boolean {
        switch (op) {
            case compareToken.EQUALS: return a === b;
            case compareToken.NOT_EQUAL: return a !== b;
            case compareToken.GREATER_THAN: return a > b;
            case compareToken.GREATER_THAN_OR_EQUAL: return a >= b;
            case compareToken.LESS_THAN: return a < b;
            case compareToken.LESS_THAN_OR_EQUAL: return a <= b;
        }
    }

    private _rawCompareStrings(a: string, b: string, op: compareToken): boolean {
        switch (op) {
            case compareToken.EQUALS: return a === b;
            case compareToken.NOT_EQUAL: return a !== b;
            case compareToken.GREATER_THAN: return a > b;
            case compareToken.GREATER_THAN_OR_EQUAL: return a >= b;
            case compareToken.LESS_THAN: return a < b;
            case compareToken.LESS_THAN_OR_EQUAL: return a <= b;
        }
    }

    private _materialize(): void {
        if (this._numericData === null) {
            return;
        }
        const data = this._numericData;
        const rowCount = this._rowCount;
        const columnCount = this._columnCount;
        const result: BaseValueObject[][] = new Array(rowCount);
        for (let r = 0; r < rowCount; r++) {
            const row = new Array<BaseValueObject>(columnCount);
            const offset = r * columnCount;
            for (let c = 0; c < columnCount; c++) {
                row[c] = NumberValueObject.create(data[offset + c]);
            }
            result[r] = row;
        }
        this._values = result;
        this._numericData = null;
    }

    clone() {
        return this.map((o) => {
            return o;
        }) as ArrayValueObject;
    }

    getRowCount() {
        return this._rowCount;
    }

    setRowCount(rowCount: number) {
        this._rowCount = rowCount;
    }

    getColumnCount() {
        return this._columnCount;
    }

    setColumnCount(columnCount: number) {
        this._columnCount = columnCount;
    }

    setCurrent(row: number, column: number) {
        this._currentRow = row;
        this._currentColumn = column;
    }

    setUnitId(unitId: string) {
        this._unitId = unitId;
    }

    getUnitId() {
        return this._unitId;
    }

    setSheetId(sheetId: string) {
        this._sheetId = sheetId;
    }

    getSheetId() {
        return this._sheetId;
    }

    getCurrentRow() {
        return this._currentRow;
    }

    getCurrentColumn() {
        return this._currentColumn;
    }

    override getArrayValue() {
        this._materialize();
        return this._values;
    }

    override setArrayValue(value: BaseValueObject[][]) {
        this._clearCache();
        this._numericData = null;
        this._values = value;
    }

    override isArray() {
        return true;
    }

    setDefaultValue(value: Nullable<BaseValueObject>) {
        this._defaultValue = value;
    }

    get(row: number, column: number) {
        if (this._numericData !== null) {
            const idx = row * this._columnCount + column;
            if (idx >= 0 && idx < this._numericData.length) {
                return NumberValueObject.create(this._numericData[idx]);
            }
            return this._defaultValue;
        }
        return this._values[row]?.[column] || this._defaultValue;
    }

    getRealValue(row: number, column: number) {
        if (this._numericData !== null) {
            const idx = row * this._columnCount + column;
            if (idx >= 0 && idx < this._numericData.length) {
                return NumberValueObject.create(this._numericData[idx]);
            }
            return null;
        }
        const rowValues = this._values[row];
        if (rowValues == null) {
            return null;
        }

        const v = rowValues[column];
        if (v == null) {
            return null;
        }
        return v;
    }

    getValueOrDefault(row: number, column: number) {
        return this.get(row, column) || this._defaultValue;
    }

    set(row: number, column: number, value: Nullable<BaseValueObject>) {
        if (row >= this._rowCount || column >= this._columnCount) {
            throw new Error('Exceeding array bounds.');
        }

        this._materialize();
        this._clearCache();

        this._values[row][column] = value;
    }

    getRangePosition() {
        const startRow = 0;
        const rowCount = this.getRowCount();
        const startColumn = 0;
        const columnCount = this.getColumnCount();

        return {
            startRow,
            endRow: rowCount - 1,
            startColumn,
            endColumn: columnCount - 1,
        };
    }

    iterator(
        callback: (valueObject: Nullable<BaseValueObject>, rowIndex: number, columnIndex: number) => Nullable<boolean>
    ) {
        const { startRow, endRow, startColumn, endColumn } = this.getRangePosition();

        if (this._numericData !== null) {
            const data = this._numericData;
            const cols = this._columnCount;
            for (let r = startRow; r <= endRow; r++) {
                const offset = r * cols;
                for (let c = startColumn; c <= endColumn; c++) {
                    if (callback(NumberValueObject.create(data[offset + c]), r, c) === false) {
                        return;
                    }
                }
            }
            return;
        }

        const valueList = this._values;

        for (let r = startRow; r <= endRow; r++) {
            for (let c = startColumn; c <= endColumn; c++) {
                if (callback(valueList[r]?.[c] || this._defaultValue, r, c) === false) {
                    return;
                }
            }
        }
    }

    iteratorReverse(
        callback: (valueObject: Nullable<BaseValueObject>, rowIndex: number, columnIndex: number) => Nullable<boolean>
    ) {
        const { startRow, endRow, startColumn, endColumn } = this.getRangePosition();

        if (this._numericData !== null) {
            const data = this._numericData;
            const cols = this._columnCount;
            for (let r = endRow; r >= startRow; r--) {
                const offset = r * cols;
                for (let c = endColumn; c >= startColumn; c--) {
                    if (callback(NumberValueObject.create(data[offset + c]), r, c) === false) {
                        return;
                    }
                }
            }
            return;
        }

        const valueList = this._values;

        for (let r = endRow; r >= startRow; r--) {
            for (let c = endColumn; c >= startColumn; c--) {
                if (callback(valueList[r]?.[c] || this._defaultValue, r, c) === false) {
                    return;
                }
            }
        }
    }

    getLastTruePosition() {
        let rangeSingle: Nullable<{ row: number; column: number }>;

        if (this._numericData !== null) {
            const data = this._numericData;
            const cols = this._columnCount;
            for (let i = data.length - 1; i >= 0; i--) {
                if (data[i] !== 0) {
                    return { row: Math.floor(i / cols), column: i % cols };
                }
            }
            return rangeSingle;
        }

        this.iteratorReverse((value, rowIndex, columnIndex) => {
            if (value?.isBoolean() && (value as BaseValueObject).getValue() === true) {
                rangeSingle = {
                    row: rowIndex,
                    column: columnIndex,
                };

                return false;
            }
        });

        return rangeSingle;
    }

    getFirstTruePosition() {
        let rangeSingle: Nullable<{ row: number; column: number }>;

        if (this._numericData !== null) {
            const data = this._numericData;
            const cols = this._columnCount;
            for (let i = 0; i < data.length; i++) {
                if (data[i] !== 0) {
                    return { row: Math.floor(i / cols), column: i % cols };
                }
            }
            return rangeSingle;
        }

        this.iterator((value, rowIndex, columnIndex) => {
            if (value?.isBoolean() && (value as BaseValueObject).getValue() === true) {
                rangeSingle = {
                    row: rowIndex,
                    column: columnIndex,
                };

                return false;
            }
        });

        return rangeSingle;
    }

    getFirstCell() {
        const { startRow, startColumn } = this.getRangePosition();
        return this.get(startRow, startColumn) || this._defaultValue || NullValueObject.create();
    }

    getLastCell() {
        const { endRow, endColumn } = this.getRangePosition();
        return this.get(endRow, endColumn) || this._defaultValue || NullValueObject.create();
    }

    /**
     * Referring to matrix calculations,
     * extract the matching values from a true/false matrix based on parameters and store them in a two-dimensional array.
     * implement x[x<10]
     * https://numpy.org/doc/stable/user/basics.indexing.html
     * @param takeArray
     */
    pick(takeArray: ArrayValueObject) {
        const newValue = this.pickRaw(takeArray);

        return this._createNewArray(newValue, 1, newValue[0].length);
    }

    pickRaw(takeArray: ArrayValueObject) {
        const takeArrayRowCount = takeArray.getRowCount();
        const takeArrayColumnCount = takeArray.getColumnCount();

        if (takeArrayRowCount !== this._rowCount || takeArrayColumnCount !== this._columnCount) {
            return [[NullValueObject.create()]];
        }

        const newValue: Nullable<BaseValueObject>[][] = [];

        newValue[0] = [];

        for (let r = 0; r < takeArrayRowCount; r++) {
            for (let c = 0; c < takeArrayColumnCount; c++) {
                const takeCell = takeArray.get(r, c);

                if (takeCell == null || takeCell.isError()) {
                    continue;
                }

                const takeCellValue = (takeCell as BaseValueObject).getValue();
                if (takeCellValue === true || takeCellValue === 1) {
                    const value = this.get(r, c);
                    newValue[0].push(value);
                }
            }
        }

        return newValue;
    }

    /**
     * Flatten a 2D array.
     * https://numpy.org/doc/stable/reference/generated/numpy.chararray.flatten.html#numpy.chararray.flatten
     */
    flatten() {
        if (this._flattenCache != null) {
            return this._flattenCache;
        }

        const newValue: Nullable<BaseValueObject>[][] = [];
        newValue[0] = [];
        for (let r = 0; r < this._rowCount; r++) {
            for (let c = 0; c < this._columnCount; c++) {
                const value = this.get(r, c);

                newValue[0].push(value);
            }
        }

        const arrayV = this._createNewArray(newValue, 1, newValue[0].length);

        arrayV.setDefaultValue(this._defaultValue);

        this._flattenCache = arrayV;

        return arrayV;
    }

    /**
     * Flatten a 2D array.
     * In Excel, errors and blank cells are ignored, which results in a binary search that cannot strictly adhere to the number of cells.
     */
    flattenPosition() {
        if (this._flattenPosition != null) {
            return this._flattenPosition;
        }

        const stringValue: BaseValueObject[] = [];
        const numberValue: BaseValueObject[] = [];
        const stringPosition: number[] = [];
        const numberPosition: number[] = [];
        let index = 0;
        for (let r = 0; r < this._rowCount; r++) {
            for (let c = 0; c < this._columnCount; c++) {
                const value = this.get(r, c);

                if (value == null || value.isError() || value.isNull()) {
                    index++;
                    continue;
                }

                if (value.isString()) {
                    stringValue.push(value);
                    stringPosition.push(index++);
                } else {
                    numberValue.push(value);
                    numberPosition.push(index++);
                }
            }
        }

        // const stringArray = this._createNewArray(stringValue, 1, stringValue[0].length);
        // const numberArray = this._createNewArray(numberValue, 1, numberValue[0].length);

        const result = {
            stringArray: stringValue,
            numberArray: numberValue,
            stringPosition,
            numberPosition,
        };

        this._flattenPosition = result;

        return result;
    }

    /**
     * I'm looking to perform slicing operations on 2D arrays, similar to the functionality provided by NumPy.
     * https://numpy.org/doc/stable/user/basics.indexing.html
     * @rowParam start:stop:step
     * @columnParam start:stop:step
     * @param takeArray
     */
    slice(rowParam: Nullable<Array<Nullable<number>>>, columnParam: Nullable<Array<Nullable<number>>>) {
        let rowStart = 0;
        let rowStop = this._rowCount;
        let rowStep = 1;

        let columnStart = 0;
        let columnStop = this._columnCount;
        let columnStep = 1;

        if (rowParam != null) {
            rowStart = rowParam[0] || 0;
            rowStop = rowParam[1] || this._rowCount;
            rowStep = rowParam[2] || 1;
        }

        if (columnParam != null) {
            columnStart = columnParam[0] || 0;
            columnStop = columnParam[1] || this._columnCount;
            columnStep = columnParam[2] || 1;
        }

        if (rowStart >= this._rowCount || columnStart >= this._columnCount) {
            return;
        }

        const cacheKey = `${rowStart}_${rowStop}_${rowStep}_${columnStart}_${columnStop}_${columnStep}`;

        const cache = this._sliceCache.get(cacheKey);

        if (cache != null) {
            return cache;
        }

        const result: BaseValueObject[][] = [];

        const array = this._values;

        let result_row_index = 0;
        let result_column_index = 0;
        for (let r = rowStart; r < rowStop; r += rowStep) {
            result_column_index = 0;
            if (result[result_row_index] == null) {
                result[result_row_index] = [];
            }
            for (let c = columnStart; c < columnStop; c += columnStep) {
                if (!array[r]) {
                    return;
                };

                let cell = array[r][c] || this._defaultValue;

                if (cell == null) {
                    cell = NullValueObject.create();
                }

                result[result_row_index][result_column_index] = cell;
                result_column_index++;
            }
            result_row_index++;
        }

        if (result.length === 0 || result[0].length === 0) {
            return;
        }

        /**
         * When iterating, arrayValue will create an inverted index cache for cell reference data.
         * If the array is created by ref, it will hit the cache,
         * but arrays derived from another array will lose the caching ability.
         * Here, based on the slice range, it is determined whether to set row and column to -1;
         * -1 means that caching is not enabled.
         */
        const startRow = rowStep > 1 ? -1 : rowStart + this._currentRow;

        const startColumn = columnStep > 1 ? -1 : columnStart + this._currentColumn;

        const newResultArray = this._createNewArray(result, result.length, result[0].length, startRow, startColumn);

        // Synchronize defaultValue
        newResultArray.setDefaultValue(this._defaultValue);

        this._sliceCache.set(cacheKey, newResultArray);

        return newResultArray;
    }

    sortByRow(index: number) {
        // new Intl.Collator('zh', { numeric: true }).compare;
        const result: Nullable<BaseValueObject>[][] = this._transposeArray(this._values);

        result.sort(this._sort(index));

        this._clearCache();

        this._values = this._transposeArray(result);
    }

    sortByColumn(index: number) {
        this._clearCache();
        this._values.sort(this._sort(index));
    }

    transpose() {
        const transposeArray = this._transposeArray(this._values);

        const rowCount = this._rowCount;
        const columnCount = this._columnCount;

        const newArray = this._createNewArray(transposeArray, columnCount, rowCount);

        newArray.setDefaultValue(this._defaultValue);
        return newArray;
    }

    /**
     * Due to the inability to effectively utilize the cache,
     * the sequential matching approach is only used for special matches in XLOOKUP and XMATCH.
     * For example, when match_mode is set to 1 and -1 for an exact match. If not found, it returns the next smaller item.
     */

    orderSearch(
        valueObject: BaseValueObject,
        searchType: ArrayOrderSearchType = ArrayOrderSearchType.MIN,
        isDesc = false,
        isFuzzyMatching = false
    ) {
        if (searchType === ArrayOrderSearchType.NORMAL && !isFuzzyMatching) {
            const raw = valueObject.getValue() as string | number | boolean;
            const target = typeof raw === 'string' ? raw.toLocaleLowerCase() : raw;
            const index = this.getEqualSearchIndex(true);
            const pos = index.get(target);
            return pos || undefined;
        }

        let result: Nullable<BaseValueObject>;
        let maxOrMin: Nullable<BaseValueObject>;
        let resultPosition: Nullable<{ row: number; column: number }>;

        let maxOrMinPosition: Nullable<{ row: number; column: number }>;

        const _handleMatch = (itemValue: Nullable<BaseValueObject>, row: number, column: number) => {
            if (itemValue == null) {
                return true;
            }

            let matchObject: Nullable<BaseValueObject>;
            if (isFuzzyMatching === true) {
                matchObject = itemValue.compare(valueObject as StringValueObject, compareToken.EQUALS);
            } else {
                matchObject = itemValue.isEqual(valueObject);
            }

            if (matchObject?.getValue() === true) {
                result = itemValue;
                resultPosition = { row, column };
                return false;
            }

            if (searchType === ArrayOrderSearchType.MAX) {
                if (itemValue.isGreaterThan(valueObject).getValue() === true) {
                    if (
                        maxOrMin == null ||
                        itemValue
                            .minus(valueObject)
                            .abs()
                            .isLessThanOrEqual(maxOrMin.minus(valueObject).abs())
                            .getValue() === true
                    ) {
                        maxOrMin = itemValue;
                        maxOrMinPosition = { row, column };
                    }
                }
            } else if (searchType === ArrayOrderSearchType.MIN) {
                if (itemValue.isLessThan(valueObject).getValue() === true) {
                    if (
                        maxOrMin == null ||
                        itemValue
                            .minus(valueObject)
                            .abs()
                            .isLessThanOrEqual(maxOrMin.minus(valueObject).abs())
                            .getValue() === true
                    ) {
                        maxOrMin = itemValue;
                        maxOrMinPosition = { row, column };
                    }
                }
            }
        };

        if (isDesc) {
            this.iteratorReverse((itemValue, r, c) => {
                return _handleMatch(itemValue, r, c);
            });
        } else {
            this.iterator((itemValue, r, c) => {
                return _handleMatch(itemValue, r, c);
            });
        }

        if (result != null) {
            return resultPosition;
        }

        if (maxOrMin != null) {
            return maxOrMinPosition;
        }
    }

    binarySearch(valueObject: BaseValueObject, searchType: ArrayBinarySearchType = ArrayBinarySearchType.MIN, matchType: ArrayOrderSearchType = ArrayOrderSearchType.MIN) {
        if (valueObject.isError()) {
            return;
        }

        const { stringArray, stringPosition, numberArray, numberPosition } = this.flattenPosition();

        if (valueObject.isString()) {
            return this._binarySearch(valueObject, stringArray, stringPosition, searchType, matchType);
        }

        const result = this._binarySearch(valueObject, numberArray, numberPosition, searchType, matchType);
        // if (result == null) {
        //     result = this._binarySearch(valueObject, stringArray, stringPosition, searchType);
        // }
        return result;

        // const stringMatrix = stringArray.getArrayValue();
        // const valueStringArray = stringMatrix[0];

        // const numberMatrix = numberArray.getArrayValue();
        // const valueNumberArray = numberMatrix[0];
    }

    /**
     * searchType defaults to ascending order
     *
     * matchType defaults to the maximum value less than the search value, which is used for the default matching mode of VLOOKUP/LOOKUP/HLOOKUP.
     * @param valueObject
     * @param searchArray
     * @param positionArray
     * @param searchType
     * @param matchType
     * @returns
     */
    private _binarySearch(
        valueObject: BaseValueObject,
        searchArray: BaseValueObject[],
        positionArray: number[],
        searchType: ArrayBinarySearchType = ArrayBinarySearchType.MIN,
        matchType: ArrayOrderSearchType = ArrayOrderSearchType.MIN
    ) {
        const compareFunc = getCompare();

        // Convert to number if possible, otherwise use string
        const value = Number(valueObject.getValue());
        const isValueNumber = !Number.isNaN(value);

        let start = 0;
        let end = searchArray.length - 1;
        let exactMatchIndex = -1;
        let nearestSmallerIndex = -1;
        let nearestLargerIndex = -1;

        while (start <= end) {
            const middle = Math.floor((start + end) / 2);
            const compareTo = searchArray[middle];

            let compareResult: number;
            if (compareTo.isNull()) {
                compareResult = searchType === ArrayBinarySearchType.MIN ? 1 : -1;
            } else {
                const compareToValue = compareTo.getValue();
                if (isValueNumber) {
                    const compareToNumber = Number(compareToValue);
                    compareResult = Number.isNaN(compareToNumber) ? 1 : Math.sign(compareToNumber - value);
                } else {
                    compareResult = compareFunc(compareToValue.toString().toLocaleLowerCase(), valueObject.getValue().toString().toLocaleLowerCase());
                }
            }

            // Reverse comparison result if searchType is MAX (descending order)
            if (searchType === ArrayBinarySearchType.MAX) {
                compareResult = -compareResult;
            }

            if (compareResult === 0) {
                // Exact match found
                exactMatchIndex = middle;
                break;
            }

            if (compareResult < 0) {
                // compareTo < value
                nearestSmallerIndex = middle;
                start = middle + 1;
            } else {
                // compareTo > value
                nearestLargerIndex = middle;
                end = middle - 1;
            }
        }

        // Determine the result based on matchType
        if (matchType === ArrayOrderSearchType.NORMAL) {
            return exactMatchIndex !== -1 ? positionArray[exactMatchIndex] : undefined;
        }

        if (matchType === ArrayOrderSearchType.MIN) {
            if (exactMatchIndex !== -1) return positionArray[exactMatchIndex];
            return searchType === ArrayBinarySearchType.MIN ?
                positionArray[nearestSmallerIndex]
                : positionArray[nearestLargerIndex];
        }

        if (matchType === ArrayOrderSearchType.MAX) {
            if (exactMatchIndex !== -1) return positionArray[exactMatchIndex];
            return searchType === ArrayBinarySearchType.MIN ?
                positionArray[nearestLargerIndex]
                : positionArray[nearestSmallerIndex];
        }

        // If no suitable match found based on matchType
        return undefined;
    }

    override sum() {
        if (this._numericData !== null) {
            const data = this._numericData;
            let total = 0;
            for (let i = 0, len = data.length; i < len; i++) {
                total += data[i];
            }
            return NumberValueObject.create(total);
        }

        let accumulatorAll: BaseValueObject = NumberValueObject.create(0);
        this.iterator((valueObject) => {
            if (valueObject == null || valueObject.isString() || valueObject.isBoolean() || valueObject.isNull()) {
                return true; // continue
            }

            if (valueObject.isError()) {
                accumulatorAll = valueObject;
                return false; // break
            }
            accumulatorAll = (accumulatorAll as NumberValueObject).plus(
                valueObject as BaseValueObject
            ) as BaseValueObject;
        });

        return accumulatorAll;
    }

    override max() {
        let accumulatorAll: BaseValueObject = NumberValueObject.create(Number.NEGATIVE_INFINITY);
        this.iterator((valueObject) => {
            if (valueObject == null) {
                return true; // continue
            }

            if (valueObject.isError()) {
                accumulatorAll = valueObject;
                return false; // break
            }

            // ignore boolean value in array, but not ignore in normal value object
            if (valueObject.isString() || valueObject.isNull() || valueObject.isBoolean()) {
                return true; // continue
            }

            const result = accumulatorAll.isLessThan(valueObject) as BooleanValueObject;

            if (result.getValue()) {
                accumulatorAll = valueObject as NumberValueObject;
            }
        });

        return accumulatorAll;
    }

    override min() {
        let accumulatorAll: BaseValueObject = NumberValueObject.create(Number.POSITIVE_INFINITY);

        this.iterator((valueObject) => {
            if (valueObject == null) {
                return true; // continue
            }

            if (valueObject.isError()) {
                accumulatorAll = valueObject;
                return false; // break
            }

            // ignore boolean value in array, but not ignore in normal value object
            if (valueObject.isString() || valueObject.isNull() || valueObject.isBoolean()) {
                return true; // continue
            }

            const result = accumulatorAll.isGreaterThan(valueObject) as BooleanValueObject;

            if (result.getValue()) {
                accumulatorAll = valueObject as NumberValueObject;
            }
        });

        return accumulatorAll;
    }

    override count() {
        let accumulatorAll: BaseValueObject = NumberValueObject.create(0);
        this.iterator((valueObject) => {
            // 'test', ' ',  blank cell, TRUE and FALSE are ignored
            if (valueObject == null || valueObject.isError() || valueObject.isString() || valueObject.isNull() || valueObject.isBoolean()) {
                return true; // continue
            }
            accumulatorAll = accumulatorAll.plusBy(1) as BaseValueObject;
        });

        return accumulatorAll;
    }

    override countA() {
        let accumulatorAll: BaseValueObject = NumberValueObject.create(0);
        this.iterator((valueObject) => {
            if (valueObject == null || valueObject.isNull()) {
                return true; // continue
            }

            accumulatorAll = accumulatorAll.plusBy(1);
        });

        return accumulatorAll;
    }

    override countBlank() {
        let accumulatorAll: BaseValueObject = NumberValueObject.create(0);
        this.iterator((valueObject) => {
            if (valueObject == null || valueObject.isNull() || (valueObject.getValue() === '')) {
                accumulatorAll = accumulatorAll.plusBy(1) as BaseValueObject;
            }
        });

        return accumulatorAll;
    }

    override getNegative(): BaseValueObject {
        const arrayValueObject = ArrayValueObject.create('{0}');
        return arrayValueObject.minus(this);
    }

    override getReciprocal(): BaseValueObject {
        const arrayValueObject = ArrayValueObject.create('{1}');

        return arrayValueObject.divided(this);

        // return NumberValueObject.create(1).divided(this);
    }

    override plus(valueObject: BaseValueObject): BaseValueObject {
        return this._batchOperator(valueObject, BatchOperatorType.PLUS);
    }

    override minus(valueObject: BaseValueObject): BaseValueObject {
        return this._batchOperator(valueObject, BatchOperatorType.MINUS);
    }

    override multiply(valueObject: BaseValueObject): BaseValueObject {
        return this._batchOperator(valueObject, BatchOperatorType.MULTIPLY);
    }

    override divided(valueObject: BaseValueObject): BaseValueObject {
        return this._batchOperator(valueObject, BatchOperatorType.DIVIDED);
    }

    override mod(valueObject: BaseValueObject): BaseValueObject {
        return this._batchOperator(valueObject, BatchOperatorType.MOD);
    }

    override modInverse(valueObject: BaseValueObject): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return valueObject.mod(currentValue);
        });
    }

    override compare(valueObject: BaseValueObject, operator: compareToken, isCaseSensitive?: boolean): BaseValueObject {
        return this._batchOperator(valueObject, BatchOperatorType.COMPARE, operator, isCaseSensitive);
    }

    override concatenateFront(valueObject: BaseValueObject): BaseValueObject {
        return this._batchOperator(valueObject, BatchOperatorType.CONCATENATE_FRONT);
    }

    override concatenateBack(valueObject: BaseValueObject): BaseValueObject {
        return this._batchOperator(valueObject, BatchOperatorType.CONCATENATE_BACK);
    }

    override map(callbackFn: callbackMapFnType): BaseValueObject {
        const wrappedCallbackFn = (currentValue: Nullable<BaseValueObject>, r: number, c: number) => {
            if (currentValue == null) {
                return NullValueObject.create();
            }
            if (currentValue.isError()) {
                return currentValue as ErrorValueObject;
            } else {
                return callbackFn(currentValue, r, c);
            }
        };

        return this.mapValue(wrappedCallbackFn);
    }

    // READNOW: MEMORY HOTSPOT #1 — mapValue allocates a full BaseValueObject[][]
    // for every formula operation (comparison, arithmetic, etc.). With SUMPRODUCT
    // over 10K rows, each sub-expression creates a 10K-element array. Nested
    // operations (e.g. (A=B)*(C*D)) compound: each * and = creates its own copy.
    // For 10K formulas each scanning 10K rows, this produces ~100M+ array objects.
    // Root cause of OOM in ArrayValueObject allocation (99% of heap in profiling).
    override mapValue(callbackFn: callbackMapFnType): BaseValueObject {
        const rowCount = this._rowCount;
        const columnCount = this._columnCount;

        const result = new Array<BaseValueObject[]>(rowCount);

        if (this._numericData !== null) {
            const data = this._numericData;
            for (let r = 0; r < rowCount; r++) {
                const rowList = new Array<BaseValueObject>(columnCount);
                const offset = r * columnCount;
                for (let c = 0; c < columnCount; c++) {
                    rowList[c] = callbackFn(NumberValueObject.create(data[offset + c]), r, c);
                }
                result[r] = rowList;
            }
            return this._createNewArray(result, rowCount, columnCount);
        }

        for (let r = 0; r < rowCount; r++) {
            const rowList = new Array<BaseValueObject>(columnCount);
            for (let c = 0; c < columnCount; c++) {
                const row = this._values?.[r];

                if (row == null) {
                    rowList[c] = ErrorValueObject.create(ErrorType.VALUE);
                } else {
                    const currentValue = row[c] || this._defaultValue;

                    if (currentValue) {
                        rowList[c] = callbackFn(currentValue, r, c);
                    } else {
                        rowList[c] = NullValueObject.create();
                    }
                }
            }
            result[r] = rowList;
        }

        return this._createNewArray(result, rowCount, columnCount);
    }

    override pow(valueObject: BaseValueObject): BaseValueObject {
        return this._batchOperator(valueObject, BatchOperatorType.POW);
    }

    /**
     *
     * @param valueObject In the case of an inverse, it is certainly not an array.
     * @returns
     */
    override powInverse(valueObject: BaseValueObject): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return valueObject.pow(currentValue);
        });
    }

    override sqrt(): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return currentValue.sqrt();
        });
    }

    override cbrt(): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return currentValue.cbrt();
        });
    }

    override cos(): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return currentValue.cos();
        });
    }

    override cosh(): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return currentValue.cosh();
        });
    }

    override acos(): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return currentValue.acos();
        });
    }

    override acosh(): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return currentValue.acosh();
        });
    }

    override sin(): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return currentValue.sin();
        });
    }

    override sinh(): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return currentValue.sinh();
        });
    }

    override asin(): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return currentValue.asin();
        });
    }

    override asinh(): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return currentValue.asinh();
        });
    }

    override tan(): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return currentValue.tan();
        });
    }

    override tanh(): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return currentValue.tanh();
        });
    }

    override atan(): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return currentValue.atan();
        });
    }

    override atanh(): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return currentValue.atanh();
        });
    }

    override atan2(valueObject: BaseValueObject): BaseValueObject {
        return this._batchOperator(valueObject, BatchOperatorType.ATAN2);
    }

    override atan2Inverse(valueObject: BaseValueObject): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return valueObject.atan2(currentValue);
        });
    }

    override mean(ddof: number = 0): BaseValueObject {
        const sum = this.sum();

        // Like sum, ignore strings and booleans
        const count = this.count();

        return sum.divided(ddof === 0 ? count : count.minusBy(1));
    }

    override median(): BaseValueObject {
        const numberArray = this.flattenPosition().numberArray;

        const allValue = this._createNewArray([numberArray], 1, numberArray.length);

        const count = allValue.getColumnCount();

        if (count <= 1) {
            return allValue.get(0, 0) || NullValueObject.create();
        }

        allValue.sortByRow(0);

        if (count % 2 === 0) {
            const medianRight = allValue.get(0, count / 2) || NullValueObject.create();
            const medianLeft = allValue.get(0, count / 2 - 1) || NullValueObject.create();

            return medianRight.plus(medianLeft).divided(NumberValueObject.create(2));
        }

        return allValue.get(0, (count - 1) / 2) || NullValueObject.create();
    }

    /**
     * ┌──────────────┬────────────────────────────────┬───────────────────┐
     * │ Function     │ Ignore logical values and text │ Type              │
     * ├──────────────┼────────────────────────────────┼───────────────────┤
     * │ VAR.S (VAR)  │ TRUE                           │ sample            │
     * │ VAR.P (VARP) │ TRUE                           │ entire population │
     * │ VARA         │ FALSE                          │ sample            │
     * │ VARPA        │ FALSE                          │ entire population │
     * └──────────────┴────────────────────────────────┴───────────────────┘
     *
     * for VARPA and VARA, strings and FALSE are counted as 0, TRUE is counted as 1
     * for VAR.S/VAR, or VAR.P/VARP, strings,TRUE and FALSE are ignored
     * Since sum ignores strings and booleans, they are ignored here too, and VAR.S and VAR.P are used more
     *
     * VAR.S assumes that its arguments are a sample of the population, like numpy.var(data, ddof=1)
     * VAR.P assumes that its arguments are the entire population, like numpy.var(data, ddof=0)
     * numpy.var uses ddof=0 (Delta Degrees of Freedom) by default, so we use ddof=0 here
     *
     */
    override var(ddof: number = 0): BaseValueObject {
        const mean = this.mean();

        // let isError = null;
        const squaredDifferences: BaseValueObject[][] = [[]];
        this.iterator((valueObject: Nullable<BaseValueObject>) => {
            if (valueObject == null || valueObject.isError() || valueObject.isString() || valueObject.isBoolean() || valueObject.isNull()) {
                return;
            }

            const baseValueObject = valueObject.minus(mean).pow(NumberValueObject.create(2));

            if (baseValueObject.isError()) {
                return;
            }

            squaredDifferences[0].push(baseValueObject);
        });

        const { _unitId, _sheetId, _currentRow, _currentColumn } = this;

        const squaredDifferencesArrayObject = ArrayValueObject.create({
            calculateValueList: squaredDifferences,
            rowCount: 1,
            columnCount: squaredDifferences[0].length,
            unitId: _unitId,
            sheetId: _sheetId,
            row: _currentRow,
            column: _currentColumn,
        });

        return squaredDifferencesArrayObject.mean(ddof);
    }

    /**
     * STDEV.P (STDEVP): ddof=0, ignore strings and booleans
     * STDEV.S (STDEV): ddof=1, ignore strings and booleans
     *
     * STDEVPA: ddof=0,
     * STDEVA: ddof=1,
     * @returns
     */
    override std(ddof: number = 0): BaseValueObject {
        const variance = this.var(ddof);

        if (variance.isError()) {
            return variance;
        }

        return variance.sqrt();
    }

    override log(): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return currentValue.log();
        });
    }

    override log10(): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return currentValue.log10();
        });
    }

    override exp(): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return currentValue.exp();
        });
    }

    override abs(): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return currentValue.abs();
        });
    }

    override round(valueObject: BaseValueObject): BaseValueObject {
        return this._batchOperator(valueObject, BatchOperatorType.ROUND);
    }

    override roundInverse(valueObject: BaseValueObject): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return valueObject.round(currentValue);
        });
    }

    override floor(valueObject: BaseValueObject): BaseValueObject {
        return this._batchOperator(valueObject, BatchOperatorType.FLOOR);
    }

    override floorInverse(valueObject: BaseValueObject): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return valueObject.floor(currentValue);
        });
    }

    override ceil(valueObject: BaseValueObject): BaseValueObject {
        return this._batchOperator(valueObject, BatchOperatorType.CEIL);
    }

    override ceilInverse(valueObject: BaseValueObject): BaseValueObject {
        return this.map((currentValue) => {
            if (currentValue.isError()) {
                return currentValue;
            }
            return valueObject.ceil(currentValue);
        });
    }

    toValue() {
        if (this._numericData !== null) {
            const rows = this._rowCount;
            const cols = this._columnCount;
            const result: (string | number | boolean | null)[][] = new Array(rows);
            for (let r = 0; r < rows; r++) {
                const row: (string | number | boolean | null)[] = new Array(cols);
                const offset = r * cols;
                for (let c = 0; c < cols; c++) {
                    row[c] = this._numericData[offset + c];
                }
                result[r] = row;
            }
            return result;
        }
        return transformToValue(this._values);
    }

    private _clearCache() {
        this._flattenCache = null;
        this._sortedNumericAsc = null;
        this._sortedNumericDesc = null;
        this._equalSearchFirstIndex = null;
        this._equalSearchLastIndex = null;
        this._sliceCache.clear();
    }

    private _sort(index: number) {
        const compare = getCompare();
        return (a: Nullable<BaseValueObject>[], b: Nullable<BaseValueObject>[]) => {
            const columnA = a[index];
            const columnB = b[index];

            if (columnA == null) {
                return 1;
            }

            if (columnB == null) {
                return -1;
            }

            if (columnA.isError() && columnA.isError()) {
                return 0;
            }
            if (columnA.isError()) {
                return 1;
            }
            if (columnB.isError()) {
                return -1;
            }
            return compare(
                (columnA as BaseValueObject).getValue() as string,
                (columnB as BaseValueObject).getValue() as string
            );
        };
    }

    private _transposeArray(array: Nullable<BaseValueObject>[][]) {
        // Create a new 2D array as the transposed matrix
        const rows = array.length;
        const cols = array[0].length;
        const transposedArray: Nullable<BaseValueObject>[][] = [];

        // Traverse the columns of the original two-dimensional array
        for (let col = 0; col < cols; col++) {
            // Create new row
            transposedArray[col] = [] as BaseValueObject[];

            // Traverse the rows of the original two-dimensional array
            for (let row = 0; row < rows; row++) {
                // Assign elements to new rows
                transposedArray[col][row] = array[row][col];
            }
        }

        return transposedArray;
    }

    private _extractNumericData(): Float64Array | null {
        if (!isTypedArrayOptimizationEnabled()) return null;
        if (this._numericData !== null) return this._numericData;
        const rowCount = this._rowCount;
        const columnCount = this._columnCount;
        const len = rowCount * columnCount;
        const out = new Float64Array(len);
        for (let r = 0; r < rowCount; r++) {
            const row = this._values[r];
            if (row == null) return null;
            const offset = r * columnCount;
            for (let c = 0; c < columnCount; c++) {
                const cell = row[c] || this._defaultValue;
                if (cell == null || cell.isNull() || cell.isError() || cell.isString()) return null;
                out[offset + c] = cell.isBoolean() ? (cell.getValue() ? 1 : 0) : (cell.getValue() as number);
            }
        }
        return out;
    }

    private _typedCompareEquals(scalarValue: BaseValueObject): ArrayValueObject | null {
        const rowCount = this._rowCount;
        const columnCount = this._columnCount;
        const len = rowCount * columnCount;
        const out = new Float64Array(len);
        const target = scalarValue.getValue();

        if (typeof target === 'string') return null;

        if (this._numericData !== null) {
            if (typeof target !== 'number') return null;
            const src = this._numericData;
            for (let i = 0; i < len; i++) {
                out[i] = src[i] === target ? 1 : 0;
            }
        } else {
            for (let r = 0; r < rowCount; r++) {
                const row = this._values[r];
                if (row == null) continue;
                const offset = r * columnCount;
                for (let c = 0; c < columnCount; c++) {
                    const cell = row[c];
                    if (cell == null || cell.isNull()) {
                        out[offset + c] = target === '' || target === null ? 1 : 0;
                    } else if (cell.isError()) {
                        return null;
                    } else {
                        const cv = cell.getValue();
                        if (cv === target) {
                            out[offset + c] = 1;
                        } else if (typeof cv === 'string' && typeof target === 'number') {
                            out[offset + c] = Number(cv) === target ? 1 : 0;
                        } else if (typeof cv === 'number' && typeof target === 'string') {
                            out[offset + c] = cv === Number(target) ? 1 : 0;
                        } else {
                            out[offset + c] = 0;
                        }
                    }
                }
            }
        }

        return ArrayValueObject.createNumberArray(out, rowCount, columnCount, this._unitId, this._sheetId, this._currentRow, this._currentColumn);
    }

    private _typedArithmeticFromValues(scalar: number, batchOperatorType: BatchOperatorType): ArrayValueObject | null {
        const rowCount = this._rowCount;
        const columnCount = this._columnCount;
        const len = rowCount * columnCount;
        const out = new Float64Array(len);

        for (let r = 0; r < rowCount; r++) {
            const row = this._values[r];
            if (row == null) return null;
            const offset = r * columnCount;
            for (let c = 0; c < columnCount; c++) {
                const cell = row[c];
                if (cell == null || cell.isNull() || cell.isError() || cell.isString()) return null;
                const v = cell.isBoolean() ? (cell.getValue() ? 1 : 0) : (cell.getValue() as number);
                switch (batchOperatorType) {
                    case BatchOperatorType.MULTIPLY:
                        out[offset + c] = v * scalar;
                        break;
                    case BatchOperatorType.PLUS:
                        out[offset + c] = v + scalar;
                        break;
                    case BatchOperatorType.MINUS:
                        out[offset + c] = v - scalar;
                        break;
                    case BatchOperatorType.DIVIDED:
                        if (scalar === 0) return null;
                        out[offset + c] = v / scalar;
                        break;
                    default: return null;
                }
            }
        }
        return ArrayValueObject.createNumberArray(out, rowCount, columnCount, this._unitId, this._sheetId, this._currentRow, this._currentColumn);
    }

    private _typedBinaryScalar(
        scalar: number,
        batchOperatorType: BatchOperatorType,
        operator?: compareToken
    ): ArrayValueObject | null {
        const src = this._numericData;
        if (src === null) {
            return null;
        }
        const len = src.length;
        const out = new Float64Array(len);
        switch (batchOperatorType) {
            case BatchOperatorType.MULTIPLY:
                for (let i = 0; i < len; i++) out[i] = src[i] * scalar;
                break;
            case BatchOperatorType.PLUS:
                for (let i = 0; i < len; i++) out[i] = src[i] + scalar;
                break;
            case BatchOperatorType.MINUS:
                for (let i = 0; i < len; i++) out[i] = src[i] - scalar;
                break;
            case BatchOperatorType.DIVIDED:
                if (scalar === 0) return null;
                for (let i = 0; i < len; i++) out[i] = src[i] / scalar;
                break;
            case BatchOperatorType.COMPARE: {
                if (!operator) return null;
                let cmpFn: (a: number, b: number) => number;
                switch (operator) {
                    case compareToken.EQUALS:
                        cmpFn = (a, b) => a === b ? 1 : 0;
                        break;
                    case compareToken.NOT_EQUAL:
                        cmpFn = (a, b) => a !== b ? 1 : 0;
                        break;
                    case compareToken.GREATER_THAN:
                        cmpFn = (a, b) => a > b ? 1 : 0;
                        break;
                    case compareToken.GREATER_THAN_OR_EQUAL:
                        cmpFn = (a, b) => a >= b ? 1 : 0;
                        break;
                    case compareToken.LESS_THAN:
                        cmpFn = (a, b) => a < b ? 1 : 0;
                        break;
                    case compareToken.LESS_THAN_OR_EQUAL:
                        cmpFn = (a, b) => a <= b ? 1 : 0;
                        break;
                    default: return null;
                }
                for (let i = 0; i < len; i++) out[i] = cmpFn(src[i], scalar);
                break;
            }
            default:
                return null;
        }
        return ArrayValueObject.createNumberArray(out, this._rowCount, this._columnCount, this._unitId, this._sheetId, this._currentRow, this._currentColumn);
    }

    private _typedBinaryArray(
        other: Float64Array,
        otherRowCount: number,
        otherColumnCount: number,
        batchOperatorType: BatchOperatorType,
        operator?: compareToken
    ): ArrayValueObject | null {
        const src = this._numericData;
        if (src === null) {
            return null;
        }
        if (this._rowCount !== otherRowCount || this._columnCount !== otherColumnCount) {
            return null;
        }
        const len = src.length;
        const out = new Float64Array(len);
        switch (batchOperatorType) {
            case BatchOperatorType.MULTIPLY:
                for (let i = 0; i < len; i++) out[i] = src[i] * other[i];
                break;
            case BatchOperatorType.PLUS:
                for (let i = 0; i < len; i++) out[i] = src[i] + other[i];
                break;
            case BatchOperatorType.MINUS:
                for (let i = 0; i < len; i++) out[i] = src[i] - other[i];
                break;
            case BatchOperatorType.DIVIDED:
                for (let i = 0; i < len; i++) out[i] = other[i] === 0 ? Number.NaN : src[i] / other[i];
                break;
            case BatchOperatorType.COMPARE: {
                if (!operator) return null;
                let cmpFn: (a: number, b: number) => number;
                switch (operator) {
                    case compareToken.EQUALS:
                        cmpFn = (a, b) => a === b ? 1 : 0;
                        break;
                    case compareToken.NOT_EQUAL:
                        cmpFn = (a, b) => a !== b ? 1 : 0;
                        break;
                    case compareToken.GREATER_THAN:
                        cmpFn = (a, b) => a > b ? 1 : 0;
                        break;
                    case compareToken.GREATER_THAN_OR_EQUAL:
                        cmpFn = (a, b) => a >= b ? 1 : 0;
                        break;
                    case compareToken.LESS_THAN:
                        cmpFn = (a, b) => a < b ? 1 : 0;
                        break;
                    case compareToken.LESS_THAN_OR_EQUAL:
                        cmpFn = (a, b) => a <= b ? 1 : 0;
                        break;
                    default: return null;
                }
                for (let i = 0; i < len; i++) out[i] = cmpFn(src[i], other[i]);
                break;
            }
            default:
                return null;
        }
        return ArrayValueObject.createNumberArray(out, this._rowCount, this._columnCount, this._unitId, this._sheetId, this._currentRow, this._currentColumn);
    }

    private _batchOperator(
        valueObject: BaseValueObject,
        batchOperatorType: BatchOperatorType,
        operator?: compareToken,
        isCaseSensitive?: boolean
    ): BaseValueObject {
        let rowCount = this._rowCount;
        let columnCount = this._columnCount;

        if (valueObject.isArray()) {
            const valueRowCount = (valueObject as ArrayValueObject).getRowCount();
            const valueColumnCount = (valueObject as ArrayValueObject).getColumnCount();

            rowCount = Math.max(valueRowCount, rowCount);
            columnCount = Math.max(valueColumnCount, columnCount);

            if (this._numericData !== null && (valueObject as ArrayValueObject)._numericData !== null) {
                const typed = this._typedBinaryArray(
                    (valueObject as ArrayValueObject)._numericData!,
                    valueRowCount,
                    valueColumnCount,
                    batchOperatorType,
                    operator
                );
                if (typed) return typed;
            }

            if (this._numericData !== null && (valueObject as ArrayValueObject)._numericData === null
                && this._rowCount === valueRowCount && this._columnCount === valueColumnCount) {
                const otherTyped = (valueObject as ArrayValueObject)._extractNumericData();
                if (otherTyped) {
                    const typed = this._typedBinaryArray(otherTyped, valueRowCount, valueColumnCount, batchOperatorType, operator);
                    if (typed) return typed;
                }
            }

            if (this._numericData === null && (valueObject as ArrayValueObject)._numericData !== null
                && this._rowCount === valueRowCount && this._columnCount === valueColumnCount) {
                const selfTyped = this._extractNumericData();
                if (selfTyped) {
                    const result = ArrayValueObject.createNumberArray(selfTyped, this._rowCount, this._columnCount, this._unitId, this._sheetId, this._currentRow, this._currentColumn);
                    const typed = result._typedBinaryArray((valueObject as ArrayValueObject)._numericData!, valueRowCount, valueColumnCount, batchOperatorType, operator);
                    if (typed) return typed;
                }
            }

            if (valueRowCount === 1 && valueColumnCount === 1) {
                const v = (valueObject as ArrayValueObject).getFirstCell() as BaseValueObject;

                if (this._numericData !== null && v.isNumber()) {
                    const typed = this._typedBinaryScalar(v.getValue() as number, batchOperatorType, operator);
                    if (typed) return typed;
                }

                if (batchOperatorType === BatchOperatorType.COMPARE && operator === compareToken.EQUALS && !isCaseSensitive) {
                    const typed = this._typedCompareEquals(v);
                    if (typed) return typed;
                }

                const valueList = new Array<BaseValueObject>(columnCount);
                for (let c = 0; c < columnCount; c++) {
                    valueList[c] = v;
                }
                const result = Array.from({ length: rowCount }, () => new Array<BaseValueObject>(columnCount));
                for (let c = 0; c < columnCount; c++) {
                    this._batchOperatorValue(
                        valueList[c],
                        c,
                        result,
                        batchOperatorType,
                        operator,
                        isCaseSensitive
                    );
                }

                const newArray = this._createNewArray(result, rowCount, columnCount);
                newArray.setDefaultValue(BooleanValueObject.create(false));
                return newArray;
            } else if (valueRowCount === 1 && this._columnCount > 1) {
                const list = (valueObject as ArrayValueObject).getArrayValue();
                const valueList = new Array<BaseValueObject>(columnCount);
                for (let c = 0; c < columnCount; c++) {
                    valueList[c] = list[0][c] as BaseValueObject;
                }
                const result = Array.from({ length: rowCount }, () => new Array<BaseValueObject>(columnCount));
                for (let c = 0; c < columnCount; c++) {
                    this._batchOperatorValue(
                        valueList[c],
                        c,
                        result,
                        batchOperatorType,
                        operator,
                        isCaseSensitive
                    );
                }

                const newArray = this._createNewArray(result, rowCount, columnCount);
                newArray.setDefaultValue(BooleanValueObject.create(false));
                return newArray;
            } else {
                return this._batchOperatorArray(valueObject, batchOperatorType, operator, isCaseSensitive);
            }
        } else {
            if (this._numericData !== null && valueObject.isNumber()) {
                const typed = this._typedBinaryScalar(valueObject.getValue() as number, batchOperatorType, operator);
                if (typed) return typed;
            }

            if (batchOperatorType === BatchOperatorType.COMPARE && operator === compareToken.EQUALS && !isCaseSensitive) {
                const typed = this._typedCompareEquals(valueObject);
                if (typed) return typed;
            }

            if (this._numericData === null && batchOperatorType !== BatchOperatorType.COMPARE
                && batchOperatorType !== BatchOperatorType.CONCATENATE_FRONT
                && batchOperatorType !== BatchOperatorType.CONCATENATE_BACK
                && valueObject.isNumber()) {
                const typed = this._typedArithmeticFromValues(valueObject.getValue() as number, batchOperatorType);
                if (typed) return typed;
            }

            const valueList = new Array<BaseValueObject>(columnCount);
            for (let c = 0; c < columnCount; c++) {
                valueList[c] = valueObject;
            }
            const result = Array.from({ length: rowCount }, () => new Array<BaseValueObject>(columnCount));

            for (let c = 0; c < columnCount; c++) {
                this._batchOperatorValue(
                    valueList[c],
                    c,
                    result,
                    batchOperatorType,
                    operator,
                    isCaseSensitive
                );
            }

            const newArray = this._createNewArray(result, rowCount, columnCount);
            newArray.setDefaultValue(BooleanValueObject.create(false));
            return newArray;
        }
    }

    // eslint-disable-next-line max-lines-per-function
    private _batchOperatorValue(
        valueObject: BaseValueObject,
        column: number,
        result: BaseValueObject[][],
        batchOperatorType: BatchOperatorType,
        operator?: compareToken,
        isCaseSensitive?: boolean
    ) {
        const rowCount = this._rowCount;

        const unitId = this.getUnitId();
        const sheetId = this.getSheetId();
        const startRow = this.getCurrentRow();
        const startColumn = this.getCurrentColumn();
        /**
         * If comparison operations are conducted for a single numerical value,
         * then retrieve the judgment from the inverted index. This enhances performance.
         */
        if (batchOperatorType === BatchOperatorType.COMPARE) {
            const { rowsInCache, rowsNotInCache } = CELL_INVERTED_INDEX_CACHE.canUseCache(
                unitId,
                sheetId,
                column + startColumn,
                startRow,
                startRow + rowCount - 1
            );

            if (rowsInCache.length > 0) {
                if (operator === compareToken.EQUALS) {
                    // TODO@DR-Univer: When comparing equal with two parameters, one of them is error, and the logic here is wrong
                    const rowPositions = CELL_INVERTED_INDEX_CACHE.getCellPositions(
                        unitId,
                        sheetId,
                        column + startColumn,
                        valueObject.getValue(),
                        rowsInCache
                    );

                    if (rowPositions != null) {
                        rowPositions.forEach((row) => {
                            if (row < startRow || row > startRow + rowCount - 1) {
                                return;
                            }

                            const r = row - startRow;
                            if (result[r] == null) {
                                result[r] = [];
                            }
                            result[r][column] = BooleanValueObject.create(true);
                        });
                    }
                } else {
                    const rowValuePositions = CELL_INVERTED_INDEX_CACHE.getCellValuePositions(
                        unitId,
                        sheetId,
                        column + startColumn
                    );

                    if (rowValuePositions != null) {
                        rowValuePositions.forEach((rowPositions, rowValue) => {
                            let currentValue: Nullable<BaseValueObject> = NullValueObject.create(); // handle blank cell
                            if (ERROR_TYPE_SET.has(rowValue as ErrorType)) {
                                currentValue = ErrorValueObject.create(rowValue as ErrorType);
                            } else if (typeof rowValue === 'string') {
                                currentValue = StringValueObject.create(rowValue);
                            } else if (typeof rowValue === 'number') {
                                currentValue = NumberValueObject.create(rowValue);
                            } else if (typeof rowValue === 'boolean') {
                                currentValue = BooleanValueObject.create(rowValue);
                            }

                            let matchResult;
                            if (currentValue.isError()) {
                                matchResult = currentValue;
                            } else if (valueObject.isError()) {
                                matchResult = valueObject;
                            } else {
                                matchResult = currentValue.compare(valueObject, operator as compareToken, isCaseSensitive);
                            }

                            if (matchResult.isError() || (matchResult as BooleanValueObject).getValue() === true) {
                                rowPositions.forEach((index) => {
                                    if (index >= startRow && index <= startRow + rowCount - 1) {
                                        if (result[index - startRow] == null) {
                                            result[index - startRow] = [];
                                        }
                                        result[index - startRow][column] = matchResult;
                                    }
                                });
                            }
                        });
                    }
                }

                // handle the not in cache rows
                if (rowsNotInCache.length > 0) {
                    for (const interval of rowsNotInCache) {
                        const [start, end] = interval;

                        for (let r = start; r <= end; r++) {
                            this.__batchOperatorRowValue(
                                valueObject,
                                column,
                                result,
                                batchOperatorType,
                                r - startRow,
                                unitId,
                                sheetId,
                                startRow,
                                startColumn,
                                operator,
                                isCaseSensitive
                            );
                        }

                        CELL_INVERTED_INDEX_CACHE.setContinueBuildingCache(
                            unitId,
                            sheetId,
                            column + startColumn,
                            start,
                            end
                        );
                    }
                }

                return;
            }
        }

        for (let r = 0; r < rowCount; r++) {
            this.__batchOperatorRowValue(
                valueObject,
                column,
                result,
                batchOperatorType,
                r,
                unitId,
                sheetId,
                startRow,
                startColumn,
                operator,
                isCaseSensitive
            );
        }

        CELL_INVERTED_INDEX_CACHE.setContinueBuildingCache(
            unitId,
            sheetId,
            column + startColumn,
            startRow,
            startRow + rowCount - 1
        );
    }

    // eslint-disable-next-line
    private __batchOperatorRowValue(
        valueObject: BaseValueObject,
        column: number,
        result: BaseValueObject[][],
        batchOperatorType: BatchOperatorType,
        r: number,
        unitId: string,
        sheetId: string,
        startRow: number,
        startColumn: number,
        operator?: compareToken,
        isCaseSensitive?: boolean
    ) {
        const currentValue = this.getValueOrDefault(r, column);

        if (result[r] == null) {
            result[r] = [];
        }

        if (currentValue && valueObject) {
            if (currentValue.isError()) {
                result[r][column] = currentValue as ErrorValueObject;
            } else if (valueObject.isError()) {
                // 1 + #NAME? gets #NAME?
                result[r][column] = valueObject;
            } else {
                switch (batchOperatorType) {
                    case BatchOperatorType.PLUS:
                        result[r][column] = currentValue.plus(valueObject);
                        break;
                    case BatchOperatorType.MINUS:
                        result[r][column] = currentValue.minus(valueObject);
                        break;
                    case BatchOperatorType.MULTIPLY:
                        result[r][column] = currentValue.multiply(valueObject);
                        break;
                    case BatchOperatorType.DIVIDED:
                        result[r][column] = currentValue.divided(valueObject);
                        break;
                    case BatchOperatorType.MOD:
                        result[r][column] = currentValue.mod(valueObject);
                        break;
                    case BatchOperatorType.COMPARE:
                        if (!operator) {
                            result[r][column] = ErrorValueObject.create(ErrorType.VALUE);
                        } else {
                            result[r][column] = currentValue.compare(valueObject, operator as compareToken, isCaseSensitive);
                        }
                        break;
                    case BatchOperatorType.CONCATENATE_FRONT:
                        result[r][column] = currentValue.concatenateFront(valueObject);
                        break;
                    case BatchOperatorType.CONCATENATE_BACK:
                        result[r][column] = currentValue.concatenateBack(valueObject);
                        break;
                    case BatchOperatorType.POW:
                        result[r][column] = currentValue.pow(valueObject);
                        break;
                    case BatchOperatorType.ROUND:
                        result[r][column] = currentValue.round(valueObject);
                        break;
                    case BatchOperatorType.FLOOR:
                        result[r][column] = currentValue.floor(valueObject);
                        break;
                    case BatchOperatorType.ATAN2:
                        result[r][column] = currentValue.atan2(valueObject);
                        break;
                    case BatchOperatorType.CEIL:
                        result[r][column] = currentValue.ceil(valueObject);
                        break;
                }
            }
        } else {
            result[r][column] = ErrorValueObject.create(ErrorType.NA);
        }

        /**
         * Inverted indexing enhances matching performance.
         */
        if (currentValue == null) {
            return;
        }

        if (currentValue.isError()) {
            CELL_INVERTED_INDEX_CACHE.set(
                unitId,
                sheetId,
                column + startColumn,
                (currentValue as ErrorValueObject).getErrorType(),
                r + startRow
            );
        } else {
            CELL_INVERTED_INDEX_CACHE.set(
                unitId,
                sheetId,
                column + startColumn,
                currentValue.getValue(),
                r + startRow
            );
        }
    }

    // eslint-disable-next-line max-lines-per-function, complexity
    private _batchOperatorArray(
        valueObject: BaseValueObject,
        batchOperatorType: BatchOperatorType,
        operator?: compareToken,
        isCaseSensitive?: boolean
    ) {
        let rowCount = (valueObject as ArrayValueObject).getRowCount();
        let columnCount = (valueObject as ArrayValueObject).getColumnCount();

        if (rowCount < this._rowCount) {
            rowCount = this._rowCount;
        }

        if (columnCount < this._columnCount) {
            columnCount = this._columnCount;
        }

        const result = new Array<BaseValueObject[]>(rowCount);

        const currentCalculateType = this._checkArrayCalculateType(this as ArrayValueObject);

        const opCalculateType = this._checkArrayCalculateType(valueObject as ArrayValueObject);

        for (let r = 0; r < rowCount; r++) {
            const rowList = new Array<BaseValueObject>(columnCount);
            for (let c = 0; c < columnCount; c++) {
                let currentValue: Nullable<BaseValueObject>;
                if (currentCalculateType === ArrayCalculateType.SINGLE) {
                    currentValue = this.getValueOrDefault(0, 0);
                } else if (currentCalculateType === ArrayCalculateType.ROW) {
                    currentValue = this.getValueOrDefault(0, c);
                } else if (currentCalculateType === ArrayCalculateType.COLUMN) {
                    currentValue = this.getValueOrDefault(r, 0);
                } else {
                    currentValue = this.getValueOrDefault(r, c);
                }

                let opValue: Nullable<BaseValueObject>;
                if (opCalculateType === ArrayCalculateType.SINGLE) {
                    opValue = (valueObject as ArrayValueObject).getValueOrDefault(0, 0);
                } else if (opCalculateType === ArrayCalculateType.ROW) {
                    opValue = (valueObject as ArrayValueObject).getValueOrDefault(0, c);
                } else if (opCalculateType === ArrayCalculateType.COLUMN) {
                    opValue = (valueObject as ArrayValueObject).getValueOrDefault(r, 0);
                } else {
                    opValue = (valueObject as ArrayValueObject).getValueOrDefault(r, c);
                }

                if (currentValue && opValue) {
                    if (currentValue.isError()) {
                        rowList[c] = currentValue as ErrorValueObject;
                    } else if (opValue.isError()) {
                        rowList[c] = opValue as ErrorValueObject;
                    } else {
                        switch (batchOperatorType) {
                            case BatchOperatorType.PLUS:
                                rowList[c] = currentValue.plus(opValue);
                                break;
                            case BatchOperatorType.MINUS:
                                rowList[c] = currentValue.minus(opValue);
                                break;
                            case BatchOperatorType.MULTIPLY:
                                rowList[c] = currentValue.multiply(opValue);
                                break;
                            case BatchOperatorType.DIVIDED:
                                rowList[c] = currentValue.divided(opValue);
                                break;
                            case BatchOperatorType.MOD:
                                rowList[c] = currentValue.mod(opValue);
                                break;
                            case BatchOperatorType.COMPARE:
                                if (!operator) {
                                    rowList[c] = ErrorValueObject.create(ErrorType.VALUE);
                                } else {
                                    rowList[c] = currentValue.compare(opValue, operator as compareToken, isCaseSensitive);
                                }
                                break;
                            case BatchOperatorType.CONCATENATE_FRONT:
                                rowList[c] = currentValue.concatenateFront(opValue);
                                break;
                            case BatchOperatorType.CONCATENATE_BACK:
                                rowList[c] = currentValue.concatenateBack(opValue);
                                break;
                            case BatchOperatorType.POW:
                                rowList[c] = currentValue.pow(opValue);
                                break;
                            case BatchOperatorType.ROUND:
                                rowList[c] = currentValue.round(opValue);
                                break;
                            case BatchOperatorType.ATAN2:
                                rowList[c] = currentValue.atan2(opValue);
                                break;
                            case BatchOperatorType.FLOOR:
                                rowList[c] = currentValue.floor(opValue);
                                break;
                            case BatchOperatorType.CEIL:
                                rowList[c] = currentValue.ceil(opValue);
                                break;
                        }
                    }
                }
                // else if (currentValue) {
                //     rowList[c] = currentValue;
                // } else if (opValue) {
                //     rowList[c] = opValue;
                // }
                else {
                    rowList[c] = ErrorValueObject.create(ErrorType.NA);
                }
            }
            result[r] = rowList;
        }

        return this._createNewArray(result, rowCount, columnCount);
    }

    private _checkArrayCalculateType(valueObject: ArrayValueObject) {
        if (valueObject.getRowCount() === 1 && valueObject.getColumnCount() === 1) {
            return ArrayCalculateType.SINGLE;
        }

        if (valueObject.getRowCount() === 1) {
            return ArrayCalculateType.ROW;
        }
        if (valueObject.getColumnCount() === 1) {
            return ArrayCalculateType.COLUMN;
        }

        return ArrayCalculateType.PRODUCT;
    }

    private _formatValue(rawValue: string | IArrayValueObject) {
        if (typeof rawValue !== 'string') {
            rawValue = rawValue as IArrayValueObject;

            this._rowCount = rawValue.rowCount;

            this._columnCount = rawValue.columnCount;

            this._unitId = rawValue.unitId;

            this._sheetId = rawValue.sheetId;

            this._currentRow = rawValue.row;

            this._currentColumn = rawValue.column;

            return rawValue.calculateValueList;
        }
        rawValue = rawValue.slice(1, -1) as string;
        const rowArray = rawValue.split(';');
        const rowArrayCount = rowArray.length;
        const result: BaseValueObject[][] = [];
        let maxColumnCount = 0;
        for (let r = 0; r < rowArrayCount; r++) {
            const columnRaw = rowArray[r];
            const columnArray = columnRaw.split(',');
            const columnArrayCount = columnArray.length;

            if (maxColumnCount < columnArrayCount) {
                maxColumnCount = columnArrayCount;
            }

            const row: BaseValueObject[] = [];
            for (let c = 0; c < columnArrayCount; c++) {
                const cellRaw = columnArray[c].trim();
                row.push(ValueObjectFactory.create(cellRaw));
            }
            result.push(row);
        }

        this._rowCount = rowArrayCount;

        this._columnCount = maxColumnCount;

        return result;
    }

    // READNOW: Every mapValue and _batchOperator call ends here, wrapping the
    // intermediate BaseValueObject[][] in a new ArrayValueObject. The 2D array
    // is held in calculateValueList and stays in memory until GC. No pooling
    // or reuse — each operation allocates fresh.
    private _createNewArray(
        result: Nullable<BaseValueObject>[][],
        rowCount: number,
        columnCount: number,
        row: number = -1,
        column: number = -1
    ) {
        if (this._currentColumn === -1 || this._currentRow === -1) {
            row = -1;
            column = -1;
        }

        const arrayValueObjectData: IArrayValueObject = {
            calculateValueList: result,
            rowCount,
            columnCount,
            unitId: this.getUnitId(),
            sheetId: this.getSheetId(),
            row,
            column,
        };

        return ArrayValueObject.create(arrayValueObjectData);
    }
}

export class ValueObjectFactory {
    static create(rawValue: string | number | boolean | null, isIgnoreNumberPattern: boolean = false): BaseValueObject {
        if (rawValue == null) {
            return NullValueObject.create();
        }
        if (typeof rawValue === 'boolean') {
            return BooleanValueObject.create(rawValue);
        }
        if (typeof rawValue === 'string') {
            const rawValueUpper = rawValue.toLocaleUpperCase().trim();
            if (ERROR_TYPE_SET.has(rawValueUpper as ErrorType)) {
                return ErrorValueObject.create(rawValueUpper as ErrorType);
            }
            if (rawValueUpper === BooleanValue.TRUE || rawValueUpper === BooleanValue.FALSE) {
                return createBooleanValueObjectByRawValue(rawValue);
            }
            if (isRealNum(rawValue)) {
                return NumberValueObject.create(Number(rawValue));
            }

            // value ignore whether it is a number pattern
            if (!isIgnoreNumberPattern) {
                const { isNumberPattern, value, pattern } = stringIsNumberPattern(rawValue);
                if (isNumberPattern) {
                    return NumberValueObject.create(value as number, pattern as string);
                }
            }

            const rawValueSingleLine = rawValue.replace(/\n/g, '').replace(/\r/g, '');
            if (!isStringWrappedByDoubleQuotes(rawValueSingleLine) && regexTestArrayValue(rawValueSingleLine)) {
                return ArrayValueObject.create(rawValueSingleLine);
            }

            return createStringValueObjectByRawValue(rawValue);
        }
        if (typeof rawValue === 'number') {
            return createNumberValueObjectByRawValue(rawValue);
        }
        return ErrorValueObject.create(ErrorType.VALUE);
    }
}

function isStringWrappedByDoubleQuotes(input: string) {
    const trimmedInput = input.trim();
    return trimmedInput.startsWith('"') && trimmedInput.endsWith('"');
}
