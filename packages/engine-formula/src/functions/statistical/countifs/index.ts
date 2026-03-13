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

import type { FunctionVariantType } from '../../../engine/reference-object/base-reference-object';
import type { BaseValueObject } from '../../../engine/value-object/base-value-object';
import { isFormulaHashCacheEnabled, isScalarFastPathEnabled } from '../../../basics/common';
import { ErrorType } from '../../../basics/error-type';
import { compareToken } from '../../../basics/token';
import { expandArrayValueObject } from '../../../engine/utils/array-object';
import { isWildcard } from '../../../engine/utils/compare';
import { findCompareToken } from '../../../engine/utils/object-compare';
import { getBooleanResults, isSameValueObjectType, parsePairedRangeAndCriteria } from '../../../engine/utils/value-object';
import { ArrayValueObject } from '../../../engine/value-object/array-value-object';
import { ErrorValueObject } from '../../../engine/value-object/base-value-object';
import { NumberValueObject } from '../../../engine/value-object/primitive-object';
import { BaseFunction } from '../../base-function';

const COUNTIFS_HASH_CACHE = new Map<string, Map<string, number>>();

export function clearCountifsHashCache() {
    COUNTIFS_HASH_CACHE.clear();
}

function typedKey(v: unknown): string {
    if (typeof v === 'string') return `s${v}`;
    if (typeof v === 'number') return `n${v}`;
    if (typeof v === 'boolean') return v ? 'b1' : 'b0';
    return 'x';
}

export class Countifs extends BaseFunction {
    override minParams = 2;

    override maxParams = 255;

    override needsReferenceObject = true;

    override calculate(...variants: FunctionVariantType[]): BaseValueObject {
        const {
            isError,
            errorObject,
            rangeIsDifferentSize,
            criteriaMaxRowLength,
            criteriaMaxColumnLength,
            variants: _variants,
        } = parsePairedRangeAndCriteria(variants);

        if (isError) {
            return errorObject as ErrorValueObject;
        }

        if (rangeIsDifferentSize) {
            if (criteriaMaxRowLength === 1 && criteriaMaxColumnLength === 1) {
                return ErrorValueObject.create(ErrorType.VALUE);
            }

            return expandArrayValueObject(criteriaMaxRowLength, criteriaMaxColumnLength, ErrorValueObject.create(ErrorType.VALUE));
        }

        if (criteriaMaxRowLength === 1 && criteriaMaxColumnLength === 1 && isScalarFastPathEnabled()) {
            return this._scalarCountifs(_variants);
        }

        const booleanResults = getBooleanResults(_variants, criteriaMaxRowLength, criteriaMaxColumnLength, true);

        return this._aggregateResults(booleanResults);
    }

    /**
     * Fast path: all criteria are scalar (1×1). Iterate ranges once per cell,
     * checking all criteria per row — zero intermediate array allocations.
     */
    private _scalarCountifs(variants: BaseValueObject[]): BaseValueObject {
        const pairCount = variants.length / 2;
        const ranges: ArrayValueObject[] = new Array(pairCount);
        const operators: compareToken[] = new Array(pairCount);
        const criteriaObjects: BaseValueObject[] = new Array(pairCount);

        for (let p = 0; p < pairCount; p++) {
            const range = variants[p * 2];
            const criteria = variants[p * 2 + 1];

            if (!range.isArray()) {
                return this._fallbackCountifs(variants);
            }
            ranges[p] = range as ArrayValueObject;

            const criteriaValue = criteria.isArray()
                ? (criteria as ArrayValueObject).get(0, 0) || criteria
                : criteria;

            if (criteriaValue.isError()) {
                return criteriaValue;
            }

            if (criteriaValue.isString()) {
                const [op, obj] = findCompareToken(`${criteriaValue.getValue()}`);
                operators[p] = op;
                criteriaObjects[p] = obj;
            } else {
                operators[p] = compareToken.EQUALS;
                criteriaObjects[p] = criteriaValue;
            }
        }

        if (operators.every((op) => op === compareToken.EQUALS) && isFormulaHashCacheEnabled()) {
            const hashResult = this._hashCountifs(ranges, criteriaObjects);
            if (hashResult !== null) return hashResult;
        }

        const rowCount = ranges[0].getRowCount();
        const colCount = ranges[0].getColumnCount();
        let count = 0;

        const criteriaRawValues: (string | number | boolean)[] = new Array(pairCount);
        let canUseRawPath = true;
        for (let p = 0; p < pairCount; p++) {
            const raw = criteriaObjects[p].getValue() as string | number | boolean;
            criteriaRawValues[p] = typeof raw === 'string' ? raw.toLocaleLowerCase() : raw;
            if (typeof raw === 'string' && isWildcard(raw)) {
                canUseRawPath = false;
            }
        }

        if (canUseRawPath && isScalarFastPathEnabled()) {
            for (let r = 0; r < rowCount; r++) {
                for (let c = 0; c < colCount; c++) {
                    let allMatch = true;
                    for (let p = 0; p < pairCount; p++) {
                        if (!ranges[p].rawCompare(r, c, criteriaRawValues[p], operators[p])) {
                            allMatch = false;
                            break;
                        }
                    }
                    if (allMatch) count++;
                }
            }
        } else {
            for (let r = 0; r < rowCount; r++) {
                for (let c = 0; c < colCount; c++) {
                    let allMatch = true;

                    for (let p = 0; p < pairCount; p++) {
                        const cellValue = ranges[p].get(r, c);
                        if (!cellValue || cellValue.isError()) {
                            allMatch = false;
                            break;
                        }

                        if (!isSameValueObjectType(cellValue, criteriaObjects[p])) {
                            const op = operators[p];
                            if (op === compareToken.EQUALS || op === compareToken.NOT_EQUAL) {
                                if (cellValue.isNumber() && criteriaObjects[p].isString()) {
                                    const criteriaNum = criteriaObjects[p].convertToNumberObjectValue();
                                    if (criteriaNum.isNumber()) {
                                        const cmp = cellValue.compare(criteriaNum, op);
                                        if (cmp.isError() || cmp.getValue() !== true) {
                                            allMatch = false;
                                        }
                                        continue;
                                    }
                                }
                                if (criteriaObjects[p].isNumber() && cellValue.isString()) {
                                    const cellNum = cellValue.convertToNumberObjectValue();
                                    if (cellNum.isNumber()) {
                                        const cmp = cellNum.compare(criteriaObjects[p], op);
                                        if (cmp.isError() || cmp.getValue() !== true) {
                                            allMatch = false;
                                        }
                                        continue;
                                    }
                                }
                                allMatch = op === compareToken.NOT_EQUAL;
                            } else {
                                allMatch = false;
                            }
                            if (!allMatch) break;
                            continue;
                        }

                        const cmp = cellValue.compare(criteriaObjects[p], operators[p]);
                        if (cmp.isError() || cmp.getValue() !== true) {
                            allMatch = false;
                            break;
                        }
                    }

                    if (allMatch) count++;
                }
            }
        }

        return NumberValueObject.create(count);
    }

    private _hashCountifs(ranges: ArrayValueObject[], criteriaObjects: BaseValueObject[]): BaseValueObject | null {
        const pairCount = ranges.length;

        let cacheKey = '';
        for (let p = 0; p < pairCount; p++) {
            const r = ranges[p];
            cacheKey += `${r.getUnitId()}_${r.getSheetId()}_${r.getCurrentRow()}_${r.getCurrentColumn()}_${r.getRowCount()}_${r.getColumnCount()};`;
        }

        let hashMap = COUNTIFS_HASH_CACHE.get(cacheKey);
        if (!hashMap) {
            hashMap = new Map<string, number>();
            const rowCount = ranges[0].getRowCount();
            const colCount = ranges[0].getColumnCount();

            for (let r = 0; r < rowCount; r++) {
                for (let c = 0; c < colCount; c++) {
                    let compositeKey = '';
                    let valid = true;

                    for (let p = 0; p < pairCount; p++) {
                        const cell = ranges[p].get(r, c);
                        if (!cell || cell.isError()) {
                            valid = false;
                            break;
                        }
                        if (p > 0) compositeKey += '\x00';
                        compositeKey += typedKey(cell.getValue());
                    }

                    if (valid) {
                        hashMap.set(compositeKey, (hashMap.get(compositeKey) || 0) + 1);
                    }
                }
            }
            COUNTIFS_HASH_CACHE.set(cacheKey, hashMap);
        }

        let lookupKey = '';
        for (let p = 0; p < pairCount; p++) {
            if (p > 0) lookupKey += '\x00';
            lookupKey += typedKey(criteriaObjects[p].getValue());
        }

        return NumberValueObject.create(hashMap.get(lookupKey) || 0);
    }

    private _fallbackCountifs(variants: BaseValueObject[]): BaseValueObject {
        const booleanResults = getBooleanResults(variants, 1, 1, true);
        return this._aggregateResults(booleanResults);
    }

    private _aggregateResults(booleanResults: BaseValueObject[][]): BaseValueObject {
        const results = booleanResults.map((row) => {
            return row.map((booleanResult) => {
                return countTrueValue(booleanResult as ArrayValueObject);
            });
        });

        if (results.length === 1 && results[0].length === 1) {
            return results[0][0];
        }

        return ArrayValueObject.create({
            calculateValueList: results,
            rowCount: results.length,
            columnCount: results[0].length,
            unitId: this.unitId || '',
            sheetId: this.subUnitId || '',
            row: this.row,
            column: this.column,
        });
    }
}

export function countTrueValue(array: ArrayValueObject) {
    let count = 0;
    array.iterator((value) => {
        if (!value) return;
        const v = value.getValue();
        if (v === true || v === 1) {
            count++;
        }
    });
    return NumberValueObject.create(count);
}
