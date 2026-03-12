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

import type { BaseReferenceObject, FunctionVariantType } from '../../../engine/reference-object/base-reference-object';
import type { ArrayValueObject } from '../../../engine/value-object/array-value-object';
import type { BaseValueObject } from '../../../engine/value-object/base-value-object';
import { ErrorType } from '../../../basics/error-type';
import { compareToken } from '../../../basics/token';
import { findCompareToken, valueObjectCompare } from '../../../engine/utils/object-compare';
import { filterSameValueObjectResult } from '../../../engine/utils/value-object';
import { ErrorValueObject } from '../../../engine/value-object/base-value-object';
import { NumberValueObject } from '../../../engine/value-object/primitive-object';
import { BaseFunction } from '../../base-function';

const SUMIF_HASH_CACHE = new Map<string, Map<string, number>>();

export function clearSumifHashCache() {
    SUMIF_HASH_CACHE.clear();
}

function typedKey(v: unknown): string {
    if (typeof v === 'string') return `s${v}`;
    if (typeof v === 'number') return `n${v}`;
    if (typeof v === 'boolean') return v ? 'b1' : 'b0';
    return 'x';
}

export class Sumif extends BaseFunction {
    override minParams = 2;

    override maxParams = 3;

    override needsReferenceObject = true;

    override calculate(range: FunctionVariantType, criteria: FunctionVariantType, sumRange?: FunctionVariantType): BaseValueObject {
        if (!range.isReferenceObject() || (sumRange && !sumRange.isReferenceObject())) {
            return ErrorValueObject.create(ErrorType.VALUE);
        }

        let _criteria = criteria;

        if (criteria.isReferenceObject()) {
            _criteria = (criteria as BaseReferenceObject).toArrayValueObject();
        }

        if (_criteria.isArray()) {
            const resultArray = (_criteria as ArrayValueObject).mapValue((criteriaObject) => this._handleSingleObject(range, criteriaObject, sumRange));

            if ((resultArray as ArrayValueObject).getRowCount() === 1 && (resultArray as ArrayValueObject).getColumnCount() === 1) {
                return (resultArray as ArrayValueObject).get(0, 0) as BaseValueObject;
            }

            return resultArray;
        }

        return this._handleSingleObject(range, _criteria as BaseValueObject, sumRange);
    }

    private _handleSingleObject(range: FunctionVariantType, criteria: BaseValueObject, sumRange?: FunctionVariantType): BaseValueObject {
        if (!criteria.isError() && !process.env.DISABLE_FORMULA_HASH_CACHE) {
            let op = compareToken.EQUALS;
            let criteriaObj = criteria;
            if (criteria.isString()) {
                const [extractedOp, extractedObj] = findCompareToken(`${criteria.getValue()}`);
                op = extractedOp;
                criteriaObj = extractedObj;
            }
            if (op === compareToken.EQUALS) {
                const hashResult = this._hashSumif(range as BaseReferenceObject, criteriaObj, sumRange as BaseReferenceObject | undefined);
                if (hashResult !== null) return hashResult;
            }
        }

        const _range = (range as BaseReferenceObject).toArrayValueObject();

        let resultArrayObject = valueObjectCompare(_range, criteria);

        resultArrayObject = filterSameValueObjectResult(resultArrayObject as ArrayValueObject, _range, criteria);

        const rangeRowCount = _range.getRowCount();
        const rangeColumnCount = _range.getColumnCount();

        let _sumRange = _range;

        if (sumRange) {
            _sumRange = (sumRange as BaseReferenceObject).toArrayValueObject();

            const sumRangeRowCount = _sumRange.getRowCount();
            const sumRangeColumnCount = _sumRange.getColumnCount();

            if (rangeRowCount !== sumRangeRowCount || rangeColumnCount !== sumRangeColumnCount) {
                const rangeData = (sumRange as BaseReferenceObject).getRangeData();
                rangeData.endRow = rangeData.startRow + rangeRowCount - 1;
                rangeData.endColumn = rangeData.startColumn + rangeColumnCount - 1;

                (sumRange as BaseReferenceObject).setRangeData(rangeData);

                _sumRange = (sumRange as BaseReferenceObject).toArrayValueObject();
            }
        }

        return _sumRange.pick(resultArrayObject as ArrayValueObject).sum();
    }

    private _hashSumif(range: BaseReferenceObject, criteria: BaseValueObject, sumRange?: BaseReferenceObject): BaseValueObject | null {
        const _range = range.toArrayValueObject();
        let _sumRange = _range;
        if (sumRange) {
            _sumRange = sumRange.toArrayValueObject();
        }

        const rangeRow = _range.getCurrentRow();
        const rangeCol = _range.getCurrentColumn();
        const sumRow = _sumRange.getCurrentRow();
        const sumCol = _sumRange.getCurrentColumn();
        const cacheKey = `${_range.getUnitId()}_${_range.getSheetId()}_${rangeRow}_${rangeCol}_${_range.getRowCount()}_${_range.getColumnCount()}_${sumRow}_${sumCol}`;

        let hashMap = SUMIF_HASH_CACHE.get(cacheKey);
        if (!hashMap) {
            hashMap = new Map<string, number>();
            const rowCount = _range.getRowCount();
            const colCount = _range.getColumnCount();

            for (let r = 0; r < rowCount; r++) {
                for (let c = 0; c < colCount; c++) {
                    const cell = _range.get(r, c);
                    if (!cell || cell.isError()) continue;

                    const sumCell = _sumRange.get(r, c);
                    if (!sumCell || sumCell.isError() || !sumCell.isNumber()) continue;

                    const key = typedKey(cell.getValue());
                    const sumVal = sumCell.getValue() as number;
                    hashMap.set(key, (hashMap.get(key) || 0) + sumVal);
                }
            }
            SUMIF_HASH_CACHE.set(cacheKey, hashMap);
        }

        const lookupKey = typedKey(criteria.getValue());
        return NumberValueObject.create(hashMap.get(lookupKey) || 0);
    }
}
