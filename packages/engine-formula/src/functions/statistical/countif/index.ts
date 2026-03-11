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

const COUNTIF_HASH_CACHE = new Map<string, Map<string, number>>();

export function clearCountifHashCache() {
    COUNTIF_HASH_CACHE.clear();
}

function typedKey(v: unknown): string {
    if (typeof v === 'string') return `s${v}`;
    if (typeof v === 'number') return `n${v}`;
    if (typeof v === 'boolean') return v ? 'b1' : 'b0';
    return 'x';
}

export class Countif extends BaseFunction {
    override minParams = 2;

    override maxParams = 2;

    override needsReferenceObject = true;

    override calculate(range: FunctionVariantType, criteria: FunctionVariantType): BaseValueObject {
        if (!range.isReferenceObject()) {
            return ErrorValueObject.create(ErrorType.VALUE);
        }

        let _criteria = criteria;

        if (criteria.isReferenceObject()) {
            _criteria = (criteria as BaseReferenceObject).toArrayValueObject();
        }

        if (_criteria.isArray()) {
            const resultArray = (_criteria as ArrayValueObject).mapValue((criteriaObject) => this._handleSingleObject(range, criteriaObject));

            if ((resultArray as ArrayValueObject).getRowCount() === 1 && (resultArray as ArrayValueObject).getColumnCount() === 1) {
                return (resultArray as ArrayValueObject).get(0, 0) as BaseValueObject;
            }

            return resultArray;
        }

        return this._handleSingleObject(range, _criteria as BaseValueObject);
    }

    private _handleSingleObject(range: FunctionVariantType, criteria: BaseValueObject): BaseValueObject {
        if (!criteria.isError()) {
            let op = compareToken.EQUALS;
            let criteriaObj = criteria;
            if (criteria.isString()) {
                const [extractedOp, extractedObj] = findCompareToken(`${criteria.getValue()}`);
                op = extractedOp;
                criteriaObj = extractedObj;
            }
            if (op === compareToken.EQUALS) {
                const hashResult = this._hashCountif(range as BaseReferenceObject, criteriaObj);
                if (hashResult !== null) return hashResult;
            }
        }

        const _range = (range as BaseReferenceObject).toArrayValueObject();

        let resultArrayObject = valueObjectCompare(_range, criteria);

        resultArrayObject = filterSameValueObjectResult(resultArrayObject as ArrayValueObject, _range, criteria);

        const picked = (_range as ArrayValueObject).pick(resultArrayObject as ArrayValueObject);
        return this._countA(picked);
    }

    private _hashCountif(range: BaseReferenceObject, criteria: BaseValueObject): BaseValueObject | null {
        const _range = range.toArrayValueObject();
        const cacheKey = `${_range.getUnitId()}_${_range.getSheetId()}_${_range.getCurrentRow()}_${_range.getCurrentColumn()}_${_range.getRowCount()}_${_range.getColumnCount()}`;

        let hashMap = COUNTIF_HASH_CACHE.get(cacheKey);
        if (!hashMap) {
            hashMap = new Map<string, number>();
            const rowCount = _range.getRowCount();
            const colCount = _range.getColumnCount();

            for (let r = 0; r < rowCount; r++) {
                for (let c = 0; c < colCount; c++) {
                    const cell = _range.get(r, c);
                    if (!cell || cell.isError()) continue;
                    const key = typedKey(cell.getValue());
                    hashMap.set(key, (hashMap.get(key) || 0) + 1);
                }
            }
            COUNTIF_HASH_CACHE.set(cacheKey, hashMap);
        }

        const lookupKey = typedKey(criteria.getValue());
        return NumberValueObject.create(hashMap.get(lookupKey) || 0);
    }

    private _countA(array: ArrayValueObject) {
        let accumulatorAll: BaseValueObject = NumberValueObject.create(0);
        array.iterator((valueObject) => {
            if (valueObject == null) {
                return true; // continue
            }

            accumulatorAll = accumulatorAll.plusBy(1);
        });

        return accumulatorAll;
    }
}
