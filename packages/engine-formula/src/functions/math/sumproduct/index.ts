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

import type { BaseValueObject } from '../../../engine/value-object/base-value-object';
import { ErrorType } from '../../../basics/error-type';
import { ArrayValueObject } from '../../../engine/value-object/array-value-object';
import { ErrorValueObject } from '../../../engine/value-object/base-value-object';
import { NumberValueObject } from '../../../engine/value-object/primitive-object';
import { BaseFunction } from '../../base-function';

export class Sumproduct extends BaseFunction {
    override minParams = 1;

    override maxParams = 255;

    // READNOW: MEMORY HOTSPOT #3 — SUMPRODUCT creates intermediate number[][]
    // arrays per variant. For SUMPRODUCT((A=B)*(C*D)), the comparison (A=B)
    // already materialized a full ArrayValueObject via mapValue before reaching
    // here. Then this method allocates another number[][] to accumulate products.
    // With 10K rows and 3 variants, that's 30K+ row arrays. The real cost is
    // upstream in mapValue — by the time we're here, the arrays are already in
    // memory. But the concat+reduce on line 90 creates yet another copy.
    override calculate(array1: BaseValueObject, ...variants: BaseValueObject[]) {
        if (array1.isError()) {
            return array1;
        }

        const _array1 = this._initArray1(array1);

        if (variants.length > 0) {
            const rowCount = _array1.getRowCount();
            const columnCount = _array1.getColumnCount();

            for (let i = 0; i < variants.length; i++) {
                if (variants[i].isError()) {
                    return variants[i];
                }

                let variantRowCount = 1;
                let variantColumnCount = 1;

                if (variants[i].isArray()) {
                    variantRowCount = (variants[i] as ArrayValueObject).getRowCount();
                    variantColumnCount = (variants[i] as ArrayValueObject).getColumnCount();
                }

                if (variantRowCount !== rowCount || variantColumnCount !== columnCount) {
                    return ErrorValueObject.create(ErrorType.VALUE);
                }
            }

            let result = 0;

            for (let r = 0; r < rowCount; r++) {
                for (let c = 0; c < columnCount; c++) {
                    const array1ValueObject = _array1.get(r, c) as BaseValueObject;

                    if (array1ValueObject.isError()) {
                        return array1ValueObject;
                    }

                    let product = array1ValueObject.isNumber() ? array1ValueObject.getValue() as number : 0;

                    for (let i = 0; i < variants.length; i++) {
                        let variantValueObject = variants[i] as BaseValueObject;

                        if (variants[i].isArray()) {
                            variantValueObject = (variants[i] as ArrayValueObject).get(r, c) as BaseValueObject;
                        }

                        if (variantValueObject.isError()) {
                            return variantValueObject;
                        }

                        if (variantValueObject.isNumber()) {
                            product *= variantValueObject.getValue() as number;
                        } else {
                            product = 0;
                        }
                    }

                    result += product;
                }
            }

            return NumberValueObject.create(result);
        } else {
            return _array1.sum();
        }
    }

    private _initArray1(array1: BaseValueObject): ArrayValueObject {
        let _array1 = array1;

        if (!_array1.isArray()) {
            _array1 = ArrayValueObject.create({
                calculateValueList: [[_array1]],
                rowCount: 1,
                columnCount: 1,
                unitId: '',
                sheetId: '',
                row: 0,
                column: 0,
            });
        }

        return _array1 as ArrayValueObject;
    }
}
