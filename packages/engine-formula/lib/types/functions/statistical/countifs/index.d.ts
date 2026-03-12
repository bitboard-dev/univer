import { FunctionVariantType } from '../../../engine/reference-object/base-reference-object';
import { BaseValueObject } from '../../../engine/value-object/base-value-object';
import { ArrayValueObject } from '../../../engine/value-object/array-value-object';
import { NumberValueObject } from '../../../engine/value-object/primitive-object';
import { BaseFunction } from '../../base-function';
export declare function clearCountifsHashCache(): void;
export declare class Countifs extends BaseFunction {
    minParams: number;
    maxParams: number;
    needsReferenceObject: boolean;
    calculate(...variants: FunctionVariantType[]): BaseValueObject;
    /**
     * Fast path: all criteria are scalar (1×1). Iterate ranges once per cell,
     * checking all criteria per row — zero intermediate array allocations.
     */
    private _scalarCountifs;
    private _hashCountifs;
    private _fallbackCountifs;
    private _aggregateResults;
}
export declare function countTrueValue(array: ArrayValueObject): NumberValueObject;
