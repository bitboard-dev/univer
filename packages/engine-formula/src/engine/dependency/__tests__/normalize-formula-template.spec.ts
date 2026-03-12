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

import { describe, expect, it } from 'vitest';
import { normalizeFormulaTemplate } from '../formula-dependency';

describe('normalizeFormulaTemplate', () => {
    it('normalizes relative cell references to delta placeholders', () => {
        expect(normalizeFormulaTemplate('=A2+B2', 2)).toBe('=A{0}+B{0}');
        expect(normalizeFormulaTemplate('=A3+B3', 3)).toBe('=A{0}+B{0}');
    });

    it('produces matching templates for fill-down formulas', () => {
        const t1 = normalizeFormulaTemplate('=A2+B2*C2', 2);
        const t2 = normalizeFormulaTemplate('=A3+B3*C3', 3);
        const t3 = normalizeFormulaTemplate('=A100+B100*C100', 100);
        expect(t1).toBe(t2);
        expect(t2).toBe(t3);
    });

    it('preserves absolute row references ($row) unchanged', () => {
        // $2 has a $ before the digit, so [A-Z]+\d+ cannot match across the $
        expect(normalizeFormulaTemplate('=A$2', 2)).toBe('=A$2');
        expect(normalizeFormulaTemplate('=$A$2', 2)).toBe('=$A$2');
    });

    it('normalizes relative row in mixed references ($col + relative row)', () => {
        // $A2 → column is absolute ($A) but row is relative (2)
        // The regex sees: $ (skip), A (match [A-Z]+), 2 (match \d+) → normalizes
        expect(normalizeFormulaTemplate('=$A2', 2)).toBe('=$A{0}');
        expect(normalizeFormulaTemplate('=$A3', 3)).toBe('=$A{0}');
    });

    it('handles absolute ranges like $C$2:$C$2001 without normalizing', () => {
        const f = '=COUNTIFS($C$2:$C$2001,"x")';
        expect(normalizeFormulaTemplate(f, 2)).toBe('=COUNTIFS($C$2:$C$2001,"x")');
        expect(normalizeFormulaTemplate(f, 999)).toBe('=COUNTIFS($C$2:$C$2001,"x")');
    });

    it('handles the sales-pipeline SUMPRODUCT pattern', () => {
        const f1 = '=SUMPRODUCT(($B$2:$B$2001=Opportunities!B2)*$F$2:$F$2001*$G$2:$G$2001)';
        const f2 = '=SUMPRODUCT(($B$2:$B$2001=Opportunities!B3)*$F$2:$F$2001*$G$2:$G$2001)';
        const f50 = '=SUMPRODUCT(($B$2:$B$2001=Opportunities!B50)*$F$2:$F$2001*$G$2:$G$2001)';
        const t1 = normalizeFormulaTemplate(f1, 2);
        const t2 = normalizeFormulaTemplate(f2, 3);
        const t50 = normalizeFormulaTemplate(f50, 50);
        expect(t1).toBe(t2);
        expect(t2).toBe(t50);
    });

    it('handles the sales-pipeline COUNTIFS pattern', () => {
        const f1 = '=IF(COUNTIFS($C$2:$C$2001,Opportunities!C2,$E$2:$E$2001,"Closed Won")>0,1,0)';
        const f2 = '=IF(COUNTIFS($C$2:$C$2001,Opportunities!C3,$E$2:$E$2001,"Closed Won")>0,1,0)';
        expect(normalizeFormulaTemplate(f1, 2)).toBe(normalizeFormulaTemplate(f2, 3));
    });

    it('handles simple relative formulas like =E2*0.3+D2/F2', () => {
        const f1 = '=E2*0.3+IF(F2>0,D2/F2,0)*0.3+IF(G2>0,D2/G2,0)*0.4';
        const f2 = '=E3*0.3+IF(F3>0,D3/F3,0)*0.3+IF(G3>0,D3/G3,0)*0.4';
        expect(normalizeFormulaTemplate(f1, 2)).toBe(normalizeFormulaTemplate(f2, 3));
    });

    it('handles cross-row references (formula references a different row)', () => {
        // =A2+A1 on row 2 → A{0}+A{-1}
        // =A3+A2 on row 3 → A{0}+A{-1}
        expect(normalizeFormulaTemplate('=A2+A1', 2)).toBe('=A{0}+A{-1}');
        expect(normalizeFormulaTemplate('=A3+A2', 3)).toBe('=A{0}+A{-1}');
    });

    it('produces different templates for structurally different formulas', () => {
        const t1 = normalizeFormulaTemplate('=A2+B2', 2);
        const t2 = normalizeFormulaTemplate('=A2*B2', 2);
        expect(t1).not.toBe(t2);
    });

    it('does not match function names (no digits in standard function names)', () => {
        const f = '=SUMPRODUCT(A2)';
        const t = normalizeFormulaTemplate(f, 2);
        expect(t).toBe('=SUMPRODUCT(A{0})');
        expect(t).toContain('SUMPRODUCT');
    });

    it('does not false-match lowercase sheet names', () => {
        const f = '=Opportunities!A2';
        const t = normalizeFormulaTemplate(f, 2);
        expect(t).toBe('=Opportunities!A{0}');
        expect(t).toContain('Opportunities');
    });

    it('handles multi-letter columns', () => {
        expect(normalizeFormulaTemplate('=AA10+AB10', 10)).toBe('=AA{0}+AB{0}');
        expect(normalizeFormulaTemplate('=AA11+AB11', 11)).toBe('=AA{0}+AB{0}');
    });
});
