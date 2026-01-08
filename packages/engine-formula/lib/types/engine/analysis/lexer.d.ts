import { Disposable } from '@univerjs/core';
import { ErrorType } from '../../basics/error-type';
import { IFormulaCurrentConfigService } from '../../services/current-data.service';
import { IDefinedNamesService } from '../../services/defined-names.service';
import { LexerTreeBuilder } from './lexer-tree-builder';
export declare class Lexer extends Disposable {
    private readonly _definedNamesService;
    private readonly _lexerTreeBuilder;
    private readonly _formulaCurrentConfigService;
    constructor(_definedNamesService: IDefinedNamesService, _lexerTreeBuilder: LexerTreeBuilder, _formulaCurrentConfigService: IFormulaCurrentConfigService);
    treeBuilder(formulaString: string, transformSuffix?: boolean): ErrorType.VALUE | import('./lexer-node').LexerNode | (string | import('./lexer-node').LexerNode)[] | undefined;
    private _simpleCheckDefinedName;
    private _checkDefinedNameDirty;
    private _injectDefinedName;
}
