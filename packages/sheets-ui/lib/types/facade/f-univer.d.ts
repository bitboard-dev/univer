import { IDisposable, Injector } from '@univerjs/core';
import { IColumnsHeaderCfgParam, IRowsHeaderCfgParam, SheetExtension } from '@univerjs/engine-render';
import { FUniver } from '@univerjs/core/facade';
import { FSheetHooks } from '@univerjs/sheets/facade';
/**
 * @ignore
 */
export interface IFUniverSheetsUIMixin {
    /**
     * @deprecated use same API in FWorkbook and FWorkSheet.
     */
    customizeColumnHeader(cfg: IColumnsHeaderCfgParam): void;
    /**
     * @deprecated use same API in FWorkbook and FWorkSheet.
     */
    customizeRowHeader(cfg: IRowsHeaderCfgParam): void;
    /**
     * Register sheet row header render extensions.
     * @param {string} unitId The unit id of the spreadsheet.
     * @param {SheetExtension[]} extensions The extensions to register.
     * @returns {IDisposable} The disposable instance.
     */
    registerSheetRowHeaderExtension(unitId: string, ...extensions: SheetExtension[]): IDisposable;
    /**
     * Register sheet column header render extensions.
     * @param {string} unitId The unit id of the spreadsheet.
     * @param {SheetExtension[]} extensions The extensions to register.
     * @returns {IDisposable} The disposable instance.
     */
    registerSheetColumnHeaderExtension(unitId: string, ...extensions: SheetExtension[]): IDisposable;
    /**
     * Register sheet main render extensions.
     * @param {string} unitId The unit id of the spreadsheet.
     * @param {SheetExtension[]} extensions The extensions to register.
     * @returns {IDisposable} The disposable instance.
     */
    registerSheetMainExtension(unitId: string, ...extensions: SheetExtension[]): IDisposable;
    /**
     * @deprecated use `univerAPI.addEvent` as instead.
     */
    getSheetHooks(): FSheetHooks;
    /**
     * Paste clipboard data or custom data into the active sheet at the current selection position.
     * @param {string} [htmlContent] - The HTML content from the clipboard or custom data.
     * @param {string} [textContent] - The plain text content from the clipboard or custom data.
     * @param {File[]} [files] - The files from the clipboard or custom data.
     * @return {Promise<boolean>} A promise that resolves to true if the paste operation was successful, otherwise false.
     * @example
     * ```typescript
     * // Listen for the paste event and call the pasteIntoSheet method
     * document.addEventListener('paste', async (event) => {
     *   const htmlContent = event.clipboardData.getData('text/html');
     *   const textContent = event.clipboardData.getData('text/plain');
     *   const files = Array.from(event.clipboardData.items)
     *     .map((item) => item.kind === 'file' ? item.getAsFile() : undefined)
     *     .filter(Boolean);
     *   await univerAPI.pasteIntoSheet(htmlContent, textContent, files);
     * });
     *
     * // Or paste custom data
     * univerAPI.pasteIntoSheet('<b>Bold Text</b>', 'Bold Text');
     * ```
     */
    pasteIntoSheet(htmlContent?: string, textContent?: string, files?: File[]): Promise<boolean>;
}
export declare class FUniverSheetsUIMixin extends FUniver implements IFUniverSheetsUIMixin {
    private _initSheetUIEvent;
    private _initObserverListener;
    /**
     * @ignore
     */
    _initialize(injector: Injector): void;
    private _generateClipboardCopyParam;
    private _beforeClipboardChange;
    private _clipboardChanged;
    private _generateClipboardPasteParam;
    private _generateClipboardPasteParamAsync;
    private _beforeClipboardPaste;
    private _clipboardPaste;
    private _beforeClipboardPasteAsync;
    private _clipboardPasteAsync;
    customizeColumnHeader(cfg: IColumnsHeaderCfgParam): void;
    customizeRowHeader(cfg: IRowsHeaderCfgParam): void;
    registerSheetRowHeaderExtension(unitId: string, ...extensions: SheetExtension[]): IDisposable;
    registerSheetColumnHeaderExtension(unitId: string, ...extensions: SheetExtension[]): IDisposable;
    registerSheetMainExtension(unitId: string, ...extensions: SheetExtension[]): IDisposable;
    /**
     * Get sheet render component from render by unitId and view key.
     * @private
     * @param {string} unitId The unit id of the spreadsheet.
     * @param {SHEET_VIEW_KEY} viewKey The view key of the spreadsheet.
     * @returns {Nullable<RenderComponentType>} The render component.
     */
    private _getSheetRenderComponent;
    /**
     * Get sheet hooks.
     * @returns {FSheetHooks} FSheetHooks instance
     */
    getSheetHooks(): FSheetHooks;
    pasteIntoSheet(htmlContent?: string, textContent?: string, files?: File[]): Promise<boolean>;
}
declare module '@univerjs/core/facade' {
    interface FUniver extends IFUniverSheetsUIMixin {
    }
}
