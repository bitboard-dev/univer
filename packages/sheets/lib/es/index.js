var Ci = Object.defineProperty;
var pi = (n, e, t) => e in n ? Ci(n, e, { enumerable: !0, configurable: !0, writable: !0, value: t }) : n[e] = t;
var R = (n, e, t) => pi(n, typeof e != "symbol" ? e + "" : e, t);
import { CellValueType as re, isTextFormat as mo, isRealNum as No, isBooleanString as Ps, willLoseNumericPrecision as Si, CommandType as v, IUniverInstanceService as M, Tools as A, createInterceptorKey as ft, Disposable as ue, UniverInstanceType as B, InterceptorEffectEnum as he, InterceptorManager as Oo, toDisposable as Ne, remove as je, composeInterceptors as wi, DisposableCollection as Et, BooleanNumber as te, HorizontalAlign as Do, BorderStyleTypes as Ns, Inject as $, IResourceManagerService as Rn, RTree as Ii, generateRandomId as Ut, insertMatrixArray as Tn, ObjectMatrix as Y, Rectangle as N, moveMatrixArray as Os, sliceMatrixArray as Ds, concatMatrixArray as As, spliceArray as xs, Range as X, normalizeTextRuns as vi, isSafeNumeric as Mi, RANGE_TYPE as j, createRowColIter as _i, ICommandService as E, IUndoRedoService as V, RxDisposable as $s, sequenceExecute as L, selectionToArray as yi, createIdentifier as Ao, IContextService as Ws, ErrorService as Vt, LocaleService as it, cellToRange as Vs, PermissionStatus as q, IPermissionService as Ze, isICellData as Ls, mapObjectMatrix as bi, Dimension as Re, getArrayLength as Hs, Direction as pe, Injector as so, queryObjectMatrix as at, CellModeEnum as Ei, mergeWorksheetSnapshotWithDefault as Ui, BorderType as de, FontWeight as ki, FontItalic as Ti, ILogService as Fs, DocumentDataModel as Pi, BuildTextUtils as Ni, TextX as Oi, IConfigService as xo, CustomCommandExecutionError as Di, IAuthzIoService as Ai, UserManagerService as xi, LRUMap as $i, Optional as Wi, isDefaultFormat as Vi, DependentOn as Li, Plugin as Hi, merge as Fi, IS_ROW_STYLE_PRECEDE_COLUMN_STYLE as Bi, AUTO_HEIGHT_FOR_MERGED_CELLS as ji, registerDependencies as zi, mergeOverrideWithDependencies as Gi, touchDependencies as vn } from "@univerjs/core";
import { Subject as De, BehaviorSubject as gt, merge as jt, shareReplay as Ki, takeUntil as et, switchMap as tt, of as Ie, distinctUntilChanged as Ji, skip as Bs, map as yo, first as qi, filter as os } from "rxjs";
import { SetDefinedNameMutation as en, RemoveDefinedNameMutation as Pn, SetDefinedNameMutationFactory as Yi, IDefinedNamesService as js, LexerTreeBuilder as Xi, operatorToken as Zi, deserializeRangeWithSheet as Qi, sequenceNodeType as ea, deserializeRangeWithSheetWithCache as ta, SetFormulaCalculationResultMutation as na, handleNumfmtInCell as oa, stripErrorMargin as sa, UniverFormulaEnginePlugin as ra } from "@univerjs/engine-formula";
import { takeUntil as ss, filter as rs, map as ia } from "rxjs/operators";
import { DataSyncPrimaryController as aa } from "@univerjs/rpc";
function ua(n, e, t) {
  var r, i, a;
  if (e.t) return e.t;
  if (e.v === null) return null;
  const o = n.getStyleByCell(e), s = n.getStyleByCell(t);
  if (t.t === re.FORCE_STRING) {
    if (!mo((r = s == null ? void 0 : s.n) == null ? void 0 : r.pattern) && e.v !== void 0) {
      if (No(e.v))
        return re.NUMBER;
      if (Ps(`${e.v}`))
        return re.BOOLEAN;
    }
    return re.FORCE_STRING;
  }
  return la(o) ? mo((i = o == null ? void 0 : o.n) == null ? void 0 : i.pattern) ? re.STRING : is(e, t) : mo((a = s == null ? void 0 : s.n) == null ? void 0 : a.pattern) ? re.STRING : is(e, t);
}
function is(n, e) {
  return n.v !== void 0 ? as(n.v, n.t) : as(e.v, e.t);
}
function la(n) {
  var e;
  return !!((e = n == null ? void 0 : n.n) != null && e.pattern);
}
function as(n, e) {
  return n === null ? null : typeof n == "string" ? No(n) ? (+n == 0 || +n == 1) && e === re.BOOLEAN ? re.BOOLEAN : e !== re.STRING && e !== re.FORCE_STRING && Si(n) ? re.FORCE_STRING : re.NUMBER : Ps(n) ? re.BOOLEAN : re.STRING : typeof n == "number" ? (n === 0 || n === 1) && e === re.BOOLEAN ? re.BOOLEAN : re.NUMBER : typeof n == "boolean" ? re.BOOLEAN : re.FORCE_STRING;
}
const We = (n, e) => {
  if (n.get(M).getUniverSheetInstance(e.unitId) == null)
    throw new Error("universheet is null error!");
  return {
    unitId: e.unitId,
    subUnitId: e.subUnitId,
    ranges: A.deepClone(e.ranges)
  };
}, ne = {
  id: "sheet.mutation.add-worksheet-merge",
  type: v.MUTATION,
  handler: (n, e) => {
    const o = n.get(M).getUniverSheetInstance(e.unitId);
    if (o == null)
      throw new Error("universheet is null error!");
    const s = o.getSheetBySheetId(e.subUnitId);
    if (!s) return !1;
    const i = s.getConfig().mergeData, a = e.ranges;
    for (let u = 0; u < a.length; u++)
      i.push(a[u]);
    return s.getSpanModel().rebuild(i), !0;
  }
}, ca = ft("CELL_CONTENT"), da = ft("ROW_FILTERED"), ht = {
  CELL_CONTENT: ca,
  ROW_FILTERED: da
};
var ha = /* @__PURE__ */ ((n) => (n[n.DATA_VALIDATION = 9] = "DATA_VALIDATION", n[n.NUMFMT = 10] = "NUMFMT", n[n.CELL_IMAGE = 11] = "CELL_IMAGE", n))(ha || {});
const zs = "sheet.interceptor.range-theme-id", us = "sheet.interceptor.ignore-range-theme";
var ga = Object.getOwnPropertyDescriptor, ma = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? ga(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, fa = (n, e) => (t, o) => e(t, o, n);
const ls = ft("BEFORE_CELL_EDIT"), fo = ft("AFTER_CELL_EDIT"), Ro = ft("VALIDATE_CELL");
let G = class extends ue {
  /** @ignore */
  constructor(e) {
    super();
    R(this, "_interceptorsByName", /* @__PURE__ */ new Map());
    R(this, "_commandInterceptors", []);
    R(this, "_rangeInterceptors", []);
    R(this, "_autoHeightInterceptors", []);
    R(this, "_beforeCommandInterceptor", []);
    R(this, "_afterCommandInterceptors", []);
    R(this, "_workbookDisposables", /* @__PURE__ */ new Map());
    R(this, "_worksheetDisposables", /* @__PURE__ */ new Map());
    R(this, "_interceptorsDirty", !1);
    R(this, "_composedInterceptorByKey", /* @__PURE__ */ new Map());
    R(this, "writeCellInterceptor", new Oo({
      BEFORE_CELL_EDIT: ls,
      AFTER_CELL_EDIT: fo,
      VALIDATE_CELL: Ro
    }));
    this._univerInstanceService = e, this.disposeWithMe(this._univerInstanceService.getTypeOfUnitAdded$(B.UNIVER_SHEET).subscribe((t) => {
      this._interceptWorkbook(t);
    })), this.disposeWithMe(this._univerInstanceService.getTypeOfUnitDisposed$(B.UNIVER_SHEET).subscribe(
      (t) => this._disposeWorkbookInterceptor(t)
    )), this.intercept(ht.CELL_CONTENT, {
      priority: -1,
      effect: he.Style | he.Value,
      handler: (t) => t
    }), this.disposeWithMe(this.writeCellInterceptor.intercept(fo, {
      priority: -1,
      handler: (t) => t
    })), this.disposeWithMe(this.writeCellInterceptor.intercept(ls, {
      priority: -1,
      handler: (t) => t
    })), this.disposeWithMe(this.writeCellInterceptor.intercept(Ro, {
      priority: -1,
      handler: (t) => t
    }));
  }
  dispose() {
    super.dispose(), this._workbookDisposables.forEach((e) => e.dispose()), this._workbookDisposables.clear(), this._worksheetDisposables.clear(), this._interceptorsByName.clear();
  }
  // #region intercept command execution
  /**
   * Add a listener function to a specific command to add affiliated mutations. It should be called in controllers.
   *
   * Pairs with {@link onCommandExecute}.
   *
   * @param interceptor
   * @returns
   */
  interceptCommand(e) {
    if (this._commandInterceptors.includes(e))
      throw new Error("[SheetInterceptorService]: Interceptor already exists!");
    return this._commandInterceptors.push(e), this._commandInterceptors.sort((t, o) => {
      var s, r;
      return ((s = o.priority) != null ? s : 0) - ((r = t.priority) != null ? r : 0);
    }), this.disposeWithMe(Ne(() => je(this._commandInterceptors, e)));
  }
  /**
   * When command is executing, call this method to gether undo redo mutations from upper features.
   * @param command
   * @returns
   */
  onCommandExecute(e) {
    const t = this._commandInterceptors.map((o) => o.getMutations(e));
    return {
      preUndos: t.map((o) => {
        var s;
        return (s = o.preUndos) != null ? s : [];
      }).flat(),
      undos: t.map((o) => o.undos).flat(),
      preRedos: t.map((o) => {
        var s;
        return (s = o.preRedos) != null ? s : [];
      }).flat(),
      redos: t.map((o) => o.redos).flat()
    };
  }
  interceptAfterCommand(e) {
    if (this._afterCommandInterceptors.includes(e))
      throw new Error("[SheetInterceptorService]: Interceptor already exists!");
    return this._afterCommandInterceptors.push(e), this._afterCommandInterceptors.sort((t, o) => {
      var s, r;
      return ((s = o.priority) != null ? s : 0) - ((r = t.priority) != null ? r : 0);
    }), this.disposeWithMe(Ne(() => je(this._afterCommandInterceptors, e)));
  }
  afterCommandExecute(e) {
    const t = this._afterCommandInterceptors.map((o) => o.getMutations(e));
    return {
      undos: t.map((o) => o.undos).flat(),
      redos: t.map((o) => o.redos).flat()
    };
  }
  interceptAutoHeight(e) {
    if (this._autoHeightInterceptors.includes(e))
      throw new Error("[SheetInterceptorService]: Interceptor already exists!");
    return this._autoHeightInterceptors.push(e), this._autoHeightInterceptors.sort((t, o) => {
      var s, r;
      return ((s = o.priority) != null ? s : 0) - ((r = t.priority) != null ? r : 0);
    }), this.disposeWithMe(Ne(() => je(this._autoHeightInterceptors, e)));
  }
  generateMutationsOfAutoHeight(e) {
    const t = this._autoHeightInterceptors.map((o) => o.getMutations(e));
    return {
      preUndos: t.map((o) => {
        var s;
        return (s = o.preUndos) != null ? s : [];
      }).flat(),
      undos: t.map((o) => o.undos).flat(),
      preRedos: t.map((o) => {
        var s;
        return (s = o.preRedos) != null ? s : [];
      }).flat(),
      redos: t.map((o) => o.redos).flat()
    };
  }
  /**
   * Add a listener function to a specific command to determine if the command can execute mutations. It should be
   * called in controllers.
   *
   * Pairs with {@link beforeCommandExecute}.
   *
   * @param interceptor
   * @returns
   */
  interceptBeforeCommand(e) {
    if (this._beforeCommandInterceptor.includes(e))
      throw new Error("[SheetInterceptorService]: Interceptor already exists!");
    return this._beforeCommandInterceptor.push(e), this._beforeCommandInterceptor.sort((t, o) => {
      var s, r;
      return ((s = o.priority) != null ? s : 0) - ((r = t.priority) != null ? r : 0);
    }), this.disposeWithMe(Ne(() => je(this._beforeCommandInterceptor, e)));
  }
  /**
   * before command execute, call this method to get the flag of whether it can be executed the command，
   * @param info ICommandInfo
   * @returns Promise<boolean>
   */
  async beforeCommandExecute(e) {
    return (await Promise.all(this._beforeCommandInterceptor.map((o) => o.performCheck(e)))).every((o) => o);
  }
  // #endregion
  // #region intercept ranges - mainly for pivot table currently (2024/10/28).
  /**
   * By adding callbacks to some Ranges can get some additional mutations, such as clearing all plugin data in a certain area.
   * @param interceptor IRangeInterceptors
   * @returns IDisposable
   */
  interceptRanges(e) {
    if (this._rangeInterceptors.includes(e))
      throw new Error("[SheetInterceptorService]: Interceptor already exists!");
    return this._rangeInterceptors.push(e), this._rangeInterceptors.sort((t, o) => {
      var s, r;
      return ((s = o.priority) != null ? s : 0) - ((r = t.priority) != null ? r : 0);
    }), this.disposeWithMe(Ne(() => je(this._rangeInterceptors, e)));
  }
  generateMutationsByRanges(e) {
    const t = this._rangeInterceptors.map((o) => o.getMutations(e));
    return {
      preUndos: t.map((o) => {
        var s;
        return (s = o.preUndos) != null ? s : [];
      }).flat(),
      undos: t.map((o) => o.undos).flat(),
      preRedos: t.map((o) => {
        var s;
        return (s = o.preRedos) != null ? s : [];
      }).flat(),
      redos: t.map((o) => o.redos).flat()
    };
  }
  // #endregion
  // #region intercept on writing cell
  onWriteCell(e, t, o, s, r) {
    const i = {
      subUnitId: t.getSheetId(),
      unitId: e.getUnitId(),
      workbook: e,
      worksheet: t,
      row: o,
      col: s,
      origin: A.deepClone(r)
    };
    return this.writeCellInterceptor.fetchThroughInterceptors(fo)(r, i);
  }
  // #endregion
  onValidateCell(e, t, o, s) {
    const r = {
      subUnitId: t.getSheetId(),
      unitId: e.getUnitId(),
      workbook: e,
      worksheet: t,
      row: o,
      col: s
    };
    return this.writeCellInterceptor.fetchThroughInterceptors(Ro)(Promise.resolve(!0), r);
  }
  intercept(e, t) {
    const o = e;
    this._interceptorsByName.has(o) || this._interceptorsByName.set(o, []);
    const s = this._interceptorsByName.get(o);
    s.push(t);
    const r = s.sort((i, a) => {
      var u, l;
      return ((u = a.priority) != null ? u : 0) - ((l = i.priority) != null ? l : 0);
    });
    if (this._interceptorsDirty = !0, o === ht.CELL_CONTENT) {
      const i = he.Style | he.Value;
      this._interceptorsByName.set(`${o}-${i}`, r);
      const a = he.Style | he.Value;
      return this._interceptorsByName.set(
        `${o}-${he.Style}`,
        r.filter((u) => ((u.effect || a) & he.Style) > 0)
      ), this._interceptorsByName.set(
        `${o}-${he.Value}`,
        r.filter((u) => ((u.effect || a) & he.Value) > 0)
      ), this.disposeWithMe(Ne(() => {
        je(this._interceptorsByName.get(o), t), je(this._interceptorsByName.get(`${o}-${i}`), t), je(this._interceptorsByName.get(`${o}-${he.Style}`), t), je(this._interceptorsByName.get(`${o}-${he.Value}`), t);
      }));
    } else
      return this._interceptorsByName.set(o, r), this.disposeWithMe(Ne(() => je(this._interceptorsByName.get(o), t)));
  }
  fetchThroughInterceptors(e, t, o, s) {
    const r = t === void 0 ? e : `${e}-${t}`, i = o != null ? o : r;
    let a = this._composedInterceptorByKey.get(i);
    if (!a || this._interceptorsDirty) {
      let u = this._interceptorsByName.get(r);
      u && s && (u = u.filter(s)), a = wi(u || []), this._composedInterceptorByKey.set(i, a);
    }
    return a;
  }
  _interceptWorkbook(e) {
    const t = new Et(), o = e.getUnitId(), s = this, r = (i) => {
      const a = i.getSheetId();
      i.__interceptViewModel((u) => {
        const l = new Et();
        s._worksheetDisposables.set(cs(o, i), l), l.add(u.registerCellContentInterceptor({
          getCell(c, d, h, g, m) {
            const f = i.getCellRaw(c, d);
            return s.fetchThroughInterceptors(ht.CELL_CONTENT, h, g, m)(
              f,
              {
                unitId: o,
                subUnitId: a,
                row: c,
                col: d,
                worksheet: i,
                workbook: e,
                rawData: f
              }
            );
          }
        })), l.add(u.registerRowFilteredInterceptor({
          getRowFiltered(c) {
            return !!s.fetchThroughInterceptors(ht.ROW_FILTERED)(
              !1,
              {
                unitId: o,
                subUnitId: a,
                row: c,
                workbook: e,
                worksheet: i
              }
            );
          }
        }));
      });
    };
    e.getSheets().forEach((i) => r(i)), t.add(e.sheetCreated$.subscribe((i) => r(i))), t.add(Ne(() => e.getSheets().forEach((i) => this._disposeSheetInterceptor(o, i)))), t.add(e.sheetDisposed$.subscribe((i) => this._disposeSheetInterceptor(o, i))), this._workbookDisposables.set(o, t);
  }
  _disposeWorkbookInterceptor(e) {
    const t = e.getUnitId(), o = this._workbookDisposables.get(t);
    o && (o.dispose(), this._workbookDisposables.delete(t));
  }
  _disposeSheetInterceptor(e, t) {
    const o = cs(e, t), s = this._worksheetDisposables.get(o);
    s && (s.dispose(), this._worksheetDisposables.delete(o));
  }
};
G = ma([
  fa(0, M)
], G);
function cs(n, e) {
  return `${n}|${e.getSheetId()}`;
}
const ge = (n) => {
  const e = {};
  return n.bg && (e.bg = { ...n.bg }), n.ol && (e.ol = { ...n.ol }), n.bd && (e.bd = { ...n.bd }), n.cl && (e.cl = { ...n.cl }), n.ht && (e.ht = n.ht), n.vt && (e.vt = n.vt), n.bl !== void 0 && (e.bl = n.bl), e;
};
function Ra(n) {
  const e = {};
  if (n.length === 1)
    return n[0];
  for (const t of n)
    t.bg && (e.bg = t.bg), t.ol && (e.ol = t.ol), t.bd && (e.bd = { ...e.bd, ...t.bd }), t.cl && (e.cl = t.cl), t.ht && (e.ht = t.ht), t.vt && (e.vt = t.vt), t.bl !== void 0 && (e.bl = t.bl);
  return e;
}
const ce = {
  wholeStyle: 1,
  headerRowStyle: 2,
  headerColumnStyle: 4,
  firstRowStyle: 8,
  secondRowStyle: 16,
  lastRowStyle: 32,
  firstColumnStyle: 128,
  secondColumnStyle: 256,
  lastColumnStyle: 512
};
class Rt {
  /**
   * @constructor
   * @param {string} name The name of the range theme style, it used to identify the range theme style.
   * @param {IRangeThemeStyleJSON} [options] The options to initialize the range theme style.
   */
  constructor(e, t) {
    R(this, "_name");
    /**
     * @property {Nullable<IRangeThemeStyleItem>} wholeStyle effect for the whole range.
     */
    R(this, "wholeStyle", null);
    /**
     * @property {Nullable<IRangeThemeStyleItem>} headerRowStyle effect for the header row.
     */
    R(this, "headerRowStyle", null);
    /**
     * @property {Nullable<IRangeThemeStyleItem>} headerColumnStyle effect for the header column.
     */
    R(this, "headerColumnStyle", null);
    /**
     * @property {Nullable<IRangeThemeStyleItem>} firstRowStyle effect for the first row.
     */
    R(this, "firstRowStyle", null);
    /**
     * @property {Nullable<IRangeThemeStyleItem>} secondRowStyle effect for the second row.
     */
    R(this, "secondRowStyle", null);
    /**
     * @property {Nullable<IRangeThemeStyleItem>} lastRowStyle effect for the last row.
     */
    R(this, "lastRowStyle", null);
    /**
     * @property {Nullable<IRangeThemeStyleItem>} firstColumnStyle effect for the first column.
     */
    R(this, "firstColumnStyle", null);
    /**
     * @property {Nullable<IRangeThemeStyleItem>} secondColumnStyle effect for the second column.
     */
    R(this, "secondColumnStyle", null);
    /**
     * @property {Nullable<IRangeThemeStyleItem>} lastColumnStyle effect for the last column.
     */
    R(this, "lastColumnStyle", null);
    /**
     * @property {Nullable<IRangeThemeStyleItem>} quickly get merge style
     */
    R(this, "_mergeCacheMap", /* @__PURE__ */ new Map());
    t && this.fromJson({ ...t, name: e }), this._name = e;
  }
  /**
   * Gets the name of the range theme style.The name is read only, and use to identifier the range theme style.
   * @returns {string} The name of the range theme style.
   */
  getName() {
    return this._name;
  }
  getWholeStyle() {
    return this.wholeStyle;
  }
  setWholeStyle(e) {
    this.wholeStyle = e, this._resetStyleCache();
  }
  getFirstRowStyle() {
    return this.firstRowStyle;
  }
  setFirstRowStyle(e) {
    this.firstRowStyle = e, this._resetStyleCache();
  }
  getSecondRowStyle() {
    return this.secondRowStyle;
  }
  setSecondRowStyle(e) {
    this.secondRowStyle = e, this._resetStyleCache();
  }
  getLastRowStyle() {
    return this.lastRowStyle;
  }
  setLastRowStyle(e) {
    this.lastRowStyle = e, this._resetStyleCache();
  }
  getFirstColumnStyle() {
    return this.firstColumnStyle;
  }
  setFirstColumnStyle(e) {
    this.firstColumnStyle = e, this._resetStyleCache();
  }
  getSecondColumnStyle() {
    return this.secondColumnStyle;
  }
  setSecondColumnStyle(e) {
    this.secondColumnStyle = e, this._resetStyleCache();
  }
  getLastColumnStyle() {
    return this.lastColumnStyle;
  }
  setLastColumnStyle(e) {
    this.lastColumnStyle = e, this._resetStyleCache();
  }
  getHeaderRowStyle() {
    return this.headerRowStyle;
  }
  setHeaderRowStyle(e) {
    this.headerRowStyle = e, this._resetStyleCache();
  }
  getHeaderColumnStyle() {
    return this.headerColumnStyle;
  }
  setHeaderColumnStyle(e) {
    this.headerColumnStyle = e, this._resetStyleCache();
  }
  getStyle(e, t, o, s, r) {
    let i = 0;
    return o && (i = i | ce.lastRowStyle), s && (i = i | ce.lastColumnStyle), e >= 0 && t >= 0 && (i = i | ce.wholeStyle), e % 2 === 1 && (i = i | (r ? ce.secondRowStyle : ce.firstRowStyle)), e % 2 === 0 && (i = i | (r ? ce.firstRowStyle : ce.secondRowStyle)), e === 0 && (i = i | ce.headerRowStyle), t === 0 && (i = i | ce.headerColumnStyle), t % 2 === 1 && (i = i | ce.firstColumnStyle), t % 2 === 0 && (i = i | ce.secondColumnStyle), i === 0 ? null : this._getMergeStyle(i);
  }
  _getMergeStyle(e) {
    let t = this._mergeCacheMap.get(e);
    return t || (t = this._mergeStyle(e), this._mergeCacheMap.set(e, t)), t;
  }
  _mergeStyle(e) {
    const t = [];
    return this.wholeStyle && e & ce.wholeStyle && t.push(this.wholeStyle), this.firstColumnStyle && e & ce.firstColumnStyle && t.push(this.firstColumnStyle), this.secondColumnStyle && e & ce.secondColumnStyle && t.push(this.secondColumnStyle), this.firstRowStyle && e & ce.firstRowStyle && t.push(this.firstRowStyle), this.secondRowStyle && e & ce.secondRowStyle && t.push(this.secondRowStyle), this.headerColumnStyle && e & ce.headerColumnStyle && t.push(this.headerColumnStyle), this.lastColumnStyle && e & ce.lastColumnStyle && t.push(this.lastColumnStyle), this.headerRowStyle && e & ce.headerRowStyle && t.push(this.headerRowStyle), this.lastRowStyle && e & ce.lastRowStyle && t.push(this.lastRowStyle), Ra(t);
  }
  _resetStyleCache() {
    this._mergeCacheMap.clear();
  }
  toJson() {
    const e = {
      name: this._name
    };
    return this.wholeStyle && (e.wholeStyle = ge(this.wholeStyle)), this.headerRowStyle && (e.headerRowStyle = ge(this.headerRowStyle)), this.headerColumnStyle && (e.headerColumnStyle = ge(this.headerColumnStyle)), this.firstRowStyle && (e.firstRowStyle = ge(this.firstRowStyle)), this.secondRowStyle && (e.secondRowStyle = ge(this.secondRowStyle)), this.lastRowStyle && (e.lastRowStyle = ge(this.lastRowStyle)), this.firstColumnStyle && (e.firstColumnStyle = ge(this.firstColumnStyle)), this.secondColumnStyle && (e.secondColumnStyle = ge(this.secondColumnStyle)), this.lastColumnStyle && (e.lastColumnStyle = ge(this.lastColumnStyle)), e;
  }
  fromJson(e) {
    this._name = e.name, e.wholeStyle && (this.wholeStyle = ge(e.wholeStyle)), e.headerRowStyle && (this.headerRowStyle = ge(e.headerRowStyle)), e.headerColumnStyle && (this.headerColumnStyle = ge(e.headerColumnStyle)), e.firstRowStyle && (this.firstRowStyle = ge(e.firstRowStyle)), e.secondRowStyle && (this.secondRowStyle = ge(e.secondRowStyle)), e.lastRowStyle && (this.lastRowStyle = ge(e.lastRowStyle)), e.firstColumnStyle && (this.firstColumnStyle = ge(e.firstColumnStyle)), e.secondColumnStyle && (this.secondColumnStyle = ge(e.secondColumnStyle)), e.lastColumnStyle && (this.lastColumnStyle = ge(e.lastColumnStyle));
  }
  dispose() {
    this._mergeCacheMap.clear();
  }
}
const Ca = (n, e, t) => new Rt(`light-${n}`, {
  headerRowStyle: {
    bg: {
      rgb: e
    }
  },
  firstColumnStyle: {
    bg: {
      rgb: "rgb(255, 255, 255)"
    }
  },
  secondColumnStyle: {
    bg: {
      rgb: t
    }
  },
  lastRowStyle: {
    bg: {
      rgb: e
    }
  }
}), pa = (n, e, t) => new Rt(`middle-${n}`, {
  headerRowStyle: {
    bg: {
      rgb: e
    }
  },
  headerColumnStyle: {
    bg: {
      rgb: t
    }
  },
  secondRowStyle: {
    bg: {
      rgb: t
    }
  },
  lastRowStyle: {
    bg: {
      rgb: e
    }
  },
  lastColumnStyle: {
    bg: {
      rgb: t
    }
  }
}), Sa = (n, e, t, o) => new Rt(`dark-${n}`, {
  headerRowStyle: {
    bg: {
      rgb: e
    },
    cl: {
      rgb: "rgb(255, 255, 255)"
    },
    ht: Do.CENTER,
    bl: te.TRUE
  },
  firstRowStyle: {
    bg: {
      rgb: t
    }
  },
  secondRowStyle: {
    bg: {
      rgb: o
    }
  },
  lastRowStyle: {
    bg: {
      rgb: e
    }
  }
}), wa = [
  {
    baseName: "blue",
    header: "rgb(164, 202, 254)",
    color: "rgb(225, 239, 254)"
  },
  {
    baseName: "grey",
    header: "rgb(205, 208, 216)",
    color: "rgb(238, 239, 241)"
  },
  {
    baseName: "red",
    header: "rgb(248, 180, 180)",
    color: "rgb(253, 232, 232)"
  },
  {
    baseName: "orange",
    header: "rgb(253, 186, 140)",
    color: "rgb(254, 236, 220)"
  },
  {
    baseName: "yellow",
    header: "rgb(250, 200, 21)",
    color: "rgb(255, 244, 185)"
  },
  {
    baseName: "green",
    header: "rgb(132, 225, 188)",
    color: "rgb(222, 247, 236)"
  },
  {
    baseName: "azure",
    header: "rgb(126, 220, 226)",
    color: "rgb(213, 245, 246)"
  },
  {
    baseName: "indigo",
    header: "rgb(186, 198, 248)",
    color: "rgb(233, 237, 255)"
  },
  {
    baseName: "purple",
    header: "rgb(202, 191, 253)",
    color: "rgb(237, 235, 254)"
  },
  {
    baseName: "magenta",
    header: "rgb(248, 180, 217)",
    color: "rgb(252, 232, 243)"
  }
], Ia = [
  {
    baseName: "blue",
    rowHeader: "rgb(63, 131, 248)",
    colHeader: "rgb(195, 221, 253)"
  },
  {
    baseName: "grey",
    rowHeader: "rgb(95, 101, 116)",
    colHeader: "rgb(227, 229, 234)"
  },
  {
    baseName: "red",
    rowHeader: "rgb(240, 82, 82)",
    colHeader: "rgb(251, 213, 213)"
  },
  {
    baseName: "orange",
    rowHeader: "rgb(255, 90, 31)",
    colHeader: "rgb(252, 217, 189)"
  },
  {
    baseName: "yellow",
    rowHeader: "rgb(212, 157, 15)",
    colHeader: "rgb(252, 220, 106)"
  },
  {
    baseName: "green",
    rowHeader: "rgb(13, 164, 113)",
    colHeader: "rgb(188, 240, 218)"
  },
  {
    baseName: "azure",
    rowHeader: "rgb(6, 148, 162)",
    colHeader: "rgb(175, 236, 239)"
  },
  {
    baseName: "indigo",
    rowHeader: "rgb(70, 106, 247)",
    colHeader: "rgb(210, 218, 250)"
  },
  {
    baseName: "purple",
    rowHeader: "rgb(144, 97, 249)",
    colHeader: "rgb(220, 215, 254)"
  },
  {
    baseName: "magenta",
    rowHeader: "rgb(231, 70, 148)",
    colHeader: "rgb(250, 209, 232)"
  }
], va = [
  {
    baseName: "blue",
    rowHeader: "rgb(30, 66, 159)",
    firstRow: "rgb(195, 221, 253)",
    secondRow: "rgb(118, 169, 250)"
  },
  {
    baseName: "grey",
    rowHeader: "rgb(44, 48, 64)",
    firstRow: "rgb(227, 229, 234)",
    secondRow: "rgb(151, 157, 172)"
  },
  {
    baseName: "red",
    rowHeader: "rgb(155, 28, 28)",
    firstRow: "rgb(251, 213, 213)",
    secondRow: "rgb(249, 128, 128)"
  },
  {
    baseName: "orange",
    rowHeader: "rgb(180, 52, 3)",
    firstRow: "rgb(252, 217, 189)",
    secondRow: "rgb(255, 138, 76)"
  },
  {
    baseName: "yellow",
    rowHeader: "rgb(154, 109, 21)",
    firstRow: "rgb(252, 220, 106)",
    secondRow: "rgb(212, 157, 15)"
  },
  {
    baseName: "green",
    rowHeader: "rgb(4, 108, 78)",
    firstRow: "rgb(188, 240, 218)",
    secondRow: "rgb(49, 196, 141)"
  },
  {
    baseName: "azure",
    rowHeader: "rgb(3, 102, 114)",
    firstRow: "rgb(175, 236, 239)",
    secondRow: "rgb(22, 189, 202)"
  },
  {
    baseName: "indigo",
    rowHeader: "rgb(16, 51, 191)",
    firstRow: "rgb(210, 218, 250)",
    secondRow: "rgb(98, 128, 249)"
  },
  {
    baseName: "purple",
    rowHeader: "rgb(74, 29, 150)",
    firstRow: "rgb(220, 215, 254)",
    secondRow: "rgb(172, 148, 250)"
  },
  {
    baseName: "magenta",
    rowHeader: "rgb(153, 21, 75)",
    firstRow: "rgb(250, 209, 232)",
    secondRow: "rgb(241, 126, 184)"
  }
], Ma = wa.map(({ baseName: n, header: e, color: t }) => Ca(n, e, t)), _a = Ia.map(({ baseName: n, rowHeader: e, colHeader: t }) => pa(n, e, t)), ya = va.map(({ baseName: n, rowHeader: e, firstRow: t, secondRow: o }) => Sa(n, e, t, o)), ba = [
  ...Ma,
  ..._a,
  ...ya
], Gs = {
  headerRowStyle: {
    bg: {
      rgb: "rgb(68,114,196)"
    },
    cl: {
      rgb: "rgb(255,255,255)"
    },
    ht: Do.CENTER,
    bl: te.TRUE
  },
  firstRowStyle: {
    bg: {
      rgb: "rgb(217,225,242)"
    }
  }
}, Ea = new Rt("default", Gs), Ua = new Rt("default-last-row", {
  ...Gs,
  lastRowStyle: {
    bd: {
      t: {
        s: Ns.THIN,
        cl: {
          rgb: "rgb(68,114,196)"
        }
      }
    },
    ht: Do.CENTER,
    bl: te.TRUE
  }
});
class ka {
  constructor() {
    R(this, "_toggleRanges", []);
  }
  /**
   * Refresh the cache based on the given range and visibility function.
   * This method calculates toggle ranges for rows that are visible within the specified range.
   * Hidden rows are excluded from the toggle calculation.
   * @param range The range of rows to refresh (startRow and endRow are required).
   * @param visibleFunc A function to determine if a row is visible.
   */
  refresh(e, t) {
    const { startRow: o, endRow: s } = e, r = [];
    let i = 0, a = !1, u = -1;
    for (let l = o; l <= s; l++) {
      if (!t(l)) {
        i++, i % 2 === 1 ? a = !0 : (a = !1, u !== -1 && (r.push([u, l - 1]), u = -1));
        continue;
      }
      i % 2 === 1 ? a ? u === -1 && (u = l) : (a = !0, u = l) : a && (r.push([u, l - 2]), a = !1, u = -1), l === s && a && r.push([u, l]);
    }
    this._toggleRanges = r;
  }
  /**
   * This function returns the toggle ranges. Only for testing purposes. In production, you should use `getIsToggled` to check if a row is toggled.
   * @returns [IToggleRange[]] The toggle ranges calculated by the last refresh.
   */
  getToggleRanges() {
    return this._toggleRanges.concat();
  }
  /**
   * Check if the given row is toggled (odd/even state).
   * This method uses binary search to efficiently determine if the row is within a toggle range.
   * @param row The row to check.
   * @returns True if the row is toggled (odd), false otherwise (even or hidden).
   */
  getIsToggled(e) {
    let t = 0, o = this._toggleRanges.length - 1;
    for (; t <= o; ) {
      const s = Math.floor((t + o) / 2), [r, i] = this._toggleRanges[s];
      if (e < r)
        o = s - 1;
      else if (e > i)
        t = s + 1;
      else
        return !0;
    }
    return !1;
  }
}
var Ta = Object.getOwnPropertyDescriptor, Pa = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? Ta(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, Co = (n, e) => (t, o) => e(t, o, n);
const Na = "SHEET_RANGE_THEME_MODEL_PLUGIN";
let ve = class extends ue {
  constructor(e, t, o) {
    super();
    R(this, "_rangeThemeStyleMap", /* @__PURE__ */ new Map());
    R(this, "_rangeThemeStyleRuleMap", /* @__PURE__ */ new Map());
    R(this, "_rTreeCollection", /* @__PURE__ */ new Map());
    R(this, "_defaultRangeThemeMap", /* @__PURE__ */ new Map());
    /**
     * This map is used to cache zebra crossing toggle ranges for each unitId and subUnitId, IRangeThemeStyleRule id
     */
    R(this, "_zebraCrossingCacheMap", /* @__PURE__ */ new Map());
    R(this, "_rowVisibleFuncSet", /* @__PURE__ */ new Map());
    R(this, "_rangeThemeMapChanged$", new De());
    R(this, "rangeThemeMapChange$", this._rangeThemeMapChanged$.asObservable());
    this._sheetInterceptorService = e, this._resourceManagerService = t, this._univerInstanceService = o, this._registerIntercept(), this._initSnapshot(), this._initDefaultTheme();
  }
  _initDefaultTheme() {
    this.registerDefaultRangeTheme(Ea), this.registerDefaultRangeTheme(Ua);
    for (const e of ba)
      this.registerDefaultRangeTheme(e);
  }
  _ensureRangeThemeStyleMap(e) {
    return this._rangeThemeStyleMap.has(e) || this._rangeThemeStyleMap.set(e, /* @__PURE__ */ new Map()), this._rangeThemeStyleMap.get(e);
  }
  _ensureRangeThemeStyleRuleMap(e) {
    return this._rangeThemeStyleRuleMap.has(e) || this._rangeThemeStyleRuleMap.set(e, /* @__PURE__ */ new Map()), this._rangeThemeStyleRuleMap.get(e);
  }
  _ensureRTreeCollection(e) {
    return this._rTreeCollection.has(e) || this._rTreeCollection.set(e, new Ii()), this._rTreeCollection.get(e);
  }
  getDefaultRangeThemeStyle(e) {
    return this._defaultRangeThemeMap.get(e);
  }
  getCustomRangeThemeStyle(e, t) {
    return this._ensureRangeThemeStyleMap(e).get(t);
  }
  _getSheetRowVisibleFuncSet(e, t) {
    this._rowVisibleFuncSet.has(e) || this._rowVisibleFuncSet.set(e, /* @__PURE__ */ new Map());
    const o = this._rowVisibleFuncSet.get(e);
    return o.has(t) || o.set(t, /* @__PURE__ */ new Set()), o.get(t);
  }
  _getSheetRowVisibleHasInit(e, t) {
    var o;
    return !!(this._rowVisibleFuncSet.has(e) && ((o = this._rowVisibleFuncSet.get(e)) != null && o.has(t)));
  }
  refreshSheetRowVisibleFuncSet(e, t) {
    const o = this._getSheetRowVisibleFuncSet(e, t);
    o.clear();
    const s = this._univerInstanceService.getUnit(e);
    if (s) {
      const r = s.getSheetBySheetId(t);
      if (r) {
        const i = r.getRowCount(), a = r.getRowManager();
        for (let u = 1; u <= i; u++)
          r.getRowVisible(u) ? a.getRowHeight(u) === 0 && o.add(u) : o.add(u);
      }
    }
  }
  _ensureZebraCrossingCache(e, t, o) {
    this._zebraCrossingCacheMap.has(e) || this._zebraCrossingCacheMap.set(e, /* @__PURE__ */ new Map());
    const s = this._zebraCrossingCacheMap.get(e);
    s.has(t) || s.set(t, /* @__PURE__ */ new Map());
    const r = s.get(t);
    return r.has(o) || r.set(o, new ka()), r.get(o);
  }
  /**
   * Register range theme styles
   * @param {string} themeName
   * @param {IRangeThemeRangeInfo} rangeInfo
   */
  registerRangeThemeRule(e, t) {
    const { unitId: o, subUnitId: s, range: r } = t, i = Ut(), a = this._ensureRangeThemeStyleRuleMap(o), u = this._ensureRTreeCollection(o);
    a.set(i, { rangeInfo: t, themeName: e }), u.insert({ unitId: o, sheetId: s, range: r, id: i }), this._getSheetRowVisibleHasInit(o, s) || this.refreshSheetRowVisibleFuncSet(o, s);
    const l = this._ensureZebraCrossingCache(o, s, i), c = this._getSheetRowVisibleFuncSet(o, s);
    l.refresh(r, (d) => !c.has(d));
  }
  getRegisteredRangeThemeStyle(e) {
    const { unitId: t, subUnitId: o, range: s } = e, r = this._ensureRTreeCollection(t), i = Array.from(r.bulkSearch([{ unitId: t, sheetId: o, range: s }]));
    if (i[0]) {
      const u = this._ensureRangeThemeStyleRuleMap(t).get(i[0]);
      if (u)
        return u.themeName;
    }
  }
  refreshZebraCrossingCacheBySheet(e, t) {
    this._zebraCrossingCacheMap.has(e) || this._zebraCrossingCacheMap.set(e, /* @__PURE__ */ new Map());
    const o = this._zebraCrossingCacheMap.get(e);
    o.has(t) || o.set(t, /* @__PURE__ */ new Map());
    const s = o.get(t);
    s && s.forEach((r, i) => {
      const u = this._ensureRangeThemeStyleRuleMap(e).get(i);
      u ? r.refresh(u.rangeInfo.range, (l) => !this._getSheetRowVisibleFuncSet(e, t).has(l)) : s.delete(i);
    });
  }
  removeRangeThemeRule(e, t) {
    const { unitId: o, subUnitId: s, range: r } = t, i = this._ensureRTreeCollection(o), a = Array.from(i.bulkSearch([{ unitId: o, sheetId: s, range: r }])), u = this._ensureRangeThemeStyleRuleMap(o);
    for (let l = 0; l < a.length; l++) {
      const c = u.get(a[l]);
      if (c && c.themeName === e) {
        u.delete(a[l]), i.remove({ unitId: o, sheetId: s, range: r, id: a[l] });
        const d = this._zebraCrossingCacheMap.get(o);
        if (d) {
          const h = d.get(s);
          h && h.delete(a[l]);
        }
        break;
      }
    }
  }
  registerDefaultRangeTheme(e) {
    this._defaultRangeThemeMap.set(e.getName(), e), this._rangeThemeMapChanged$.next({ type: "add", styleName: e.getName() });
  }
  unRegisterDefaultRangeTheme(e) {
    this._defaultRangeThemeMap.delete(e), this._rangeThemeMapChanged$.next({ type: "remove", styleName: e });
  }
  getRegisteredRangeThemes() {
    return Array.from(this._defaultRangeThemeMap.keys());
  }
  /**
   * Register custom range theme style.
   * @param {string} unitId The unit id.
   * @param {RangeThemeStyle} rangeThemeStyle The range theme style.
   */
  registerRangeThemeStyle(e, t) {
    this._ensureRangeThemeStyleMap(e).set(t.getName(), t), this._rangeThemeMapChanged$.next({ type: "add", styleName: t.getName() });
  }
  /**
   *  Unregister custom range theme style.
   * @param {string} unitId The unit id.
   * @param {string} name The name of the range theme style.
   */
  unregisterRangeThemeStyle(e, t) {
    this._ensureRangeThemeStyleMap(e).delete(t), this._rangeThemeMapChanged$.next({ type: "remove", styleName: t });
  }
  /**
   * Gets all custom register themes
   * @param {string} unitId Which unit to register the range theme style.
   * @return {string[]} The array of all custom registered themes.
   */
  getALLRegisteredTheme(e) {
    return Array.from(this._ensureRangeThemeStyleMap(e).keys());
  }
  getRangeThemeStyle(e, t) {
    return this._defaultRangeThemeMap.has(t) ? this._defaultRangeThemeMap.get(t) : this._ensureRangeThemeStyleMap(e).get(t);
  }
  getCellStyle(e, t, o, s) {
    const r = { startRow: o, startColumn: s, endRow: o, endColumn: s }, i = this._ensureRTreeCollection(e), a = Array.from(i.bulkSearch([{ unitId: e, sheetId: t, range: r }]));
    if (a[0]) {
      const l = this._ensureRangeThemeStyleRuleMap(e).get(a[0]);
      if (l) {
        const { rangeInfo: c, themeName: d } = l, h = o - c.range.startRow, g = s - c.range.startColumn, m = this.getRangeThemeStyle(e, d), C = this._ensureZebraCrossingCache(e, t, a[0]).getIsToggled(o);
        if (m)
          return m.getStyle(h, g, o === c.range.endRow, s === c.range.endColumn, C);
      }
    }
  }
  _registerIntercept() {
    this.disposeWithMe(this._sheetInterceptorService.intercept(ht.CELL_CONTENT, {
      id: zs,
      effect: he.Style,
      handler: (e, t, o) => {
        const { row: s, col: r, unitId: i, subUnitId: a } = t, u = this.getCellStyle(i, a, s, r);
        if (u) {
          const l = !e || e === t.rawData ? { ...t.rawData } : e;
          return l.themeStyle = u, o(l);
        }
        return o(e);
      }
    }));
  }
  toJson(e) {
    const t = this._ensureRangeThemeStyleRuleMap(e), o = this._ensureRangeThemeStyleMap(e);
    if (o.size === 0 && t.size === 0)
      return "{}";
    const s = {};
    t.forEach((i, a) => {
      s[a] = i;
    });
    const r = {};
    return o.forEach((i, a) => {
      r[a] = i.toJson();
    }), JSON.stringify({
      rangeThemeStyleRuleMap: s,
      rangeThemeStyleMapJson: r
    });
  }
  fromJSON(e, t) {
    const { rangeThemeStyleRuleMap: o, rangeThemeStyleMapJson: s } = t;
    o && Object.keys(o).forEach((r) => {
      const i = o[r], { themeName: a, rangeInfo: u } = i;
      a.startsWith("table") || (this.registerRangeThemeRule(a, u), this._ensureRTreeCollection(u.unitId).insert({ unitId: r, sheetId: u.subUnitId, range: u.range, id: r }));
    }), s && Object.keys(s).forEach((r) => {
      const i = s[r], a = new Rt(i.name);
      a.fromJson(i), this._ensureRangeThemeStyleMap(e).set(a.getName(), a);
    });
  }
  deleteUnitId(e) {
    this._rangeThemeStyleMap.delete(e), this._rangeThemeStyleRuleMap.delete(e), this._rTreeCollection.delete(e);
  }
  _initSnapshot() {
    this.disposeWithMe(this._resourceManagerService.registerPluginResource({
      toJson: (e) => this.toJson(e),
      parseJson: (e) => {
        if (!e)
          return {};
        try {
          return JSON.parse(e);
        } catch {
          return {};
        }
      },
      businesses: [B.UNIVER_SHEET],
      pluginName: Na,
      onLoad: (e, t) => {
        this.fromJSON(e, t);
      },
      onUnLoad: (e) => {
        this.deleteUnitId(e);
      }
    }));
  }
  dispose() {
    super.dispose(), this._rangeThemeStyleMap.clear(), this._rangeThemeStyleRuleMap.clear(), this._defaultRangeThemeMap.clear(), this._rTreeCollection.clear(), this._zebraCrossingCacheMap.clear(), this._rowVisibleFuncSet.clear();
  }
};
ve = Pa([
  Co(0, $(G)),
  Co(1, $(Rn)),
  Co(2, $(M))
], ve);
function Ks(n, e) {
  const { unitId: t } = e, o = t ? n.getUnit(t, B.UNIVER_SHEET) : n.getCurrentUnitOfType(B.UNIVER_SHEET);
  return o ? {
    workbook: o,
    unitId: o.getUnitId()
  } : null;
}
function P(n, e = {}) {
  const { unitId: t, subUnitId: o } = e, s = t ? n.getUnit(t, B.UNIVER_SHEET) : n.getCurrentUnitOfType(B.UNIVER_SHEET);
  if (!s) return null;
  const r = o ? s.getSheetBySheetId(o) : s.getActiveSheet(!0);
  return r ? {
    worksheet: r,
    workbook: s,
    unitId: s.getUnitId(),
    subUnitId: r.getSheetId()
  } : null;
}
function Qe(n, e) {
  const { unitId: t, subUnitId: o } = e, s = n.getUnit(t, B.UNIVER_SHEET);
  if (!s) return null;
  const r = s.getSheetBySheetId(o);
  return r ? {
    worksheet: r,
    workbook: s
  } : null;
}
const tn = {
  id: "sheet.mutation.set-worksheet-range-theme-style",
  type: v.MUTATION,
  handler: (n, e) => {
    const { unitId: t, subUnitId: o, range: s, themeName: r } = e, i = n.get(M), a = P(i), u = n.get(ve);
    return a ? (u.registerRangeThemeRule(r, { range: s, unitId: t, subUnitId: o }), !0) : !1;
  }
}, Oa = (n, e) => {
  const t = Qe(n.get(M), e);
  if (!t)
    throw new Error("[SetWorksheetRangeThemeStyleMutation]: worksheet is null error!");
  const { worksheet: o } = t;
  return {
    unitId: e.unitId,
    subUnitId: o.getSheetId(),
    range: e.range,
    themeName: e.themeName
  };
}, nn = {
  id: "sheet.mutation.remove-worksheet-range-theme-style",
  type: v.MUTATION,
  handler: (n, e) => {
    const { unitId: t, subUnitId: o, range: s, themeName: r } = e, i = n.get(M), a = P(i), u = n.get(ve);
    return a ? (u.removeRangeThemeRule(r, { range: s, unitId: t, subUnitId: o }), !0) : !1;
  }
}, Da = (n, e) => {
  const t = Qe(n.get(M), e);
  if (!t)
    throw new Error("[DeleteWorksheetRangeThemeStyleMutationFactory]: worksheet is null error!");
  const { worksheet: o } = t;
  return {
    unitId: e.unitId,
    subUnitId: o.getSheetId(),
    range: e.range,
    themeName: e.themeName
  };
}, $o = (n, e) => {
  if (n.get(M).getUniverSheetInstance(e.unitId) == null)
    throw new Error("universheet is null error!");
  return {
    unitId: e.unitId,
    subUnitId: e.subUnitId,
    range: e.range
  };
}, Ue = {
  id: "sheet.mutation.insert-row",
  type: v.MUTATION,
  handler: (n, e) => {
    var m;
    const { unitId: t, subUnitId: o, range: s, rowInfo: r } = e, a = n.get(M).getUniverSheetInstance(t);
    if (a == null)
      throw new Error("universheet is null error!");
    const u = a.getSheetBySheetId(o);
    if (u == null)
      throw new Error("worksheet is null error!");
    const l = u.getRowManager().getRowData(), c = {
      h: u.getConfig().defaultRowHeight,
      hd: 0
    }, d = s.startRow, h = s.endRow - s.startRow + 1;
    for (let f = d; f < d + h; f++)
      r ? Tn(f, (m = r[f - s.startRow]) != null ? m : c, l) : Tn(f, c, l);
    return u.setRowCount(u.getRowCount() + h), u.getCellMatrix().insertRows(s.startRow, h), !0;
  }
}, ro = (n, e) => {
  if (n.get(M).getUniverSheetInstance(e.unitId) == null)
    throw new Error("universheet is null error!");
  return {
    unitId: e.unitId,
    subUnitId: e.subUnitId,
    range: e.range
  };
}, ke = {
  id: "sheet.mutation.insert-col",
  type: v.MUTATION,
  handler: (n, e) => {
    var m;
    const o = n.get(M).getUniverSheetInstance(e.unitId);
    if (o == null)
      throw new Error("universheet is null error!");
    const s = o.getSheetBySheetId(e.subUnitId);
    if (!s) return !1;
    const r = s.getColumnManager(), { range: i, colInfo: a } = e, l = r.getColumnData(), c = i.startColumn, d = i.endColumn - i.startColumn + 1, h = s.getConfig().defaultColumnWidth;
    for (let f = c; f < c + d; f++) {
      const C = {
        w: h,
        hd: 0
      };
      a ? Tn(f, (m = a[f - i.startColumn]) != null ? m : C, l) : Tn(f, C, l);
    }
    return s.setColumnCount(s.getColumnCount() + i.endColumn - i.startColumn + 1), s.getCellMatrix().insertColumns(i.startColumn, d), !0;
  }
}, wt = {
  id: "sheet.mutation.move-range",
  type: v.MUTATION,
  handler: (n, e) => {
    const { from: t, to: o } = e;
    if (!t || !o)
      return !1;
    const r = n.get(M).getCurrentUnitForType(B.UNIVER_SHEET);
    if (!r)
      return !1;
    const i = r.getSheetBySheetId(e.from.subUnitId), a = r.getSheetBySheetId(e.to.subUnitId);
    if (!i || !a)
      return !1;
    const u = i.getCellMatrix(), l = a.getCellMatrix();
    return new Y(t.value).forValue((c, d, h) => {
      h == null ? u.realDeleteValue(c, d) : u.setValue(c, d, h);
    }), new Y(o.value).forValue((c, d, h) => {
      h == null ? l.realDeleteValue(c, d) : l.setValue(c, d, h);
    }), !0;
  }
};
function Aa(n, e) {
  const { unitId: t, subUnitId: o, sourceRange: s, targetRange: r } = e, i = s.startRow > r.startRow, a = s.endRow - s.startRow + 1;
  return i ? {
    unitId: t,
    subUnitId: o,
    sourceRange: N.clone(r),
    targetRange: {
      ...s,
      endRow: s.endRow + a,
      startRow: s.startRow + a
    }
  } : {
    unitId: t,
    subUnitId: o,
    targetRange: N.clone(s),
    sourceRange: {
      ...r,
      endRow: r.endRow - a,
      startRow: r.startRow - a
    }
  };
}
const Je = {
  id: "sheet.mutation.move-rows",
  type: v.MUTATION,
  handler: (n, e) => {
    const { unitId: t, subUnitId: o, sourceRange: s, targetRange: r } = e, a = n.get(M).getUniverSheetInstance(t);
    if (!a)
      throw new Error("[MoveRowMutation] univerSheet is null!");
    const u = a.getSheetBySheetId(o);
    if (!u)
      throw new Error("[MoveRowMutation] worksheet is null!");
    const l = s.startRow, c = s.endRow - s.startRow + 1, d = r.startRow, h = u.getRowManager().getRowData();
    return Os(l, c, d, h), u.getCellMatrix().moveRows(l, c, d), !0;
  }
};
function xa(n, e) {
  const { unitId: t, subUnitId: o, sourceRange: s, targetRange: r } = e, i = s.startColumn > r.startColumn, a = s.endColumn - s.startColumn + 1;
  return i ? {
    unitId: t,
    subUnitId: o,
    sourceRange: N.clone(r),
    targetRange: {
      ...s,
      endColumn: s.endColumn + a,
      startColumn: s.startColumn + a
    }
  } : {
    unitId: t,
    subUnitId: o,
    targetRange: N.clone(s),
    sourceRange: {
      ...r,
      startColumn: r.startColumn - a,
      endColumn: r.endColumn - a
    }
  };
}
const qe = {
  id: "sheet.mutation.move-columns",
  type: v.MUTATION,
  handler: (n, e) => {
    const { unitId: t, subUnitId: o, sourceRange: s, targetRange: r } = e, a = n.get(M).getUniverSheetInstance(t);
    if (!a)
      throw new Error("[MoveColumnMutation] univerSheet is null!");
    const u = a.getSheetBySheetId(o);
    if (!u)
      throw new Error("[MoveColumnMutation] worksheet is null!");
    const l = s.startColumn, c = s.endColumn - s.startColumn + 1, d = r.startColumn, h = u.getColumnManager().getColumnData();
    return Os(l, c, d, h), u.getCellMatrix().moveColumns(l, c, d), !0;
  }
}, $a = (n, e) => {
  const s = e.getRowManager().getRowData(), r = {}, i = n.range, a = Ds(i.startRow, i.endRow, s), u = As(r, a);
  return {
    unitId: n.unitId,
    subUnitId: n.subUnitId,
    range: n.range,
    rowInfo: u
  };
}, Te = {
  id: "sheet.mutation.remove-rows",
  type: v.MUTATION,
  handler: (n, e) => {
    const o = n.get(M).getUniverSheetInstance(e.unitId);
    if (o == null)
      throw new Error("universheet is null error!");
    const s = o.getSheetBySheetId(e.subUnitId);
    if (!s) return !1;
    const r = e.range, a = s.getRowManager().getRowData();
    for (let c = r.startRow; c <= r.endRow; c++)
      s.getRowFiltered(c);
    const u = r.endRow - r.startRow + 1;
    return xs(r.startRow, u, a), s.getCellMatrix().removeRows(r.startRow, u), s.setRowCount(s.getRowCount() - u), !0;
  }
}, Wa = (n, e) => {
  const o = n.get(M).getUniverSheetInstance(e.unitId);
  if (o == null)
    throw new Error("universheet is null error!");
  const s = o.getSheetBySheetId(e.subUnitId);
  if (s == null)
    throw new Error("worksheet is null error!");
  const a = s.getColumnManager().getColumnData(), u = {}, l = e.range, c = Ds(l.startColumn, l.endColumn, a), d = As(u, c);
  return {
    unitId: e.unitId,
    subUnitId: e.subUnitId,
    range: e.range,
    colInfo: d
  };
}, Me = {
  id: "sheet.mutation.remove-col",
  type: v.MUTATION,
  handler: (n, e) => {
    const o = n.get(M).getUniverSheetInstance(e.unitId);
    if (o == null)
      throw new Error("universheet is null error!");
    const s = o.getSheetBySheetId(e.subUnitId);
    if (!s) return !1;
    const r = e.range, a = s.getColumnManager().getColumnData(), u = r.endColumn - r.startColumn + 1;
    return xs(r.startColumn, u, a), s.setColumnCount(s.getColumnCount() - u), s.getCellMatrix().removeColumns(r.startColumn, u), !0;
  }
}, ye = (n, e) => {
  const o = n.get(M).getUniverSheetInstance(e.unitId);
  if (o == null)
    throw new Error("universheet is null error!");
  const s = o.getSheetBySheetId(e.subUnitId);
  if (s == null)
    throw new Error("worksheet is null error!");
  const i = s.getConfig().mergeData, a = e.ranges, u = [];
  for (let l = 0; l < a.length; l++)
    for (let c = i.length - 1; c >= 0; c--) {
      const d = i[c], h = a[l];
      N.intersects(d, h) && u.push(i[c]);
    }
  return {
    unitId: e.unitId,
    subUnitId: e.subUnitId,
    ranges: u
  };
}, oe = {
  id: "sheet.mutation.remove-worksheet-merge",
  type: v.MUTATION,
  handler: (n, e) => {
    const o = n.get(M).getUniverSheetInstance(e.unitId);
    if (o == null)
      throw new Error("universheet is null error!");
    const s = o.getSheetBySheetId(e.subUnitId);
    if (!s) return !1;
    const i = s.getConfig().mergeData, a = e.ranges;
    for (let u = 0; u < a.length; u++)
      for (let l = i.length - 1; l >= 0; l--) {
        const c = i[l], d = a[u];
        N.intersects(c, d) && i.splice(l, 1);
      }
    return s.getSpanModel().rebuild(i), !0;
  }
}, Va = (n) => {
  const { order: e } = n, t = {};
  return Object.keys(e).forEach((o) => {
    t[e[Number(o)]] = Number(o);
  }), {
    ...n,
    order: t
  };
}, Nn = {
  id: "sheet.mutation.reorder-range",
  type: v.MUTATION,
  handler: (n, e) => {
    const { subUnitId: t, unitId: o, range: s, order: r } = e, u = n.get(M).getUnit(o).getSheetBySheetId(t);
    if (!u)
      return !1;
    const l = new Y();
    X.foreach(s, (d, h) => {
      if (r.hasOwnProperty(d)) {
        const g = r[d], m = A.deepClone(u.getCellRaw(g, h));
        l.setValue(d, h, m);
      }
    });
    const c = u.getCellMatrix();
    return l.forValue((d, h, g) => {
      c.setValue(d, h, g);
    }), !0;
  }
};
function La(n, e) {
  if (n == null)
    return n;
  const t = A.deepClone(n);
  if (e == null)
    return t;
  const o = {};
  return "h" in e && (o.h = t.h), "ia" in e && (o.ia = t.ia), "ah" in e && (o.ah = t.ah), "hd" in e && (o.hd = t.hd), "s" in e && (o.s = t.s), "custom" in e && (o.custom = t.custom), o;
}
function Ha(n, e) {
  if (n == null)
    return n;
  const t = A.deepClone(n);
  if (e == null)
    return t;
  const o = {};
  return "w" in e && (o.w = t.w), "hd" in e && (o.hd = t.hd), "s" in e && (o.s = t.s), "custom" in e && (o.custom = t.custom), o;
}
const Fa = (n, e) => {
  const { unitId: t, subUnitId: o, columnData: s } = n, r = {}, i = e.getColumnManager();
  for (const a in s) {
    const u = s[a], l = i.getColumn(Number(a));
    r[a] = Ha(l, u);
  }
  return {
    unitId: t,
    subUnitId: o,
    columnData: r
  };
}, zt = {
  id: "sheet.mutation.set-col-data",
  type: v.MUTATION,
  handler: (n, e) => {
    const { columnData: t } = e, o = n.get(M), s = P(o, e);
    if (!s) return !1;
    const { worksheet: r } = s, i = r.getColumnManager();
    for (const a in t) {
      const u = t[a];
      if (u == null) {
        i.removeColumn(Number(a));
        continue;
      }
      const l = i.getColumnOrCreate(Number(a));
      Object.assign(l, u);
    }
    return !0;
  }
}, Ba = (n, e) => {
  if (n.get(M).getUniverSheetInstance(e.unitId) == null)
    throw new Error("universheet is null error!");
  return {
    unitId: e.unitId,
    subUnitId: e.subUnitId,
    ranges: e.ranges
  };
}, on = {
  id: "sheet.mutation.set-col-hidden",
  type: v.MUTATION,
  handler: (n, e) => {
    const o = n.get(M).getUniverSheetInstance(e.unitId);
    if (!o)
      return !1;
    const s = o.getSheetBySheetId(e.subUnitId).getColumnManager();
    for (let r = 0; r < e.ranges.length; r++) {
      const i = e.ranges[r];
      for (let a = i.startColumn; a < i.endColumn + 1; a++) {
        const u = s.getColumnOrCreate(a);
        u != null && (u.hd = te.TRUE);
      }
    }
    return !0;
  }
}, ja = (n, e) => {
  if (n.get(M).getUniverSheetInstance(e.unitId) == null)
    throw new Error("universheet is null error!");
  return {
    unitId: e.unitId,
    subUnitId: e.subUnitId,
    ranges: e.ranges
  };
}, sn = {
  id: "sheet.mutation.set-col-visible",
  type: v.MUTATION,
  handler: (n, e) => {
    const o = n.get(M).getUniverSheetInstance(e.unitId);
    if (!o)
      return !1;
    const s = o.getSheetBySheetId(e.subUnitId).getColumnManager();
    for (let r = 0; r < e.ranges.length; r++) {
      const i = e.ranges[r];
      for (let a = i.startColumn; a < i.endColumn + 1; a++) {
        const u = s.getColumnOrCreate(a);
        u != null && (u.hd = te.FALSE);
      }
    }
    return !0;
  }
}, Gt = {
  id: "sheet.mutation.set-gridlines-color",
  type: v.MUTATION,
  handler: (n, e) => {
    const t = P(n.get(M), e);
    if (!t) return !1;
    const { worksheet: o } = t, s = o.getConfig();
    return s.gridlinesColor = e.color, !0;
  }
};
function za(n, e, t) {
  var i;
  const o = n.getStyleByCell(e);
  o == null && delete e.s, typeof t.s == "string" && (t.s = n.get(t.s));
  const s = En(o, t.s ? t.s : null);
  s && (A.removeNull(s), Object.entries(s).forEach(([a, u]) => {
    typeof u == "object" && u !== null && Object.keys(u).length === 0 && delete s[a];
  })), A.isEmptyObject(s) ? delete e.s : e.s = n.setValue(s);
  const r = t.v ? `${t.v}\r
` : "";
  !t.p && e.p && (r && r !== ((i = e.p.body) == null ? void 0 : i.dataStream) ? delete e.p : Ja(e.p, t.s ? t.s : null));
}
function Ga(n, e) {
  if (!e || !Object.keys(e).length)
    return n;
  const t = A.deepClone(n != null ? n : {});
  for (const o in e)
    o === "bd" ? t[o] = Ka(t[o] || {}, e[o]) : o in t || (t[o] = null);
  return t;
}
function Ka(n, e) {
  if (!e || !Object.keys(e).length)
    return n;
  for (const t in e)
    t in n || (n[t] = null);
  return n;
}
function En(n, e, t = !1) {
  if (e === null) return e;
  if (e === void 0) return n;
  const o = A.deepClone(n) || {};
  for (const s in e)
    t && ["bd", "tr", "td", "ht", "vt", "tb", "pd", "bg"].includes(s) || (s in o && s === "bd" ? o[s] = Object.assign(o[s], e[s]) : o[s] = e[s]);
  return "cl" in o && ("ul" in o && o.ul && (o.ul.cl = o.cl), "ol" in o && o.ol && (o.ol.cl = o.cl), "st" in o && o.st && (o.st.cl = o.cl)), o;
}
function Js(n, e) {
  return n.some((t) => t.startIndex === e) ? Js(n, e + 1) : e;
}
function Ja(n, e) {
  var i;
  if (n.body == null)
    return;
  Array.isArray(n.body.textRuns) || (n.body.textRuns = []);
  let t = 0;
  const o = [], s = ((i = n.body) == null ? void 0 : i.paragraphs) || [];
  for (const a of n.body.textRuns) {
    const { st: u, ed: l, ts: c = {} } = a;
    if (t < u) {
      const h = {
        st: t,
        ed: u
      }, g = En({}, e, !0);
      g && A.removeNull(g), A.isEmptyObject(g) || (h.ts = g), o.push(h);
    }
    const d = En(c, e, !0);
    d && A.removeNull(d), A.isEmptyObject(d) ? delete a.ts : a.ts = d, o.push(a), t = Js(s, l);
  }
  const r = n.body.dataStream.endsWith(`\r
`) ? n.body.dataStream.length - 2 : n.body.dataStream.length;
  if (t < r) {
    const a = {
      st: t,
      ed: r
    }, u = En({}, e, !0);
    u && A.removeNull(u), A.isEmptyObject(u) || (a.ts = u), o.push(a);
  }
  n.body.textRuns = vi(o);
}
function ds(n, e) {
  return e.v === void 0 || e.v === null ? e.v : n === re.NUMBER ? Number(e.v) : n === re.BOOLEAN ? qa(e.v) ? 1 : 0 : n === re.STRING || n === re.FORCE_STRING ? `${e.v}` : e.v;
}
function qa(n) {
  if (typeof n == "string") {
    if (n.toUpperCase() === "TRUE")
      return !0;
    if (n.toUpperCase() === "FALSE")
      return !1;
    if (Mi(n)) {
      if (Number(n) === 0)
        return !1;
      if (Number(n) === 1)
        return !0;
    }
  }
  if (typeof n == "number") {
    if (n === 0)
      return !1;
    if (n === 1)
      return !0;
  }
  return typeof n == "boolean" ? n : null;
}
function Ya(n) {
  return n == null ? null : (n.f === void 0 && (n.f = null), n.si === void 0 && (n.si = null), n.p === void 0 && (n.p = null), n.v === void 0 && (n.v = null), n.t === void 0 && (n.t = null), n.s === void 0 && (n.s = null), n.custom === void 0 && (n.custom = null), n);
}
const Ae = (n, e) => {
  const { unitId: t, subUnitId: o, cellValue: s } = e, i = n.get(M).getUniverSheetInstance(t);
  if (i == null)
    throw new Error("workbook is null error!");
  const a = i.getSheetBySheetId(o);
  if (a == null)
    throw new Error("worksheet is null error!");
  const u = a.getCellMatrix(), l = i.getStyles(), c = new Y();
  return new Y(s).forValue((h, g, m) => {
    const f = A.deepClone(u == null ? void 0 : u.getValue(h, g)) || {}, C = l.getStyleByCell(f), p = l.getStyleByCell(m);
    f.s = Ga(C, p), c.setValue(h, g, Ya(f));
  }), {
    ...e,
    options: {},
    cellValue: c.getMatrix()
  };
}, ee = {
  id: "sheet.mutation.set-range-values",
  type: v.MUTATION,
  handler: (n, e) => {
    const { cellValue: t, subUnitId: o, unitId: s } = e, i = n.get(M).getUnit(s);
    if (!i)
      return !1;
    const a = i.getSheetBySheetId(o);
    if (!a)
      return !1;
    const u = a.getCellMatrix(), l = i.getStyles();
    return new Y(t).forValue((d, h, g) => {
      if (!g)
        u.realDeleteValue(d, h);
      else {
        let m = u.getValue(d, h) || {};
        m = Za(g, m, l), A.isEmptyObject(m) ? u.realDeleteValue(d, h) : u.setValue(d, h, m);
      }
    }), !0;
  }
}, Xa = /* @__PURE__ */ new Set(["f", "p", "si", "custom", "ref"]);
function Za(n, e, t) {
  const o = ua(t, n, e);
  return Object.keys(n).forEach((s) => {
    const r = s;
    if (Xa.has(r)) {
      const i = n[r];
      Qa(e, r, i);
    } else r === "v" ? n.v !== void 0 && (e.v = ds(o, n)) : r === "s" && za(t, e, n);
  }), e.v !== void 0 && (e.t = o, e.v = ds(o, e)), e.v === null && (delete e.t, delete e.v), e;
}
function Qa(n, e, t) {
  t === void 0 || (t === null ? delete n[e] : n[e] = t);
}
const eu = (n, e) => {
  const { unitId: t, subUnitId: o, rowData: s } = n, r = {}, i = e.getRowManager();
  for (const a in s) {
    const u = s[a], l = i.getRow(Number(a));
    r[a] = La(l, u);
  }
  return {
    unitId: t,
    subUnitId: o,
    rowData: r
  };
}, Kt = {
  id: "sheet.mutation.set-row-data",
  type: v.MUTATION,
  handler: (n, e) => {
    const { rowData: t } = e, o = n.get(M), s = P(o, e);
    if (!s) return !1;
    const { worksheet: r } = s, i = r.getRowManager();
    for (const a in t) {
      const u = t[a];
      if (u == null) {
        i.removeRow(Number(a));
        continue;
      }
      const l = i.getRowOrCreate(Number(a));
      Object.assign(l, u);
    }
    return !0;
  }
}, tu = (n, e) => {
  if (n.get(M).getUniverSheetInstance(e.unitId) == null)
    throw new Error("universheet is null error!");
  return {
    unitId: e.unitId,
    subUnitId: e.subUnitId,
    ranges: e.ranges
  };
}, kt = {
  id: "sheet.mutation.set-row-visible",
  type: v.MUTATION,
  handler: (n, e) => {
    const o = n.get(M).getUniverSheetInstance(e.unitId);
    if (o == null)
      throw new Error("universheet is null error!");
    const s = o.getSheetBySheetId(e.subUnitId).getRowManager();
    for (let r = 0; r < e.ranges.length; r++) {
      const i = e.ranges[r];
      for (let a = i.startRow; a < i.endRow + 1; a++) {
        const u = s.getRowOrCreate(a);
        u != null && (u.hd = 0);
      }
    }
    return !0;
  }
}, nu = (n, e) => {
  if (n.get(M).getUniverSheetInstance(e.unitId) == null)
    throw new Error("universheet is null error!");
  return {
    unitId: e.unitId,
    subUnitId: e.subUnitId,
    ranges: e.ranges
  };
}, Tt = {
  id: "sheet.mutation.set-row-hidden",
  type: v.MUTATION,
  handler: (n, e) => {
    const o = n.get(M).getUniverSheetInstance(e.unitId);
    if (o == null)
      throw new Error("universheet is null error!");
    const s = o.getSheetBySheetId(e.subUnitId).getRowManager();
    for (let r = 0; r < e.ranges.length; r++) {
      const i = e.ranges[r];
      for (let a = i.startRow; a < i.endRow + 1; a++) {
        const u = s.getRowOrCreate(a);
        u != null && (u.hd = 1);
      }
    }
    return !0;
  }
}, qs = (n, e) => {
  const { unitId: t, subUnitId: o, ranges: s } = n, r = {}, i = e.getColumnManager();
  for (let a = 0; a < s.length; a++) {
    const u = s[a];
    for (let l = u.startColumn; l < u.endColumn + 1; l++)
      r[l] = i.getColumnWidth(l);
  }
  return {
    unitId: t,
    subUnitId: o,
    ranges: s,
    colWidth: r
  };
}, mt = {
  id: "sheet.mutation.set-worksheet-col-width",
  type: v.MUTATION,
  handler: (n, e) => {
    const t = n.get(M), o = P(t, e);
    if (!o) return !1;
    const { worksheet: s } = o, r = s.getColumnManager(), i = e.ranges;
    for (let a = 0; a < i.length; a++) {
      const u = i[a];
      for (let l = u.startColumn; l < u.endColumn + 1; l++)
        s.getColVisible(l) && (typeof e.colWidth == "number" ? r.setColumnWidth(l, e.colWidth) : A.isDefine(e.colWidth[l]) && r.setColumnWidth(l, e.colWidth[l]));
    }
    return !0;
  }
}, ou = (n, e) => {
  const t = Qe(n.get(M), e);
  if (!t)
    throw new Error("[SetWorksheetColumnCountUndoMutationFactory]: worksheet is null error!");
  return {
    unitId: e.unitId,
    subUnitId: e.subUnitId,
    columnCount: t.worksheet.getColumnCount()
  };
}, Jt = {
  id: "sheet.mutation.set-worksheet-column-count",
  type: v.MUTATION,
  handler: (n, e) => {
    const t = n.get(M), o = Qe(t, e);
    return o ? (o.worksheet.setColumnCount(e.columnCount), !0) : !1;
  }
}, qt = {
  id: "sheet.mutation.set-worksheet-default-style",
  type: v.MUTATION,
  handler: (n, e) => {
    const { defaultStyle: t } = e, o = n.get(M), s = P(o);
    if (!s) return !1;
    const { worksheet: r } = s;
    return r ? (r.setDefaultCellStyle(t), !0) : !1;
  }
}, su = (n, e) => {
  const t = Qe(n.get(M), e);
  if (!t)
    throw new Error("[SetWorksheetDefaultStyleMutationFactory]: worksheet is null error!");
  const { worksheet: o } = t;
  return {
    unitId: e.unitId,
    subUnitId: o.getSheetId(),
    defaultStyle: o.getDefaultCellStyle()
  };
}, ru = (n, e) => {
  const t = Qe(n.get(M), e);
  if (!t)
    throw new Error("[SetWorksheetRowCountUndoMutationFactory]: worksheet is null error!");
  return {
    unitId: e.unitId,
    subUnitId: e.subUnitId,
    rowCount: t.worksheet.getRowCount()
  };
}, Yt = {
  id: "sheet.mutation.set-worksheet-row-count",
  type: v.MUTATION,
  handler: (n, e) => {
    const t = n.get(M), o = Qe(t, e);
    return o ? (o.worksheet.setRowCount(e.rowCount), !0) : !1;
  }
}, Ys = (n, e) => {
  var a, u;
  const { unitId: t, subUnitId: o, ranges: s } = n, r = {}, i = e.getRowManager();
  for (const { startRow: l, endRow: c } of s)
    for (let d = l; d < c + 1; d++)
      r[d] = (u = (a = i.getRow(d)) == null ? void 0 : a.h) != null ? u : e.getConfig().defaultRowHeight;
  return {
    unitId: t,
    subUnitId: o,
    ranges: s,
    rowHeight: r
  };
}, Wo = (n, e) => {
  var a;
  const { unitId: t, subUnitId: o, ranges: s } = n, r = {}, i = e.getRowManager();
  for (const { startRow: u, endRow: l } of s)
    for (let c = u; c <= l; c++)
      r[c] = (a = i.getRow(c)) == null ? void 0 : a.ia;
  return {
    unitId: t,
    subUnitId: o,
    ranges: s,
    autoHeightInfo: r
  };
}, gh = (n, e) => {
  var a, u;
  const { unitId: t, subUnitId: o, rowsAutoHeightInfo: s } = n, r = [], i = e.getRowManager();
  for (const l of s) {
    const { row: c } = l;
    r.push({
      row: c,
      autoHeight: (u = (a = i.getRow(c)) == null ? void 0 : a.ah) != null ? u : e.getConfig().defaultRowHeight
    });
  }
  return {
    unitId: t,
    subUnitId: o,
    rowsAutoHeightInfo: r
  };
}, nt = {
  id: "sheet.mutation.set-worksheet-row-height",
  type: v.MUTATION,
  handler: (n, e) => {
    const { ranges: t, rowHeight: o } = e, s = n.get(M), r = P(s, e);
    if (!r) return !1;
    const { worksheet: i } = r, a = i.getRowManager();
    for (const { startRow: u, endRow: l } of t)
      for (let c = u; c <= l; c++)
        typeof o == "number" ? a.setRowHeight(c, o) : A.isDefine(o[c]) && a.setRowHeight(c, o[c]);
    return !0;
  }
}, Fe = {
  id: "sheet.mutation.set-worksheet-row-is-auto-height",
  type: v.MUTATION,
  handler: (n, e) => {
    var a;
    const { ranges: t, autoHeightInfo: o } = e, s = n.get(M), r = P(s, e);
    if (!r) return !1;
    const i = r.worksheet.getRowManager();
    for (const { startRow: u, endRow: l } of t)
      for (let c = u; c <= l; c++) {
        const d = i.getRowOrCreate(c);
        typeof o == "number" ? d.ia = o : d.ia = (a = o[c]) != null ? a : void 0;
      }
    return !0;
  }
}, Xs = {
  id: "sheet.mutation.set-worksheet-row-auto-height",
  type: v.MUTATION,
  handler: (n, e) => {
    const { rowsAutoHeightInfo: t } = e, o = n.get(M), s = P(o, e);
    if (!s) return !1;
    const r = s.worksheet.getRowManager();
    for (const { row: i, autoHeight: a } of t) {
      const u = r.getRowOrCreate(i);
      u.ah = a;
    }
    return !0;
  }
}, Xt = {
  id: "sheet.mutation.toggle-gridlines",
  type: v.MUTATION,
  handler: (n, e) => {
    const t = P(n.get(M), e);
    if (!t) return !1;
    const { worksheet: o } = t, s = o.getConfig();
    return s.showGridlines = e.showGridlines, !0;
  }
}, Cn = {
  id: "sheet.operation.set-worksheet-active",
  type: v.OPERATION,
  handler: (n, e) => {
    const t = n.get(M).getUniverSheetInstance(e.unitId);
    if (!t) return !1;
    const o = t.getWorksheets();
    for (const [, s] of o)
      if (s.getSheetId() === e.subUnitId)
        return t.setActiveSheet(s), !0;
    return !1;
  }
};
var iu = /* @__PURE__ */ ((n) => (n.SET_WORKSHEET_ROW_HEIGHT = "sheet.mutation.set-worksheet-row-height", n.SET_WORKSHEET_ROW_IS_AUTO_HEIGHT = "sheet.mutation.set-worksheet-row-is-auto-height", n.SET_WORKSHEET_ROW_AUTO_HEIGHT = "sheet.mutation.set-worksheet-row-auto-height", n.SET_WORKSHEET_COL_WIDTH = "sheet.mutation.set-worksheet-col-width", n.SET_WORKSHEET_ACTIVE = "sheet.operation.set-worksheet-active", n.MOVE_ROWS = "sheet.mutation.move-rows", n.MOVE_COLUMNS = "sheet.mutation.move-columns", n.SET_COL_HIDDEN = "sheet.mutation.set-col-hidden", n.SET_COL_VISIBLE = "sheet.mutation.set-col-visible", n.SET_ROW_HIDDEN = "sheet.mutation.set-row-hidden", n.SET_ROW_VISIBLE = "sheet.mutation.set-row-visible", n.INSERT_COL = "sheet.mutation.insert-col", n.INSERT_ROW = "sheet.mutation.insert-row", n.REMOVE_COL = "sheet.mutation.remove-col", n.REMOVE_ROW = "sheet.mutation.remove-rows", n.TOGGLE_GRIDLINES = "sheet.mutation.toggle-gridlines", n.SET_GRIDLINES_COLOR = "sheet.mutation.set-gridlines-color", n))(iu || {}), au = /* @__PURE__ */ ((n) => (n.SET_RANGE_VALUES = "sheet.mutation.set-range-values", n.MOVE_RANGE = "sheet.mutation.move-range", n.REMOVE_WORKSHEET_MERGE = "sheet.mutation.remove-worksheet-merge", n.ADD_WORKSHEET_MERGE = "sheet.mutation.add-worksheet-merge", n.REORDER_RANGE = "sheet.mutation.reorder-range", n.SET_WORKSHEET_DEFAULT_STYLE = "sheet.mutation.set-worksheet-default-style", n.SET_ROW_DATA = "sheet.mutation.set-row-data", n.SET_COL_DATA = "sheet.mutation.set-col-data", n.SET_WORKSHEET_RANGE_THEME_STYLE = "sheet.mutation.set-worksheet-range-theme-style", n.DELETE_WORKSHEET_RANGE_THEME_STYLE = "sheet.mutation.delete-worksheet-range-theme-style", n))(au || {});
const mh = [
  nt.id,
  Fe.id,
  Xs.id,
  mt.id,
  Cn.id,
  Je.id,
  qe.id,
  on.id,
  sn.id,
  Tt.id,
  kt.id,
  ke.id,
  Ue.id,
  Me.id,
  Te.id,
  Xt.id,
  Gt.id,
  Yt.id,
  Jt.id
], fh = [
  ee.id,
  wt.id,
  oe.id,
  ne.id,
  Nn.id,
  qt.id,
  Kt.id,
  zt.id,
  tn.id,
  nn.id
];
function Rh(n) {
  switch (n.id) {
    case "sheet.mutation.set-range-values": {
      const e = n.params, t = new Y(e.cellValue).getDataRange();
      return t.endRow === -1 ? [] : e.cellValue ? [{
        unitId: e.unitId,
        subUnitId: e.subUnitId,
        range: t
      }] : [];
    }
    case "sheet.mutation.move-range": {
      const e = n.params;
      return [{
        unitId: e.unitId,
        subUnitId: e.from.subUnitId,
        range: new Y(e.from.value).getRange()
      }, {
        unitId: e.unitId,
        subUnitId: e.to.subUnitId,
        range: new Y(e.to.value).getRange()
      }];
    }
    case "sheet.mutation.remove-worksheet-merge": {
      const e = n.params;
      return e.ranges.map((t) => ({
        unitId: e.unitId,
        subUnitId: e.subUnitId,
        range: t
      }));
    }
    case "sheet.mutation.add-worksheet-merge": {
      const e = n.params;
      return e.ranges.map((t) => ({
        unitId: e.unitId,
        subUnitId: e.subUnitId,
        range: t
      }));
    }
    case "sheet.mutation.reorder-range": {
      const e = n.params;
      return [{
        unitId: e.unitId,
        subUnitId: e.subUnitId,
        range: e.range
      }];
    }
    case "sheet.mutation.set-worksheet-default-style": {
      const e = n.params;
      return [{
        unitId: e.unitId,
        subUnitId: e.subUnitId,
        range: { startRow: 0, endRow: Number.MAX_SAFE_INTEGER, startColumn: 0, endColumn: Number.MAX_SAFE_INTEGER }
      }];
    }
    case "sheet.mutation.set-row-data": {
      const e = n.params, t = Object.keys(e.rowData).map(Number);
      return t.length === 0 ? [] : [{
        unitId: e.unitId,
        subUnitId: e.subUnitId,
        range: {
          startRow: Math.min(...t),
          endRow: Math.max(...t),
          startColumn: 0,
          endColumn: Number.MAX_SAFE_INTEGER
        }
      }];
    }
    case "sheet.mutation.set-col-data": {
      const e = n.params, t = Object.keys(e.columnData).map(Number);
      return t.length === 0 ? [] : [{
        unitId: e.unitId,
        subUnitId: e.subUnitId,
        range: {
          startRow: 0,
          endRow: Number.MAX_SAFE_INTEGER,
          startColumn: Math.min(...t),
          endColumn: Math.max(...t)
        }
      }];
    }
    case "sheet.mutation.set-worksheet-range-theme-style":
    case "sheet.mutation.delete-worksheet-range-theme-style": {
      const e = n.params;
      return [{
        unitId: e.unitId,
        subUnitId: e.subUnitId,
        range: e.range
      }];
    }
    default:
      return [];
  }
}
function Ch(n, e) {
  switch (n.id) {
    case "sheet.mutation.set-worksheet-row-height":
    case "sheet.mutation.set-worksheet-row-is-auto-height": {
      const t = n.params;
      return t.ranges.map((o) => ({
        unitId: t.unitId,
        subUnitId: t.subUnitId,
        range: {
          ...o,
          rangeType: j.ROW
        }
      }));
    }
    // Note: SET_WORKSHEET_ROW_AUTO_HEIGHT has no ranges
    case "sheet.mutation.set-worksheet-row-auto-height": {
      const t = n.params;
      return t.rowsAutoHeightInfo.map((o) => ({
        unitId: t.unitId,
        subUnitId: t.subUnitId,
        range: {
          startRow: o.row,
          endRow: o.row,
          startColumn: 0,
          endColumn: e - 1,
          rangeType: j.ROW
        }
      }));
    }
    case "sheet.mutation.set-worksheet-col-width": {
      const t = n.params;
      return t.ranges.map((o) => ({
        unitId: t.unitId,
        subUnitId: t.subUnitId,
        range: {
          ...o,
          rangeType: j.COLUMN
        }
      }));
    }
    case "sheet.mutation.move-rows":
    case "sheet.mutation.move-columns": {
      const t = n.params;
      return [{
        unitId: t.unitId,
        subUnitId: t.subUnitId,
        range: t.targetRange
      }, {
        unitId: t.unitId,
        subUnitId: t.subUnitId,
        range: t.sourceRange
      }];
    }
    case "sheet.mutation.set-col-hidden":
    case "sheet.mutation.set-col-visible": {
      const t = n.params;
      return t.ranges.map((o) => ({
        unitId: t.unitId,
        subUnitId: t.subUnitId,
        range: {
          ...o,
          rangeType: j.COLUMN
        }
      }));
    }
    case "sheet.mutation.set-row-hidden":
    case "sheet.mutation.set-row-visible": {
      const t = n.params;
      return t.ranges.map((o) => ({
        unitId: t.unitId,
        subUnitId: t.subUnitId,
        range: {
          ...o,
          rangeType: j.ROW
        }
      }));
    }
    case "sheet.mutation.insert-col": {
      const t = n.params;
      return [{
        unitId: t.unitId,
        subUnitId: t.subUnitId,
        range: {
          ...t.range,
          rangeType: j.COLUMN
        }
      }];
    }
    case "sheet.mutation.insert-row": {
      const t = n.params;
      return [{
        unitId: t.unitId,
        subUnitId: t.subUnitId,
        range: {
          ...t.range,
          rangeType: j.ROW
        }
      }];
    }
    case "sheet.mutation.remove-col": {
      const t = n.params;
      return [{
        unitId: t.unitId,
        subUnitId: t.subUnitId,
        range: {
          ...t.range,
          rangeType: j.COLUMN
        }
      }];
    }
    case "sheet.mutation.remove-rows": {
      const t = n.params;
      return [{
        unitId: t.unitId,
        subUnitId: t.subUnitId,
        range: {
          ...t.range,
          rangeType: j.ROW
        }
      }];
    }
    case "sheet.mutation.toggle-gridlines":
    case "sheet.mutation.set-gridlines-color":
      return [];
    default:
      return [];
  }
}
function hs(n) {
  return n == null ? !1 : n.v !== void 0 && n.v !== null && n.v !== "" || n.p !== void 0;
}
function io(n, e) {
  return n && n.spanAnchor ? hs(e.getValue(n.spanAnchor.startRow, n.spanAnchor.startColumn)) : hs(n);
}
function uu(n, e, t, o, s) {
  const r = n.getCellMatrix(), i = n.getSpanModel().getMergedCellRange(e, t, o, s), a = new Y();
  return r.forValue((u, l) => {
    const c = r.getValue(u, l);
    c && a.setValue(u, l, c);
  }), i.forEach((u) => {
    const { startColumn: l, startRow: c, endColumn: d, endRow: h } = u;
    _i(c, h, l, d).forEach((g, m) => {
      g === c && m === l && a.setValue(g, m, {
        ...r.getValue(g, m),
        spanAnchor: { startRow: c, endRow: h, startColumn: l, endColumn: d }
      }), (g !== c || m !== l) && (a.realDeleteValue(g, m), a.setValue(g, m, {
        spanAnchor: { startRow: c, endRow: h, startColumn: l, endColumn: d }
      }));
    });
  }), a;
}
function lu(n, e, t, o) {
  const { startRow: s, startColumn: r, endRow: i } = n;
  let a = null, u = !1;
  for (let l = s; l <= i; l++) {
    const c = e.getValue(l, r - t);
    if (u = u || io(c, e), !o && u)
      break;
    c && c.spanAnchor && (a ? a = {
      startRow: Math.min(c.spanAnchor.startRow, a.startRow),
      startColumn: Math.min(c.spanAnchor.startColumn, a.startColumn),
      endRow: Math.max(c.spanAnchor.endRow, a.endRow),
      endColumn: Math.max(c.spanAnchor.endColumn, a.endColumn)
    } : a = {
      startRow: c.spanAnchor.startRow,
      startColumn: c.spanAnchor.startColumn,
      endRow: c.spanAnchor.endRow,
      endColumn: c.spanAnchor.endColumn
    });
  }
  return u ? (n.startColumn = n.startColumn - t, {
    spanAnchor: a,
    hasValue: !0,
    range: n
  }) : {
    spanAnchor: null,
    hasValue: !1,
    range: n
  };
}
function cu(n, e, t, o) {
  const { startRow: s, endColumn: r, endRow: i } = n;
  let a = null, u = !1;
  for (let l = s; l <= i; l++) {
    const c = e.getValue(l, r + t);
    if (u = u || io(c, e), !o && u)
      break;
    c && c.spanAnchor && (a ? a = {
      startRow: Math.min(c.spanAnchor.startRow, a.startRow),
      startColumn: Math.min(c.spanAnchor.startColumn, a.startColumn),
      endRow: Math.max(c.spanAnchor.endRow, a.endRow),
      endColumn: Math.max(c.spanAnchor.endColumn, a.endColumn)
    } : a = {
      startRow: c.spanAnchor.startRow,
      startColumn: c.spanAnchor.startColumn,
      endRow: c.spanAnchor.endRow,
      endColumn: c.spanAnchor.endColumn
    });
  }
  return u ? (n.endColumn = n.endColumn + t, {
    spanAnchor: a,
    hasValue: !0,
    range: n
  }) : {
    spanAnchor: null,
    hasValue: !1,
    range: n
  };
}
function du(n, e, t, o) {
  const { startRow: s, startColumn: r, endColumn: i } = n;
  let a = null, u = !1;
  for (let l = r; l <= i; l++) {
    const c = e.getValue(s - t, l);
    if (u = u || io(c, e), !o && u)
      break;
    c && c.spanAnchor && (a ? a = {
      startRow: Math.min(c.spanAnchor.startRow, a.startRow),
      startColumn: Math.min(c.spanAnchor.startColumn, a.startColumn),
      endRow: Math.max(c.spanAnchor.endRow, a.endRow),
      endColumn: Math.max(c.spanAnchor.endColumn, a.endColumn)
    } : a = {
      startRow: c.spanAnchor.startRow,
      startColumn: c.spanAnchor.startColumn,
      endRow: c.spanAnchor.endRow,
      endColumn: c.spanAnchor.endColumn
    });
  }
  return u ? (n.startRow = n.startRow - t, {
    spanAnchor: a,
    hasValue: !0,
    range: n
  }) : {
    spanAnchor: null,
    hasValue: !1,
    range: n
  };
}
function hu(n, e, t, o) {
  const { startColumn: s, endColumn: r, endRow: i } = n;
  let a = null, u = !1;
  for (let l = s; l <= r; l++) {
    const c = e.getValue(i + t, l);
    if (u = u || io(c, e), !o && u)
      break;
    c && c.spanAnchor && (a ? a = {
      startRow: Math.min(c.spanAnchor.startRow, a.startRow),
      startColumn: Math.min(c.spanAnchor.startColumn, a.startColumn),
      endRow: Math.max(c.spanAnchor.endRow, a.endRow),
      endColumn: Math.max(c.spanAnchor.endColumn, a.endColumn)
    } : a = {
      startRow: c.spanAnchor.startRow,
      startColumn: c.spanAnchor.startColumn,
      endRow: c.spanAnchor.endRow,
      endColumn: c.spanAnchor.endColumn
    });
  }
  return u ? (n.endRow = n.endRow + t, {
    spanAnchor: a,
    hasValue: !0,
    range: n
  }) : {
    spanAnchor: null,
    hasValue: !1,
    range: n
  };
}
function ph(n, e, t) {
  const o = t.getMaxRows(), s = t.getMaxColumns(), r = uu(t, 0, 0, o - 1, s - 1), i = t.getSnapshot().mergeData.length > 0, { left: a, right: u, up: l, down: c } = e;
  let d = !0, h = { ...n };
  const g = [];
  for (; d; ) {
    if (d = !1, l && h.startRow !== 0) {
      const { hasValue: m, range: f, spanAnchor: C } = du(h, r, 1, i);
      if (C && g.push(C), m) {
        h = f, d = !0;
        continue;
      }
    }
    if (c && h.endRow !== o - 1) {
      const { hasValue: m, range: f, spanAnchor: C } = hu(h, r, 1, i);
      if (C && g.push(C), m) {
        h = f, d = !0;
        continue;
      }
    }
    if (a && h.startColumn !== 0) {
      const { hasValue: m, range: f, spanAnchor: C } = lu(h, r, 1, i);
      if (C && g.push(C), m) {
        h = f, d = !0;
        continue;
      }
    }
    if (u && h.endColumn !== s - 1) {
      const { hasValue: m, range: f, spanAnchor: C } = cu(h, r, 1, i);
      if (C && g.push(C), m) {
        h = f, d = !0;
        continue;
      }
    }
  }
  return g.length > 0 && (h = N.union(h, ...g)), h;
}
const gu = (n) => {
  const e = new Y();
  return n.forEach((t) => {
    X.foreach(t, (o, s) => {
      e.setValue(o, s, 1);
    });
  }), e.forValue((t, o) => {
    const s = e.getValue(t - 1, o);
    s && e.setValue(t, o, s + 1);
  }), e;
}, mu = (n) => {
  const e = n;
  return e.forValue((t, o) => {
    const s = n.getValue(t - 1, o);
    s && e.setValue(t, o, s + 1);
  }), e;
}, gs = (n) => {
  const e = {
    area: 0
  }, t = (o, s) => e.area < o ? (e.area = o, e.range = s, !0) : !1;
  return n.forValue((o, s, r) => {
    let i = 1, a = r;
    t(i * a, {
      startRow: o - a + 1,
      endRow: o,
      startColumn: s,
      endColumn: s
    });
    const u = {
      startRow: o - a + 1,
      endRow: o,
      startColumn: 0,
      endColumn: s
    };
    for (let l = s - 1; l >= 0 && n.getValue(o, l); l--) {
      a = Math.min(n.getValue(o, l) || 0, a), i++;
      const c = a * i;
      u.startColumn = l, u.startRow = o - a + 1, t(c, u);
    }
  }), e;
}, fu = (n, e) => {
  X.foreach(e, (t, o) => {
    n.realDeleteValue(t, o);
  });
  for (let t = e.startColumn; t <= e.endColumn; t++) {
    const o = e.endRow + 1;
    if (n.getValue(o, t) > 0) {
      n.setValue(o, t, 1);
      let r = o + 1;
      for (; n.getValue(r, t) > 0; )
        n.setValue(r, t, n.getValue(r - 1, t) + 1), r++;
    }
  }
  return n;
}, Zs = (n) => {
  const e = [];
  let t = gs(n);
  for (; t.area > 0; )
    t.range && (e.push(t.range), fu(n, t.range)), t = gs(n);
  return e;
}, Qs = (n) => {
  const e = gu(n);
  return Zs(e);
};
class Sh {
  constructor() {
    R(this, "_matrix", new Y());
  }
  add(...e) {
    return e.forEach((t) => {
      X.foreach(t, (o, s) => {
        this._matrix.setValue(o, s, 1);
      });
    }), this;
  }
  subtract(...e) {
    return e.forEach((t) => {
      X.foreach(t, (o, s) => {
        this._matrix.realDeleteValue(o, s);
      });
    }), this;
  }
  merge() {
    const e = mu(this._matrix);
    return Zs(e);
  }
}
const wh = 1.5, Ih = "rgba(255, 255, 255, 0.01)";
function vh(n) {
  const { rangeWithCoord: e, primaryWithCoord: t, style: o } = n, s = {
    range: {
      startRow: e.startRow,
      startColumn: e.startColumn,
      endRow: e.endRow,
      endColumn: e.endColumn,
      rangeType: e.rangeType,
      unitId: e.unitId,
      sheetId: e.sheetId
    },
    primary: null,
    style: o
  };
  return t != null && (s.primary = Ru(t)), s;
}
function Ru(n) {
  const { actualRow: e, actualColumn: t, isMerged: o, isMergedMainCell: s } = n, { startRow: r, startColumn: i, endRow: a, endColumn: u } = n.mergeInfo;
  return {
    actualRow: e,
    actualColumn: t,
    isMerged: o,
    isMergedMainCell: s,
    startRow: r,
    startColumn: i,
    endRow: a,
    endColumn: u
  };
}
var Cu = /* @__PURE__ */ ((n) => (n[n.Tab = 1] = "Tab", n[n.Comma = 2] = "Comma", n[n.Semicolon = 4] = "Semicolon", n[n.Space = 8] = "Space", n[n.Custom = 16] = "Custom", n))(Cu || {});
class pu {
  constructor() {
    R(this, "_tabCount", 0);
    R(this, "_commaCount", 0);
    R(this, "_semicolonCount", 0);
    R(this, "_spaceCount", 0);
  }
  add(e) {
    switch (e) {
      case "	":
        this._tabCount++;
        break;
      case ",":
        this._commaCount++;
        break;
      case ";":
        this._semicolonCount++;
        break;
      case " ":
        this._spaceCount++;
        break;
    }
  }
  update(e) {
    e && typeof e == "string" && (e.includes("	") && this._tabCount++, e.includes(",") && this._commaCount++, e.includes(";") && this._semicolonCount++, e.trim().includes(" ") && this._spaceCount++);
  }
  getDelimiter() {
    const e = Math.max(this._tabCount, this._commaCount, this._semicolonCount, this._spaceCount);
    return e === 0 || e === this._tabCount ? 1 : e === this._commaCount ? 2 : e === this._semicolonCount ? 4 : e === this._spaceCount ? 8 : 1;
  }
}
function Su(n, e, t) {
  const o = [];
  t !== void 0 && (n & 16) > 0 && o.push(t), (n & 1) > 0 && o.push("	"), (n & 2) > 0 && o.push(","), (n & 4) > 0 && o.push(";"), (n & 8) > 0 && o.push(" ");
  let s = "";
  for (const i of o)
    s += wu(i);
  let r = "[".concat(s, "]");
  return e && (r += "+"), new RegExp(r);
}
function wu(n) {
  return n.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
const Iu = (n) => {
  var t;
  return ((t = n.body) == null ? void 0 : t.dataStream.replace(/\r\n$/, "")) || "";
};
function vu(n) {
  if (n != null) {
    if (n.p)
      return Iu(n.p);
    if (n.v && typeof n.v == "string")
      return n.v;
    if (n.t && (n.t === re.FORCE_STRING || n.t === re.STRING))
      return String(n.v);
  }
}
function Mu(n, e, t, o, s = !1) {
  const r = X.transformRange(e, n), { startColumn: i, startRow: a, endColumn: u, endRow: l } = r;
  if (i !== u)
    throw new Error("The range must be in the same column.");
  if (t && (t & 16) > 0 && (o === void 0 || o.length !== 1))
    throw new Error("The custom delimiter must a character.");
  const c = t === void 0, d = c ? new pu() : null, h = [];
  for (let I = a; I <= l; I++) {
    const S = n.getCell(I, i), y = vu(S);
    h.push(y), d && d.update(y);
  }
  const g = c ? d.getDelimiter() : t, m = Su(g, s, o);
  let f = -1, C = 0, p = 0;
  const w = [];
  for (const I of h) {
    if (I !== void 0) {
      const S = String(I).split(m);
      f < 0 ? f = S.length : f = Math.max(f, S.length), w.push(S), C = p;
    } else
      w.push(void 0);
    p++;
  }
  return {
    rs: w,
    maxLength: f === -1 ? 0 : f,
    lastRow: C
  };
}
const _u = (n, e, t = "") => n.reduce(
  (o, s) => {
    const r = s && s[e];
    return typeof r != "string" ? (console.warn(s, `${e} is not string`), o) : (r ? (o[r] || (o[r] = []), o[r].push(s)) : o[t].push(s), o);
  },
  {}
), yu = (n = 0) => {
  let e = n;
  return function() {
    return e++;
  };
};
function bu(n) {
  return n == null ? !1 : n.v !== void 0 && n.v !== null && n.v !== "" || n.p !== void 0;
}
function Mh(n, e) {
  for (let t = n.startRow; t <= n.endRow; t++)
    for (let o = n.startColumn; o <= n.endColumn; o++) {
      const s = e.getCell(t, o);
      if (bu(s))
        return { startRow: t, startColumn: o, endRow: t, endColumn: o };
    }
  return null;
}
function er(n) {
  const e = new Y();
  return n.forEach((t) => {
    const { startRow: o, startColumn: s, endRow: r, endColumn: i } = t;
    for (let a = o; a <= r; a++)
      for (let u = s; u <= i; u++)
        e.setValue(a, u, null);
  }), e.clone();
}
function Eu(n) {
  const e = new Y();
  return n.forEach((t) => {
    const { startRow: o, startColumn: s, endRow: r, endColumn: i } = t;
    for (let a = o; a <= r; a++)
      for (let u = s; u <= i; u++)
        e.setValue(a, u, {
          v: null,
          p: null,
          f: null,
          si: null,
          custom: null
        });
  }), e.clone();
}
function Uu(n) {
  const e = new Y();
  return n.forEach((t) => {
    const { startRow: o, startColumn: s, endRow: r, endColumn: i } = t;
    for (let a = o; a <= r; a++)
      for (let u = s; u <= i; u++)
        e.setValue(a, u, {
          s: null
        });
  }), e.clone();
}
function ku(n, e, t, o) {
  const s = e.get(M), r = t ? s.getUnit(t, B.UNIVER_SHEET) : s.getCurrentUnitForType(B.UNIVER_SHEET), i = o ? r == null ? void 0 : r.getSheetBySheetId(o) : r == null ? void 0 : r.getActiveSheet();
  if (!i)
    return null;
  const { startRow: a, endRow: u, startColumn: l, endColumn: c } = n, d = [], h = [];
  for (let g = a; g <= u; g++)
    i.getRowFiltered(g) || d.push(g);
  for (let g = l; g <= c; g++)
    h.push(g);
  return {
    rows: d,
    cols: h
  };
}
function ao(n, e, t, o) {
  const s = [], r = [];
  for (const h of n) {
    const g = ku(h, e, t, o);
    g && (s.push(...g.rows), r.push(...g.cols));
  }
  const i = Array.from(new Set(s)).sort((h, g) => h - g), a = Array.from(new Set(r)).sort((h, g) => h - g), u = [];
  function l(h) {
    const g = [];
    let m = h[0];
    for (let f = 1; f < h.length; f++)
      h[f] !== h[f - 1] + 1 && (g.push([m, h[f - 1]]), m = h[f]);
    return g.push([m, h[h.length - 1]]), g;
  }
  const c = l(i), d = l(a);
  for (const [h, g] of c)
    for (const [m, f] of d)
      u.push({
        startRow: h,
        endRow: g,
        startColumn: m,
        endColumn: f
      });
  return u;
}
var Tu = /* @__PURE__ */ ((n) => (n.OthersCanView = "othersCanView", n.NoOneElseCanView = "noOneElseCanView", n))(Tu || {}), Pu = /* @__PURE__ */ ((n) => (n.DesignedUserCanEdit = "designedUserCanEdit", n.OnlyMe = "onlyMe", n))(Pu || {});
class me {
  constructor() {
    /**
     * Map<unitId, Map<subUnitId, Map<ruleId, IRangeProtectionRule>>>
     */
    R(this, "_model", /* @__PURE__ */ new Map());
    R(this, "_ruleChange$", new De());
    R(this, "ruleChange$", this._ruleChange$.asObservable());
    R(this, "_ruleRefresh$", new De());
    R(this, "ruleRefresh$", this._ruleRefresh$.asObservable());
    R(this, "_rangeRuleInitStateChange", new gt(!1));
    R(this, "rangeRuleInitStateChange$", this._rangeRuleInitStateChange.asObservable());
  }
  dispose() {
    this._ruleChange$.complete(), this._ruleRefresh$.complete();
  }
  ruleRefresh(e) {
    this._ruleRefresh$.next(e);
  }
  getRangeRuleInitState() {
    return this._rangeRuleInitStateChange.value;
  }
  changeRuleInitState(e) {
    this._rangeRuleInitStateChange.next(e);
  }
  addRule(e, t, o) {
    this._ensureRuleMap(e, t).set(o.id, o), this._ruleChange$.next({ unitId: e, subUnitId: t, rule: o, type: "add" });
  }
  deleteRule(e, t, o) {
    var r, i, a, u;
    const s = (i = (r = this._model.get(e)) == null ? void 0 : r.get(t)) == null ? void 0 : i.get(o);
    s && ((u = (a = this._model.get(e)) == null ? void 0 : a.get(t)) == null || u.delete(o), this._ruleChange$.next({ unitId: e, subUnitId: t, rule: s, type: "delete" }));
  }
  setRule(e, t, o, s) {
    var i, a;
    const r = this.getRule(e, t, o);
    r && ((a = (i = this._model.get(e)) == null ? void 0 : i.get(t)) == null || a.set(o, s), this._ruleChange$.next({ unitId: e, subUnitId: t, oldRule: r, rule: s, type: "set" }));
  }
  getRule(e, t, o) {
    var s, r;
    return (r = (s = this._model.get(e)) == null ? void 0 : s.get(t)) == null ? void 0 : r.get(o);
  }
  getSubunitRuleList(e, t) {
    var s;
    return [...(((s = this._model.get(e)) == null ? void 0 : s.get(t)) || /* @__PURE__ */ new Map()).values()];
  }
  getSubunitRuleListLength(e, t) {
    var s;
    const o = (s = this._model.get(e)) == null ? void 0 : s.get(t);
    return o ? o.size : 0;
  }
  _ensureRuleMap(e, t) {
    let o = this._model.get(e);
    o || (o = /* @__PURE__ */ new Map(), this._model.set(e, o));
    let s = o.get(t);
    return s || (s = /* @__PURE__ */ new Map(), o.set(t, s)), s;
  }
  toObject() {
    const e = {};
    return [...this._model.keys()].forEach((o) => {
      const s = this._model.get(o), r = [...s.keys()];
      e[o] = {}, r.forEach((i) => {
        const a = s.get(i);
        e[o][i] = [...a.values()];
      });
    }), e;
  }
  fromObject(e) {
    const t = /* @__PURE__ */ new Map();
    Object.keys(e).forEach((o) => {
      const s = e[o], r = /* @__PURE__ */ new Map();
      Object.keys(s).forEach((i) => {
        const a = s[i].reduce((u, l) => (u.set(l.id, l), u), /* @__PURE__ */ new Map());
        r.set(i, a);
      }), t.set(o, r);
    }), this._model = t;
  }
  deleteUnitModel(e) {
    this._model.delete(e);
  }
  createRuleId(e, t) {
    let o = Ut(4);
    const s = this._ensureRuleMap(e, t);
    for (; s.has(o); )
      o = Ut(4);
    return o;
  }
  getTargetByPermissionId(e, t) {
    const o = this._model.get(e);
    if (!o) return null;
    for (const [s, r] of o)
      for (const i of r.values())
        if (i.permissionId === t)
          return [e, s];
    return null;
  }
}
const _h = (n, e) => {
  const t = n.get(me), o = e.ruleIds.map((r) => t.getRule(e.unitId, e.subUnitId, r)).filter((r) => !!r);
  return { id: Be.id, params: { subUnitId: e.subUnitId, unitId: e.unitId, rules: o } };
}, st = {
  id: "sheet.mutation.delete-range-protection",
  type: v.MUTATION,
  handler: (n, e) => {
    const { unitId: t, subUnitId: o, ruleIds: s } = e, r = n.get(me);
    return s.forEach((i) => {
      r.deleteRule(t, o, i);
    }), !0;
  }
}, yh = (n) => {
  const e = { ...n, ruleIds: n.rules.map((t) => t.id) };
  return { id: st.id, params: e };
}, Be = {
  id: "sheet.mutation.add-range-protection",
  type: v.MUTATION,
  handler: (n, e) => {
    const { unitId: t, subUnitId: o, rules: s } = e, r = n.get(me);
    return s.forEach((i) => {
      r.addRule(t, o, i);
    }), !0;
  }
}, Nu = {
  type: v.COMMAND,
  id: "sheet.command.add-range-protection",
  async handler(n, e) {
    if (!e)
      return !1;
    const t = n.get(E), o = n.get(V), s = n.get(me), { rule: r, permissionId: i } = e, { unitId: a, subUnitId: u, ranges: l, description: c, viewState: d, editState: h } = r, g = [{
      ranges: l,
      permissionId: i,
      id: s.createRuleId(a, u),
      description: c,
      unitType: r.unitType,
      unitId: a,
      subUnitId: u,
      viewState: d,
      editState: h
    }];
    if (await t.executeCommand(Be.id, {
      unitId: a,
      subUnitId: u,
      rules: g
    })) {
      const f = [{ id: Be.id, params: { unitId: a, subUnitId: u, rules: g } }], C = [{ id: st.id, params: { unitId: a, subUnitId: u, ruleIds: g.map((p) => p.id) } }];
      o.pushUndoRedo({
        unitID: a,
        redoMutations: f,
        undoMutations: C
      });
    }
    return !0;
  }
};
var we = /* @__PURE__ */ ((n) => (n[n.MOVE_START = 0] = "MOVE_START", n[n.MOVING = 1] = "MOVING", n[n.MOVE_END = 2] = "MOVE_END", n[n.ONLY_SET = 3] = "ONLY_SET", n))(we || {});
class Ou extends ue {
  constructor(t) {
    super();
    /**
     * Selection data model for each worksheet.
     */
    R(this, "_worksheetSelections", /* @__PURE__ */ new Map());
    R(this, "_selectionMoveStart$", new De());
    R(this, "selectionMoveStart$", this._selectionMoveStart$.asObservable());
    R(this, "_selectionMoving$", new De());
    R(this, "selectionMoving$", this._selectionMoving$.asObservable());
    R(this, "_selectionMoveEnd$", new gt([]));
    R(this, "selectionMoveEnd$", this._selectionMoveEnd$.asObservable());
    R(this, "_selectionSet$", new gt([]));
    R(this, "selectionSet$", this._selectionSet$.asObservable());
    R(this, "selectionChanged$");
    R(this, "_beforeSelectionMoveEnd$", new gt([]));
    R(this, "beforeSelectionMoveEnd$", this._beforeSelectionMoveEnd$.asObservable());
    this._workbook = t, this.selectionChanged$ = jt(this._selectionMoveEnd$, this._selectionSet$);
  }
  dispose() {
    super.dispose(), this._beforeSelectionMoveEnd$.complete(), this._selectionMoveEnd$.complete(), this._selectionMoving$.complete(), this._selectionMoveStart$.complete(), this._selectionSet$.complete(), this._workbook = null;
  }
  addSelections(t, o) {
    const s = this.getSelectionsOfWorksheet(t);
    s.push(...o), this._selectionSet$.next(s);
  }
  /**
   * Set selectionDatas to _worksheetSelections, and emit selectionDatas by type.
   * @param sheetId
   * @param selectionDatas
   * @param type
   */
  setSelections(t, o = [], s) {
    switch (this.setSelectionsOfWorksheet(t, o), s) {
      case we.MOVE_START:
        this._selectionMoveStart$.next(o);
        break;
      case we.MOVING:
        this._selectionMoving$.next(o);
        break;
      case we.MOVE_END:
        this._beforeSelectionMoveEnd$.next(o), this._selectionMoveEnd$.next(o);
        break;
      case we.ONLY_SET: {
        this._selectionSet$.next(o);
        break;
      }
      default:
        this._selectionSet$.next(o);
        break;
    }
  }
  getCurrentSelections() {
    return this._getCurrentSelections();
  }
  /**
   * @deprecated use `getSelectionsOfWorksheet` instead.
   * @param sheetId
   * @returns
   */
  getSelectionOfWorksheet(t) {
    return this.getSelectionsOfWorksheet(t);
  }
  getSelectionsOfWorksheet(t) {
    return this._worksheetSelections.has(t) || this._worksheetSelections.set(t, []), this._worksheetSelections.get(t);
  }
  setSelectionsOfWorksheet(t, o) {
    this._worksheetSelections.set(t, [...o]);
  }
  deleteSheetSelection(t) {
    this._worksheetSelections.set(t, []);
  }
  /** Clear all selections in this workbook. */
  clear() {
    this._worksheetSelections.clear(), this._selectionSet$.next([]);
  }
  _getCurrentSelections() {
    return this.getSelectionsOfWorksheet(this._workbook.getActiveSheet().getSheetId());
  }
  getCurrentLastSelection() {
    const t = this._getCurrentSelections();
    return t[t.length - 1];
  }
}
var Du = Object.getOwnPropertyDescriptor, Au = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? Du(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, xu = (n, e) => (t, o) => e(t, o, n);
let z = class extends $s {
  constructor(e) {
    super();
    /**
     * Cache cell styles for current selections, key is `${row}_${column}`.
     */
    R(this, "_cellStylesCache", /* @__PURE__ */ new Map());
    /**
     * Selection Events, usually triggered when pointerdown in spreadsheet by selection render service after selectionModel has updated.
     */
    R(this, "selectionMoveStart$");
    /**
     * Selection Events, usually triggered when pointermove in spreadsheet by selection render service after selectionModel has updated.
     */
    R(this, "selectionMoving$");
    /**
     * Selection Events, usually triggered when pointerup in spreadsheet by selection render service after selectionModel has updated.
     */
    R(this, "selectionMoveEnd$");
    /**
     * Selection Events, usually triggered when changing unit.(focus in formula editor)
     */
    R(this, "selectionSet$");
    /**
     * Selection Events, merge moveEnd$ and selectionSet$
     */
    R(this, "selectionChanged$");
    R(this, "_workbookSelections", /* @__PURE__ */ new Map());
    this._instanceSrv = e, this._init();
  }
  get _currentSelectionPos() {
    const e = this._instanceSrv.getCurrentUnitForType(B.UNIVER_SHEET);
    if (!e) return null;
    const t = e.getActiveSheet();
    return { unitId: e.getUnitId(), sheetId: t.getSheetId() };
  }
  get currentSelectionParam() {
    return this._currentSelectionPos;
  }
  _init() {
    const e = this._instanceSrv.getCurrentTypeOfUnit$(B.UNIVER_SHEET).pipe(Ki(1), et(this.dispose$));
    this.selectionMoveStart$ = e.pipe().pipe(tt((t) => t ? this._ensureWorkbookSelection(t.getUnitId()).selectionMoveStart$ : Ie())).pipe(et(this.dispose$)), this.selectionMoving$ = e.pipe(tt((t) => t ? this._ensureWorkbookSelection(t.getUnitId()).selectionMoving$ : Ie())).pipe(et(this.dispose$)), this.selectionMoveEnd$ = e.pipe(tt((t) => t ? this._ensureWorkbookSelection(t.getUnitId()).selectionMoveEnd$ : Ie([]))).pipe(et(this.dispose$)), this.selectionSet$ = e.pipe(tt((t) => t ? this._ensureWorkbookSelection(t.getUnitId()).selectionSet$ : Ie([]))).pipe(et(this.dispose$)), this.selectionChanged$ = e.pipe(tt((t) => t ? this._ensureWorkbookSelection(t.getUnitId()).selectionChanged$ : Ie([]))).pipe(
      Ji((t, o) => t.length !== o.length ? !1 : t.length === 0 && o.length === 0 ? !0 : t.every((s, r) => JSON.stringify(s) === JSON.stringify(o[r]))),
      Bs(1)
    ).pipe(et(this.dispose$)), this.disposeWithMe(
      this._instanceSrv.getTypeOfUnitDisposed$(B.UNIVER_SHEET).pipe(et(this.dispose$)).subscribe((t) => {
        this._removeWorkbookSelection(t.getUnitId());
      })
    ), this.disposeWithMe(
      this.selectionChanged$.pipe(et(this.dispose$)).subscribe(() => {
        this._cellStylesCache.clear();
      })
    );
  }
  dispose() {
    super.dispose(), this._cellStylesCache.clear(), this._workbookSelections.forEach((e) => e.dispose()), this._workbookSelections.clear(), this.selectionMoveStart$ = Ie(null), this.selectionMoving$ = Ie(null), this.selectionMoveEnd$ = Ie([]), this.selectionSet$ = Ie(null), this.selectionChanged$ = Ie(null);
  }
  /**
   * Clear all selections in all workbooks.
   * invoked by prompt.controller
   */
  clear() {
    this._workbookSelections.forEach((e) => e.clear());
  }
  getCurrentSelections() {
    return this._getCurrentSelections();
  }
  getCurrentLastSelection() {
    const e = this._getCurrentSelections();
    return e == null ? void 0 : e[e.length - 1];
  }
  addSelections(e, t, o) {
    if (typeof e == "string") {
      this._ensureWorkbookSelection(e).addSelections(t, o);
      return;
    }
    const s = this._currentSelectionPos;
    if (!s)
      throw new Error("[SheetsSelectionsService]: cannot find current selection position!");
    const { unitId: r, sheetId: i } = s;
    this._ensureWorkbookSelection(r).addSelections(i, e);
  }
  setSelections(e, t, o, s) {
    if (typeof e == "string" && typeof t == "string") {
      const u = e;
      this._ensureWorkbookSelection(u).setSelections(
        t,
        o || [],
        s != null ? s : we.ONLY_SET
      );
      return;
    }
    const r = this._currentSelectionPos;
    if (!r)
      throw new Error("[SheetsSelectionsService]: cannot find current selection position!");
    const { unitId: i, sheetId: a } = r;
    if (typeof e == "object") {
      const u = e != null ? e : o, l = t != null ? t : we.ONLY_SET;
      this._ensureWorkbookSelection(i).setSelections(a, u, l);
    }
  }
  clearCurrentSelections() {
    this._getCurrentSelections().splice(0);
  }
  /**
   * Determine whether multiple current selections overlap
   *
   * @deprecated this should be extracted to an pure function
   */
  isOverlapping() {
    const e = this.getCurrentSelections();
    return e == null ? !1 : e.some(
      ({ range: t }, o) => e.some(({ range: s }, r) => o === r ? !1 : t.startRow <= s.endRow && t.endRow >= s.startRow && t.startColumn <= s.endColumn && t.endColumn >= s.startColumn)
    );
  }
  _getCurrentSelections() {
    const e = this._currentSelectionPos;
    if (!e)
      return [];
    const { unitId: t, sheetId: o } = e;
    return this._ensureWorkbookSelection(t).getSelectionsOfWorksheet(o);
  }
  getWorkbookSelections(e) {
    return this._ensureWorkbookSelection(e);
  }
  _ensureWorkbookSelection(e) {
    let t = this._workbookSelections.get(e);
    if (!t) {
      const o = this._instanceSrv.getUnit(e);
      if (!o)
        throw new Error(`[SheetsSelectionsService]: cannot resolve unit with id "${e}"!`);
      t = new Ou(o), this._workbookSelections.set(e, t);
    }
    return t;
  }
  _removeWorkbookSelection(e) {
    this._workbookSelections.delete(e);
  }
  /**
   * This method is used to get the common value of a specific cell style property in the current selections.
   * Used to determine the state related to color panels in the toolbar.
   * Because in Excel, only the color panels need to show the common color of the current selections, other properties based on the current selection primary cell.
   * Now only handles text color, fill color, border style, border color.
   */
  getCellStylesProperty(e) {
    var r;
    const t = (r = this._instanceSrv.getCurrentUnitForType(B.UNIVER_SHEET)) == null ? void 0 : r.getActiveSheet(), o = this.getCurrentSelections();
    if (!t || o.length === 0)
      return {
        isAllValuesSame: !1,
        value: null
      };
    let s = null;
    for (let i = 0; i < o.length; i++) {
      const a = o[i], { startRow: u, endRow: l, startColumn: c, endColumn: d } = a.range;
      for (let h = u; h <= l; h++)
        for (let g = c; g <= d; g++) {
          const m = `${h}_${g}`;
          let f;
          this._cellStylesCache.has(m) ? f = this._cellStylesCache.get(m) : (f = t.getComposedCellStyle(h, g), this._cellStylesCache.set(m, f));
          const C = f[e];
          if (s != null && !A.diffValue(s, C))
            return {
              isAllValuesSame: !1,
              value: null
            };
          s = C;
        }
    }
    return {
      isAllValuesSame: !0,
      value: s
    };
  }
};
z = Au([
  xu(0, M)
], z);
const bh = "DISABLE_NORMAL_SELECTIONS", Eh = "SELECTIONS_ENABLED", $u = "REF_SELECTIONS_ENABLED", Vo = {
  id: "sheet.command.clear-selection-all",
  type: v.COMMAND,
  handler: (n, e) => {
    var I;
    const t = n.get(M), o = n.get(E), s = n.get(z), r = n.get(V), i = n.get(G), a = t.getCurrentUnitForType(B.UNIVER_SHEET);
    if (!a) return !1;
    const u = (e == null ? void 0 : e.unitId) || a.getUnitId(), l = a.getActiveSheet();
    if (!l) return !1;
    const c = (e == null ? void 0 : e.subUnitId) || l.getSheetId(), d = (e == null ? void 0 : e.ranges) || ((I = s.getCurrentSelections()) == null ? void 0 : I.map((S) => S.range));
    if (!(d != null && d.length))
      return !1;
    const h = ao(d, n, u, c), g = [], m = [], f = {
      subUnitId: c,
      unitId: u,
      cellValue: er(h)
    }, C = Ae(
      n,
      f
    );
    g.push({
      id: ee.id,
      params: f
    }), m.push({
      id: ee.id,
      params: C
    });
    const p = i.onCommandExecute({ id: Vo.id });
    return g.push(...p.redos), m.unshift(...p.undos), L(g, o) ? (r.pushUndoRedo({
      // If there are multiple mutations that form an encapsulated project, they must be encapsulated in the same undo redo element.
      // Hooks can be used to hook the code of external controllers to add new actions.
      unitID: u,
      undoMutations: m,
      redoMutations: g
    }), !0) : !1;
  }
}, Lo = {
  id: "sheet.command.clear-selection-format",
  type: v.COMMAND,
  handler: (n, e) => {
    var I;
    const t = n.get(M), o = n.get(E), s = n.get(z), r = n.get(V), i = n.get(G), a = t.getCurrentUnitForType(B.UNIVER_SHEET);
    if (!a) return !1;
    const u = (e == null ? void 0 : e.unitId) || a.getUnitId(), l = a.getActiveSheet();
    if (!l) return !1;
    const c = (e == null ? void 0 : e.subUnitId) || l.getSheetId(), d = (e == null ? void 0 : e.ranges) || ((I = s.getCurrentSelections()) == null ? void 0 : I.map((S) => S.range));
    if (!(d != null && d.length))
      return !1;
    const h = ao(d, n, u, c), g = [], m = [], f = {
      subUnitId: c,
      unitId: u,
      cellValue: Uu(h)
    }, C = Ae(
      n,
      f
    );
    g.push({
      id: ee.id,
      params: f
    }), m.push({
      id: ee.id,
      params: C
    });
    const p = i.onCommandExecute({ id: Lo.id });
    return g.push(...p.redos), m.unshift(...p.undos), L(g, o) ? (r.pushUndoRedo({
      // If there are multiple mutations that form an encapsulated project, they must be encapsulated in the same undo redo element.
      // Hooks can be used to hook the code of external controllers to add new actions.
      unitID: u,
      undoMutations: m,
      redoMutations: g
    }), !0) : !1;
  }
};
function uo(n, e, t = !0) {
  const o = e.getMatrixWithMergedCells(...yi(n)), s = [];
  if (o.forValue((i, a, u) => {
    if (u.colSpan !== void 0 && u.rowSpan !== void 0) {
      const l = {
        startRow: i,
        startColumn: a,
        endRow: i + u.rowSpan - 1,
        endColumn: a + u.colSpan - 1
      };
      N.contains(n, l) || s.push(l);
    }
  }), s.length === 0)
    return n;
  const r = N.union(n, ...s);
  return t ? uo(r, e, t) : r;
}
function Uh(n, e, t) {
  let o = null;
  return t.getMatrixWithMergedCells(n, e, n, e).forValue((r, i, a) => (o = {
    actualRow: r,
    actualColumn: i,
    startRow: r,
    startColumn: i,
    isMerged: a.rowSpan !== void 0 || a.colSpan !== void 0,
    isMergedMainCell: a.rowSpan !== void 0 && a.colSpan !== void 0,
    endRow: r + (a.rowSpan !== void 0 ? a.rowSpan - 1 : 0),
    endColumn: i + (a.colSpan !== void 0 ? a.colSpan - 1 : 0),
    rangeType: j.NORMAL
  }, !1)), o || {
    actualColumn: e,
    actualRow: n,
    startRow: n,
    startColumn: e,
    endRow: n,
    endColumn: e,
    isMerged: !1,
    isMergedMainCell: !1,
    rangeType: j.NORMAL
  };
}
function kh(n, e, t) {
  const { startRow: o, startColumn: s, endRow: r, endColumn: i } = n;
  return Number.isNaN(o) && (n.startRow = 0), Number.isNaN(r) && (n.endRow = e - 1), Number.isNaN(s) && (n.startColumn = 0), Number.isNaN(i) && (n.endColumn = t - 1), n;
}
function _e(n, e) {
  const t = Number.isNaN(n.startRow) ? 0 : n.startRow, o = Number.isNaN(n.startColumn) ? 0 : n.startColumn, s = e.getMergedCell(t, o);
  return s ? {
    ...s,
    actualRow: t,
    actualColumn: o,
    rangeType: j.NORMAL,
    isMerged: !0,
    isMergedMainCell: !0
  } : {
    startRow: t,
    startColumn: o,
    endRow: n.startRow,
    endColumn: n.startColumn,
    actualRow: t,
    actualColumn: o,
    rangeType: j.NORMAL,
    isMerged: !1,
    isMergedMainCell: !1
  };
}
const ut = (n, e, t) => ({
  id: ie.id,
  params: {
    unitId: e.getUnitId(),
    subUnitId: t.getSheetId(),
    reveal: !0,
    selections: [{ range: n, primary: _e(n, t) }]
  }
});
function Th(n) {
  if (!n)
    return !1;
  const { range: e, primary: t } = n;
  return N.equals(e, t);
}
function Wu(n) {
  function e(t, o) {
    function s(r) {
      for (let i = r.startRow; i <= r.endRow; i++)
        if (!n.getRowFiltered(i))
          for (let a = r.startColumn; a <= r.endColumn; a++)
            o(i, a, r);
    }
    s(t);
  }
  return {
    forOperableEach: e
  };
}
const ms = (n) => n.id !== zs;
function Ct(n, e, t, o, s, r, i) {
  const a = {};
  for (let u = e; u <= t; u++)
    for (let l = o; l <= s; l++) {
      const c = r ? n.getCellWithFilteredInterceptors(i, l, us, ms) : n.getCellWithFilteredInterceptors(u, i, us, ms);
      !c || !c.s || (a[u] || (a[u] = {}), a[u][l] = { s: c.s });
    }
  return a;
}
var Vu = Object.getOwnPropertyDescriptor, Lu = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? Vu(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, Hu = (n, e) => (t, o) => e(t, o, n);
const Fu = Ao("sheets-formula.ref-selections.service");
let fs = class extends z {
  constructor(n) {
    super(n);
  }
  _init() {
    const n = this._getAliveWorkbooks$().pipe(et(this.dispose$));
    this.selectionMoveStart$ = n.pipe(tt((e) => jt(...e.map((t) => t.selectionMoveStart$)))), this.selectionMoving$ = n.pipe(tt((e) => jt(...e.map((t) => t.selectionMoving$)))), this.selectionMoveEnd$ = n.pipe(tt((e) => jt(...e.map((t) => t.selectionMoveEnd$)))), this.selectionSet$ = n.pipe(tt((e) => jt(...e.map((t) => t.selectionSet$))));
  }
  dispose() {
    super.dispose(), this.selectionMoveStart$ = Ie(null), this.selectionMoving$ = Ie(null), this.selectionMoveEnd$ = Ie(null), this.selectionSet$ = Ie(null), delete this._instanceSrv, this._workbookSelections.clear();
  }
  _getAliveWorkbooks$() {
    const n = this._instanceSrv.getAllUnitsForType(B.UNIVER_SHEET);
    n.forEach((t) => this._ensureWorkbookSelection(t.getUnitId()));
    const e = new gt(n);
    return this.disposeWithMe(this._instanceSrv.getTypeOfUnitAdded$(B.UNIVER_SHEET).subscribe((t) => {
      this._ensureWorkbookSelection(t.getUnitId()), e.next([...e.getValue(), t]);
    })), this.disposeWithMe(this._instanceSrv.getTypeOfUnitDisposed$(B.UNIVER_SHEET).subscribe((t) => {
      this._removeWorkbookSelection(t.getUnitId()), e.next(e.getValue().filter((o) => o !== t));
    })), e.pipe(yo((t) => t.map((o) => this._ensureWorkbookSelection(o.getUnitId()))));
  }
};
fs = Lu([
  Hu(0, M)
], fs);
function Bu(n, e) {
  const o = n.get(Ws).getContextValue($u);
  return n.get(o && !e ? Fu : z);
}
const ie = {
  id: "sheet.operation.set-selections",
  type: v.OPERATION,
  handler: (n, e) => {
    if (!e) return !1;
    const { selections: t, type: o, unitId: s, subUnitId: r } = e;
    return Bu(n).setSelections(s, r, [...t], o), !0;
  }
}, ju = {
  id: "sheet.command.select-range",
  type: v.COMMAND,
  handler: (n, e) => {
    if (!e) return !1;
    const { unitId: t, subUnit: o, range: s } = e, r = n.get(E), i = P(n.get(M), e);
    if (!i) return !1;
    const a = [{
      range: s,
      primary: _e(s, i.worksheet),
      style: null
    }];
    return r.syncExecuteCommand(ie.id, {
      unitId: t,
      subUnitId: o,
      selections: a
    });
  }
}, tr = "sheet.command.move-range", yt = {
  type: v.COMMAND,
  id: tr,
  // eslint-disable-next-line max-lines-per-function
  handler: async (n, e) => {
    var w, I;
    const t = n.get(E), o = n.get(V), s = n.get(M), r = n.get(Vt), i = n.get(it), a = n.get(G), u = P(s);
    if (!u || !await a.beforeCommandExecute({ id: yt.id, params: e }))
      return !1;
    const { worksheet: c, subUnitId: d, unitId: h } = u, g = Ho(
      n,
      { unitId: h, subUnitId: d, range: e.fromRange },
      { subUnitId: d, range: e.toRange }
    );
    if (g === null)
      return r.emit(i.t("sheets.info.acrossMergedCell")), !1;
    const m = a.onCommandExecute({
      id: yt.id,
      params: e
    }), f = [
      ...(w = m.preRedos) != null ? w : [],
      ...g.redos,
      ...m.redos,
      {
        id: ie.id,
        params: {
          unitId: h,
          subUnitId: d,
          selections: [{ range: e.toRange, primary: zu(e.fromRange, e.toRange, c) }],
          type: we.MOVE_END
        }
      }
    ], C = [
      ...(I = m.preUndos) != null ? I : [],
      ...g.undos,
      ...m.undos,
      {
        id: ie.id,
        params: {
          unitId: h,
          subUnitId: d,
          selections: [{ range: e.fromRange, primary: _e(e.fromRange, c) }],
          type: we.MOVE_END
        }
      }
    ];
    if (L(f, t).result) {
      const { undos: S, redos: y } = a.generateMutationsOfAutoHeight({
        unitId: h,
        subUnitId: d,
        ranges: [e.fromRange, e.toRange]
      }), b = a.afterCommandExecute({
        id: yt.id,
        params: e
      });
      return L([...b.redos, ...y], t), o.pushUndoRedo({
        unitID: h,
        undoMutations: [...C, ...b.undos, ...S],
        redoMutations: [...f, ...b.redos, ...y]
      }), !0;
    }
    return !1;
  }
};
function Ho(n, e, t, o = !1) {
  const s = [], r = [], { range: i, subUnitId: a, unitId: u } = e, { range: l, subUnitId: c } = t, h = n.get(M).getUniverSheetInstance(u), g = h == null ? void 0 : h.getSheetBySheetId(c), m = h == null ? void 0 : h.getSheetBySheetId(a), f = g == null ? void 0 : g.getCellMatrix(), C = m == null ? void 0 : m.getCellMatrix();
  if (g && m && f && C) {
    const p = uo(l, g, !1);
    if (!N.equals(l, p) && !o)
      return null;
    const w = new Y(), I = new Y(), S = new Y();
    X.foreach(i, (T, U) => {
      const D = C.getValue(T, U);
      if (w.setValue(T, U, A.deepClone(D)), D) {
        const H = h == null ? void 0 : h.getStyles().get(D.s);
        S.setValue(T, U, A.deepClone(H));
      }
      I.setValue(T, U, null);
    });
    const y = new Y(), b = new Y();
    X.foreach(l, (T, U) => {
      y.setValue(T, U, A.deepClone(f.getValue(T, U)));
    }), X.foreach(i, (T, U) => {
      const D = Vs(T, U), H = N.getRelativeRange(D, i), F = N.getPositionRange(H, l), K = A.deepClone(S.getValue(T, U)), J = A.deepClone(w.getValue(T, U));
      J && K && (J.s = K), b.setValue(F.startRow, F.startColumn, J);
    });
    const O = {
      fromRange: e.range,
      toRange: t.range,
      from: {
        value: I.getMatrix(),
        subUnitId: a
      },
      to: {
        value: b.getMatrix(),
        subUnitId: c
      },
      unitId: u
    }, k = {
      fromRange: t.range,
      toRange: e.range,
      from: {
        value: w.getMatrix(),
        subUnitId: a
      },
      to: {
        value: y.getMatrix(),
        subUnitId: c
      },
      unitId: u
    };
    s.push({ id: wt.id, params: O }), r.push({ id: wt.id, params: k });
  }
  return {
    redos: s,
    undos: r
  };
}
function zu(n, e, t) {
  const o = n.startRow, s = n.startColumn, r = t.getMergedCell(o, s), i = _e(e, t);
  if (r) {
    const a = r.endRow - r.startRow + 1, u = r.endColumn - r.startColumn + 1;
    i.endRow = i.startRow + a - 1, i.endColumn = i.startColumn + u - 1, i.actualRow = i.startRow, i.actualColumn = i.startColumn, i.isMerged = !1, i.isMergedMainCell = !0;
  }
  return i;
}
var On = /* @__PURE__ */ ((n) => (n[n.UNIVER_UNKNOWN = 0] = "UNIVER_UNKNOWN", n[n.UNIVER_DOC = 1] = "UNIVER_DOC", n[n.UNIVER_SHEET = 2] = "UNIVER_SHEET", n[n.UNIVER_SLIDE = 3] = "UNIVER_SLIDE", n[n.UNIVER_PROJECT = 4] = "UNIVER_PROJECT", n[n.UNRECOGNIZED = -1] = "UNRECOGNIZED", n))(On || {}), _ = /* @__PURE__ */ ((n) => (n[n.View = 0] = "View", n[n.Edit = 1] = "Edit", n[n.ManageCollaborator = 2] = "ManageCollaborator", n[n.Print = 3] = "Print", n[n.Duplicate = 4] = "Duplicate", n[n.Comment = 5] = "Comment", n[n.Copy = 6] = "Copy", n[n.Share = 7] = "Share", n[n.Export = 8] = "Export", n[n.MoveWorksheet = 9] = "MoveWorksheet", n[n.DeleteWorksheet = 10] = "DeleteWorksheet", n[n.HideWorksheet = 11] = "HideWorksheet", n[n.RenameWorksheet = 12] = "RenameWorksheet", n[n.CreateWorksheet = 13] = "CreateWorksheet", n[n.SetWorksheetStyle = 14] = "SetWorksheetStyle", n[n.EditWorksheetCell = 15] = "EditWorksheetCell", n[n.InsertHyperlink = 16] = "InsertHyperlink", n[n.Sort = 17] = "Sort", n[n.Filter = 18] = "Filter", n[n.PivotTable = 19] = "PivotTable", n[n.FloatImg = 20] = "FloatImg", n[n.History = 21] = "History", n[n.RwHgtClWdt = 22] = "RwHgtClWdt", n[n.ViemRwHgtClWdt = 23] = "ViemRwHgtClWdt", n[n.ViewFilter = 24] = "ViewFilter", n[n.MoveSheet = 25] = "MoveSheet", n[n.DeleteSheet = 26] = "DeleteSheet", n[n.HideSheet = 27] = "HideSheet", n[n.CopySheet = 28] = "CopySheet", n[n.RenameSheet = 29] = "RenameSheet", n[n.CreateSheet = 30] = "CreateSheet", n[n.SelectProtectedCells = 31] = "SelectProtectedCells", n[n.SelectUnProtectedCells = 32] = "SelectUnProtectedCells", n[n.SetCellStyle = 33] = "SetCellStyle", n[n.SetCellValue = 34] = "SetCellValue", n[n.SetRowStyle = 35] = "SetRowStyle", n[n.SetColumnStyle = 36] = "SetColumnStyle", n[n.InsertRow = 37] = "InsertRow", n[n.InsertColumn = 38] = "InsertColumn", n[n.DeleteRow = 39] = "DeleteRow", n[n.DeleteColumn = 40] = "DeleteColumn", n[n.EditExtraObject = 41] = "EditExtraObject", n[n.Delete = 42] = "Delete", n[n.RecoverHistory = 43] = "RecoverHistory", n[n.ViewHistory = 44] = "ViewHistory", n[n.CreatePermissionObject = 45] = "CreatePermissionObject", n[n.UNRECOGNIZED = -1] = "UNRECOGNIZED", n))(_ || {}), x = /* @__PURE__ */ ((n) => (n[n.Unkonwn = 0] = "Unkonwn", n[n.Workbook = 1] = "Workbook", n[n.Worksheet = 2] = "Worksheet", n[n.SelectRange = 3] = "SelectRange", n[n.Document = 4] = "Document", n[n.Slide = 5] = "Slide", n[n.UNRECOGNIZED = -1] = "UNRECOGNIZED", n))(x || {});
class Ve {
  constructor(e, t, o) {
    R(this, "type", x.SelectRange);
    R(this, "subType", _.Edit);
    R(this, "status", q.INIT);
    R(this, "value", !0);
    R(this, "id");
    R(this, "unitId");
    R(this, "subUnitId");
    R(this, "permissionId");
    this.unitId = e, this.subUnitId = t, this.permissionId = o, this.id = `${x.SelectRange}.${_.Edit}.${o}`;
  }
}
class Fo {
  constructor(e, t, o) {
    R(this, "type", x.SelectRange);
    R(this, "subType", _.View);
    R(this, "status", q.INIT);
    R(this, "value", !0);
    R(this, "id");
    R(this, "unitId");
    R(this, "subUnitId");
    R(this, "permissionId");
    this.unitId = e, this.subUnitId = t, this.permissionId = o, this.id = `${x.SelectRange}.${_.View}.${o}`;
  }
}
class nr {
  constructor(e) {
    R(this, "id");
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "subType", _.Comment);
    this.unitId = e, this.unitId = e, this.id = `${this.type}.${_.Comment}_${e}`;
  }
}
class or {
  constructor(e) {
    R(this, "id");
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "subType", _.Copy);
    this.unitId = e, this.unitId = e, this.id = `${this.type}.${_.Copy}_${e}`;
  }
}
class Gu {
  constructor(e) {
    R(this, "id");
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "subType", _.CopySheet);
    R(this, "status", q.INIT);
    this.unitId = e, this.unitId = e, this.id = `${this.type}.${_.CopySheet}_${e}`;
  }
}
class sr {
  constructor(e) {
    R(this, "id");
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "subType", _.CreatePermissionObject);
    this.unitId = e, this.unitId = e, this.id = `${this.type}.${_.CreatePermissionObject}_${e}`;
  }
}
class rr {
  constructor(e) {
    R(this, "id");
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "subType", _.CreateSheet);
    this.unitId = e, this.unitId = e, this.id = `${this.type}.${_.CreateSheet}_${e}`;
  }
}
class Ku {
  constructor(e) {
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.DeleteColumn);
    this.unitId = e, this.id = `${this.type}.${_.DeleteColumn}_${e}`;
  }
}
class Ju {
  constructor(e) {
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.DeleteRow);
    this.unitId = e, this.id = `${this.type}.${_.DeleteRow}_${e}`;
  }
}
class ir {
  constructor(e) {
    R(this, "id");
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "subType", _.DeleteSheet);
    this.unitId = e, this.unitId = e, this.id = `${this.type}.${_.DeleteSheet}_${e}`;
  }
}
class ar {
  constructor(e) {
    R(this, "id");
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "subType", _.Duplicate);
    this.unitId = e, this.unitId = e, this.id = `${this.type}.${_.Duplicate}_${e}`;
  }
}
class Pe {
  constructor(e) {
    R(this, "id");
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "subType", _.Edit);
    this.unitId = e, this.unitId = e, this.id = `${this.type}.${_.Edit}_${e}`;
  }
}
class ur {
  constructor(e) {
    R(this, "id");
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "subType", _.Export);
    this.unitId = e, this.unitId = e, this.id = `${this.type}.${_.Export}_${e}`;
  }
}
class Bo {
  constructor(e) {
    R(this, "id");
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "subType", _.HideSheet);
    this.unitId = e, this.unitId = e, this.id = `${this.type}.${_.HideSheet}_${e}`;
  }
}
class qu {
  constructor(e) {
    R(this, "id");
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "subType", _.History);
    this.unitId = e, this.unitId = e, this.id = `${this.type}.${_.History}_${e}`;
  }
}
class Yu {
  constructor(e) {
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.InsertColumn);
    this.unitId = e, this.id = `${this.type}.${_.InsertColumn}_${e}`;
  }
}
class Xu {
  constructor(e) {
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.InsertRow);
    this.unitId = e, this.id = `${this.type}.${_.InsertRow}_${e}`;
  }
}
class jo {
  constructor(e) {
    R(this, "id");
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "subType", _.ManageCollaborator);
    this.unitId = e, this.unitId = e, this.id = `${this.type}.${_.ManageCollaborator}_${e}`;
  }
}
class zo {
  constructor(e) {
    R(this, "id");
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "subType", _.MoveSheet);
    this.unitId = e, this.unitId = e, this.id = `${this.type}.${_.MoveSheet}_${e}`;
  }
}
class lr {
  constructor(e) {
    R(this, "id");
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "subType", _.Print);
    this.unitId = e, this.unitId = e, this.id = `${this.type}.${_.Print}_${e}`;
  }
}
class cr {
  constructor(e) {
    R(this, "id");
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "subType", _.RecoverHistory);
    this.unitId = e, this.unitId = e, this.id = `${this.type}.${_.RecoverHistory}_${e}`;
  }
}
class Go {
  constructor(e) {
    R(this, "id");
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "subType", _.RenameSheet);
    this.unitId = e, this.unitId = e, this.id = `${this.type}.${_.RenameSheet}_${e}`;
  }
}
class dr {
  constructor(e) {
    R(this, "id");
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "subType", _.Share);
    this.unitId = e, this.unitId = e, this.id = `${this.type}.${_.Share}_${e}`;
  }
}
class hr {
  constructor(e) {
    R(this, "id");
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "subType", _.View);
    this.unitId = e, this.unitId = e, this.id = `${this.type}.${_.View}_${e}`;
  }
}
class gr {
  constructor(e) {
    R(this, "id");
    R(this, "value", !0);
    R(this, "type", x.Workbook);
    R(this, "status", q.INIT);
    R(this, "subType", _.ViewHistory);
    this.unitId = e, this.unitId = e, this.id = `${this.type}.${_.ViewHistory}_${e}`;
  }
}
class mr {
  constructor(e, t) {
    R(this, "value", !0);
    R(this, "type", x.Worksheet);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.Copy);
    this.unitId = e, this.subUnitId = t, this.id = `${this.type}.${_.Copy}_${e}_${t}`;
  }
}
class fr {
  constructor(e, t) {
    R(this, "value", !0);
    R(this, "type", x.Worksheet);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.DeleteColumn);
    this.unitId = e, this.subUnitId = t, this.id = `${this.type}.${_.DeleteColumn}_${e}_${t}`;
  }
}
class Rr {
  constructor(e, t) {
    R(this, "value", !0);
    R(this, "type", x.Worksheet);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.Delete);
    this.unitId = e, this.subUnitId = t, this.id = `${this.type}.${_.Delete}_${e}_${t}`;
  }
}
class Cr {
  constructor(e, t) {
    R(this, "value", !0);
    R(this, "type", x.Worksheet);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.DeleteRow);
    this.unitId = e, this.subUnitId = t, this.id = `${this.type}.${_.DeleteRow}_${e}_${t}`;
  }
}
class Le {
  constructor(e, t) {
    R(this, "value", !0);
    R(this, "type", x.Worksheet);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.Edit);
    this.unitId = e, this.subUnitId = t, this.id = `${this.type}.${_.Edit}_${e}_${t}`;
  }
}
class pr {
  constructor(e, t) {
    R(this, "value", !0);
    R(this, "type", x.Worksheet);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.EditExtraObject);
    this.unitId = e, this.subUnitId = t, this.id = `${this.type}.${_.EditExtraObject}_${e}_${t}`;
  }
}
class Sr {
  constructor(e, t) {
    R(this, "value", !0);
    R(this, "type", x.Worksheet);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.Filter);
    this.unitId = e, this.subUnitId = t, this.id = `${this.type}.${_.Filter}_${e}_${t}`;
  }
}
class wr {
  constructor(e, t) {
    R(this, "value", !0);
    R(this, "type", x.Worksheet);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.InsertColumn);
    this.unitId = e, this.subUnitId = t, this.id = `${this.type}.${_.InsertColumn}_${e}_${t}`;
  }
}
class Ir {
  constructor(e, t) {
    R(this, "value", !0);
    R(this, "type", x.Worksheet);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.InsertHyperlink);
    this.unitId = e, this.subUnitId = t, this.id = `${this.type}.${_.InsertHyperlink}_${e}_${t}`;
  }
}
class vr {
  constructor(e, t) {
    R(this, "value", !0);
    R(this, "type", x.Worksheet);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.InsertRow);
    this.unitId = e, this.subUnitId = t, this.id = `${this.type}.${_.InsertRow}_${e}_${t}`;
  }
}
class Mr {
  constructor(e, t) {
    R(this, "value", !0);
    R(this, "type", x.Worksheet);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.ManageCollaborator);
    this.unitId = e, this.subUnitId = t, this.id = `${this.type}.${_.ManageCollaborator}_${e}_${t}`;
  }
}
class _r {
  constructor(e, t) {
    R(this, "value", !0);
    R(this, "type", x.Worksheet);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.PivotTable);
    this.unitId = e, this.subUnitId = t, this.id = `${this.type}.${_.PivotTable}_${e}_${t}`;
  }
}
class Ph {
  constructor(e, t) {
    R(this, "value", !0);
    R(this, "type", x.Worksheet);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.SelectProtectedCells);
    this.unitId = e, this.subUnitId = t, this.id = `${this.type}.${_.SelectProtectedCells}_${e}_${t}`;
  }
}
class Nh {
  constructor(e, t) {
    R(this, "value", !0);
    R(this, "type", x.Worksheet);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.SelectUnProtectedCells);
    this.unitId = e, this.subUnitId = t, this.id = `${this.type}.${_.SelectUnProtectedCells}_${e}_${t}`;
  }
}
class yr {
  constructor(e, t) {
    R(this, "value", !0);
    R(this, "type", x.Worksheet);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.SetCellStyle);
    this.unitId = e, this.subUnitId = t, this.id = `${this.type}.${_.SetCellStyle}_${e}_${t}`;
  }
}
class Dn {
  constructor(e, t) {
    R(this, "value", !0);
    R(this, "type", x.Worksheet);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.SetCellValue);
    this.unitId = e, this.subUnitId = t, this.id = `${this.type}.${_.SetCellValue}_${e}_${t}`;
  }
}
class Zt {
  constructor(e, t) {
    R(this, "value", !0);
    R(this, "type", x.Worksheet);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.SetColumnStyle);
    this.unitId = e, this.subUnitId = t, this.id = `${this.type}.${_.SetColumnStyle}_${e}_${t}`;
  }
}
class Qt {
  constructor(e, t) {
    R(this, "value", !0);
    R(this, "type", x.Worksheet);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.SetRowStyle);
    this.unitId = e, this.subUnitId = t, this.id = `${this.type}.${_.SetRowStyle}_${e}_${t}`;
  }
}
class br {
  constructor(e, t) {
    R(this, "value", !0);
    R(this, "type", x.Worksheet);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.Sort);
    this.unitId = e, this.subUnitId = t, this.id = `${this.type}.${_.Sort}_${e}_${t}`;
  }
}
class lo {
  constructor(e, t) {
    R(this, "value", !0);
    R(this, "type", x.Worksheet);
    R(this, "status", q.INIT);
    R(this, "id");
    R(this, "subType", _.View);
    this.unitId = e, this.subUnitId = t, this.id = `${this.type}.${_.View}_${e}_${t}`;
  }
}
const pn = {
  id: "sheet.command.set-range-values",
  type: v.COMMAND,
  handler: (n, e) => {
    var H;
    const t = n.get(E), o = n.get(V), s = n.get(M), r = n.get(z), i = n.get(G), a = n.get(Ze), u = P(s, e);
    if (!u) return !1;
    const { subUnitId: l, unitId: c, workbook: d, worksheet: h } = u, { value: g, range: m, redoUndoId: f } = e, C = m ? [m] : (H = r.getCurrentSelections()) == null ? void 0 : H.map((F) => F.range);
    if (!C || !C.length || !a.getPermissionPoint(new Le(c, l).id)) return !1;
    const p = new Y();
    let w;
    if (A.isArray(g))
      for (let F = 0; F < C.length; F++) {
        const { startRow: K, startColumn: J, endRow: ae, endColumn: se } = C[F];
        for (let le = 0; le <= ae - K; le++)
          for (let Se = 0; Se <= se - J; Se++)
            p.setValue(le + K, Se + J, g[le][Se]);
      }
    else if (Ls(g))
      for (let F = 0; F < C.length; F++) {
        const { startRow: K, startColumn: J, endRow: ae, endColumn: se } = C[F];
        for (let le = K; le <= ae; le++)
          for (let Se = J; Se <= se; Se++)
            p.setValue(le, Se, g);
      }
    else
      w = g;
    const I = { subUnitId: l, unitId: c, cellValue: w != null ? w : p.getMatrix() }, S = Ae(n, I), y = bi(I.cellValue, (F, K) => h.getCellHeight(F, K) || void 0);
    if (!t.syncExecuteCommand(ee.id, I)) return !1;
    const { undos: O, redos: k } = i.onCommandExecute({
      id: pn.id,
      params: I
    }), { undos: T, redos: U } = i.generateMutationsOfAutoHeight({
      unitId: c,
      subUnitId: l,
      ranges: C,
      cellHeights: new Y(y)
    });
    if (L([...k, ...U], t).result) {
      const F = ut(m != null ? m : p.getRange(), d, h);
      return o.pushUndoRedo({
        unitID: c,
        undoMutations: [
          { id: ee.id, params: S },
          ...O,
          ...T,
          F
        ],
        redoMutations: [
          { id: ee.id, params: I },
          ...k,
          ...U,
          A.deepClone(F)
        ],
        id: f
      }), !0;
    }
    return !1;
  }
};
function Er(n, e) {
  const t = [], o = [], { unitId: s, subUnitId: r, range: i, shiftDimension: a, cellValue: u = {} } = e, l = n.get(M), c = n.get(G), d = l.getUniverSheetInstance(s), h = d == null ? void 0 : d.getSheetBySheetId(r);
  if (h) {
    const g = h.getCellMatrix(), m = g.getDataRange();
    if (i.startColumn <= m.endColumn || i.startRow <= m.endRow) {
      let I, S;
      if (a === Re.COLUMNS) {
        const b = Math.min(i.endRow, m.endRow);
        let O = 0;
        for (let T = i.startRow; T <= b; T++) {
          const U = g.getRow(T), D = U ? Hs(U) - 1 : 0;
          O = Math.max(O, D);
        }
        I = {
          startRow: i.startRow,
          startColumn: i.startColumn,
          endRow: b,
          endColumn: O
        };
        const k = i.endColumn - i.startColumn + 1;
        S = {
          startRow: i.startRow,
          startColumn: I.startColumn + k,
          endRow: b,
          endColumn: I.endColumn + k
        };
      } else {
        const b = Math.min(i.endColumn, m.endColumn), O = m.endRow;
        I = {
          startRow: i.startRow,
          startColumn: i.startColumn,
          endRow: O,
          endColumn: b
        };
        const k = i.endRow - i.startRow + 1;
        S = {
          startRow: I.startRow + k,
          startColumn: i.startColumn,
          endRow: I.endRow + k,
          endColumn: b
        };
      }
      const y = Ho(
        n,
        { unitId: s, subUnitId: r, range: I },
        { subUnitId: r, range: S },
        !0
      );
      y && (t.push(...y.redos), o.push(...y.undos));
    }
    if (Object.entries(u).length === 0)
      for (let I = i.startRow; I <= i.endRow; I++) {
        u[I] || (u[I] = {});
        for (let S = i.startColumn; S <= i.endColumn; S++)
          u[I][S] = null;
      }
    const f = {
      subUnitId: r,
      unitId: s,
      cellValue: u
    }, C = Ae(
      n,
      f
    ), { undos: p, redos: w } = c.onCommandExecute({
      id: pn.id,
      params: { ...f, range: i }
    });
    t.push({ id: ee.id, params: f }, ...w), o.push({ id: ee.id, params: C }, ...p);
  }
  return {
    redo: t,
    undo: o
  };
}
function Ur(n, e) {
  const t = [], o = [], { unitId: s, subUnitId: r, range: i, shiftDimension: a } = e, u = n.get(M), l = n.get(G), c = u.getUniverSheetInstance(s), d = c == null ? void 0 : c.getSheetBySheetId(r);
  if (d) {
    const h = d.getCellMatrix(), g = h.getDataRange(), m = {
      subUnitId: r,
      unitId: s,
      cellValue: er([i])
    }, f = Ae(
      n,
      m
    ), C = l.onCommandExecute({
      id: pn.id,
      params: m
    });
    if (t.push({ id: ee.id, params: m }, ...C.redos), o.push(...C.undos, {
      id: ee.id,
      params: f
    }), i.startColumn <= g.endColumn || i.startRow <= g.endRow) {
      let p = null, w = null;
      if (a === Re.COLUMNS && i.endColumn < g.endColumn) {
        const I = Math.min(i.endRow, g.endRow);
        let S = 0;
        for (let b = i.startRow; b <= I; b++) {
          const O = h.getRow(b), k = O ? Hs(O) - 1 : 0;
          S = Math.max(S, k);
        }
        p = {
          startRow: i.startRow,
          startColumn: i.endColumn + 1,
          endRow: I,
          endColumn: S
        };
        const y = i.endColumn - i.startColumn + 1;
        w = {
          startRow: i.startRow,
          startColumn: p.startColumn - y,
          endRow: I,
          endColumn: p.endColumn - y
        };
      }
      if (a === Re.ROWS && i.endRow < g.endRow) {
        const I = Math.min(i.endColumn, g.endColumn), S = g.endRow;
        p = {
          startRow: i.endRow + 1,
          startColumn: i.startColumn,
          endRow: S,
          endColumn: I
        };
        const y = i.endRow - i.startRow + 1;
        w = {
          startRow: p.startRow - y,
          startColumn: i.startColumn,
          endRow: p.endRow - y,
          endColumn: I
        };
      }
      if (p && w) {
        const I = Ho(
          n,
          { unitId: s, subUnitId: r, range: p },
          { subUnitId: r, range: w },
          !0
        );
        I && (t.push(...I.redos), o.push(...I.undos));
      }
    }
  }
  return {
    redo: t,
    undo: o
  };
}
function Oh(n, e, t, o, s, r) {
  const { startRow: i, endRow: a, startColumn: u, endColumn: l } = e;
  if (s === Re.ROWS) {
    const c = a - i + 1;
    for (let d = t; d >= i; d--)
      for (let h = u; h <= l; h++) {
        const g = n.getValue(d, h);
        g == null ? n.realDeleteValue(d + c, h) : n.setValue(d + c, h, g);
      }
    for (let d = a; d >= i; d--)
      for (let h = u; h <= l; h++)
        r && r[d] && r[d][h] ? n.setValue(d, h, r[d][h]) : n.realDeleteValue(d, h);
  } else if (s === Re.COLUMNS) {
    const c = l - u + 1;
    for (let d = i; d <= a; d++)
      for (let h = o; h >= u; h--) {
        const g = n.getValue(d, h);
        g == null ? n.realDeleteValue(d, h + c) : n.setValue(d, h + c, g);
      }
    for (let d = i; d <= a; d++)
      for (let h = l; h >= u; h--)
        r && r[d] && r[d][h] ? n.setValue(d, h, r[d][h]) : n.realDeleteValue(d, h);
  }
}
function Dh(n, e, t, o, s) {
  const { startRow: r, endRow: i, startColumn: a, endColumn: u } = e, l = i - r + 1, c = u - a + 1;
  if (s === Re.ROWS)
    for (let d = r; d <= t; d++)
      for (let h = a; h <= u; h++) {
        const g = n.getValue(d + l, h);
        g == null ? n.realDeleteValue(d, h) : n.setValue(d, h, g);
      }
  else if (s === Re.COLUMNS)
    for (let d = r; d <= i; d++)
      for (let h = a; h <= o; h++) {
        const g = n.getValue(d, h + c);
        g == null ? n.realDeleteValue(d, h) : n.setValue(d, h, g);
      }
}
const kr = "sheet.command.delete-range-move-left", It = {
  type: v.COMMAND,
  id: kr,
  handler: async (n, e) => {
    var S, y, b;
    const t = n.get(E), o = n.get(V), s = n.get(M), r = n.get(z), i = n.get(G), a = P(s);
    if (!a) return !1;
    const { worksheet: u, workbook: l, subUnitId: c, unitId: d } = a;
    let h = e == null ? void 0 : e.range;
    if (h || (h = (S = r.getCurrentLastSelection()) == null ? void 0 : S.range), !h) return !1;
    const g = {
      range: h,
      subUnitId: c,
      unitId: d,
      shiftDimension: Re.COLUMNS
    }, m = i.onCommandExecute({
      id: It.id,
      params: { range: h }
    }), { redo: f, undo: C } = Ur(
      n,
      g
    ), p = [...(y = m.preRedos) != null ? y : [], ...f], w = [...m.undos, ...C];
    if (p.push(...m.redos), p.push(ut(h, l, u)), w.push(...(b = m.preUndos) != null ? b : []), L(p, t).result) {
      const O = i.afterCommandExecute({
        id: It.id,
        params: { range: h }
      });
      return L(O.redos, t), w.push(...O.undos), p.push(...O.redos), o.pushUndoRedo({
        unitID: d,
        undoMutations: w.reverse(),
        redoMutations: p
      }), !0;
    }
    return !1;
  }
  // all subsequent mutations should succeed inorder to make the whole process succeed
  // Promise.all([]).then(() => true),
}, Tr = "sheet.command.delete-range-move-up", vt = {
  type: v.COMMAND,
  id: Tr,
  handler: async (n, e) => {
    var S, y, b;
    const t = n.get(E), o = n.get(V), s = n.get(M), r = n.get(z), i = n.get(G), a = P(s);
    if (!a) return !1;
    const { unitId: u, subUnitId: l, workbook: c, worksheet: d } = a;
    let h = e == null ? void 0 : e.range;
    if (h || (h = (S = r.getCurrentLastSelection()) == null ? void 0 : S.range), !h) return !1;
    const g = {
      range: h,
      subUnitId: l,
      unitId: u,
      shiftDimension: Re.ROWS
    }, m = i.onCommandExecute({
      id: vt.id,
      params: { range: h }
    }), { redo: f, undo: C } = Ur(
      n,
      g
    ), p = [...(y = m.preRedos) != null ? y : [], ...f], w = [...m.undos, ...C];
    if (p.push(...m.redos), p.push(ut(h, c, d)), w.push(...(b = m.preUndos) != null ? b : []), L(p, t).result) {
      const O = i.afterCommandExecute({
        id: vt.id,
        params: { range: h }
      });
      return L(O.redos, t), w.push(...O.undos), p.push(...O.redos), o.pushUndoRedo({
        unitID: u,
        undoMutations: w.reverse(),
        redoMutations: p
      }), !0;
    }
    return !1;
  }
  // all subsequent mutations should succeed inorder to make the whole process succeed
  // Promise.all([]).then(() => true),
}, Zu = "sheet.command.insert-range-move-down", Pt = {
  type: v.COMMAND,
  id: "sheet.command.insert-range-move-down",
  // eslint-disable-next-line max-lines-per-function
  handler: async (n, e) => {
    var H, F, K;
    const t = n.get(E), o = n.get(V), s = n.get(M), r = n.get(z), i = n.get(G), a = n.get(Vt), u = n.get(it);
    if (r.isOverlapping())
      return a.emit(u.t("sheets.info.overlappingSelections")), !1;
    const l = P(s);
    if (!l) return !1;
    const { unitId: c, subUnitId: d, worksheet: h, workbook: g } = l;
    let m = e == null ? void 0 : e.range;
    if (m || (m = (H = r.getCurrentLastSelection()) == null ? void 0 : H.range), !m) return !1;
    const f = [], C = [], p = h.getCellMatrix(), w = p.getDataRange(), S = p.getSlice(w.startRow, w.endRow, m.startColumn, m.endColumn).getDataRange().endRow, y = Math.max(S + (m.endRow - m.startRow + 1) - w.endRow, 0);
    if (y > 0) {
      const J = m.startRow - 1, ae = h.getRowHeight(J), se = {
        unitId: c,
        subUnitId: d,
        range: {
          startRow: w.endRow + 1,
          endRow: w.endRow + y,
          startColumn: w.startColumn,
          endColumn: w.endColumn
        },
        rowInfo: new Array(y).fill(void 0).map(() => ({
          h: ae,
          hd: te.FALSE
        }))
      };
      f.push({
        id: Ue.id,
        params: se
      });
      const le = $o(
        n,
        se
      );
      C.push({ id: Te.id, params: le });
    }
    const b = {};
    X.foreach(m, (J, ae) => {
      const se = h.getCell(J, ae);
      se && (b[J] || (b[J] = {}), b[J][ae] = { s: se.s });
    });
    const O = {
      range: m,
      subUnitId: d,
      unitId: c,
      shiftDimension: Re.ROWS,
      cellValue: b
    }, { redo: k, undo: T } = Er(
      n,
      O
    );
    f.push(...k), C.push(...T);
    const U = i.onCommandExecute({
      id: Pt.id,
      params: { range: m }
    });
    if (f.push(...U.redos), f.push(ut(m, g, h)), C.push(...(F = U.preUndos) != null ? F : []), f.unshift(...(K = U.preRedos) != null ? K : []), C.unshift(...U.undos), L(f, t)) {
      const J = i.afterCommandExecute({
        id: Pt.id,
        params: { range: m }
      });
      return L(J.redos, t), C.push(...J.undos), f.push(...J.redos), o.pushUndoRedo({
        unitID: c,
        undoMutations: C.reverse(),
        redoMutations: f
      }), !0;
    }
    return !1;
  }
  // all subsequent mutations should succeed inorder to make the whole process succeed
  // Promise.all([]).then(() => true),
}, Ko = "sheet.command.insert-range-move-right", rn = {
  type: v.COMMAND,
  id: Ko,
  // eslint-disable-next-line max-lines-per-function
  handler: async (n, e) => {
    var H, F, K;
    const t = n.get(E), o = n.get(V), s = n.get(M), r = n.get(z), i = n.get(G), a = n.get(Vt), u = n.get(it);
    if (r.isOverlapping())
      return a.emit(u.t("sheets.info.overlappingSelections")), !1;
    const l = P(s);
    if (!l) return !1;
    const { workbook: c, worksheet: d, unitId: h, subUnitId: g } = l;
    let m = e == null ? void 0 : e.range;
    if (m || (m = (H = r.getCurrentLastSelection()) == null ? void 0 : H.range), !m) return !1;
    const f = [], C = [], p = d.getCellMatrix(), w = p.getDataRange(), S = p.getSlice(m.startRow, m.endRow, w.startColumn, w.endColumn).getDataRange().endColumn, y = Math.max(
      S + (m.endColumn - m.startColumn + 1) - w.endColumn,
      0
    );
    if (y > 0) {
      const J = m.startColumn - 1, ae = d.getColumnWidth(J), se = {
        unitId: h,
        subUnitId: g,
        range: {
          startRow: w.startRow + 1,
          endRow: w.endRow,
          startColumn: w.endColumn + 1,
          endColumn: w.endColumn + y
        },
        colInfo: new Array(y).fill(void 0).map(() => ({
          w: ae,
          hd: te.FALSE
        }))
      };
      f.push({
        id: ke.id,
        params: se
      });
      const le = ro(
        n,
        se
      );
      C.push({ id: Me.id, params: le });
    }
    const b = {};
    X.foreach(m, (J, ae) => {
      const se = d.getCell(J, ae);
      !se || !se.s || (b[J] || (b[J] = {}), b[J][ae] = { s: se.s });
    });
    const O = {
      range: m,
      subUnitId: g,
      unitId: h,
      shiftDimension: Re.COLUMNS,
      cellValue: b
    }, { redo: k, undo: T } = Er(
      n,
      O
    );
    f.push(...k), C.push(...T);
    const U = i.onCommandExecute({
      id: rn.id,
      params: { range: m }
    });
    if (f.push(...U.redos), f.push(ut(m, c, d)), C.push(...(F = U.preUndos) != null ? F : []), f.unshift(...(K = U.preRedos) != null ? K : []), C.unshift(...U.undos), L(f, t).result) {
      const J = i.afterCommandExecute({
        id: rn.id,
        params: { range: m }
      });
      return L(J.redos, t), C.push(...J.undos), f.push(...J.redos), o.pushUndoRedo({
        unitID: h,
        undoMutations: C.reverse(),
        redoMutations: f
      }), !0;
    }
    return !1;
  }
  // all subsequent mutations should succeed in order to make the whole process succeed
  // Promise.all([]).then(() => true),
}, Pr = "sheet.command.insert-row", Ye = {
  type: v.COMMAND,
  id: Pr,
  handler: async (n, e) => {
    const t = n.get(E), o = n.get(G), { range: s, direction: r, unitId: i, subUnitId: a, cellValue: u } = e;
    return await o.beforeCommandExecute({
      id: Ye.id,
      params: e
    }) ? t.syncExecuteCommand(Nr.id, {
      range: s,
      direction: r,
      unitId: i,
      subUnitId: a,
      cellValue: u
    }) : !1;
  }
}, Nr = {
  type: v.COMMAND,
  id: "sheet.command.insert-row-by-range",
  handler: (n, e) => {
    var k, T, U, D;
    const t = n.get(E), o = n.get(V), s = n.get(M), r = n.get(G), i = P(s, e);
    if (!i) return !1;
    const { workbook: a, worksheet: u } = i, { range: l, direction: c, unitId: d, subUnitId: h, cellValue: g } = e, { startRow: m, endRow: f } = l;
    l.rangeType = j.ROW;
    const C = c === pe.UP ? m : m - 1, p = u.getRowHeight(C), w = {
      unitId: d,
      subUnitId: h,
      range: l,
      rowInfo: new Array(f - m + 1).fill(void 0).map(() => ({
        h: p,
        hd: te.FALSE
      }))
      // row height should inherit from the anchor row
    }, I = $o(
      n,
      w
    ), S = [{ id: Ue.id, params: w }], y = [{ id: Te.id, params: I }];
    g && Object.keys(g).length > 0 && S.push({
      id: ee.id,
      params: {
        unitId: d,
        subUnitId: h,
        cellValue: g
      }
    });
    const b = r.onCommandExecute({
      id: Ye.id,
      params: e
    });
    if (S.unshift(...(k = b.preRedos) != null ? k : []), S.push(...(T = b.redos) != null ? T : []), S.push(ut(l, a, u)), y.unshift(...(U = b.preUndos) != null ? U : []), y.push(...(D = b.undos) != null ? D : []), L(S, t).result) {
      const H = r.afterCommandExecute({
        id: Ye.id,
        params: e
      });
      return L(H.redos, t), S.push(...H.redos), y.push(...H.undos), o.pushUndoRedo({
        unitID: e.unitId,
        undoMutations: y,
        redoMutations: S
      }), !0;
    }
    return !1;
  }
}, Qu = {
  type: v.COMMAND,
  id: "sheet.command.insert-row-before",
  handler: async (n, e) => {
    var C;
    const o = (C = n.get(z).getCurrentSelections()) == null ? void 0 : C.map((p) => p.range);
    let s;
    if ((o == null ? void 0 : o.length) === 1)
      s = o[0];
    else
      return !1;
    const r = n.get(M), i = P(r);
    if (!i) return !1;
    const { worksheet: a, subUnitId: u, unitId: l } = i, c = e.value || 0, d = s.startRow, h = s.startRow + c - 1, g = 0, m = a.getColumnCount() - 1, f = {
      unitId: l,
      subUnitId: u,
      direction: pe.UP,
      range: {
        startRow: d,
        endRow: h,
        startColumn: g,
        endColumn: m
      },
      // copy styles from the row above
      cellValue: Ct(a, d, h, g, m, !0, d - 1)
    };
    return n.get(E).executeCommand(Ye.id, f);
  }
}, el = {
  type: v.COMMAND,
  id: "sheet.command.insert-row-after",
  handler: async (n) => {
    var f;
    const t = (f = n.get(z).getCurrentSelections()) == null ? void 0 : f.map((C) => C.range);
    let o;
    if ((t == null ? void 0 : t.length) === 1)
      o = t[0];
    else
      return !1;
    const s = n.get(M), r = P(s);
    if (!r) return !1;
    const { worksheet: i, unitId: a, subUnitId: u } = r, l = o.endRow - o.startRow + 1, c = o.endRow + 1, d = o.endRow + l, h = 0, g = i.getColumnCount() - 1, m = {
      unitId: a,
      subUnitId: u,
      direction: pe.DOWN,
      range: {
        startRow: c,
        endRow: d,
        startColumn: h,
        endColumn: g,
        rangeType: j.ROW
      },
      // copy styles from the row below
      cellValue: Ct(i, c, d, h, g, !0, o.endRow)
    };
    return n.get(E).executeCommand(Ye.id, m);
  }
}, tl = {
  type: v.COMMAND,
  id: "sheet.command.insert-multi-rows-above",
  handler: async (n, e) => {
    var p;
    const o = (p = n.get(z).getCurrentSelections()) == null ? void 0 : p.map((w) => w.range);
    let s;
    if ((o == null ? void 0 : o.length) === 1)
      s = o[0];
    else
      return !1;
    const r = n.get(M), i = P(r);
    if (!i) return !1;
    const { worksheet: a, unitId: u, subUnitId: l } = i, c = e.value || 0, d = s.startRow, h = s.startRow + c - 1, g = 0, m = a.getColumnCount() - 1, f = Ct(a, d, h, g, m, !0, d - 1), C = {
      unitId: u,
      subUnitId: l,
      direction: pe.UP,
      range: {
        startRow: d,
        endRow: h,
        startColumn: g,
        endColumn: m,
        rangeType: j.ROW
      },
      // copy styles from the row above
      cellValue: f
    };
    return n.get(E).executeCommand(Ye.id, C);
  }
}, nl = {
  type: v.COMMAND,
  id: "sheet.command.insert-multi-rows-after",
  handler: async (n, e) => {
    var C;
    const o = (C = n.get(z).getCurrentSelections()) == null ? void 0 : C.map((p) => p.range);
    let s;
    if ((o == null ? void 0 : o.length) === 1)
      s = o[0];
    else
      return !1;
    const r = n.get(M), i = P(r);
    if (!i) return !1;
    const { worksheet: a, unitId: u, subUnitId: l } = i, c = e.value || 0, d = s.endRow + 1, h = s.endRow + c, g = 0, m = a.getColumnCount() - 1, f = {
      unitId: u,
      subUnitId: l,
      direction: pe.DOWN,
      range: {
        startRow: d,
        endRow: h,
        startColumn: g,
        endColumn: m,
        rangeType: j.ROW
      },
      // copy styles from the row below
      cellValue: Ct(a, d, h, g, m, !0, s.endRow)
    };
    return n.get(E).executeCommand(Ye.id, f);
  }
}, Or = "sheet.command.insert-col", Xe = {
  type: v.COMMAND,
  id: Or,
  handler: async (n, e) => {
    const t = n.get(E), o = n.get(G), { range: s, direction: r, subUnitId: i, unitId: a, cellValue: u } = e;
    return await o.beforeCommandExecute({
      id: Xe.id,
      params: e
    }) ? t.syncExecuteCommand(Dr.id, {
      range: s,
      direction: r,
      unitId: a,
      subUnitId: i,
      cellValue: u
    }) : !1;
  }
}, Dr = {
  type: v.COMMAND,
  id: "sheet.command.insert-col-by-range",
  handler: (n, e) => {
    var O, k, T, U;
    const t = n.get(E), o = n.get(V), s = n.get(M), r = n.get(G), { range: i, direction: a, subUnitId: u, unitId: l, cellValue: c } = e, { startColumn: d, endColumn: h } = e.range;
    i.rangeType = j.COLUMN;
    const g = s.getUniverSheetInstance(e.unitId), m = g.getSheetBySheetId(e.subUnitId), f = a === pe.LEFT ? d : d - 1, C = m.getColumnWidth(f), p = {
      unitId: l,
      subUnitId: u,
      range: i,
      colInfo: new Array(h - d + 1).fill(void 0).map(() => ({
        w: C,
        hd: te.FALSE
      }))
    }, w = ro(
      n,
      p
    ), I = [{ id: ke.id, params: p }], S = [{ id: Me.id, params: w }];
    c && I.push({
      id: ee.id,
      params: {
        unitId: l,
        subUnitId: u,
        cellValue: c
      }
    });
    const y = r.onCommandExecute({
      id: Xe.id,
      params: e
    });
    if (I.unshift(...(O = y.preRedos) != null ? O : []), I.push(...(k = y.redos) != null ? k : []), I.push(ut(i, g, m)), S.unshift(...(T = y.preUndos) != null ? T : []), S.push(...(U = y.undos) != null ? U : []), L(I, t).result) {
      const D = r.afterCommandExecute({
        id: Xe.id,
        params: e
      });
      return L(D.redos, t), I.push(...D.redos), S.push(...D.undos), o.pushUndoRedo({
        unitID: e.unitId,
        undoMutations: S.filter(Boolean),
        redoMutations: I.filter(Boolean)
      }), !0;
    }
    return !1;
  }
}, ol = {
  type: v.COMMAND,
  id: "sheet.command.insert-col-before",
  handler: async (n, e) => {
    const o = n.get(z).getCurrentSelections();
    let s;
    if ((o == null ? void 0 : o.length) === 1)
      s = o[0].range;
    else
      return !1;
    const r = n.get(M), i = P(r);
    if (!i) return !1;
    const { worksheet: a, unitId: u, subUnitId: l } = i, c = e.value || 0, d = s.startColumn, h = s.startColumn + c - 1, g = 0, m = a.getRowCount() - 1, f = {
      unitId: u,
      subUnitId: l,
      direction: pe.LEFT,
      range: {
        startColumn: d,
        endColumn: h,
        startRow: g,
        endRow: m,
        rangeType: j.COLUMN
      },
      // copy styles from the column before
      cellValue: Ct(a, g, m, d, h, !1, d - 1)
    };
    return n.get(E).executeCommand(Xe.id, f);
  }
}, sl = {
  type: v.COMMAND,
  id: "sheet.command.insert-col-after",
  handler: async (n) => {
    const t = n.get(z).getCurrentSelections();
    let o;
    if ((t == null ? void 0 : t.length) === 1)
      o = t[0].range;
    else
      return !1;
    const s = n.get(M), r = P(s);
    if (!r) return !1;
    const { worksheet: i, unitId: a, subUnitId: u } = r, l = o.endColumn - o.startColumn + 1, c = o.endColumn + 1, d = o.endColumn + l, h = 0, g = i.getRowCount() - 1, m = {
      unitId: a,
      subUnitId: u,
      direction: pe.RIGHT,
      range: {
        startColumn: c,
        endColumn: d,
        startRow: h,
        endRow: g
      },
      // copy styles from the column after
      cellValue: Ct(i, h, g, c, d, !1, o.endColumn)
    };
    return n.get(E).executeCommand(Xe.id, m);
  }
}, rl = {
  type: v.COMMAND,
  id: "sheet.command.insert-multi-cols-before",
  handler: async (n, e) => {
    const o = n.get(z).getCurrentSelections();
    let s;
    if ((o == null ? void 0 : o.length) === 1)
      s = o[0].range;
    else
      return !1;
    const r = n.get(M), i = P(r);
    if (!i) return !1;
    const { worksheet: a, unitId: u, subUnitId: l } = i, c = e.value || 0, d = s.startColumn, h = s.startColumn + c - 1, g = 0, m = a.getRowCount() - 1, f = {
      unitId: u,
      subUnitId: l,
      direction: pe.LEFT,
      range: {
        startColumn: d,
        endColumn: h,
        startRow: g,
        endRow: m,
        rangeType: j.COLUMN
      },
      // copy styles from the column before
      cellValue: Ct(a, g, m, d, h, !1, d - 1)
    };
    return n.get(E).executeCommand(Xe.id, f);
  }
}, il = {
  type: v.COMMAND,
  id: "sheet.command.insert-multi-cols-right",
  handler: async (n, e) => {
    const o = n.get(z).getCurrentSelections();
    let s;
    if ((o == null ? void 0 : o.length) === 1)
      s = o[0].range;
    else
      return !1;
    const r = n.get(M), i = P(r);
    if (!i) return !1;
    const { worksheet: a, unitId: u, subUnitId: l } = i, c = e.value || 0, d = s.endColumn + 1, h = s.endColumn + c, g = 0, m = a.getRowCount() - 1, f = {
      unitId: u,
      subUnitId: l,
      direction: pe.RIGHT,
      range: {
        startColumn: d,
        endColumn: h,
        startRow: g,
        endRow: m
      },
      // copy styles from the column after
      cellValue: Ct(a, g, m, d, h, !1, s.endColumn)
    };
    return n.get(E).executeCommand(Xe.id, f);
  }
}, An = "sheet.command.remove-row", Ar = {
  type: v.COMMAND,
  id: "sheet.command.remove-row-by-range",
  handler: (n, e) => {
    var C, p, w;
    if (!e)
      return !1;
    const t = n.get(M), o = P(t, e);
    if (!o) return !1;
    const { workbook: s, worksheet: r } = o, i = n.get(G), { range: a, unitId: u, subUnitId: l } = e, c = ao([a], n, u, l).reverse(), d = [], h = [];
    c.forEach((I) => {
      const S = [], y = [], b = {
        unitId: u,
        subUnitId: l,
        range: I
      }, O = $a(
        b,
        r
      ), k = r.getCellMatrix().getSlice(I.startRow, I.endRow, 0, r.getColumnCount() - 1), T = {
        unitId: u,
        subUnitId: l,
        cellValue: k.getMatrix()
      };
      y.push({ id: Te.id, params: b }), S.push({ id: Ue.id, params: O }), S.push({ id: ee.id, params: T }), h.push(...y), d.unshift(...S);
    });
    const g = i.onCommandExecute({
      id: An,
      params: { range: a }
    }), m = n.get(E);
    if (L(
      [
        ...(C = g.preRedos) != null ? C : [],
        ...h,
        ...g.redos,
        ut(a, s, r)
      ],
      m
    ).result) {
      const I = i.afterCommandExecute({
        id: An,
        params: { range: a }
      });
      return L(I.redos, m), n.get(V).pushUndoRedo({
        unitID: u,
        undoMutations: [
          ...(p = g.preUndos) != null ? p : [],
          ...d,
          ...g.undos,
          ...I.undos
        ],
        redoMutations: [
          ...(w = g.preRedos) != null ? w : [],
          ...h,
          ...g.redos,
          ...I.redos
        ]
      }), !0;
    }
    return !1;
  }
}, co = {
  type: v.COMMAND,
  id: An,
  handler: async (n, e) => {
    var h;
    const t = n.get(z), o = n.get(G), s = n.get(E);
    let r = e == null ? void 0 : e.range;
    if (r || (r = (h = t.getCurrentLastSelection()) == null ? void 0 : h.range), !r) return !1;
    const i = n.get(M), a = P(i);
    if (!a) return !1;
    const { worksheet: u, subUnitId: l, unitId: c } = a;
    return r = {
      ...r,
      startColumn: 0,
      endColumn: Math.max(u.getMaxColumns() - 1, 0)
    }, await o.beforeCommandExecute({
      id: co.id,
      params: { range: r }
    }) ? s.syncExecuteCommand(Ar.id, {
      range: r,
      unitId: c,
      subUnitId: l
    }) : !1;
  }
}, xn = "sheet.command.remove-col", xr = {
  type: v.COMMAND,
  id: "sheet.command.remove-col-by-range",
  handler: (n, e) => {
    var p, w, I;
    if (!e)
      return !1;
    const t = n.get(M), o = P(t, e);
    if (!o) return !1;
    const { workbook: s, worksheet: r } = o, i = n.get(G), { range: a, unitId: u, subUnitId: l } = e, c = {
      unitId: u,
      subUnitId: l,
      range: a
    }, d = Wa(n, c), h = r.getCellMatrix().getSlice(0, r.getRowCount() - 1, a.startColumn, a.endColumn), g = {
      unitId: u,
      subUnitId: l,
      cellValue: h.getMatrix()
    }, m = i.onCommandExecute({
      id: xn,
      params: { range: a }
    }), f = n.get(E);
    if (L(
      [
        ...(p = m.preRedos) != null ? p : [],
        { id: Me.id, params: c },
        ...m.redos,
        ut(a, s, r)
      ],
      f
    ).result) {
      const S = i.afterCommandExecute({
        id: xn,
        params: { range: a }
      });
      return L(S.redos, f), n.get(V).pushUndoRedo({
        unitID: u,
        undoMutations: [
          ...(w = m.preUndos) != null ? w : [],
          { id: ke.id, params: d },
          { id: ee.id, params: g },
          ...m.undos,
          ...S.undos
        ],
        redoMutations: [
          ...(I = m.preRedos) != null ? I : [],
          { id: Me.id, params: c },
          ...m.redos,
          ...S.redos
        ]
      }), !0;
    }
    return !1;
  }
}, ho = {
  type: v.COMMAND,
  id: xn,
  handler: async (n, e) => {
    var h;
    const t = n.get(z), o = n.get(G), s = n.get(E);
    let r = e == null ? void 0 : e.range;
    if (r || (r = (h = t.getCurrentLastSelection()) == null ? void 0 : h.range), !r) return !1;
    const i = n.get(M), a = P(i);
    if (!a) return !1;
    const { worksheet: u, subUnitId: l, unitId: c } = a;
    return r = {
      ...r,
      startRow: 0,
      endRow: Math.max(u.getMaxRows() - 1, 0)
    }, await o.beforeCommandExecute({
      id: ho.id,
      params: { range: r }
    }) ? s.syncExecuteCommand(xr.id, {
      range: r,
      unitId: c,
      subUnitId: l
    }) : !1;
  }
}, al = (n, e) => {
  const t = n.get(M), { subUnitId: o, unitId: s } = e, r = Qe(t, e);
  if (!r)
    throw new Error("[RemoveSheetUndoMutationFactory]: Worksheet is null error!");
  const { workbook: i, worksheet: a } = r, u = a.getConfig();
  return {
    index: i.getConfig().sheetOrder.findIndex((d) => d === o),
    sheet: u,
    unitId: s
  };
}, Lt = {
  id: "sheet.mutation.remove-sheet",
  type: v.MUTATION,
  handler: (n, e) => {
    const t = n.get(M), { subUnitId: o, unitId: s } = e, r = t.getUniverSheetInstance(s);
    return r ? r.removeSheet(o) : !1;
  }
};
function ul(n, e) {
  return e.getMergeData().some((t) => t.startRow < n && n <= t.endRow);
}
function ll(n, e) {
  return e.getMergeData().some((t) => t.startColumn < n && n <= t.endColumn);
}
const $r = "sheet.command.move-rows", an = {
  id: $r,
  type: v.COMMAND,
  // eslint-disable-next-line max-lines-per-function
  handler: (n, e) => {
    var H, F;
    const t = n.get(z), {
      fromRange: { startRow: o },
      toRange: { startRow: s },
      range: r
    } = e, i = r ? [Vr(r)] : t.getCurrentSelections(), a = i == null ? void 0 : i.filter(
      (K) => K.range.rangeType === j.ROW && K.range.startRow <= o && o <= K.range.endRow
    );
    if ((a == null ? void 0 : a.length) !== 1)
      return !1;
    const u = n.get(G), l = n.get(M), c = P(l, e);
    if (!c) return !1;
    const { workbook: d, worksheet: h } = c, g = d.getUnitId(), m = h.getSheetId(), f = n.get(Vt), C = n.get(it), p = a[0].range, w = a[0].primary, I = uo(p, h, !1);
    if (!N.equals(p, I))
      return f.emit(C.t("sheets.info.partOfCell")), !1;
    if (ul(s, h))
      return f.emit(C.t("sheets.info.acrossMergedCell")), !1;
    const S = {
      ...p,
      startRow: s,
      endRow: s + p.endRow - p.startRow
    }, y = {
      unitId: g,
      subUnitId: m,
      sourceRange: p,
      targetRange: S
    }, b = Aa(n, y), O = n.get(E), k = u.onCommandExecute({ id: an.id, params: e }), T = [
      ...(H = k.preRedos) != null ? H : [],
      { id: Je.id, params: y }
    ], U = [
      ...(F = k.preUndos) != null ? F : [],
      { id: Je.id, params: b }
    ];
    if (w) {
      const J = s - o < 0, ae = p.endRow - p.startRow + 1, se = J ? S : {
        ...S,
        startRow: S.startRow - ae,
        endRow: S.endRow - ae
      }, le = {
        unitId: g,
        subUnitId: m,
        type: we.MOVE_END,
        selections: [{
          range: se,
          primary: _e(se, h),
          style: null
        }]
      }, Se = {
        unitId: g,
        subUnitId: m,
        type: we.MOVE_END,
        selections: [{ range: p, primary: w, style: null }]
      };
      T.push({ id: ie.id, params: le }), U.push({ id: ie.id, params: Se });
    }
    if (T.push(...k.redos), U.push(...k.undos), L(T, O).result) {
      const K = u.afterCommandExecute({
        id: an.id,
        params: e
      });
      return L(K.redos, O), T.push(...K.redos), U.push(...K.undos), n.get(V).pushUndoRedo({
        unitID: g,
        undoMutations: U,
        redoMutations: T
      }), !0;
    }
    return !1;
  }
}, Wr = "sheet.command.move-cols", un = {
  id: Wr,
  type: v.COMMAND,
  // eslint-disable-next-line max-lines-per-function
  handler: (n, e) => {
    var H, F;
    const t = n.get(z), {
      fromRange: { startColumn: o },
      toRange: { startColumn: s },
      range: r
    } = e, i = r ? [Vr(r)] : t.getCurrentSelections(), a = i == null ? void 0 : i.filter(
      (K) => K.range.rangeType === j.COLUMN && K.range.startColumn <= o && o <= K.range.endColumn
    );
    if ((a == null ? void 0 : a.length) !== 1)
      return !1;
    const u = n.get(G), l = n.get(M), c = P(l, e);
    if (!c) return !1;
    const { workbook: d, worksheet: h } = c, g = d.getUnitId(), m = h.getSheetId(), f = n.get(Vt), C = n.get(it), p = a[0].range, w = a[0].primary, I = uo(p, h, !1);
    if (!N.equals(p, I))
      return f.emit(C.t("sheets.info.partOfCell")), !1;
    if (ll(s, h))
      return f.emit(C.t("sheets.info.acrossMergedCell")), !1;
    const S = {
      ...p,
      startColumn: s,
      endColumn: s + p.endColumn - p.startColumn
    }, y = {
      unitId: g,
      subUnitId: m,
      sourceRange: p,
      targetRange: S
    }, b = xa(n, y), O = n.get(E), k = u.onCommandExecute({ id: un.id, params: e }), T = [
      ...(H = k.preRedos) != null ? H : [],
      { id: qe.id, params: y }
    ], U = [
      ...(F = k.preUndos) != null ? F : [],
      { id: qe.id, params: b }
    ];
    if (w) {
      const K = p.endColumn - p.startColumn + 1, se = s - o < 0 ? S : {
        ...S,
        startColumn: S.startColumn - K,
        endColumn: S.endColumn - K
      }, le = {
        unitId: g,
        subUnitId: m,
        type: we.MOVE_END,
        selections: [{ range: se, primary: _e(se, h), style: null }]
      }, Se = {
        unitId: g,
        subUnitId: m,
        type: we.MOVE_END,
        selections: [{ range: p, primary: w, style: null }]
      };
      T.push({ id: ie.id, params: le }), U.push({ id: ie.id, params: Se });
    }
    if (T.push(...k.redos), U.push(...k.undos), L(T, O).result) {
      const K = u.afterCommandExecute({
        id: un.id,
        params: e
      });
      return L(K.redos, O), T.push(...K.redos), U.push(...K.undos), n.get(V).pushUndoRedo({
        unitID: g,
        undoMutations: U,
        redoMutations: T
      }), !0;
    }
    return !1;
  }
};
function Vr(n) {
  return {
    range: n,
    primary: null,
    style: null
  };
}
var cl = Object.getOwnPropertyDescriptor, dl = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? cl(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, hl = (n, e) => (t, o) => e(t, o, n);
let rt = class extends ue {
  constructor(e) {
    super();
    R(this, "_sheetSkeletonStore", /* @__PURE__ */ new Map());
    this._injector = e, this.disposeWithMe(() => {
      this._sheetSkeletonStore = /* @__PURE__ */ new Map();
    });
  }
  getSkeleton(e, t) {
    if (this._sheetSkeletonStore.has(e))
      return this._sheetSkeletonStore.get(e).get(t);
  }
  setSkeleton(e, t, o) {
    this._sheetSkeletonStore.has(e) || this._sheetSkeletonStore.set(e, /* @__PURE__ */ new Map()), this._sheetSkeletonStore.get(e).set(t, o);
  }
  deleteSkeleton(e, t) {
    this._sheetSkeletonStore.has(e) && this._sheetSkeletonStore.get(e).delete(t);
  }
};
rt = dl([
  hl(0, $(so))
], rt);
function Jo(n, e) {
  const t = new Y();
  return n.map((o) => X.transformRange(o, e)).forEach((o) => {
    X.foreach(o, (s, r) => {
      const i = e.getCellHeight(s, r);
      i && t.setValue(s, r, i);
    });
  }), t;
}
const gl = 1e4;
function Sn(n, e) {
  if (!e)
    return { suitableRanges: n, remainingRanges: [] };
  const t = e.worksheet.getColumnCount(), o = Math.ceil(gl / t), s = [], r = [], i = e.getOffsetRelativeToRowCol(0, e.scrollY).row, a = n.map((l) => {
    let c;
    return i >= l.startRow && i <= l.endRow ? c = 0 : i < l.startRow ? c = l.startRow - i : c = i - l.endRow, {
      range: l,
      distance: c,
      rowCount: l.endRow - l.startRow + 1
      // Number of rows in the range
    };
  });
  a.sort((l, c) => l.distance !== c.distance ? l.distance - c.distance : l.rowCount - c.rowCount);
  let u = 0;
  for (const l of a)
    if (u + l.rowCount <= o)
      s.push(l.range), u += l.rowCount;
    else {
      const c = o - u;
      if (c > 0) {
        const d = {
          ...l.range,
          endRow: l.range.startRow + c - 1
        }, h = {
          ...l.range,
          startRow: l.range.startRow + c
        };
        s.push(d), r.push(h), u = o;
      } else
        r.push(l.range);
    }
  return { suitableRanges: s, remainingRanges: r };
}
const Lr = "sheet.command.reorder-range", bo = {
  id: Lr,
  type: v.COMMAND,
  handler: (n, e) => {
    var I, S;
    const { subUnitId: t, unitId: o, range: s, order: r } = e, i = n.get(E), a = {
      id: Nn.id,
      params: {
        unitId: o,
        subUnitId: t,
        order: r,
        range: s
      }
    }, u = {
      id: Nn.id,
      params: Va(a.params)
    }, l = n.get(G), c = l.onCommandExecute({ id: bo.id, params: e }), d = [
      ...(I = c.preRedos) != null ? I : [],
      a,
      ...c.redos
    ], h = [
      ...(S = c.preUndos) != null ? S : [],
      u,
      ...c.undos
    ], g = L(d, i), { suitableRanges: m, remainingRanges: f } = Sn([s], n.get(rt).getSkeleton(o, t)), { undos: C, redos: p } = l.generateMutationsOfAutoHeight({
      unitId: o,
      subUnitId: t,
      ranges: [s],
      autoHeightRanges: m,
      lazyAutoHeightRanges: f
    }), w = l.afterCommandExecute({ id: bo.id, params: e });
    return g.result ? (L([...w.redos, ...p], i), n.get(V).pushUndoRedo({
      unitID: o,
      undoMutations: [...h, ...w.undos, ...C],
      redoMutations: [...d, ...w.redos, ...p]
    }), !0) : !1;
  }
}, W = {
  MoveRangeCommandId: tr,
  InsertRowCommandId: Pr,
  InsertColCommandId: Or,
  RemoveColCommandId: xn,
  RemoveRowCommandId: An,
  DeleteRangeMoveLeftCommandId: kr,
  DeleteRangeMoveUpCommandId: Tr,
  InsertRangeMoveDownCommandId: Zu,
  InsertRangeMoveRightCommandId: Ko,
  MoveColsCommandId: Wr,
  MoveRowsCommandId: $r,
  ReorderRangeCommandId: Lr
};
var Q = /* @__PURE__ */ ((n) => (n[n.Set = 0] = "Set", n[n.Delete = 1] = "Delete", n[n.HorizontalMove = 2] = "HorizontalMove", n[n.VerticalMove = 3] = "VerticalMove", n[n.Unknown = 4] = "Unknown", n))(Q || {});
const Mn = Number.MAX_SAFE_INTEGER, Ke = (n) => {
  const e = { ...n }, t = Number.isNaN(e.startRow) && Number.isNaN(e.endRow) && !Number.isNaN(e.startColumn) && !Number.isNaN(e.endColumn), o = Number.isNaN(e.startColumn) && Number.isNaN(e.endColumn) && !Number.isNaN(e.startRow) && !Number.isNaN(e.endRow);
  return (e.rangeType === j.COLUMN || t) && (e.startRow = 0, e.endRow = Mn), (e.rangeType === j.ROW || o) && (e.startColumn = 0, e.endColumn = Mn), e.rangeType === j.ALL && (e.startColumn = 0, e.endColumn = Mn, e.startRow = 0, e.endRow = Mn), e;
}, Oe = (n) => {
  let e = n.rangeType;
  return n.rangeType === j.COLUMN ? e = j.ROW : n.rangeType === j.ROW && (e = j.COLUMN), {
    startRow: n.startColumn,
    endRow: n.endColumn,
    startColumn: n.startRow,
    endColumn: n.endRow,
    rangeType: e
  };
}, $n = (n, e, t) => {
  const o = { ...t }, s = { ...e }, r = (f, C) => {
    const p = Math.max(f.start, C.start), w = Math.min(f.end, C.end);
    return w < p ? null : { start: p, end: w };
  }, i = (f) => f.end - f.start + 1, a = (f, C) => ({
    start: f.start - C.start,
    end: f.start - C.start + f.end - f.start
  }), u = (f, C) => ({
    start: C.start + f.start,
    end: C.start + f.start + f.end - f.start
  }), l = e.start > n.start;
  if (l) {
    const f = Math.min(n.end, e.start) - n.start + 1;
    s.start -= f, s.end -= f;
  }
  const c = i(n), d = c, h = r(n, o), g = h && i(h) >= i(o);
  if (n.end < o.start)
    o.start -= c, o.end -= c;
  else if (h) {
    const f = i(h);
    if (g) {
      const C = a(o, n), p = u(C, s);
      o.start = p.start, o.end = p.end;
    } else h.start > n.start ? l ? (o.end -= f + c, o.start -= c) : o.end -= f : l ? o.end -= f : o.start > n.start && o.end > n.end ? (o.start -= c, o.end -= c + f) : o.end -= f;
  }
  const m = r(s, o);
  return g || (s.start <= o.start ? (o.start += d, o.end += d) : m && (l ? s.end <= o.start || s.start <= o.start && s.end >= o.start ? (o.start += d, o.end += d) : s.start >= o.start && s.start <= o.end && (o.end += d) : o.start < s.start && o.end > s.start ? o.end += d : (o.start >= s.end || o.start >= s.start && o.start <= s.end) && (o.end += d, o.start += d))), {
    step: o.start - t.start,
    length: i(o) - i(t)
  };
}, Hr = (n, e) => {
  const { fromRange: t, toRange: o } = n.params || {};
  if (!o || !t)
    return [];
  const s = Ke(t), r = Ke(o), i = Ke(e), a = $n(
    { start: s.startRow, end: s.endRow },
    { start: r.startRow, end: r.endRow },
    { start: i.startRow, end: i.endRow }
  );
  return a === null ? [
    {
      type: Q.Delete
    }
  ] : [
    {
      type: Q.VerticalMove,
      step: a.step || 0,
      length: a.length || 0
    }
  ];
}, ml = (n, e) => {
  const { fromRange: t, toRange: o } = n.params || {};
  if (!t || !o)
    return [e];
  const s = t.startRow, r = t.endRow - t.startRow + 1, i = o.startRow, a = new Y();
  return X.foreach(e, (l, c) => {
    a.setValue(l, c, 1);
  }), a.moveRows(s, r, i), at(a, (l) => l === 1);
}, fl = (n, e) => {
  const { range: t, order: o } = n.params || {};
  if (!t || !o)
    return [e];
  const s = new Y();
  X.foreach(e, (a, u) => {
    s.setValue(a, u, 1);
  });
  const r = new Y();
  return X.foreach(t, (a, u) => {
    var l;
    if (Object.prototype.hasOwnProperty.call(o, a)) {
      const c = o[a], d = (l = s.getValue(c, u)) != null ? l : 0;
      r.setValue(a, u, d);
    }
  }), r.forValue((a, u, l) => {
    s.setValue(a, u, l);
  }), at(s, (a) => a === 1);
}, Fr = (n, e) => {
  const { fromRange: t, toRange: o } = n.params || {};
  if (!o || !t)
    return [];
  const s = Ke(t), r = Ke(o), i = Ke(e), a = $n(
    { start: s.startColumn, end: s.endColumn },
    { start: r.startColumn, end: r.endColumn },
    { start: i.startColumn, end: i.endColumn }
  );
  return a === null ? [
    {
      type: Q.Delete
    }
  ] : [
    {
      type: Q.HorizontalMove,
      step: a.step || 0,
      length: a.length || 0
    }
  ];
}, Rl = (n, e) => {
  const { fromRange: t, toRange: o } = n.params || {};
  if (!t || !o)
    return [e];
  const s = t.startColumn, r = t.endColumn - t.startColumn + 1, i = o.startColumn, a = new Y();
  return X.foreach(e, (u, l) => {
    a.setValue(u, l, 1);
  }), a.moveColumns(s, r, i), at(a, (u) => u === 1);
}, Cl = (n, e) => {
  var r, i;
  const t = (r = n.params) == null ? void 0 : r.toRange, o = (i = n.params) == null ? void 0 : i.fromRange;
  if (!t || !o)
    return [];
  const s = [];
  if (N.contains(t, e) && s.push({
    type: Q.Delete
  }), N.contains(o, e)) {
    s.push({
      type: Q.Delete
    });
    const a = N.getRelativeRange(e, o), u = N.getPositionRange(a, t);
    return [
      {
        type: Q.Set,
        range: u
      }
    ];
  }
  return s;
}, pl = (n, e) => {
  var d, h;
  const t = (d = n.params) == null ? void 0 : d.toRange, o = (h = n.params) == null ? void 0 : h.fromRange;
  if (!t || !o)
    return [e];
  if (!N.intersects(o, e) && !N.intersects(t, e))
    return [e];
  if (N.contains(o, e)) {
    const g = N.getRelativeRange(e, o);
    return [N.getPositionRange(g, t)];
  }
  const s = new Y();
  X.foreach(e, (g, m) => {
    s.setValue(g, m, 1);
  });
  const r = new Y(), i = N.getIntersects(o, e);
  i && X.foreach(i, (g, m) => {
    s.getValue(g, m) && (s.setValue(g, m, void 0), r.setValue(g, m, 1));
  });
  const a = t.startColumn - o.startColumn, u = t.startRow - o.startRow, l = {
    startColumn: t.startColumn - a,
    endColumn: t.endColumn - a,
    startRow: t.startRow - u,
    endRow: t.endRow - u
  };
  return l && X.foreach(l, (g, m) => {
    var p;
    const f = g + u, C = m + a;
    s.setValue(f, C, (p = r.getValue(g, m)) != null ? p : 0);
  }), at(s, (g) => g === 1);
}, Nt = (n, e) => {
  const t = Ke(n), o = Ke(e), s = (i) => i.endColumn - i.startColumn + 1, r = (i) => i.endRow - i.startRow + 1;
  if (t.startRow <= o.startRow && t.endRow >= o.endRow) {
    if (
      // 2
      o.startColumn < t.startColumn && o.endColumn >= t.startColumn && o.endColumn <= t.endColumn || // 6
      o.startColumn < t.startColumn && o.endColumn >= t.endColumn
    ) {
      const i = N.getIntersects(o, t);
      if (i)
        return { step: 0, length: -s(i) };
    }
    if (o.startColumn >= t.startColumn && o.endColumn <= t.endColumn && r(t) >= r(o))
      return null;
    if (o.startColumn >= t.startColumn && o.startColumn <= t.endColumn && o.endColumn > t.endColumn) {
      const i = N.getIntersects(o, t);
      if (i) {
        const a = -s(i);
        return { step: -(s(t) - s(i)), length: a };
      }
    }
    if (o.startColumn > t.endColumn)
      return { step: -s(t), length: 0 };
  }
  return { step: 0, length: 0 };
}, Br = (n, e) => {
  var r;
  const t = (r = n.params) == null ? void 0 : r.range;
  if (!t)
    return [];
  const o = [], s = Nt(t, e);
  if (!s)
    o.push({ type: Q.Delete });
  else {
    const { step: i, length: a } = s;
    o.push({
      type: Q.HorizontalMove,
      step: i,
      length: a
    });
  }
  return o;
}, Sl = (n, e, t) => {
  var i;
  const o = (i = n.params) == null ? void 0 : i.range;
  if (!o)
    return [];
  const s = [];
  if (t && t.length > 0) {
    let a = o.startRow;
    for (let u = o.startRow; u <= o.endRow; u++) {
      if (t.includes(u)) {
        if (u === a) {
          a = u + 1;
          continue;
        }
        r({
          ...o,
          startRow: a,
          endRow: u - 1
        }), a = u + 1;
        continue;
      }
      u === o.endRow && r({
        ...o,
        startRow: a,
        endRow: o.endRow
      });
    }
  } else
    r(o);
  function r(a) {
    const u = Nt(Oe(a), Oe(e));
    if (!u)
      s.push({ type: Q.Delete });
    else {
      const { step: l, length: c } = u;
      s.push({
        type: Q.VerticalMove,
        step: l,
        length: c
      });
    }
  }
  return s;
}, wl = (n, e) => {
  const { range: t, order: o } = n.params || {};
  if (!t || !o)
    return [];
  if (N.contains(t, e) && e.endRow === e.startRow) {
    const s = [], r = e.startRow;
    for (const i in o)
      if (o[i] === r) {
        const a = Number(i);
        return s.push({
          type: Q.VerticalMove,
          step: a - r,
          length: 0
        }), s;
      }
    return [];
  }
  return [];
}, Ot = (n, e) => {
  const t = Ke(n), o = Ke(e), s = (r) => r.endColumn - r.startColumn + 1;
  return t.startRow <= o.startRow && t.endRow >= o.endRow ? (
    // 2
    // Case 2: Overlap on the left side
    // Target range starts before the insert range and ends within the insert range boundaries
    // targetRange:  |----------|
    // insertRange:         |-------|
    // insertRange:
    o.startColumn < t.startColumn && o.endColumn >= t.startColumn && o.endColumn <= t.endColumn || // 6
    // Case 6: Fully overlapping on both sides
    // Target range starts before the insert range and ends after the insert range
    // targetRange:  |----------------|
    // insertRange:         |-------|
    o.startColumn < t.startColumn && o.endColumn >= t.endColumn ? { step: 0, length: s(t) } : (
      // 3
      // Case 3: Fully contained
      // Target range is completely within the insert range
      // targetRange:      |---|
      // insertRange:    |-------|
      o.startColumn >= t.startColumn && o.endColumn <= t.endColumn || // 4
      // Case 4: Overlap on the right side
      // Target range starts within the insert range and ends after the insert range
      // targetRange:         |---------|
      // insertRange:    |-------|
      o.startColumn >= t.startColumn && o.startColumn <= t.endColumn && o.endColumn > t.endColumn || //5
      // Case 5: No overlap (target range starts after the insert range ends)
      // targetRange:                |-------|
      // insertRange:    |-------|
      o.startColumn >= t.endColumn ? { step: s(t), length: 0 } : { step: 0, length: 0 }
    )
  ) : { step: 0, length: 0 };
};
function Il(n, e, t) {
  const o = [];
  if (N.contains(e, t) && o.push({
    type: Q.Delete
  }), N.contains(n, t)) {
    o.push({
      type: Q.Delete
    });
    const s = N.getRelativeRange(t, n), r = N.getPositionRange(s, e);
    return [
      {
        type: Q.Set,
        range: r
      }
    ];
  }
  return o;
}
const vl = (n, e) => {
  var a;
  const t = (a = n.params) == null ? void 0 : a.range;
  if (!t)
    return [];
  const o = [], s = Ot(Oe(t), Oe(e)), { step: r, length: i } = s;
  return o.push({
    type: Q.VerticalMove,
    step: r,
    length: i
  }), o;
}, Ml = (n, e) => {
  var a;
  const t = (a = n.params) == null ? void 0 : a.range;
  if (!t)
    return [];
  const o = [], s = Ot(t, e), { step: r, length: i } = s;
  return o.push({
    type: Q.HorizontalMove,
    step: r,
    length: i
  }), o;
}, _l = (n, e) => {
  var a;
  const t = (a = n.params) == null ? void 0 : a.range;
  if (!t)
    return [];
  const o = [], s = Ot(Oe(t), Oe(e)), { step: r, length: i } = s;
  return o.push({
    type: Q.VerticalMove,
    step: r,
    length: i
  }), o;
}, yl = (n, e) => {
  var u;
  const t = (u = n.params) == null ? void 0 : u.range;
  if (!t)
    return [e];
  const o = t.endRow - t.startRow + 1, s = {
    ...t,
    startRow: t.startRow,
    endRow: Number.POSITIVE_INFINITY
  }, r = N.subtract(e, s), i = N.getIntersects(s, e);
  if (!i)
    return [e];
  const a = new Y();
  return r.forEach((l) => {
    X.foreach(l, (c, d) => {
      a.setValue(c, d, 1);
    });
  }), i && X.foreach(i, (l, c) => {
    a.setValue(l + o, c, 1);
  }), at(a, (l) => l === 1);
}, bl = (n, e) => {
  var a;
  const t = (a = n.params) == null ? void 0 : a.range;
  if (!t)
    return [];
  const o = [], s = Ot(t, e), { step: r, length: i } = s;
  return o.push({
    type: Q.HorizontalMove,
    step: r,
    length: i
  }), o;
}, El = (n, e) => {
  var u;
  const t = (u = n.params) == null ? void 0 : u.range;
  if (!t)
    return [e];
  const o = t.endColumn - t.startColumn + 1, s = {
    ...t,
    startColumn: t.startColumn,
    endColumn: Number.POSITIVE_INFINITY
  }, r = N.subtract(e, s), i = N.getIntersects(s, e);
  if (!i)
    return [e];
  const a = new Y();
  return r.forEach((l) => {
    X.foreach(l, (c, d) => {
      a.setValue(c, d, 1);
    });
  }), i && X.foreach(i, (l, c) => {
    a.setValue(l, c + o, 1);
  }), at(a, (l) => l === 1);
}, Ul = (n, e) => {
  var r;
  const t = (r = n.params) == null ? void 0 : r.range;
  if (!t)
    return [];
  const o = [], s = Nt(t, e);
  if (!s)
    o.push({ type: Q.Delete });
  else {
    const { step: i, length: a } = s;
    o.push({
      type: Q.HorizontalMove,
      step: i,
      length: a
    });
  }
  return o;
}, kl = (n, e) => {
  var l;
  const t = (l = n.params) == null ? void 0 : l.range;
  if (!t)
    return [e];
  const o = {
    startRow: t.startRow,
    endRow: t.endRow,
    startColumn: t.startColumn,
    endColumn: Number.POSITIVE_INFINITY
  }, s = t.endColumn - t.startColumn + 1, r = N.getIntersects(t, e), i = N.subtract(e, o), a = N.getIntersects(o, e);
  if (!r && !a)
    return [e];
  const u = new Y();
  return a && X.foreach(a, (c, d) => {
    u.setValue(c, d - s, 1);
  }), r && X.foreach(r, (c, d) => {
    u.setValue(c, d - s, 0);
  }), i.forEach((c) => {
    X.foreach(c, (d, h) => {
      u.setValue(d, h, 1);
    });
  }), at(u, (c) => c === 1);
}, Tl = (n, e) => {
  var r;
  const t = (r = n.params) == null ? void 0 : r.range;
  if (!t)
    return [];
  const o = [], s = Nt(Oe(t), Oe(e));
  if (!s)
    o.push({ type: Q.Delete });
  else {
    const { step: i, length: a } = s;
    o.push({
      type: Q.VerticalMove,
      step: i,
      length: a
    });
  }
  return o;
}, Pl = (n, e) => {
  var l;
  const t = (l = n.params) == null ? void 0 : l.range;
  if (!t)
    return [e];
  const o = {
    ...t,
    startRow: t.startRow,
    endRow: Number.POSITIVE_INFINITY
  }, s = t.endRow - t.startRow + 1, r = N.getIntersects(t, e), i = N.subtract(e, o), a = N.getIntersects(o, e);
  if (!r && !a)
    return [e];
  const u = new Y();
  return a && X.foreach(a, (c, d) => {
    u.setValue(c - s, d, 1);
  }), r && X.foreach(r, (c, d) => {
    u.setValue(c - s, d, 0);
  }), i.forEach((c) => {
    X.foreach(c, (d, h) => {
      u.setValue(d, h, 1);
    });
  }), at(u, (c) => c === 1);
}, Nl = (n, e) => {
  var s;
  const t = (s = n.ranges) != null ? s : [n.range], o = new Y();
  return X.foreach(e, (r, i) => {
    o.setValue(r, i, 1);
  }), t.forEach((r) => {
    const i = r.startRow, u = r.endRow - i + 1;
    o.removeRows(i, u);
  }), at(o, (r) => r === 1);
}, Ol = (n, e) => {
  const t = n.params, o = t.range.startRow, s = t.range.endRow - t.range.startRow + 1;
  return e.startRow >= o ? [{
    startRow: e.startRow + s,
    endRow: e.endRow + s,
    startColumn: e.startColumn,
    endColumn: e.endColumn
  }] : e.endRow < o ? [e] : [{
    startRow: e.startRow,
    endRow: e.endRow + s,
    startColumn: e.startColumn,
    endColumn: e.endColumn
  }];
}, Dl = (n, e) => {
  const t = n.params, o = t.range.startColumn, s = t.range.endColumn - t.range.startColumn + 1;
  return e.startColumn >= o ? [{
    startRow: e.startRow,
    endRow: e.endRow,
    startColumn: e.startColumn + s,
    endColumn: e.endColumn + s
  }] : e.endColumn < o ? [e] : [{
    startRow: e.startRow,
    endRow: e.endRow,
    startColumn: e.startColumn,
    endColumn: e.endColumn + s
  }];
}, Dt = (n, e) => {
  let t = { ...e };
  return n.forEach((o) => {
    switch (o.type) {
      case Q.Delete: {
        t = null;
        break;
      }
      case Q.HorizontalMove: {
        if (!t)
          return;
        t.startColumn += o.step, t.endColumn += o.step + (o.length || 0);
        break;
      }
      case Q.VerticalMove: {
        if (!t)
          return;
        t.startRow += o.step, t.endRow += o.step + (o.length || 0);
        break;
      }
      case Q.Set: {
        t = o.range;
        break;
      }
    }
  }), t && (t.endColumn < t.startColumn || t.endRow < t.startRow) ? null : t;
}, Rs = (n, e) => {
  let t = [];
  switch (e.id) {
    case W.DeleteRangeMoveLeftCommandId: {
      t = Ul(e, n);
      break;
    }
    case W.DeleteRangeMoveUpCommandId: {
      t = Tl(e, n);
      break;
    }
    case W.InsertColCommandId: {
      t = Ml(e, n);
      break;
    }
    case W.InsertRangeMoveDownCommandId: {
      t = _l(e, n);
      break;
    }
    case W.InsertRangeMoveRightCommandId: {
      t = bl(e, n);
      break;
    }
    case W.InsertRowCommandId: {
      t = vl(e, n);
      break;
    }
    case W.MoveColsCommandId: {
      t = Fr(e, n);
      break;
    }
    case W.MoveRangeCommandId: {
      t = Cl(e, n);
      break;
    }
    case W.MoveRowsCommandId: {
      t = Hr(e, n);
      break;
    }
    case W.RemoveColCommandId: {
      t = Br(e, n);
      break;
    }
    case W.RemoveRowCommandId: {
      t = Sl(e, n);
      break;
    }
    case W.ReorderRangeCommandId: {
      t = wl(e, n);
      break;
    }
  }
  return Dt(t, n);
}, Ah = (n, e, t) => [It.id, vt.id].includes(e.id) || jr(e, t).some((r) => N.intersects(r, n)) ? Rs(n, e) : n, Cs = (n, e) => {
  let t = [];
  switch (e.id) {
    case W.DeleteRangeMoveLeftCommandId:
      return kl(e, n);
    case W.DeleteRangeMoveUpCommandId:
      return Pl(e, n);
    case W.InsertRangeMoveDownCommandId:
      return yl(e, n);
    case W.InsertRangeMoveRightCommandId:
      return El(e, n);
    case W.InsertColCommandId:
      return Dl(e, n);
    case W.InsertRowCommandId:
      return Ol(e, n);
    case W.MoveColsCommandId:
      return Rl(e, n);
    case W.MoveRangeCommandId:
      return pl(e, n);
    case W.MoveRowsCommandId:
      return ml(e, n);
    case W.ReorderRangeCommandId:
      return fl(e, n);
    case W.RemoveColCommandId: {
      t = Br(e, n);
      break;
    }
    case W.RemoveRowCommandId:
      return Nl(e.params, n);
  }
  const o = Dt(t, n);
  return o ? [o] : [];
}, xh = (n, e, t) => [It.id, vt.id, Pt.id, Ko].includes(e.id) || jr(e, t).some((r) => N.intersects(r, n)) ? Cs(n, e) : n;
function Al(n, e) {
  const { id: t, params: o } = e;
  let s = {
    length: 0,
    step: 0,
    type: Q.Unknown
  };
  switch (t) {
    case Lt.id:
      s.type = Q.Delete;
      break;
    case Je.id:
      s = $n(
        { start: o.sourceRange.startRow, end: o.sourceRange.endRow },
        { start: o.targetRange.startRow, end: o.targetRange.endRow },
        { start: n.startRow, end: n.endRow }
      ), s.type = Q.VerticalMove;
      break;
    case qe.id:
      s = $n(
        { start: o.sourceRange.startColumn, end: o.sourceRange.endColumn },
        { start: o.targetRange.startColumn, end: o.targetRange.endColumn },
        { start: n.startColumn, end: n.endColumn }
      ), s.type = Q.HorizontalMove;
      break;
    case Me.id:
      s = Nt(o.range, n), s ? s.type = Q.HorizontalMove : s = { step: 0, length: 0, type: Q.Delete };
      break;
    case Te.id:
      s = Nt(Oe(o.range), Oe(n)), s ? s.type = Q.VerticalMove : s = { step: 0, length: 0, type: Q.Delete };
      break;
    case Ue.id:
      s = Ot(Oe(o.range), Oe(n)), s.type = Q.VerticalMove;
      break;
    case ke.id:
      s = Ot(o.range, n), s.type = Q.HorizontalMove;
      break;
    case wt.id:
      {
        const r = o.fromRange || new Y(o.from).getRange(), i = o.toRange || new Y(o.to).getRange();
        s = Il(
          r,
          i,
          n
        );
      }
      break;
  }
  return s ? Array.isArray(s) ? Dt(s, n) : Dt([s], n) : n;
}
function jr(n, e) {
  var o, s, r, i, a, u;
  const { selectionManagerService: t } = e;
  switch (n.id) {
    case W.MoveColsCommandId: {
      const l = n.params;
      return [
        l.fromRange,
        {
          ...l.toRange,
          startColumn: l.toRange.startColumn - 0.5,
          endColumn: l.toRange.endColumn - 0.5
        }
      ];
    }
    case W.MoveRowsCommandId: {
      const l = n.params;
      return [
        l.fromRange,
        {
          ...l.toRange,
          startRow: l.toRange.startRow - 0.5,
          endRow: l.toRange.startRow - 0.5
        }
      ];
    }
    case W.MoveRangeCommandId: {
      const l = n;
      return [l.params.fromRange, l.params.toRange];
    }
    case W.InsertRowCommandId: {
      const c = n.params.range;
      return [
        {
          ...c,
          startRow: c.startRow - 0.5,
          endRow: c.endRow - 0.5
        }
      ];
    }
    case W.InsertColCommandId: {
      const c = n.params.range;
      return [
        {
          ...c,
          startColumn: c.startColumn - 0.5,
          endColumn: c.endColumn - 0.5
        }
      ];
    }
    case W.RemoveRowCommandId:
      return [n.params.range];
    case W.RemoveColCommandId:
      return [n.params.range];
    case W.DeleteRangeMoveUpCommandId:
    case W.InsertRangeMoveDownCommandId: {
      const c = ((o = n.params) == null ? void 0 : o.range) || ((r = (s = t.getCurrentSelections()) == null ? void 0 : s.map((d) => d.range)) == null ? void 0 : r[0]);
      return c ? [c] : [];
    }
    case W.DeleteRangeMoveLeftCommandId:
    case W.InsertRangeMoveRightCommandId: {
      const c = ((i = n.params) == null ? void 0 : i.range) || ((u = (a = t.getCurrentSelections()) == null ? void 0 : a.map((d) => d.range)) == null ? void 0 : u[0]);
      return c ? [c] : [];
    }
    case W.ReorderRangeCommandId: {
      const l = n, { range: c, order: d } = l.params, h = [];
      for (let g = c.startRow; g <= c.endRow; g++)
        g in d && h.push({
          startRow: g,
          endRow: g,
          startColumn: c.startColumn,
          endColumn: c.endColumn
        });
      return h;
    }
  }
}
function xl(n) {
  switch (n.id) {
    case qe.id: {
      const e = n.params;
      return [
        e.sourceRange,
        {
          ...e.targetRange,
          startColumn: e.targetRange.startColumn - 0.5,
          endColumn: e.targetRange.startColumn - 0.5
        }
      ];
    }
    case Je.id: {
      const e = n.params;
      return [
        e.sourceRange,
        {
          ...e.targetRange,
          startRow: e.targetRange.startRow - 0.5,
          endRow: e.targetRange.startRow - 0.5
        }
      ];
    }
    case wt.id: {
      const e = n.params;
      return [new Y(e.from.value).getRange(), new Y(e.to.value).getRange()];
    }
    case ke.id: {
      const t = n.params.range;
      return [
        {
          ...t,
          startColumn: t.startColumn - 0.5,
          endColumn: t.startColumn - 0.5
        }
      ];
    }
    case Ue.id: {
      const t = n.params.range;
      return [
        {
          ...t,
          startRow: t.startRow - 0.5,
          endRow: t.startRow - 0.5
        }
      ];
    }
    case Me.id:
      return [n.params.range];
    case Te.id:
      return [n.params.range];
  }
}
function $h(n, e) {
  var s, r, i, a, u, l;
  const t = n.get(M), o = n.get(z);
  switch (e.id) {
    case W.MoveColsCommandId: {
      const c = e.params, d = P(t, {
        unitId: c.unitId,
        subUnitId: c.subUnitId
      });
      return {
        unitId: d.unitId,
        subUnitId: d.subUnitId,
        ranges: [
          c.fromRange,
          {
            ...c.toRange,
            startColumn: c.fromRange.startColumn < c.toRange.startColumn ? c.fromRange.endColumn + 1 : c.toRange.startColumn,
            endColumn: c.fromRange.startColumn < c.toRange.startColumn ? c.toRange.endColumn - 1 : c.fromRange.startColumn - 1
          }
        ]
      };
    }
    case W.MoveRowsCommandId: {
      const c = e.params, d = P(t, {
        unitId: c.unitId,
        subUnitId: c.subUnitId
      });
      return {
        unitId: d.unitId,
        subUnitId: d.subUnitId,
        ranges: [
          c.fromRange,
          {
            ...c.toRange,
            startRow: c.fromRange.startRow < c.toRange.startRow ? c.fromRange.endRow + 1 : c.toRange.startRow,
            endRow: c.fromRange.startRow < c.toRange.startRow ? c.toRange.endRow - 1 : c.fromRange.startRow - 1
          }
        ]
      };
    }
    case W.MoveRangeCommandId: {
      const c = e.params, d = P(t);
      return {
        unitId: d.unitId,
        subUnitId: d.subUnitId,
        ranges: [c.fromRange, c.toRange]
      };
    }
    case W.InsertRowCommandId: {
      const c = e.params, d = c.range;
      return {
        unitId: c.unitId,
        subUnitId: c.subUnitId,
        ranges: [
          ...d.startRow > 0 ? [{
            ...d,
            startRow: d.startRow - 1,
            endRow: d.endRow - 1
          }] : [],
          {
            ...d,
            startRow: d.startRow,
            endRow: Number.MAX_SAFE_INTEGER
          }
        ]
      };
    }
    case W.InsertColCommandId: {
      const c = e.params, d = c.range;
      return {
        unitId: c.unitId,
        subUnitId: c.subUnitId,
        ranges: [
          ...d.startColumn > 0 ? [{
            ...d,
            startColumn: d.startColumn - 1,
            endColumn: d.endColumn - 1
          }] : [],
          {
            ...d,
            startColumn: d.startColumn,
            endColumn: Number.MAX_SAFE_INTEGER
          }
        ]
      };
    }
    case W.RemoveRowCommandId: {
      const d = e.params.range, h = P(t);
      return {
        unitId: h.unitId,
        subUnitId: h.subUnitId,
        ranges: [
          d,
          {
            ...d,
            startRow: d.endRow + 1,
            endRow: Number.MAX_SAFE_INTEGER
          }
        ]
      };
    }
    case W.RemoveColCommandId: {
      const d = e.params.range, h = P(t);
      return {
        unitId: h.unitId,
        subUnitId: h.subUnitId,
        ranges: [
          d,
          {
            ...d,
            startColumn: d.endColumn + 1,
            endColumn: Number.MAX_SAFE_INTEGER
          }
        ]
      };
    }
    case W.DeleteRangeMoveUpCommandId:
    case W.InsertRangeMoveDownCommandId: {
      const c = e, d = P(t), h = ((s = c.params) == null ? void 0 : s.range) || ((i = (r = o.getCurrentSelections()) == null ? void 0 : r.map((g) => g.range)) == null ? void 0 : i[0]);
      return h ? {
        unitId: d.unitId,
        subUnitId: d.subUnitId,
        ranges: [
          h,
          {
            ...h,
            startRow: h.endRow + 1,
            endRow: Number.MAX_SAFE_INTEGER
          }
        ]
      } : {
        unitId: d.unitId,
        subUnitId: d.subUnitId,
        ranges: []
      };
    }
    case W.DeleteRangeMoveLeftCommandId:
    case W.InsertRangeMoveRightCommandId: {
      const d = ((a = e.params) == null ? void 0 : a.range) || ((l = (u = o.getCurrentSelections()) == null ? void 0 : u.map((g) => g.range)) == null ? void 0 : l[0]), h = P(t);
      return d ? {
        unitId: h.unitId,
        subUnitId: h.subUnitId,
        ranges: [
          d,
          {
            ...d,
            startColumn: d.endColumn + 1,
            endColumn: Number.MAX_SAFE_INTEGER
          }
        ]
      } : {
        unitId: h.unitId,
        subUnitId: h.subUnitId,
        ranges: []
      };
    }
    case W.ReorderRangeCommandId: {
      const c = e, { range: d, order: h } = c.params, g = [];
      for (let f = d.startRow; f <= d.endRow; f++)
        f in h && g.push({
          startRow: f,
          endRow: f,
          startColumn: d.startColumn,
          endColumn: d.endColumn
        });
      const m = P(t);
      return {
        unitId: m.unitId,
        subUnitId: m.subUnitId,
        ranges: g
      };
    }
  }
}
var $l = Object.getOwnPropertyDescriptor, Wl = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? $l(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, _n = (n, e) => (t, o) => e(t, o, n);
const Vl = ft("MERGE_REDO"), Ll = ft("MERGE_UNDO"), ps = Math.floor(Number.MAX_SAFE_INTEGER / 10);
class Hl extends ue {
  constructor(e, t, o, s, r = !1) {
    super(), this._unitId = e, this._subUnitId = t, this._range = o, this._callback = s, this._skipIntersects = r;
  }
  onMutation(e) {
    var s, r;
    if (((s = e.params) == null ? void 0 : s.unitId) !== this._unitId)
      return;
    if (e.id === wt.id) {
      const i = e.params;
      if (i.from.subUnitId !== this._subUnitId || i.to.subUnitId !== this._subUnitId)
        return;
    } else if (((r = e.params) == null ? void 0 : r.subUnitId) !== this._subUnitId)
      return;
    if (!this._range)
      return;
    if (this._skipIntersects) {
      if (e.id === Lt.id)
        return;
      const i = xl(e);
      if (i != null && i.some((a) => N.intersects(a, this._range)))
        return;
    }
    const t = Al(this._range, e);
    if (t && N.equals(t, this._range))
      return !1;
    const o = this._range;
    this._range = t, this._callback(o, t);
  }
}
let At = class extends ue {
  constructor(e, t, o, s) {
    super();
    R(this, "interceptor", new Oo({ MERGE_REDO: Vl, MERGE_UNDO: Ll }));
    R(this, "_watchRanges", /* @__PURE__ */ new Set());
    R(this, "_refRangeManagerMap", /* @__PURE__ */ new Map());
    R(this, "_serializer", Fl());
    // eslint-disable-next-line max-lines-per-function
    R(this, "_onRefRangeChange", () => {
      this._sheetInterceptorService.interceptCommand({
        // eslint-disable-next-line max-lines-per-function
        getMutations: (e) => {
          const t = this._univerInstanceService.getCurrentUnitForType(B.UNIVER_SHEET).getActiveSheet(), o = Ss(this._univerInstanceService), s = ws(this._univerInstanceService);
          if (!t || !o || !s)
            return { redos: [], undos: [], preRedos: [], preUndos: [] };
          const a = ((() => {
            switch (e.id) {
              case W.MoveColsCommandId: {
                const h = e.params, g = Math.min(h.fromRange.startColumn, h.toRange.startColumn);
                return this._checkRange(
                  [{ ...h.fromRange, startColumn: g, endColumn: t.getColumnCount() - 1 }],
                  o,
                  s
                );
              }
              case W.MoveRowsCommandId: {
                const h = e.params, g = Math.min(h.fromRange.startRow, h.toRange.startRow);
                return this._checkRange(
                  [{ ...h.fromRange, startRow: g, endRow: t.getRowCount() - 1 }],
                  o,
                  s
                );
              }
              case W.MoveRangeCommandId: {
                const h = e;
                return this._checkRange(
                  [h.params.fromRange, h.params.toRange],
                  o,
                  s
                );
              }
              case W.InsertRowCommandId: {
                const m = {
                  startRow: e.params.range.startRow,
                  endRow: t.getRowCount() - 1,
                  startColumn: 0,
                  endColumn: t.getColumnCount() - 1,
                  rangeType: j.ROW
                };
                return this._checkRange([m], o, s);
              }
              case W.InsertColCommandId: {
                const g = e.params.range.startColumn, m = {
                  startRow: 0,
                  endRow: t.getRowCount() - 1,
                  startColumn: g,
                  endColumn: t.getColumnCount() - 1,
                  rangeType: j.COLUMN
                };
                return this._checkRange([m], o, s);
              }
              case W.RemoveRowCommandId: {
                const m = {
                  startRow: e.params.range.startRow,
                  endRow: t.getRowCount() - 1,
                  startColumn: 0,
                  endColumn: t.getColumnCount() - 1,
                  rangeType: j.ROW
                };
                return this._checkRange([m], o, s);
              }
              case W.RemoveColCommandId: {
                const g = e.params.range.startColumn, m = {
                  startRow: 0,
                  endRow: t.getRowCount() - 1,
                  startColumn: g,
                  endColumn: t.getColumnCount() - 1,
                  rangeType: j.COLUMN
                };
                return this._checkRange([m], o, s);
              }
              case W.DeleteRangeMoveUpCommandId:
              case W.InsertRangeMoveDownCommandId: {
                const g = e.params.range || Is(this._selectionManagerService)[0], m = {
                  startRow: g.startRow,
                  startColumn: g.startColumn,
                  endColumn: g.endColumn,
                  endRow: ps
                };
                return this._checkRange([m], o, s);
              }
              case W.DeleteRangeMoveLeftCommandId:
              case W.InsertRangeMoveRightCommandId: {
                const g = e.params.range || Is(this._selectionManagerService)[0], m = {
                  startRow: g.startRow,
                  startColumn: g.startColumn,
                  endColumn: ps,
                  endRow: g.endRow
                };
                return this._checkRange([m], o, s);
              }
              case W.ReorderRangeCommandId: {
                const h = e, { range: g, order: m } = h.params, f = [];
                for (let C = g.startRow; C <= g.endRow; C++)
                  C in m && f.push({
                    startRow: C,
                    endRow: C,
                    startColumn: g.startColumn,
                    endColumn: g.endColumn
                  });
                return this._checkRange(f, o, s);
              }
            }
          })() || []).reduce(
            (h, g) => {
              const m = g(e);
              return h.push(m), h;
            },
            []
          ).reduce(
            (h, g) => {
              var m, f;
              return h.redos.push(...g.redos), h.undos.push(...g.undos), h.preRedos.push(...(m = g.preRedos) != null ? m : []), h.preUndos.push(...(f = g.preUndos) != null ? f : []), h;
            },
            { redos: [], undos: [], preUndos: [], preRedos: [] }
          ), u = this.interceptor.fetchThroughInterceptors(this.interceptor.getInterceptPoints().MERGE_REDO)(
            a.preRedos,
            null
          ) || [], l = this.interceptor.fetchThroughInterceptors(this.interceptor.getInterceptPoints().MERGE_REDO)(
            a.redos,
            null
          ) || [], c = this.interceptor.fetchThroughInterceptors(this.interceptor.getInterceptPoints().MERGE_UNDO)(
            a.preUndos,
            null
          ) || [], d = this.interceptor.fetchThroughInterceptors(this.interceptor.getInterceptPoints().MERGE_UNDO)(
            a.undos,
            null
          ) || [];
          return { redos: l, undos: d, preRedos: u, preUndos: c };
        }
      });
    });
    R(this, "_checkRange", (e, t, o) => {
      const s = vs(t, o), r = this._refRangeManagerMap.get(s);
      if (r) {
        const i = /* @__PURE__ */ new Set();
        return [...r.keys()].forEach((u) => {
          const l = r.get(u), c = this._serializer.deserialize(u), d = {
            ...c,
            startRow: +c.startRow,
            endRow: +c.endRow,
            startColumn: +c.startColumn,
            endColumn: +c.endColumn,
            rangeType: c.rangeType && +c.rangeType
          };
          e.some((h) => N.intersects(h, d)) && l && l.forEach((h) => {
            i.add(h);
          });
        }), [...i];
      }
      return [];
    });
    /**
     * Listens to an area and triggers a fall back when movement occurs
     * @param {IRange} range the area that needs to be monitored
     * @param {RefRangCallback} callback the callback function that is executed when the range changes
     * @param {string} [_unitId]
     * @param {string} [_subUnitId]
     * @memberof RefRangeService
     */
    R(this, "registerRefRange", (e, t, o, s) => {
      const r = o || Ss(this._univerInstanceService), i = s || ws(this._univerInstanceService);
      if (!r || !i)
        return Ne(() => {
        });
      const a = vs(r, i), u = this._serializer.serialize(e);
      let l = this._refRangeManagerMap.get(a);
      l || (l = /* @__PURE__ */ new Map(), this._refRangeManagerMap.set(a, l));
      const c = l.get(u);
      return c ? c.add(t) : l.set(u, /* @__PURE__ */ new Set([t])), Ne(() => {
        const d = l.get(u);
        d && (d.delete(t), d.size || (l.delete(u), l.size || this._refRangeManagerMap.delete(a)));
      });
    });
    this._commandService = e, this._sheetInterceptorService = t, this._univerInstanceService = o, this._selectionManagerService = s, this._onRefRangeChange(), this.interceptor.intercept(this.interceptor.getInterceptPoints().MERGE_REDO, {
      priority: -1,
      handler: (r) => r
    }), this.interceptor.intercept(this.interceptor.getInterceptPoints().MERGE_UNDO, {
      priority: -1,
      handler: (r) => r
    });
  }
  watchRange(e, t, o, s, r) {
    let i;
    this._watchRanges.size === 0 && (i = this._commandService.onCommandExecuted((c) => {
      if (c.type !== v.MUTATION) return !1;
      for (const d of this._watchRanges)
        d.onMutation(c);
    }));
    const a = new Hl(e, t, o, s, r);
    this._watchRanges.add(a);
    const u = Ne(() => {
      this._watchRanges.delete(a), this._watchRanges.size === 0 && (i == null || i.dispose(), i = null);
    }), l = this.disposeWithMe(u);
    return Ne(() => {
      l.dispose(), u.dispose();
    });
  }
};
At = Wl([
  _n(0, E),
  _n(1, $(G)),
  _n(2, $(M)),
  _n(3, $(z))
], At);
function Ss(n) {
  return n.getCurrentUnitForType(B.UNIVER_SHEET).getUnitId();
}
function ws(n) {
  var e;
  return (e = n.getCurrentUnitForType(B.UNIVER_SHEET).getActiveSheet()) == null ? void 0 : e.getSheetId();
}
function Is(n) {
  var e;
  return ((e = n.getCurrentSelections()) == null ? void 0 : e.map((t) => t.range)) || [];
}
function vs(n, e) {
  return `${n}_${e}`;
}
function Fl() {
  const n = ["startRow", "startColumn", "endRow", "endColumn", "rangeType"];
  return {
    deserialize: (t) => {
      const o = n.reduce(
        (r, i, a) => (r[String(a)] = i, r),
        {}
      );
      return t.split("_").reduce(
        (r, i, a) => {
          const u = String(a);
          return i && o[u] && (r[o[u]] = i), r;
        },
        {}
      );
    },
    serialize: (t) => n.reduce((o, s, r) => {
      const i = t[s];
      return i !== void 0 ? `${o}${r > 0 ? "_" : ""}${i}` : `${o}`;
    }, "")
  };
}
var Bl = Object.getOwnPropertyDescriptor, jl = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? Bl(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, _t = (n, e) => (t, o) => e(t, o, n);
const zl = [ke.id, Ue.id, Me.id, Te.id], Gl = [Je.id, qe.id];
function zr(n, e) {
  let t = n;
  if (e !== void 0) {
    const o = [];
    for (let s = 0; s < t.length; s++) {
      const { startRow: r, endRow: i, startColumn: a, endColumn: u } = t[s];
      if (e === Re.ROWS)
        for (let l = r; l <= i; l++) {
          const c = {
            startRow: l,
            endRow: l,
            startColumn: a,
            endColumn: u
          };
          o.push(c);
        }
      else if (e === Re.COLUMNS)
        for (let l = a; l <= u; l++) {
          const c = {
            startRow: r,
            endRow: i,
            startColumn: l,
            endColumn: l
          };
          o.push(c);
        }
    }
    t = o;
  }
  return t;
}
const Kl = ft("mergeCellPermissionCheck");
let Wn = class extends ue {
  constructor(e, t, o, s, r, i) {
    super();
    R(this, "disposableCollection", new Et());
    R(this, "interceptor", new Oo({ MERGE_CELL_INTERCEPTOR_CHECK: Kl }));
    this._commandService = e, this._refRangeService = t, this._univerInstanceService = o, this._injector = s, this._sheetInterceptorService = r, this._selectionManagerService = i, this._onRefRangeChange(), this._initCommandInterceptor(), this._commandExecutedListener();
  }
  _initCommandInterceptor() {
    const e = this;
    this._sheetInterceptorService.interceptCommand({
      getMutations(t) {
        var o;
        switch (t.id) {
          case Vo.id:
          case Lo.id: {
            const s = e._univerInstanceService.getCurrentUnitForType(B.UNIVER_SHEET), r = s.getUnitId(), i = s == null ? void 0 : s.getActiveSheet();
            if (!i)
              return { redos: [], undos: [] };
            const a = i.getSheetId(), u = i.getConfig().mergeData, l = (o = e._selectionManagerService.getCurrentSelections()) == null ? void 0 : o.map((c) => c.range);
            if (l && l.length > 0 && l.some(
              (d) => u.some((h) => N.intersects(h, d))
            )) {
              const d = {
                unitId: r,
                subUnitId: a,
                ranges: l
              }, h = ye(e._injector, d), g = [
                { id: oe.id, params: d }
              ], m = [
                { id: ne.id, params: h }
              ];
              return { redos: g, undos: m };
            }
          }
        }
        return { redos: [], undos: [] };
      }
    }), this._sheetInterceptorService.interceptRanges({
      getMutations: ({ unitId: t, subUnitId: o, ranges: s }) => {
        const r = [], i = [], a = { redos: r, undos: i };
        if (!s || !s.length)
          return a;
        const u = P(this._univerInstanceService, { unitId: t, subUnitId: o });
        if (!u)
          return a;
        const { worksheet: l } = u, d = l.getMergeData().filter((h) => s.some((g) => N.intersects(h, g)));
        return d.length ? (r.push({
          id: oe.id,
          params: {
            unitId: t,
            subUnitId: o,
            ranges: d
          }
        }), i.push({
          id: ne.id,
          params: {
            unitId: t,
            subUnitId: o,
            ranges: d
          }
        }), { undos: i, redos: r }) : a;
      }
    });
  }
  refRangeHandle(e, t, o) {
    switch (e.id) {
      case W.MoveColsCommandId: {
        const s = e.params;
        return this._handleMoveColsCommand(s, t, o);
      }
      case W.MoveRowsCommandId: {
        const s = e.params;
        return this._handleMoveRowsCommand(s, t, o);
      }
      case Ye.id: {
        const s = e.params, r = s.unitId || t, i = s.subUnitId || o;
        return this._handleInsertRowCommand(s, r, i);
      }
      case Xe.id: {
        const s = e.params, r = s.unitId || t, i = s.subUnitId || o;
        return this._handleInsertColCommand(s, r, i);
      }
      case ho.id: {
        const s = e.params;
        return this._handleRemoveColCommand(s, t, o);
      }
      case co.id: {
        const s = e.params;
        return this._handleRemoveRowCommand(s, t, o);
      }
      case yt.id: {
        const s = e.params;
        return this._handleMoveRangeCommand(s, t, o);
      }
      case rn.id: {
        const s = e.params;
        return this._handleInsertRangeMoveRightCommand(s, t, o);
      }
      case Pt.id: {
        const s = e.params;
        return this._handleInsertRangeMoveDownCommand(s, t, o);
      }
      case vt.id: {
        const s = e.params;
        return this._handleDeleteRangeMoveUpCommand(s, t, o);
      }
      case It.id: {
        const s = e.params;
        return this._handleDeleteRangeMoveLeftCommand(s, t, o);
      }
    }
    return { redos: [], undos: [] };
  }
  _onRefRangeChange() {
    const e = (t, o) => {
      const s = this._univerInstanceService.getUniverSheetInstance(t);
      if (!s)
        return;
      const r = s == null ? void 0 : s.getSheetBySheetId(o);
      if (!r)
        return;
      this.disposableCollection.dispose();
      const i = r.getMergeData(), a = (u) => this.refRangeHandle(u, t, o);
      i.forEach((u) => {
        this.disposableCollection.add(this._refRangeService.registerRefRange(u, a, t, o));
      });
    };
    this.disposeWithMe(
      this._commandService.onCommandExecuted((t) => {
        if (t.id === Cn.id) {
          const o = t.params, s = o.subUnitId, r = o.unitId;
          if (!s || !r)
            return;
          e(r, s);
        }
        if (t.id === ne.id) {
          const o = t.params, s = o.subUnitId, r = o.unitId;
          if (!s || !r)
            return;
          e(o.unitId, o.subUnitId);
        }
      })
    ), this._univerInstanceService.getCurrentTypeOfUnit$(B.UNIVER_SHEET).pipe(qi((t) => !!t)).subscribe((t) => {
      const o = t.getActiveSheet();
      o && e(t.getUnitId(), o.getSheetId());
    });
  }
  _handleMoveRowsCommand(e, t, o) {
    const s = xe(this._univerInstanceService, t);
    if (!s)
      return this._handleNull();
    const r = $e(s, o);
    if (!r)
      return this._handleNull();
    const i = [...r.getMergeData()], a = { unitId: t, subUnitId: o, ranges: [] }, u = { unitId: t, subUnitId: o, ranges: [] }, { fromRange: l } = e, { startRow: c, endRow: d } = l;
    if (i.forEach((m) => {
      if (c <= m.startRow && d >= m.endRow) {
        a.ranges.push(m);
        const f = Hr({ id: W.MoveRowsCommandId, params: e }, m), C = Dt(f, m);
        C && u.ranges.push(C);
      }
    }), a.ranges.length === 0)
      return this._handleNull();
    const h = ye(this._injector, a), g = We(this._injector, u);
    return {
      preRedos: [{ id: oe.id, params: a }],
      redos: [{ id: ne.id, params: u }],
      preUndos: [{ id: oe.id, params: g }],
      undos: [{ id: ne.id, params: h }]
    };
  }
  _handleMoveColsCommand(e, t, o) {
    const s = xe(this._univerInstanceService, t);
    if (!s)
      return this._handleNull();
    const r = $e(s, o);
    if (!r)
      return this._handleNull();
    const i = [...r.getMergeData()], a = { unitId: t, subUnitId: o, ranges: [] }, u = { unitId: t, subUnitId: o, ranges: [] }, { fromRange: l } = e, { startColumn: c, endColumn: d } = l;
    if (i.forEach((m) => {
      if (c <= m.startColumn && d >= m.endColumn) {
        a.ranges.push(m);
        const f = Fr({ id: W.MoveColsCommandId, params: e }, m), C = Dt(f, m);
        C && u.ranges.push(C);
      }
    }), a.ranges.length === 0)
      return this._handleNull();
    const h = ye(this._injector, a), g = We(this._injector, u);
    return {
      preRedos: [{ id: oe.id, params: a }],
      redos: [{ id: ne.id, params: u }],
      preUndos: [{ id: oe.id, params: g }],
      undos: [{ id: ne.id, params: h }]
    };
  }
  _handleMoveRangeCommand(e, t, o) {
    const s = xe(this._univerInstanceService, t);
    if (!s)
      return this._handleNull();
    const r = $e(s, o);
    if (!r)
      return this._handleNull();
    const i = r.getMergeData(), a = i.filter((g) => N.intersects(g, e.fromRange)), u = i.filter((g) => N.intersects(g, e.toRange)), l = a.map((g) => N.getRelativeRange(g, e.fromRange)).map((g) => N.getPositionRange(g, e.toRange)), c = zr(l).filter(
      (g) => !i.some((m) => N.equals(g, m))
    ), d = [
      {
        id: oe.id,
        params: {
          unitId: t,
          subUnitId: o,
          ranges: a
        }
      },
      {
        id: oe.id,
        params: {
          unitId: t,
          subUnitId: o,
          ranges: u
        }
      },
      {
        id: ne.id,
        params: {
          unitId: t,
          subUnitId: o,
          ranges: c
        }
      }
    ], h = [
      {
        id: oe.id,
        params: {
          unitId: t,
          subUnitId: o,
          ranges: c
        }
      },
      {
        id: ne.id,
        params: {
          unitId: t,
          subUnitId: o,
          ranges: u
        }
      },
      {
        id: ne.id,
        params: {
          unitId: t,
          subUnitId: o,
          ranges: a
        }
      }
    ];
    return { redos: d, undos: h };
  }
  _handleInsertRowCommand(e, t, o) {
    const s = xe(this._univerInstanceService, t);
    if (!s)
      return this._handleNull();
    const r = $e(s, o);
    if (!r)
      return this._handleNull();
    const { range: i } = e, { startRow: a, endRow: u } = i, l = A.deepClone(r.getMergeData()).reduce((p, w) => (a > w.startRow && a <= w.endRow && p.push(w), p), []);
    if (l.length === 0)
      return this._handleNull();
    const c = A.deepClone(r.getMergeData()).reduce((p, w) => {
      if (a > w.startRow && a <= w.endRow) {
        const I = u - a + 1;
        w.endRow += I, this._checkIsMergeCell(w) && p.push(w);
      }
      return p;
    }, []), d = {
      unitId: t,
      subUnitId: o,
      ranges: l
    }, h = ye(
      this._injector,
      d
    ), g = {
      unitId: t,
      subUnitId: o,
      ranges: c
    }, m = We(
      this._injector,
      g
    ), f = [
      { id: oe.id, params: d },
      { id: ne.id, params: g }
    ], C = [
      { id: oe.id, params: m },
      { id: ne.id, params: h }
    ];
    return { redos: f, undos: C };
  }
  _handleInsertColCommand(e, t, o) {
    const { range: s } = e, r = xe(this._univerInstanceService, t);
    if (!r)
      return this._handleNull();
    const i = $e(r, o);
    if (!i)
      return this._handleNull();
    const { startColumn: a, endColumn: u } = s, l = A.deepClone(i.getMergeData()).reduce((p, w) => (a > w.startColumn && a <= w.endColumn && p.push(w), p), []);
    if (l.length === 0)
      return this._handleNull();
    const c = A.deepClone(i.getMergeData()).reduce((p, w) => {
      if (a > w.startColumn && a <= w.endColumn) {
        const I = u - a + 1;
        w.endColumn += I, this._checkIsMergeCell(w) && p.push(w);
      }
      return p;
    }, []), d = {
      unitId: t,
      subUnitId: o,
      ranges: l
    }, h = ye(
      this._injector,
      d
    ), g = {
      unitId: t,
      subUnitId: o,
      ranges: c
    }, m = We(
      this._injector,
      g
    ), f = [
      { id: oe.id, params: d },
      { id: ne.id, params: g }
    ], C = [
      { id: oe.id, params: m },
      { id: ne.id, params: h }
    ];
    return { redos: f, undos: C };
  }
  _handleRemoveColCommand(e, t, o) {
    const s = xe(this._univerInstanceService, t);
    if (!s)
      return this._handleNull();
    const r = $e(s, o);
    if (!r)
      return this._handleNull();
    const { range: i } = e, { startColumn: a, endColumn: u } = i, l = A.deepClone(r.getMergeData()).reduce((I, S) => (N.intersects(i, S) && I.push(S), I), []);
    if (l.length === 0)
      return this._handleNull();
    const c = A.deepClone(r.getMergeData()).reduce((I, S) => {
      if (N.intersects(i, S)) {
        if (a <= S.startColumn && u >= S.endColumn)
          return I;
        a >= S.startColumn && u <= S.endColumn ? S.endColumn -= u - a + 1 : a < S.startColumn ? (S.startColumn = a, S.endColumn -= u - a + 1) : u > S.endColumn && (S.endColumn = a - 1), this._checkIsMergeCell(S) && I.push(S);
      }
      return I;
    }, []), d = {
      unitId: t,
      subUnitId: o,
      ranges: l
    }, h = ye(
      this._injector,
      d
    ), g = {
      unitId: t,
      subUnitId: o,
      ranges: c
    }, m = We(
      this._injector,
      g
    ), f = [{ id: oe.id, params: d }], C = [{ id: ne.id, params: g }], p = [{ id: oe.id, params: m }], w = [{ id: ne.id, params: h }];
    return { preUndos: p, undos: w, preRedos: f, redos: C };
  }
  _handleRemoveRowCommand(e, t, o) {
    const { range: s } = e, r = xe(this._univerInstanceService, t);
    if (!r)
      return this._handleNull();
    const i = $e(r, o);
    if (!i)
      return this._handleNull();
    const { startRow: a, endRow: u } = s, l = A.deepClone(i.getMergeData()).reduce((I, S) => (N.intersects(s, S) && I.push(S), I), []);
    if (l.length === 0)
      return this._handleNull();
    const c = A.deepClone(i.getMergeData()).reduce((I, S) => {
      if (N.intersects(s, S)) {
        if (a <= S.startRow && u >= S.endRow)
          return I;
        a >= S.startRow && u <= S.endRow ? S.endRow -= u - a + 1 : a < S.startRow ? (S.startRow = a, S.endRow -= u - a + 1) : u > S.endRow && (S.endRow = a - 1), this._checkIsMergeCell(S) && I.push(S);
      }
      return I;
    }, []), d = {
      unitId: t,
      subUnitId: o,
      ranges: l
    }, h = ye(
      this._injector,
      d
    ), g = {
      unitId: t,
      subUnitId: o,
      ranges: c
    }, m = We(
      this._injector,
      g
    ), f = [{ id: oe.id, params: d }], C = [{ id: ne.id, params: g }], p = [{ id: oe.id, params: m }], w = [{ id: ne.id, params: h }];
    return { preUndos: p, undos: w, preRedos: f, redos: C };
  }
  _handleInsertRangeMoveRightCommand(e, t, o) {
    const s = xe(this._univerInstanceService, t);
    if (!s)
      return this._handleNull();
    const r = $e(s, o);
    if (!r)
      return this._handleNull();
    const i = e.range, a = r.getMaxColumns() - 1, u = r.getMergeData(), l = [], c = [];
    u.forEach((f) => {
      const { startRow: C, endRow: p, startColumn: w, endColumn: I } = i;
      if (N.intersects(
        {
          startRow: C,
          startColumn: w,
          endRow: p,
          endColumn: a
        },
        f
      ) && (l.push(f), N.contains(
        {
          startRow: C,
          startColumn: w,
          endRow: p,
          endColumn: a
        },
        f
      ))) {
        const b = I - w + 1;
        c.push({
          startRow: f.startRow,
          startColumn: f.startColumn + b,
          endRow: f.endRow,
          endColumn: f.endColumn + b
        });
      }
    });
    const d = {
      unitId: t,
      subUnitId: o,
      ranges: l
    }, h = ye(
      this._injector,
      d
    ), g = {
      unitId: t,
      subUnitId: o,
      ranges: c
    }, m = We(
      this._injector,
      g
    );
    return {
      preRedos: [
        { id: oe.id, params: d }
      ],
      redos: [
        {
          id: ne.id,
          params: g
        }
      ],
      preUndos: [
        { id: oe.id, params: m }
      ],
      undos: [
        {
          id: ne.id,
          params: h
        }
      ]
    };
  }
  _handleInsertRangeMoveDownCommand(e, t, o) {
    const s = xe(this._univerInstanceService, t);
    if (!s)
      return this._handleNull();
    const r = $e(s, o);
    if (!r)
      return this._handleNull();
    const i = e.range, a = r.getMaxRows() - 1, u = r.getMergeData(), l = [], c = [];
    u.forEach((I) => {
      const { startRow: S, startColumn: y, endColumn: b, endRow: O } = i;
      if (N.intersects({ startRow: S, startColumn: y, endRow: a, endColumn: b }, I) && (l.push(I), N.contains({ startRow: S, startColumn: y, endRow: a, endColumn: b }, I))) {
        const U = O - S + 1;
        c.push({
          startRow: I.startRow + U,
          startColumn: I.startColumn,
          endRow: I.endRow + U,
          endColumn: I.endColumn
        });
      }
    });
    const d = {
      unitId: t,
      subUnitId: o,
      ranges: l
    }, h = ye(
      this._injector,
      d
    ), g = {
      unitId: t,
      subUnitId: o,
      ranges: c
    }, m = We(
      this._injector,
      g
    ), f = [
      {
        id: oe.id,
        params: d
      }
    ], C = [
      {
        id: ne.id,
        params: g
      }
    ], p = [
      {
        id: oe.id,
        params: m
      }
    ], w = [
      {
        id: ne.id,
        params: h
      }
    ];
    return { redos: C, undos: w, preRedos: f, preUndos: p };
  }
  _handleDeleteRangeMoveUpCommand(e, t, o) {
    const s = xe(this._univerInstanceService, t);
    if (!s)
      return this._handleNull();
    const r = $e(s, o);
    if (!r)
      return this._handleNull();
    const i = e.range, a = r.getMaxRows() - 1, u = r.getMergeData(), l = [], c = [];
    u.forEach((I) => {
      const { startRow: S, startColumn: y, endColumn: b, endRow: O } = i;
      if (N.intersects({ startRow: S, startColumn: y, endRow: a, endColumn: b }, I) && (l.push(I), N.contains({ startRow: S, startColumn: y, endRow: a, endColumn: b }, I))) {
        const U = O - S + 1, D = N.moveVertical(I, -U);
        c.push(D);
      }
    });
    const d = {
      unitId: t,
      subUnitId: o,
      ranges: l
    }, h = ye(
      this._injector,
      d
    ), g = {
      unitId: t,
      subUnitId: o,
      ranges: c
    }, m = We(
      this._injector,
      g
    ), f = [
      {
        id: oe.id,
        params: d
      }
    ], C = [
      {
        id: ne.id,
        params: g
      }
    ], p = [
      {
        id: oe.id,
        params: m
      }
    ], w = [
      {
        id: ne.id,
        params: h
      }
    ];
    return { redos: C, undos: w, preRedos: f, preUndos: p };
  }
  _handleDeleteRangeMoveLeftCommand(e, t, o) {
    const s = xe(this._univerInstanceService, t);
    if (!s)
      return this._handleNull();
    const r = $e(s, o);
    if (!r)
      return this._handleNull();
    const i = e.range, a = r.getMaxColumns() - 1, u = r.getMergeData(), l = [], c = [];
    u.forEach((f) => {
      const { startRow: C, endRow: p, startColumn: w, endColumn: I } = i;
      if (N.intersects(
        {
          startRow: C,
          startColumn: w,
          endRow: p,
          endColumn: a
        },
        f
      ) && (l.push(f), N.contains(
        {
          startRow: C,
          startColumn: w,
          endRow: p,
          endColumn: a
        },
        f
      ))) {
        const b = I - w + 1;
        c.push({
          startRow: f.startRow,
          startColumn: f.startColumn - b,
          endRow: f.endRow,
          endColumn: f.endColumn - b
        });
      }
    });
    const d = {
      unitId: t,
      subUnitId: o,
      ranges: l
    }, h = ye(this._injector, d), g = {
      unitId: t,
      subUnitId: o,
      ranges: c
    }, m = We(this._injector, g);
    return {
      preRedos: [
        { id: oe.id, params: d }
      ],
      redos: [
        {
          id: ne.id,
          params: g
        }
      ],
      undos: [
        {
          id: ne.id,
          params: h
        }
      ],
      preUndos: [
        { id: oe.id, params: m }
      ]
    };
  }
  _checkIsMergeCell(e) {
    return !(e.startRow === e.endRow && e.startColumn === e.endColumn);
  }
  _handleNull() {
    return { redos: [], undos: [] };
  }
  _commandExecutedListener() {
    this.disposeWithMe(this._commandService.onCommandExecuted((e) => {
      if (Gl.includes(e.id)) {
        if (!e.params) return;
        const t = this._univerInstanceService.getUniverSheetInstance(e.params.unitId);
        if (!t) return;
        const o = t.getSheetBySheetId(e.params.subUnitId);
        if (!o) return;
        const { sourceRange: s, targetRange: r } = e.params, i = s.startColumn === r.startColumn && s.endColumn === r.endColumn, a = i ? s.endRow - s.startRow + 1 : s.endColumn - s.startColumn + 1, u = i ? s.startRow : s.startColumn, l = i ? r.startRow : r.startColumn, c = o.getConfig().mergeData, d = [];
        c.forEach((f) => {
          let { startRow: C, endRow: p, startColumn: w, endColumn: I, rangeType: S } = f;
          N.intersects(f, s) || (i ? u < C && l > p ? (C -= a, p -= a) : u > p && l <= C && (C += a, p += a) : u < w && l > I ? (w -= a, I -= a) : u > I && l <= w && (w += a, I += a)), f.startRow === f.endRow && f.startColumn === f.endColumn || d.push({ startRow: C, endRow: p, startColumn: w, endColumn: I, rangeType: S });
        }), o.setMergeData(d), this.disposableCollection.dispose();
        const { unitId: h, subUnitId: g } = e.params, m = (f) => this.refRangeHandle(f, h, g);
        d.forEach((f) => {
          this.disposableCollection.add(this._refRangeService.registerRefRange(f, m, h, g));
        });
      }
      if (zl.includes(e.id)) {
        const t = this._univerInstanceService.getUniverSheetInstance(e.params.unitId);
        if (!t) return;
        const o = t.getSheetBySheetId(e.params.subUnitId);
        if (!o) return;
        const s = o.getConfig().mergeData, r = e.params;
        if (!r) return;
        const { range: i } = r, a = e.id.includes("row"), u = e.id.includes("insert"), l = a ? i.startRow : i.startColumn, c = a ? i.endRow : i.endColumn, d = c - l + 1, h = [];
        s.forEach((C) => {
          let { startRow: p, endRow: w, startColumn: I, endColumn: S, rangeType: y } = C;
          u ? a ? l <= p && (p += d, w += d) : l <= I && (I += d, S += d) : a ? c < p && (p -= d, w -= d) : c < I && (I -= d, S -= d), C.startRow === C.endRow && C.startColumn === C.endColumn || h.push({ startRow: p, endRow: w, startColumn: I, endColumn: S, rangeType: y });
        }), o.setMergeData(h), this.disposableCollection.dispose();
        const { unitId: g, subUnitId: m } = e.params, f = (C) => this.refRangeHandle(C, g, m);
        h.forEach((C) => {
          this.disposableCollection.add(this._refRangeService.registerRefRange(C, f, g, m));
        });
      }
    }));
  }
};
Wn = jl([
  _t(0, $(E)),
  _t(1, $(At)),
  _t(2, $(M)),
  _t(3, $(so)),
  _t(4, $(G)),
  _t(5, $(z))
], Wn);
function xe(n, e) {
  return e ? n.getUniverSheetInstance(e) : n.getCurrentUnitForType(B.UNIVER_SHEET);
}
function $e(n, e) {
  return e ? n.getSheetBySheetId(e) : n.getActiveSheet();
}
function Jl(n, e) {
  return e.some((t) => ql(n, t));
}
function ql(n, e) {
  const { startRow: t, startColumn: o, endColumn: s, endRow: r } = e, i = n.getMatrixWithMergedCells(t, o, r, s);
  let a = !1;
  return i.forValue((u, l, c) => {
    if (c && (u !== t || l !== o) && n.cellHasValue(c))
      return a = !0, !1;
  }), a;
}
function Yl(n, e, t, o) {
  const s = [], r = [], i = t.getSheetId();
  return o.forEach((a) => {
    const u = Xl(t, a), l = {
      unitId: e,
      subUnitId: i,
      cellValue: u.getData()
    }, c = Ae(
      n,
      l
    );
    s.push({ id: ee.id, params: c }), r.push({ id: ee.id, params: l });
  }), {
    undos: s,
    redos: r
  };
}
function Xl(n, e) {
  const { startRow: t, startColumn: o, endColumn: s, endRow: r } = e, i = n.getMatrixWithMergedCells(t, o, r, s, Ei.Raw), a = new Y();
  return i.forValue((u, l, c) => {
    c && (u !== t || l !== o) && a.setValue(u, l, null);
  }), a;
}
const go = {
  type: v.COMMAND,
  id: "sheet.command.add-worksheet-merge",
  handler: (n, e) => {
    const t = n.get(E), o = n.get(V), s = n.get(M), r = e.unitId, i = e.subUnitId, a = e.selections, u = zr(a, e.value), l = s.getUniverSheetInstance(r).getSheetBySheetId(i), c = [], d = [], h = Jl(l, u), g = {
      unitId: r,
      subUnitId: i,
      ranges: u
    }, m = {
      unitId: r,
      subUnitId: i,
      ranges: u
    };
    c.push({ id: oe.id, params: g }), c.push({ id: ne.id, params: m });
    const f = ye(n, g), C = We(n, m);
    if (d.push({ id: oe.id, params: C }), d.push({ id: ne.id, params: f }), h) {
      const w = Yl(n, r, l, u);
      c.unshift(...w.redos), d.push(...w.undos);
    }
    return L(c, t).result ? (o.pushUndoRedo({
      unitID: r,
      undoMutations: d,
      redoMutations: c
    }), !0) : !1;
  }
}, Wh = {
  type: v.COMMAND,
  id: "sheet.command.add-worksheet-merge-all",
  handler: async (n) => {
    var l;
    const e = n.get(E), o = (l = n.get(z).getCurrentSelections()) == null ? void 0 : l.map((c) => c.range);
    if (!(o != null && o.length))
      return !1;
    const r = n.get(M).getCurrentUnitForType(B.UNIVER_SHEET);
    if (!r) return !1;
    const i = r.getActiveSheet();
    if (!i) return !1;
    const a = r.getUnitId(), u = i.getSheetId();
    return e.executeCommand(go.id, {
      selections: o,
      unitId: a,
      subUnitId: u
    });
  }
}, Vh = {
  type: v.COMMAND,
  id: "sheet.command.add-worksheet-merge-vertical",
  handler: async (n) => {
    var l;
    const e = n.get(E), o = (l = n.get(z).getCurrentSelections()) == null ? void 0 : l.map((c) => c.range);
    if (!(o != null && o.length))
      return !1;
    const r = n.get(M).getCurrentUnitForType(B.UNIVER_SHEET);
    if (!r) return !1;
    const i = r.getActiveSheet();
    if (!i) return !1;
    const a = r.getUnitId(), u = i.getSheetId();
    return e.executeCommand(go.id, {
      value: Re.COLUMNS,
      selections: o,
      unitId: a,
      subUnitId: u
    });
  }
}, Lh = {
  type: v.COMMAND,
  id: "sheet.command.add-worksheet-merge-horizontal",
  handler: async (n) => {
    var l;
    const e = n.get(E), o = (l = n.get(z).getCurrentSelections()) == null ? void 0 : l.map((c) => c.range);
    if (!(o != null && o.length))
      return !1;
    const r = n.get(M).getCurrentUnitForType(B.UNIVER_SHEET);
    if (!r) return !1;
    const i = r.getActiveSheet();
    if (!i) return !1;
    const a = r.getUnitId(), u = i.getSheetId();
    return e.executeCommand(go.id, {
      value: Re.ROWS,
      selections: o,
      unitId: a,
      subUnitId: u
    });
  }
};
function Hh(n, e, t, o, s) {
  const r = n.get(M), i = P(r, { unitId: e, subUnitId: t });
  if (!i) return;
  const { worksheet: a } = i;
  if (a.getMergeData().some((d) => o.some((h) => N.intersects(h, d))))
    throw new Error("The ranges to be merged overlap with the existing merged cells");
  n.get(E).executeCommand(go.id, {
    unitId: e,
    subUnitId: t,
    selections: o,
    defaultMerge: s
  });
}
class lt {
  constructor() {
    /**
     *
     * Map<unitId, Map<subUnitId, Map<subUnitId, IWorksheetProtectionRule>>>
     */
    R(this, "_model", /* @__PURE__ */ new Map());
    R(this, "_ruleChange", new De());
    R(this, "_ruleRefresh", new De());
    R(this, "_resetOrder", new De());
    R(this, "ruleChange$", this._ruleChange.asObservable());
    R(this, "ruleRefresh$", this._ruleRefresh.asObservable());
    R(this, "resetOrder$", this._resetOrder.asObservable());
    R(this, "_worksheetRuleInitStateChange", new gt(!1));
    R(this, "worksheetRuleInitStateChange$", this._worksheetRuleInitStateChange.asObservable());
  }
  changeRuleInitState(e) {
    this._worksheetRuleInitStateChange.next(e);
  }
  getSheetRuleInitState() {
    return this._worksheetRuleInitStateChange.value;
  }
  addRule(e, t) {
    this._ensureSubUnitMap(e).set(t.subUnitId, t), this._ruleChange.next({ unitId: e, rule: t, type: "add", subUnitId: t.subUnitId });
  }
  deleteRule(e, t) {
    var s, r, i;
    const o = (r = (s = this._model) == null ? void 0 : s.get(e)) == null ? void 0 : r.get(t);
    o && ((i = this._model.get(e)) == null || i.delete(t), this._ruleChange.next({ unitId: e, rule: o, type: "delete", subUnitId: t }));
  }
  setRule(e, t, o) {
    var r, i;
    const s = this.getRule(e, t);
    s && ((i = (r = this._model) == null ? void 0 : r.get(e)) == null || i.set(t, o), this._ruleChange.next({ unitId: e, oldRule: s, rule: o, type: "set", subUnitId: t }));
  }
  getRule(e, t) {
    var o, s;
    return (s = (o = this._model) == null ? void 0 : o.get(e)) == null ? void 0 : s.get(t);
  }
  toObject() {
    const e = {};
    return [...this._model.keys()].forEach((o) => {
      const s = this._model.get(o);
      s != null && s.size && (e[o] = [], [...s.keys()].forEach((i) => {
        const a = s.get(i);
        a && e[o].push(a);
      }));
    }), e;
  }
  fromObject(e) {
    const t = /* @__PURE__ */ new Map();
    Object.keys(e).forEach((o) => {
      const s = e[o];
      if (s != null && s.length) {
        const r = /* @__PURE__ */ new Map();
        s.forEach((i) => {
          r.set(i.subUnitId, i);
        }), t.set(o, r);
      }
    }), this._model = t;
  }
  deleteUnitModel(e) {
    this._model.delete(e);
  }
  _ensureSubUnitMap(e) {
    let t = this._model.get(e);
    return t || (t = /* @__PURE__ */ new Map(), this._model.set(e, t)), t;
  }
  ruleRefresh(e) {
    this._ruleRefresh.next(e);
  }
  resetOrder() {
    this._resetOrder.next(Math.random());
  }
  getTargetByPermissionId(e, t) {
    const o = this._model.get(e);
    if (!o) return null;
    for (const [s, r] of o)
      if (r.permissionId === t)
        return [e, s];
  }
}
const Mt = {
  id: "sheet.mutation.add-worksheet-protection",
  type: v.MUTATION,
  handler: (n, e) => {
    const { unitId: t, rule: o } = e;
    return n.get(lt).addRule(t, o), !0;
  }
}, xt = {
  id: "sheet.mutation.delete-worksheet-protection",
  type: v.MUTATION,
  handler: (n, e) => {
    const { unitId: t, subUnitId: o } = e;
    return n.get(lt).deleteRule(t, o), !0;
  }
}, Zl = {
  type: v.COMMAND,
  id: "sheet.command.add-worksheet-protection",
  async handler(n, e) {
    if (!e)
      return !1;
    const t = n.get(E), o = n.get(V), { rule: s, unitId: r } = e, i = s.subUnitId;
    if (await t.executeCommand(Mt.id, {
      unitId: r,
      rule: s,
      subUnitId: s.subUnitId
    })) {
      const u = [{ id: Mt.id, params: { unitId: r, rule: s, subUnitId: s.subUnitId } }], l = [{ id: xt.id, params: { unitId: r, subUnitId: i } }];
      o.pushUndoRedo({
        unitID: r,
        redoMutations: u,
        undoMutations: l
      });
    }
    return !0;
  }
}, Ql = {
  type: v.COMMAND,
  id: "sheet.command.set-worksheet-range-theme-style",
  handler: (n, e) => {
    const t = n.get(E), o = n.get(V), { unitId: s } = e, r = Oa(n, e);
    return t.syncExecuteCommand(tn.id, e) ? (o.pushUndoRedo({
      unitID: s,
      undoMutations: [{ id: nn.id, params: r }],
      redoMutations: [{ id: tn.id, params: e }]
    }), !0) : !1;
  }
}, ec = "sheet.command.append-row", tc = {
  type: v.COMMAND,
  id: ec,
  handler: (n, e) => {
    const t = n.get(E), o = n.get(V), { unitId: s, subUnitId: r, cellValue: i, insertRowNums: a, insertColumnNums: u, maxRows: l, maxColumns: c } = e, d = {
      unitId: s,
      subUnitId: r,
      cellValue: i
    }, h = Ae(
      n,
      d
    ), g = [{ id: ee.id, params: d }], m = [{ id: ee.id, params: h }];
    if (a) {
      const C = {
        unitId: s,
        subUnitId: r,
        range: {
          startRow: l,
          endRow: l,
          startColumn: 0,
          endColumn: c - 1
        }
      }, p = $o(
        n,
        C
      );
      g.unshift({ id: Ue.id, params: C }), m.push({ id: Te.id, params: p });
    }
    if (u) {
      const C = {
        unitId: s,
        subUnitId: r,
        range: {
          startRow: 0,
          endRow: l - 1,
          startColumn: c,
          endColumn: c - 1 + u
        }
      }, p = ro(
        n,
        C
      );
      g.unshift({ id: ke.id, params: C }), m.push({ id: Me.id, params: p });
    }
    return L(g, t).result ? (o.pushUndoRedo({
      unitID: s,
      undoMutations: m,
      redoMutations: g
    }), !0) : !1;
  }
}, qo = {
  id: "sheet.command.clear-selection-content",
  type: v.COMMAND,
  handler: (n, e) => {
    var I;
    const t = n.get(M), o = n.get(E), s = n.get(z), r = n.get(V), i = n.get(G), a = t.getCurrentUnitForType(B.UNIVER_SHEET);
    if (!a) return !1;
    const u = (e == null ? void 0 : e.unitId) || a.getUnitId(), l = a.getActiveSheet();
    if (!l) return !1;
    const c = (e == null ? void 0 : e.subUnitId) || l.getSheetId(), d = (e == null ? void 0 : e.ranges) || ((I = s.getCurrentSelections()) == null ? void 0 : I.map((S) => S.range));
    if (!(d != null && d.length))
      return !1;
    const h = ao(d, n, u, c), g = {
      subUnitId: c,
      unitId: u,
      cellValue: Eu(h)
    }, m = Ae(
      n,
      g
    ), f = i.onCommandExecute({ id: qo.id }), C = [{ id: ee.id, params: g }, ...f.redos], p = [...f.undos, { id: ee.id, params: m }];
    return L(C, o).result ? (r.pushUndoRedo({
      // If there are multiple mutations that form an encapsulated project, they must be encapsulated in the same undo redo element.
      // Hooks can be used to hook the code of external controllers to add new actions.
      unitID: u,
      undoMutations: p,
      redoMutations: C
    }), !0) : !1;
  }
}, Gr = (n, e) => ({
  subUnitId: e.sheet.id,
  unitId: e.unitId,
  subUnitName: e.sheet.name
}), ln = {
  id: "sheet.mutation.insert-sheet",
  type: v.MUTATION,
  handler: (n, e) => {
    const t = n.get(M), { sheet: o, index: s, unitId: r, styles: i } = e, a = t.getUniverSheetInstance(r);
    return a ? (i && a.addStyles(i), a.addWorksheet(o.id, s, o)) : !1;
  }
}, Kr = {
  type: v.COMMAND,
  id: "sheet.command.copy-sheet",
  handler: (n, e) => {
    var S, y;
    const t = n.get(E), o = n.get(V), s = n.get(M), r = n.get(G), i = n.get(it), a = P(s, e);
    if (!a)
      return !1;
    const { workbook: u, worksheet: l, unitId: c, subUnitId: d } = a, h = A.deepClone(l.getConfig());
    h.name = nc(u, i, h.name), h.id = Ut();
    const m = {
      index: u.getSheetIndex(l) + 1,
      sheet: h,
      unitId: c
    }, f = Gr(
      n,
      m
    ), C = r.onCommandExecute({
      id: Kr.id,
      params: { unitId: c, subUnitId: d, targetSubUnitId: h.id }
    }), p = [
      ...(S = C.preRedos) != null ? S : [],
      { id: ln.id, params: m },
      ...C.redos
    ], w = [
      ...(y = C.preUndos) != null ? y : [],
      { id: Lt.id, params: f },
      ...C.undos
    ];
    return L(p, t).result ? (o.pushUndoRedo({
      unitID: c,
      undoMutations: w,
      redoMutations: p
    }), !0) : !1;
  }
};
function nc(n, e, t) {
  let o = `${t} ${e.t("sheets.tabs.sheetCopy", "")}`, s = 2;
  for (; n.checkSheetName(o); )
    o = `${t} ${e.t("sheets.tabs.sheetCopy", `${s}`)}`, s++;
  return o;
}
const oc = {
  type: v.COMMAND,
  id: "sheet.command.delete-range-protection",
  async handler(n, e) {
    if (!e)
      return !1;
    const t = n.get(E), o = n.get(V), { unitId: s, subUnitId: r, rule: i } = e, a = {
      unitId: s,
      subUnitId: r,
      ruleIds: [i.id]
    };
    return await t.executeCommand(st.id, a) && o.pushUndoRedo({
      unitID: s,
      redoMutations: [{ id: st.id, params: a }],
      undoMutations: [{ id: Be.id, params: { unitId: s, subUnitId: r, rules: [i] } }]
    }), !0;
  }
}, sc = {
  type: v.COMMAND,
  id: "sheet.command.delete-worksheet-protection",
  handler(n, e) {
    if (!e)
      return !1;
    const t = n.get(E), o = n.get(V), { rule: s, unitId: r, subUnitId: i } = e;
    t.executeCommand(xt.id, {
      unitId: r,
      subUnitId: i
    });
    const a = [{ id: xt.id, params: { unitId: r, subUnitId: i } }], u = [{ id: Mt.id, params: { unitId: r, rule: s, subUnitId: i } }];
    return o.pushUndoRedo({
      unitID: r,
      redoMutations: a,
      undoMutations: u
    }), !0;
  }
}, rc = {
  type: v.COMMAND,
  id: "sheet.command.remove-worksheet-range-theme-style",
  handler: (n, e) => {
    const t = n.get(E), o = n.get(V), { unitId: s } = e, r = Da(n, e);
    return t.syncExecuteCommand(nn.id, e) ? (o.pushUndoRedo({
      unitID: s,
      undoMutations: [{ id: tn.id, params: r }],
      redoMutations: [{ id: nn.id, params: e }]
    }), !0) : !1;
  }
}, ic = {
  id: "sheet.command.insert-defined-name",
  type: v.COMMAND,
  handler: (n, e) => {
    const t = n.get(E), o = n.get(V);
    if (!e) return !1;
    const s = {
      ...e
    };
    return t.syncExecuteCommand(en.id, s) ? (o.pushUndoRedo({
      unitID: e.unitId,
      undoMutations: [{ id: Pn.id, params: s }],
      redoMutations: [{ id: en.id, params: s }]
    }), !0) : !1;
  }
}, ac = {
  id: "sheet.command.insert-sheet",
  type: v.COMMAND,
  handler: (n, e) => {
    var C;
    const t = n.get(E), o = n.get(V), s = n.get(M), r = n.get(it), i = Ks(s, { unitId: e == null ? void 0 : e.unitId });
    if (!i) return !1;
    const { unitId: a, workbook: u } = i;
    let l = u.getSheets().length;
    const c = e == null ? void 0 : e.sheet, d = c == null ? void 0 : c.id, h = Ui(c || {});
    e ? (l = (C = e.index) != null ? C : l, h.id = d || Ut(), h.name = (c == null ? void 0 : c.name) || u.generateNewSheetName(`${r.t("sheets.tabs.sheet")}`)) : (h.id = Ut(), h.name = u.generateNewSheetName(`${r.t("sheets.tabs.sheet")}`));
    const g = {
      index: l,
      sheet: h,
      unitId: a
    }, m = Gr(
      n,
      g
    );
    return t.syncExecuteCommand(ln.id, g) ? (o.pushUndoRedo({
      unitID: a,
      undoMutations: [{ id: Lt.id, params: m }],
      redoMutations: [{ id: ln.id, params: g }]
    }), !0) : !1;
  }
}, cn = {
  id: "sheet.mutation.register-worksheet-range-theme-style",
  type: v.MUTATION,
  handler: (n, e) => {
    const { unitId: t, rangeThemeStyleJson: o, themeName: s } = e, r = n.get(M), i = P(r), a = n.get(ve);
    if (!i) return !1;
    const u = new Rt(s, o);
    return a.registerRangeThemeStyle(t, u), !0;
  }
}, Yo = {
  id: "sheet.mutation.unregister-worksheet-range-theme-style",
  type: v.MUTATION,
  handler: (n, e) => {
    const { unitId: t, themeName: o } = e, s = n.get(M), r = P(s), i = n.get(ve);
    return r ? (i.unregisterRangeThemeStyle(t, o), !0) : !1;
  }
}, uc = {
  id: "sheet.command.register-worksheet-range-theme-style",
  type: v.COMMAND,
  handler: (n, e) => {
    if (!e)
      return !1;
    const { unitId: t, rangeThemeStyle: o } = e, s = n.get(M), r = n.get(E), i = n.get(V);
    if (!P(s)) return !1;
    const u = {
      unitId: t,
      themeName: o.getName(),
      rangeThemeStyleJson: o.toJson()
    }, l = {
      unitId: t,
      themeName: o.getName()
    };
    return r.syncExecuteCommand(cn.id, u) && i.pushUndoRedo({
      unitID: t,
      undoMutations: [{ id: Yo.id, params: l }],
      redoMutations: [{ id: cn.id, params: u }]
    }), !0;
  }
}, Jr = {
  id: "sheet.command.remove-defined-name",
  type: v.COMMAND,
  handler: (n, e) => {
    var c, d;
    const t = n.get(E), o = n.get(V), s = n.get(G);
    if (!e) return !1;
    const r = {
      ...e
    }, i = s.onCommandExecute({ id: Jr.id, params: e }), a = [
      ...(c = i.preRedos) != null ? c : [],
      { id: Pn.id, params: r },
      ...i.redos
    ], u = [
      ...(d = i.preUndos) != null ? d : [],
      { id: en.id, params: r },
      ...i.undos
    ];
    return L(a, t) ? (o.pushUndoRedo({
      unitID: e.unitId,
      undoMutations: u.filter(Boolean),
      redoMutations: a.filter(Boolean)
    }), !0) : !1;
  }
}, Xo = {
  id: "sheet.command.remove-sheet",
  type: v.COMMAND,
  handler: (n, e) => {
    var p, w;
    const t = n.get(E), o = n.get(V), s = n.get(M), r = n.get(G), i = P(s, e);
    if (!i) return !1;
    const { unitId: a, subUnitId: u, workbook: l, worksheet: c } = i;
    if (l.getSheets().length <= 1) return !1;
    const d = {
      subUnitId: u,
      unitId: a,
      subUnitName: c.getName()
    }, h = al(
      n,
      d
    ), g = r.onCommandExecute({
      id: Xo.id,
      params: { unitId: a, subUnitId: u }
    }), m = [...(p = g.preRedos) != null ? p : [], { id: Lt.id, params: d }, ...g.redos], f = [...(w = g.preUndos) != null ? w : [], { id: ln.id, params: h }, ...g.undos];
    return L(m, t).result ? (o.pushUndoRedo({
      unitID: a,
      undoMutations: f,
      redoMutations: m
    }), !0) : !1;
  }
}, lc = {
  type: v.COMMAND,
  id: "sheet.command.remove-worksheet-merge",
  // eslint-disable-next-line max-lines-per-function
  handler: (n, e) => {
    var D;
    const t = n.get(z), o = n.get(E), s = n.get(V), r = n.get(M), i = (e == null ? void 0 : e.ranges) || ((D = t.getCurrentSelections()) == null ? void 0 : D.map((H) => H.range));
    if (!(i != null && i.length)) return !1;
    const a = P(r);
    if (!a) return !1;
    const { subUnitId: u, unitId: l, worksheet: c } = a, d = {
      unitId: l,
      subUnitId: u,
      ranges: i
    }, g = c.getConfig().mergeData.filter((H) => i.some((F) => N.intersects(F, H)));
    if (!g.length) return !1;
    const m = ye(
      n,
      d
    ), f = t.getCurrentSelections();
    if (!(f != null && f.length)) return !1;
    const C = A.deepClone(f), p = A.deepClone(f), w = p[p.length - 1], { startRow: I, startColumn: S } = w.range;
    w.primary = {
      startRow: I,
      startColumn: S,
      endRow: I,
      endColumn: S,
      actualRow: I,
      actualColumn: S,
      isMerged: !1,
      isMergedMainCell: !1
    };
    const y = cc(c, g), b = {
      unitId: l,
      subUnitId: u,
      cellValue: y.redoParams.getMatrix()
    }, O = {
      unitId: l,
      subUnitId: u,
      cellValue: y.undoParams.getMatrix()
    }, k = [
      { id: oe.id, params: m },
      { id: ee.id, params: b },
      { id: ie.id, params: { selections: p } }
    ], T = [
      { id: ne.id, params: m },
      { id: ee.id, params: O },
      { id: ie.id, params: { selections: C } }
    ];
    return L(k, o) ? (s.pushUndoRedo({
      unitID: l,
      undoMutations: T,
      redoMutations: k
    }), !0) : !1;
  }
};
function cc(n, e) {
  const t = new Y(), o = new Y();
  return e.forEach((s) => {
    const { startRow: r, startColumn: i, endColumn: a, endRow: u } = s, l = n.getCellMatrix().getValue(r, i);
    if (l != null && l.s)
      for (let c = r; c <= u; c++)
        for (let d = i; d <= a; d++)
          (c !== r || d !== i) && (t.setValue(c, d, { s: l.s }), o.setValue(c, d, null));
  }), {
    redoParams: t,
    undoParams: o
  };
}
class Ht {
  constructor() {
    R(this, "_borderInfo", {
      type: de.ALL,
      color: "#000000",
      style: Ns.THIN,
      activeBorderType: !1
    });
    R(this, "_borderInfo$", new gt(this._borderInfo));
    R(this, "borderInfo$", this._borderInfo$.asObservable());
  }
  dispose() {
    this._borderInfo$.complete();
  }
  setType(e) {
    this._borderInfo.type = e, this.setActiveBorderType(!0), this._refresh();
  }
  setColor(e) {
    this._borderInfo.color = e, this._refresh();
  }
  setStyle(e) {
    this._borderInfo.style = e, this._refresh();
  }
  setActiveBorderType(e) {
    this._borderInfo.activeBorderType = e;
  }
  getBorderInfo() {
    return this._borderInfo;
  }
  _refresh() {
    this._borderInfo$.next(this._borderInfo);
  }
}
function Vn(n, e) {
  const { startRow: t, startColumn: o, endRow: s, endColumn: r } = n;
  for (let i = t; i <= s; i++)
    for (let a = o; a <= r; a++)
      e(i, a);
}
const Zo = (n, e, t, o) => {
  const { mr: s, worksheet: r } = n;
  e.startRow < 0 || e.startColumn < 0 || Vn(e, (i, a) => {
    var c, d;
    const u = r.getMergedCell(i, a);
    let l = t;
    if (u && (t.bc_tr || t.ml_tr || t.bl_tr || t.tl_mr || t.tl_bc || t.tl_br)) {
      if (o) {
        const h = A.deepClone(
          (c = s.getValue(u.startRow, u.startColumn)) == null ? void 0 : c.s
        );
        l = h != null && h.bd ? Object.assign(h.bd, t) : t;
      }
      s.setValue(u.startRow, u.startColumn, {
        s: {
          bd: l
        }
      });
    } else {
      if (o) {
        const h = A.deepClone((d = s.getValue(i, a)) == null ? void 0 : d.s);
        l = h != null && h.bd ? Object.assign(h.bd, t) : t;
      }
      s.setValue(i, a, { s: { bd: l } });
    }
  });
}, dc = (n) => {
  const e = {
    startRow: n.startRow - 1,
    startColumn: n.startColumn,
    endRow: n.startRow - 1,
    endColumn: n.endColumn
  }, t = {
    startRow: n.startRow,
    startColumn: n.startColumn - 1,
    endRow: n.endRow,
    endColumn: n.startColumn - 1
  }, o = {
    startRow: n.endRow + 1,
    startColumn: n.startColumn,
    endRow: n.endRow + 1,
    endColumn: n.endColumn
  }, s = {
    startRow: n.startRow,
    startColumn: n.endColumn + 1,
    endRow: n.endRow,
    endColumn: n.endColumn + 1
  }, r = {
    startRow: n.startRow,
    startColumn: n.startColumn,
    endRow: n.startRow,
    endColumn: n.endColumn
  }, i = {
    startRow: n.startRow,
    startColumn: n.startColumn,
    endRow: n.endRow,
    endColumn: n.startColumn
  }, a = {
    startRow: n.endRow,
    startColumn: n.startColumn,
    endRow: n.endRow,
    endColumn: n.endColumn
  }, u = {
    startRow: n.startRow,
    startColumn: n.endColumn,
    endRow: n.endRow,
    endColumn: n.endColumn
  };
  return {
    topRangeOut: e,
    leftRangeOut: t,
    bottomRangeOut: o,
    rightRangeOut: s,
    topRange: r,
    leftRange: i,
    bottomRange: a,
    rightRange: u
  };
};
function hc(n, e, t) {
  const { style: o, color: s, type: r } = n.getBorderInfo(), i = r === de.TOP || r === de.ALL || r === de.OUTSIDE, a = r === de.LEFT || r === de.ALL || r === de.OUTSIDE, u = r === de.BOTTOM || r === de.ALL || r === de.OUTSIDE, l = r === de.RIGHT || r === de.ALL || r === de.OUTSIDE, c = r === de.VERTICAL || r === de.ALL || r === de.INSIDE, d = r === de.HORIZONTAL || r === de.ALL || r === de.INSIDE, h = r.indexOf("tlbr") > -1, g = r.indexOf("tlbc") > -1, m = r.indexOf("tlmr") > -1, f = r.indexOf("bltr") > -1, C = r.indexOf("mltr") > -1, p = r.indexOf("bctr") > -1, w = t[0], {
    topRangeOut: I,
    leftRangeOut: S,
    bottomRangeOut: y,
    rightRangeOut: b,
    topRange: O,
    leftRange: k,
    bottomRange: T,
    rightRange: U
  } = dc(w), D = new Y(), { worksheet: H, unitId: F, subUnitId: K } = e;
  return {
    worksheet: H,
    unitId: F,
    subUnitId: K,
    style: o,
    color: s,
    type: r,
    top: i,
    left: a,
    right: l,
    bottom: u,
    vertical: c,
    horizontal: d,
    tl_br: h,
    tl_bc: g,
    tl_mr: m,
    bl_tr: f,
    ml_tr: C,
    bc_tr: p,
    topRangeOut: I,
    leftRangeOut: S,
    bottomRangeOut: y,
    rightRangeOut: b,
    topRange: O,
    leftRange: k,
    bottomRange: T,
    rightRange: U,
    range: w,
    mr: D,
    borderStyle: {
      s: o,
      cl: {
        rgb: s
      }
    }
  };
}
const gc = (n) => {
  const { range: e, mr: t, borderStyle: o, vertical: s, horizontal: r, worksheet: i } = n;
  s && Vn(e, (a, u) => {
    var c, d, h;
    const l = i.getMergedCell(a, u);
    if (l) {
      const g = (c = t.getValue(l.startRow, l.startColumn)) == null ? void 0 : c.s;
      l.startColumn !== e.startColumn && t.setValue(a, u, {
        s: {
          bd: g != null && g.bd ? Object.assign(g.bd, { l: A.deepClone(o) }) : { l: A.deepClone(o) }
        }
      });
    } else {
      if (u !== e.endColumn) {
        const g = (d = t.getValue(a, u)) == null ? void 0 : d.s;
        t.setValue(a, u, {
          s: {
            bd: g != null && g.bd ? Object.assign(g.bd, { r: A.deepClone(o) }) : { r: A.deepClone(o) }
          }
        });
      }
      if (u !== e.startColumn) {
        const g = (h = t.getValue(a, u)) == null ? void 0 : h.s;
        t.setValue(a, u, {
          s: {
            bd: g != null && g.bd ? Object.assign(g.bd, { l: A.deepClone(o) }) : { l: A.deepClone(o) }
          }
        });
      }
    }
  }), r && Vn(e, (a, u) => {
    var c, d, h;
    const l = i.getMergedCell(a, u);
    if (l) {
      const g = (c = t.getValue(l.startRow, l.startColumn)) == null ? void 0 : c.s;
      l.startRow !== e.startRow && t.setValue(a, u, {
        s: {
          bd: g != null && g.bd ? Object.assign(g.bd, { t: A.deepClone(o) }) : { t: A.deepClone(o) }
        }
      });
    } else {
      if (a !== e.endRow) {
        const g = (d = t.getValue(a, u)) == null ? void 0 : d.s;
        t.setValue(a, u, {
          s: {
            bd: g != null && g.bd ? Object.assign(g.bd, { b: A.deepClone(o) }) : { b: A.deepClone(o) }
          }
        });
      }
      if (a !== e.startRow) {
        const g = (h = t.getValue(a, u)) == null ? void 0 : h.s;
        t.setValue(a, u, {
          s: {
            bd: g != null && g.bd ? Object.assign(g.bd, { t: A.deepClone(o) }) : { t: A.deepClone(o) }
          }
        });
      }
    }
  });
};
function mc(n) {
  const { borderStyle: e, tl_br: t, tl_bc: o, tl_mr: s, bl_tr: r, ml_tr: i, bc_tr: a } = n, u = (l, c, d) => {
    Zo(n, l, c, d);
  };
  t && u(n.range, { tl_br: A.deepClone(e) }, !0), o && u(n.range, { tl_bc: A.deepClone(e) }, !0), s && u(n.range, { tl_mr: A.deepClone(e) }, !0), r && u(n.range, { bl_tr: A.deepClone(e) }, !0), i && u(n.range, { ml_tr: A.deepClone(e) }, !0), a && u(n.range, { bc_tr: A.deepClone(e) }, !0);
}
const fc = (n) => {
  const { top: e, left: t, right: o, bottom: s, borderStyle: r, bottomRange: i, topRange: a, leftRange: u, rightRange: l, bottomRangeOut: c, topRangeOut: d, leftRangeOut: h, rightRangeOut: g } = n, m = (f, C, p) => {
    Zo(n, f, C, p);
  };
  e && (m(d, { b: null }), m(a, { t: A.deepClone(r) }, !0)), s && (m(c, { t: null }), m(i, { b: A.deepClone(r) }, !0)), t && (m(h, { r: null }), m(u, { l: A.deepClone(r) }, !0)), o && (m(g, { l: null }), m(l, { r: A.deepClone(r) }, !0));
}, Rc = (n) => {
  const { range: e, worksheet: t, mr: o, top: s, bottom: r, left: i, right: a, vertical: u, horizontal: l, tl_br: c, tl_bc: d, tl_mr: h, bl_tr: g, ml_tr: m, bc_tr: f, topRange: C, bottomRange: p, leftRange: w, rightRange: I, topRangeOut: S, bottomRangeOut: y, leftRangeOut: b, rightRangeOut: O } = n, k = (T, U, D) => {
    Zo(n, T, U, D);
  };
  !s && !r && !i && !a && !u && !l && !c && !d && !h && !g && !m && !f && (Vn(e, (T, U) => {
    var H, F, K, J, ae, se, le, Se;
    const D = t.getMergedCell(T, U);
    if (D) {
      if (D.endColumn !== e.endColumn) {
        const Z = (H = o.getValue(D.startRow, D.startColumn)) == null ? void 0 : H.s;
        o.setValue(T, U, {
          s: {
            bd: Z != null && Z.bd ? Object.assign(Z.bd, { r: null }) : { r: null }
          }
        });
      }
      if (D.startColumn !== e.startColumn) {
        const Z = (F = o.getValue(D.startRow, D.startColumn)) == null ? void 0 : F.s;
        o.setValue(T, U, {
          s: {
            bd: Z != null && Z.bd ? Object.assign(Z.bd, { l: null }) : { l: null }
          }
        });
      }
      if (D.endRow !== e.endRow) {
        const Z = (K = o.getValue(D.startRow, D.startColumn)) == null ? void 0 : K.s;
        o.setValue(T, U, {
          s: {
            bd: Z != null && Z.bd ? Object.assign(Z.bd, { b: null }) : { b: null }
          }
        });
      }
      if (D.startRow !== e.startRow) {
        const Z = (J = o.getValue(D.startRow, D.startColumn)) == null ? void 0 : J.s;
        o.setValue(T, U, {
          s: {
            bd: Z != null && Z.bd ? Object.assign(Z.bd, { t: null }) : { t: null }
          }
        });
      }
    } else {
      if (U !== e.endColumn) {
        const Z = (ae = o.getValue(T, U)) == null ? void 0 : ae.s;
        o.setValue(T, U, {
          s: {
            bd: Z != null && Z.bd ? Object.assign(Z.bd, { r: null }) : { r: null }
          }
        });
      }
      if (U !== e.startColumn) {
        const Z = (se = o.getValue(T, U)) == null ? void 0 : se.s;
        o.setValue(T, U, {
          s: {
            bd: Z != null && Z.bd ? Object.assign(Z.bd, { l: null }) : { l: null }
          }
        });
      }
      if (T !== e.endRow) {
        const Z = (le = o.getValue(T, U)) == null ? void 0 : le.s;
        o.setValue(T, U, {
          s: {
            bd: Z != null && Z.bd ? Object.assign(Z.bd, { b: null }) : { b: null }
          }
        });
      }
      if (T !== e.startRow) {
        const Z = (Se = o.getValue(T, U)) == null ? void 0 : Se.s;
        o.setValue(T, U, {
          s: {
            bd: Z != null && Z.bd ? Object.assign(Z.bd, { t: null }) : { t: null }
          }
        });
      }
    }
  }), k(S, { b: null }), k(C, { t: null }, !0), k(y, { t: null }), k(p, { b: null }, !0), k(b, { r: null }), k(w, { l: null }, !0), k(O, { l: null }), k(I, { r: null }, !0), k(e, { tl_br: null }, !0), k(e, { tl_bc: null }, !0), k(e, { tl_mr: null }, !0), k(e, { bl_tr: null }, !0), k(e, { ml_tr: null }, !0), k(e, { bc_tr: null }, !0));
}, wn = {
  id: "sheet.command.set-border",
  type: v.COMMAND,
  handler: (n, e) => {
    var p;
    const t = n.get(E), o = n.get(V), s = n.get(M), r = n.get(z), i = n.get(Ht), a = P(s, e);
    if (!a) return !1;
    const u = (e == null ? void 0 : e.ranges) || ((p = r.getCurrentSelections()) == null ? void 0 : p.map((w) => w.range));
    if (!(u != null && u.length))
      return !1;
    const { activeBorderType: l } = i.getBorderInfo();
    if (!l) return !1;
    const c = hc(i, a, u);
    gc(c), fc(c), mc(c), Rc(c);
    const { unitId: d, subUnitId: h, mr: g } = c, m = {
      unitId: d,
      subUnitId: h,
      cellValue: g.getData()
    }, f = Ae(
      n,
      m
    );
    return t.syncExecuteCommand(ee.id, m) ? (o.pushUndoRedo({
      unitID: d,
      undoMutations: [{ id: ee.id, params: f }],
      redoMutations: [{ id: ee.id, params: m }]
    }), !0) : !1;
  }
}, Cc = {
  id: "sheet.command.set-border-position",
  type: v.COMMAND,
  handler: (n, e) => {
    if (!e.value) return !1;
    const t = n.get(E);
    return n.get(Ht).setType(e.value), t.syncExecuteCommand(wn.id);
  }
}, pc = {
  id: "sheet.command.set-border-style",
  type: v.COMMAND,
  handler: (n, e) => {
    const t = n.get(E);
    return n.get(Ht).setStyle(e.value), t.syncExecuteCommand(wn.id);
  }
}, Sc = {
  id: "sheet.command.set-border-color",
  type: v.COMMAND,
  handler: (n, e) => {
    const t = n.get(E);
    return n.get(Ht).setColor(e.value), t.syncExecuteCommand(wn.id);
  }
}, wc = {
  id: "sheet.command.set-border-basic",
  type: v.COMMAND,
  handler: (n, e) => {
    const { unitId: t, subUnitId: o, value: s, ranges: r } = e, { type: i, color: a, style: u } = s, l = n.get(E), c = n.get(Ht);
    return c.setType(i), a && c.setColor(a), c.setStyle(u), l.syncExecuteCommand(wn.id, {
      unitId: t,
      subUnitId: o,
      ranges: r
    });
  }
}, Ic = {
  type: v.COMMAND,
  id: "sheet.command.set-col-data",
  handler: (n, e) => {
    const t = n.get(E), o = n.get(V), s = n.get(M), r = P(s, e);
    if (!r) return !1;
    const { columnData: i } = e, { unitId: a, subUnitId: u, worksheet: l } = r, c = {
      subUnitId: u,
      unitId: a,
      columnData: i
    }, d = Fa(c, l);
    return t.syncExecuteCommand(zt.id, c) ? (o.pushUndoRedo({
      unitID: a,
      undoMutations: [{ id: zt.id, params: d }],
      redoMutations: [{ id: zt.id, params: c }]
    }), !0) : !1;
  }
}, dn = {
  type: v.COMMAND,
  id: "sheet.command.set-col-visible-on-cols",
  handler: (n, e) => {
    var p, w;
    const { unitId: t, subUnitId: o, ranges: s } = e, r = n.get(G), i = n.get(E), a = n.get(M), u = P(a, { unitId: t, subUnitId: o });
    if (!u) return !1;
    const { worksheet: l } = u, c = {
      unitId: t,
      subUnitId: o,
      ranges: s
    }, d = {
      unitId: t,
      subUnitId: o,
      reveal: !0,
      selections: s.map((I) => ({ range: I, primary: _e(I, l), style: null }))
    }, h = ja(n, c), g = {
      unitId: t,
      subUnitId: o,
      selections: Yr(s).map((I) => ({
        range: I,
        primary: _e(I, l),
        style: null
      }))
    }, m = L([
      { id: sn.id, params: c },
      { id: ie.id, params: d }
    ], i), f = r.onCommandExecute({
      id: dn.id,
      params: e
    }), C = L([...f.redos], i);
    if (m.result && C.result) {
      const I = r.afterCommandExecute({
        id: dn.id,
        params: e
      });
      return L(I.redos, i), n.get(V).pushUndoRedo({
        unitID: t,
        undoMutations: [
          { id: on.id, params: h },
          { id: ie.id, params: g },
          ...(p = f.undos) != null ? p : [],
          ...I.undos
        ],
        redoMutations: [
          ...(w = f.preRedos) != null ? w : [],
          { id: sn.id, params: c },
          { id: ie.id, params: d },
          ...f.redos,
          ...I.redos
        ]
      }), !0;
    }
    return !0;
  }
}, qr = {
  type: v.COMMAND,
  id: "sheet.command.set-selected-cols-visible",
  handler: (n) => {
    var l;
    const e = n.get(z), t = n.get(E), o = (l = e.getCurrentSelections()) == null ? void 0 : l.map((c) => c.range).filter((c) => c.rangeType === j.COLUMN);
    if (!(o != null && o.length)) return !1;
    const s = P(n.get(M));
    if (!s) return !1;
    const { worksheet: r, unitId: i, subUnitId: a } = s, u = o.map((c) => r.getHiddenCols(c.startColumn, c.endColumn)).flat();
    return t.executeCommand(dn.id, {
      unitId: i,
      subUnitId: a,
      ranges: u
    });
  }
}, Eo = {
  type: v.COMMAND,
  id: "sheet.command.set-col-hidden",
  handler: (n, e) => {
    var w, I, S, y;
    const t = n.get(z), o = n.get(G), s = n.get(M), r = n.get(E);
    let i = (w = e == null ? void 0 : e.ranges) != null && w.length ? e.ranges : (I = t.getCurrentSelections()) == null ? void 0 : I.map((b) => b.range).filter((b) => b.rangeType === j.COLUMN);
    if (!(i != null && i.length)) return !1;
    const a = P(s, e);
    if (!a) return !1;
    const { worksheet: u, unitId: l, subUnitId: c } = a;
    i = vc(a.worksheet, i);
    const d = { unitId: l, subUnitId: c, ranges: i }, h = {
      unitId: l,
      subUnitId: c,
      selections: Yr(i).map((b) => ({
        range: b,
        primary: _e(b, u),
        style: null
      }))
    }, g = Ba(n, d), m = {
      unitId: l,
      subUnitId: c,
      reveal: !0,
      selections: i.map((b) => ({
        range: b,
        primary: _e(b, u),
        style: null
      }))
    }, f = L([
      { id: on.id, params: d },
      { id: ie.id, params: h }
    ], r), C = o.onCommandExecute({
      id: Eo.id,
      params: d
    }), p = L([...C.redos], r);
    if (f.result && p.result) {
      const b = o.afterCommandExecute({
        id: Eo.id,
        params: d
      });
      return L(b.redos, r), n.get(V).pushUndoRedo({
        unitID: l,
        undoMutations: [
          { id: sn.id, params: g },
          { id: ie.id, params: m },
          ...(S = C.undos) != null ? S : [],
          ...b.undos
        ],
        redoMutations: [
          ...(y = C.preRedos) != null ? y : [],
          { id: on.id, params: d },
          { id: ie.id, params: h },
          ...C.redos,
          ...b.redos
        ]
      }), !0;
    }
    return !1;
  }
};
function vc(n, e) {
  const t = n.getRowCount() - 1, o = n.getHiddenCols(), s = [];
  return e.forEach((r) => {
    const i = o.filter((a) => a.startColumn >= r.startColumn && a.endColumn <= r.endColumn);
    if (i.length) {
      let a = r.startColumn;
      i.forEach((u) => {
        u.startColumn > a && (s.push({ startColumn: a, endColumn: u.startColumn - 1, startRow: 0, endRow: t }), a = u.endColumn + 1);
      }), a <= r.endColumn && s.push({ startColumn: a, endColumn: r.endColumn, startRow: 0, endRow: t });
    } else
      s.push(r);
  }), s;
}
function Yr(n) {
  return Mc(n).map((t) => {
    const o = t.startColumn === 0 ? t.endColumn + 1 : t.startColumn - 1;
    return {
      ...t,
      startColumn: o,
      endColumn: o
    };
  });
}
function Mc(n) {
  const e = [];
  let t;
  return n.sort((o, s) => o.startColumn - s.startColumn).forEach((o) => {
    if (!t) {
      t = o;
      return;
    }
    t.endColumn === o.startColumn - 1 ? t.endColumn = o.endColumn : (e.push(t), t = o);
  }), e.push(t), e;
}
const Xr = {
  id: "sheet.command.set-defined-name",
  type: v.COMMAND,
  handler: (n, e) => {
    var d, h;
    const t = n.get(E), o = n.get(V), s = n.get(G);
    if (!e) return !1;
    const r = {
      ...e
    }, i = Yi(n, e), a = s.onCommandExecute({ id: Xr.id, params: e }), u = [
      ...(d = a.preRedos) != null ? d : [],
      { id: Pn.id, params: i },
      { id: en.id, params: r },
      ...a.redos
    ], l = [
      ...(h = a.preUndos) != null ? h : [],
      { id: Pn.id, params: r },
      { id: en.id, params: i },
      ...a.undos
    ];
    return L(u, t) ? (o.pushUndoRedo({
      unitID: e.unitId,
      undoMutations: l.filter(Boolean),
      redoMutations: u.filter(Boolean)
    }), !0) : !1;
  }
}, Zr = (n, e) => {
  const o = n.get(M).getUniverSheetInstance(e.unitId);
  if (o == null)
    throw new Error("universheet is null error!");
  const s = o.getSheetBySheetId(e.subUnitId);
  if (s == null)
    throw new Error("worksheet is null error!");
  const i = s.getConfig().freeze;
  return {
    unitId: e.unitId,
    subUnitId: e.subUnitId,
    ...i
  };
}, ot = {
  id: "sheet.mutation.set-frozen",
  type: v.MUTATION,
  handler: (n, e) => {
    const o = n.get(M).getUniverSheetInstance(e.unitId);
    if (o == null)
      throw new Error("universheet is null error!");
    const s = o.getSheetBySheetId(e.subUnitId);
    if (!s) return !1;
    const r = s.getConfig(), { startRow: i, startColumn: a, ySplit: u, xSplit: l } = e;
    return r.freeze = { startRow: i, startColumn: a, ySplit: u, xSplit: l }, !0;
  }
}, _c = {
  type: v.COMMAND,
  id: "sheet.command.set-frozen",
  handler: (n, e) => {
    const t = n.get(E), o = n.get(V), s = n.get(M), r = P(s, { unitId: e.unitId, subUnitId: e.subUnitId });
    if (!r) return !1;
    const { unitId: i, subUnitId: a, worksheet: u } = r, { startColumn: l, startRow: c, xSplit: d, ySplit: h } = e;
    if (c >= u.getRowCount() || l >= u.getColumnCount() || d >= u.getColumnCount() || h >= u.getRowCount())
      return !1;
    const g = {
      unitId: i,
      subUnitId: a,
      ...e
    }, m = Zr(n, g);
    return t.syncExecuteCommand(ot.id, g) ? (o.pushUndoRedo({
      unitID: i,
      undoMutations: [{ id: ot.id, params: m }],
      redoMutations: [{ id: ot.id, params: g }]
    }), !0) : !1;
  }
}, yc = {
  type: v.COMMAND,
  id: "sheet.command.cancel-frozen",
  handler: (n, e) => {
    const t = n.get(E), o = n.get(M), s = n.get(V), r = P(o, { unitId: e == null ? void 0 : e.unitId, subUnitId: e == null ? void 0 : e.subUnitId });
    if (!r) return !1;
    const { unitId: i, subUnitId: a } = r, u = {
      unitId: i,
      subUnitId: a,
      startRow: -1,
      startColumn: -1,
      xSplit: 0,
      ySplit: 0
    }, l = Zr(n, u);
    return t.syncExecuteCommand(ot.id, u) && s.pushUndoRedo({
      unitID: i,
      undoMutations: [{ id: ot.id, params: l }],
      redoMutations: [{ id: ot.id, params: u }]
    }), !0;
  }
}, bc = {
  type: v.COMMAND,
  id: "sheet.command.set-gridlines-color",
  handler: (n, e) => {
    const t = n.get(E), o = n.get(V), s = n.get(M), r = P(s);
    if (!r) return !1;
    const { worksheet: i } = r, a = i.getConfig().gridlinesColor;
    if (a === (e == null ? void 0 : e.color)) return !1;
    const { unitId: u, subUnitId: l } = r, c = {
      color: e == null ? void 0 : e.color,
      unitId: u,
      subUnitId: l
    }, d = {
      color: a,
      unitId: u,
      subUnitId: l
    };
    return t.syncExecuteCommand(Gt.id, c) ? (o.pushUndoRedo({
      unitID: u,
      undoMutations: [{ id: Gt.id, params: d }],
      redoMutations: [{ id: Gt.id, params: c }]
    }), !0) : !1;
  }
}, fe = {
  id: "sheet.mutation.set-range-protection",
  type: v.MUTATION,
  handler: (n, e) => {
    const { unitId: t, subUnitId: o, rule: s, ruleId: r } = e;
    return n.get(me).setRule(t, o, r, s), !0;
  }
}, Fh = (n, e) => {
  const { unitId: t, subUnitId: o, ruleId: s } = e, i = n.get(me).getRule(t, o, s);
  return i ? {
    id: fe.id,
    params: {
      ...e,
      rule: i
    }
  } : null;
}, bt = {
  id: "sheet.mutation.set-worksheet-protection",
  type: v.MUTATION,
  handler: (n, e) => {
    const { unitId: t, subUnitId: o, rule: s } = e;
    return n.get(lt).setRule(t, o, s), !0;
  }
}, Ec = {
  type: v.COMMAND,
  id: "sheet.command.set-protection",
  async handler(n, e) {
    if (!e)
      return !1;
    const t = n.get(E), o = n.get(V), s = n.get(me), { rule: r, oldRule: i } = e, { unitId: a, subUnitId: u } = r, l = [], c = [];
    return (i == null ? void 0 : i.unitType) === r.unitType ? r.unitType === x.Worksheet ? (l.push({ id: bt.id, params: { unitId: a, subUnitId: u, rule: r } }), c.push({ id: bt.id, params: { unitId: a, subUnitId: u, rule: i } })) : (l.push({ id: fe.id, params: { unitId: a, subUnitId: u, rule: r, ruleId: r.id } }), c.push({ id: fe.id, params: { unitId: a, subUnitId: u, ruleId: i.id, rule: i } })) : (i && (i.unitType === x.Worksheet ? (l.push({ id: xt.id, params: { unitId: a, subUnitId: u } }), c.push({ id: Mt.id, params: { unitId: a, rule: i, subUnitId: i.subUnitId } })) : i.unitType === x.SelectRange && (l.push({ id: st.id, params: { unitId: a, subUnitId: u, ruleIds: [i.id] } }), c.push({ id: Be.id, params: { unitId: a, subUnitId: u, rules: [i] } }))), r.unitType === x.Worksheet ? (l.push({ id: Mt.id, params: { unitId: a, rule: r, subUnitId: r.subUnitId } }), c.unshift({ id: xt.id, params: { unitId: a, subUnitId: u } })) : r.unitType === x.SelectRange && (r.id = s.createRuleId(a, u), l.push({ id: Be.id, params: { unitId: a, subUnitId: u, rules: [r] } }), c.unshift({ id: st.id, params: { unitId: a, subUnitId: u, ruleIds: [r.id] } }))), L(l, t) && o.pushUndoRedo({
      unitID: a,
      undoMutations: c,
      redoMutations: l
    }), !0;
  }
}, Uc = {
  type: v.COMMAND,
  id: "sheet.command.set-row-data",
  handler: (n, e) => {
    const t = n.get(E), o = n.get(V), s = n.get(M), r = P(s, e);
    if (!r) return !1;
    const { rowData: i } = e, { unitId: a, subUnitId: u, worksheet: l } = r, c = {
      subUnitId: u,
      unitId: a,
      rowData: i
    }, d = eu(c, l);
    return t.syncExecuteCommand(Kt.id, c) ? (o.pushUndoRedo({
      unitID: a,
      undoMutations: [{ id: Kt.id, params: d }],
      redoMutations: [{ id: Kt.id, params: c }]
    }), !0) : !1;
  }
}, hn = {
  type: v.COMMAND,
  id: "sheet.command.set-specific-rows-visible",
  handler: (n, e) => {
    var p, w, I;
    const { unitId: t, subUnitId: o, ranges: s } = e, r = n.get(E), i = n.get(V), a = n.get(G), u = P(n.get(M), { unitId: t, subUnitId: o });
    if (!u) return !1;
    const { worksheet: l } = u, c = { unitId: t, subUnitId: o, ranges: s }, d = {
      unitId: t,
      subUnitId: o,
      reveal: !0,
      selections: s.map((S) => ({
        range: S,
        primary: _e(S, l),
        style: null
      }))
    }, h = tu(n, c), g = {
      unitId: t,
      subUnitId: o,
      selections: ei(s).map((S) => ({
        range: S,
        primary: _e(S, l),
        style: null
      }))
    }, m = L(
      [
        { id: kt.id, params: c },
        { id: ie.id, params: d }
      ],
      r
    ), f = a.onCommandExecute({
      id: hn.id,
      params: e
    }), C = L([...f.redos], r);
    if (m.result && C.result) {
      const S = a.afterCommandExecute({
        id: hn.id,
        params: e
      });
      return L(S.redos, r), i.pushUndoRedo({
        unitID: t,
        undoMutations: [
          ...(p = f.preUndos) != null ? p : [],
          { id: Tt.id, params: h },
          { id: ie.id, params: g },
          ...(w = f.undos) != null ? w : [],
          ...S.undos
        ],
        redoMutations: [
          ...(I = f.preRedos) != null ? I : [],
          { id: kt.id, params: c },
          { id: ie.id, params: d },
          ...f.redos,
          ...S.redos
        ]
      }), !0;
    }
    return !0;
  }
}, Qr = {
  type: v.COMMAND,
  id: "sheet.command.set-selected-rows-visible",
  handler: async (n) => {
    var c;
    const e = n.get(z), t = n.get(M), o = n.get(E), s = (c = e.getCurrentSelections()) == null ? void 0 : c.map((d) => d.range).filter((d) => d.rangeType === j.ROW);
    if (!(s != null && s.length)) return !1;
    const r = P(t);
    if (!r) return !1;
    const { worksheet: i, unitId: a, subUnitId: u } = r, l = s.map((d) => i.getHiddenRows(d.startRow, d.endRow)).flat();
    return o.executeCommand(hn.id, {
      unitId: a,
      subUnitId: u,
      ranges: l
    });
  }
}, Uo = {
  type: v.COMMAND,
  id: "sheet.command.set-rows-hidden",
  handler: (n, e) => {
    var w, I, S, y, b, O;
    const t = n.get(z), o = n.get(E), s = n.get(V), r = n.get(M), i = n.get(G);
    let a = (w = e == null ? void 0 : e.ranges) != null && w.length ? e.ranges : (I = t.getCurrentSelections()) == null ? void 0 : I.map((k) => k.range).filter((k) => k.rangeType === j.ROW);
    if (!(a != null && a.length)) return !1;
    const u = P(r, e);
    if (!u) return !1;
    a = kc(u.worksheet, a);
    const { unitId: l, subUnitId: c, worksheet: d } = u, h = { unitId: l, subUnitId: c, ranges: a }, g = {
      unitId: l,
      subUnitId: c,
      selections: ei(a).map((k) => ({
        range: k,
        primary: _e(k, d),
        style: null
      }))
    }, m = nu(n, h), f = {
      unitId: l,
      subUnitId: c,
      reveal: !0,
      selections: a.map((k) => ({
        range: k,
        primary: _e(k, d),
        style: null
      }))
    }, C = i.onCommandExecute({ id: Uo.id, params: h });
    if (L([
      ...(S = C.preRedos) != null ? S : [],
      { id: Tt.id, params: h },
      { id: ie.id, params: g },
      ...C.redos
    ], o).result) {
      const k = i.afterCommandExecute({
        id: Uo.id,
        params: h
      });
      return L(k.redos, o), s.pushUndoRedo({
        unitID: l,
        undoMutations: [
          ...(y = C.preUndos) != null ? y : [],
          { id: kt.id, params: m },
          { id: ie.id, params: f },
          ...(b = C.undos) != null ? b : [],
          ...k.undos
        ],
        redoMutations: [
          ...(O = C.preRedos) != null ? O : [],
          { id: Tt.id, params: h },
          { id: ie.id, params: g },
          ...C.redos,
          ...k.redos
        ]
      }), !0;
    }
    return !0;
  }
};
function kc(n, e) {
  const t = n.getMaxColumns() - 1, o = n.getHiddenRows(), s = [];
  return e.forEach((r) => {
    const i = o.filter((a) => a.startRow >= r.startRow && a.endRow <= r.endRow);
    if (i.length) {
      let a = r.startRow;
      i.forEach((u) => {
        u.startRow > a && (s.push({ startRow: a, endRow: u.startRow - 1, startColumn: 0, endColumn: t }), a = u.endRow + 1);
      }), a <= r.endRow && s.push({ startRow: a, endRow: r.endRow, startColumn: 0, endColumn: t });
    } else
      s.push(r);
  }), s;
}
function ei(n) {
  return Tc(n).map((t) => {
    const o = t.startRow === 0 ? t.endRow + 1 : t.startRow - 1;
    return {
      ...t,
      startRow: o,
      endRow: o
    };
  });
}
function Tc(n) {
  const e = [];
  let t;
  return n.sort((o, s) => o.startRow - s.startRow).forEach((o) => {
    if (!t) {
      t = o;
      return;
    }
    o.startRow === t.endRow + 1 ? t.endRow = o.endRow : (e.push(t), t = o);
  }), e.push(t), e;
}
const Pc = ["ff", "fs", "tr", "tb"], Ce = {
  type: v.COMMAND,
  id: "sheet.command.set-style",
  // eslint-disable-next-line max-lines-per-function
  handler: (n, e) => {
    var T;
    const t = n.get(M), o = P(t, e);
    if (!o) return !1;
    const { unitId: s, subUnitId: r, worksheet: i } = o, { range: a, style: u } = e, l = n.get(E), c = n.get(V), d = n.get(z), h = a ? [a] : (T = d.getCurrentSelections()) == null ? void 0 : T.map((U) => U.range);
    if (!(h != null && h.length))
      return !1;
    const g = new Y(), m = Wu(i);
    if (A.isArray(u.value))
      for (let U = 0; U < h.length; U++)
        m.forOperableEach(h[U], (D, H, F) => {
          g.setValue(D, H, {
            s: {
              [u.type]: u.value[D - F.startRow][H - F.startColumn]
            }
          });
        });
    else
      for (let U = 0; U < h.length; U++) {
        const D = {
          s: {
            [u.type]: u.value
          }
        };
        m.forOperableEach(h[U], (H, F) => g.setValue(H, F, D));
      }
    const f = {
      subUnitId: r,
      unitId: s,
      cellValue: g.getMatrix()
    }, C = n.get(rt).getSkeleton(s, r), p = Ae(
      n,
      f
    ), w = l.syncExecuteCommand(
      ee.id,
      f
    ), I = n.get(G);
    let S = [], y = [];
    if (Pc.includes(e == null ? void 0 : e.style.type)) {
      const { suitableRanges: U, remainingRanges: D } = Sn(h, C), H = Jo(U, i), { undos: F, redos: K } = I.generateMutationsOfAutoHeight({
        unitId: s,
        subUnitId: r,
        ranges: U,
        autoHeightRanges: U,
        lazyAutoHeightRanges: D,
        cellHeights: H
      });
      S = F, y = K;
    }
    const { undos: b, redos: O } = I.onCommandExecute({
      id: Ce.id,
      params: e
    }), k = L([...O, ...y], l);
    return w && k.result ? (c.pushUndoRedo({
      unitID: f.unitId,
      undoMutations: [{ id: ee.id, params: p }, ...b, ...S],
      redoMutations: [{ id: ee.id, params: f }, ...O, ...y]
    }), !0) : !1;
  }
}, Bh = {
  type: v.COMMAND,
  id: "sheet.command.set-bold",
  handler: (n) => {
    const e = n.get(z).getCurrentLastSelection();
    if (!e) return !1;
    const t = P(n.get(M));
    if (!t) return !1;
    const { worksheet: o } = t, { actualRow: s, actualColumn: r } = e.primary, a = {
      style: {
        type: "bl",
        value: o.getRange(s, r).getFontWeight() === ki.BOLD ? te.FALSE : te.TRUE
      }
    };
    return n.get(E).syncExecuteCommand(Ce.id, a);
  }
}, jh = {
  type: v.COMMAND,
  id: "sheet.command.set-italic",
  handler: (n) => {
    const e = n.get(z).getCurrentLastSelection();
    if (!e) return !1;
    const t = P(n.get(M));
    if (!t) return !1;
    const { worksheet: o } = t;
    let s = !0;
    if (e.primary) {
      const { startRow: i, startColumn: a } = e.primary;
      s = o.getRange(i, a).getFontStyle() === Ti.ITALIC;
    }
    const r = {
      style: {
        type: "it",
        value: s ? te.FALSE : te.TRUE
      }
    };
    return n.get(E).syncExecuteCommand(Ce.id, r);
  }
}, zh = {
  type: v.COMMAND,
  id: "sheet.command.set-underline",
  handler: (n) => {
    const e = n.get(z).getCurrentLastSelection();
    if (!e) return !1;
    const t = P(n.get(M));
    if (!t) return !1;
    const { worksheet: o } = t;
    let s = !0;
    e.primary && (s = !!o.getRange(e.primary.startRow, e.primary.startColumn).getUnderline().s);
    const r = {
      style: {
        type: "ul",
        value: {
          s: s ? te.FALSE : te.TRUE
        }
      }
    };
    return n.get(E).syncExecuteCommand(Ce.id, r);
  }
}, Gh = {
  type: v.COMMAND,
  id: "sheet.command.set-stroke",
  handler: (n) => {
    const e = n.get(z).getCurrentLastSelection();
    if (!e) return !1;
    const t = P(n.get(M));
    if (!t) return !1;
    const { worksheet: o } = t;
    let s = !0;
    e.primary && (s = !!o.getRange(e.primary.actualRow, e.primary.actualColumn).getStrikeThrough().s);
    const r = {
      style: {
        type: "st",
        value: { s: s ? te.FALSE : te.TRUE }
      }
    };
    return n.get(E).syncExecuteCommand(Ce.id, r);
  }
}, Kh = {
  type: v.COMMAND,
  id: "sheet.command.set-overline",
  handler: (n) => {
    const e = n.get(z).getCurrentLastSelection();
    if (!e) return !1;
    const t = P(n.get(M));
    if (!t) return !1;
    const { worksheet: o } = t;
    let s = !0;
    e.primary && (s = !!o.getRange(e.primary.startRow, e.primary.startColumn).getOverline().s);
    const r = {
      style: {
        type: "ol",
        value: {
          s: s ? te.FALSE : te.TRUE
        }
      }
    };
    return n.get(E).syncExecuteCommand(Ce.id, r);
  }
}, Jh = {
  type: v.COMMAND,
  id: "sheet.command.set-font-family",
  handler: (n, e) => {
    if (!e)
      return !1;
    const t = n.get(E), o = {
      style: {
        type: "ff",
        value: e.value
      }
    };
    return t.syncExecuteCommand(Ce.id, o);
  }
}, qh = {
  type: v.COMMAND,
  id: "sheet.command.set-font-size",
  handler: (n, e) => {
    if (!e)
      return !1;
    const t = n.get(E), o = {
      style: {
        type: "fs",
        value: e.value
      }
    };
    return t.syncExecuteCommand(Ce.id, o);
  }
}, Nc = {
  type: v.COMMAND,
  id: "sheet.command.set-text-color",
  handler: (n, e) => {
    if (!e)
      return !1;
    const t = n.get(E), o = {
      style: {
        type: "cl",
        value: {
          rgb: e.value
        }
      }
    };
    return t.syncExecuteCommand(Ce.id, o);
  }
}, Oc = {
  type: v.COMMAND,
  id: "sheet.command.reset-text-color",
  handler: (n) => {
    const e = n.get(E), t = {
      style: {
        type: "cl",
        value: {
          rgb: null
          // use null to reset text color
        }
      }
    };
    return e.syncExecuteCommand(Ce.id, t);
  }
}, Dc = {
  type: v.COMMAND,
  id: "sheet.command.set-background-color",
  handler: (n, e) => {
    if (!e || !e.value)
      return !1;
    const t = n.get(E), o = {
      style: {
        type: "bg",
        value: {
          rgb: e.value
        }
      }
    };
    return t.syncExecuteCommand(Ce.id, o);
  }
}, Ac = {
  type: v.COMMAND,
  id: "sheet.command.reset-background-color",
  handler: (n) => {
    const e = n.get(E), t = {
      style: {
        type: "bg",
        value: {
          rgb: null
          // use null to reset background color
        }
      }
    };
    return e.syncExecuteCommand(Ce.id, t);
  }
}, xc = {
  type: v.COMMAND,
  id: "sheet.command.set-vertical-text-align",
  handler: (n, e) => {
    if (!e)
      return !1;
    const t = n.get(E), o = {
      unitId: e.unitId,
      subUnitId: e.subUnitId,
      range: e.range,
      style: {
        type: "vt",
        value: e.value
      }
    };
    return t.syncExecuteCommand(Ce.id, o);
  }
}, $c = {
  type: v.COMMAND,
  id: "sheet.command.set-horizontal-text-align",
  handler: (n, e) => {
    if (!e)
      return !1;
    const t = n.get(E), o = {
      unitId: e.unitId,
      subUnitId: e.subUnitId,
      range: e.range,
      style: {
        type: "ht",
        value: e.value
      }
    };
    return t.syncExecuteCommand(Ce.id, o);
  }
}, Wc = {
  type: v.COMMAND,
  id: "sheet.command.set-text-wrap",
  handler: (n, e) => {
    if (!e)
      return !1;
    const t = n.get(E), o = {
      unitId: e.unitId,
      subUnitId: e.subUnitId,
      range: e.range,
      style: {
        type: "tb",
        value: e.value
      }
    };
    return t.syncExecuteCommand(Ce.id, o);
  }
}, Vc = {
  type: v.COMMAND,
  id: "sheet.command.set-text-rotation",
  handler: (n, e) => {
    if (!e)
      return !1;
    const t = typeof e.value == "number" ? { a: e.value } : { a: 0, v: te.TRUE }, o = n.get(E), s = {
      unitId: e.unitId,
      subUnitId: e.subUnitId,
      range: e.range,
      style: {
        type: "tr",
        value: t
      }
    };
    return o.syncExecuteCommand(Ce.id, s);
  }
}, Lc = (n, e) => {
  const r = n.get(M).getUniverSheetInstance(e.unitId).getSheetBySheetId(e.subUnitId).getConfig().tabColor;
  return {
    ...A.deepClone(e),
    color: r
  };
}, Un = {
  id: "sheet.mutation.set-tab-color",
  type: v.MUTATION,
  handler: (n, e) => {
    const t = n.get(M).getUniverSheetInstance(e.unitId);
    if (!t) return !1;
    const o = t.getSheetBySheetId(e.subUnitId);
    return o ? (o.getConfig().tabColor = e.color, !0) : !1;
  }
}, Hc = {
  type: v.COMMAND,
  id: "sheet.command.set-tab-color",
  handler: (n, e) => {
    const t = n.get(E), o = n.get(V), s = P(n.get(M));
    if (!s) return !1;
    const { unitId: r, subUnitId: i } = s, a = {
      color: e.value,
      unitId: r,
      subUnitId: i
    }, u = Lc(n, a);
    return t.syncExecuteCommand(Un.id, a) ? (o.pushUndoRedo({
      unitID: r,
      undoMutations: [{ id: Un.id, params: u }],
      redoMutations: [{ id: Un.id, params: a }]
    }), !0) : !1;
  }
}, ti = {
  id: "sheet.mutation.set-workbook-name",
  type: v.MUTATION,
  handler: (n, e) => {
    const t = n.get(M).getUnit(e.unitId, B.UNIVER_SHEET);
    return t ? (t.setName(e.name), !0) : !1;
  }
}, ni = {
  type: v.COMMAND,
  id: "sheet.command.set-workbook-name",
  handler: (n, e) => {
    var u;
    const t = n.get(E), o = n.get(G);
    if (!Ks(n.get(M), e)) return !1;
    const r = o.onCommandExecute({
      id: ni.id,
      params: e
    }), i = {
      name: e.name,
      unitId: e.unitId
    }, a = [
      ...(u = r.preRedos) != null ? u : [],
      { id: ti.id, params: i },
      ...r.redos
    ];
    return L(a, t).result;
  }
}, Fc = 4, oi = {
  type: v.COMMAND,
  id: "sheet.command.set-worksheet-activate",
  handler: (n, e, t) => {
    const o = n.get(E), s = P(n.get(M), e);
    if (!s) return !1;
    const { unitId: r, subUnitId: i } = s;
    return new Promise((a) => {
      setTimeout(() => {
        const u = o.syncExecuteCommand(Cn.id, {
          unitId: r,
          subUnitId: i
        }, t);
        a(u);
      }, Fc);
    });
  }
}, Ln = {
  type: v.COMMAND,
  id: "sheet.command.delta-column-width",
  // eslint-disable-next-line max-lines-per-function
  handler: async (n, e) => {
    const o = n.get(z).getCurrentSelections();
    if (!(o != null && o.length))
      return !1;
    const s = n.get(E), r = n.get(V), i = P(n.get(M));
    if (!i) return !1;
    const { worksheet: a, unitId: u, subUnitId: l } = i, { anchorCol: c, deltaX: d } = e, g = a.getColumnWidth(c) + d, m = o.length === 1 && o[0].range.rangeType === j.ALL, f = o.filter((J) => J.range.rangeType === j.COLUMN), C = m ? j.ALL : f.some(({ range: J }) => {
      const { startColumn: ae, endColumn: se } = J;
      return ae <= c && c <= se;
    }) ? j.COLUMN : j.NORMAL;
    let p;
    if (C === j.ALL) {
      const J = a.getRowCount(), ae = new Array(a.getColumnCount()).fill(void 0).map(
        (se, le) => ({ startRow: 0, endRow: J - 1, startColumn: le, endColumn: le })
      );
      p = {
        subUnitId: l,
        unitId: u,
        colWidth: g,
        ranges: ae
      };
    } else C === j.COLUMN ? p = {
      subUnitId: l,
      unitId: u,
      ranges: f.map((J) => N.clone(J.range)),
      colWidth: g
    } : p = {
      subUnitId: l,
      unitId: u,
      colWidth: g,
      ranges: [
        {
          startRow: 0,
          endRow: a.getMaxRows() - 1,
          startColumn: c,
          endColumn: c
        }
      ]
    };
    const w = n.get(rt).getSkeleton(u, l), { suitableRanges: I, remainingRanges: S } = Sn(p.ranges, w);
    Jo(I, a);
    const y = n.get(G), { undos: b, redos: O } = y.onCommandExecute({
      id: Ln.id,
      params: p
    }), k = qs(
      p,
      a
    ), T = s.syncExecuteCommand(
      mt.id,
      p
    ), { undos: U, redos: D } = y.generateMutationsOfAutoHeight({
      unitId: u,
      subUnitId: l,
      ranges: I,
      autoHeightRanges: I,
      lazyAutoHeightRanges: S
    }), { undos: H, redos: F } = n.get(G).afterCommandExecute({
      id: Ln.id,
      params: p
    }), K = L([...O, ...F, ...D], s);
    return T && K.result && r.pushUndoRedo({
      unitID: u,
      undoMutations: [{ id: mt.id, params: k }, ...b, ...H, ...U],
      redoMutations: [{ id: mt.id, params: p }, ...O, ...F, ...D]
    }), !0;
  }
}, Hn = {
  type: v.COMMAND,
  id: "sheet.command.set-worksheet-col-width",
  handler: (n, e) => {
    var y, b, O, k;
    const t = n.get(z), o = n.get(E), s = n.get(V), r = n.get(G), i = (y = e == null ? void 0 : e.ranges) != null && y.length ? e.ranges : (b = t.getCurrentSelections()) == null ? void 0 : b.map((T) => T.range);
    if (!(i != null && i.length)) return !1;
    const a = P(n.get(M), e);
    if (!a) return !1;
    const { subUnitId: u, unitId: l, worksheet: c } = a, d = n.get(rt).getSkeleton(l, u), h = {
      subUnitId: u,
      unitId: l,
      ranges: i,
      colWidth: e.value
    }, { suitableRanges: g, remainingRanges: m } = Sn(h.ranges, d);
    Jo(g, c);
    const f = qs(h, c), C = o.syncExecuteCommand(mt.id, h), { undos: p, redos: w } = r.generateMutationsOfAutoHeight({
      unitId: l,
      subUnitId: u,
      ranges: g,
      autoHeightRanges: g,
      lazyAutoHeightRanges: m
    }), I = r.onCommandExecute({
      id: Hn.id,
      params: h
    }), S = L([...I.redos, ...w], o);
    if (C && S.result) {
      const T = r.afterCommandExecute({
        id: Hn.id,
        params: h
      });
      return L(T.redos, o), s.pushUndoRedo({
        unitID: l,
        undoMutations: [
          ...(O = I.preUndos) != null ? O : [],
          { id: mt.id, params: f },
          ...I.undos,
          ...T.undos,
          ...p
        ],
        redoMutations: [
          ...(k = I.preRedos) != null ? k : [],
          { id: mt.id, params: h },
          ...I.redos,
          ...T.redos,
          ...w
        ]
      }), !0;
    }
    return !1;
  }
};
v.COMMAND;
const Bc = {
  type: v.COMMAND,
  id: "sheet.command.set-worksheet-column-count",
  handler: (n, e) => {
    const { unitId: t, subUnitId: o, columnCount: s } = e, r = n.get(E), i = n.get(V), a = n.get(M);
    if (!P(a, e)) return !1;
    const l = {
      unitId: t,
      subUnitId: o,
      columnCount: s
    }, c = ou(n, l);
    return r.syncExecuteCommand(Jt.id, l) ? (i.pushUndoRedo({
      unitID: t,
      undoMutations: [{ id: Jt.id, params: c }],
      redoMutations: [{ id: Jt.id, params: l }]
    }), !0) : !1;
  }
}, jc = {
  type: v.COMMAND,
  id: "sheet.command.set-worksheet-default-style",
  handler: (n, e) => {
    const t = n.get(E), o = n.get(V), { unitId: s } = e, r = su(n, e);
    return t.syncExecuteCommand(qt.id, e) ? (o.pushUndoRedo({
      unitID: s,
      undoMutations: [{ id: qt.id, params: r }],
      redoMutations: [{ id: qt.id, params: e }]
    }), !0) : !1;
  }
}, si = (n, e) => {
  const t = Qe(n.get(M), e);
  if (!t)
    throw new Error("[SetWorksheetHideMutationFactory]: worksheet is null error!");
  const { worksheet: o } = t;
  return {
    hidden: o.isSheetHidden(),
    unitId: e.unitId,
    subUnitId: o.getSheetId()
  };
}, St = {
  id: "sheet.mutation.set-worksheet-hidden",
  type: v.MUTATION,
  handler: (n, e) => {
    const t = n.get(M).getUniverSheetInstance(e.unitId);
    if (t == null)
      return !1;
    const o = t.getSheetBySheetId(e.subUnitId);
    return o ? (o.getConfig().hidden = e.hidden, !0) : !1;
  }
}, zc = {
  type: v.COMMAND,
  id: "sheet.command.set-worksheet-hidden",
  handler: (n, e) => {
    const t = n.get(E), o = n.get(V), s = n.get(Vt), r = n.get(it), i = P(n.get(M), e);
    if (!i) return !1;
    const { workbook: a, worksheet: u, unitId: l, subUnitId: c } = i;
    if (u.getConfig().hidden === te.TRUE) return !1;
    const h = {
      unitId: l,
      subUnitId: c,
      hidden: te.TRUE
    }, g = si(n, h);
    return a.getSheets().filter((p) => p.getConfig().hidden === te.FALSE).length === 1 ? (s.emit(r.t("sheets.info.hideSheet")), !1) : t.syncExecuteCommand(St.id, h) ? (o.pushUndoRedo({
      unitID: l,
      undoMutations: [{ id: St.id, params: g }],
      redoMutations: [{ id: St.id, params: h }]
    }), !0) : !1;
  }
}, Gc = (n, e) => {
  const t = Qe(n.get(M), e);
  if (!t)
    throw new Error("[SetWorksheetNameMutationFactory]: worksheet is null error!");
  const { worksheet: o } = t;
  return {
    unitId: e.unitId,
    name: o.getName(),
    subUnitId: o.getSheetId()
  };
}, Fn = {
  id: "sheet.mutation.set-worksheet-name",
  type: v.MUTATION,
  handler: (n, e) => {
    const t = n.get(M).getUniverSheetInstance(e.unitId);
    if (t == null)
      return !1;
    const o = t.getSheetBySheetId(e.subUnitId);
    return o ? (o.getConfig().name = e.name, !0) : !1;
  }
}, Qo = {
  type: v.COMMAND,
  id: "sheet.command.set-worksheet-name",
  handler: (n, e) => {
    var m, f;
    const t = n.get(E), o = n.get(V), s = n.get(G), r = P(n.get(M), e);
    if (!r) return !1;
    const { unitId: i, subUnitId: a } = r, u = {
      subUnitId: a,
      name: e.name,
      unitId: i
    }, l = Gc(
      n,
      u
    ), c = s.onCommandExecute({
      id: Qo.id,
      params: e
    }), d = [
      ...(m = c.preRedos) != null ? m : [],
      { id: Fn.id, params: u },
      ...c.redos
    ], h = [
      ...(f = c.preUndos) != null ? f : [],
      { id: Fn.id, params: l },
      ...c.undos
    ];
    return L(d, t).result ? (o.pushUndoRedo({
      unitID: i,
      undoMutations: h,
      redoMutations: d
    }), !0) : !1;
  }
}, Kc = (n, e) => ({
  ...A.deepClone(e),
  toOrder: e.fromOrder,
  fromOrder: e.toOrder
}), kn = {
  id: "sheet.mutation.set-worksheet-order",
  type: v.MUTATION,
  handler: (n, e) => {
    const t = n.get(M).getUniverSheetInstance(e.unitId);
    if (!t) return !1;
    const o = t.getConfig();
    return o.sheetOrder.splice(e.fromOrder, 1), o.sheetOrder.splice(e.toOrder, 0, e.subUnitId), !0;
  }
}, ri = {
  type: v.COMMAND,
  id: "sheet.command.set-worksheet-order",
  handler: (n, e) => {
    const t = n.get(E), o = n.get(V), s = P(n.get(M), e);
    if (!s) return !1;
    const { workbook: r, unitId: i, subUnitId: a } = s, l = {
      fromOrder: r.getConfig().sheetOrder.indexOf(a),
      toOrder: e.order,
      unitId: i,
      subUnitId: a
    }, c = Kc(n, l);
    return t.syncExecuteCommand(kn.id, l) ? (o.pushUndoRedo({
      unitID: i,
      undoMutations: [{ id: kn.id, params: c }],
      redoMutations: [{ id: kn.id, params: l }]
    }), !0) : !1;
  }
};
class In {
  constructor() {
    R(this, "_model", /* @__PURE__ */ new Map());
    R(this, "_pointChange", new De());
    R(this, "pointChange$", this._pointChange.asObservable());
  }
  addRule(e) {
    this._ensureSubUnitMap(e.unitId).set(e.subUnitId, e), this._pointChange.next(e);
  }
  deleteRule(e, t) {
    var s, r, i;
    const o = (s = this._model.get(e)) == null ? void 0 : s.get(t);
    o && ((i = (r = this._model) == null ? void 0 : r.get(e)) == null || i.delete(t), this._pointChange.next(o));
  }
  getRule(e, t) {
    var o, s;
    return (s = (o = this._model) == null ? void 0 : o.get(e)) == null ? void 0 : s.get(t);
  }
  toObject() {
    const e = {};
    return [...this._model.keys()].forEach((o) => {
      const s = this._model.get(o);
      s != null && s.size && (e[o] = [], [...s.keys()].forEach((i) => {
        const a = s.get(i);
        a && e[o].push(a);
      }));
    }), e;
  }
  fromObject(e) {
    const t = /* @__PURE__ */ new Map();
    Object.keys(e).forEach((o) => {
      const s = e[o];
      if (s != null && s.length) {
        const r = /* @__PURE__ */ new Map();
        s.forEach((i) => {
          r.set(i.subUnitId, i);
        }), t.set(o, r);
      }
    }), this._model = t;
  }
  deleteUnitModel(e) {
    this._model.delete(e);
  }
  _ensureSubUnitMap(e) {
    let t = this._model.get(e);
    return t || (t = /* @__PURE__ */ new Map(), this._model.set(e, t)), t;
  }
  getTargetByPermissionId(e, t) {
    const o = this._model.get(e);
    if (!o) return null;
    for (const [s, r] of o)
      if (r.permissionId === t)
        return [e, s];
  }
}
class ii {
  constructor(e, t, o) {
    R(this, "type", x.SelectRange);
    R(this, "subType", _.Delete);
    R(this, "status", q.INIT);
    R(this, "value", !0);
    R(this, "id");
    R(this, "unitId");
    R(this, "subUnitId");
    R(this, "permissionId");
    this.unitId = e, this.subUnitId = t, this.permissionId = o, this.id = `${x.SelectRange}.${_.Delete}.${o}`;
  }
}
class ai {
  constructor(e, t, o) {
    R(this, "type", x.SelectRange);
    R(this, "subType", _.ManageCollaborator);
    R(this, "status", q.INIT);
    R(this, "value", !0);
    R(this, "id");
    R(this, "unitId");
    R(this, "subUnitId");
    R(this, "permissionId");
    this.unitId = e, this.subUnitId = t, this.permissionId = o, this.id = `${x.SelectRange}.${_.ManageCollaborator}.${o}`;
  }
}
const Ee = () => [Fo, Ve, ai, ii], pt = [_.Edit, _.View, _.ManageCollaborator, _.Delete], Yh = (n = "unitId", e = "subUnitId", t = "permissionId") => Ee().reduce((o, s) => {
  const r = new s(n, e, t);
  return o[r.subType] = r.value, o;
}, {}), gn = () => [
  Pe,
  lr,
  nr,
  hr,
  or,
  ur,
  jo,
  rr,
  ir,
  Go,
  Bo,
  ar,
  dr,
  zo,
  Gu,
  gr,
  cr,
  sr,
  Xu,
  Yu,
  Ju,
  Ku
], Jc = [
  _.Edit,
  _.Print,
  _.Comment,
  _.View,
  _.Copy,
  _.Export,
  _.ManageCollaborator,
  _.CreateSheet,
  _.DeleteSheet,
  _.RenameSheet,
  _.HideSheet,
  _.Duplicate,
  _.Share,
  _.MoveSheet,
  _.CopySheet,
  _.RecoverHistory,
  _.ViewHistory,
  _.CreatePermissionObject,
  _.InsertRow,
  _.InsertColumn,
  _.DeleteRow,
  _.DeleteColumn
], be = () => [
  Le,
  lo,
  Mr,
  Rr
], He = () => [
  mr,
  fr,
  Cr,
  pr,
  Sr,
  wr,
  vr,
  Ir,
  _r,
  yr,
  Dn,
  Zt,
  Qt,
  br
], po = [
  _.Copy,
  _.DeleteColumn,
  _.DeleteRow,
  _.EditExtraObject,
  _.Filter,
  _.InsertColumn,
  _.InsertRow,
  _.InsertHyperlink,
  _.PivotTable,
  _.SetCellStyle,
  _.SetCellValue,
  _.SetColumnStyle,
  _.SetRowStyle,
  _.Sort
];
var qc = Object.getOwnPropertyDescriptor, Yc = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? qc(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, ct = (n, e) => (t, o) => e(t, o, n);
const Xc = "SHEET_WORKSHEET_PROTECTION_PLUGIN", Zc = "SHEET_WORKSHEET_PROTECTION_POINT_PLUGIN";
let Bn = class extends $s {
  constructor(n, e, t, o, s, r, i, a) {
    super(), this._permissionService = n, this._univerInstanceService = e, this._injector = t, this._worksheetProtectionRuleModel = o, this._worksheetProtectionPointRuleModel = s, this._resourceManagerService = r, this._rangeProtectionRuleModel = i, this._logService = a, this._init(), this._initRuleChange(), this._initRuleSnapshot(), this._initPointSnapshot();
  }
  _init() {
    const n = (e) => {
      const t = e.getUnitId(), o = (s) => {
        const r = s.getSheetId();
        [...be(), ...He()].forEach((i) => {
          const a = new i(t, r);
          this._permissionService.addPermissionPoint(a);
        }), this._logService.debug("[WorksheetPermissionService]", "Initialization completed", t, r);
      };
      e.getSheets().forEach((s) => {
        o(s);
      }), e.sheetCreated$.subscribe((s) => {
        o(s);
      }), e.sheetDisposed$.subscribe((s) => {
        const r = s.getSheetId();
        this._rangeProtectionRuleModel.getSubunitRuleList(t, r).forEach((a) => {
          [...Ee()].forEach((u) => {
            const l = new u(t, r, a.permissionId);
            this._permissionService.deletePermissionPoint(l.id);
          });
        }), [...be(), ...He()].forEach((a) => {
          const u = new a(t, r);
          this._permissionService.deletePermissionPoint(u.id);
        });
      });
    };
    this._univerInstanceService.getAllUnitsForType(B.UNIVER_SHEET).forEach((e) => {
      n(e);
    }), this._univerInstanceService.getTypeOfUnitAdded$(B.UNIVER_SHEET).pipe(ss(this.dispose$)).subscribe(n), this._univerInstanceService.getTypeOfUnitDisposed$(B.UNIVER_SHEET).pipe(ss(this.dispose$)).subscribe((e) => {
      e.getSheets().forEach((t) => {
        const o = e.getUnitId(), s = t.getSheetId();
        be().forEach((r) => {
          const i = new r(o, s);
          this._permissionService.deletePermissionPoint(i.id);
        });
      });
    });
  }
  _initRuleChange() {
    this.disposeWithMe(
      this._worksheetProtectionRuleModel.ruleChange$.subscribe((n) => {
        switch (n.type) {
          case "add":
            break;
          case "delete": {
            be().forEach((e) => {
              const t = new e(n.unitId, n.subUnitId);
              this._permissionService.updatePermissionPoint(t.id, !0);
            });
            break;
          }
          case "set": {
            be().forEach((e) => {
              const t = new e(n.unitId, n.subUnitId);
              this._permissionService.updatePermissionPoint(t.id, n.rule);
            });
            break;
          }
        }
      })
    );
  }
  _initRuleSnapshot() {
    const n = () => {
      const t = this._worksheetProtectionRuleModel.toObject();
      return JSON.stringify(t);
    }, e = (t) => {
      if (!t)
        return {};
      try {
        return JSON.parse(t);
      } catch {
        return {};
      }
    };
    this.disposeWithMe(
      this._resourceManagerService.registerPluginResource({
        toJson: n,
        parseJson: e,
        pluginName: Xc,
        businesses: [On.UNIVER_SHEET],
        onLoad: (t, o) => {
          this._worksheetProtectionRuleModel.fromObject(o), Object.keys(o).forEach((s) => {
            be().forEach((r) => {
              const i = new r(t, s);
              i.value = !1, this._permissionService.addPermissionPoint(i);
            });
          }), this._worksheetProtectionRuleModel.changeRuleInitState(!0);
        },
        onUnLoad: (t) => {
          const o = this._univerInstanceService.getUnit(t);
          o && (o.getSheets().forEach((s) => {
            const r = s.getSheetId();
            [...be(), ...He()].forEach((i) => {
              const a = new i(t, r);
              this._permissionService.deletePermissionPoint(a.id);
            });
          }), gn().forEach((s) => {
            const r = new s(t);
            this._permissionService.deletePermissionPoint(r.id);
          })), this._worksheetProtectionRuleModel.deleteUnitModel(t);
        }
      })
    );
  }
  _initPointSnapshot() {
    const n = () => {
      const t = this._worksheetProtectionPointRuleModel.toObject();
      return JSON.stringify(t);
    }, e = (t) => {
      if (!t)
        return {};
      try {
        return JSON.parse(t);
      } catch {
        return {};
      }
    };
    this.disposeWithMe(
      this._resourceManagerService.registerPluginResource({
        toJson: n,
        parseJson: e,
        pluginName: Zc,
        businesses: [On.UNIVER_SHEET],
        onLoad: (t, o) => {
          this._worksheetProtectionPointRuleModel.fromObject(o), Object.keys(o).forEach((s) => {
            He().forEach((r) => {
              const i = new r(t, s);
              this._permissionService.addPermissionPoint(i);
            });
          });
        },
        onUnLoad: (t) => {
          this._worksheetProtectionPointRuleModel.deleteUnitModel(t);
        }
      })
    );
  }
};
Bn = Yc([
  ct(0, $(Ze)),
  ct(1, $(M)),
  ct(2, $(so)),
  ct(3, $(lt)),
  ct(4, $(In)),
  ct(5, $(Rn)),
  ct(6, $(me)),
  ct(7, $(Fs))
], Bn);
const es = {
  id: "sheet.mutation.set-worksheet-permission-points",
  type: v.MUTATION,
  handler: (n, e) => {
    const { rule: t } = e;
    return n.get(In).addRule(t), !0;
  }
}, Qc = {
  type: v.COMMAND,
  id: "sheet.command.set-worksheet-permission-points",
  async handler(n, e) {
    if (!e)
      return !1;
    const t = n.get(E), { rule: o } = e;
    return t.executeCommand(es.id, {
      rule: o,
      unitId: o.unitId,
      subUnitId: o.subUnitId
    }), !0;
  }
}, ed = {
  type: v.COMMAND,
  id: "sheet.command.set-worksheet-protection",
  async handler(n, e) {
    if (!e)
      return !1;
    const t = n.get(E), o = n.get(V), { rule: s, permissionId: r, oldRule: i } = e, { unitId: a, subUnitId: u } = s, l = { ...s, permissionId: r };
    if (await t.executeCommand(bt.id, {
      unitId: a,
      subUnitId: u,
      newRule: l
    })) {
      const d = [{ id: bt.id, params: { unitId: a, subUnitId: u, newRule: l } }], h = [{ id: bt.id, params: { unitId: a, subUnitId: u, rule: i } }];
      o.pushUndoRedo({
        unitID: a,
        redoMutations: d,
        undoMutations: h
      });
    }
    return !0;
  }
}, td = (n, e) => {
  const r = n.get(M).getUniverSheetInstance(e.unitId).getSheetBySheetId(e.subUnitId).getConfig().rightToLeft;
  return {
    ...A.deepClone(e),
    rightToLeft: r
  };
}, So = {
  id: "sheet.mutation.set-worksheet-right-to-left",
  type: v.MUTATION,
  handler: (n, e) => {
    const t = n.get(M).getUniverSheetInstance(e.unitId);
    if (!t) return !1;
    const o = t.getSheetBySheetId(e.subUnitId);
    if (!o) return !1;
    const s = o.getConfig();
    return s.rightToLeft = e.rightToLeft, !0;
  }
}, Xh = {
  type: v.COMMAND,
  id: "sheet.command.set-worksheet-right-to-left",
  handler: async (n, e) => {
    var d;
    const t = n.get(E), o = n.get(V), s = P(n.get(M), e);
    if (!s) return !1;
    const { unitId: r, subUnitId: i } = s;
    let a = te.FALSE;
    e && (a = (d = e.rightToLeft) != null ? d : te.FALSE);
    const u = {
      rightToLeft: a,
      unitId: r,
      subUnitId: i
    }, l = td(
      n,
      u
    );
    return t.syncExecuteCommand(
      So.id,
      u
    ) ? (o.pushUndoRedo({
      unitID: r,
      undoMutations: [{ id: So.id, params: l }],
      redoMutations: [
        { id: So.id, params: u }
      ]
    }), !0) : !1;
  }
}, nd = {
  type: v.COMMAND,
  id: "sheet.command.set-worksheet-row-count",
  handler: (n, e) => {
    const { unitId: t, subUnitId: o, rowCount: s } = e, r = n.get(E), i = n.get(V), a = n.get(M);
    if (!P(a, e)) return !1;
    const l = {
      unitId: t,
      subUnitId: o,
      rowCount: s
    }, c = ru(n, l);
    return r.syncExecuteCommand(Yt.id, l) ? (i.pushUndoRedo({
      unitID: t,
      undoMutations: [{ id: Yt.id, params: c }],
      redoMutations: [{ id: Yt.id, params: l }]
    }), !0) : !1;
  }
}, jn = {
  type: v.COMMAND,
  id: "sheet.command.delta-row-height",
  // eslint-disable-next-line max-lines-per-function
  handler: async (n, e) => {
    var T, U;
    const o = n.get(z).getCurrentSelections(), s = n.get(G);
    if (!(o != null && o.length))
      return !1;
    const r = P(n.get(M));
    if (!r) return !1;
    const { worksheet: i, subUnitId: a, unitId: u } = r, { anchorRow: l, deltaY: c } = e, h = i.getRowHeight(l) + c, g = o.length === 1 && o[0].range.rangeType === j.ALL, m = o.filter((D) => D.range.rangeType === j.ROW), f = g ? j.ALL : m.some(({ range: D }) => {
      const { startRow: H, endRow: F } = D;
      return H <= l && l <= F;
    }) ? j.ROW : j.NORMAL;
    let C;
    if (f === j.ALL) {
      const D = i.getColumnCount(), H = new Array(i.getRowCount()).fill(void 0).map(
        (F, K) => ({ startRow: K, endRow: K, startColumn: 0, endColumn: D - 1 })
      );
      C = {
        subUnitId: a,
        unitId: u,
        rowHeight: h,
        ranges: H
      };
    } else f === j.ROW ? C = {
      subUnitId: a,
      unitId: u,
      ranges: m.map((D) => N.clone(D.range)),
      rowHeight: h
    } : C = {
      subUnitId: a,
      unitId: u,
      rowHeight: h,
      ranges: [
        {
          startRow: l,
          endRow: l,
          startColumn: 0,
          endColumn: i.getMaxColumns() - 1
        }
      ]
    };
    const p = Ys(C, i), w = {
      unitId: u,
      subUnitId: a,
      ranges: C.ranges,
      autoHeightInfo: te.FALSE
    }, I = Wo(w, i), S = n.get(E), y = n.get(V), b = s.onCommandExecute({
      id: jn.id,
      params: C
    }), O = L([
      {
        id: nt.id,
        params: C
      },
      {
        id: Fe.id,
        params: w
      }
    ], S), k = L([...b.redos], S);
    if (O.result && k.result) {
      const D = s.afterCommandExecute({
        id: jn.id,
        params: C
      });
      return L(D.redos, S), y.pushUndoRedo({
        unitID: u,
        undoMutations: [
          ...(T = b.preUndos) != null ? T : [],
          {
            id: nt.id,
            params: p
          },
          {
            id: Fe.id,
            params: I
          },
          ...b.undos,
          ...D.undos
        ],
        redoMutations: [
          ...(U = b.preRedos) != null ? U : [],
          {
            id: nt.id,
            params: C
          },
          {
            id: Fe.id,
            params: w
          },
          ...b.redos,
          ...D.redos
        ]
      }), !0;
    }
    return !1;
  }
}, zn = {
  type: v.COMMAND,
  id: "sheet.command.set-row-height",
  // eslint-disable-next-line max-lines-per-function
  handler: (n, e) => {
    var I, S, y, b;
    const t = n.get(z), o = n.get(E), s = n.get(V), r = n.get(M), i = n.get(G), a = (I = e == null ? void 0 : e.ranges) != null && I.length ? e.ranges : (S = t.getCurrentSelections()) == null ? void 0 : S.map((O) => O.range);
    if (!(a != null && a.length))
      return !1;
    const u = P(r, e);
    if (!u) return !1;
    const { unitId: l, subUnitId: c, worksheet: d } = u, h = {
      subUnitId: c,
      unitId: l,
      ranges: a,
      rowHeight: e.value
    }, g = Ys(h, d), m = {
      unitId: l,
      subUnitId: c,
      ranges: h.ranges,
      autoHeightInfo: te.FALSE
    }, f = Wo(m, d), C = L([
      {
        id: nt.id,
        params: h
      },
      {
        id: Fe.id,
        params: m
      }
    ], o), p = i.onCommandExecute({
      id: zn.id,
      params: h
    }), w = L([...p.redos], o);
    if (C.result && w.result) {
      const O = i.afterCommandExecute({
        id: zn.id,
        params: h
      });
      return L(O.redos, o), s.pushUndoRedo({
        unitID: l,
        undoMutations: [
          ...(y = p.preRedos) != null ? y : [],
          {
            id: nt.id,
            params: g
          },
          {
            id: Fe.id,
            params: f
          },
          ...p.undos,
          ...O.undos
        ],
        redoMutations: [
          ...(b = p.preRedos) != null ? b : [],
          {
            id: nt.id,
            params: h
          },
          {
            id: Fe.id,
            params: m
          },
          ...p.redos,
          ...O.redos
        ]
      }), !0;
    }
    return !1;
  }
}, ts = {
  type: v.COMMAND,
  id: "sheet.command.set-row-is-auto-height",
  handler: (n, e) => {
    var O, k;
    const t = n.get(E), o = n.get(V), s = n.get(z), r = n.get(M), i = P(r, e);
    if (!i) return !1;
    const { unitId: a, subUnitId: u, worksheet: l } = i, c = (O = e == null ? void 0 : e.ranges) != null && O.length ? e.ranges : (k = s.getCurrentSelections()) == null ? void 0 : k.map((T) => T.range);
    if (!(c != null && c.length))
      return !1;
    const d = {
      unitId: a,
      subUnitId: u,
      ranges: c,
      autoHeightInfo: te.TRUE
      // Hard code first, maybe it will change by the menu item in the future.
    }, h = Wo(d, l), g = t.syncExecuteCommand(
      Fe.id,
      d
    ), m = n.get(rt).getSkeleton(a, u), { suitableRanges: f, remainingRanges: C } = Sn(d.ranges, m), p = n.get(G), { undos: w, redos: I } = p.generateMutationsOfAutoHeight({
      unitId: a,
      subUnitId: u,
      ranges: f,
      autoHeightRanges: f,
      lazyAutoHeightRanges: C
    }), { undos: S, redos: y } = p.onCommandExecute({
      id: ts.id,
      params: d
    }), b = L([...y, ...I], t);
    return g && b.result ? (o.pushUndoRedo({
      unitID: a,
      undoMutations: [{ id: Fe.id, params: h }, ...S, ...w],
      redoMutations: [{ id: Fe.id, params: d }, ...y, ...I]
    }), !0) : !1;
  }
}, ui = {
  type: v.COMMAND,
  id: "sheet.command.set-worksheet-show",
  handler: (n, e) => {
    const { unitId: t, subUnitId: o } = e, s = n.get(E), r = n.get(V), i = n.get(M);
    if (!P(n.get(M))) return !1;
    const u = i.getCurrentUnitForType(B.UNIVER_SHEET);
    if (!u) return !1;
    const l = u.getSheetBySheetId(o);
    if (!l || l.getConfig().hidden === te.FALSE) return !1;
    const d = {
      unitId: t,
      subUnitId: o,
      hidden: te.FALSE
    }, h = si(n, d), g = s.syncExecuteCommand(St.id, d), m = {
      unitId: t,
      subUnitId: o
    }, f = s.syncExecuteCommand(
      Cn.id,
      m
    );
    return g && f ? (r.pushUndoRedo({
      unitID: t,
      undoMutations: [
        { id: St.id, params: h }
        // { id: SetWorksheetActiveOperation.id, params: unActiveMutationParams },
      ],
      redoMutations: [
        // { id: SetWorksheetActiveOperation.id, params: activeSheetMutationParams },
        { id: St.id, params: d }
      ]
    }), !0) : !1;
  }
}, od = {
  type: v.COMMAND,
  id: "sheet.command.split-text-to-columns",
  // eslint-disable-next-line max-lines-per-function
  handler: (n, e) => {
    const { unitId: t, subUnitId: o, range: s, delimiter: r, customDelimiter: i, treatMultipleDelimitersAsOne: a } = e, u = n.get(E), l = n.get(M), c = n.get(V);
    if (!P(n.get(M))) return !1;
    const h = l.getCurrentUnitForType(B.UNIVER_SHEET);
    if (!h) return !1;
    const g = h.getSheetBySheetId(o);
    if (!g) return !1;
    const { lastRow: m, rs: f, maxLength: C } = Mu(g, s, r, i, a), p = g.getColumnCount(), { startColumn: w } = X.transformRange(s, g);
    if (s.startColumn !== s.endColumn)
      return !1;
    const I = [], S = [], y = w + C + 1 - p;
    if (y > 0) {
      const D = {
        unitId: t,
        subUnitId: o,
        range: {
          startRow: 0,
          endRow: g.getRowCount() - 1,
          startColumn: p - 1,
          endColumn: p - 1 + y
        }
      };
      I.push({
        id: ke.id,
        params: D
      });
      const H = ro(
        n,
        D
      );
      S.push({ id: Me.id, params: H });
    }
    const b = {
      startRow: s.startRow,
      endRow: m,
      startColumn: w,
      endColumn: w + C
    }, O = new Y();
    for (let D = b.startRow; D <= b.endRow; D++)
      for (let H = b.startColumn; H <= b.endColumn; H++) {
        const F = f[D - b.startRow];
        H === 0 && (F == null ? void 0 : F.length) === 1 ? O.setValue(D, H, g.getCell(D, H)) : O.setValue(D, H, {
          v: (F == null ? void 0 : F[H - b.startColumn]) || null,
          p: null,
          f: null,
          si: null,
          custom: null
        });
      }
    const k = {
      unitId: t,
      subUnitId: o,
      cellValue: O.clone()
    }, T = Ae(n, k);
    return I.push({
      id: ee.id,
      params: k
    }), S.unshift({
      id: ee.id,
      params: T
    }), L(I, u).result ? (c.pushUndoRedo({
      unitID: t,
      undoMutations: S,
      redoMutations: I
    }), !0) : !1;
  }
}, sd = {
  id: "sheet.command.toggle-cell-checkbox",
  type: v.COMMAND,
  handler: (n, e) => {
    if (!e)
      return !1;
    const { unitId: t, subUnitId: o, row: s, col: r, paragraphIndex: i } = e, u = n.get(M).getUnit(t, B.UNIVER_SHEET), l = u == null ? void 0 : u.getSheetBySheetId(o), c = n.get(V), d = n.get(E);
    if (!l)
      return !1;
    const h = l.getCell(s, r);
    if (!(h != null && h.p))
      return !1;
    const g = A.deepClone(h.p), m = new Pi(g), f = Ni.paragraph.bullet.toggleChecklist({
      document: m,
      paragraphIndex: i
    });
    if (!f)
      return !1;
    Oi.apply(m.getBody(), f.serialize());
    const C = {
      unitId: t,
      subUnitId: o,
      cellValue: {
        [s]: {
          [r]: {
            p: g,
            t: re.STRING
          }
        }
      }
    }, p = {
      id: ee.id,
      params: C
    }, w = Ae(n, C), I = {
      id: ee.id,
      params: w
    }, S = [p], y = [I];
    return c.pushUndoRedo({
      redoMutations: S,
      undoMutations: y,
      unitID: t
    }), d.syncExecuteCommand(p.id, p.params);
  }
}, rd = {
  type: v.COMMAND,
  id: "sheet.command.toggle-gridlines",
  handler: (n, e) => {
    const t = n.get(E), o = n.get(V), s = n.get(M), r = P(s);
    if (!r) return !1;
    const { worksheet: i } = r, a = i.getConfig().showGridlines;
    if (a === (e == null ? void 0 : e.showGridlines)) return !1;
    const { unitId: u, subUnitId: l } = r, c = {
      showGridlines: a === te.TRUE ? te.FALSE : te.TRUE,
      unitId: u,
      subUnitId: l
    }, d = {
      showGridlines: a,
      unitId: u,
      subUnitId: l
    };
    return t.syncExecuteCommand(Xt.id, c) ? (o.pushUndoRedo({
      unitID: u,
      undoMutations: [{ id: Xt.id, params: d }],
      redoMutations: [{ id: Xt.id, params: c }]
    }), !0) : !1;
  }
}, id = {
  id: "sheet.command.unregister-worksheet-range-theme-style",
  type: v.COMMAND,
  handler: (n, e) => {
    var h;
    if (!e)
      return !1;
    const { unitId: t, themeName: o } = e, s = n.get(M), r = n.get(E), i = n.get(V), a = n.get(ve);
    if (!P(s)) return !1;
    const l = {
      unitId: t,
      themeName: o
    }, c = {
      unitId: t,
      themeName: o,
      rangeThemeStyleJson: (h = a.getRangeThemeStyle(t, o)) == null ? void 0 : h.toJson()
    };
    return r.syncExecuteCommand(cn.id, e) && i.pushUndoRedo({
      unitID: t,
      undoMutations: [{ id: cn.id, params: c }],
      redoMutations: [{ id: Yo.id, params: l }]
    }), !0;
  }
}, ad = {
  id: "sheet.mutation.add-range-theme",
  type: v.MUTATION,
  handler: (n, e) => {
    if (!e)
      return !1;
    const { styleJSON: t, unitId: o } = e, s = n.get(ve), r = new Rt(t.name);
    return r.fromJson(t), s.registerRangeThemeStyle(o, r), !0;
  }
}, ud = {
  id: "sheet.mutation.empty",
  type: v.MUTATION,
  handler: () => !0
}, ld = {
  id: "sheet.operation.mark-dirty-row-auto-height",
  type: v.OPERATION,
  handler: () => !0
}, cd = {
  id: "sheet.operation.cancel-mark-dirty-row-auto-height",
  type: v.OPERATION,
  handler: () => !0
}, $t = Ao("INumfmtService"), Zh = (n, e) => {
  const t = n.get($t), { values: o, unitId: s, subUnitId: r } = e, i = [], a = [];
  Object.keys(o).forEach((l) => {
    o[l].ranges.forEach((d) => {
      X.foreach(d, (h, g) => {
        const m = t.getValue(s, r, h, g);
        m ? i.push({
          pattern: m.pattern,
          row: h,
          col: g
        }) : a.push({ startColumn: g, endColumn: g, startRow: h, endRow: h });
      });
    });
  });
  const u = [];
  if (i.length) {
    const l = ko(s, r, i);
    Object.keys(l.values).forEach((c) => {
      const d = l.values[c];
      d.ranges = Qs(d.ranges);
    }), u.push({
      id: ns.id,
      params: ko(s, r, i)
    });
  }
  return a.length && u.push({
    id: li.id,
    params: {
      unitId: s,
      subUnitId: r,
      ranges: a
    }
  }), u;
}, ns = {
  id: "sheet.mutation.set.numfmt",
  type: v.MUTATION,
  handler: (n, e) => {
    if (!e)
      return !1;
    const { values: t, refMap: o } = e, s = n.get($t), r = e.unitId, i = e.subUnitId, a = Object.keys(t).reduce(
      (u, l) => {
        const c = o[l], d = t[l].ranges;
        return c && u.push({
          ...c,
          ranges: d
        }), u;
      },
      []
    );
    return s.setValues(r, i, a), !0;
  }
}, li = {
  id: "sheet.mutation.remove.numfmt",
  type: v.MUTATION,
  handler: (n, e) => {
    if (!e)
      return !1;
    const { unitId: t, subUnitId: o, ranges: s } = e;
    return n.get($t).deleteValues(t, o, s), !0;
  }
}, Qh = (n, e) => {
  const t = n.get($t), { ranges: o, unitId: s, subUnitId: r } = e, i = [];
  if (o.forEach((u) => {
    X.foreach(u, (l, c) => {
      const d = t.getValue(s, r, l, c);
      d && i.push({
        pattern: d.pattern,
        row: l,
        col: c
      });
    });
  }), !i.length)
    return [];
  const a = ko(s, r, i);
  return Object.keys(a.values).forEach((u) => {
    const l = a.values[u];
    l.ranges = Qs(l.ranges);
  }), [{ id: ns.id, params: a }];
}, ko = (n, e, t) => {
  const o = _u(t, "pattern"), s = {}, r = {}, i = yu();
  return Object.keys(o).forEach((a) => {
    const u = o[a], l = i();
    s[l] = {
      pattern: a
    }, u.forEach((c) => {
      r[l] || (r[l] = { ranges: [] }), r[l].ranges.push(Vs(c.row, c.col));
    });
  }), { unitId: n, subUnitId: e, refMap: s, values: r };
}, dd = {
  id: "sheet.mutation.remove-range-theme",
  type: v.MUTATION,
  handler: (n, e) => {
    if (!e)
      return !1;
    const { styleName: t, unitId: o } = e;
    return n.get(ve).unregisterRangeThemeStyle(o, t), !0;
  }
}, hd = {
  id: "sheet.mutation.set-range-theme",
  type: v.MUTATION,
  handler: (n, e) => {
    if (!e)
      return !1;
    const { unitId: t, styleName: o, style: s } = e, i = n.get(ve).getRangeThemeStyle(t, o);
    return i && (s.headerRowStyle && i.setHeaderRowStyle(s.headerRowStyle), s.firstRowStyle && i.setFirstRowStyle(s.firstRowStyle), s.secondRowStyle && i.setSecondRowStyle(s.secondRowStyle), s.lastRowStyle && i.setLastRowStyle(s.lastRowStyle)), !0;
  }
}, gd = {
  id: "sheet.operation.scroll-to-cell",
  type: v.OPERATION,
  handler: () => !0
}, eg = (n, e, t) => {
  const s = n.get(z).getCurrentSelections(), { value: r, selections: i, unitId: a, subUnitId: u } = e;
  if (s) {
    const c = s[(s == null ? void 0 : s.length) - 1].primary;
    if (c) {
      const { actualColumn: d, actualRow: h } = c;
      let { startRow: g, startColumn: m, endRow: f, endColumn: C } = i[i.length - 1];
      if (r === Re.COLUMNS) {
        const S = t.find((y) => y.startColumn === d && y.endColumn === d && h === y.startRow);
        S && (C = S.endColumn, g = S.startRow, f = S.endRow);
      } else if (r === Re.ROWS) {
        const S = t.find((y) => y.startRow === h && y.endRow === h && d === y.startColumn);
        S && (f = S.endRow, m = S.startColumn, C = S.endColumn);
      }
      const p = {
        startRow: g,
        startColumn: m,
        endRow: f,
        endColumn: C,
        actualRow: h,
        actualColumn: d,
        isMerged: !0,
        isMergedMainCell: g === h && m === d
      }, w = s.map((S, y, b) => ({
        range: S.range,
        style: null,
        primary: y === b.length - 1 ? p : null
      })), I = {
        unitId: a,
        subUnitId: u,
        type: we.ONLY_SET,
        selections: w
      };
      return {
        id: ie.id,
        params: I
      };
    }
    return null;
  }
  return null;
}, tg = (n, e) => {
  const o = n.get(z).getCurrentSelections(), { unitId: s, subUnitId: r } = e;
  if (o && o[(o == null ? void 0 : o.length) - 1].primary) {
    const u = {
      unitId: s,
      subUnitId: r,
      type: we.ONLY_SET,
      selections: [...o]
    };
    return {
      id: ie.id,
      params: u
    };
  }
  return null;
}, md = "maxCellsPerSheet", fd = 3e6;
var Rd = Object.getOwnPropertyDescriptor, Cd = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? Rd(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, Ms = (n, e) => (t, o) => e(t, o, n);
const pd = "SHEET_DEFINED_NAME_PLUGIN", ng = "AllDefaultWorkbook";
let Gn = class extends ue {
  constructor(n, e) {
    super(), this._definedNamesService = n, this._resourceManagerService = e, this._initialize();
  }
  _initialize() {
    this._initSnapshot();
  }
  _initSnapshot() {
    const n = (t) => {
      const o = this._definedNamesService.getDefinedNameMap(t);
      return o ? JSON.stringify(o) : "";
    }, e = (t) => {
      if (!t)
        return {};
      try {
        return JSON.parse(t);
      } catch {
        return {};
      }
    };
    this.disposeWithMe(
      this._resourceManagerService.registerPluginResource({
        pluginName: pd,
        businesses: [B.UNIVER_SHEET],
        toJson: (t) => n(t),
        parseJson: (t) => e(t),
        onUnLoad: (t) => {
          this._definedNamesService.removeUnitDefinedName(t);
        },
        onLoad: (t, o) => {
          this._definedNamesService.registerDefinedNames(t, o);
        }
      })
    );
  }
};
Gn = Cd([
  Ms(0, js),
  Ms(1, Rn)
], Gn);
const ci = "sheets.config", _s = {};
var Sd = Object.getOwnPropertyDescriptor, wd = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? Sd(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, wo = (n, e) => (t, o) => e(t, o, n);
const Id = [
  ot.id
], vd = [
  Ue.id,
  ke.id,
  Te.id,
  Me.id,
  Je.id,
  qe.id
];
let Kn = class extends ue {
  constructor(e, t, o) {
    var r, i;
    super();
    R(this, "_d", new Et());
    R(this, "_enabled", !0);
    this._univerInstanceService = e, this._commandService = t, this._configService = o;
    const s = (i = (r = this._configService.getConfig(ci)) == null ? void 0 : r.freezeSync) != null ? i : !0;
    this.setEnabled(s);
  }
  getEnabled() {
    return this._enabled;
  }
  setEnabled(e) {
    e ? this._d.dispose() : this._initOnlyLocalListener(), this._enabled = e;
  }
  _initOnlyLocalListener() {
    this._d.add(
      this._commandService.beforeCommandExecuted((e, t) => {
        Id.includes(e.id) && (t || (t = {}), t.onlyLocal = !0);
      })
    ), this._d.add(
      this._commandService.onCommandExecuted((e, t) => {
        if (vd.includes(e.id) && (t != null && t.fromCollab)) {
          const { id: o, params: s } = e;
          o === Ue.id ? this._handleInsertRowMutation(s, t) : o === ke.id ? this._handleInsertColMutation(s, t) : o === Te.id ? this._handleRemoveRowMutation(s, t) : o === Me.id ? this._handleRemoveColMutation(s, t) : o === Je.id ? this._handleMoveRowsMutation(s, t) : o === qe.id && this._handleMoveColsMutation(s, t);
        }
      })
    );
  }
  _handleInsertRowMutation(e, t) {
    const { range: o, unitId: s, subUnitId: r } = e, i = this._getFreeze(s, r);
    if (i && o.startRow < i.startRow) {
      const a = o.endRow - o.startRow + 1, u = {
        ...i,
        startRow: Math.max(1, i.startRow + a),
        ySplit: Math.max(1, i.ySplit + a)
      };
      this._sequenceExecute(s, r, u, t);
    }
  }
  _handleInsertColMutation(e, t) {
    const { range: o, unitId: s, subUnitId: r } = e, i = this._getFreeze(s, r);
    if (i && o.startColumn < i.startColumn) {
      const a = o.endColumn - o.startColumn + 1, u = {
        ...i,
        startColumn: Math.max(1, i.startColumn + a),
        xSplit: Math.max(1, i.xSplit + a)
      };
      this._sequenceExecute(s, r, u, t);
    }
  }
  _handleRemoveRowMutation(e, t) {
    const { range: o, unitId: s, subUnitId: r } = e, i = this._getFreeze(s, r);
    if (i && o.startRow < i.startRow) {
      const a = Math.min(i.startRow, o.endRow + 1) - o.startRow, u = {
        ...i,
        startRow: Math.max(1, i.startRow - a),
        ySplit: Math.max(1, i.ySplit - a)
      };
      this._sequenceExecute(s, r, u, t);
    }
  }
  _handleRemoveColMutation(e, t) {
    const { range: o, unitId: s, subUnitId: r } = e, i = this._getFreeze(s, r);
    if (i && o.startColumn < i.startColumn) {
      const a = Math.min(i.startColumn, o.endColumn + 1) - o.startColumn, u = {
        ...i,
        startColumn: Math.max(1, i.startColumn - a),
        xSplit: Math.max(1, i.xSplit - a)
      };
      this._sequenceExecute(s, r, u, t);
    }
  }
  _handleMoveRowsMutation(e, t) {
    const { sourceRange: o, targetRange: s, unitId: r, subUnitId: i } = e, a = this._getFreeze(r, i);
    if (!a || a.startRow <= 0 || o.startRow >= a.startRow && s.startRow >= a.startRow || o.endRow < a.startRow && s.endRow < a.startRow)
      return;
    const u = o.endRow - o.startRow + 1, l = Math.max(
      Math.min(a.startRow, o.endRow + 1) - o.startRow,
      0
    ), c = { ...a };
    s.startRow >= a.startRow ? (c.startRow = Math.max(1, a.startRow - l), c.ySplit = Math.max(1, a.ySplit - l)) : (c.startRow = a.startRow + u - l, c.ySplit = a.ySplit + u - l), this._sequenceExecute(r, i, c, t);
  }
  _handleMoveColsMutation(e, t) {
    const { sourceRange: o, targetRange: s, unitId: r, subUnitId: i } = e, a = this._getFreeze(r, i);
    if (!a || a.startColumn <= 0 || o.startColumn >= a.startColumn && s.startColumn >= a.startColumn || o.endColumn < a.startColumn && s.endColumn < a.startColumn)
      return;
    const u = o.endColumn - o.startColumn + 1, l = Math.max(
      Math.min(a.startColumn, o.endColumn + 1) - o.startColumn,
      0
    ), c = { ...a };
    s.startColumn >= a.startColumn ? (c.startColumn = Math.max(1, a.startColumn - l), c.xSplit = Math.max(1, a.xSplit - l)) : (c.startColumn = a.startColumn + u - l, c.xSplit = a.xSplit + u - l), this._sequenceExecute(r, i, c, t);
  }
  _getFreeze(e, t) {
    const o = this._univerInstanceService.getUnit(e, B.UNIVER_SHEET);
    if (!o) return null;
    const s = o.getSheetBySheetId(t);
    return s ? s.getFreeze() : null;
  }
  _sequenceExecute(e, t, o, s) {
    L([
      {
        id: ot.id,
        params: {
          ...o,
          unitId: e,
          subUnitId: t,
          resetScroll: !1
        }
      }
    ], this._commandService, s);
  }
};
Kn = wd([
  wo(0, $(M)),
  wo(1, E),
  wo(2, xo)
], Kn);
var Md = Object.getOwnPropertyDescriptor, _d = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? Md(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, ze = (n, e) => (t, o) => e(t, o, n);
let Jn = class extends ue {
  constructor(e, t, o, s, r, i, a, u, l, c) {
    super();
    R(this, "disposableCollection", new Et());
    R(this, "_triggerPermissionUIEvent$", new De());
    R(this, "triggerPermissionUIEvent$", this._triggerPermissionUIEvent$.asObservable());
    this._commandService = e, this._univerInstanceService = t, this._permissionService = o, this._selectionManagerService = s, this._rangeProtectionRuleModel = r, this._worksheetProtectionRuleModel = i, this._localeService = a, this._lexerTreeBuilder = u, this._contextService = l, this._definedNamesService = c, this._initialize();
  }
  blockExecuteWithoutPermission(e) {
    throw this._triggerPermissionUIEvent$.next(e), new Di("have no permission");
  }
  _getPermissionCheck(e, t) {
    let o = !0, s = "";
    switch (e) {
      case pn.id:
        Ls(t.value) && t.value.f ? (o = this._permissionCheckWithFormula(t), s = this._localeService.t("permission.dialog.formulaErr")) : o = this._permissionCheckBySetRangeValue({
          workbookTypes: [Pe],
          rangeTypes: [Ve],
          worksheetTypes: [Dn, Le]
        }, t);
        break;
      case qo.id:
        o = this.permissionCheckWithRanges({
          workbookTypes: [Pe],
          rangeTypes: [Ve],
          worksheetTypes: [Dn, Le]
        }), s = this._localeService.t("permission.dialog.editErr");
        break;
      case Ln.id:
      case Hn.id:
        o = this.permissionCheckWithoutRange({
          worksheetTypes: [Zt]
        }), s = this._localeService.t("permission.dialog.setRowColStyleErr");
        break;
      case jn.id:
      case zn.id:
      case ts.id:
        o = this.permissionCheckWithoutRange({
          worksheetTypes: [Qt]
        }), s = this._localeService.t("permission.dialog.setRowColStyleErr");
        break;
      case un.id:
      case an.id:
        o = this._permissionCheckByMoveCommand(t), s = this._localeService.t("permission.dialog.moveRowColErr");
        break;
      case yt.id:
        o = this._permissionCheckByMoveRangeCommand(t), s = this._localeService.t("permission.dialog.moveRangeErr");
        break;
      case ri.id:
        o = this._permissionCheckByWorksheetCommand([Pe, zo]), s = this._localeService.t("permission.dialog.operatorSheetErr"), o === !1 && this._worksheetProtectionRuleModel.resetOrder();
        break;
      case Qo.id:
        o = this._permissionCheckByWorksheetCommand([Pe, Go]), s = this._localeService.t("permission.dialog.operatorSheetErr"), o === !1 && this._worksheetProtectionRuleModel.resetOrder();
        break;
      case ui.id:
        {
          const { unitId: r, subUnitId: i } = t;
          o = this._permissionCheckByWorksheetCommand([Pe, Bo], r, i), s = this._localeService.t("permission.dialog.operatorSheetErr"), o === !1 && this._worksheetProtectionRuleModel.resetOrder();
        }
        break;
      case dn.id:
        o = this.permissionCheckWithRanges({
          workbookTypes: [Pe],
          rangeTypes: [Ve],
          worksheetTypes: [Le, Zt]
        }, t.ranges), s = this._localeService.t("permission.dialog.setRowColStyleErr");
        break;
      case hn.id:
        o = this.permissionCheckWithRanges({
          workbookTypes: [Pe],
          rangeTypes: [Ve],
          worksheetTypes: [Le, Qt]
        }, t.ranges), s = this._localeService.t("permission.dialog.setRowColStyleErr");
        break;
      case qr.id:
        o = this.permissionCheckWithRanges({
          workbookTypes: [Pe],
          rangeTypes: [Ve],
          worksheetTypes: [Le, Zt]
        }), s = this._localeService.t("permission.dialog.setRowColStyleErr");
        break;
      case Qr.id:
        o = this.permissionCheckWithRanges({
          workbookTypes: [Pe],
          rangeTypes: [Ve],
          worksheetTypes: [Le, Qt]
        }), s = this._localeService.t("permission.dialog.setRowColStyleErr");
        break;
      case rn.id:
        o = this._permissionCheckWithInsertRangeMove("right"), s = this._localeService.t("permission.dialog.insertOrDeleteMoveRangeErr");
        break;
      case Pt.id:
        o = this._permissionCheckWithInsertRangeMove("bottom"), s = this._localeService.t("permission.dialog.insertOrDeleteMoveRangeErr");
        break;
      case It.id:
        o = this._permissionCheckWithInsertRangeMove("left"), s = this._localeService.t("permission.dialog.insertOrDeleteMoveRangeErr");
        break;
      case vt.id:
        o = this._permissionCheckWithInsertRangeMove("top"), s = this._localeService.t("permission.dialog.insertOrDeleteMoveRangeErr");
        break;
    }
    o || this.blockExecuteWithoutPermission(s);
  }
  _initialize() {
    this._commandExecutedListener();
  }
  _commandExecutedListener() {
    this.disposeWithMe(
      this._commandService.beforeCommandExecuted((e) => {
        this._getPermissionCheck(e.id, e == null ? void 0 : e.params);
      })
    ), this.disposeWithMe(
      this._commandService.onCommandExecuted((e) => {
        var t;
        if (e.id === Fn.id) {
          const o = e.params, { unitId: s = (t = this._univerInstanceService.getCurrentUnitForType(B.UNIVER_SHEET)) == null ? void 0 : t.getUnitId(), subUnitId: r } = o;
          if (!s || !r)
            return;
          const i = this._worksheetProtectionRuleModel.getRule(s, r), a = this._rangeProtectionRuleModel.getSubunitRuleList(s, r);
          i && this._worksheetProtectionRuleModel.ruleRefresh(i.permissionId), a.length && this._rangeProtectionRuleModel.ruleRefresh(r);
        }
      })
    );
  }
  _permissionCheckWithInsertRangeMove(e) {
    var l;
    const t = P(this._univerInstanceService);
    if (!t)
      return !1;
    const { worksheet: o, unitId: s, subUnitId: r } = t, i = A.deepClone((l = this._selectionManagerService.getCurrentLastSelection()) == null ? void 0 : l.range);
    return !(!i || (e === "top" || e === "bottom" ? i.endRow = o.getRowCount() - 1 : (e === "left" || e === "right") && (i.endColumn = o.getColumnCount() - 1), this._rangeProtectionRuleModel.getSubunitRuleList(s, r).map((c) => c.ranges).flat().some((c) => N.getIntersects(i, c))));
  }
  _permissionCheckByWorksheetCommand(e, t, o) {
    var l, c;
    const s = P(this._univerInstanceService, { unitId: t, subUnitId: o });
    if (!s)
      return !1;
    const { unitId: r, subUnitId: i } = s, a = this._worksheetProtectionRuleModel.getRule(r, i), u = this._rangeProtectionRuleModel.getSubunitRuleList(r, i).length > 0;
    return a || u ? (c = (l = this._permissionService.getPermissionPoint(new jo(r).id)) == null ? void 0 : l.value) != null ? c : !1 : this._permissionService.composePermission(e.map((d) => new d(r).id)).every((d) => d.value);
  }
  permissionCheckWithoutRange(e) {
    var h, g, m, f;
    const t = P(this._univerInstanceService);
    if (!t)
      return !1;
    const { worksheet: o, unitId: s, subUnitId: r } = t, i = this._selectionManagerService.getCurrentLastSelection();
    if (!i)
      return !0;
    const a = (g = (h = i == null ? void 0 : i.primary) == null ? void 0 : h.actualRow) != null ? g : 0, u = (f = (m = i == null ? void 0 : i.primary) == null ? void 0 : m.actualColumn) != null ? f : 0, { workbookTypes: l, worksheetTypes: c, rangeTypes: d } = e;
    return !(l && l.some((p) => {
      var S, y;
      const w = new p(s);
      return ((y = (S = this._permissionService.getPermissionPoint(w.id)) == null ? void 0 : S.value) != null ? y : !1) === !1;
    }) === !0 || c && c.some((p) => {
      var S, y;
      const w = new p(s, r);
      return ((y = (S = this._permissionService.getPermissionPoint(w.id)) == null ? void 0 : S.value) != null ? y : !1) === !1;
    }) === !0 || d && d.some((p) => {
      var b, O, k, T, U;
      const w = (O = (b = o.getCell(a, u)) == null ? void 0 : b.selectionProtection) == null ? void 0 : O[0];
      if (!(w != null && w.ruleId))
        return !1;
      const I = (k = this._rangeProtectionRuleModel.getRule(s, r, w.ruleId)) == null ? void 0 : k.permissionId;
      if (!I)
        return !1;
      const S = new p(s, r, I);
      return ((U = (T = this._permissionService.getPermissionPoint(S.id)) == null ? void 0 : T.value) != null ? U : !1) === !1;
    }) === !0);
  }
  permissionCheckWithRanges(e, t, o, s) {
    var g;
    const r = P(this._univerInstanceService);
    if (!r)
      return !1;
    const { workbook: i, worksheet: a } = r;
    o || (o = i.getUnitId()), s || (s = a.getSheetId());
    const u = t != null ? t : (g = this._selectionManagerService.getCurrentSelections()) == null ? void 0 : g.map((m) => m.range);
    if (!u)
      return !1;
    const { workbookTypes: l, worksheetTypes: c, rangeTypes: d } = e, h = [];
    return l && h.push(...l.map((m) => new m(o).id)), c && h.push(...c.map((m) => new m(o, s).id)), d && this._rangeProtectionRuleModel.getSubunitRuleList(o, s).forEach((m) => {
      u.some((C) => m.ranges.some((p) => N.intersects(p, C))) && h.push(...d.map((C) => new C(o, s, m.permissionId).id));
    }), h.length ? this._permissionService.composePermission(h).every((m) => m.value) : !0;
  }
  _permissionCheckByMoveCommand(e) {
    const t = P(this._univerInstanceService);
    if (!t)
      return !1;
    const { worksheet: o, unitId: s, subUnitId: r } = t, i = e.toRange;
    i.endRow === o.getRowCount() - 1 ? i.endColumn = i.startColumn : i.endRow = i.startRow;
    const a = this._rangeProtectionRuleModel.getSubunitRuleList(s, r).reduce((u, l) => [...u, ...l.ranges], []).filter((u) => N.intersects(u, i));
    return a.length > 0 ? !1 : (a.forEach((u) => {
      var l, c;
      for (let d = u.startRow; d <= u.endRow; d++)
        for (let h = u.startColumn; h <= u.endColumn; h++) {
          const g = (c = (l = o.getCell(d, h)) == null ? void 0 : l.selectionProtection) == null ? void 0 : c[0];
          if ((g == null ? void 0 : g[_.Edit]) === !1)
            return !1;
        }
    }), !0);
  }
  _permissionCheckByMoveRangeCommand(e) {
    const t = P(this._univerInstanceService);
    if (!t)
      return !1;
    const { worksheet: o, unitId: s, subUnitId: r } = t, i = e.toRange, a = this._rangeProtectionRuleModel.getSubunitRuleList(s, r).reduce((u, l) => [...u, ...l.ranges], []).filter((u) => N.intersects(u, i));
    return a.length > 0 ? !1 : (a.forEach((u) => {
      var l, c;
      for (let d = u.startRow; d <= u.endRow; d++)
        for (let h = u.startColumn; h <= u.endColumn; h++) {
          const g = (c = (l = o.getCell(d, h)) == null ? void 0 : l.selectionProtection) == null ? void 0 : c[0];
          if ((g == null ? void 0 : g[_.Edit]) === !1)
            return !1;
        }
    }), !0);
  }
  _permissionCheckBySetRangeValue(e, t) {
    let o = [];
    t.range ? o = [t.range] : o = [new Y(t.value).getDataRange()];
    const { unitId: s, subUnitId: r } = t;
    return this.permissionCheckWithRanges(e, o, s, r);
  }
  _permissionCheckWithFormula(e) {
    var r, i, a, u, l;
    const t = e.value, o = e.range, s = t.f;
    if (s) {
      const c = s.substring(1), d = this._univerInstanceService.getCurrentUnitForType(B.UNIVER_SHEET), h = (r = e.unitId) != null ? r : d.getUnitId(), g = this._definedNamesService.getValueByName(h, c);
      if (g) {
        let m = g.formulaOrRefString;
        m.startsWith(Zi.EQUALS) && (m = m.slice(1));
        const f = m.split(",");
        for (let C = 0; C < f.length; C++) {
          const p = f[C], w = Qi(p);
          if (w.sheetName) {
            const I = d.getSheetBySheetName(w.sheetName);
            if (!I)
              return !0;
            const { startRow: S, endRow: y, startColumn: b, endColumn: O } = w.range;
            for (let k = S; k <= y; k++)
              for (let T = b; T <= O; T++) {
                const U = (a = (i = I.getCell(k, T)) == null ? void 0 : i.selectionProtection) == null ? void 0 : a[0];
                if ((U == null ? void 0 : U[_.View]) === !1)
                  return !1;
              }
          }
        }
        return !0;
      } else {
        const m = this._lexerTreeBuilder.sequenceNodesBuilder(s);
        if (!m)
          return !0;
        for (let f = 0; f < m.length; f++) {
          const C = m[f];
          if (typeof C == "string" || C.nodeType !== ea.REFERENCE)
            continue;
          const { token: p } = C, w = ta(p), I = w.unitId ? this._univerInstanceService.getUnit(w.unitId) : this._univerInstanceService.getCurrentUnitForType(B.UNIVER_SHEET);
          if (!I) return !0;
          let S = w.sheetName ? I.getSheetBySheetName(w.sheetName) : I.getActiveSheet();
          const y = I.getUnitId();
          if (w.sheetName) {
            if (S = I.getSheetBySheetName(w.sheetName), !S)
              return !0;
            const U = S == null ? void 0 : S.getSheetId();
            if (!this._permissionService.getPermissionPoint(new lo(y, U).id)) return !1;
          }
          if (!S)
            return !0;
          const { startRow: b, endRow: O, startColumn: k, endColumn: T } = w.range;
          for (let U = b; U <= O; U++)
            for (let D = k; D <= T; D++) {
              const H = (l = (u = S.getCell(U, D)) == null ? void 0 : u.selectionProtection) == null ? void 0 : l[0];
              if ((H == null ? void 0 : H[_.View]) === !1)
                return !1;
            }
        }
        return !0;
      }
    }
    if (o) {
      const c = P(this._univerInstanceService);
      if (!c)
        return !1;
      const d = e.unitId || c.unitId, h = e.subUnitId || c.subUnitId, m = this._rangeProtectionRuleModel.getSubunitRuleList(d, h).filter((C) => C.ranges.some((p) => N.intersects(p, o))).map((C) => new Ve(d, h, C.permissionId).id);
      if (!this._permissionService.composePermission(m).every((C) => C.value))
        return !1;
    }
    return !0;
  }
};
Jn = _d([
  ze(0, E),
  ze(1, M),
  ze(2, Ze),
  ze(3, $(z)),
  ze(4, $(me)),
  ze(5, $(lt)),
  ze(6, $(it)),
  ze(7, $(Xi)),
  ze(8, Ws),
  ze(9, js)
], Jn);
var yd = Object.getOwnPropertyDescriptor, bd = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? yd(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, Ft = (n, e) => (t, o) => e(t, o, n);
let mn = class extends ue {
  constructor(e, t, o, s, r) {
    super();
    R(this, "_unitPermissionInitStateChange", new gt(!1));
    R(this, "unitPermissionInitStateChange$", this._unitPermissionInitStateChange.asObservable());
    this._permissionService = e, this._univerInstanceService = t, this._rangeProtectionRuleModel = o, this._worksheetProtectionRuleModel = s, this._worksheetProtectionPointModel = r, this._init();
  }
  _init() {
    const e = (t) => {
      const o = t.getUnitId();
      gn().forEach((s) => {
        const r = new s(o);
        this._permissionService.addPermissionPoint(r);
      });
    };
    this._univerInstanceService.getAllUnitsForType(B.UNIVER_SHEET).forEach((t) => {
      e(t);
    }), this.disposeWithMe(this._univerInstanceService.getTypeOfUnitAdded$(B.UNIVER_SHEET).subscribe((t) => {
      e(t);
    })), this.disposeWithMe(this._univerInstanceService.getTypeOfUnitDisposed$(B.UNIVER_SHEET).subscribe((t) => {
      const o = t.getUnitId();
      t.getSheets().forEach((s) => {
        const r = s.getSheetId();
        this._rangeProtectionRuleModel.getSubunitRuleList(o, r).forEach((a) => {
          [...Ee()].forEach((u) => {
            const l = new u(o, r, a.permissionId);
            this._permissionService.deletePermissionPoint(l.id);
          });
        }), [...be(), ...He()].forEach((a) => {
          const u = new a(o, r);
          this._permissionService.deletePermissionPoint(u.id);
        });
      }), gn().forEach((s) => {
        const r = new s(o);
        this._permissionService.deletePermissionPoint(r.id);
      }), this._rangeProtectionRuleModel.deleteUnitModel(o), this._worksheetProtectionPointModel.deleteUnitModel(o), this._worksheetProtectionRuleModel.deleteUnitModel(o);
    }));
  }
  changeUnitInitState(e) {
    this._unitPermissionInitStateChange.next(e);
  }
};
mn = bd([
  Ft(0, $(Ze)),
  Ft(1, $(M)),
  Ft(2, $(me)),
  Ft(3, $(lt)),
  Ft(4, $(In))
], mn);
var Ed = Object.getOwnPropertyDescriptor, Ud = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? Ed(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, Ge = (n, e) => (t, o) => e(t, o, n);
let qn = class extends ue {
  constructor(n, e, t, o, s, r, i, a, u, l) {
    super(), this._univerInstanceService = n, this._permissionService = e, this._authzIoService = t, this._rangeProtectionRuleModel = o, this._worksheetProtectionRuleModel = s, this._userManagerService = r, this._worksheetProtectionPointRuleModel = i, this._workbookPermissionService = a, this._undoRedoService = u, this._commandService = l;
  }
  initPermission() {
    this._initRangePermissionFromSnapshot(), this._initRangePermissionChange(), this._initWorksheetPermissionFromSnapshot(), this._initWorksheetPermissionChange(), this._initWorksheetPermissionPointsChange(), this._initWorkbookPermissionFromSnapshot(), this._initUserChange(), this._refreshPermissionByCollaCreate();
  }
  refreshRangeProtectPermission() {
    this._initRangePermissionFromSnapshot();
  }
  async _initRangePermissionFromSnapshot() {
    const n = async (e) => {
      const t = [], o = e.getUnitId(), s = e.getSheets(), r = /* @__PURE__ */ new Map();
      if (s.forEach((i) => {
        const a = i.getSheetId();
        this._rangeProtectionRuleModel.getSubunitRuleList(o, a).forEach((u) => {
          r.set(u.permissionId, u), t.push({
            objectID: u.permissionId,
            unitID: o,
            objectType: x.SelectRange,
            actions: pt
          });
        });
      }), !t.length) {
        this._rangeProtectionRuleModel.changeRuleInitState(!0);
        return;
      }
      this._authzIoService.batchAllowed(t).then((i) => {
        i.forEach((a) => {
          const u = r.get(a.objectID);
          u && Ee().forEach((l) => {
            const c = new l(o, u.subUnitId, a.objectID), d = c.subType, h = a.actions.find((g) => g.action === d);
            (h == null ? void 0 : h.allowed) !== void 0 && this._permissionService.updatePermissionPoint(c.id, h.allowed);
          });
        }), this._rangeProtectionRuleModel.changeRuleInitState(!0);
      });
    };
    await Promise.all(this._univerInstanceService.getAllUnitsForType(B.UNIVER_SHEET).map((e) => n(e))), this._rangeProtectionRuleModel.changeRuleInitState(!0);
  }
  _initRangePermissionChange() {
    this.disposeWithMe(
      this._rangeProtectionRuleModel.ruleChange$.subscribe((n) => {
        n.type !== "delete" ? this._authzIoService.allowed({
          objectID: n.rule.permissionId,
          unitID: n.unitId,
          objectType: x.SelectRange,
          actions: pt
        }).then((e) => {
          Ee().forEach((t) => {
            if (n.type === "set") {
              const { rule: a, oldRule: u } = n;
              if (a.permissionId === (u == null ? void 0 : u.permissionId))
                return;
            }
            const o = n.rule, s = new t(o.unitId, o.subUnitId, o.permissionId), r = s.subType, i = e.find((a) => a.action === r);
            i && this._permissionService.updatePermissionPoint(s.id, i.allowed);
          }), this._rangeProtectionRuleModel.ruleRefresh(n.rule.permissionId);
        }) : this._rangeProtectionRuleModel.getSubunitRuleList(n.unitId, n.subUnitId).length === 0 && (this._worksheetProtectionPointRuleModel.deleteRule(n.unitId, n.subUnitId), [...He()].forEach((t) => {
          const o = new t(n.unitId, n.subUnitId);
          this._permissionService.updatePermissionPoint(o.id, o.value);
        }));
      })
    );
  }
  async initWorkbookPermissionChange(n) {
    var t;
    const e = n || ((t = this._univerInstanceService.getCurrentUnitForType(B.UNIVER_SHEET)) == null ? void 0 : t.getUnitId());
    if (e)
      return this._authzIoService.allowed({
        objectID: e,
        objectType: x.Workbook,
        unitID: e,
        actions: Jc
      }).then((o) => {
        gn().forEach((s) => {
          const r = new s(e), i = r.subType, a = o.find((u) => u.action === i);
          a && this._permissionService.updatePermissionPoint(r.id, a.allowed);
        });
      });
  }
  async _initWorkbookPermissionFromSnapshot() {
    await Promise.all(this._univerInstanceService.getAllUnitsForType(B.UNIVER_SHEET).map((n) => this.initWorkbookPermissionChange(n.getUnitId()))), this._workbookPermissionService.changeUnitInitState(!0);
  }
  _initWorksheetPermissionChange() {
    this.disposeWithMe(
      this._worksheetProtectionRuleModel.ruleChange$.subscribe((n) => {
        n.type !== "delete" ? this._authzIoService.allowed({
          objectID: n.rule.permissionId,
          unitID: n.unitId,
          objectType: x.Worksheet,
          actions: pt
        }).then((e) => {
          be().forEach((t) => {
            const o = new t(n.unitId, n.subUnitId), s = o.subType, r = e.find((i) => i.action === s);
            r && this._permissionService.updatePermissionPoint(o.id, r.allowed);
          }), this._worksheetProtectionRuleModel.ruleRefresh(n.rule.permissionId);
        }) : ([...be(), ...He()].forEach((e) => {
          const t = new e(n.unitId, n.subUnitId);
          this._permissionService.updatePermissionPoint(t.id, !0);
        }), this._worksheetProtectionPointRuleModel.deleteRule(n.unitId, n.subUnitId));
      })
    );
  }
  _initWorksheetPermissionPointsChange() {
    this.disposeWithMe(
      this._worksheetProtectionPointRuleModel.pointChange$.subscribe((n) => {
        this._authzIoService.allowed({
          objectID: n.permissionId,
          unitID: n.unitId,
          objectType: x.Worksheet,
          actions: po
        }).then((e) => {
          He().forEach((t) => {
            const o = new t(n.unitId, n.subUnitId), s = o.subType, r = e.find((i) => i.action === s);
            r && this._permissionService.updatePermissionPoint(o.id, r.allowed);
          });
        });
      })
    );
  }
  async _initWorksheetPermissionFromSnapshot() {
    const n = async (e) => {
      const t = [], o = e.getUnitId(), s = e.getSheets(), r = /* @__PURE__ */ new Map();
      if (s.forEach((i) => {
        const a = i.getSheetId(), u = this._worksheetProtectionRuleModel.getRule(o, a);
        u && (r.set(u.permissionId, u), t.push({
          objectID: u.permissionId,
          unitID: o,
          objectType: x.Worksheet,
          actions: pt
        }));
        const l = this._worksheetProtectionPointRuleModel.getRule(o, a);
        l && (r.set(l.permissionId, l), t.push({
          objectID: l.permissionId,
          unitID: o,
          objectType: x.Worksheet,
          actions: po
        }));
      }), !t.length) {
        this._worksheetProtectionRuleModel.changeRuleInitState(!0);
        return;
      }
      this._authzIoService.batchAllowed(t).then((i) => {
        i.forEach((a) => {
          const u = r.get(a.objectID);
          u && [...be(), ...He()].forEach((l) => {
            const c = new l(o, u.subUnitId), d = c.subType, h = a.actions.find((g) => g.action === d);
            (h == null ? void 0 : h.allowed) !== void 0 && this._permissionService.updatePermissionPoint(c.id, h.allowed);
          });
        }), this._worksheetProtectionRuleModel.changeRuleInitState(!0);
      });
    };
    await Promise.all(this._univerInstanceService.getAllUnitsForType(B.UNIVER_SHEET).map((e) => n(e))), this._worksheetProtectionRuleModel.changeRuleInitState(!0);
  }
  _initUserChange() {
    this.disposeWithMe(
      // When the user changes, the permission points are updated. The first modification needs to be filtered here because it is a Behavior type, but in fact the user information is ready when this controller is initialized.
      this._userManagerService.currentUser$.pipe(Bs(1)).subscribe(() => {
        const n = this._permissionService.getAllPermissionPoint();
        this._permissionService.clearPermissionMap(), this._worksheetProtectionRuleModel.changeRuleInitState(!1), this._univerInstanceService.getAllUnitsForType(B.UNIVER_SHEET).forEach((t) => {
          const o = t.getUnitId();
          gn().forEach((s) => {
            let r = new s(o);
            n.has(r.id) && (r = n.get(r.id)), this._permissionService.addPermissionPoint(r);
          }), t.getSheets().forEach((s) => {
            const r = s.getSheetId();
            [...be(), ...He()].forEach((a) => {
              let u = new a(o, r);
              n.has(u.id) && (u = n.get(u.id)), this._permissionService.addPermissionPoint(u);
            }), this._rangeProtectionRuleModel.getSubunitRuleList(o, r).forEach((a) => {
              Ee().forEach((u) => {
                let l = new u(o, r, a.permissionId);
                n.has(l.id) && (l = n.get(l.id)), this._permissionService.addPermissionPoint(l);
              });
            });
          }), this._initWorkbookPermissionFromSnapshot(), this._initWorksheetPermissionFromSnapshot(), this._initRangePermissionFromSnapshot();
        });
      })
    );
  }
  refreshPermission(n, e) {
    const t = this._worksheetProtectionRuleModel.getTargetByPermissionId(n, e);
    let o = !1;
    if (t) {
      const [i, a] = t;
      this._authzIoService.allowed({
        objectID: e,
        unitID: n,
        objectType: x.Worksheet,
        actions: pt
      }).then((u) => {
        let l = "";
        be().forEach((c) => {
          var m;
          const d = new c(n, a), h = d.subType, g = u.find((f) => f.action === h);
          g && (((m = this._permissionService.getPermissionPoint(d.id)) == null ? void 0 : m.value) !== g.allowed && (o = !0), this._permissionService.updatePermissionPoint(d.id, g.allowed), l += `${g.action}_${g.allowed}`);
        }), this._worksheetProtectionRuleModel.ruleRefresh(`${e}_${l}`), o && this._undoRedoService.clearUndoRedo(n);
      });
    }
    const s = this._worksheetProtectionPointRuleModel.getTargetByPermissionId(n, e);
    if (s) {
      const [i, a] = s;
      this._authzIoService.allowed({
        objectID: e,
        unitID: n,
        objectType: x.Worksheet,
        actions: po
      }).then((u) => {
        He().forEach((l) => {
          var g;
          const c = new l(n, a), d = c.subType, h = u.find((m) => m.action === d);
          h && (((g = this._permissionService.getPermissionPoint(c.id)) == null ? void 0 : g.value) !== h.allowed && (o = !0), this._permissionService.updatePermissionPoint(c.id, h.allowed));
        }), o && this._undoRedoService.clearUndoRedo(n);
      });
    }
    const r = this._rangeProtectionRuleModel.getTargetByPermissionId(n, e);
    if (r) {
      const [i, a] = r;
      this._authzIoService.allowed({
        objectID: e,
        unitID: n,
        objectType: x.SelectRange,
        actions: pt
      }).then((u) => {
        let l = "";
        Ee().forEach((c) => {
          var m;
          const d = new c(n, a, e), h = d.subType, g = u.find((f) => f.action === h);
          g && (((m = this._permissionService.getPermissionPoint(d.id)) == null ? void 0 : m.value) !== g.allowed && (o = !0), this._permissionService.updatePermissionPoint(d.id, g.allowed), l += `${g.action}_${g.allowed}`);
        }), this._rangeProtectionRuleModel.ruleRefresh(`${e}_${l}`), o && this._undoRedoService.clearUndoRedo(n);
      });
    }
  }
  _refreshPermissionByCollaCreate() {
    this.disposeWithMe(
      this._commandService.onCommandExecuted((n, e) => {
        if (e != null && e.fromCollab && (n.id === Be.id || n.id === Mt.id || n.id === es.id)) {
          const t = n.params;
          this._undoRedoService.clearUndoRedo(t.unitId);
        }
      })
    );
  }
};
qn = Ud([
  Ge(0, M),
  Ge(1, Ze),
  Ge(2, Ai),
  Ge(3, $(me)),
  Ge(4, $(lt)),
  Ge(5, $(xi)),
  Ge(6, $(In)),
  Ge(7, $(mn)),
  Ge(8, $(V)),
  Ge(9, $(E))
], qn);
var kd = Object.getOwnPropertyDescriptor, Td = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? kd(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, Io = (n, e) => (t, o) => e(t, o, n);
let Yn = class extends ue {
  constructor(e, t, o) {
    super();
    R(this, "_zebraCacheUpdateSubject", new De());
    this._commandService = e, this._sheetRangeThemeModel = t, this._univerInstanceService = o, this._init();
  }
  _init() {
    this._initializeCommandListener(), this._initTriggerCacheUpdateListener();
  }
  /**
   * Update the zebra crossing cache for a specific unit and sub-unit.
   * @param {string} unitId - The ID of the unit.
   * @param {string} subUnitId - The ID of the sub-unit.
   */
  updateZebraCrossingCache(e, t) {
    this._zebraCacheUpdateSubject.next({ unitId: e, subUnitId: t });
  }
  _initializeCommandListener() {
    this.disposeWithMe(this._commandService.onCommandExecuted((e) => {
      const { id: t } = e;
      let o, s;
      switch (t) {
        case Ue.id:
          {
            const r = e.params;
            o = r.unitId, s = r.subUnitId;
          }
          break;
        case kt.id:
          {
            const r = e.params;
            o = r.unitId, s = r.subUnitId;
          }
          break;
        case Tt.id:
          {
            const r = e.params;
            o = r.unitId, s = r.subUnitId;
          }
          break;
        case Te.id:
          {
            const r = e.params;
            o = r.unitId, s = r.subUnitId;
          }
          break;
        case nt.id:
          {
            const r = e.params;
            o = r.unitId, s = r.subUnitId;
          }
          break;
      }
      o && s && (this._sheetRangeThemeModel.refreshSheetRowVisibleFuncSet(o, s), this._sheetRangeThemeModel.refreshZebraCrossingCacheBySheet(o, s));
    }));
  }
  _initTriggerCacheUpdateListener() {
    this.disposeWithMe(
      this._zebraCacheUpdateSubject.subscribe(({ unitId: e, subUnitId: t }) => {
        this._sheetRangeThemeModel.refreshSheetRowVisibleFuncSet(e, t), this._sheetRangeThemeModel.refreshZebraCrossingCacheBySheet(e, t);
      })
    );
  }
};
Yn = Td([
  Io(0, $(E)),
  Io(1, $(ve)),
  Io(2, $(M))
], Yn);
var Pd = Object.getOwnPropertyDescriptor, Nd = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? Pd(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, ys = (n, e) => (t, o) => e(t, o, n);
let fn = class {
  constructor(n, e) {
    R(this, "_cache", new $i(1e4));
    this._selectionProtectionRuleModel = n, this._permissionService = e, this._init();
  }
  _init() {
    this._permissionService.permissionPointUpdate$.pipe(
      rs((n) => n.type === x.SelectRange),
      rs((n) => Ee().some((e) => n instanceof e)),
      ia((n) => n)
    ).subscribe((n) => {
      const e = this._selectionProtectionRuleModel.getSubunitRuleList(n.unitId, n.subUnitId);
      for (const t of e)
        t.permissionId === n.permissionId && t.ranges.forEach((o) => {
          X.foreach(o, (s, r) => {
            const i = this._createKey(n.unitId, n.subUnitId, s, r);
            this._cache.delete(i);
          });
        });
    }), this._selectionProtectionRuleModel.ruleChange$.subscribe((n) => {
      var e;
      n.rule.ranges.forEach((t) => {
        X.foreach(t, (o, s) => {
          const r = this._createKey(n.unitId, n.subUnitId, o, s);
          this._cache.delete(r);
        });
      }), n.type === "set" && ((e = n.oldRule) == null || e.ranges.forEach((t) => {
        X.foreach(t, (o, s) => {
          const r = this._createKey(n.unitId, n.subUnitId, o, s);
          this._cache.delete(r);
        });
      }));
    });
  }
  _createKey(n, e, t, o) {
    return `${n}_${e}_${t}_${o}`;
  }
  getCellInfo(n, e, t, o) {
    const s = this._selectionProtectionRuleModel.getSubunitRuleList(n, e), r = [];
    if (!s || !s.length)
      return r;
    const i = this._createKey(n, e, t, o), a = this._cache.get(i);
    if (a)
      return a;
    const u = [];
    for (const l of s)
      if (l.ranges.some((c) => c.startRow <= t && c.endRow >= t && c.startColumn <= o && c.endColumn >= o)) {
        const c = Ee().reduce((d, h) => {
          var f;
          const g = new h(n, e, l.permissionId), m = this._permissionService.getPermissionPoint(g.id);
          return d[g.subType] = (f = m == null ? void 0 : m.value) != null ? f : g.value, d;
        }, {});
        u.push({ ...c, ruleId: l.id, ranges: l.ranges });
      }
    return this._cache.set(i, u), u;
  }
  clear() {
    this._cache.clear();
  }
};
fn = Nd([
  ys(0, $(me)),
  ys(1, $(Ze))
], fn);
var Od = Object.getOwnPropertyDescriptor, Dd = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? Od(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, vo = (n, e) => (t, o) => e(t, o, n);
let Wt = class extends ue {
  constructor(e, t, o) {
    super();
    R(this, "_cellRuleCache", /* @__PURE__ */ new Map());
    R(this, "_permissionIdCache", /* @__PURE__ */ new Map());
    R(this, "_cellInfoCache", /* @__PURE__ */ new Map());
    //  {unitId:{subUnitId:{[row/col]:{permissionId1:{edit:true},permissionId2:{edit:true},permissionId3:{edit:false}}}}}
    R(this, "_rowInfoCache", /* @__PURE__ */ new Map());
    R(this, "_colInfoCache", /* @__PURE__ */ new Map());
    this._ruleModel = e, this._permissionService = t, this._univerInstanceService = o, this._initUpdateCellRuleCache(), this._initUpdateCellInfoCache(), this._initUpdateRowColInfoCache(), this._initCache();
  }
  _initCache() {
    this._univerInstanceService.getAllUnitsForType(B.UNIVER_SHEET).forEach((e) => {
      e.getSheets().forEach((t) => {
        const o = e.getUnitId(), s = t.getSheetId();
        this.reBuildCache(o, s);
      });
    });
  }
  _initUpdateCellInfoCache() {
    this._permissionService.permissionPointUpdate$.pipe(
      os((e) => e.type === x.SelectRange),
      yo((e) => e)
    ).subscribe((e) => {
      const { subUnitId: t, unitId: o, permissionId: s } = e, r = this._permissionIdCache.get(s);
      if (!r)
        return;
      const i = this._ruleModel.getRule(o, t, r);
      if (!i)
        return;
      const a = this._ensureCellInfoMap(o, t);
      i.ranges.forEach((u) => {
        const { startRow: l, endRow: c, startColumn: d, endColumn: h } = u;
        for (let g = l; g <= c; g++)
          for (let m = d; m <= h; m++)
            a.delete(`${g}-${m}`);
      });
    }), this._ruleModel.ruleChange$.subscribe((e) => {
      var r;
      const { unitId: t, subUnitId: o } = e, s = this._ensureCellInfoMap(t, o);
      e.rule.ranges.forEach((i) => {
        X.foreach(i, (a, u) => {
          s.delete(`${a}-${u}`);
        });
      }), e.type === "set" && ((r = e.oldRule) == null || r.ranges.forEach((i) => {
        X.foreach(i, (a, u) => {
          this._cellInfoCache.delete(`${a}-${u}`);
        });
      }));
    });
  }
  _initUpdateCellRuleCache() {
    this._ruleModel.ruleChange$.subscribe((e) => {
      const { type: t } = e;
      t === "add" ? this._addCellRuleCache(e) : t === "delete" ? this._deleteCellRuleCache(e) : (this._deleteCellRuleCache({ ...e, rule: e.oldRule }), this._addCellRuleCache(e));
    });
  }
  _ensureRuleMap(e, t) {
    let o = this._cellRuleCache.get(e);
    o || (o = /* @__PURE__ */ new Map(), this._cellRuleCache.set(e, o));
    let s = o.get(t);
    return s || (s = /* @__PURE__ */ new Map(), o.set(t, s)), s;
  }
  _ensureCellInfoMap(e, t) {
    let o = this._cellInfoCache.get(e);
    o || (o = /* @__PURE__ */ new Map(), this._cellInfoCache.set(e, o));
    let s = o.get(t);
    return s || (s = /* @__PURE__ */ new Map(), o.set(t, s)), s;
  }
  _ensureRowColInfoMap(e, t, o) {
    let s = o === "row" ? this._rowInfoCache.get(e) : this._colInfoCache.get(e);
    s || (s = /* @__PURE__ */ new Map(), o === "row" ? this._rowInfoCache.set(e, s) : this._colInfoCache.set(e, s));
    let r = s.get(t);
    return r || (r = /* @__PURE__ */ new Map(), s.set(t, r)), r;
  }
  _addCellRuleCache(e) {
    const { subUnitId: t, unitId: o, rule: s } = e, r = this._ensureRuleMap(o, t);
    s.ranges.forEach((i) => {
      const { startRow: a, endRow: u, startColumn: l, endColumn: c } = i;
      for (let d = a; d <= u; d++)
        for (let h = l; h <= c; h++)
          r.set(`${d}-${h}`, s.id);
    }), this._permissionIdCache.set(s.permissionId, s.id);
  }
  _deleteCellRuleCache(e) {
    const { subUnitId: t, unitId: o, rule: s } = e, r = this._ensureRuleMap(o, t), i = this._ensureCellInfoMap(o, t);
    s.ranges.forEach((a) => {
      const { startRow: u, endRow: l, startColumn: c, endColumn: d } = a;
      for (let h = u; h <= l; h++)
        for (let g = c; g <= d; g++)
          r.delete(`${h}-${g}`), i.delete(`${h}-${g}`);
    }), this._permissionIdCache.delete(s.permissionId);
  }
  _getSelectionActions(e, t, o) {
    var l, c, d, h, g, m, f, C, p, w, I, S;
    const s = (d = (c = this._permissionService.getPermissionPoint((l = new Ve(e, t, o.permissionId)) == null ? void 0 : l.id)) == null ? void 0 : c.value) != null ? d : !0, r = (m = (g = this._permissionService.getPermissionPoint((h = new Fo(e, t, o.permissionId)) == null ? void 0 : h.id)) == null ? void 0 : g.value) != null ? m : !0, i = (p = (C = this._permissionService.getPermissionPoint((f = new ai(e, t, o.permissionId)) == null ? void 0 : f.id)) == null ? void 0 : C.value) != null ? p : !1, a = (S = (I = this._permissionService.getPermissionPoint((w = new ii(e, t, o.permissionId)) == null ? void 0 : w.id)) == null ? void 0 : I.value) != null ? S : !1;
    return {
      [_.Edit]: s,
      [_.View]: r,
      [_.ManageCollaborator]: i,
      [_.Delete]: a
    };
  }
  reBuildCache(e, t) {
    const o = this._ensureRuleMap(e, t), s = this._ensureCellInfoMap(e, t);
    o.clear(), s.clear();
    const r = this._ensureRowColInfoMap(e, t, "row"), i = this._ensureRowColInfoMap(e, t, "col");
    r.clear(), i.clear(), this._ruleModel.getSubunitRuleList(e, t).forEach((a) => {
      const u = this._getSelectionActions(e, t, a), l = {
        ...u,
        ruleId: a.id,
        ranges: a.ranges
      };
      a.ranges.forEach((c) => {
        const { startRow: d, endRow: h, startColumn: g, endColumn: m } = c;
        for (let f = d; f <= h; f++) {
          const C = r.get(`${f}`);
          C ? C.set(a.id, u) : r.set(`${f}`, /* @__PURE__ */ new Map([[a.id, u]]));
          for (let p = g; p <= m; p++) {
            o.set(`${f}-${p}`, a.id), s.set(`${f}-${p}`, l);
            const w = i.get(`${p}`);
            w ? w.set(a.id, u) : i.set(`${p}`, /* @__PURE__ */ new Map([[a.id, u]]));
          }
        }
      }), this._permissionIdCache.set(a.permissionId, a.id);
    });
  }
  getRowPermissionInfo(e, t, o, s) {
    var a;
    const r = (a = this._rowInfoCache.get(e)) == null ? void 0 : a.get(t);
    if (!r)
      return !0;
    const i = r.get(`${o}`);
    return i ? s.every((u) => {
      for (const l of i.values())
        if (l[u] === !1)
          return !1;
      return !0;
    }) : !0;
  }
  getColPermissionInfo(e, t, o, s) {
    var a;
    const r = (a = this._colInfoCache.get(e)) == null ? void 0 : a.get(t);
    if (!r)
      return !0;
    const i = r.get(`${o}`);
    return i ? s.every((u) => {
      for (const l of i.values())
        if (l[u] === !1)
          return !1;
      return !0;
    }) : !0;
  }
  _initUpdateRowColInfoCache() {
    this._permissionService.permissionPointUpdate$.pipe(
      os((e) => e.type === x.SelectRange),
      yo((e) => e)
    ).subscribe({
      next: (e) => {
        const { subUnitId: t, unitId: o, permissionId: s } = e, r = this._permissionIdCache.get(s);
        if (!r)
          return;
        const i = this._ruleModel.getRule(o, t, r);
        if (!i)
          return;
        const a = this._ensureRowColInfoMap(o, t, "row"), u = this._ensureRowColInfoMap(o, t, "col"), l = this._getSelectionActions(o, t, i);
        i.ranges.forEach((c) => {
          const { startRow: d, endRow: h, startColumn: g, endColumn: m } = c;
          for (let f = d; f <= h; f++) {
            const C = a.get(`${f}`);
            C ? C.set(r, l) : a.set(`${f}`, /* @__PURE__ */ new Map([[r, l]]));
            for (let p = g; p <= m; p++) {
              const w = u.get(`${p}`);
              w ? w.set(r, l) : u.set(`${p}`, /* @__PURE__ */ new Map([[r, l]]));
            }
          }
        });
      }
    }), this._ruleModel.ruleChange$.subscribe((e) => {
      if (e.type === "delete") {
        const { unitId: t, subUnitId: o, rule: s } = e, r = this._ensureRowColInfoMap(t, o, "row"), i = this._ensureRowColInfoMap(t, o, "col");
        s.ranges.forEach((a) => {
          const { startRow: u, endRow: l, startColumn: c, endColumn: d } = a;
          for (let h = u; h <= l; h++) {
            const g = r.get(`${h}`);
            g == null || g.delete(s.id);
            for (let m = c; m <= d; m++) {
              const f = i.get(`${m}`);
              f == null || f.delete(s.id);
            }
          }
        });
      }
    });
  }
  getCellInfo(e, t, o, s) {
    var l, c;
    const r = this._ensureCellInfoMap(e, t), i = r.get(`${o}-${s}`);
    if (i)
      return i;
    const a = (c = (l = this._cellRuleCache.get(e)) == null ? void 0 : l.get(t)) == null ? void 0 : c.get(`${o}-${s}`);
    if (!a)
      return;
    const u = this._ruleModel.getRule(e, t, a);
    if (u) {
      const h = {
        ...this._getSelectionActions(e, t, u),
        ruleId: a,
        ranges: u.ranges
      };
      return r.set(`${o}-${s}`, h), h;
    }
  }
  deleteUnit(e) {
    this._cellRuleCache.delete(e), this._cellInfoCache.delete(e), this._rowInfoCache.delete(e), this._colInfoCache.delete(e);
    const t = this._univerInstanceService.getUnit(e);
    t == null || t.getSheets().forEach((o) => {
      const s = o.getSheetId();
      this._ruleModel.getSubunitRuleList(e, s).forEach((r) => {
        this._permissionIdCache.delete(r.permissionId);
      });
    });
  }
};
Wt = Dd([
  vo(0, $(me)),
  vo(1, $(Ze)),
  vo(2, $(M))
], Wt);
const di = "ONLY_REGISTER_FORMULA_RELATED_MUTATIONS_KEY";
var Ad = Object.getOwnPropertyDescriptor, xd = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? Ad(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, Mo = (n, e) => (t, o) => e(t, o, n);
let Xn = class extends ue {
  // eslint-disable-next-line max-lines-per-function
  constructor(n, e, t) {
    var s;
    super(), this._commandService = n, this._configService = e, this._dataSyncPrimaryController = t, [
      ee,
      ke,
      Ue,
      ln,
      wt,
      Je,
      qe,
      Me,
      Te,
      Lt,
      oe,
      li,
      ne,
      ti,
      Fn,
      ns,
      Nn,
      ud,
      Tt,
      // formula SUBTOTAL
      kt,
      ld,
      cd
    ].forEach((r) => {
      var i;
      this._commandService.registerCommand(r), (i = this._dataSyncPrimaryController) == null || i.registerSyncingMutations(r);
    }), ((s = this._configService.getConfig(di)) != null ? s : !1) || [
      tc,
      Vo,
      qo,
      Lo,
      Kr,
      It,
      vt,
      Ln,
      jn,
      sl,
      ol,
      rl,
      il,
      Dr,
      Xe,
      Pt,
      rn,
      el,
      Qu,
      nl,
      tl,
      Nr,
      Ye,
      ac,
      un,
      yt,
      an,
      Ar,
      ho,
      xr,
      co,
      Xo,
      bo,
      lc,
      Ac,
      Oc,
      Dc,
      wc,
      Sc,
      wn,
      Cc,
      pc,
      Eo,
      on,
      sn,
      Hn,
      Ic,
      zt,
      _c,
      ot,
      yc,
      $c,
      pn,
      zn,
      Uo,
      Uc,
      Kt,
      qr,
      Qr,
      dn,
      hn,
      Ce,
      Hc,
      Un,
      Nc,
      Vc,
      Wc,
      xc,
      ni,
      oi,
      Cn,
      zc,
      St,
      Qo,
      ri,
      kn,
      Xs,
      nt,
      ts,
      Fe,
      mt,
      // SetWorksheetColIsAutoWidthCommand,
      nd,
      Yt,
      Bc,
      Jt,
      ju,
      ie,
      gd,
      ic,
      Jr,
      Xr,
      ui,
      rd,
      Xt,
      bc,
      Gt,
      // permissions range protection
      Qc,
      Mt,
      bt,
      xt,
      es,
      Nu,
      Ec,
      oc,
      Zl,
      sc,
      ed,
      Be,
      st,
      fe,
      sd,
      qt,
      jc,
      od,
      // range theme
      nn,
      tn,
      Yo,
      cn,
      id,
      uc,
      Ql,
      rc,
      ad,
      hd,
      dd
    ].forEach((r) => this.disposeWithMe(this._commandService.registerCommand(r))), this._configService.setConfig(md, fd);
  }
};
Xn = xd([
  Mo(0, E),
  Mo(1, xo),
  Mo(2, Wi(aa))
], Xn);
var $d = Object.getOwnPropertyDescriptor, Wd = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? $d(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, bs = (n, e) => (t, o) => e(t, o, n);
let Zn = class extends ue {
  constructor(n, e) {
    super(), this._univerInstanceService = n, this._commandService = e, this._initialize();
  }
  _initialize() {
    this.disposeWithMe(
      this._commandService.onCommandExecuted((n) => {
        if (n.id !== na.id)
          return;
        const e = n.params, { unitData: t } = e, o = Object.keys(t), s = [];
        for (let i = 0; i < o.length; i++) {
          const a = o[i], u = t[a];
          if (u == null)
            continue;
          const l = Object.keys(u);
          for (let c = 0; c < l.length; c++) {
            const d = l[c], h = u[d];
            if (h == null)
              continue;
            const g = this._getMergedCellData(a, d, h), m = {
              subUnitId: d,
              unitId: a,
              cellValue: g
            };
            s.push({
              id: ee.id,
              params: m
            });
          }
        }
        return s.every(
          (i) => this._commandService.executeCommand(i.id, i.params, {
            onlyLocal: !0,
            fromFormula: !0
          })
        );
      })
    );
  }
  /**
   * Priority that mainly deals with number format in unitData
   * @param unitId
   * @param sheetId
   * @param cellData
   * @returns
   */
  _getMergedCellData(n, e, t) {
    const o = this._univerInstanceService.getUniverSheetInstance(n), s = o == null ? void 0 : o.getStyles(), r = o == null ? void 0 : o.getSheetBySheetId(e), i = r == null ? void 0 : r.getCellMatrix(), a = new Y(t);
    return a.forValue((u, l, c) => {
      const d = i == null ? void 0 : i.getValue(u, l), h = oa(d, c, s);
      a.setValue(u, l, h);
    }), a.getMatrix();
  }
};
Zn = Wd([
  bs(0, $(M)),
  bs(1, E)
], Zn);
var Vd = Object.getOwnPropertyDescriptor, Ld = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? Vd(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, Hd = (n, e) => (t, o) => e(t, o, n);
let Qn = class extends ue {
  constructor(n) {
    super(), this._sheetInterceptorService = n, this._initialize();
  }
  _initialize() {
    this._initInterceptorCellContent();
  }
  _initInterceptorCellContent() {
    this.disposeWithMe(
      this._sheetInterceptorService.intercept(ht.CELL_CONTENT, {
        priority: 11,
        effect: he.Value | he.Style,
        handler: (n, e, t) => {
          var s;
          if (!n)
            return t(n);
          const o = e.workbook.getStyles().getStyleByCell(n);
          return Vi((s = o == null ? void 0 : o.n) == null ? void 0 : s.pattern) && (n == null ? void 0 : n.t) === re.NUMBER && n.v !== void 0 && n.v !== null && No(n.v) && ((!n || n === e.rawData) && (n = { ...e.rawData }), n.v = sa(Number(n.v))), t(n);
        }
      })
    );
  }
};
Qn = Ld([
  Hd(0, $(G))
], Qn);
var Fd = Object.getOwnPropertyDescriptor, Bd = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? Fd(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, yn = (n, e) => (t, o) => e(t, o, n);
let eo = class extends ue {
  constructor(n, e, t, o) {
    super(), this._permissionService = n, this._worksheetProtectionRuleModel = e, this._sheetInterceptorService = t, this._rangeProtectionCache = o, this._initViewModelByRangeInterceptor(), this._initViewModelBySheetInterceptor();
  }
  _initViewModelByRangeInterceptor() {
    this.disposeWithMe(this._sheetInterceptorService.intercept(ht.CELL_CONTENT, {
      // permissions are placed at a high level to prioritize whether to filter subsequent renderings.
      priority: 999,
      effect: he.Value | he.Style,
      handler: (n, e, t) => {
        const { unitId: o, subUnitId: s, row: r, col: i } = e, a = this._rangeProtectionCache.getCellInfo(o, s, r, i);
        if (a) {
          const u = a[_.View] === !1, l = !n || n === e.rawData ? { ...e.rawData } : n;
          return l.selectionProtection = [a], u ? (delete l.s, delete l.v, delete l.p, l) : t(l);
        }
        return t(n);
      }
    }));
  }
  _initViewModelBySheetInterceptor() {
    this.disposeWithMe(this._sheetInterceptorService.intercept(ht.CELL_CONTENT, {
      // permissions are placed at a high level to prioritize whether to filter subsequent renderings.
      priority: 999,
      effect: he.Value | he.Style,
      handler: (n, e, t) => {
        var i, a, u, l, c;
        const { unitId: o, subUnitId: s } = e, r = this._worksheetProtectionRuleModel.getRule(o, s);
        if (r != null && r.permissionId) {
          const d = [{
            [_.View]: (a = (i = this._permissionService.getPermissionPoint(new lo(o, s).id)) == null ? void 0 : i.value) != null ? a : !1,
            [_.Edit]: (l = (u = this._permissionService.getPermissionPoint(new Le(o, s).id)) == null ? void 0 : u.value) != null ? l : !1
          }], h = !((c = d[0]) != null && c[_.View]), g = !n || n === e.rawData ? { ...n } : n;
          return g.hasWorksheetRule = !0, g.selectionProtection = d, h ? (delete g.s, delete g.v, delete g.p, g) : t(g);
        }
        return t(n);
      }
    }));
  }
};
eo = Bd([
  yn(0, Ze),
  yn(1, $(lt)),
  yn(2, $(G)),
  yn(3, $(Wt))
], eo);
const Es = Ao("univer.exclusive-range-service");
class jd extends ue {
  constructor() {
    super(...arguments);
    /**
     * Exclusive range data structure is as follows: unitId -> sheetId -> feature -> range
     */
    R(this, "_exclusiveRanges", /* @__PURE__ */ new Map());
    R(this, "_exclusiveRangesChange$", new De());
    R(this, "exclusiveRangesChange$", this._exclusiveRangesChange$.asObservable());
  }
  _ensureUnitMap(t) {
    return this._exclusiveRanges.has(t) || this._exclusiveRanges.set(t, /* @__PURE__ */ new Map()), this._exclusiveRanges.get(t);
  }
  _ensureSubunitMap(t, o) {
    const s = this._ensureUnitMap(t);
    return s.has(o) || s.set(o, /* @__PURE__ */ new Map()), s.get(o);
  }
  _ensureFeature(t, o, s) {
    const r = this._ensureSubunitMap(t, o);
    return r.has(s) || r.set(s, []), r.get(s);
  }
  addExclusiveRange(t, o, s, r) {
    const i = this._ensureFeature(t, o, s);
    i.push(...r), this._exclusiveRangesChange$.next({ unitId: t, subUnitId: o, ranges: i.map((a) => a.range) });
  }
  getExclusiveRanges(t, o, s) {
    var r, i;
    return (i = (r = this._exclusiveRanges.get(t)) == null ? void 0 : r.get(o)) == null ? void 0 : i.get(s);
  }
  clearExclusiveRanges(t, o, s) {
    const r = this.getExclusiveRanges(t, o, s);
    this._exclusiveRangesChange$.next({ unitId: t, subUnitId: o, ranges: (r == null ? void 0 : r.map((i) => i.range)) || [] }), this._ensureFeature(t, o, s), this._exclusiveRanges.get(t).get(o).set(s, []);
  }
  clearExclusiveRangesByGroupId(t, o, s, r) {
    const i = this.getExclusiveRanges(t, o, s);
    this._exclusiveRangesChange$.next({ unitId: t, subUnitId: o, ranges: (i == null ? void 0 : i.map((u) => u.range)) || [] });
    const a = this.getExclusiveRanges(t, o, s);
    if (a) {
      const u = a.filter((l) => l.groupId !== r);
      this._exclusiveRanges.get(t).get(o).set(s, u);
    }
  }
  getInterestGroupId(t) {
    const o = [];
    return t.forEach((s) => {
      var l;
      const r = s.range, { unitId: i, sheetId: a } = r;
      if (!i || !a) return;
      const u = (l = this._exclusiveRanges.get(i)) == null ? void 0 : l.get(a);
      if (u)
        for (const c of u.keys()) {
          const d = u.get(c);
          if (d) {
            for (const h of d)
              if (N.intersects(r, h.range)) {
                o.push(c);
                break;
              }
          }
        }
    }), o;
  }
}
var zd = Object.getOwnPropertyDescriptor, Gd = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? zd(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, _o = (n, e) => (t, o) => e(t, o, n);
let To = class extends ue {
  constructor(n, e, t) {
    super(), this._resourceManagerService = n, this._univerInstanceService = e, this._logService = t;
  }
  getValue(n, e, t, o) {
    const s = this._univerInstanceService.getUniverSheetInstance(n);
    if (!s)
      return;
    const r = s == null ? void 0 : s.getSheetBySheetId(e);
    if (!r)
      return;
    const i = s.getStyles(), a = r.getCellRaw(t, o);
    if (a != null && a.s) {
      const u = i.get(a.s);
      if (u != null && u.n)
        return u.n;
    }
    return null;
  }
  deleteValues(n, e, t) {
    const o = this._univerInstanceService.getUniverSheetInstance(n);
    if (!o)
      return;
    const s = o == null ? void 0 : o.getSheetBySheetId(e);
    if (!s)
      return;
    const r = o.getStyles();
    t.forEach((i) => {
      X.foreach(i, (a, u) => {
        const l = s.getCellRaw(a, u);
        if (!l)
          return;
        const c = l == null ? void 0 : l.s, h = { ...c && r.get(c) || {} };
        delete h.n;
        const g = r.setValue(h);
        l.s = g;
      });
    });
  }
  setValues(n, e, t) {
    const o = this._univerInstanceService.getUniverSheetInstance(n);
    if (!o)
      return;
    const s = o == null ? void 0 : o.getSheetBySheetId(e);
    if (!s)
      return;
    const r = o.getStyles(), i = s.getCellMatrix();
    t.forEach((a) => {
      a.ranges.forEach((u) => {
        X.foreach(u, (l, c) => {
          const d = s.getCellRaw(l, c);
          if (d) {
            const g = { ...r.getStyleByCell(d) || {}, n: { pattern: a.pattern } }, m = r.setValue(g);
            d.s = m;
          } else {
            const h = { n: { pattern: a.pattern } }, g = r.setValue(h);
            g && i.setValue(l, c, { s: g });
          }
        });
      });
    });
  }
};
To = Gd([
  _o(0, Rn),
  _o(1, M),
  _o(2, Fs)
], To);
var Kd = Object.getOwnPropertyDescriptor, Jd = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? Kd(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, dt = (n, e) => (t, o) => e(t, o, n);
const Us = [ke.id, Ue.id, Me.id, Te.id], ks = [Je.id, qe.id];
let to = class extends ue {
  constructor(e, t, o, s, r, i, a, u) {
    super();
    R(this, "disposableCollection", new Et());
    this._selectionProtectionRuleModel = e, this._univerInstanceService = t, this._commandService = o, this._refRangeService = s, this._selectionProtectionRenderModel = r, this._rangeProtectionCache = i, this._sheetInterceptorService = a, this._rangeProtectionRuleModel = u, this._onRefRangeChange(), this._correctPermissionRange(), this._initReBuildCache(), this._initRemoveSheet();
  }
  _onRefRangeChange() {
    const e = (o, s) => {
      const r = this._univerInstanceService.getCurrentUnitForType(B.UNIVER_SHEET);
      if (!r || !(r == null ? void 0 : r.getSheetBySheetId(s)))
        return;
      this.disposableCollection.dispose();
      const a = (l) => this.refRangeHandle(l, o, s);
      this._selectionProtectionRuleModel.getSubunitRuleList(o, s).reduce((l, c) => [...l, ...c.ranges], []).forEach((l) => {
        this.disposableCollection.add(this._refRangeService.registerRefRange(l, a, o, s));
      });
    };
    this.disposeWithMe(
      this._commandService.onCommandExecuted((o) => {
        if (o.id === oi.id) {
          const s = o.params, r = s.subUnitId, i = s.unitId;
          if (!r || !i)
            return;
          e(i, r);
        }
        if (o.id === fe.id || o.id === Be.id) {
          const s = o.params, r = s.subUnitId, i = s.unitId;
          if (!r || !i)
            return;
          e(i, r);
        }
      })
    );
    const t = this._univerInstanceService.getCurrentUnitForType(B.UNIVER_SHEET);
    if (t) {
      const o = t.getActiveSheet();
      if (!o) return;
      e(t.getUnitId(), o.getSheetId());
    }
  }
  refRangeHandle(e, t, o) {
    switch (e.id) {
      case an.id:
        return this._getRefRangeMutationsByMoveRows(e.params, t, o);
      case un.id:
        return this._getRefRangeMutationsByMoveCols(e.params, t, o);
      case Ye.id:
        return this._getRefRangeMutationsByInsertRows(e.params, t, o);
      case Xe.id:
        return this._getRefRangeMutationsByInsertCols(e.params, t, o);
      case ho.id:
        return this._getRefRangeMutationsByDeleteCols(e.params, t, o);
      case co.id:
        return this._getRefRangeMutationsByDeleteRows(e.params, t, o);
    }
    return { redos: [], undos: [] };
  }
  _getRefRangeMutationsByDeleteCols(e, t, o) {
    const s = this._selectionProtectionRuleModel.getSubunitRuleList(t, o).filter((i) => i.ranges.some((a) => N.intersects(a, e.range))), r = e.range;
    if (s.length) {
      const i = [], a = [];
      return s.forEach((u) => {
        const l = A.deepClone(u), c = l.ranges.reduce((d, h) => {
          if (N.intersects(h, r)) {
            const g = A.deepClone(h), { startColumn: m, endColumn: f } = r;
            if (m <= g.startColumn && f >= g.endColumn)
              return d;
            m >= g.startColumn && f <= g.endColumn ? g.endColumn -= f - m + 1 : m < g.startColumn ? (g.startColumn = m, g.endColumn -= f - m + 1) : f > g.endColumn && (g.endColumn = m - 1), this._checkIsRightRange(g) && d.push(g);
          }
          return d;
        }, []);
        l.ranges = c, l.ranges.length ? (i.push({ id: fe.id, params: { unitId: t, subUnitId: o, rule: l, ruleId: u.id } }), a.push({ id: fe.id, params: { unitId: t, subUnitId: o, rule: u, ruleId: u.id } })) : (i.push({ id: st.id, params: { unitId: t, subUnitId: o, ruleIds: [u.id] } }), a.push({ id: Be.id, params: { unitId: t, subUnitId: o, name: "", rules: [u] } }));
      }), { redos: i, undos: a };
    }
    return { undos: [], redos: [] };
  }
  _getRefRangeMutationsByDeleteRows(e, t, o) {
    const s = this._selectionProtectionRuleModel.getSubunitRuleList(t, o).filter((i) => i.ranges.some((a) => N.intersects(a, e.range))), r = e.range;
    if (s.length) {
      const i = [], a = [];
      return s.forEach((u) => {
        const l = A.deepClone(u), c = l.ranges.reduce((d, h) => {
          if (N.intersects(h, r)) {
            const g = A.deepClone(h), { startRow: m, endRow: f } = r;
            if (m <= g.startRow && f >= g.endRow)
              return d;
            m >= g.startRow && f <= g.endRow ? g.endRow -= f - m + 1 : m < g.startRow ? (g.startRow = m, g.endRow -= f - m + 1) : f > g.endRow && (g.endRow = m - 1), this._checkIsRightRange(g) && d.push(g);
          }
          return d;
        }, []);
        l.ranges = c, i.push({ id: fe.id, params: { unitId: t, subUnitId: o, rule: l, ruleId: u.id } }), a.push({ id: fe.id, params: { unitId: t, subUnitId: o, rule: u, ruleId: u.id } });
      }), { redos: i, undos: a };
    }
    return { undos: [], redos: [] };
  }
  _getRefRangeMutationsByInsertCols(e, t, o) {
    const s = e.range.startColumn, r = e.range.endColumn - e.range.startColumn + 1, i = this._selectionProtectionRuleModel.getSubunitRuleList(t, o).filter((a) => a.ranges.some((u) => s > u.startColumn && s <= u.endColumn));
    if (i.length) {
      const a = [], u = [];
      return i.forEach((l) => {
        const c = A.deepClone(l);
        let d = !1;
        c.ranges.forEach((h) => {
          s > h.startColumn && s <= h.endColumn && (h.endColumn += r, d = !0);
        }), d && (a.push({ id: fe.id, params: { unitId: t, subUnitId: o, rule: c, ruleId: l.id } }), u.push({ id: fe.id, params: { unitId: t, subUnitId: o, rule: l, ruleId: l.id } }));
      }), { redos: a, undos: u };
    }
    return { undos: [], redos: [] };
  }
  _getRefRangeMutationsByInsertRows(e, t, o) {
    const s = e.range.startRow, r = e.range.endRow - e.range.startRow + 1, i = this._selectionProtectionRuleModel.getSubunitRuleList(t, o).filter((a) => a.ranges.some((u) => s > u.startRow && s <= u.endRow));
    if (i.length) {
      const a = [], u = [];
      return i.forEach((l) => {
        const c = A.deepClone(l);
        let d = !1;
        c.ranges.forEach((h) => {
          s > h.startRow && s <= h.endRow && (h.endRow += r, d = !0);
        }), d && (a.push({ id: fe.id, params: { unitId: t, subUnitId: o, rule: c, ruleId: l.id } }), u.push({ id: fe.id, params: { unitId: t, subUnitId: o, rule: l, ruleId: l.id } }));
      }), { redos: a, undos: u };
    }
    return { undos: [], redos: [] };
  }
  _getRefRangeMutationsByMoveRows(e, t, o) {
    const s = e.toRange, r = s.startRow, i = s.endRow - s.startRow + 1, a = this._selectionProtectionRuleModel.getSubunitRuleList(t, o).filter((u) => u.ranges.some((l) => r > l.startRow && r <= l.endRow));
    if (a.length) {
      const u = [], l = [];
      return a.forEach((c) => {
        const d = A.deepClone(c), g = e.fromRange.startRow;
        let m = !1;
        d.ranges.forEach((f) => {
          r > f.startRow && r <= f.endRow && (g < f.startRow && (f.startRow = f.startRow - i, f.endRow = f.endRow - i), f.endRow += i, m = !0);
        }), m && (u.push({ id: fe.id, params: { unitId: t, subUnitId: o, rule: d, ruleId: c.id } }), l.push({ id: fe.id, params: { unitId: t, subUnitId: o, rule: c, ruleId: c.id } }));
      }), { redos: u, undos: l };
    }
    return { undos: [], redos: [] };
  }
  _getRefRangeMutationsByMoveCols(e, t, o) {
    const s = e.toRange, r = s.startColumn, i = s.endColumn - s.startColumn + 1, a = this._selectionProtectionRuleModel.getSubunitRuleList(t, o).filter((u) => u.ranges.some((l) => r > l.startColumn && r <= l.endColumn));
    if (a.length) {
      const u = [], l = [];
      return a.forEach((c) => {
        const d = A.deepClone(c), g = e.fromRange.startColumn;
        let m = !1;
        d.ranges.forEach((f) => {
          r > f.startColumn && r <= f.endColumn && (g < f.startColumn && (f.startColumn = f.startColumn - i, f.endColumn = f.endColumn - i), f.endColumn += i, m = !0);
        }), m && (u.push({ id: fe.id, params: { unitId: t, subUnitId: o, rule: d, ruleId: c.id } }), l.push({ id: fe.id, params: { unitId: t, subUnitId: o, rule: c, ruleId: c.id } }));
      }), { redos: u, undos: l };
    }
    return { undos: [], redos: [] };
  }
  _correctPermissionRange() {
    this.disposeWithMe(this._commandService.onCommandExecuted((e) => {
      if (ks.includes(e.id)) {
        if (!e.params) return;
        const t = this._univerInstanceService.getCurrentUnitForType(B.UNIVER_SHEET);
        if (!t) return;
        const o = t.getSheetBySheetId(e.params.subUnitId);
        if (!o) return;
        const { sourceRange: s, targetRange: r } = e.params, i = s.startColumn === r.startColumn && s.endColumn === r.endColumn, a = i ? s.endRow - s.startRow + 1 : s.endColumn - s.startColumn + 1, u = i ? s.startRow : s.startColumn, l = i ? r.startRow : r.startColumn;
        this._selectionProtectionRuleModel.getSubunitRuleList(t.getUnitId(), o.getSheetId()).forEach((f) => {
          f.ranges.forEach((p) => {
            let { startRow: w, endRow: I, startColumn: S, endColumn: y } = p;
            N.intersects(p, s) || (i ? u < w && l > I ? (w -= a, I -= a) : u > I && l <= w && (w += a, I += a) : u < S && l > y ? (S -= a, y -= a) : u > y && l <= S && (S += a, y += a)), this._checkIsRightRange({ startRow: w, endRow: I, startColumn: S, endColumn: y }) && (p.startColumn = S, p.endColumn = y, p.startRow = w, p.endRow = I);
          });
        }), this.disposableCollection.dispose();
        const { unitId: d, subUnitId: h } = e.params, g = (f) => this.refRangeHandle(f, d, h);
        this._selectionProtectionRuleModel.getSubunitRuleList(d, h).reduce((f, C) => [...f, ...C.ranges], []).forEach((f) => {
          this.disposableCollection.add(this._refRangeService.registerRefRange(f, g, d, h));
        }), this._selectionProtectionRenderModel.clear();
      }
      if (Us.includes(e.id)) {
        const t = this._univerInstanceService.getUniverSheetInstance(e.params.unitId);
        if (!t) return;
        const o = t.getSheetBySheetId(e.params.subUnitId);
        if (!o) return;
        const s = e.params;
        if (!s) return;
        const { range: r } = s, i = e.id.includes("row"), a = e.id.includes("insert"), u = i ? r.startRow : r.startColumn, l = i ? r.endRow : r.endColumn, c = l - u + 1;
        this._selectionProtectionRuleModel.getSubunitRuleList(t.getUnitId(), o.getSheetId()).forEach((C) => {
          C.ranges.forEach((w) => {
            let { startRow: I, endRow: S, startColumn: y, endColumn: b } = w;
            a ? i ? u <= I && (I += c, S += c) : u <= y && (y += c, b += c) : i ? l < I && (I -= c, S -= c) : l < y && (y -= c, b -= c), this._checkIsRightRange({ startRow: I, endRow: S, startColumn: y, endColumn: b }) && (w.startColumn = y, w.endColumn = b, w.startRow = I, w.endRow = S);
          });
        }), this.disposableCollection.dispose();
        const { unitId: h, subUnitId: g } = e.params, m = (C) => this.refRangeHandle(C, h, g);
        this._selectionProtectionRuleModel.getSubunitRuleList(h, g).reduce((C, p) => [...C, ...p.ranges], []).forEach((C) => {
          this.disposableCollection.add(this._refRangeService.registerRefRange(C, m, h, g));
        }), this._selectionProtectionRenderModel.clear();
      }
    }));
  }
  _checkIsRightRange(e) {
    return e.startRow <= e.endRow && e.startColumn <= e.endColumn;
  }
  _initReBuildCache() {
    this.disposeWithMe(this._commandService.onCommandExecuted((e) => {
      if (Us.includes(e.id) || ks.includes(e.id)) {
        const { unitId: t, subUnitId: o } = e.params;
        this._rangeProtectionCache.reBuildCache(t, o);
      }
    }));
  }
  _initRemoveSheet() {
    this._sheetInterceptorService.interceptCommand(
      {
        getMutations: (e) => {
          const t = [], o = [], s = [], r = [];
          if (e.id === Xo.id) {
            const i = e.params, a = [], u = [];
            this._rangeProtectionRuleModel.getSubunitRuleList(i.unitId, i.subUnitId).forEach((l) => {
              a.push(l.id), u.push(l);
            }), a.length && u.length && (s.push({ id: st.id, params: { unitId: i.unitId, subUnitId: i.subUnitId, ruleIds: a } }), t.push({ id: Be.id, params: { unitId: i.unitId, subUnitId: i.subUnitId, name: "", rules: u } }));
          }
          return {
            redos: o,
            undos: t,
            preRedos: s,
            preUndos: r
          };
        }
      }
    );
  }
};
to = Jd([
  dt(0, $(me)),
  dt(1, $(M)),
  dt(2, E),
  dt(3, $(At)),
  dt(4, $(fn)),
  dt(5, $(Wt)),
  dt(6, $(G)),
  dt(7, $(me))
], to);
var qd = Object.getOwnPropertyDescriptor, Yd = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? qd(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, Bt = (n, e) => (t, o) => e(t, o, n);
const Xd = "SHEET_RANGE_PROTECTION_PLUGIN";
let no = class extends ue {
  constructor(n, e, t, o, s) {
    super(), this._selectionProtectionRuleModel = n, this._permissionService = e, this._resourceManagerService = t, this._selectionProtectionCache = o, this._univerInstanceService = s, this._initSnapshot(), this._initRuleChange();
  }
  _initRuleChange() {
    this.disposeWithMe(
      this._selectionProtectionRuleModel.ruleChange$.subscribe((n) => {
        switch (n.type) {
          case "add": {
            Ee().forEach((e) => {
              const t = new e(n.unitId, n.subUnitId, n.rule.permissionId);
              this._permissionService.addPermissionPoint(t);
            });
            break;
          }
          case "delete": {
            Ee().forEach((e) => {
              const t = new e(n.unitId, n.subUnitId, n.rule.permissionId);
              this._permissionService.deletePermissionPoint(t.id);
            });
            break;
          }
          case "set": {
            n.oldRule.permissionId !== n.rule.permissionId && Ee().forEach((e) => {
              const t = new e(n.unitId, n.subUnitId, n.oldRule.permissionId);
              this._permissionService.deletePermissionPoint(t.id);
              const o = new e(n.unitId, n.subUnitId, n.rule.permissionId);
              this._permissionService.addPermissionPoint(o);
            });
            break;
          }
        }
      })
    );
  }
  _initSnapshot() {
    const n = (t) => {
      const s = this._selectionProtectionRuleModel.toObject()[t];
      return s ? JSON.stringify(s) : "";
    }, e = (t) => {
      if (!t)
        return {};
      try {
        return JSON.parse(t);
      } catch {
        return {};
      }
    };
    this.disposeWithMe(
      this._resourceManagerService.registerPluginResource({
        toJson: n,
        parseJson: e,
        pluginName: Xd,
        businesses: [On.UNIVER_SHEET],
        onLoad: (t, o) => {
          const s = this._selectionProtectionRuleModel.toObject();
          s[t] = o, this._selectionProtectionRuleModel.fromObject(s);
          const r = [];
          Object.keys(o).forEach((i) => {
            const a = o[i];
            this._selectionProtectionRuleModel.getSubunitRuleList(t, i).forEach((u) => {
              r.push({
                objectID: u.permissionId,
                unitID: t,
                objectType: x.SelectRange,
                actions: pt
              });
            }), a.forEach((u) => {
              Ee().forEach((l) => {
                const c = new l(t, i, u.permissionId);
                c.value = !1, this._permissionService.addPermissionPoint(c);
              });
            }), this._selectionProtectionCache.reBuildCache(t, i);
          });
        },
        onUnLoad: (t) => {
          this._selectionProtectionCache.deleteUnit(t);
        }
      })
    );
  }
};
no = Yd([
  Bt(0, $(me)),
  Bt(1, $(Ze)),
  Bt(2, $(Rn)),
  Bt(3, $(Wt)),
  Bt(4, $(M))
], no);
var Zd = Object.getOwnPropertyDescriptor, Qd = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? Zd(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, eh = (n, e) => (t, o) => e(t, o, n);
let Po = class extends ue {
  constructor(n) {
    super(), this._sheetRangeThemeModel = n;
  }
  /**
   * Register a custom range theme style.
   * @param {string} unitId Which unit to register the range theme style.
   * @param {RangeThemeStyle} rangeThemeStyle The range theme style to register.
   */
  registerRangeTheme(n, e) {
    this._sheetRangeThemeModel.registerRangeThemeStyle(n, e);
  }
  removeRangeThemeRule(n, e) {
    this._sheetRangeThemeModel.removeRangeThemeRule(n, e);
  }
  /**
   * Get custom register themes name list
   * @param {string} unitId Which unit to register the range theme style.
   * @returns {string[]} The list of custom register themes name.
   */
  getALLRegisterThemes(n) {
    return this._sheetRangeThemeModel.getALLRegisteredTheme(n);
  }
  /**
   * Register range theme style to the range.
   * @param {string} themeName The defined theme name.
   * @param {IRangeThemeRangeInfo} rangeInfo The range info to apply the theme style.
   */
  registerRangeThemeStyle(n, e) {
    this._sheetRangeThemeModel.registerRangeThemeRule(n, e);
  }
  /**
   * Get applied range theme style name.
   * @param {IRangeThemeRangeInfo} rangeInfo The range info to get the applied theme style.
   * @returns {string | undefined} The applied theme style name or not exist.
   */
  getAppliedRangeThemeStyle(n) {
    return this._sheetRangeThemeModel.getRegisteredRangeThemeStyle(n);
  }
  /**
   * Get registered build-in range theme style
   */
  getRegisteredRangeThemes() {
    return this._sheetRangeThemeModel.getRegisteredRangeThemes();
  }
};
Po = Qd([
  eh(0, $(ve))
], Po);
var th = Object.defineProperty, nh = Object.getOwnPropertyDescriptor, oh = (n, e, t) => e in n ? th(n, e, { enumerable: !0, configurable: !0, writable: !0, value: t }) : n[e] = t, sh = (n, e, t, o) => {
  for (var s = o > 1 ? void 0 : o ? nh(e, t) : e, r = n.length - 1, i; r >= 0; r--)
    (i = n[r]) && (s = i(s) || s);
  return s;
}, Ts = (n, e) => (t, o) => e(t, o, n), hi = (n, e, t) => oh(n, typeof e != "symbol" ? e + "" : e, t);
const rh = "SHEET_PLUGIN";
let oo = class extends Hi {
  constructor(n = _s, e, t) {
    super(), this._config = n, this._injector = e, this._configService = t;
    const { ...o } = Fi(
      {},
      _s,
      this._config
    );
    this._configService.setConfig(ci, o), this._initConfig(), this._initDependencies();
  }
  _initConfig() {
    var n, e, t;
    (n = this._config) != null && n.onlyRegisterFormulaRelatedMutations && this._configService.setConfig(di, !0), (e = this._config) != null && e.isRowStylePrecedeColumnStyle && this._configService.setConfig(Bi, !0), (t = this._config) != null && t.autoHeightForMergedCells && this._configService.setConfig(ji, !0);
  }
  _initDependencies() {
    var e;
    const n = [
      // services
      [Ht],
      [z],
      [At],
      [mn],
      [$t, { useClass: To }],
      [G],
      [Po],
      [rt],
      // controllers
      [Xn],
      [Wn],
      [Qn],
      [Gn],
      [Yn],
      [Kn],
      // permission
      [Bn],
      [lt],
      [In],
      [eo],
      [qn],
      [Jn],
      // range theme
      [ve],
      // range protection
      [fn],
      [me],
      [Wt],
      [to],
      [no],
      [Es, {
        useClass: jd,
        deps: [z]
      }]
    ];
    (e = this._config) != null && e.notExecuteFormula || n.push([Zn]), zi(this._injector, Gi(n, this._config.override)), vn(this._injector, [
      [G],
      [no],
      [Es],
      [qn],
      [Kn]
    ]);
  }
  onStarting() {
    vn(this._injector, [
      [Xn],
      [Wn],
      [mn],
      [Bn],
      [eo],
      [rt]
    ]);
  }
  onRendered() {
    vn(this._injector, [
      [$t]
    ]);
  }
  onReady() {
    vn(this._injector, [
      [Zn],
      [Gn],
      [Yn],
      [ve],
      [Qn],
      [fn],
      [to],
      [At],
      [Jn]
    ]);
  }
};
hi(oo, "pluginName", rh);
hi(oo, "type", B.UNIVER_SHEET);
oo = sh([
  Li(ra),
  Ts(1, $(so)),
  Ts(2, xo)
], oo);
const og = {
  /**
   * The permission point for adding or editing workbook comments
   */
  WorkbookCommentPermission: nr,
  /**
   * The permission point for copy in workbook
   */
  WorkbookCopyPermission: or,
  /**
   * The permission point for creating protect in a workbook
   */
  WorkbookCreateProtectPermission: sr,
  /**
   * The permission point for creating new sheets in a workbook
   */
  WorkbookCreateSheetPermission: rr,
  /**
   * The permission point for deleting sheets in a workbook
   */
  WorkbookDeleteSheetPermission: ir,
  /**
   * The permission point for duplicating a sheet in a workbook
   */
  WorkbookDuplicatePermission: ar,
  /**
   * The permission point for editing workbook content
   */
  WorkbookEditablePermission: Pe,
  /**
   * The permission point for exporting workbook data
   */
  WorkbookExportPermission: ur,
  /**
   * The permission point for hiding sheets in a workbook
   */
  WorkbookHideSheetPermission: Bo,
  /**
   * The permission point for viewing and managing workbook history
   */
  WorkbookHistoryPermission: qu,
  /**
   * The permission point for managing collaborators in a workbook
   */
  WorkbookManageCollaboratorPermission: jo,
  /**
   * The permission point for moving sheets within a workbook
   */
  WorkbookMoveSheetPermission: zo,
  /**
   * The permission point for printing a workbook
   */
  WorkbookPrintPermission: lr,
  /**
   * The permission point for recovering a previous history state of a workbook
   */
  WorkbookRecoverHistoryPermission: cr,
  /**
   * The permission point for renaming sheets in a workbook
   */
  WorkbookRenameSheetPermission: Go,
  /**
   * The permission point for sharing a workbook with others
   */
  WorkbookSharePermission: dr,
  /**
   * The permission point for viewing the history of a workbook
   */
  WorkbookViewHistoryPermission: gr,
  /**
   * The permission point for viewing a workbook
   */
  WorkbookViewPermission: hr,
  /**
   * The permission point for copying contents from a worksheet
   */
  WorksheetCopyPermission: mr,
  /**
   * The permission point for deleting columns in a worksheet
   */
  WorksheetDeleteColumnPermission: fr,
  /**
   * The permission point for deleting worksheet protection rules
   */
  WorksheetDeleteProtectionPermission: Rr,
  /**
   * The permission point for deleting rows in a worksheet
   */
  WorksheetDeleteRowPermission: Cr,
  /**
   * The permission point for editing extra objects (e.g. shapes) in a worksheet
   */
  WorksheetEditExtraObjectPermission: pr,
  /**
   * The permission point for editing the content of a worksheet
   */
  WorksheetEditPermission: Le,
  /**
   * The permission point for applying filters in a worksheet
   */
  WorksheetFilterPermission: Sr,
  /**
   * The permission point for inserting columns into a worksheet
   */
  WorksheetInsertColumnPermission: wr,
  /**
   * The permission point for inserting hyperlinks in a worksheet
   */
  WorksheetInsertHyperlinkPermission: Ir,
  /**
   * The permission point for inserting rows into a worksheet
   */
  WorksheetInsertRowPermission: vr,
  /**
   * The permission point for managing collaborators of a worksheet
   */
  WorksheetManageCollaboratorPermission: Mr,
  /**
   * The permission point for creating or modifying pivot tables in a worksheet
   */
  WorksheetPivotTablePermission: _r,
  /**
   * The permission point for setting the style of cells in a worksheet
   */
  WorksheetSetCellStylePermission: yr,
  /**
   * The permission point for setting the value of cells in a worksheet
   */
  WorksheetSetCellValuePermission: Dn,
  /**
   * The permission point for setting the style of columns in a worksheet
   */
  WorksheetSetColumnStylePermission: Zt,
  /**
   * The permission point for setting the style of rows in a worksheet
   */
  WorksheetSetRowStylePermission: Qt,
  /**
   * The permission point for performing sort operations on a worksheet
   */
  WorksheetSortPermission: br,
  /**
   * The permission point for viewing the content of a worksheet
   */
  WorksheetViewPermission: lo,
  /**
   * The permission point for editing the range protection settings
   */
  RangeProtectionPermissionEditPoint: Ve,
  /**
   * The permission point for viewing the range protection settings
   */
  RangeProtectionPermissionViewPoint: Fo
}, sg = (n, e, t, o) => {
  const s = n.get(Ze), r = n.get(me), i = s.getPermissionPoint(new Pe(e).id);
  if (!(i != null && i.value))
    return !1;
  const a = s.getPermissionPoint(new Le(e, t).id);
  if (!(a != null && a.value))
    return !1;
  const l = r.getSubunitRuleList(e, t).filter((c) => c.ranges.some((d) => o.some((h) => N.intersects(d, h))));
  return l.length ? l.every((c) => {
    const d = c.permissionId, h = s.getPermissionPoint(new Ve(e, t, d).id);
    return !!(h != null && h.value);
  }) : !0;
}, gi = (n, e, t, o = 1, s = !0, r = !0) => {
  const i = X.transformRange(n, e), { startRow: a, endRow: u } = i;
  let l = t.startRow - o, c = e.getMergedCell(l, t.startColumn), d = !c || c.startRow === l && c.startColumn === t.startColumn;
  for (; !e.getRowVisible(l) || !d; )
    l--, c = e.getMergedCell(l, t.startColumn), d = !c || c.startRow === l && c.startColumn === t.startColumn;
  if (l >= a)
    return { ...t, startRow: l, endRow: l };
  if (r) {
    const h = { ...t, startRow: u, endRow: u };
    return fi(n, e, h, o, s, !1);
  }
}, mi = (n, e, t, o = 1, s = !0, r = !0) => {
  const i = X.transformRange(n, e), { startRow: a, endRow: u } = i;
  let l = t.endRow + o, c = e.getMergedCell(l, t.startColumn), d = !c || c.startRow === l && c.startColumn === t.startColumn;
  for (; !e.getRowVisible(l) || !d; )
    l++, c = e.getMergedCell(l, t.startColumn), d = !c || c.startRow === l && c.startColumn === t.startColumn;
  if (l <= u)
    return { ...t, startRow: l, endRow: l };
  if (r) {
    const h = { ...t, startRow: a, endRow: a };
    return Ri(n, e, h, o, s, !1);
  }
}, fi = (n, e, t, o = 1, s = !0, r = !0) => {
  const i = X.transformRange(n, e), { startColumn: a, endColumn: u } = i;
  let l = t.startColumn - o, c = e.getMergedCell(t.startRow, l), d = !c || c.startRow === t.startRow && c.startColumn === l;
  for (; !e.getColVisible(l) || !d; )
    l--, c = e.getMergedCell(t.startRow, l), d = !c || c.startRow === t.startRow && c.startColumn === l;
  if (l >= a)
    return { ...t, startColumn: l, endColumn: l };
  if (r) {
    const h = { ...t, startColumn: u, endColumn: u };
    return gi(n, e, h, o, s, !1);
  }
}, Ri = (n, e, t, o = 1, s = !0, r = !0) => {
  const i = X.transformRange(n, e), { startColumn: a, endColumn: u } = i;
  let l = t.endColumn + o, c = e.getMergedCell(t.startRow, l), d = !c || c.startRow === t.startRow && c.startColumn === l;
  for (; !e.getColVisible(l) || !d; )
    l++, c = e.getMergedCell(t.startRow, l), d = !c || c.startRow === t.startRow && c.startColumn === l;
  if (l <= u)
    return { ...t, endColumn: l, startColumn: l };
  if (r) {
    const h = { ...t, startColumn: a, endColumn: a };
    return mi(n, e, h, o, s, !1);
  }
};
function bn(n, e, t) {
  let o = null;
  return t.getMatrixWithMergedCells(n, e, n, e).forValue((r, i, a) => (o = {
    actualRow: r,
    actualColumn: i,
    startRow: r,
    startColumn: i,
    isMerged: a.rowSpan !== void 0 || a.colSpan !== void 0,
    isMergedMainCell: a.rowSpan !== void 0 && a.colSpan !== void 0,
    endRow: r + (a.rowSpan !== void 0 ? a.rowSpan - 1 : 0),
    endColumn: i + (a.colSpan !== void 0 ? a.colSpan - 1 : 0),
    rangeType: j.NORMAL
  }, !1)), o || {
    actualColumn: e,
    actualRow: n,
    startRow: n,
    startColumn: e,
    endRow: n,
    endColumn: e,
    isMerged: !1,
    isMergedMainCell: !1,
    rangeType: j.NORMAL
  };
}
const ih = (n, e, t, o, s = 1) => {
  switch (o) {
    case pe.UP:
      return gi(n, e, t, s);
    case pe.DOWN:
      return mi(n, e, t, s);
    case pe.LEFT:
      return fi(n, e, t, s);
    case pe.RIGHT:
      return Ri(n, e, t, s);
  }
}, rg = (n, e, t) => {
  let o, s = -1, r;
  for (let p = 0; p < n.length; p++)
    if (n[p].primary) {
      o = n[p], s = p, r = o.primary;
      break;
    }
  if (s === -1)
    return null;
  const i = e === pe.LEFT || e === pe.UP, a = i ? s - 1 >= 0 ? s - 1 : n.length - 1 : s + 1 < n.length ? s + 1 : 0, u = n[a];
  if (!o || !r)
    return null;
  const l = { ...r }, { startRow: c, startColumn: d, endRow: h, endColumn: g } = o.range, m = i ? l.startRow === c && l.startColumn === d : l.endRow === h && l.endColumn === g, f = m && i;
  if (!N.equals(o.range, l)) {
    const p = m ? u.range : ih(o.range, t, l, e);
    if (!p)
      return null;
    const w = f ? bn(p.endRow, p.endColumn, t) : bn(p.startRow, p.startColumn, t);
    return {
      startRow: w.startRow,
      startColumn: w.startColumn,
      endRow: w.endRow,
      endColumn: w.endColumn
    };
  }
  const C = f ? bn(u.range.endRow, u.range.endColumn, t) : bn(u.range.startRow, u.range.startColumn, t);
  return {
    startRow: C.startRow,
    startColumn: C.startColumn,
    endRow: C.endRow,
    endColumn: C.endColumn
  };
};
export {
  fo as AFTER_CELL_EDIT,
  eg as AddMergeRedoSelectionsOperationFactory,
  We as AddMergeUndoMutationFactory,
  tg as AddMergeUndoSelectionsOperationFactory,
  Nu as AddRangeProtectionCommand,
  Be as AddRangeProtectionMutation,
  ad as AddRangeThemeMutation,
  Wh as AddWorksheetMergeAllCommand,
  go as AddWorksheetMergeCommand,
  Lh as AddWorksheetMergeHorizontalCommand,
  ne as AddWorksheetMergeMutation,
  Vh as AddWorksheetMergeVerticalCommand,
  Zl as AddWorksheetProtectionCommand,
  Mt as AddWorksheetProtectionMutation,
  tc as AppendRowCommand,
  ls as BEFORE_CELL_EDIT,
  Ht as BorderStyleManagerService,
  mh as COMMAND_LISTENER_SKELETON_CHANGE,
  fh as COMMAND_LISTENER_VALUE_CHANGE,
  yc as CancelFrozenCommand,
  cd as CancelMarkDirtyRowAutoHeightMutation,
  Vo as ClearSelectionAllCommand,
  qo as ClearSelectionContentCommand,
  Lo as ClearSelectionFormatCommand,
  Kr as CopySheetCommand,
  bh as DISABLE_NORMAL_SELECTIONS,
  Gn as DefinedNameDataController,
  It as DeleteRangeMoveLeftCommand,
  vt as DeleteRangeMoveUpCommand,
  oc as DeleteRangeProtectionCommand,
  st as DeleteRangeProtectionMutation,
  sc as DeleteWorksheetProtectionCommand,
  xt as DeleteWorksheetProtectionMutation,
  rc as DeleteWorksheetRangeThemeStyleCommand,
  nn as DeleteWorksheetRangeThemeStyleMutation,
  Da as DeleteWorksheetRangeThemeStyleMutationFactory,
  Ln as DeltaColumnWidthCommand,
  jn as DeltaRowHeightCommand,
  Pu as EditStateEnum,
  W as EffectRefRangId,
  ud as EmptyMutation,
  jd as ExclusiveRangeService,
  yh as FactoryAddRangeProtectionMutation,
  _h as FactoryDeleteRangeProtectionMutation,
  Fh as FactorySetRangeProtectionMutation,
  Es as IExclusiveRangeService,
  ht as INTERCEPTOR_POINT,
  $t as INumfmtService,
  Fu as IRefSelectionsService,
  sl as InsertColAfterCommand,
  ol as InsertColBeforeCommand,
  Dr as InsertColByRangeCommand,
  Xe as InsertColCommand,
  ke as InsertColMutation,
  ro as InsertColMutationUndoFactory,
  ic as InsertDefinedNameCommand,
  rl as InsertMultiColsLeftCommand,
  il as InsertMultiColsRightCommand,
  tl as InsertMultiRowsAboveCommand,
  nl as InsertMultiRowsAfterCommand,
  Pt as InsertRangeMoveDownCommand,
  rn as InsertRangeMoveRightCommand,
  el as InsertRowAfterCommand,
  Qu as InsertRowBeforeCommand,
  Nr as InsertRowByRangeCommand,
  Ye as InsertRowCommand,
  Ue as InsertRowMutation,
  $o as InsertRowMutationUndoFactory,
  ac as InsertSheetCommand,
  ln as InsertSheetMutation,
  Gr as InsertSheetUndoMutationFactory,
  ha as InterceptCellContentPriority,
  md as MAX_CELL_PER_SHEET_KEY,
  Kl as MERGE_CELL_INTERCEPTOR_CHECK,
  ld as MarkDirtyRowAutoHeightMutation,
  Wn as MergeCellController,
  un as MoveColsCommand,
  qe as MoveColsMutation,
  xa as MoveColsMutationUndoFactory,
  yt as MoveRangeCommand,
  wt as MoveRangeMutation,
  an as MoveRowsCommand,
  Je as MoveRowsMutation,
  Aa as MoveRowsMutationUndoFactory,
  To as NumfmtService,
  Q as OperatorType,
  og as PermissionPointsDefinitions,
  $u as REF_SELECTIONS_ENABLED,
  Sh as RangeMergeUtil,
  Wt as RangeProtectionCache,
  ii as RangeProtectionPermissionDeleteProtectionPoint,
  Ve as RangeProtectionPermissionEditPoint,
  ai as RangeProtectionPermissionManageCollaPoint,
  Fo as RangeProtectionPermissionViewPoint,
  to as RangeProtectionRefRangeService,
  fn as RangeProtectionRenderModel,
  me as RangeProtectionRuleModel,
  no as RangeProtectionService,
  Rt as RangeThemeStyle,
  At as RefRangeService,
  fs as RefSelectionsService,
  uc as RegisterWorksheetRangeThemeStyleCommand,
  cn as RegisterWorksheetRangeThemeStyleMutation,
  xr as RemoveColByRangeCommand,
  ho as RemoveColCommand,
  Me as RemoveColMutation,
  Jr as RemoveDefinedNameCommand,
  ye as RemoveMergeUndoMutationFactory,
  li as RemoveNumfmtMutation,
  dd as RemoveRangeThemeMutation,
  Ar as RemoveRowByRangeCommand,
  co as RemoveRowCommand,
  Te as RemoveRowMutation,
  Xo as RemoveSheetCommand,
  Lt as RemoveSheetMutation,
  al as RemoveSheetUndoMutationFactory,
  lc as RemoveWorksheetMergeCommand,
  oe as RemoveWorksheetMergeMutation,
  bo as ReorderRangeCommand,
  Nn as ReorderRangeMutation,
  Va as ReorderRangeUndoMutationFactory,
  Ac as ResetBackgroundColorCommand,
  Oc as ResetTextColorCommand,
  ng as SCOPE_WORKBOOK_VALUE_DEFINED_NAME,
  Eh as SELECTIONS_ENABLED,
  Ih as SELECTION_CONTROL_BORDER_BUFFER_COLOR,
  wh as SELECTION_CONTROL_BORDER_BUFFER_WIDTH,
  gd as ScrollToCellOperation,
  ju as SelectRangeCommand,
  we as SelectionMoveType,
  Dc as SetBackgroundColorCommand,
  Bh as SetBoldCommand,
  wc as SetBorderBasicCommand,
  Sc as SetBorderColorCommand,
  wn as SetBorderCommand,
  Cc as SetBorderPositionCommand,
  pc as SetBorderStyleCommand,
  Ic as SetColDataCommand,
  zt as SetColDataMutation,
  Fa as SetColDataMutationFactory,
  Eo as SetColHiddenCommand,
  on as SetColHiddenMutation,
  sn as SetColVisibleMutation,
  Hn as SetColWidthCommand,
  Xr as SetDefinedNameCommand,
  Jh as SetFontFamilyCommand,
  qh as SetFontSizeCommand,
  _c as SetFrozenCommand,
  ot as SetFrozenMutation,
  Zr as SetFrozenMutationFactory,
  bc as SetGridlinesColorCommand,
  Gt as SetGridlinesColorMutation,
  $c as SetHorizontalTextAlignCommand,
  jh as SetItalicCommand,
  ns as SetNumfmtMutation,
  Kh as SetOverlineCommand,
  Ec as SetProtectionCommand,
  fe as SetRangeProtectionMutation,
  hd as SetRangeThemeMutation,
  pn as SetRangeValuesCommand,
  ee as SetRangeValuesMutation,
  Ae as SetRangeValuesUndoMutationFactory,
  Uc as SetRowDataCommand,
  Kt as SetRowDataMutation,
  eu as SetRowDataMutationFactory,
  zn as SetRowHeightCommand,
  Uo as SetRowHiddenCommand,
  Tt as SetRowHiddenMutation,
  kt as SetRowVisibleMutation,
  qr as SetSelectedColsVisibleCommand,
  Qr as SetSelectedRowsVisibleCommand,
  ie as SetSelectionsOperation,
  dn as SetSpecificColsVisibleCommand,
  hn as SetSpecificRowsVisibleCommand,
  Gh as SetStrikeThroughCommand,
  Ce as SetStyleCommand,
  Hc as SetTabColorCommand,
  Un as SetTabColorMutation,
  Nc as SetTextColorCommand,
  Vc as SetTextRotationCommand,
  Wc as SetTextWrapCommand,
  zh as SetUnderlineCommand,
  xc as SetVerticalTextAlignCommand,
  ni as SetWorkbookNameCommand,
  ti as SetWorkbookNameMutation,
  oi as SetWorksheetActivateCommand,
  Cn as SetWorksheetActiveOperation,
  mt as SetWorksheetColWidthMutation,
  qs as SetWorksheetColWidthMutationFactory,
  Bc as SetWorksheetColumnCountCommand,
  Jt as SetWorksheetColumnCountMutation,
  ou as SetWorksheetColumnCountUndoMutationFactory,
  jc as SetWorksheetDefaultStyleCommand,
  qt as SetWorksheetDefaultStyleMutation,
  su as SetWorksheetDefaultStyleMutationFactory,
  zc as SetWorksheetHideCommand,
  St as SetWorksheetHideMutation,
  Qo as SetWorksheetNameCommand,
  Fn as SetWorksheetNameMutation,
  ri as SetWorksheetOrderCommand,
  kn as SetWorksheetOrderMutation,
  Qc as SetWorksheetPermissionPointsCommand,
  es as SetWorksheetPermissionPointsMutation,
  ed as SetWorksheetProtectionCommand,
  bt as SetWorksheetProtectionMutation,
  Ql as SetWorksheetRangeThemeStyleCommand,
  tn as SetWorksheetRangeThemeStyleMutation,
  Oa as SetWorksheetRangeThemeStyleMutationFactory,
  Xh as SetWorksheetRightToLeftCommand,
  So as SetWorksheetRightToLeftMutation,
  Xs as SetWorksheetRowAutoHeightMutation,
  gh as SetWorksheetRowAutoHeightMutationFactory,
  nd as SetWorksheetRowCountCommand,
  Yt as SetWorksheetRowCountMutation,
  ru as SetWorksheetRowCountUndoMutationFactory,
  nt as SetWorksheetRowHeightMutation,
  ts as SetWorksheetRowIsAutoHeightCommand,
  Fe as SetWorksheetRowIsAutoHeightMutation,
  ui as SetWorksheetShowCommand,
  G as SheetInterceptorService,
  Jn as SheetPermissionCheckController,
  qn as SheetPermissionInitController,
  ve as SheetRangeThemeModel,
  Po as SheetRangeThemeService,
  iu as SheetSkeletonChangeType,
  rt as SheetSkeletonService,
  au as SheetValueChangeType,
  Kn as SheetsFreezeSyncController,
  z as SheetsSelectionsService,
  Cu as SplitDelimiterEnum,
  od as SplitTextToColumnsCommand,
  sd as ToggleCellCheckboxCommand,
  rd as ToggleGridlinesCommand,
  Xt as ToggleGridlinesMutation,
  _ as UnitAction,
  x as UnitObject,
  oo as UniverSheetsPlugin,
  id as UnregisterWorksheetRangeThemeStyleCommand,
  Yo as UnregisterWorksheetRangeThemeStyleMutation,
  Ro as VALIDATE_CELL,
  Tu as ViewStateEnum,
  nr as WorkbookCommentPermission,
  or as WorkbookCopyPermission,
  Gu as WorkbookCopySheetPermission,
  sr as WorkbookCreateProtectPermission,
  rr as WorkbookCreateSheetPermission,
  Ku as WorkbookDeleteColumnPermission,
  Ju as WorkbookDeleteRowPermission,
  ir as WorkbookDeleteSheetPermission,
  ar as WorkbookDuplicatePermission,
  Pe as WorkbookEditablePermission,
  ur as WorkbookExportPermission,
  Bo as WorkbookHideSheetPermission,
  qu as WorkbookHistoryPermission,
  Yu as WorkbookInsertColumnPermission,
  Xu as WorkbookInsertRowPermission,
  jo as WorkbookManageCollaboratorPermission,
  zo as WorkbookMoveSheetPermission,
  mn as WorkbookPermissionService,
  lr as WorkbookPrintPermission,
  cr as WorkbookRecoverHistoryPermission,
  Go as WorkbookRenameSheetPermission,
  Ou as WorkbookSelectionModel,
  dr as WorkbookSharePermission,
  gr as WorkbookViewHistoryPermission,
  hr as WorkbookViewPermission,
  mr as WorksheetCopyPermission,
  fr as WorksheetDeleteColumnPermission,
  Rr as WorksheetDeleteProtectionPermission,
  Cr as WorksheetDeleteRowPermission,
  pr as WorksheetEditExtraObjectPermission,
  Le as WorksheetEditPermission,
  Sr as WorksheetFilterPermission,
  wr as WorksheetInsertColumnPermission,
  Ir as WorksheetInsertHyperlinkPermission,
  vr as WorksheetInsertRowPermission,
  Mr as WorksheetManageCollaboratorPermission,
  Bn as WorksheetPermissionService,
  _r as WorksheetPivotTablePermission,
  In as WorksheetProtectionPointModel,
  lt as WorksheetProtectionRuleModel,
  Ph as WorksheetSelectProtectedCellsPermission,
  Nh as WorksheetSelectUnProtectedCellsPermission,
  yr as WorksheetSetCellStylePermission,
  Dn as WorksheetSetCellValuePermission,
  Zt as WorksheetSetColumnStylePermission,
  Qt as WorksheetSetRowStylePermission,
  br as WorksheetSortPermission,
  lo as WorksheetViewPermission,
  Yn as ZebraCrossingCacheController,
  Hh as addMergeCellsUtil,
  Al as adjustRangeOnMutation,
  uo as alignToMergedCellsBorders,
  pt as baseProtectionActions,
  as as checkCellValueType,
  sg as checkRangesEditablePermission,
  Ru as convertPrimaryWithCoordToPrimary,
  vh as convertSelectionDataToRange,
  Ct as copyRangeStyles,
  mu as createTopMatrixFromMatrix,
  gu as createTopMatrixFromRanges,
  Jc as defaultWorkbookPermissionPoints,
  po as defaultWorksheetPermissionPoint,
  ph as expandToContinuousRange,
  Qh as factoryRemoveNumfmtUndoMutation,
  Zh as factorySetNumfmtUndoMutation,
  Zs as findAllRectangle,
  Mh as findFirstNonEmptyCell,
  ut as followSelectionOperation,
  er as generateNullCell,
  Eu as generateNullCellValue,
  zr as getAddMergeMutationRangeByType,
  Ee as getAllRangePermissionPoint,
  gn as getAllWorkbookPermissionPoint,
  be as getAllWorksheetPermissionPoint,
  He as getAllWorksheetPermissionPointByPointPanel,
  Uh as getCellAtRowCol,
  Yh as getDefaultRangePermission,
  Er as getInsertRangeMutations,
  Ho as getMoveRangeUndoRedoMutations,
  rg as getNextPrimaryCell,
  _e as getPrimaryForRange,
  Ur as getRemoveRangeMutations,
  Bu as getSelectionsService,
  $h as getSeparateEffectedRangesOnCommand,
  P as getSheetCommandTarget,
  Ks as getSheetCommandTargetWorkbook,
  Qe as getSheetMutationTarget,
  Ch as getSkeletonChangedEffectedRange,
  Rh as getValueChangedEffectedRange,
  ao as getVisibleRanges,
  Ot as handleBaseInsertRange,
  $n as handleBaseMoveRowsCols,
  Nt as handleBaseRemoveRange,
  Cs as handleCommonDefaultRangeChangeWithEffectRefCommands,
  xh as handleCommonRangeChangeWithEffectRefCommandsSkipNoInterests,
  Rs as handleDefaultRangeChangeWithEffectRefCommands,
  Ah as handleDefaultRangeChangeWithEffectRefCommandsSkipNoInterests,
  Ul as handleDeleteRangeMoveLeft,
  Tl as handleDeleteRangeMoveUp,
  Dh as handleDeleteRangeMutation,
  Br as handleIRemoveCol,
  Sl as handleIRemoveRow,
  Ml as handleInsertCol,
  _l as handleInsertRangeMoveDown,
  bl as handleInsertRangeMoveRight,
  Oh as handleInsertRangeMutation,
  vl as handleInsertRow,
  Fr as handleMoveCols,
  Cl as handleMoveRange,
  Hr as handleMoveRows,
  Th as isSingleCellSelection,
  Qs as rangeMerge,
  ku as rangeToDiscreteRange,
  Oe as rotateRange,
  Dt as runRefRangeMutations,
  kh as setEndForRange,
  Mu as splitRangeText,
  ko as transformCellsToRange
};
