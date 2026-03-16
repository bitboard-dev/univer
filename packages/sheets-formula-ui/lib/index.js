var $r = Object.defineProperty;
var Pr = (t, e, n) => e in t ? $r(t, e, { enumerable: !0, configurable: !0, writable: !0, value: n }) : t[e] = n;
var q = (t, e, n) => Pr(t, typeof e != "symbol" ? e + "" : e, n);
import { CommandType as Te, ICommandService as me, createIdentifier as Ur, IContextService as Qt, IUniverInstanceService as ne, Rectangle as qn, DOCS_NORMAL_EDITOR_UNIT_ID_KEY as Jt, DOCS_FORMULA_BAR_EDITOR_UNIT_ID_KEY as jn, DEFAULT_EMPTY_DOCUMENT_VALUE as Kn, isRealNum as Wr, CellValueType as _n, getCellValueType as Vr, Inject as B, Disposable as rt, ObjectMatrix as Ae, Range as Hr, Tools as en, LocaleService as ot, isICellData as Br, isFormulaString as ke, isFormulaId as tt, generateRandomId as Ft, Direction as de, Injector as ft, UniverInstanceType as K, ThemeService as gt, ILogService as qr, toDisposable as Yn, ColorKit as tn, RxDisposable as jr, InterceptorEffectEnum as Kr, FOCUSING_DOC as Yr, FOCUSING_UNIVER_EDITOR as Zr, DisposableCollection as Be, RANGE_TYPE as Se, getBodySlice as Rn, EDITOR_ACTIVATED as En, createInternalEditorID as zr, BuildTextUtils as Gr, IConfigService as Zn, RichTextBuilder as Xr, DependentOn as Qr, Plugin as Jr, merge as eo, registerDependencies as to, touchDependencies as no } from "@univerjs/core";
import { SheetPasteCommand as ro, PREDEFINED_HOOK_NAME as Je, IEditorBridgeService as nn, SetCellEditVisibleOperation as zn, HoverManagerService as oo, CellAlertManagerService as so, CellAlertType as io, IAutoFillService as co, APPLY_TYPE as ao, DATA_TYPE as In, ISheetClipboardService as lo, COPY_TYPE as Gn, SheetSkeletonManagerService as $t, attachSelectionWithCoord as Yt, SelectionControl as Xn, SELECTION_SHAPE_DEPTH as uo, useActiveWorkbook as ho, getCurrentRangeDisable$ as fo, PASTE_SPECIAL_MENU_ID as go, whenFormulaEditorActivated as mt, whenSheetEditorFocused as mo, SheetsUIPart as po, BaseSelectionRenderService as So, getCoordByOffset as bn, checkInHeaderRanges as yn, getAllSelection as vo, genNormalSelectionStyle as Qn, getSheetObject as Co, MoveSelectionCommand as Tn, JumpOver as xn, ExpandSelectionCommand as Nn, EMBEDDING_FORMULA_EDITOR as _o, IMarkSelectionService as Ro, RANGE_SELECTOR_COMPONENT_KEY as Eo, EMBEDDING_FORMULA_EDITOR_COMPONENT_KEY as Io } from "@univerjs/sheets-ui";
import { sequenceNodeType as j, serializeRange as ve, FormulaDataModel as pt, LexerTreeBuilder as Ce, ErrorType as ue, extractFormulaError as Jn, SetFormulaCalculationResultMutation as bo, SetArrayFormulaDataMutation as yo, SetFormulaCalculationStopMutation as To, FunctionType as er, matchToken as qe, deserializeRangeWithSheetWithCache as xo, matchRefDrawToken as No, isFormulaLexerToken as Oo, deserializeRangeWithSheet as st, serializeRangeToRefString as Mo, serializeRangeWithSheet as dt, serializeRangeWithSpreadsheet as wo, generateStringWithSequence as ko, operatorToken as tr, UniverFormulaEnginePlugin as Ao } from "@univerjs/engine-formula";
import { Subject as ct, debounceTime as rn, combineLatestWith as Do, map as nr, switchMap as Lo, of as On, Observable as Fo, BehaviorSubject as rr, throttleTime as $o, filter as or, distinctUntilChanged as Po, merge as Uo } from "rxjs";
import { IEditorService as je, DocSelectionRenderService as sr, ReplaceTextRunsCommand as Mn, MoveSelectionOperation as Wo, MoveCursorOperation as Vo, useKeyboardEvent as Ho, useResize as Bo, DocBackScrollRenderController as qo, RichTextEditor as jo } from "@univerjs/docs-ui";
import { DeviceInputEventType as xe, IRenderManagerService as Ne, ScrollTimerType as Bt, SHEET_VIEWPORT_KEY as wn, Vector2 as kn } from "@univerjs/engine-render";
import { SheetsSelectionsService as on, getSheetCommandTarget as ir, getCellAtRowCol as Ko, SetSelectionsOperation as sn, SheetInterceptorService as cn, ReorderRangeCommand as Yo, SetRangeValuesMutation as Tt, SetRangeValuesUndoMutationFactory as Zo, BEFORE_CELL_EDIT as zo, SetWorksheetRowAutoHeightMutation as Go, INTERCEPTOR_POINT as Xo, WorksheetSetCellValuePermission as Qo, WorksheetEditPermission as Jo, RangeProtectionPermissionEditPoint as es, WorkbookEditablePermission as ts, IRefSelectionsService as Pt, SelectionMoveType as qt, convertSelectionDataToRange as ns, setEndForRange as rs, REF_SELECTIONS_ENABLED as An, SetWorksheetActiveOperation as Dn } from "@univerjs/sheets";
import { InsertFunctionCommand as os, TriggerCalculationController as ss, IDescriptionService as Ut, QuickSumCommand as is, UniverSheetsFormulaPlugin as cs } from "@univerjs/sheets-formula";
import { ISidebarService as an, IZenZoneService as as, useDependency as T, useObservable as pe, ProgressBar as ls, MenuItemType as us, IClipboardInterfaceService as Ln, KeyCode as F, MetaKeys as P, IMenuManagerService as ds, IShortcutService as Wt, IUIPartsService as cr, ComponentManager as ar, connectInjector as lr, useEvent as te, RectPopup as Zt, IContextMenuService as hs, useUpdateEffect as fs, BuiltInUIPart as gs } from "@univerjs/ui";
import { jsx as M, jsxs as V, Fragment as ms } from "react/jsx-runtime";
import { useCallback as xt, useState as W, useRef as X, createElement as Ke, forwardRef as Le, useEffect as $, useMemo as ie, useLayoutEffect as Nt, useImperativeHandle as ps } from "react";
import { clsx as re, scrollbarClassName as nt, borderLeftClassName as Ss, Select as vs, Input as ur, borderClassName as ln, Button as et, borderTopClassName as Cs, Tooltip as _s, Dialog as Rs } from "@univerjs/design";
import { DocSelectionManagerService as dr } from "@univerjs/docs";
import { debounceTime as Es } from "rxjs/operators";
const un = {
  id: "sheet.command.paste-formula",
  type: Te.COMMAND,
  handler: async (t) => t.get(me).executeCommand(ro.id, {
    value: Je.SPECIAL_PASTE_FORMULA
  })
}, it = {
  id: "formula-ui.operation.select-editor-formula",
  type: Te.OPERATION,
  handler: (t, e) => !0
};
var Is = Object.getOwnPropertyDescriptor, bs = (t, e, n, o) => {
  for (var r = o > 1 ? void 0 : o ? Is(e, n) : e, s = t.length - 1, i; s >= 0; s--)
    (i = t[s]) && (r = i(r) || r);
  return r;
}, ys = (t, e) => (n, o) => e(n, o, t);
const Ts = "FORMULA_PROMPT_ACTIVATED", Vt = Ur("formula-ui.prompt-service");
let zt = class {
  constructor(t) {
    q(this, "_search$", new ct());
    q(this, "_help$", new ct());
    q(this, "_navigate$", new ct());
    q(this, "_accept$", new ct());
    q(this, "_acceptFormulaName$", new ct());
    q(this, "search$", this._search$.asObservable());
    q(this, "help$", this._help$.asObservable());
    q(this, "navigate$", this._navigate$.asObservable());
    q(this, "accept$", this._accept$.asObservable());
    q(this, "acceptFormulaName$", this._acceptFormulaName$.asObservable());
    q(this, "_searching", !1);
    q(this, "_helping", !1);
    q(this, "_sequenceNodes", []);
    q(this, "_isLockedOnSelectionChangeRefString", !1);
    q(this, "_isLockedOnSelectionInsertRefString", !1);
    this._contextService = t;
  }
  dispose() {
    this._search$.complete(), this._help$.complete(), this._navigate$.complete(), this._accept$.complete(), this._acceptFormulaName$.complete(), this._sequenceNodes = [];
  }
  search(t) {
    this._contextService.setContextValue(Ts, t.visible), this._searching = t.visible, this._search$.next(t);
  }
  isSearching() {
    return this._searching;
  }
  help(t) {
    this._helping = t.visible, this._help$.next(t);
  }
  isHelping() {
    return this._helping;
  }
  navigate(t) {
    this._navigate$.next(t);
  }
  accept(t) {
    this._accept$.next(t);
  }
  acceptFormulaName(t) {
    this._acceptFormulaName$.next(t);
  }
  getSequenceNodes() {
    return [...this._sequenceNodes];
  }
  setSequenceNodes(t) {
    this._sequenceNodes = t;
  }
  clearSequenceNodes() {
    this._sequenceNodes = [];
  }
  getCurrentSequenceNode(t) {
    return this._sequenceNodes[this.getCurrentSequenceNodeIndex(t)];
  }
  getCurrentSequenceNodeByIndex(t) {
    return this._sequenceNodes[t];
  }
  /**
   * Query the text coordinates in the sequenceNodes and determine the actual insertion index.
   * @param strIndex
   */
  getCurrentSequenceNodeIndex(t) {
    let e = 0;
    const n = this._sequenceNodes[0];
    for (let o = 0, r = this._sequenceNodes.length; o < r; o++) {
      const s = this._sequenceNodes[o];
      if (typeof s == "string")
        e++;
      else {
        const { endIndex: i } = s;
        e = i;
      }
      if (t <= e)
        return typeof n == "string" && t !== 0 ? o + 1 : o;
    }
    return this._sequenceNodes.length;
  }
  /**
   * Synchronize the reference text based on the changes of the selection.
   * @param nodeIndex
   * @param refString
   */
  updateSequenceRef(t, e) {
    const n = this._sequenceNodes[t];
    if (typeof n == "string" || n.nodeType !== j.REFERENCE)
      return;
    const o = e.length - n.token.length, r = { ...n };
    r.token = e, r.endIndex += o, this._sequenceNodes[t] = r;
    for (let s = t + 1, i = this._sequenceNodes.length; s < i; s++) {
      const c = this._sequenceNodes[s];
      if (typeof c == "string")
        continue;
      const l = { ...c };
      l.startIndex += o, l.endIndex += o, this._sequenceNodes[s] = l;
    }
  }
  /**
   * When the cursor is on the right side of a formula token,
   * you can add reference text to the formula by drawing a selection.
   * @param index
   * @param refString
   */
  insertSequenceRef(t, e) {
    const n = e.length, o = this.getCurrentSequenceNodeIndex(t);
    this._sequenceNodes.splice(o, 0, {
      token: e,
      startIndex: t,
      endIndex: t + n - 1,
      nodeType: j.REFERENCE
    });
    for (let r = o + 1, s = this._sequenceNodes.length; r < s; r++) {
      const i = this._sequenceNodes[r];
      if (typeof i == "string")
        continue;
      const c = { ...i };
      c.startIndex += n, c.endIndex += n, this._sequenceNodes[r] = c;
    }
  }
  /**
   * Insert a string at the cursor position in the text corresponding to the sequenceNodes.
   * @param index
   * @param content
   */
  insertSequenceString(t, e) {
    const n = this.getCurrentSequenceNodeIndex(t), o = e.split("");
    this._sequenceNodes.splice(n, 0, ...o);
    const r = o.length;
    for (let s = n + r, i = this._sequenceNodes.length; s < i; s++) {
      const c = this._sequenceNodes[s];
      if (typeof c == "string")
        continue;
      const l = { ...c };
      l.startIndex += r, l.endIndex += r, this._sequenceNodes[s] = l;
    }
  }
  enableLockedSelectionChange() {
    this._isLockedOnSelectionChangeRefString = !0;
  }
  disableLockedSelectionChange() {
    this._isLockedOnSelectionChangeRefString = !1;
  }
  isLockedSelectionChange() {
    return this._isLockedOnSelectionChangeRefString;
  }
  enableLockedSelectionInsert() {
    this._isLockedOnSelectionInsertRefString = !0;
  }
  disableLockedSelectionInsert() {
    this._isLockedOnSelectionInsertRefString = !1;
  }
  isLockedSelectionInsert() {
    return this._isLockedOnSelectionInsertRefString;
  }
};
zt = bs([
  ys(0, Qt)
], zt);
const xs = {
  id: "formula-ui.operation.help-function",
  type: Te.OPERATION,
  handler: async (t, e) => (t.get(Vt).help(e), !0)
}, Ns = {
  id: "formula-ui.operation.insert-function",
  type: Te.OPERATION,
  // eslint-disable-next-line
  handler: async (t, e) => {
    var C, _;
    const n = t.get(on), o = t.get(je), r = n.getCurrentSelections();
    if (!r || !r.length)
      return !1;
    const s = ir(t.get(ne));
    if (!s) return !1;
    const { worksheet: i, unitId: c, subUnitId: l } = s, m = i.getCellMatrix(), { value: g } = e, h = t.get(me);
    t.get(nn);
    const u = [], d = [];
    let a = null, v = 0, S = 0, p = "";
    if (r.length === 1 && (ws(r[0].range) || ks(r[0].range) && $n(m, r[0].range))) {
      const { range: I, primary: f } = r[0], b = (C = f == null ? void 0 : f.actualRow) != null ? C : I.startRow, y = (_ = f == null ? void 0 : f.actualColumn) != null ? _ : I.startColumn;
      a = I, v = b, S = y;
      const A = Fn(m, b, y);
      A && (p = ve(A));
    } else
      r.some((I) => {
        var y, A;
        const { range: f, primary: b } = I;
        if ($n(m, f)) {
          const D = (y = b == null ? void 0 : b.actualRow) != null ? y : f.startRow, x = (A = b == null ? void 0 : b.actualColumn) != null ? A : f.startColumn, N = Fn(m, D, x);
          if (!N)
            return a = f, v = D, S = x, !0;
          const k = ve(N), E = `=${g}(${k})`;
          u.push({
            range: f,
            primary: {
              row: D,
              column: x
            },
            formula: E
          });
        } else {
          const { startRow: D, startColumn: x, endRow: N, endColumn: k } = f;
          if (D === N) {
            const E = As(m, D, k, i.getColumnCount() - 1), O = E === k ? k - 1 : k, R = ve({
              startRow: D,
              endRow: N,
              startColumn: x,
              endColumn: O
            }), w = `=${g}(${R})`;
            d.push({
              range: f,
              primary: {
                row: D,
                column: E
              },
              formula: w
            });
          } else {
            let E = -1;
            for (let R = x; R <= k; R++) {
              const w = Ds(m, R, N, i.getRowCount() - 1);
              E = Math.max(E, w);
            }
            const O = E === N ? N - 1 : N;
            for (let R = x; R <= k; R++) {
              const w = ve({
                startRow: D,
                endRow: O,
                startColumn: R,
                endColumn: R
              }), L = `=${g}(${w})`;
              d.push({
                range: f,
                primary: {
                  row: E,
                  column: R
                },
                formula: L
              });
            }
          }
        }
        return !1;
      });
    if (a) {
      const I = Ko(v, S, i), f = {
        range: qn.clone(a),
        primary: {
          startRow: I.startRow,
          startColumn: I.startColumn,
          endRow: I.endRow,
          endColumn: I.endColumn,
          actualRow: v,
          actualColumn: S,
          isMerged: I.isMerged,
          isMergedMainCell: I.startRow === v && I.startColumn === S
        }
      }, b = {
        unitId: c,
        subUnitId: l,
        selections: [f]
      };
      await h.executeCommand(sn.id, b);
      const y = o.getEditor(Jt), A = o.getEditor(jn);
      h.syncExecuteCommand(zn.id, {
        visible: !0,
        unitId: c,
        eventType: xe.Dblclick
      });
      const D = `=${g}(${p}`;
      y == null || y.replaceText(D), A == null || A.replaceText(D, !1);
    }
    return u.length === 0 && d.length === 0 ? !1 : h.executeCommand(os.id, {
      list: u,
      listOfRangeHasNumber: d
    });
  }
};
function Fn(t, e, n) {
  const o = Os(t, e, n);
  if (o !== e)
    return {
      startRow: o,
      endRow: e - 1,
      startColumn: n,
      endColumn: n
    };
  const r = Ms(t, e, n);
  return r !== n ? {
    startRow: e,
    endRow: e,
    startColumn: r,
    endColumn: n - 1
  } : null;
}
function Os(t, e, n) {
  let o = !1;
  if (e === 0) return e;
  for (let r = e - 1; r >= 0; r--) {
    const s = t.getValue(r, n);
    if (ht(s) && !o) {
      if (r === 0) return 0;
      o = !0;
    } else {
      if (o && !ht(s))
        return r + 1;
      if (o && r === 0)
        return 0;
    }
  }
  return e;
}
function Ms(t, e, n) {
  let o = !1;
  if (n === 0) return n;
  for (let r = n - 1; r >= 0; r--) {
    const s = t.getValue(e, r);
    if (ht(s) && !o) {
      if (r === 0) return 0;
      o = !0;
    } else {
      if (o && !ht(s))
        return r + 1;
      if (o && r === 0)
        return 0;
    }
  }
  return n;
}
function ht(t) {
  if (t != null && t.p) {
    const e = t == null ? void 0 : t.p.body;
    if (e == null)
      return !1;
    const n = e.dataStream, r = n.substring(n.length - 2, n.length) === Kn ? n.substring(0, n.length - 2) : n;
    return Wr(r);
  }
  return t && (t.t === _n.NUMBER || Vr(t) === _n.NUMBER);
}
function ws(t) {
  return t.startRow === t.endRow && t.startColumn === t.endColumn;
}
function ks(t) {
  return t.startRow !== t.endRow && t.startColumn !== t.endColumn;
}
function $n(t, e) {
  for (let n = e.startRow; n <= e.endRow; n++)
    for (let o = e.startColumn; o <= e.endColumn; o++)
      if (ht(t.getValue(n, o)))
        return !1;
  return !0;
}
function As(t, e, n, o) {
  for (let r = n; r <= o; r++)
    if (!t.getValue(e, r))
      return r;
  return o;
}
function Ds(t, e, n, o) {
  for (let r = n; r <= o; r++)
    if (!t.getValue(r, e))
      return r;
  return o;
}
const hr = "SHEET_FORMULA_UI_PLUGIN", fr = `${hr}_MORE_FUNCTIONS_COMPONENT`, Ls = {
  id: "formula-ui.operation.more-functions",
  type: Te.OPERATION,
  handler: async (t) => (t.get(an).open({
    header: { title: "formula.insert.tooltip" },
    children: { label: fr }
  }), !0)
}, gr = {
  id: "formula-ui.operation.change-ref-to-absolute",
  type: Te.OPERATION,
  handler: async (t) => !0
}, Fs = {
  id: "formula-ui.operation.search-function",
  type: Te.OPERATION,
  handler: async (t, e) => (t.get(Vt).search(e), !0)
};
var $s = Object.getOwnPropertyDescriptor, Ps = (t, e, n, o) => {
  for (var r = o > 1 ? void 0 : o ? $s(e, n) : e, s = t.length - 1, i; s >= 0; s--)
    (i = t[s]) && (r = i(r) || r);
  return r;
}, It = (t, e) => (n, o) => e(n, o, t);
let Ot = class extends rt {
  constructor(t, e, n, o) {
    super(), this._sheetInterceptorService = t, this._univerInstanceService = e, this._formulaDataModel = n, this._lexerTreeBuilder = o, this._initialize();
  }
  _initialize() {
    this.disposeWithMe(this._sheetInterceptorService.interceptCommand({
      getMutations: (t) => t.id === Yo.id ? this._reorderFormula(t.params) : {
        redos: [],
        undos: []
      }
    }));
  }
  _reorderFormula(t) {
    const e = [], n = [], { unitId: o, subUnitId: r, range: s, order: i } = t, c = this._univerInstanceService.getUniverSheetInstance(o), l = c == null ? void 0 : c.getSheetBySheetId(r);
    if (!l)
      return {
        redos: e,
        undos: n
      };
    const m = l.getCellMatrix(), g = new Ae(), h = new Ae();
    let u = !1;
    return Hr.foreach(s, (d, a) => {
      let v = d;
      i.hasOwnProperty(d) && (v = i[d]);
      const S = m.getValue(v, a);
      if (S != null && S.f || S != null && S.si) {
        u = !0;
        const p = this._formulaDataModel.getFormulaStringByCell(v, a, r, o), C = this._lexerTreeBuilder.moveFormulaRefOffset(
          p,
          0,
          d - v
        ), _ = en.deepClone(S);
        _.f = C, _.si = null, g.setValue(d, a, _);
      } else
        g.setValue(d, a, S);
      h.setValue(d, a, m.getValue(d, a));
    }), u ? (e.push({
      id: Tt.id,
      params: {
        unitId: o,
        subUnitId: r,
        cellValue: g.getMatrix()
      }
    }), n.push({
      id: Tt.id,
      params: {
        unitId: o,
        subUnitId: r,
        cellValue: h.getMatrix()
      }
    }), {
      redos: e,
      undos: n
    }) : {
      redos: e,
      undos: n
    };
  }
};
Ot = Ps([
  It(0, B(cn)),
  It(1, B(ne)),
  It(2, B(pt)),
  It(3, B(Ce))
], Ot);
const mr = "sheets-formula-ui.base.config", Pn = {};
var Us = Object.getOwnPropertyDescriptor, Ws = (t, e, n, o) => {
  for (var r = o > 1 ? void 0 : o ? Us(e, n) : e, s = t.length - 1, i; s >= 0; s--)
    (i = t[s]) && (r = i(r) || r);
  return r;
}, at = (t, e) => (n, o) => e(n, o, t);
const jt = "SHEET_FORMULA_ALERT", Vs = {
  [ue.DIV_BY_ZERO]: "divByZero",
  [ue.NAME]: "name",
  [ue.VALUE]: "value",
  [ue.NUM]: "num",
  [ue.NA]: "na",
  [ue.CYCLE]: "cycle",
  [ue.REF]: "ref",
  [ue.SPILL]: "spill",
  [ue.CALC]: "calc",
  [ue.ERROR]: "error",
  [ue.CONNECT]: "connect",
  [ue.NULL]: "null"
};
let Gt = class extends rt {
  constructor(t, e, n, o, r, s) {
    super(), this._context = t, this._hoverManagerService = e, this._cellAlertManagerService = n, this._localeService = o, this._formulaDataModel = r, this._zenZoneService = s, this._init();
  }
  _init() {
    this._initCellAlertPopup(), this._initZenService();
  }
  _initCellAlertPopup() {
    this.disposeWithMe(this._hoverManagerService.currentCell$.pipe(rn(100)).subscribe((t) => {
      var e, n, o, r, s;
      if (t) {
        const c = this._context.unit.getActiveSheet();
        if (!c) return;
        const l = c.getCell(t.location.row, t.location.col), m = (r = (o = (n = (e = this._formulaDataModel.getArrayFormulaCellData()) == null ? void 0 : e[t.location.unitId]) == null ? void 0 : n[t.location.subUnitId]) == null ? void 0 : o[t.location.row]) == null ? void 0 : r[t.location.col];
        if (Br(l)) {
          const g = Jn(l, !!m);
          if (!g) {
            this._hideAlert();
            return;
          }
          const h = this._cellAlertManagerService.currentAlert.get(jt), u = (s = h == null ? void 0 : h.alert) == null ? void 0 : s.location;
          if (u && u.row === t.location.row && u.col === t.location.col && u.subUnitId === t.location.subUnitId && u.unitId === t.location.unitId)
            return;
          this._cellAlertManagerService.showAlert({
            type: io.ERROR,
            title: this._localeService.t("formula.error.title"),
            message: this._localeService.t(`formula.error.${Vs[g]}`),
            location: t.location,
            width: 200,
            height: 74,
            key: jt
          });
          return;
        }
      }
      this._hideAlert();
    }));
  }
  _initZenService() {
    this.disposeWithMe(this._zenZoneService.visible$.subscribe((t) => {
      t && this._hideAlert();
    }));
  }
  _hideAlert() {
    this._cellAlertManagerService.removeAlert(jt);
  }
};
Gt = Ws([
  at(1, B(oo)),
  at(2, B(so)),
  at(3, B(ot)),
  at(4, B(pt)),
  at(5, as)
], Gt);
var Hs = Object.getOwnPropertyDescriptor, Bs = (t, e, n, o) => {
  for (var r = o > 1 ? void 0 : o ? Hs(e, n) : e, s = t.length - 1, i; s >= 0; s--)
    (i = t[s]) && (r = i(r) || r);
  return r;
}, Un = (t, e) => (n, o) => e(n, o, t);
let Mt = class extends rt {
  constructor(t, e) {
    super(), this._autoFillService = t, this._lexerTreeBuilder = e, this._registerAutoFill();
  }
  _registerAutoFill() {
    const t = {
      type: In.FORMULA,
      priority: 1001,
      match: (e) => ke(e == null ? void 0 : e.f) || tt(e == null ? void 0 : e.si),
      isContinue: (e, n) => e.type === In.FORMULA,
      applyFunctions: {
        [ao.COPY]: (e, n, o, r, s) => {
          const { data: i, index: c } = e;
          return this._fillCopyFormula(i, n, o, c, r, s);
        }
      }
    };
    this._autoFillService.registerRule(t);
  }
  _fillCopyFormula(t, e, n, o, r, s) {
    var m, g;
    const i = js(r), c = [], l = /* @__PURE__ */ new Map();
    for (let h = 1; h <= e; h++) {
      const u = (h - 1) % t.length, d = o[u], a = en.deepClone(t[u]);
      if (a) {
        const v = ((m = t[u]) == null ? void 0 : m.f) || "", S = ((g = t[u]) == null ? void 0 : g.si) || "", p = ke(v);
        if (tt(S))
          a.si = S, a.f = null, a.v = null, a.p = null, a.t = null, c.push(a);
        else if (p) {
          let _ = l.get(u);
          if (_)
            a.si = _, a.f = null, a.v = null, a.p = null, a.t = null;
          else {
            _ = Ft(6), l.set(u, _);
            const { offsetX: I, offsetY: f } = qs(i, e, n, s, d), b = this._lexerTreeBuilder.moveFormulaRefOffset(
              v,
              I,
              f
            );
            a.si = _, a.f = b, a.v = null, a.p = null, a.t = null;
          }
          c.push(a);
        }
      }
    }
    return c;
  }
};
Mt = Bs([
  Un(0, co),
  Un(1, B(Ce))
], Mt);
function qs(t, e, n, o, r) {
  const { source: s, target: i } = o, { rows: c } = i, { rows: l } = s;
  let m = 0, g = 0;
  switch (n) {
    case de.UP:
      g = c[r] - l[r];
      break;
    case de.RIGHT:
      m = t;
      break;
    case de.DOWN:
      g = c[r] - l[r];
      break;
    case de.LEFT:
      m = -t * e;
      break;
  }
  return { offsetX: m, offsetY: g };
}
function js(t) {
  let e = 0;
  for (const n in t)
    t[n].forEach((o) => {
      e += o.data.length;
    });
  return e;
}
var Ks = Object.getOwnPropertyDescriptor, Ys = (t, e, n, o) => {
  for (var r = o > 1 ? void 0 : o ? Ks(e, n) : e, s = t.length - 1, i; s >= 0; s--)
    (i = t[s]) && (r = i(r) || r);
  return r;
}, lt = (t, e) => (n, o) => e(n, o, t);
const Zs = "default-paste-formula";
let wt = class extends rt {
  constructor(t, e, n, o, r) {
    super(), this._currentUniverSheet = t, this._lexerTreeBuilder = e, this._sheetClipboardService = n, this._injector = o, this._formulaDataModel = r, this._initialize();
  }
  _initialize() {
    this._registerClipboardHook();
  }
  _registerClipboardHook() {
    this.disposeWithMe(this._sheetClipboardService.addClipboardHook(this._pasteFormulaHook())), this.disposeWithMe(this._sheetClipboardService.addClipboardHook(this._pasteWithFormulaHook()));
  }
  _pasteFormulaHook() {
    return {
      id: Je.SPECIAL_PASTE_FORMULA,
      priority: 10,
      specialPasteInfo: { label: "specialPaste.formula" },
      onPasteCells: (t, e, n, o) => this._onPasteCells(t, e, n, o, !0)
    };
  }
  _pasteWithFormulaHook() {
    return {
      id: Zs,
      priority: 10,
      onPasteCells: (t, e, n, o) => this._onPasteCells(t, e, n, o, !1)
    };
  }
  _onPasteCells(t, e, n, o, r) {
    var u;
    if ([
      Je.SPECIAL_PASTE_FORMAT,
      Je.SPECIAL_PASTE_COL_WIDTH
    ].includes(o.pasteType))
      return {
        undos: [],
        redos: []
      };
    const i = this._currentUniverSheet.getCurrentUnitForType(K.UNIVER_SHEET), c = e.unitId || i.getUnitId(), l = e.subUnitId || ((u = i.getActiveSheet()) == null ? void 0 : u.getSheetId());
    if (!c || !l)
      return {
        undos: [],
        redos: []
      };
    const m = e.range, g = n, h = {
      copyType: o.copyType || Gn.COPY,
      copyRange: t == null ? void 0 : t.range,
      pasteType: o.pasteType
    };
    return this._injector.invoke((d) => zs(
      c,
      l,
      m,
      g,
      d,
      h,
      this._lexerTreeBuilder,
      this._formulaDataModel,
      r,
      t
    ));
  }
};
wt = Ys([
  lt(0, ne),
  lt(1, B(Ce)),
  lt(2, lo),
  lt(3, B(ft)),
  lt(4, B(pt))
], wt);
function zs(t, e, n, o, r, s, i, c, l = !1, m) {
  const g = [], h = [], u = Gs(t, e, n, o, s, i, c, m);
  if (!u.hasValue())
    return {
      undos: [],
      redos: []
    };
  const d = {
    unitId: t,
    subUnitId: e,
    cellValue: u.getData()
  };
  g.push({
    id: Tt.id,
    params: d
  });
  const a = Zo(
    r,
    d
  );
  return h.push({
    id: Tt.id,
    params: a
  }), {
    undos: h,
    redos: g
  };
}
function Gs(t, e, n, o, r, s, i, c) {
  return c ? r.pasteType === Je.SPECIAL_PASTE_VALUE ? Qs(t, e, n, o, i, c) : r.pasteType === Je.SPECIAL_PASTE_FORMULA ? Js(t, e, n, o, s, i, c) : ei(t, e, n, o, r.copyType, s, i, c) : Xs(t, e, n, o, i);
}
function Xs(t, e, n, o, r) {
  const s = new Ae(), i = r.getSheetFormulaData(t, e);
  return o.forValue((c, l, m) => {
    var d;
    const g = n.rows[c], h = n.cols[l], u = {};
    ke(m.v) ? (u.v = null, u.f = `${m.v}`, u.si = null, u.p = null, s.setValue(g, h, u)) : (d = i == null ? void 0 : i[g]) != null && d[h] && (u.v = m.v, u.f = null, u.si = null, u.p = null, s.setValue(g, h, u));
  }), s;
}
function Qs(t, e, n, o, r, s) {
  var m, g;
  const i = new Ae(), c = (g = (m = r.getArrayFormulaCellData()) == null ? void 0 : m[s.unitId]) == null ? void 0 : g[s.subUnitId], l = r.getSheetFormulaData(t, e);
  return o.forValue((h, u, d) => {
    var _, I;
    const a = s.range.rows[h % s.range.rows.length], v = s.range.cols[u % s.range.cols.length], S = n.rows[h], p = n.cols[u], C = {};
    if (ke(d.f) || tt(d.si))
      C.v = d.v, C.f = null, C.si = null, C.p = null, i.setValue(S, p, C);
    else if ((_ = c == null ? void 0 : c[a]) != null && _[v]) {
      const f = c[a][v];
      C.v = f.v, C.f = null, C.si = null, C.p = null, i.setValue(S, p, C);
    } else if ((I = l == null ? void 0 : l[S]) != null && I[p]) {
      if (C.v = d.v, C.f = null, C.si = null, C.p = null, d.p) {
        const f = pr(d);
        f && (C.v = f);
      }
      i.setValue(S, p, C);
    }
  }), i;
}
function Js(t, e, n, o, r, s, i) {
  const c = new Ae(), l = /* @__PURE__ */ new Map();
  return o.forValue((m, g, h) => {
    const u = n.rows[m], d = n.cols[g], a = {};
    if (tt(h.si)) {
      if (i.unitId !== t || i.subUnitId !== e) {
        const v = s.getFormulaStringByCell(
          i.range.rows[m % i.range.rows.length],
          i.range.cols[g % i.range.cols.length],
          i.subUnitId,
          i.unitId
        ), S = n.cols[g] - i.range.cols[g % i.range.cols.length], p = n.rows[m] - i.range.rows[m % i.range.rows.length], C = r.moveFormulaRefOffset(v || "", S, p);
        a.si = null, a.f = C;
      } else
        a.si = h.si, a.f = null;
      a.v = null, a.p = null, c.setValue(u, d, a);
    } else if (ke(h.f)) {
      const v = `${m % i.range.rows.length}_${g % i.range.cols.length}`;
      let S = l.get(v);
      if (S)
        a.si = S, a.f = null;
      else {
        S = Ft(6), l.set(v, S);
        const p = n.cols[g] - i.range.cols[g % i.range.cols.length], C = n.rows[m] - i.range.rows[m % i.range.rows.length], _ = r.moveFormulaRefOffset(h.f || "", p, C);
        a.si = S, a.f = _;
      }
      a.v = null, a.p = null, c.setValue(u, d, a);
    } else {
      if (a.v = h.v, a.f = null, a.si = null, a.p = null, h.p) {
        const v = pr(h);
        v && (a.v = v);
      }
      c.setValue(u, d, a);
    }
  }), c;
}
function ei(t, e, n, o, r, s, i, c) {
  const l = new Ae(), m = /* @__PURE__ */ new Map(), g = i.getSheetFormulaData(t, e), h = [];
  return r === Gn.CUT ? o.forValue((u, d, a) => {
    const v = n.rows[u], S = n.cols[d], p = {};
    if (tt(a.si)) {
      if (ke(a.f))
        h.push(a.si), p.f = a.f, p.si = a.si;
      else if (h.includes(a.si))
        p.f = null, p.si = a.si;
      else {
        const C = i.getFormulaStringByCell(
          c.range.rows[u % c.range.rows.length],
          c.range.cols[d % c.range.cols.length],
          c.subUnitId,
          c.unitId
        );
        p.f = C, p.si = null;
      }
      p.v = null, p.p = null, l.setValue(v, S, p);
    } else ke(a.f) && (p.f = a.f, p.si = null, p.v = null, p.p = null, l.setValue(v, S, p));
  }) : o.forValue((u, d, a) => {
    var C;
    const v = n.rows[u], S = n.cols[d], p = {};
    if (tt(a.si)) {
      if (c.unitId !== t || c.subUnitId !== e) {
        const _ = i.getFormulaStringByCell(
          c.range.rows[u % c.range.rows.length],
          c.range.cols[d % c.range.cols.length],
          c.subUnitId,
          c.unitId
        ), I = n.cols[d] - c.range.cols[d % c.range.cols.length], f = n.rows[u] - c.range.rows[u % c.range.rows.length], b = s.moveFormulaRefOffset(_ || "", I, f);
        p.si = null, p.f = b;
      } else
        p.si = a.si, p.f = null;
      p.v = null, p.p = null, l.setValue(v, S, p);
    } else if (ke(a.f)) {
      const _ = `${u % c.range.rows.length}_${d % c.range.cols.length}`;
      let I = m.get(_);
      if (I)
        p.si = I, p.f = null;
      else {
        I = Ft(6), m.set(_, I);
        const f = n.cols[d] - c.range.cols[d % c.range.cols.length], b = n.rows[u] - c.range.rows[u % c.range.rows.length], y = s.moveFormulaRefOffset(a.f || "", f, b);
        p.si = I, p.f = y;
      }
      p.v = null, p.p = null, l.setValue(v, S, p);
    } else (C = g == null ? void 0 : g[v]) != null && C[S] && (p.v = a.v, p.f = null, p.si = null, p.p = a.p, l.setValue(v, S, p));
  }), h.length > 0 && new Ae(g).forValue((u, d, a) => {
    if (!(c.range.rows.includes(u) && c.range.cols.includes(d)) && !(n.rows.includes(u) && n.cols.includes(d)) && h.includes(a == null ? void 0 : a.si)) {
      const v = i.getFormulaStringByCell(
        u,
        d,
        c.subUnitId,
        c.unitId
      );
      l.setValue(u, d, {
        f: v,
        si: null,
        v: null,
        p: null
      });
    }
  }), l;
}
function pr(t) {
  if (t != null && t.p) {
    const e = t == null ? void 0 : t.p.body;
    if (e == null)
      return;
    const n = e.dataStream;
    return n.substring(n.length - 2, n.length) === Kn ? n.substring(0, n.length - 2) : n;
  }
}
var ti = Object.getOwnPropertyDescriptor, ni = (t, e, n, o) => {
  for (var r = o > 1 ? void 0 : o ? ti(e, n) : e, s = t.length - 1, i; s >= 0; s--)
    (i = t[s]) && (r = i(r) || r);
  return r;
}, Ve = (t, e) => (n, o) => e(n, o, t);
let kt = class extends rt {
  constructor(e, n, o, r, s, i, c, l) {
    super();
    q(this, "_previousShape");
    q(this, "_skeleton");
    this._context = e, this._sheetInterceptorService = n, this._formulaDataModel = o, this._themeService = r, this._renderManagerService = s, this._sheetSkeletonManagerService = i, this._commandService = c, this._logService = l, this._initSkeletonChangeListener(), this._initInterceptorEditorStart(), this._commandExecutedListener();
  }
  _initSkeletonChangeListener() {
    this.disposeWithMe(
      this._sheetSkeletonManagerService.currentSkeleton$.subscribe((e) => {
        var n, o;
        if (e == null)
          this._logService.debug("[FormulaEditorShowController]: should not receive currentSkeleton$ as null!");
        else {
          const { skeleton: r } = e, s = (o = (n = this._skeleton) == null ? void 0 : n.worksheet) == null ? void 0 : o.getSheetId();
          if (this._changeRuntime(r), s !== r.worksheet.getSheetId())
            this._removeArrayFormulaRangeShape();
          else {
            const { unitId: i, sheetId: c } = e;
            this._updateArrayFormulaRangeShape(i, c);
          }
        }
      })
    );
  }
  _changeRuntime(e) {
    this._skeleton = e;
  }
  _initInterceptorEditorStart() {
    this.disposeWithMe(
      Yn(
        this._sheetInterceptorService.writeCellInterceptor.intercept(zo, {
          handler: (e, n, o) => {
            var a, v, S, p;
            const { row: r, col: s, unitId: i, subUnitId: c, worksheet: l } = n, m = this._formulaDataModel.getArrayFormulaRange(), g = this._formulaDataModel.getArrayFormulaCellData();
            if (this._removeArrayFormulaRangeShape(), e == null)
              return o(e);
            let h = null;
            const u = this._formulaDataModel.getFormulaStringByCell(r, s, c, i);
            if (u !== null && (h = { f: u }), e.v != null && e.v !== "" && ((S = (v = (a = g[i]) == null ? void 0 : a[c]) == null ? void 0 : v[r]) == null ? void 0 : S[s]) == null)
              return h ? { ...e, ...h } : o(e);
            const d = (p = m == null ? void 0 : m[i]) == null ? void 0 : p[c];
            return d != null && (h = this._displayArrayFormulaRangeShape(d, r, s, i, c, l, h)), h ? { ...e, ...h } : o(e);
          }
        })
      )
    );
  }
  _commandExecutedListener() {
    this.disposeWithMe(this._commandService.onCommandExecuted((e, n) => {
      (e.id === bo.id || e.id === yo.id && n && n.remove) && this._removeArrayFormulaRangeShape();
    })), this.disposeWithMe(
      this._commandService.beforeCommandExecuted((e) => {
        Go.id === e.id && requestIdleCallback(() => {
          const n = e.params, { unitId: o, subUnitId: r, rowsAutoHeightInfo: s } = n;
          this._refreshArrayFormulaRangeShapeByRow(o, r, s);
        });
      })
    );
  }
  _displayArrayFormulaRangeShape(e, n, o, r, s, i, c) {
    const l = this._formulaDataModel.getSheetFormulaData(r, s);
    return new Ae(e).forValue((m, g, h) => {
      var S;
      if (h == null)
        return !0;
      const { startRow: u, startColumn: d, endRow: a, endColumn: v } = h;
      if (m === n && g === o)
        return this._createArrayFormulaRangeShape(h, r), !1;
      if (n >= u && n <= a && o >= d && o <= v) {
        const p = i.getCell(u, d);
        if ((p == null ? void 0 : p.v) === ue.SPILL)
          return;
        const C = (S = l == null ? void 0 : l[m]) == null ? void 0 : S[g];
        return C == null || C.f == null ? !0 : (c == null && (c = {
          f: C.f,
          isInArrayFormulaRange: !0
        }), this._createArrayFormulaRangeShape(h, r), !1);
      }
    }), c;
  }
  _createArrayFormulaRangeShape(e, n) {
    const o = this._renderManagerService.getRenderById(n), r = this._sheetSkeletonManagerService.getCurrentSkeleton();
    if (!o || !r) return;
    const { scene: s } = o;
    if (!s) return;
    const i = {
      range: e,
      primary: null,
      style: {
        strokeWidth: 1,
        stroke: this._themeService.getColorFromTheme("primary.600"),
        fill: new tn(this._themeService.getColorFromTheme("white")).setAlpha(0).toString(),
        widgets: {}
      }
    }, c = Yt(i, r), { rowHeaderWidth: l, columnHeaderHeight: m } = r, g = new Xn(s, uo.FORMULA_EDITOR_SHOW, this._themeService, {
      highlightHeader: !1,
      rowHeaderWidth: l,
      columnHeaderHeight: m
    });
    g.updateRangeBySelectionWithCoord(c), g.setEvent(!1), this._previousShape = g;
  }
  _removeArrayFormulaRangeShape() {
    this._previousShape != null && (this._previousShape.dispose(), this._previousShape = null);
  }
  _refreshArrayFormulaRangeShape(e, n) {
    if (this._previousShape) {
      const { startRow: o, endRow: r, startColumn: s, endColumn: i } = this._previousShape.getRange(), c = { startRow: o, endRow: r, startColumn: s, endColumn: i };
      this._removeArrayFormulaRangeShape(), this._createArrayFormulaRangeShape(c, e);
    }
  }
  _checkCurrentSheet(e, n) {
    const o = this._sheetSkeletonManagerService.getCurrentSkeleton();
    if (!o) return !1;
    const r = o.worksheet;
    return r ? r.unitId === e && r.getSheetId() === n : !1;
  }
  _updateArrayFormulaRangeShape(e, n) {
    this._checkCurrentSheet(e, n) && this._previousShape && this._refreshArrayFormulaRangeShape(e);
  }
  _refreshArrayFormulaRangeShapeByRow(e, n, o) {
    if (!this._checkCurrentSheet(e, n) || !this._previousShape) return;
    const { startRow: r, endRow: s, startColumn: i, endColumn: c } = this._previousShape.getRange();
    for (let l = 0; l < o.length; l++) {
      const { row: m } = o[l];
      if (r >= m) {
        const g = {
          startRow: r,
          endRow: s,
          startColumn: i,
          endColumn: c
        };
        this._refreshArrayFormulaRangeShape(e, g);
        break;
      }
    }
  }
};
kt = ni([
  Ve(1, B(cn)),
  Ve(2, B(pt)),
  Ve(3, B(gt)),
  Ve(4, Ne),
  Ve(5, B($t)),
  Ve(6, me),
  Ve(7, qr)
], kt);
var ri = Object.getOwnPropertyDescriptor, oi = (t, e, n, o) => {
  for (var r = o > 1 ? void 0 : o ? ri(e, n) : e, s = t.length - 1, i; s >= 0; s--)
    (i = t[s]) && (r = i(r) || r);
  return r;
}, Wn = (t, e) => (n, o) => e(n, o, t);
const si = {
  tl: {
    size: 6,
    color: "#409f11"
  }
};
let At = class extends jr {
  constructor(t, e) {
    super(), this._sheetInterceptorService = t, this._formulaDataModel = e, this.disposeWithMe(this._sheetInterceptorService.intercept(
      Xo.CELL_CONTENT,
      {
        effect: Kr.Style,
        handler: (n, o, r) => {
          var c, l, m, g;
          const s = (g = (m = (l = (c = this._formulaDataModel.getArrayFormulaCellData()) == null ? void 0 : c[o.unitId]) == null ? void 0 : l[o.subUnitId]) == null ? void 0 : m[o.row]) == null ? void 0 : g[o.col];
          return !Jn(n, !!s) || !n || (n === o.rawData && (n = { ...o.rawData }), n.markers = {
            ...n == null ? void 0 : n.markers,
            ...si
          }), r(n);
        },
        priority: 10
      }
    ));
  }
};
At = oi([
  Wn(0, B(cn)),
  Wn(1, B(pt))
], At);
function ii() {
  const t = T(ss), e = T(me), n = pe(t.progress$), o = xt(() => {
    e.executeCommand(To.id);
  }, [e]), r = xt(() => {
    t.clearProgress();
  }, [t]);
  return /* @__PURE__ */ M(ls, { progress: n, onTerminate: o, onClearProgress: r });
}
function ci(t, e) {
  return Object.keys(t).filter((n) => isNaN(Number(n)) && n !== "DefinedName").map((n) => ({
    label: e.t(`formula.functionType.${n.toLocaleLowerCase()}`),
    value: `${t[n]}`
  }));
}
function Sr(t) {
  if (!t.require && !t.repeat)
    return `[${t.name}]`;
  if (t.require && !t.repeat)
    return t.name;
  if (!t.require && t.repeat)
    return `[${t.name},...]`;
  if (t.require && t.repeat)
    return `${t.name},...`;
}
function vr(t) {
  const { prefix: e, value: n } = t;
  return /* @__PURE__ */ V("div", { children: [
    /* @__PURE__ */ V("span", { children: [
      e,
      "("
    ] }),
    n && n.map((o, r) => /* @__PURE__ */ V("span", { children: [
      /* @__PURE__ */ M("span", { children: Sr(o) }),
      r === n.length - 1 ? "" : ","
    ] }, r)),
    ")"
  ] });
}
function ut(t) {
  const { className: e, value: n, title: o } = t;
  return /* @__PURE__ */ V("div", { className: "univer-mb-2 univer-text-xs", children: [
    /* @__PURE__ */ M(
      "div",
      {
        className: re("univer-mb-2 univer-font-medium univer-text-gray-500 dark:!univer-text-gray-300", e),
        children: o
      }
    ),
    /* @__PURE__ */ M(
      "div",
      {
        className: "univer-break-all univer-text-gray-900 dark:!univer-text-white",
        children: n
      }
    )
  ] });
}
function ai(t) {
  const { functionInfo: e, onChange: n } = t;
  if (!e) return null;
  const [o, r] = W([]), [s, i] = W(e.functionParameter), [c, l] = W(-1);
  return /* @__PURE__ */ V("div", { children: [
    /* @__PURE__ */ M("div", { className: re("univer-h-[364px] univer-overflow-y-auto", nt), children: s.map((m, g) => /* @__PURE__ */ V("div", { children: [
      /* @__PURE__ */ M("div", { className: "univer-text-sm", children: m.name }),
      /* @__PURE__ */ M("div", { className: "univer-mb-2 univer-mt-1" })
    ] }, g)) }),
    /* @__PURE__ */ M("div", { className: re("univer-flex-1 univer-p-3", Ss), children: /* @__PURE__ */ M(
      ut,
      {
        title: c === -1 ? /* @__PURE__ */ M(vr, { prefix: e.functionName, value: s }) : s[c].name,
        value: c === -1 ? e.description : s[c].detail
      }
    ) })
  ] });
}
function Ye({ ref: t, ...e }) {
  const { icon: n, id: o, className: r, extend: s, ...i } = e, c = `univerjs-icon univerjs-icon-${o} ${r || ""}`.trim(), l = X(`_${di()}`);
  return Cr(n, `${o}`, {
    defIds: n.defIds,
    idSuffix: l.current
  }, {
    ref: t,
    className: c,
    ...i
  }, s);
}
function Cr(t, e, n, o, r) {
  return Ke(t.tag, {
    key: e,
    ...li(t, n, r),
    ...o
  }, (ui(t, n).children || []).map((s, i) => Cr(s, `${e}-${t.tag}-${i}`, n, void 0, r)));
}
function li(t, e, n) {
  const o = { ...t.attrs };
  n != null && n.colorChannel1 && o.fill === "colorChannel1" && (o.fill = n.colorChannel1), t.tag === "mask" && o.id && (o.id = o.id + e.idSuffix), Object.entries(o).forEach(([s, i]) => {
    s === "mask" && typeof i == "string" && (o[s] = i.replace(/url\(#(.*)\)/, `url(#$1${e.idSuffix})`));
  });
  const { defIds: r } = e;
  return !r || r.length === 0 || (t.tag === "use" && o["xlink:href"] && (o["xlink:href"] = o["xlink:href"] + e.idSuffix), Object.entries(o).forEach(([s, i]) => {
    typeof i == "string" && (o[s] = i.replace(/url\(#(.*)\)/, `url(#$1${e.idSuffix})`));
  })), o;
}
function ui(t, e) {
  var o;
  const { defIds: n } = e;
  return !n || n.length === 0 ? t : t.tag === "defs" && ((o = t.children) != null && o.length) ? {
    ...t,
    children: t.children.map((r) => typeof r.attrs.id == "string" && n && n.includes(r.attrs.id) ? {
      ...r,
      attrs: {
        ...r.attrs,
        id: r.attrs.id + e.idSuffix
      }
    } : r)
  } : t;
}
function di() {
  return Math.random().toString(36).substring(2, 8);
}
Ye.displayName = "UniverIcon";
const hi = {
  tag: "svg",
  attrs: {
    xmlns: "http://www.w3.org/2000/svg",
    fill: "none",
    viewBox: "0 0 16 16",
    width: "1em",
    height: "1em"
  },
  children: [{
    tag: "path",
    attrs: {
      fill: "currentColor",
      d: "M14.1544 3.75557C14.3887 3.98988 14.3887 4.36978 14.1544 4.6041L6.51409 12.2444C6.40157 12.3569 6.24896 12.4201 6.08983 12.4201C5.9307 12.4201 5.77808 12.3569 5.66556 12.2444L1.84541 8.42425C1.6111 8.18993 1.6111 7.81003 1.84541 7.57572C2.07973 7.34141 2.45963 7.34141 2.69394 7.57572L6.08983 10.9716L13.3059 3.75557C13.5402 3.52126 13.9201 3.52126 14.1544 3.75557Z",
      fillRule: "evenodd",
      clipRule: "evenodd"
    }
  }]
}, _r = Le(function(e, n) {
  return Ke(Ye, Object.assign({}, e, {
    id: "check-mark-icon",
    ref: n,
    icon: hi
  }));
});
_r.displayName = "CheckMarkIcon";
const fi = {
  tag: "svg",
  attrs: {
    xmlns: "http://www.w3.org/2000/svg",
    fill: "none",
    viewBox: "0 0 16 16",
    width: "1em",
    height: "1em"
  },
  children: [{
    tag: "path",
    attrs: {
      fill: "currentColor",
      d: "M3.71274 2.86421C3.47843 2.6299 3.09853 2.6299 2.86421 2.86421C2.6299 3.09853 2.6299 3.47843 2.86421 3.71274L7.15154 8.00007L2.86421 12.2874C2.6299 12.5217 2.6299 12.9016 2.86421 13.1359C3.09853 13.3702 3.47843 13.3702 3.71274 13.1359L8.00007 8.84859L12.2874 13.1359C12.5217 13.3702 12.9016 13.3702 13.1359 13.1359C13.3702 12.9016 13.3702 12.5217 13.1359 12.2874L8.84859 8.00007L13.1359 3.71274C13.3702 3.47843 13.3702 3.09853 13.1359 2.86421C12.9016 2.6299 12.5217 2.6299 12.2874 2.86421L8.00007 7.15154L3.71274 2.86421Z"
    }
  }]
}, Rr = Le(function(e, n) {
  return Ke(Ye, Object.assign({}, e, {
    id: "close-icon",
    ref: n,
    icon: fi
  }));
});
Rr.displayName = "CloseIcon";
const gi = {
  tag: "svg",
  attrs: {
    xmlns: "http://www.w3.org/2000/svg",
    fill: "none",
    viewBox: "0 0 16 16",
    width: "1em",
    height: "1em"
  },
  children: [
    {
      tag: "path",
      attrs: {
        fill: "currentColor",
        d: "M5.3313 1.4667C5.3313 1.13533 5.59993 0.866699 5.9313 0.866699H10.069C10.4004 0.866699 10.669 1.13533 10.669 1.4667C10.669 1.79807 10.4004 2.0667 10.069 2.0667H5.9313C5.59993 2.0667 5.3313 1.79807 5.3313 1.4667Z"
      }
    },
    {
      tag: "path",
      attrs: {
        fill: "currentColor",
        d: "M1.09985 3.64443C1.09985 3.31306 1.36848 3.04443 1.69985 3.04443H14.2999C14.6312 3.04443 14.8999 3.31306 14.8999 3.64443C14.8999 3.9758 14.6312 4.24443 14.2999 4.24443H1.69985C1.36848 4.24443 1.09985 3.9758 1.09985 3.64443Z"
      }
    },
    {
      tag: "path",
      attrs: {
        fill: "currentColor",
        d: "M6.12398 8.30171C6.35829 8.0674 6.73819 8.0674 6.97251 8.30171L8.00007 9.32928L9.02764 8.30171C9.26195 8.0674 9.64185 8.0674 9.87617 8.30171C10.1105 8.53603 10.1105 8.91593 9.87617 9.15024L8.8486 10.1778L9.87617 11.2054C10.1105 11.4397 10.1105 11.8196 9.87617 12.0539C9.64185 12.2882 9.26195 12.2882 9.02764 12.0539L8.00007 11.0263L6.97251 12.0539C6.73819 12.2882 6.35829 12.2882 6.12398 12.0539C5.88966 11.8196 5.88966 11.4397 6.12398 11.2054L7.15154 10.1778L6.12398 9.15024C5.88966 8.91593 5.88966 8.53603 6.12398 8.30171Z"
      }
    },
    {
      tag: "path",
      attrs: {
        fill: "currentColor",
        d: "M4.75332 5.22217C3.86966 5.22217 3.15332 5.93851 3.15332 6.82217V12.5331C3.15332 13.9691 4.31738 15.1332 5.75332 15.1332H10.2465C11.6825 15.1332 12.8465 13.9691 12.8465 12.5331V6.82217C12.8465 5.93851 12.1302 5.22217 11.2465 5.22217H4.75332ZM4.35332 6.82217C4.35332 6.60125 4.53241 6.42217 4.75332 6.42217H11.2465C11.4674 6.42217 11.6465 6.60125 11.6465 6.82217V12.5331C11.6465 13.3063 11.0197 13.9332 10.2465 13.9332H5.75332C4.98012 13.9332 4.35332 13.3063 4.35332 12.5331V6.82217Z",
        fillRule: "evenodd",
        clipRule: "evenodd"
      }
    }
  ]
}, Er = Le(function(e, n) {
  return Ke(Ye, Object.assign({}, e, {
    id: "delete-icon",
    ref: n,
    icon: gi
  }));
});
Er.displayName = "DeleteIcon";
const mi = {
  tag: "svg",
  attrs: {
    xmlns: "http://www.w3.org/2000/svg",
    fill: "none",
    viewBox: "0 0 16 16",
    width: "1em",
    height: "1em"
  },
  children: [{
    tag: "path",
    attrs: {
      fill: "currentColor",
      d: "M8.6 1.99991C8.60001 1.66854 8.33138 1.39991 8.00001 1.3999C7.66864 1.3999 7.40001 1.66853 7.4 1.9999L7.39996 7.3999H1.9999C1.66853 7.3999 1.3999 7.66853 1.3999 7.9999C1.3999 8.33127 1.66853 8.5999 1.9999 8.5999H7.39995L7.3999 13.9999C7.3999 14.3313 7.66853 14.5999 7.9999 14.5999C8.33127 14.5999 8.5999 14.3313 8.5999 13.9999L8.59995 8.5999H13.9999C14.3313 8.5999 14.5999 8.33127 14.5999 7.9999C14.5999 7.66853 14.3313 7.3999 13.9999 7.3999H8.59996L8.6 1.99991Z"
    }
  }]
}, Ir = Le(function(e, n) {
  return Ke(Ye, Object.assign({}, e, {
    id: "increase-icon",
    ref: n,
    icon: mi
  }));
});
Ir.displayName = "IncreaseIcon";
const pi = {
  tag: "svg",
  attrs: {
    xmlns: "http://www.w3.org/2000/svg",
    fill: "none",
    viewBox: "0 0 16 16",
    width: "1em",
    height: "1em"
  },
  children: [{
    tag: "path",
    attrs: {
      fill: "currentColor",
      d: "M5.90913 3.57564C6.14345 3.34132 6.52335 3.34132 6.75766 3.57564L10.7577 7.57564C10.992 7.80995 10.992 8.18985 10.7577 8.42417L6.75766 12.4242C6.52335 12.6585 6.14345 12.6585 5.90913 12.4242C5.67482 12.1899 5.67482 11.81 5.90913 11.5756L9.48487 7.9999L5.90913 4.42417C5.67482 4.18985 5.67482 3.80995 5.90913 3.57564Z",
      fillRule: "evenodd",
      clipRule: "evenodd"
    }
  }]
}, br = Le(function(e, n) {
  return Ke(Ye, Object.assign({}, e, {
    id: "more-icon",
    ref: n,
    icon: pi
  }));
});
br.displayName = "MoreIcon";
const Si = {
  tag: "svg",
  attrs: {
    xmlns: "http://www.w3.org/2000/svg",
    fill: "none",
    viewBox: "0 0 16 16",
    width: "1em",
    height: "1em"
  },
  children: [{
    tag: "path",
    attrs: {
      fill: "currentColor",
      d: "M12.6185 12.4423C12.5907 12.2749 12.7773 12.15 12.9343 12.2308L15.4242 13.5126C15.6102 13.6084 15.5544 13.8745 15.3439 13.8955L14.2456 14.184L13.4521 15.1286C13.3495 15.2939 13.085 15.2463 13.0534 15.0568L12.6185 12.4423Z"
    }
  }, {
    tag: "path",
    attrs: {
      fill: "currentColor",
      d: "M1 3.6C1 2.16406 2.16406 1 3.6 1H12.3C13.7359 1 14.9 2.16406 14.9 3.6V5.81156C14.9003 5.81881 14.9004 5.82609 14.9004 5.8334C14.9004 5.84071 14.9003 5.84799 14.9 5.85524V10.045C14.9003 10.0522 14.9004 10.0595 14.9004 10.0668C14.9004 10.3982 14.6318 10.6668 14.3004 10.6668H11.1668C10.8907 10.6668 10.6668 10.8907 10.6668 11.1668V14.3C10.6668 14.6314 10.3982 14.9 10.0668 14.9L10.05 14.8998L3.6 14.9C2.16406 14.9 1 13.7359 1 12.3V3.6ZM13.2 5.2334C13.4761 5.2334 13.7 5.00954 13.7 4.7334V3.6C13.7 2.8268 13.0732 2.2 12.3 2.2H11.1668C10.8907 2.2 10.6668 2.42386 10.6668 2.7V4.7334C10.6668 5.00954 10.8907 5.2334 11.1668 5.2334H13.2ZM10.6668 6.9334C10.6668 6.65726 10.8907 6.4334 11.1668 6.4334H13.2C13.4761 6.4334 13.7 6.65726 13.7 6.9334V8.9668C13.7 9.24294 13.4761 9.4668 13.2 9.4668H11.1668C10.8907 9.4668 10.6668 9.24294 10.6668 8.9668V6.9334ZM8.9668 5.2334C9.24294 5.2334 9.4668 5.00954 9.4668 4.7334V2.7C9.4668 2.42386 9.24294 2.2 8.9668 2.2H6.9334C6.65726 2.2 6.4334 2.42386 6.4334 2.7V4.7334C6.4334 5.00954 6.65726 5.2334 6.9334 5.2334L8.9668 5.2334ZM6.4334 6.9334C6.4334 6.65726 6.65726 6.4334 6.9334 6.4334L8.9668 6.4334C9.24294 6.4334 9.4668 6.65726 9.4668 6.9334V8.9668C9.4668 9.24294 9.24294 9.4668 8.9668 9.4668L6.9334 9.4668C6.65726 9.4668 6.4334 9.24294 6.4334 8.9668V6.9334ZM4.7334 5.2334C5.00954 5.2334 5.2334 5.00954 5.2334 4.7334V2.7C5.2334 2.42386 5.00954 2.2 4.7334 2.2H3.6C2.8268 2.2 2.2 2.8268 2.2 3.6V4.7334C2.2 5.00954 2.42386 5.2334 2.7 5.2334H4.7334ZM2.2 6.9334C2.2 6.65726 2.42386 6.4334 2.7 6.4334H4.7334C5.00954 6.4334 5.2334 6.65725 5.2334 6.9334V8.9668C5.2334 9.24294 5.00954 9.4668 4.7334 9.4668H2.7C2.42386 9.4668 2.2 9.24294 2.2 8.9668V6.9334ZM5.2334 11.1668C5.2334 10.8907 5.00954 10.6668 4.7334 10.6668H2.7C2.42386 10.6668 2.2 10.8907 2.2 11.1668V12.3C2.2 13.0732 2.8268 13.7 3.6 13.7H4.7334C5.00954 13.7 5.2334 13.4761 5.2334 13.2V11.1668ZM9.4668 11.1668C9.4668 10.8907 9.24294 10.6668 8.9668 10.6668H6.9334C6.65726 10.6668 6.4334 10.8907 6.4334 11.1668V13.2C6.4334 13.4761 6.65726 13.7 6.9334 13.7H8.9668C9.24294 13.7 9.4668 13.4761 9.4668 13.2V11.1668Z",
      fillRule: "evenodd",
      clipRule: "evenodd"
    }
  }]
}, yr = Le(function(e, n) {
  return Ke(Ye, Object.assign({}, e, {
    id: "select-range-icon",
    ref: n,
    icon: Si
  }));
});
yr.displayName = "SelectRangeIcon";
function vi(t) {
  const { onChange: e } = t, n = "-1", [o, r] = W(""), [s, i] = W([]), [c, l] = W(0), [m, g] = W(n), [h, u] = W(0), [d, a] = W(null), v = T(Ut), S = T(ot), p = T(an), C = pe(p.sidebarOptions$), _ = ci(er, S);
  _.unshift({
    label: S.t("formula.moreFunctions.allFunctions"),
    value: n
  });
  const I = S.t("formula.prompt.required"), f = S.t("formula.prompt.optional");
  $(() => {
    A(n);
  }, []), $(() => {
    y(0);
  }, [s]), $(() => {
    C != null && C.visible && (r(""), i([]), l(0), g(n), u(0), a(null), A(n));
  }, [C]);
  const b = (E) => {
    if (o.trim() === "") return E;
    const O = new RegExp(`(${o.toLocaleUpperCase()})`);
    return E.split(O).filter(Boolean).map((w, L) => w.match(O) ? /* @__PURE__ */ M("span", { className: "univer-text-red-500", children: w }, L) : w);
  }, y = (E) => {
    if (s.length === 0) {
      a(null);
      return;
    }
    u(E);
    const O = v.getFunctionInfo(s[E].name);
    if (!O) {
      a(null);
      return;
    }
    a(O), e(O);
  };
  function A(E) {
    g(E);
    const O = v.getSearchListByType(+E);
    i(O);
  }
  function D(E) {
    r(E);
    const O = v.getSearchListByName(E);
    i(O);
  }
  function x(E) {
    if (E.stopPropagation(), E.key === "ArrowDown") {
      const O = c + 1;
      l(O === s.length ? 0 : O);
    } else if (E.key === "ArrowUp") {
      const O = c - 1;
      l(O === -1 ? s.length - 1 : O);
    } else E.key === "Enter" && y(c);
  }
  const N = (E) => {
    l(E);
  }, k = () => {
    l(-1);
  };
  return /* @__PURE__ */ V("div", { children: [
    /* @__PURE__ */ V("div", { className: "univer-flex univer-items-center univer-justify-between univer-gap-2", children: [
      /* @__PURE__ */ M(vs, { value: m, options: _, onChange: A }),
      /* @__PURE__ */ M(
        ur,
        {
          placeholder: S.t("formula.moreFunctions.searchFunctionPlaceholder"),
          onKeyDown: x,
          value: o,
          onChange: D,
          size: "small",
          allowClear: !0
        }
      )
    ] }),
    s.length > 0 && /* @__PURE__ */ M(
      "ul",
      {
        className: re("univer-mb-0 univer-mt-2 univer-box-border univer-max-h-72 univer-w-full univer-select-none univer-list-none univer-overflow-y-auto univer-rounded univer-p-3 univer-outline-none", ln, nt),
        onKeyDown: x,
        tabIndex: -1,
        children: s.map(({ name: E }, O) => /* @__PURE__ */ V(
          "li",
          {
            className: re("univer-relative univer-box-border univer-cursor-pointer univer-rounded univer-px-7 univer-py-1 univer-text-sm univer-text-gray-900 univer-transition-colors dark:!univer-text-white", {
              "univer-bg-gray-200 dark:!univer-bg-gray-600": c === O
            }),
            onMouseEnter: () => N(O),
            onMouseLeave: k,
            onClick: () => y(O),
            children: [
              h === O && /* @__PURE__ */ M(
                _r,
                {
                  className: "univer-absolute univer-left-1.5 univer-top-1/2 univer-inline-flex -univer-translate-y-1/2 univer-text-base univer-text-primary-600"
                }
              ),
              /* @__PURE__ */ M("span", { className: "univer-block", children: b(E) })
            ]
          },
          O
        ))
      }
    ),
    d && /* @__PURE__ */ V("div", { className: re("univer-mx-0 univer-my-2 univer-overflow-y-auto", nt), children: [
      /* @__PURE__ */ M(ut, { title: d.functionName, value: d.description }),
      /* @__PURE__ */ M(
        ut,
        {
          title: S.t("formula.moreFunctions.syntax"),
          value: /* @__PURE__ */ M(vr, { prefix: d.functionName, value: d.functionParameter })
        }
      ),
      /* @__PURE__ */ M(
        ut,
        {
          title: S.t("formula.prompt.helpExample"),
          value: `${d.functionName}(${d.functionParameter.map((E) => E.example).join(",")})`
        }
      ),
      d.functionParameter && d.functionParameter.map((E) => /* @__PURE__ */ M(
        ut,
        {
          title: E.name,
          value: `${E.require ? I : f} ${E.detail}`
        },
        E.name
      ))
    ] })
  ] });
}
function Ci() {
  const t = ho(), [e, n] = W(!0), [o, r] = W(!1), [s, i] = W(null);
  T(nn);
  const c = T(ot), l = T(je), m = T(ne), g = T(me);
  function h() {
    n(!e), r(!o);
  }
  function u() {
    const d = ir(m);
    if (!d) return;
    g.executeCommand(zn.id, {
      visible: !0,
      unitId: d.unitId,
      eventType: xe.Dblclick
    });
    const a = l.getEditor(Jt), v = l.getEditor(jn), S = `=${s == null ? void 0 : s.functionName}(`;
    a == null || a.replaceText(S), v == null || v.replaceText(S, !1);
  }
  return /* @__PURE__ */ V(
    "div",
    {
      "data-u-comp": "sheets-formula-functions-panel",
      className: "univer-box-border univer-flex univer-h-full univer-flex-col univer-justify-between univer-py-2",
      children: [
        e && /* @__PURE__ */ M(vi, { onChange: i }),
        o && /* @__PURE__ */ M(ai, { functionInfo: s, onChange: () => {
        } }),
        /* @__PURE__ */ V("div", { className: "univer-flex univer-justify-end", children: [
          o && /* @__PURE__ */ M(
            et,
            {
              variant: "primary",
              onClick: h,
              className: "univer-mb-5 univer-ml-4 univer-mr-0 univer-mt-0",
              children: c.t("formula.moreFunctions.next")
            }
          ),
          o && /* @__PURE__ */ M(et, { onClick: h, className: "univer-mb-5 univer-ml-4 univer-mr-0 univer-mt-0", children: c.t("formula.moreFunctions.prev") }),
          e && !!t && /* @__PURE__ */ M(
            et,
            {
              variant: "primary",
              onClick: u,
              className: "univer-mb-5 univer-ml-4 univer-mr-0 univer-mt-0",
              children: c.t("formula.moreFunctions.confirm")
            }
          )
        ] })
      ]
    }
  );
}
function _i(t) {
  return t.get(ne).getCurrentTypeOfUnit$(K.UNIVER_SHEET).pipe(
    Lo((o) => o ? t.get(Ln) ? new Fo((s) => s.next(!t.get(Ln).supportClipboard)) : On(!0) : On(!0))
  );
}
function Ri(t) {
  return {
    id: un.id,
    type: us.BUTTON,
    title: "formula.operation.pasteFormula",
    disabled$: _i(t).pipe(
      Do(fo(t, {
        workbookTypes: [ts],
        rangeTypes: [es],
        worksheetTypes: [Qo, Jo]
      })),
      nr(([e, n]) => e || n)
    )
  };
}
const Ei = {
  [go]: {
    [un.id]: {
      order: 4,
      menuItemFactory: Ri
    }
  }
}, Ii = "meta_key_ctrl_And_Shift";
function bi(t) {
  return t.getContextValue(Yr) && t.getContextValue(Zr);
}
const Ht = [
  F.ARROW_DOWN,
  F.ARROW_UP,
  F.ARROW_LEFT,
  F.ARROW_RIGHT
], yi = [...Ht, F.ENTER, F.TAB, F.ESC];
function Ti() {
  const t = [];
  for (const e of yi)
    t.push({
      id: it.id,
      binding: e,
      preconditions: (n) => mt(n),
      staticParameters: {
        eventType: xe.Keyboard,
        keycode: e
      }
    });
  return t;
}
function xi() {
  const t = [];
  for (const e of Ht)
    t.push({
      id: it.id,
      binding: e | P.SHIFT,
      preconditions: (n) => mt(n),
      staticParameters: {
        eventType: xe.Keyboard,
        keycode: e,
        metaKey: P.SHIFT
      }
    });
  return t;
}
function Ni() {
  const t = [];
  for (const e of Ht)
    t.push({
      id: it.id,
      binding: e | P.CTRL_COMMAND,
      preconditions: (n) => mt(n),
      staticParameters: {
        eventType: xe.Keyboard,
        keycode: e,
        metaKey: P.CTRL_COMMAND
      }
    });
  return t;
}
function Oi() {
  const t = [];
  for (const e of Ht)
    t.push({
      id: it.id,
      binding: e | P.SHIFT | P.CTRL_COMMAND,
      preconditions: (n) => mt(n),
      staticParameters: {
        eventType: xe.Keyboard,
        keycode: e,
        metaKey: Ii
      }
    });
  return t;
}
const Mi = {
  id: gr.id,
  binding: F.F4,
  preconditions: (t) => mt(t)
};
function wi() {
  const t = [];
  for (const e of [F.ENTER, F.TAB, F.ARROW_DOWN, F.ARROW_UP])
    t.push({
      id: it.id,
      binding: e,
      preconditions: (n) => bi(n),
      staticParameters: {
        eventType: xe.Keyboard,
        keycode: e,
        isSingleEditor: !0
      }
    });
  return t;
}
const ki = {
  id: is.id,
  binding: P.ALT | F.EQUAL,
  preconditions: mo,
  mac: P.CTRL_COMMAND | P.ALT | F.EQUAL,
  description: "shortcut.sheets-formula-ui.quick-sum",
  group: "4_sheet-edit"
};
var Ai = Object.getOwnPropertyDescriptor, Di = (t, e, n, o) => {
  for (var r = o > 1 ? void 0 : o ? Ai(e, n) : e, s = t.length - 1, i; s >= 0; s--)
    (i = t[s]) && (r = i(r) || r);
  return r;
}, He = (t, e) => (n, o) => e(n, o, t);
let Dt = class extends rt {
  constructor(t, e, n, o, r, s, i) {
    super(), this._injector = t, this._menuManagerService = e, this._commandService = n, this._shortcutService = o, this._uiPartsService = r, this._renderManagerService = s, this._componentManager = i, this._initialize();
  }
  _initialize() {
    this._registerCommands(), this._registerMenus(), this._registerShortcuts(), this._registerComponents(), this._registerRenderModules();
  }
  _registerMenus() {
    this._menuManagerService.mergeMenu(Ei);
  }
  _registerCommands() {
    [
      un,
      Ns,
      Ls,
      Fs,
      xs,
      it,
      gr
    ].forEach((t) => this.disposeWithMe(this._commandService.registerCommand(t)));
  }
  _registerShortcuts() {
    [
      ...Ti(),
      ...xi(),
      ...Ni(),
      ...Oi(),
      ...wi(),
      ki,
      Mi
    ].forEach((t) => {
      this.disposeWithMe(this._shortcutService.registerShortcut(t));
    });
  }
  _registerComponents() {
    this.disposeWithMe(this._uiPartsService.registerComponent(po.FORMULA_AUX, () => lr(ii, this._injector))), this._componentManager.register(fr, Ci);
  }
  _registerRenderModules() {
    this.disposeWithMe(this._renderManagerService.registerRenderModule(K.UNIVER_SHEET, [kt]));
  }
};
Dt = Di([
  He(0, B(ft)),
  He(1, ds),
  He(2, me),
  He(3, Wt),
  He(4, cr),
  He(5, Ne),
  He(6, B(ar))
], Dt);
class Tr {
  constructor() {
    q(this, "_currentSelector$", new rr(null));
    q(this, "currentSelector$", this._currentSelector$.asObservable());
  }
  showRangeSelectorDialog(e) {
    const n = e.callback, o = new Promise((r) => {
      e.callback = (s) => {
        r(s), n(s);
      };
    });
    return this._currentSelector$.next(e), o;
  }
}
var Li = Object.getOwnPropertyDescriptor, Fi = (t, e, n, o) => {
  for (var r = o > 1 ? void 0 : o ? Li(e, n) : e, s = t.length - 1, i; s >= 0; s--)
    (i = t[s]) && (r = i(r) || r);
  return r;
}, Qe = (t, e) => (n, o) => e(n, o, t);
let De = class extends So {
  constructor(e, n, o, r, s, i, c) {
    super(
      n,
      o,
      r,
      s,
      i
    );
    q(this, "_workbookSelections");
    q(this, "_eventDisposables");
    this._context = e, this._contextService = i, this._refSelectionsService = c, this._workbookSelections = this._refSelectionsService.getWorkbookSelections(this._context.unitId), this._initSelectionChangeListener(), this._initSkeletonChangeListener(), this._initUserActionSyncListener(), this._setSelectionStyle($i(this._themeService)), this._remainLastEnabled = !0, this._highlightHeader = !1;
  }
  getLocation() {
    return this._skeleton.getLocation();
  }
  setRemainLastEnabled(e) {
    this._remainLastEnabled = e;
  }
  /**
   * This is set to true when you need to add a new selection.
   * @param {boolean} enabled
   * @memberof RefSelectionsRenderService
   */
  setSkipLastEnabled(e) {
    this._skipLastEnabled = e;
  }
  clearLastSelection() {
    const e = this._selectionControls[this._selectionControls.length - 1];
    e && (e.dispose(), this._selectionControls.pop());
  }
  /**
   * Call this method and user will be able to select on the canvas to update selections.
   */
  enableSelectionChanging() {
    return this._disableSelectionChanging(), this._eventDisposables = this._initCanvasEventListeners(), Yn(() => this._disableSelectionChanging());
  }
  _disableSelectionChanging() {
    var e;
    (e = this._eventDisposables) == null || e.dispose(), this._eventDisposables = null;
  }
  disableSelectionChanging() {
    this._disableSelectionChanging();
  }
  _initCanvasEventListeners() {
    const e = this._getSheetObject(), { spreadsheetRowHeader: n, spreadsheetColumnHeader: o, spreadsheet: r, spreadsheetLeftTopPlaceholder: s } = e, { scene: i } = this._context, c = new Be();
    return c.add(r == null ? void 0 : r.onPointerDown$.subscribeEvent((l, m) => {
      this.inRefSelectionMode() && (this._onPointerDown(l, r.zIndex + 1, Se.NORMAL, this._getActiveViewport(l)), l.button !== 2 && m.stopPropagation());
    })), c.add(
      n == null ? void 0 : n.onPointerDown$.subscribeEvent((l, m) => {
        if (!this.inRefSelectionMode()) return;
        const g = this._sheetSkeletonManagerService.getCurrent().skeleton, { row: h } = bn(l.offsetX, l.offsetY, i, g);
        yn(this._workbookSelections.getCurrentSelections(), h, Se.ROW) || (this._onPointerDown(l, (r.zIndex || 1) + 1, Se.ROW, this._getActiveViewport(l), Bt.Y), l.button !== 2 && m.stopPropagation());
      })
    ), c.add(o == null ? void 0 : o.onPointerDown$.subscribeEvent((l, m) => {
      if (!this.inRefSelectionMode()) return;
      const g = this._sheetSkeletonManagerService.getCurrent().skeleton, { column: h } = bn(l.offsetX, l.offsetY, i, g);
      yn(this._workbookSelections.getCurrentSelections(), h, Se.COLUMN) || (this._onPointerDown(l, (r.zIndex || 1) + 1, Se.COLUMN, this._getActiveViewport(l), Bt.X), l.button !== 2 && m.stopPropagation());
    })), c.add(s == null ? void 0 : s.onPointerDown$.subscribeEvent((l, m) => {
      if (this._reset(), !this.inRefSelectionMode()) return;
      const g = this._sheetSkeletonManagerService.getCurrent().skeleton, h = vo(g);
      this._addSelectionControlByModelData(h), this._selectionMoveStart$.next(this.getSelectionDataWithStyle());
      const u = i.onPointerUp$.subscribeEvent(() => {
        u.unsubscribe(), this._selectionMoveEnd$.next(this.getSelectionDataWithStyle());
      });
      l.button !== 2 && m.stopPropagation();
    })), c;
  }
  /**
   * Add a selection in spreadsheet, create a new SelectionControl and then update this control by range derives from selection.
   * For ref selection, create selectionShapeExtension to handle user action.
   * @param {ISelectionWithCoord} selectionWithStyle
   */
  _addSelectionControlByModelData(e) {
    var i;
    const n = this._skeleton, o = (i = e.style) != null ? i : Qn(this._themeService), r = this._scene;
    return e.style = o, this.newSelectionControl(r, n, e);
  }
  _initSelectionChangeListener() {
    this.disposeWithMe(this._refSelectionsService.selectionSet$.subscribe((e) => {
      this._reset(), this._skeleton && this.resetSelectionsByModelData(e || []);
    }));
  }
  /**
   * Update selectionModel in this._workbookSelections by user action in spreadsheet area.
   */
  _initUserActionSyncListener() {
    this.disposeWithMe(this.selectionMoveStart$.subscribe((e) => {
      this._updateSelections(e, qt.MOVE_START);
    })), this.disposeWithMe(this.selectionMoving$.subscribe((e) => {
      this._updateSelections(e, qt.MOVING);
    })), this.disposeWithMe(this.selectionMoveEnd$.subscribe((e) => {
      this._updateSelections(e, qt.MOVE_END);
    }));
  }
  _updateSelections(e, n) {
    const r = this._context.unit.getActiveSheet().getSheetId();
    e.length !== 0 && this._workbookSelections.setSelections(
      r,
      e.map((s) => ns(s)),
      n
    );
  }
  _initSkeletonChangeListener() {
    this.disposeWithMe(this._sheetSkeletonManagerService.currentSkeleton$.subscribe((e) => {
      var i;
      if (!e)
        return;
      const { skeleton: n } = e, { scene: o } = this._context, r = o.getViewport(wn.VIEW_MAIN);
      this._skeleton && ((i = this._skeleton.worksheet) == null ? void 0 : i.getSheetId()) !== n.worksheet.getSheetId() && this._reset(), this._changeRuntime(n, o, r);
      const s = this._workbookSelections.getCurrentSelections();
      this.resetSelectionsByModelData(s);
    }));
  }
  _getActiveViewport(e) {
    const n = this._getSheetObject();
    return n == null ? void 0 : n.scene.getActiveViewportByCoord(kn.FromArray([e.offsetX, e.offsetY]));
  }
  _getSheetObject() {
    return Co(this._context.unit, this._context);
  }
  /**
   * Handle pointer down event, bind pointermove & pointerup handler.
   * then trigger selectionMoveStart$.
   *
   * @param evt
   * @param _zIndex
   * @param rangeType
   * @param viewport
   * @param scrollTimerType
   */
  // eslint-disable-next-line complexity, max-lines-per-function
  _onPointerDown(e, n = 0, o = Se.NORMAL, r, s = Bt.ALL) {
    var x;
    this._rangeType = o;
    const i = this._skeleton, c = this._scene;
    if (!c || !i)
      return;
    r && (this._activeViewport = r);
    const { offsetX: l, offsetY: m } = e, g = c.getViewport(wn.VIEW_MAIN);
    if (!g) return;
    const h = c.getCoordRelativeToViewport(kn.FromArray([l, m])), { x: u, y: d } = h;
    this._startViewportPosX = u, this._startViewportPosY = d;
    const a = c.getScrollXYInfoByViewport(h), { scaleX: v, scaleY: S } = c.getAncestorScale(), p = this._skeleton.getCellByOffset(u, d, v, S, a);
    if (!p) return;
    switch (o) {
      case Se.NORMAL:
        break;
      case Se.ROW:
        p.startColumn = 0, p.endColumn = this._skeleton.getColumnCount() - 1;
        break;
      case Se.COLUMN:
        p.startRow = 0, p.endRow = this._skeleton.getRowCount() - 1;
        break;
      case Se.ALL:
        p.startRow = 0, p.startColumn = 0, p.endRow = this._skeleton.getRowCount() - 1, p.endColumn = this._skeleton.getColumnCount() - 1;
    }
    const C = { range: p, primary: p, style: null };
    C.range.rangeType = o;
    const _ = Yt(C, this._skeleton);
    this._startRangeWhenPointerDown = { ..._.rangeWithCoord };
    const I = { ..._.rangeWithCoord, rangeType: o };
    let f = this.getActiveSelectionControl();
    const b = this.getSelectionControls();
    for (const N of b) {
      if (e.button === 2 && qn.contains(N.model, I)) {
        f = N;
        return;
      }
      if (N.model.isEqual(I)) {
        f = N;
        break;
      }
    }
    this._checkClearPreviousControls(e);
    const y = f == null ? void 0 : f.model.currentCell, A = e.shiftKey && y, D = this._remainLastEnabled && !e.ctrlKey && !e.shiftKey && !this._skipLastEnabled && !this._singleSelectionEnabled;
    A && y ? this._makeSelectionByTwoCells(
      y,
      I,
      i,
      o,
      f
      // Get updated in this method
    ) : D && f ? f.updateRangeBySelectionWithCoord(_) : f = this.newSelectionControl(c, i, C);
    for (let N = 0; N < this.getSelectionControls().length - 1; N++)
      this.getSelectionControls()[N].clearHighlight();
    this._selectionMoveStart$.next(this.getSelectionDataWithStyle()), c.disableObjectsEvent(), this._clearUpdatingListeners(), this._addEndingListeners(), (x = c.getTransformer()) == null || x.clearSelectedObjects(), this._setupPointerMoveListener(g, f, o, s, u, d), this._escapeShortcutDisposable = this._shortcutService.forceEscape(), this._scenePointerUpSub = c.onPointerUp$.subscribeEvent(() => {
      var N;
      this._clearUpdatingListeners(), this._selectionMoveEnd$.next(this.getSelectionDataWithStyle()), (N = this._escapeShortcutDisposable) == null || N.dispose(), this._escapeShortcutDisposable = null;
    });
  }
  /**
   * Diff between normal selection, no highlightHeader for ref selections.
   * @param scene
   * @param skeleton
   * @param selectionWithCoord
   * @returns {SelectionControl} selectionControl just created
   */
  newSelectionControl(e, n, o) {
    const r = this.getSelectionControls().length, { rowHeaderWidth: s, columnHeaderHeight: i } = n, c = new Xn(e, r, this._themeService, {
      highlightHeader: this._highlightHeader,
      enableAutoFill: !1,
      rowHeaderWidth: s,
      columnHeaderHeight: i
    }), l = Yt(o, n);
    return c.updateRangeBySelectionWithCoord(l), this._selectionControls.push(c), c.setControlExtension({
      skeleton: n,
      scene: e,
      themeService: this._themeService,
      injector: this._injector,
      selectionHooks: {
        selectionMoveEnd: () => {
          this._selectionMoveEnd$.next(this.getSelectionDataWithStyle());
        }
      }
    }), c;
  }
};
De = Fi([
  Qe(1, B(ft)),
  Qe(2, B(gt)),
  Qe(3, Wt),
  Qe(4, B($t)),
  Qe(5, Qt),
  Qe(6, Pt)
], De);
function $i(t) {
  const e = Qn(t);
  return e.widgets = { tl: !0, tc: !0, tr: !0, ml: !0, mr: !0, bl: !0, bc: !0, br: !0 }, e;
}
const dn = (t, e, n = !0) => {
  let o = -1;
  return t.reduce((r, s, i) => {
    if (r.isFinish)
      return r;
    const c = r.currentIndex;
    if (typeof s != "string")
      r.currentIndex += s.token.length;
    else {
      const l = s.length;
      r.currentIndex += l;
    }
    return (n ? r.currentIndex === e : e > c && e <= r.currentIndex) && (o = i, r.isFinish = !0), r;
  }, { currentIndex: 0, isFinish: !1 }), o;
}, xr = (t, e) => {
  const n = t[e];
  let o = -1;
  if (!n || typeof n == "string" || n.nodeType !== j.REFERENCE) return -1;
  for (let r = 0; r <= e; r++) {
    const s = t[r];
    typeof s != "string" && s.nodeType === j.REFERENCE && o++;
  }
  return o;
}, Pi = (t, e = 100) => {
  $(() => {
    let n = null;
    const o = () => {
      n === null && (n = window.setTimeout(() => {
        t(), n = null;
      }, e));
    };
    return window.addEventListener("scroll", o), window.addEventListener("resize", o), () => {
      n !== null && clearTimeout(n), window.removeEventListener("scroll", o), window.removeEventListener("resize", o);
    };
  }, [t, e]);
};
function Nr(t, e, n) {
  const o = T(je), r = ie(() => new rr({ left: -999, top: -999, right: -999, bottom: -999 }), []), s = T(an), i = T(ne), c = te(() => {
    var _;
    const l = o.getEditor(t);
    if (!l)
      return;
    const m = l.getBoundingClientRect(), { marginTop: g = 0, marginBottom: h = 0 } = l.getDocumentData().documentStyle, u = l.getSkeleton();
    if (!u) return;
    const d = (_ = u.getSkeletonData()) == null ? void 0 : _.pages[0].height;
    let { left: a, top: v, right: S, bottom: p } = m;
    v = v + g, p = d ? v + d : p - h;
    const C = r.getValue();
    if (!(C.left === a && C.top === v && C.right === S && C.bottom === p))
      return r.next({ left: a - 1, right: S + 1, top: v - 1, bottom: p + 1 }), m;
  });
  return $(() => {
    e && c();
  }, [t, o, i.unitAdded$, c, e, ...n != null ? n : []]), Pi(c), $(() => {
    const l = s.scrollEvent$.pipe($o(100)).subscribe(c);
    return () => {
      l.unsubscribe();
    };
  }, []), [r, c];
}
const Ze = (t) => {
  const e = X(t);
  return e.current = t, e;
}, Ui = (t, e, n) => {
  const o = T(Vt), r = T(Ut), s = T(Ce), [i, c] = W(), [l, m] = W(-1), [g, h] = W(!0), u = Ze(g), d = X(e);
  d.current = e;
  const a = () => {
    c(void 0), m(-1), h(!1);
  };
  return $(() => {
    const v = s.sequenceNodesBuilder(e.slice(1));
    o.setSequenceNodes(v != null ? v : []);
  }, [e]), $(() => {
    if (n && t) {
      const v = n.selectionChange$.pipe(rn(50)).subscribe((p) => {
        if (p.textRanges.length === 1) {
          const [C] = p.textRanges;
          if (C.collapsed && u.current) {
            const { startOffset: _ } = C, I = o.getCurrentSequenceNodeIndex(_ - 2), f = o.getCurrentSequenceNodeByIndex(I), b = o.getCurrentSequenceNodeByIndex(I + 1);
            if (f)
              if (typeof f != "string" && f.nodeType === 3 && !r.hasDefinedNameDescription(f.token.trim()) && b === qe.OPEN_BRACKET) {
                const y = r.getFunctionInfo(f.token);
                c(y), m(-1);
                return;
              } else {
                const y = s.getFunctionAndParameter(`${d.current}A`, _ - 1);
                if (y) {
                  const { functionName: A, paramIndex: D } = y, x = r.getFunctionInfo(A);
                  c(x), m(D);
                  return;
                }
              }
          }
        }
        c(void 0), m(-1);
      }), S = n.selectionChange$.pipe(
        or((p) => p.textRanges.length === 1),
        nr((p) => p.textRanges[0].startOffset),
        Po()
      ).subscribe(() => {
        h(!0);
      });
      return () => {
        v.unsubscribe(), S.unsubscribe();
      };
    }
  }, [n, t]), $(() => {
    t || a();
  }, [t]), {
    functionInfo: i,
    paramIndex: l,
    reset: a
  };
}, Wi = ({ onClick: t }) => /* @__PURE__ */ M(
  "div",
  {
    className: "univer-z-[15] univer-box-border univer-h-[18px] univer-cursor-pointer univer-overflow-visible univer-whitespace-nowrap univer-rounded-l univer-border univer-border-r-0 univer-border-gray-600 univer-bg-primary-600 univer-p-0.5 univer-text-xs univer-font-bold univer-leading-[13px] univer-text-white",
    onClick: t,
    children: "?"
  }
), Kt = ({ className: t, title: e, value: n }) => /* @__PURE__ */ V("div", { className: "univer-my-2", children: [
  /* @__PURE__ */ M(
    "div",
    {
      className: re("univer-mb-2 univer-text-sm univer-font-medium univer-text-gray-900 dark:!univer-text-white", t),
      children: e
    }
  ),
  /* @__PURE__ */ M(
    "div",
    {
      className: "univer-whitespace-pre-wrap univer-break-words univer-text-xs univer-text-gray-500",
      children: n
    }
  )
] }), Vi = (t) => {
  const { prefix: e, value: n, active: o, onClick: r } = t;
  return /* @__PURE__ */ V("div", { children: [
    /* @__PURE__ */ V("span", { children: [
      e,
      "("
    ] }),
    n && n.map((s, i) => /* @__PURE__ */ V("span", { children: [
      /* @__PURE__ */ M(
        "span",
        {
          className: o === i ? "univer-text-primary-500" : "",
          onClick: () => r(i),
          children: Sr(s)
        }
      ),
      i === n.length - 1 ? "" : ","
    ] }, s.name)),
    ")"
  ] });
}, Vn = () => {
};
function Hi(t) {
  const { onParamsSwitch: e = Vn, onClose: n = Vn, isFocus: o, editor: r, formulaText: s } = t, { functionInfo: i, paramIndex: c, reset: l } = Ui(o, s, r), m = T(nn), g = !pe(m.helpFunctionVisible$), [h, u] = W(!0), d = T(ot), a = d.t("formula.prompt.required"), v = d.t("formula.prompt.optional"), S = r.getEditorId(), [p] = Nr(S, !!i, [i, c]);
  function C(f) {
    e && e(f);
  }
  const _ = te((f) => {
    m.helpFunctionVisible$.next(!f);
  }), I = () => {
    _(!0), n();
  };
  return i ? g ? /* @__PURE__ */ M(Zt, { portal: !0, anchorRect$: p, direction: "left-center", children: /* @__PURE__ */ M(Wi, { onClick: () => _(!1) }) }, "hidden") : /* @__PURE__ */ M(Zt, { portal: !0, onClickOutside: () => l(), anchorRect$: p, direction: "vertical", children: /* @__PURE__ */ V(
    "div",
    {
      className: re("univer-m-0 univer-box-border univer-w-[250px] univer-select-none univer-list-none univer-rounded-lg univer-bg-white univer-leading-5 univer-shadow-md univer-outline-none dark:!univer-bg-gray-900", ln),
      children: [
        /* @__PURE__ */ V(
          "div",
          {
            className: re("univer-wrap-anywhere univer-box-border univer-flex univer-items-center univer-justify-between univer-px-4 univer-py-3 univer-text-xs univer-font-medium univer-text-gray-900 dark:!univer-text-white", Cs),
            children: [
              /* @__PURE__ */ M(
                Vi,
                {
                  prefix: i.functionName,
                  value: i.functionParameter,
                  active: c,
                  onClick: C
                }
              ),
              /* @__PURE__ */ V("div", { className: "univer-flex", children: [
                /* @__PURE__ */ M(
                  "div",
                  {
                    className: "univer-ml-2 univer-flex univer-h-6 univer-w-6 univer-cursor-pointer univer-items-center univer-justify-center univer-rounded univer-bg-transparent univer-p-0 univer-text-xs univer-text-gray-500 univer-outline-none univer-transition-colors hover:univer-bg-gray-200 dark:hover:!univer-bg-gray-600",
                    style: { transform: h ? "rotateZ(-90deg)" : "rotateZ(90deg)" },
                    onClick: () => u(!h),
                    children: /* @__PURE__ */ M(br, {})
                  }
                ),
                /* @__PURE__ */ M(
                  "div",
                  {
                    className: "univer-ml-2 univer-flex univer-h-6 univer-w-6 univer-cursor-pointer univer-items-center univer-justify-center univer-rounded univer-bg-transparent univer-p-0 univer-text-xs univer-text-gray-600 univer-outline-none univer-transition-colors hover:univer-bg-gray-300 dark:!univer-text-gray-200 dark:hover:!univer-bg-gray-600",
                    onClick: I,
                    children: /* @__PURE__ */ M(Rr, {})
                  }
                )
              ] })
            ]
          }
        ),
        /* @__PURE__ */ M(
          "div",
          {
            className: re("univer-box-border univer-max-h-[350px] univer-overflow-y-auto univer-px-4 univer-pb-3 univer-pt-0", nt),
            style: {
              height: h ? "unset" : 0,
              padding: h ? "revert-layer" : 0
            },
            children: /* @__PURE__ */ V("div", { className: "univer-mt-3", children: [
              /* @__PURE__ */ M(
                Kt,
                {
                  title: d.t("formula.prompt.helpExample"),
                  value: `${i.functionName}(${i.functionParameter.map((f) => f.example).join(",")})`
                }
              ),
              /* @__PURE__ */ M(
                Kt,
                {
                  title: d.t("formula.prompt.helpAbstract"),
                  value: i.description
                }
              ),
              i && i.functionParameter && i.functionParameter.map((f, b) => /* @__PURE__ */ M(
                Kt,
                {
                  className: c === b ? "univer-text-primary-500" : "",
                  title: f.name,
                  value: `${f.require ? a : v} ${f.detail}`
                },
                b
              ))
            ] })
          }
        )
      ]
    }
  ) }, "show") : null;
}
const Bi = (t) => {
  const e = T(je);
  return te((o) => {
    var r, s;
    if (t) {
      e.focus(t.getEditorId());
      const i = [...t.getSelectionRanges()];
      if (en.isDefine(o))
        t.setSelectionRanges([{ startOffset: o, endOffset: o }]);
      else if (!i.length && !t.docSelectionRenderService.isOnPointerEvent) {
        const c = (s = (r = t.getDocumentData().body) == null ? void 0 : r.dataStream) != null ? s : `\r
`, l = Math.max(c.length - 2, 0);
        t.setSelectionRanges([{ startOffset: l, endOffset: l }]);
      } else
        t.setSelectionRanges(i);
    }
  });
};
function qi(t) {
  var r, s;
  const n = t.get(ne).getCurrentUniverDocInstance();
  return n != null && n.getBody() ? { dataStream: (s = (r = n.getBody()) == null ? void 0 : r.dataStream) != null ? s : "", offset: 0 } : void 0;
}
var ye = /* @__PURE__ */ ((t) => (t[t.NOT_SELECT = 0] = "NOT_SELECT", t[t.NEED_ADD = 1] = "NEED_ADD", t[t.CAN_EDIT = 2] = "CAN_EDIT", t[t.EDIT_OTHER_SHEET_REFERENCE = 3] = "EDIT_OTHER_SHEET_REFERENCE", t[t.EDIT_OTHER_WORKBOOK_REFERENCE = 4] = "EDIT_OTHER_WORKBOOK_REFERENCE", t))(ye || {});
function ji(t) {
  var y;
  const { editorId: e, isFocus: n, disableOnClick: o, unitId: r, subUnitId: s } = t, i = T(Ne), c = T(ne), l = i.getRenderById(r), m = i.getRenderById(e), g = m == null ? void 0 : m.with(sr), h = T(dr), u = T(ft), [d, a] = W(
    0
    /* NOT_SELECT */
  ), v = T(Ce), S = X(!0), p = l == null ? void 0 : l.with(De), C = Ze(d), _ = c.getUnit(r, K.UNIVER_SHEET), I = _ == null ? void 0 : _.getSheetBySheetId(s), f = te((A) => {
    p && p.setSkipLastEnabled(
      A === 1 || A === 3 || A === 4
      /* EDIT_OTHER_WORKBOOK_REFERENCE */
    ), C.current = A, a(A);
  }), b = te(() => {
    var ce, ae, he;
    const A = c.getCurrentUnitOfType(K.UNIVER_SHEET);
    if (!A) return;
    const D = A.getActiveSheet(), x = g == null ? void 0 : g.getActiveTextRange(), N = x != null && x.collapsed ? x.startOffset : -1, k = qi(u);
    if (!k) return;
    const E = (ce = k == null ? void 0 : k.dataStream) == null ? void 0 : ce.slice(0, -2), O = ((ae = v.sequenceNodesBuilder(E)) != null ? ae : []).map((Y) => typeof Y == "object" ? Y.nodeType === j.REFERENCE ? {
      ...Y,
      range: xo(Y.token)
    } : {
      ...Y,
      range: void 0
    } : Y), R = E[N - 1], w = E[N], L = O.find((Y) => typeof Y == "object" && Y.nodeType === j.REFERENCE && N === Y.endIndex + 2), U = R && No(R) && (!w || Oo(w) && w !== qe.OPEN_BRACKET), H = !!L;
    if ((E == null ? void 0 : E.substring(0, 1)) === "=" && (U || H))
      if (H) {
        if (S.current)
          return;
        const { sheetName: Y, unitId: J } = L.range, ze = (he = c.getCurrentUnitOfType(K.UNIVER_SHEET)) == null ? void 0 : he.getUnitId();
        J && J !== ze ? f(
          4
          /* EDIT_OTHER_WORKBOOK_REFERENCE */
        ) : !Y && D.getSheetId() === (I == null ? void 0 : I.getSheetId()) || Y === D.getName() ? f(
          2
          /* CAN_EDIT */
        ) : f(
          3
          /* EDIT_OTHER_SHEET_REFERENCE */
        );
      } else
        S.current = !1, f(
          1
          /* NEED_ADD */
        );
    else
      f(
        0
        /* NOT_SELECT */
      );
  });
  return $(() => {
    const A = h.textSelection$.pipe(or((D) => D.unitId === e)).subscribe(() => {
      b();
    });
    return () => A.unsubscribe();
  }, [b, h.textSelection$, e]), $(() => {
    n || (f(
      0
      /* NOT_SELECT */
    ), S.current = !0);
  }, [n, f]), $(() => {
    var D;
    if (!o) return;
    const A = (D = m == null ? void 0 : m.mainComponent) == null ? void 0 : D.onPointerDown$.subscribeEvent(() => {
      f(
        0
        /* NOT_SELECT */
      ), S.current = !0;
    });
    return () => A == null ? void 0 : A.unsubscribe();
  }, [o, (y = m == null ? void 0 : m.mainComponent) == null ? void 0 : y.onPointerDown$, f]), $(() => {
    if (!n) return;
    const A = _ == null ? void 0 : _.activeSheet$.subscribe(() => {
      b();
    }), D = c.getCurrentTypeOfUnit$(K.UNIVER_SHEET).subscribe(() => {
      b();
    });
    return () => {
      A == null || A.unsubscribe(), D == null || D.unsubscribe();
    };
  }, [b, n, _ == null ? void 0 : _.activeSheet$, c.getCurrentTypeOfUnit$]), { isSelecting: d, isSelectingRef: C };
}
const Ki = () => {
  const t = T(Ce);
  return xt((n) => t.sequenceNodesBuilder(n) || [], [t]);
};
function Yi(t, e, n) {
  const o = new tn(e).setAlpha(0.05).toRgbString();
  return {
    id: n,
    strokeWidth: 1,
    stroke: e,
    fill: o,
    widgets: { tl: !0, tc: !0, tr: !0, ml: !0, mr: !0, bl: !0, bc: !0, br: !0 },
    widgetSize: 6,
    widgetStrokeWidth: 1,
    widgetStroke: t.getColorFromTheme("white")
  };
}
function Or(t) {
  var _, I, f;
  const {
    unitId: e,
    subUnitId: n,
    currentWorkbook: o,
    refSelections: r,
    editor: s,
    refSelectionsService: i,
    refSelectionsRenderService: c,
    sheetSkeletonManagerService: l,
    themeService: m,
    univerInstanceService: g
  } = t, h = o.getUnitId(), u = g.getUnit(e, K.UNIVER_SHEET), d = u == null ? void 0 : u.getActiveSheet(), a = [];
  if (!u || !d) {
    i.setSelections(a);
    return;
  }
  const v = d.getSheetId(), S = (b) => {
    var y;
    return (y = u == null ? void 0 : u.getSheetBySheetName(b)) == null ? void 0 : y.getSheetId();
  };
  if (!((_ = l == null ? void 0 : l.getWorksheetSkeleton(v)) == null ? void 0 : _.skeleton)) return;
  const C = [];
  for (let b = 0, y = r.length; b < y; b++) {
    const A = r[b], { themeColor: D, token: x, refIndex: N, endIndex: k } = A, E = st(x), { unitId: O, sheetName: R, range: w } = E, L = S(R);
    if (!L && R || h !== e && O !== h || O && O !== h || L && L !== v || !L && v !== n)
      continue;
    const U = rs(w, d.getRowCount(), d.getColumnCount());
    U.unitId = e, U.sheetId = v, a.push({
      range: U,
      primary: null,
      style: Yi(m, D, N.toString())
    }), C.push(k);
  }
  if (s) {
    const b = (f = (I = s.getSelectionRanges()) == null ? void 0 : I[0]) == null ? void 0 : f.startOffset, y = C.findIndex((A) => A + 2 === b);
    y !== -1 ? c == null || c.setActiveSelectionIndex(y) : c == null || c.resetActiveSelectionIndex();
  }
  return a;
}
function Zi(t, e) {
  const n = T(ne), o = T(gt), r = T(Pt), s = T(Ne), i = pe(ie(() => n.getCurrentTypeOfUnit$(K.UNIVER_SHEET), [n])), c = i ? s.getRenderById(i.getUnitId()) : null, l = c == null ? void 0 : c.with(De), m = c == null ? void 0 : c.with($t), g = te((h, u) => {
    const d = n.getCurrentUnitOfType(K.UNIVER_SHEET);
    if (!d || l != null && l.selectionMoving) return;
    const a = Or({
      unitId: t,
      subUnitId: e,
      currentWorkbook: d,
      refSelections: h,
      editor: u,
      refSelectionsService: r,
      refSelectionsRenderService: l,
      sheetSkeletonManagerService: m,
      themeService: o,
      univerInstanceService: n
    });
    if (!a) return;
    ((l == null ? void 0 : l.getSelectionControls()) || []).length === a.length ? l == null || l.resetSelectionsByModelData(a) : r.setSelections(a);
  });
  return $(() => () => {
    l == null || l.resetActiveSelectionIndex();
  }, [l]), g;
}
function Mr(t = "") {
  const e = T(Ut), n = zi(), o = T(me), r = ie(() => t.length, [t]);
  return te((i, c, l = !0, m) => {
    const g = i.getDocumentData(), h = i.getEditorId();
    if (!g)
      return [];
    const u = g.body;
    if (!u)
      return [];
    const d = u.dataStream.slice(0, u.dataStream.length - 2), a = { dataStream: "", ...g.body };
    if (!d.startsWith(t)) return [];
    if (c == null || c.length === 0)
      return a.textRuns = [], o.syncExecuteCommand(Mn.id, {
        unitId: h,
        body: Rn(a, 0, a.dataStream.length - 2)
      }), [];
    {
      const { textRuns: v, refSelections: S } = Gi(e, n, c);
      r && v.forEach((_) => {
        _.ed = _.ed + r, _.st = _.st + r;
      }), a.textRuns = [{ st: 0, ed: 1, ts: { fs: 11 } }, ...v];
      const p = c.reduce((_, I) => typeof I == "string" ? `${_}${I}` : `${_}${I.token}`, "");
      a.dataStream = `${t}${p}\r
`;
      let C;
      if (l) {
        C = i.getSelectionRanges();
        const _ = a.dataStream.length - 2 + r;
        C.forEach((I) => {
          I.startOffset = Math.max(0, Math.min(I.startOffset, _)), I.endOffset = Math.max(0, Math.min(I.endOffset, _));
        });
      }
      return o.syncExecuteCommand(Mn.id, {
        unitId: h,
        body: Rn(a, 0, a.dataStream.length - 2),
        textRanges: m != null ? m : C
      }), S;
    }
  });
}
function zi() {
  const t = T(gt), e = t.getCurrentTheme();
  return ie(() => ({ formulaRefColors: ["#B87333", "#8A9A5B", "#3D3831"], numberColor: "#3D3831", stringColor: "#7C3F2A", plainTextColor: "#3D3831", functionColor: "#3D3831" }), [e, t]);
}
function Gi(t, e, n) {
  const { formulaRefColors: o, numberColor: r, stringColor: s, plainTextColor: i, functionColor: c } = e, l = [], m = [], g = /* @__PURE__ */ new Map();
  let h = 0;
  for (let u = 0, d = n.length; u < d; u++) {
    const a = n[u];
    if (typeof a == "string") {
      const I = l[l.length - 1], f = I ? I.ed : 0, b = f + a.length;
      l.push({
        st: f,
        ed: b,
        ts: {
          cl: {
            rgb: i
          },
          fs: 11
        }
      });
      continue;
    }
    if (t.hasDefinedNameDescription(a.token.trim())) {
      l.push({
        st: a.startIndex,
        ed: a.endIndex + 1,
        ts: {
          cl: {
            rgb: i
          },
          fs: 11,
          bl: 1
        }
      });
      continue;
    }
    const { startIndex: v, endIndex: S, nodeType: p, token: C } = a;
    let _ = "";
    if (p === j.REFERENCE) {
      if (g.has(C))
        _ = g.get(C);
      else {
        const I = h % o.length;
        _ = o[I], g.set(C, _), h++;
      }
      m.push({
        refIndex: u,
        themeColor: _,
        token: C,
        startIndex: a.startIndex,
        endIndex: a.endIndex,
        index: m.length
      });
    } else p === j.NUMBER ? _ = r : p === j.STRING || p === j.ARRAY ? _ = s : p === j.FUNCTION && (_ = c);
    _ && _.length > 0 ? l.push({
      st: v,
      ed: S + 1,
      ts: {
        cl: {
          rgb: _
        },
        fs: 11,
        ...p === j.REFERENCE || p === j.FUNCTION ? { bl: 1 } : {}
      }
    }) : l.push({
      st: v,
      ed: S + 1,
      ts: {
        cl: {
          rgb: i
        },
        fs: 11
      }
    });
  }
  return { textRuns: l, refSelections: m };
}
const Xi = (t, e, n, o) => {
  const r = T(me), s = T(Wt), i = X(e);
  i.current = e;
  const c = X(o);
  c.current = o, $(() => {
    if (!n || !t)
      return;
    const m = `sheet.formula-embedding-editor.${n.getEditorId()}`, g = new Be(), h = (a, v) => {
      if (c.current) {
        c.current(a, v);
        return;
      }
      let S = de.LEFT;
      a === F.ARROW_DOWN ? S = de.DOWN : a === F.ARROW_UP ? S = de.UP : a === F.ARROW_RIGHT && (S = de.RIGHT), v === P.SHIFT ? r.executeCommand(Wo.id, {
        direction: S
      }) : r.executeCommand(Vo.id, {
        direction: S
      });
    }, u = (a, v) => {
      let S = de.DOWN;
      a === F.ARROW_DOWN ? S = de.DOWN : a === F.ARROW_UP ? S = de.UP : a === F.ARROW_LEFT ? S = de.LEFT : a === F.ARROW_RIGHT && (S = de.RIGHT), i.current ? v === P.CTRL_COMMAND ? r.executeCommand(Tn.id, {
        direction: S,
        jumpOver: xn.moveGap,
        extra: "formula-editor",
        fromCurrentSelection: i.current === ye.NEED_ADD || i.current === ye.EDIT_OTHER_SHEET_REFERENCE
      }) : v === P.SHIFT ? r.executeCommand(Nn.id, {
        direction: S,
        extra: "formula-editor"
      }) : v === (P.CTRL_COMMAND | P.SHIFT) ? r.executeCommand(Nn.id, {
        direction: S,
        jumpOver: xn.moveGap,
        extra: "formula-editor"
      }) : r.executeCommand(Tn.id, {
        direction: S,
        extra: "formula-editor",
        fromCurrentSelection: i.current === ye.NEED_ADD || i.current === ye.EDIT_OTHER_SHEET_REFERENCE
      }) : h(a, v);
    };
    return g.add(r.registerCommand({
      id: m,
      type: Te.OPERATION,
      handler(a, v) {
        const { keyCode: S, metaKey: p } = v;
        u(S, p);
      }
    })), [
      { keyCode: F.ARROW_DOWN },
      { keyCode: F.ARROW_LEFT },
      { keyCode: F.ARROW_RIGHT },
      { keyCode: F.ARROW_UP },
      { keyCode: F.ARROW_DOWN, metaKey: P.SHIFT },
      { keyCode: F.ARROW_LEFT, metaKey: P.SHIFT },
      { keyCode: F.ARROW_RIGHT, metaKey: P.SHIFT },
      { keyCode: F.ARROW_UP, metaKey: P.SHIFT },
      { keyCode: F.ARROW_DOWN, metaKey: P.CTRL_COMMAND },
      { keyCode: F.ARROW_LEFT, metaKey: P.CTRL_COMMAND },
      { keyCode: F.ARROW_RIGHT, metaKey: P.CTRL_COMMAND },
      { keyCode: F.ARROW_UP, metaKey: P.CTRL_COMMAND },
      { keyCode: F.ARROW_DOWN, metaKey: P.CTRL_COMMAND | P.SHIFT },
      { keyCode: F.ARROW_LEFT, metaKey: P.CTRL_COMMAND | P.SHIFT },
      { keyCode: F.ARROW_RIGHT, metaKey: P.CTRL_COMMAND | P.SHIFT },
      { keyCode: F.ARROW_UP, metaKey: P.CTRL_COMMAND | P.SHIFT }
    ].map(({ keyCode: a, metaKey: v }) => ({
      id: m,
      binding: v ? a | v : a,
      preconditions: () => !0,
      priority: 900,
      staticParameters: {
        eventType: xe.Keyboard,
        keyCode: a,
        metaKey: v
      }
    })).forEach((a) => {
      g.add(s.registerShortcut(a));
    }), () => {
      g.dispose();
    };
  }, [r, n, t, s]);
}, Qi = (t, e, n, o, r = !0) => {
  var d;
  const s = T(Ne), i = T(Qt), c = T(hs), l = T(Pt), m = T(ne), g = pe(ie(() => m.getCurrentTypeOfUnit$(K.UNIVER_SHEET), [m])), h = s.getRenderById((d = g == null ? void 0 : g.getUnitId()) != null ? d : ""), u = h == null ? void 0 : h.with(De);
  Nt(() => {
    if (t)
      return i.setContextValue(En, !0), r && c.disable(), () => {
        const a = m.getCurrentUnitOfType(K.UNIVER_DOC);
        (a == null ? void 0 : a.getUnitId()) === o && i.setContextValue(En, !1), r && c.enable(), l.clear();
      };
  }, [i, t, l, r, o]), Nt(() => {
    if (t && e) {
      const a = u == null ? void 0 : u.enableSelectionChanging();
      return i.setContextValue(An, !0), () => {
        i.setContextValue(An, !1), a == null || a.dispose();
      };
    }
  }, [i, t, u, e]), $(() => {
    t && (u == null || u.setSkipLastEnabled(!1));
  }, [t, u]);
}, Ji = (t, e, n) => {
  const o = T(ne), r = T(on);
  return xt(() => {
    if (t) {
      const i = [...r.getWorkbookSelections(e).getSelectionsOfWorksheet(n)], c = o.getCurrentUnitForType(K.UNIVER_SHEET), l = c == null ? void 0 : c.getActiveSheet();
      (c == null ? void 0 : c.getUnitId()) !== e && o.setCurrentUnitForType(e), l && l.getSheetId() === n && r.setSelections(i);
    }
  }, [t, r, n, e, o]);
}, ec = (t) => t.reduce((e, n) => typeof n == "string" ? e + n.length : e + n.token.length, 0), Xt = (t) => t.map((e) => typeof e == "string" ? e : e.token).join(""), bt = (t, e = !1, n = "", o = !1) => !e && !o ? t.map((r) => ve(r.range)) : t.map((r) => o ? Mo(r) : r.sheetName !== "" && r.sheetName !== n ? dt(r.sheetName, r.range) : ve(r.range)), tc = (t) => {
  var g, h, u;
  const { editor: e, lexerTreeBuilder: n } = t, o = e == null ? void 0 : e.getSelectionRanges();
  if ((o == null ? void 0 : o.length) !== 1)
    return;
  const s = o[0].startOffset - 1, i = ((h = (g = e == null ? void 0 : e.getDocumentData().body) == null ? void 0 : g.dataStream) != null ? h : `\r
`).slice(0, -2), c = (u = n.sequenceNodesBuilder(i.slice(1))) != null ? u : [], l = dn(c, s, !1), m = xr(c, l);
  return {
    nodeIndex: l,
    updatingRefIndex: m,
    sequenceNodes: c,
    offset: s
  };
}, nc = (() => {
}), rc = (t, e, n, o, r, s, i, c, l, m = nc) => {
  var N;
  const g = T(Ne), h = T(ne), u = T(me), d = T(dr), a = T(gt), v = T(Ce), S = h.getUnit(o), p = te((k, E) => {
    var O, R, w;
    return (w = (R = (O = h.getUnit(k)) == null ? void 0 : O.getSheetBySheetId(E)) == null ? void 0 : R.getName()) != null ? w : "";
  }), C = ie(() => p(o, r), [p, r, o]), _ = pe(S == null ? void 0 : S.activeSheet$), I = Ze({ activeSheet: _, sheetName: C }), f = pe(ie(() => h.getCurrentTypeOfUnit$(K.UNIVER_SHEET), [h])), b = g.getRenderById((N = f == null ? void 0 : f.getUnitId()) != null ? N : ""), y = b == null ? void 0 : b.with(De), A = b == null ? void 0 : b.with($t), D = T(Pt), x = te((k, E) => {
    var H, ce, ae, he, Y, J, ze, Ge, St, _e;
    const O = tc({ editor: l, lexerTreeBuilder: v });
    if (!O) return;
    const { nodeIndex: R, updatingRefIndex: w, sequenceNodes: L, offset: U } = O;
    if (n.current === ye.NEED_ADD)
      if (U !== 0) {
        if (R === -1 && L.length)
          return;
        const Z = k[k.length - 1], Q = L.splice(R + 1), Re = (H = Z.sheetId) != null ? H : r, fe = {
          range: Z,
          unitId: (ce = Z.unitId) != null ? ce : f.getUnitId(),
          sheetName: p((ae = Z.unitId) != null ? ae : f.getUnitId(), Re)
        }, ee = Re !== r, le = (f == null ? void 0 : f.getUnitId()) !== o, Ee = bt([fe], i && (ee || le), C, le);
        L.push({ token: Ee[0], nodeType: j.REFERENCE });
        const Fe = [...L, ...Q], Oe = Xt(Fe);
        m(Oe, ec(L), E);
      } else {
        const Z = k[k.length - 1], Q = (he = Z.sheetId) != null ? he : r, Re = {
          range: Z,
          unitId: (Y = Z.unitId) != null ? Y : f.getUnitId(),
          sheetName: p((J = Z.unitId) != null ? J : f.getUnitId(), Q)
        }, fe = Q !== r, ee = (f == null ? void 0 : f.getUnitId()) !== o, le = bt([Re], i && (fe || ee), C, ee);
        L.unshift({ token: le[0], nodeType: j.REFERENCE });
        const Ee = Xt(L);
        m(Ee, le[0].length, E);
      }
    else if (n.current === ye.EDIT_OTHER_SHEET_REFERENCE || n.current === ye.EDIT_OTHER_WORKBOOK_REFERENCE) {
      const Z = k.pop();
      if (!Z) return;
      const Q = L[R];
      if (typeof Q == "object" && Q.nodeType === j.REFERENCE) {
        const Re = Q.token;
        (f == null ? void 0 : f.getUnitId()) !== o ? Q.token = wo((ze = f == null ? void 0 : f.getUnitId()) != null ? ze : "", C, Z) : Q.token = C === (_ == null ? void 0 : _.getName()) ? ve(Z) : dt(_.getName(), Z);
        const ee = U + (Q.token.length - Re.length);
        m(ko(L), ee, E);
      }
    } else {
      const Z = [...k];
      if (w !== -1) {
        const z = Z.pop();
        z && Z.splice(w, 0, z);
      }
      let Q = 0;
      const Re = L.map((z) => {
        var ge, Xe, $e, Pe;
        if (typeof z == "string")
          return z;
        if (z.nodeType === j.REFERENCE) {
          const Ue = st(z.token);
          if (Ue.sheetName || (Ue.sheetName = C), (Ue.unitId || o) !== (f == null ? void 0 : f.getUnitId()) || i && ((ge = I.current.activeSheet) == null ? void 0 : ge.getName()) !== Ue.sheetName)
            return z.token;
          const oe = Z[Q];
          if (Q++, !oe)
            return "";
          const Ie = (Xe = oe.sheetId) != null ? Xe : r, vt = {
            range: oe,
            unitId: ($e = oe.unitId) != null ? $e : f.getUnitId(),
            sheetName: p((Pe = oe.unitId) != null ? Pe : f.getUnitId(), Ie)
          }, Ct = (f == null ? void 0 : f.getUnitId()) !== o;
          return bt([vt], i && (Ie !== r || Ct), C, Ct)[0];
        }
        return z.token;
      });
      let fe = "", ee;
      Re.forEach((z, ge) => {
        fe += z, ge === R && (ee = fe.length);
      });
      const le = [];
      for (let z = Q; z <= k.length - 1; z++) {
        const ge = k[z], Xe = (Ge = ge.sheetId) != null ? Ge : r, $e = {
          range: ge,
          unitId: (St = ge.unitId) != null ? St : f.getUnitId(),
          sheetName: p((_e = ge.unitId) != null ? _e : f.getUnitId(), Xe)
        }, Pe = (f == null ? void 0 : f.getUnitId()) !== o, oe = bt([$e], i && (Xe !== r || Pe), C, Pe);
        le.push(oe[0]);
      }
      const Ee = L[L.length - 1], Fe = Ee && (typeof Ee == "string" ? !1 : Ee.nodeType === j.REFERENCE), Oe = `${fe}${le.length && Fe ? "," : ""}${le.join(",")}`;
      m(Oe, !le.length && ee ? ee : Oe.length, E);
    }
  });
  $(() => {
    if (y && t) {
      let k = !0;
      const E = (R, w) => {
        if (k) {
          k = !1;
          return;
        }
        x(R.map((L) => L.rangeWithCoord), w);
      }, O = new Be();
      return O.add(y.selectionMoving$.subscribe((R) => {
        E(R, !1);
      })), O.add(y.selectionMoveEnd$.subscribe((R) => {
        E(R, !0);
      })), () => {
        O.dispose();
      };
    }
  }, [t, x, y]), $(() => {
    if (e && y && l) {
      const k = new Be(), E = () => {
        k.dispose(), y.getSelectionControls().forEach((w, L) => {
          k.add(
            w.selectionScaling$.subscribe((U) => {
              const H = y.getSelectionDataWithStyle().map((ae) => ae.rangeWithCoord), ce = H[L];
              U.sheetId = ce.sheetId, U.unitId = ce.unitId, H[L] = U, x(H, !1);
            })
          ), k.add(
            w.selectionMoving$.subscribe((U) => {
              const H = y.getSelectionDataWithStyle().map((ae) => ae.rangeWithCoord), ce = H[L];
              U.sheetId = ce.sheetId, U.unitId = ce.unitId, H[L] = U, x(H, !0);
            })
          );
        });
      }, O = Uo(
        l.input$,
        D.selectionSet$,
        y.selectionMoveEnd$
      ).pipe(
        Es(50)
      ).subscribe(() => {
        E();
      });
      return () => {
        O.unsubscribe(), k.dispose();
      };
    }
  }, [l, e, x, y, D.selectionSet$]), y == null || y.getSelectionDataWithStyle(), $(() => {
    if (c) {
      const k = u.onCommandExecuted((E) => {
        var R;
        if (E.id !== sn.id)
          return;
        const O = E.params;
        if (O.extra === "formula-editor" && O.selections.length) {
          const w = O.selections[O.selections.length - 1];
          if (w) {
            const L = n.current === ye.NEED_ADD, U = ((R = y == null ? void 0 : y.getSelectionDataWithStyle()) != null ? R : []).map((H) => H.rangeWithCoord);
            L ? U.push(w.range) : U[U.length - 1] = w.range, x(U, !0);
          }
        }
      });
      return () => {
        k.dispose();
      };
    }
  }, [u, l, n, v, c, x, y]), $(() => {
    if (!l)
      return;
    const k = d.textSelection$.subscribe((E) => {
      E.unitId === l.getEditorId() && Or({
        unitId: o,
        subUnitId: r,
        refSelections: s.current,
        editor: l,
        refSelectionsService: D,
        refSelectionsRenderService: y,
        sheetSkeletonManagerService: A,
        themeService: a,
        univerInstanceService: h,
        currentWorkbook: f
      });
    });
    return () => k.unsubscribe();
  }, [d.textSelection$, l, s, y, D, A, r, a, o, h]);
}, oc = (t, e, n, o, r, s) => {
  const i = T(me), c = T(je), m = T(Ne).getRenderById(e), g = T(ne), h = m == null ? void 0 : m.with(De);
  $(() => {
    if (t && h)
      if (n) {
        const u = () => {
          const v = h.getSelectionControls().length;
          for (let S = 1; S <= v; S++)
            h.clearLastSelection();
          return setTimeout(() => {
            s();
          }, 30);
        }, d = i.onCommandExecuted((v) => {
          v.id === Dn.id && u();
        }), a = g.getCurrentTypeOfUnit$(K.UNIVER_SHEET).subscribe((v) => {
          u();
        });
        return () => {
          d.dispose(), a.unsubscribe();
        };
      } else {
        const u = i.beforeCommandExecuted((d) => {
          if (d.id === Dn.id) {
            o(!1), r(), s();
            const a = c.getEditor(Jt);
            a == null || a.focus();
          }
        });
        return () => {
          u.dispose();
        };
      }
  }, [t, h]);
}, sc = (t, e, n) => {
  const o = T(Ce), r = X(!0);
  $(() => {
    if (t) {
      const s = setTimeout(() => {
        r.current = !1;
      }, 500);
      return () => {
        clearTimeout(s);
      };
    }
  }, [t]), $(() => {
    if (!r.current && e) {
      const s = o.checkIfAddBracket(n);
      e(s === 0 && n.startsWith(tr.EQUALS), `${n}`);
    }
  }, [n, e]);
}, ic = (t, e = [], n) => {
  const o = T(Ut), [r, s] = W([]), [i, c] = W(""), l = X(-1), m = Ze({ nodes: e }), g = () => {
    s([]), c(""), l.current = -1;
  };
  return $(() => {
    if (n && t) {
      const u = n.input$.pipe(rn(300)).subscribe(() => {
        const d = n.getSelectionRanges();
        if (d.length === 1) {
          const a = m.current.nodes, v = d[0];
          if (v.collapsed) {
            const S = dn(a, v.startOffset - 1, !1);
            l.current = S;
            const p = a[S];
            if (p && typeof p != "string" && p.nodeType === j.FUNCTION) {
              l.current = S;
              const C = p.token, _ = o.getSearchListByNameFirstLetter(C);
              s(_), c(C);
              return;
            }
          }
        }
        l.current = -1, c(""), s((a) => a != null && a.length ? [] : a);
      });
      return () => {
        u.unsubscribe();
      };
    }
  }, [n, t]), $(() => {
    t || g();
  }, [t]), {
    searchList: r,
    searchText: i,
    handlerFormulaReplace: (u, d) => {
      const a = [...m.current.nodes];
      if (l.current !== -1) {
        const v = a.splice(l.current + 1), S = a.pop() || "";
        let p = (typeof S == "string" ? S.length : S.token.length) - u.length;
        return a.push(u), v[0] !== qe.OPEN_BRACKET && d !== er.DefinedName && (a.push(qe.OPEN_BRACKET), p--), { text: Xt([...a, ...v]), offset: p };
      }
    },
    reset: g
  };
}, cc = () => {
}, ac = Le(lc);
function lc(t, e) {
  const { isFocus: n, sequenceNodes: o, onSelect: r, editor: s, onClose: i = cc } = t, c = s.getEditorId(), l = T(Wt), m = T(me), { searchList: g, searchText: h, handlerFormulaReplace: u, reset: d } = ic(n, o, s), a = ie(() => !!g.length, [g]), v = X(void 0), [S, p] = W(0), C = X(!1), [_] = Nr(c, a, [h, g]), I = Ze({ searchList: g, active: S }), f = (x, N) => {
    const k = u(x, N);
    k && (d(), r(k));
  };
  function b(x) {
    C.current && p(x);
  }
  function y() {
    C.current && p(-1);
  }
  $(() => {
    if (!g.length)
      return;
    const x = `sheet.formula-embedding-editor.search_function.${c}`, N = new Be(), k = (E) => {
      const { searchList: O, active: R } = I.current;
      switch (E) {
        case F.ARROW_UP: {
          p((w) => {
            const L = Math.max(0, w - 1);
            return A(L), L;
          });
          break;
        }
        case F.ARROW_DOWN: {
          p((w) => {
            const L = Math.min(O.length - 1, w + 1);
            return A(L), L;
          });
          break;
        }
        case F.TAB:
        case F.ENTER: {
          const w = O[R];
          f(w.name, w.functionType);
          break;
        }
        case F.ESC: {
          d(), i();
          break;
        }
      }
    };
    return N.add(m.registerCommand({
      id: x,
      type: Te.OPERATION,
      handler(E, O) {
        const { keyCode: R } = O;
        k(R);
      }
    })), [F.ARROW_UP, F.ARROW_DOWN, F.ENTER, F.ESC, F.TAB].map((E) => ({
      id: x,
      binding: E,
      preconditions: () => !0,
      priority: 1e3,
      staticParameters: {
        eventType: xe.Keyboard,
        keyCode: E
      }
    })).forEach((E) => {
      N.add(l.registerShortcut(E));
    }), () => {
      N.dispose();
    };
  }, [g]);
  function A(x) {
    const N = v.current;
    if (!N) return;
    const k = N.children[x];
    if (!k) return;
    const O = N.getBoundingClientRect().top, R = N.offsetHeight, w = k.getBoundingClientRect(), L = w.top, U = w.height;
    if (L >= 0 && L > O && L - O + U <= R)
      return;
    const H = k.offsetTop - (R - U) / 2;
    N.scrollTo({
      top: H,
      behavior: "smooth"
    });
  }
  const D = ie(() => {
    let x = "";
    return () => {
      clearTimeout(x), C.current = !0, x = setTimeout(() => {
        C.current = !1;
      }, 300);
    };
  }, []);
  return g.length > 0 && a && /* @__PURE__ */ M(Zt, { portal: !0, anchorRect$: _, direction: "vertical", children: /* @__PURE__ */ M(
    "ul",
    {
      ref: (x) => {
        v.current = x, e && (e.current = x);
      },
      "data-u-comp": "sheets-formula-editor",
      className: re("univer-m-0 univer-box-border univer-max-h-[400px] univer-w-[250px] univer-list-none univer-overflow-y-auto univer-rounded-lg univer-bg-white univer-p-2 univer-leading-5 univer-shadow-md univer-outline-none dark:!univer-bg-gray-900", ln, nt),
      children: g.map((x, N) => /* @__PURE__ */ V(
        "li",
        {
          className: re("univer-box-border univer-cursor-pointer univer-rounded univer-px-2 univer-py-1 univer-text-gray-900 univer-transition-colors dark:!univer-text-white", {
            "univer-bg-gray-200 dark:!univer-bg-gray-600": S === N
          }),
          onMouseEnter: () => b(N),
          onMouseLeave: y,
          onMouseMove: D,
          onClick: () => {
            f(x.name, x.functionType), s && s.focus();
          },
          children: [
            /* @__PURE__ */ V("span", { className: "univer-block univer-overflow-x-hidden univer-text-ellipsis univer-text-xs", children: [
              /* @__PURE__ */ M("span", { className: "univer-text-red-500", children: x.name.substring(0, h.length) }),
              /* @__PURE__ */ M("span", { children: x.name.slice(h.length) })
            ] }),
            /* @__PURE__ */ M(
              "span",
              {
                className: "univer-block univer-text-xs univer-text-gray-400",
                children: x.desc
              }
            )
          ]
        },
        x.name
      ))
    }
  ) });
}
const uc = (t) => t.startsWith(tr.EQUALS) ? t.slice(1) : "", Hn = () => {
}, dc = Le((t, e) => {
  var gn, mn, pn, Sn;
  const {
    errorText: n,
    initValue: o,
    unitId: r,
    subUnitId: s,
    isFocus: i = !0,
    isSupportAcrossSheet: c = !1,
    onFocus: l = Hn,
    onBlur: m = Hn,
    onChange: g,
    onVerify: h,
    className: u,
    editorId: d,
    moveCursor: a = !0,
    onFormulaSelectingChange: v,
    keyboardEventConfig: S,
    onMoveInEditor: p,
    resetSelectionOnBlur: C = !0,
    autoScrollbar: _ = !0,
    isSingle: I = !0,
    disableSelectionOnClick: f = !1,
    autofocus: b = !0,
    disableContextMenu: y,
    style: A
  } = t, D = T(je), x = X(null), N = te(g);
  ps(e, () => ({
    isClickOutSide: (G) => x.current ? !x.current.contains(G.target) : !1
  }));
  const k = te(v), E = X(null), O = X(void 0), R = O.current, [w, L] = W(i), U = X(null), H = ie(() => d != null ? d : zr(`${_o}-${Ft(4)}`), []), ce = ie(() => n !== void 0, [n]), ae = T(ne), he = ae.getUnit(H);
  pe(he == null ? void 0 : he.change$);
  const Y = Ki(), J = Gr.transform.getPlainText((mn = (gn = he == null ? void 0 : he.getBody()) == null ? void 0 : gn.dataStream) != null ? mn : ""), ze = Ze(J), Ge = ie(() => uc(J), [J]), St = ie(() => Y(Ge), [Ge, Y]), { isSelecting: _e, isSelectingRef: Z } = ji({ unitId: r, subUnitId: s, editorId: H, isFocus: w, disableOnClick: f }), Q = X(""), fe = T(Ne).getRenderById(H), ee = fe == null ? void 0 : fe.with(sr), le = ee == null ? void 0 : ee.isFocusing, Ee = ie(() => ae.getCurrentTypeOfUnit$(K.UNIVER_DOC), [ae]), Fe = pe(Ee), Oe = (Fe == null ? void 0 : Fe.getUnitId()) === H, z = X([]), ge = _e, $e = (Sn = (pn = T(Zn).getConfig(mr)) == null ? void 0 : pn.functionScreenTips) != null ? Sn : !0;
  fs(() => {
    N(J);
  }, [J, N]);
  const Pe = Mr("="), Ue = Zi(r, s), oe = te((G, se = !0, Me, be) => {
    if (!O.current) return;
    Q.current = G;
    const _t = G[0] === "=" ? G.slice(1) : "", we = Y(_t), Dr = we.reduce((We, Et) => typeof Et == "object" ? `${We}${Et.token}` : `${We}${Et}`, ""), Rt = Pe(
      O.current,
      Dr === _t ? we : [],
      se,
      be
    );
    if (z.current = Rt, Me) {
      const We = be != null ? be : R == null ? void 0 : R.getSelectionRanges();
      if ((We == null ? void 0 : We.length) !== 1)
        return;
      const Lr = We[0].startOffset - 1, Fr = dn(we, Lr, !1), vn = xr(we, Fr);
      if (vn >= 0) {
        const Cn = Rt.splice(vn, 1)[0];
        Cn && Rt.push(Cn);
      }
      Ue(w ? Rt : [], O.current);
    }
  });
  $(() => {
    w && oe(J, !1, !0);
  }, [w]), $(() => {
    if (w) {
      if (Q.current === J) return;
      oe(J, !1, !0);
    }
  }, [J]), sc(w, h, J);
  const Ie = Bi(R), vt = Ji(w, r, s);
  $(() => {
    var G;
    k(_e, (G = ee == null ? void 0 : ee.isFocusing) != null ? G : !0);
  }, [k, _e]), Ho(w, S, R), Nt(() => {
    let G;
    if (U.current) {
      G = D.register({
        autofocus: b,
        editorUnitId: H,
        initialSnapshot: {
          id: H,
          body: {
            dataStream: `${o}\r
`,
            textRuns: [],
            customBlocks: [],
            customDecorations: [],
            customRanges: []
          },
          documentStyle: {}
        }
      }, U.current);
      const se = D.getEditor(H);
      O.current = se, oe(o, !1, !0);
    }
    return () => {
      G == null || G.dispose();
    };
  }, []), Nt(() => {
    i ? (L(i), Ie()) : (C && (R == null || R.blur(), vt()), L(i));
  }, [i, R, Ie, vt, C]);
  const { checkScrollBar: Ct } = Bo(R, I, _);
  Qi(w, !!(_e && Oe), r, H, y), Xi(!!(w && le && a), ge, R, p);
  const hn = te((G, se, Me) => {
    if (!le)
      return;
    const be = se !== -1 ? [{ startOffset: se + 1, endOffset: se + 1, collapsed: !0 }] : void 0;
    oe(`=${G}`, !0, Me, be), Me && (Ie(), se !== -1 && setTimeout(() => {
      const _t = { startOffset: se + 1, endOffset: se + 1 }, we = R == null ? void 0 : R.render.with(qo);
      we == null || we.scrollToRange({ ..._t, collapsed: !0 });
    }, 50), Ct());
  });
  rc(
    w && !!(_e && Oe),
    w,
    Z,
    r,
    s,
    z,
    c,
    !!ge,
    R,
    hn
  ), oc(w && !!(_e && Oe), r, c, L, m, () => {
    oe(ze.current, !1, !0);
  });
  const fn = (G) => {
    if (G) {
      const se = R == null ? void 0 : R.getSelectionRanges();
      if (se && se.length === 1) {
        const Me = se[0];
        if (Me.collapsed) {
          const be = G.offset;
          setTimeout(() => {
            R == null || R.setSelectionRanges([{ startOffset: Me.startOffset - be, endOffset: Me.endOffset - be }]);
          }, 30);
        }
      }
      Ie(), oe(`=${G.text}`);
    }
  }, Ar = () => {
    L(!0), l(), Ie();
  };
  return /* @__PURE__ */ V("div", { className: u, children: [
    /* @__PURE__ */ M(
      "div",
      {
        ref: x,
        className: re("univer-relative univer-box-border univer-flex univer-h-full univer-w-full univer-items-center univer-justify-around univer-gap-2 univer-rounded-none univer-p-0 univer-ring-1", {
          "univer-ring-primary-500": w,
          "univer-ring-red-500": ce
        }),
        children: /* @__PURE__ */ M(
          "div",
          {
            ref: U,
            className: "univer-relative univer-h-full univer-w-full",
            onMouseUp: Ar
          }
        )
      }
    ),
    n !== void 0 && /* @__PURE__ */ M("div", { className: "univer-my-1 univer-text-xs univer-text-red-500", children: n }),
    $e && R && Ge !== "" && /* @__PURE__ */ M(
      Hi,
      {
        editor: R,
        isFocus: w,
        formulaText: J,
        onClose: () => Ie()
      }
    ),
    $e && !!R && /* @__PURE__ */ M(
      ac,
      {
        isFocus: w,
        sequenceNodes: St,
        onSelect: fn,
        ref: E,
        editor: R
      }
    )
  ] });
});
function hc(t, e, n, o) {
  const r = T(Ce), s = Mr(""), i = pe(t == null ? void 0 : t.getDocumentDataModel().change$), [c, l] = W([]), m = T(Ro), g = X(""), h = T(ne);
  return $(() => {
    if (!t) return;
    const u = t.getDocumentDataModel().getPlainText();
    if (g.current === u)
      return;
    g.current = u;
    const d = r.sequenceNodesBuilder(u);
    l(d != null ? d : []);
  }, [i, t, r]), $(() => {
    var a, v;
    if (!t) return;
    if (!e) {
      const S = t.getDocumentData();
      t.setDocumentData({
        ...S,
        body: {
          ...S.body,
          dataStream: (v = (a = S.body) == null ? void 0 : a.dataStream) != null ? v : "",
          textRuns: []
        }
      });
      return;
    }
    const u = s(t, c, !1), d = new Be();
    return u.forEach((S) => {
      const p = st(S.token), C = h.getCurrentUnitForType(K.UNIVER_SHEET), _ = C == null ? void 0 : C.getActiveSheet();
      if (!p.sheetName && o !== (_ == null ? void 0 : _.getSheetId()) || p.sheetName && (_ == null ? void 0 : _.getName()) !== p.sheetName)
        return;
      const I = new tn(S.themeColor).toRgb(), f = m.addShape({
        range: p.range,
        style: {
          stroke: S.themeColor,
          fill: `rgba(${I.r}, ${I.g}, ${I.b}, 0.1)`,
          strokeDash: 12
        },
        primary: null
      });
      f && d.add(() => m.removeShape(f));
    }), () => {
      d.dispose();
    };
  }, [t, e, s, m, c]), { sequenceNodes: c };
}
function fc(t) {
  const e = T(on), { supportAcrossSheet: n = !1, keepSheetReference: o = !1, unitId: r, subUnitId: s, onChange: i } = t, l = T(ne).getUnit(r, K.UNIVER_SHEET), m = te(i), g = te((h, u) => {
    const d = l == null ? void 0 : l.getActiveSheet();
    if (!d || !n && d.getSheetId() !== s || !(h != null && h.length)) return;
    const a = o ? d.getName() : d.getSheetId() === s ? "" : d.getName(), v = h.map((S) => ({
      range: S.range,
      unitId: r,
      sheetName: a
    }));
    m(v, u);
  });
  $(() => {
    const h = new Be();
    return h.add(e.selectionMoveStart$.subscribe((u) => {
      g(u, !0);
    })), h.add(e.selectionMoving$.subscribe((u) => {
      g(u, !1);
    })), h.add(e.selectionMoveEnd$.subscribe((u) => {
      g(u, !1);
    })), () => {
      h.dispose();
    };
  }, [g, e.selectionMoveEnd$, e.selectionMoveStart$, e.selectionMoving$]);
}
const Bn = (t) => !t.some((n) => {
  if (typeof n == "string") {
    if (n !== qe.COMMA)
      return !0;
  } else if (n.nodeType !== j.REFERENCE)
    return !0;
  return !1;
}), gc = (t) => {
  if (t.endColumn < t.startColumn) {
    const e = t.endColumn;
    t.endColumn = t.startColumn, t.startColumn = e;
  }
  if (t.endRow < t.startRow) {
    const e = t.endRow;
    t.endRow = t.startRow, t.startRow = e;
  }
  return t;
};
function mc(t) {
  const {
    visible: e,
    initialValue: n,
    unitId: o,
    subUnitId: r,
    maxRangeCount: s = 1 / 0,
    supportAcrossSheet: i,
    keepSheetReference: c,
    onConfirm: l,
    onClose: m,
    onShowBySelection: g
  } = t, h = T(ot), u = T(Ce), [d, a] = W([]), [v, S] = W(0), p = X(null);
  $(() => {
    if (e && n.length) {
      const f = n.map((b) => b.sheetName ? dt(b.sheetName, b.range) : ve(b.range));
      a(f), S(f.length - 1);
    } else
      a([""]), S(0);
  }, [e]);
  const C = (f, b) => {
    const y = [...d];
    y[f] = b, a(y);
  }, _ = () => {
    a([...d, ""]), S(d.length);
  }, I = (f) => {
    d.splice(f, 1), a([...d]);
  };
  return fc({
    unitId: o,
    subUnitId: r,
    supportAcrossSheet: i,
    keepSheetReference: c,
    onChange: (f, b) => {
      if (!e && g != null && g(f))
        return;
      const y = new Set(d), A = f.map((N) => N.sheetName ? dt(N.sheetName, N.range) : ve(N.range)), D = A.filter((N) => !y.has(N));
      if (!D.length) return;
      const x = [...d];
      if (A.length > 1) {
        b || x.splice(v, 1), x.push(...D);
        const N = x.slice(0, s);
        a(N), S(N.length - 1), requestAnimationFrame(() => {
          var k;
          (k = p.current) == null || k.scrollTo({ top: p.current.scrollHeight });
        });
      } else {
        x.splice(v, 1, ...D);
        const N = x.slice(0, s);
        a(N), S(v + D.length - 1);
      }
    }
  }), /* @__PURE__ */ M(
    Rs,
    {
      width: "328px",
      open: e,
      title: h.t("rangeSelector.title"),
      draggable: !0,
      mask: !1,
      maskClosable: !1,
      footer: /* @__PURE__ */ V("footer", { className: "univer-flex univer-gap-2", children: [
        /* @__PURE__ */ M(et, { onClick: m, children: h.t("rangeSelector.cancel") }),
        /* @__PURE__ */ M(
          et,
          {
            variant: "primary",
            onClick: () => {
              l(
                d.filter((f) => {
                  const b = u.sequenceNodesBuilder(f);
                  return b && b.length === 1 && typeof b[0] != "string" && b[0].nodeType === j.REFERENCE;
                }).map((f) => st(f)).map((f) => ({ ...f, range: gc(f.range) }))
              );
            },
            children: h.t("rangeSelector.confirm")
          }
        )
      ] }),
      onClose: m,
      children: /* @__PURE__ */ V(
        "div",
        {
          ref: p,
          className: re("-univer-mx-6 univer-max-h-60 univer-overflow-y-auto univer-px-6", nt),
          children: [
            d.map((f, b) => /* @__PURE__ */ V(
              "div",
              {
                className: "univer-mb-2 univer-flex univer-items-center univer-gap-4",
                children: [
                  /* @__PURE__ */ M(
                    ur,
                    {
                      className: re("univer-w-full", {
                        "univer-border-primary-600": v === b
                      }),
                      placeholder: h.t("rangeSelector.placeHolder"),
                      onFocus: () => S(b),
                      value: f,
                      onChange: (y) => C(b, y)
                    }
                  ),
                  d.length > 1 && /* @__PURE__ */ M(
                    Er,
                    {
                      className: "univer-cursor-pointer",
                      onClick: () => I(b)
                    }
                  )
                ]
              },
              b
            )),
            d.length < s && /* @__PURE__ */ M("div", { children: /* @__PURE__ */ V(et, { variant: "link", onClick: _, children: [
              /* @__PURE__ */ M(Ir, {}),
              /* @__PURE__ */ M("span", { children: h.t("rangeSelector.addAnotherRange") })
            ] }) })
          ]
        }
      )
    }
  );
}
function pc(t) {
  return t.split(qe.COMMA).filter((e) => !!e).map((e) => st(e));
}
function Sc(t) {
  return t.map((e) => e.sheetName ? dt(e.sheetName, e.range) : ve(e.range)).join(qe.COMMA);
}
function wr(t) {
  const [e, n] = W(null), {
    onVerify: o,
    selectorRef: r,
    unitId: s,
    subUnitId: i,
    maxRangeCount: c,
    supportAcrossSheet: l,
    keepSheetReference: m,
    autoFocus: g,
    onChange: h,
    onRangeSelectorDialogVisibleChange: u,
    onClickOutside: d,
    onFocusChange: a,
    forceShowDialogWhenSelectionChanged: v,
    hideEditor: S,
    resetRange: p
  } = t, [C, _] = W(g != null ? g : !1), [I, f] = W(!1), [b, y] = W([]), A = T(ot), D = T(je), { sequenceNodes: x } = hc(e, C, s, i), N = Ze(x), k = T(me), E = te(() => {
    e == null || e.setSelectionRanges([]), e == null || e.blur(), D.blur();
  }), O = te(() => {
    var R;
    E(), y(pc((R = e == null ? void 0 : e.getDocumentDataModel().getPlainText()) != null ? R : "")), f(!0);
  });
  return $(() => {
    r && (r.current = {
      get editor() {
        return e;
      },
      focus() {
        D.focus(e.getEditorId());
      },
      blur: E,
      verify: () => Bn(N.current),
      showDialog: (R) => {
        E(), y(R), f(!0);
      },
      hideDialog: () => {
        y([]), f(!1);
      },
      getValue: () => {
        var R;
        return (R = e == null ? void 0 : e.getDocumentDataModel().getPlainText()) != null ? R : "";
      }
    });
  }, [E, e, D, r, N]), $(() => {
    var R;
    o == null || o(Bn(x), (R = e == null ? void 0 : e.getDocumentDataModel().getPlainText()) != null ? R : "");
  }, [x]), $(() => {
    u == null || u(I);
  }, [I]), $(() => {
    if (I && p)
      return () => {
        const R = {
          unitId: s,
          subUnitId: i,
          selections: p
        };
        k.executeCommand(sn.id, R);
      };
  }, [I]), /* @__PURE__ */ V(ms, { children: [
    S ? null : /* @__PURE__ */ M(
      jo,
      {
        isSingle: !0,
        ...t,
        onFocusChange: (R, w) => {
          _(R), a == null || a(R, w);
        },
        editorRef: n,
        onClickOutside: () => {
          _(!1), E(), d == null || d();
        },
        icon: /* @__PURE__ */ M(_s, { title: A.t("rangeSelector.buttonTooltip"), placement: "bottom", children: /* @__PURE__ */ M(
          yr,
          {
            className: "univer-cursor-pointer dark:!univer-text-gray-300",
            onClick: O
          }
        ) })
      }
    ),
    /* @__PURE__ */ M(
      mc,
      {
        initialValue: b,
        unitId: s,
        subUnitId: i,
        visible: I,
        maxRangeCount: c,
        onConfirm: (R) => {
          const w = Sc(R), L = Xr.newEmptyData();
          L.body.dataStream = w, e == null || e.replaceText(w, !1), h == null || h(L, w), f(!1), y([]), requestAnimationFrame(() => {
            E();
          });
        },
        onClose: () => {
          f(!1), y([]);
        },
        supportAcrossSheet: l,
        keepSheetReference: m,
        onShowBySelection: (R) => C || v ? (y(R), f(!0), !1) : !0
      }
    )
  ] });
}
const vc = () => {
  var o, r;
  const t = T(Tr), e = pe(t.currentSelector$), n = X(null);
  return $(() => {
    var s, i;
    if (e)
      return (i = n.current) == null || i.showDialog((s = e.initialValue) != null ? s : []), () => {
        var c;
        (c = n.current) == null || c.hideDialog();
      };
  }, [e]), /* @__PURE__ */ M(
    wr,
    {
      unitId: (o = e == null ? void 0 : e.unitId) != null ? o : "",
      subUnitId: (r = e == null ? void 0 : e.subUnitId) != null ? r : "",
      hideEditor: !0,
      selectorRef: n,
      onChange: (s, i) => {
        var c;
        e == null || e.callback((c = i == null ? void 0 : i.split(",").map((l) => st(l))) != null ? c : []);
      }
    }
  );
};
var Cc = Object.defineProperty, _c = Object.getOwnPropertyDescriptor, Rc = (t, e, n) => e in t ? Cc(t, e, { enumerable: !0, configurable: !0, writable: !0, value: n }) : t[e] = n, Ec = (t, e, n, o) => {
  for (var r = o > 1 ? void 0 : o ? _c(e, n) : e, s = t.length - 1, i; s >= 0; s--)
    (i = t[s]) && (r = i(r) || r);
  return r;
}, yt = (t, e) => (n, o) => e(n, o, t), kr = (t, e, n) => Rc(t, typeof e != "symbol" ? e + "" : e, n);
let Lt = class extends Jr {
  constructor(t = Pn, e, n, o, r) {
    super(), this._config = t, this._injector = e, this._renderManagerService = n, this._configService = o, this._uiPartsService = r;
    const { menu: s, ...i } = eo(
      Pn,
      this._config
    );
    s && this._configService.setConfig("menu", s, { merge: !0 }), this._configService.setConfig(mr, i, { merge: !0 });
  }
  onStarting() {
    to(this._injector, [
      [Vt, { useClass: zt }],
      [Tr],
      [Dt],
      [Mt],
      [wt],
      [kt],
      [At],
      [Ot]
    ]), this._initUIPart();
  }
  onReady() {
    [
      [De]
    ].forEach((t) => {
      this.disposeWithMe(this._renderManagerService.registerRenderModule(K.UNIVER_SHEET, t));
    });
  }
  onRendered() {
    [
      [Gt]
    ].forEach((t) => {
      this.disposeWithMe(this._renderManagerService.registerRenderModule(K.UNIVER_SHEET, t));
    }), no(this._injector, [
      [Dt],
      // FormulaProgressBar relies on TriggerCalculationController, but it is necessary to ensure that the formula calculation is done after rendered.
      [wt],
      [At]
    ]);
  }
  onSteady() {
    this._injector.get(Mt), this._injector.get(Ot);
  }
  _initUIPart() {
    const t = this._injector.get(ar);
    this.disposeWithMe(t.register(Eo, wr)), this.disposeWithMe(t.register(Io, dc)), this.disposeWithMe(this._uiPartsService.registerComponent(gs.GLOBAL, () => lr(vc, this._injector)));
  }
};
kr(Lt, "pluginName", hr);
kr(Lt, "type", K.UNIVER_SHEET);
Lt = Ec([
  Qr(Ao, cs),
  yt(1, B(ft)),
  yt(2, Ne),
  yt(3, Zn),
  yt(4, cr)
], Lt);
export {
  Ts as FORMULA_PROMPT_ACTIVATED,
  dc as FormulaEditor,
  Ot as FormulaReorderController,
  Tr as GlobalRangeSelectorService,
  xs as HelpFunctionOperation,
  Ns as InsertFunctionOperation,
  Ls as MoreFunctionsOperation,
  wr as RangeSelector,
  De as RefSelectionsRenderService,
  gr as ReferenceAbsoluteOperation,
  Fs as SearchFunctionOperation,
  it as SelectEditorFormulaOperation,
  un as SheetOnlyPasteFormulaCommand,
  Lt as UniverSheetsFormulaUIPlugin
};
