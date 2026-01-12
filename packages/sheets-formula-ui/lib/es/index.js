var $r = Object.defineProperty;
var Pr = (t, e, n) => e in t ? $r(t, e, { enumerable: !0, configurable: !0, writable: !0, value: n }) : t[e] = n;
var q = (t, e, n) => Pr(t, typeof e != "symbol" ? e + "" : e, n);
import { CommandType as Te, ICommandService as me, createIdentifier as Ur, IContextService as Qt, IUniverInstanceService as ne, Rectangle as qn, DOCS_NORMAL_EDITOR_UNIT_ID_KEY as Jt, DOCS_FORMULA_BAR_EDITOR_UNIT_ID_KEY as jn, DEFAULT_EMPTY_DOCUMENT_VALUE as Kn, isRealNum as Wr, CellValueType as _n, getCellValueType as Vr, Inject as B, Disposable as rt, ObjectMatrix as Ae, Range as Hr, Tools as en, LocaleService as ot, isICellData as Br, isFormulaString as ke, isFormulaId as tt, generateRandomId as Lt, Direction as he, Injector as ft, UniverInstanceType as j, ThemeService as gt, ILogService as qr, toDisposable as Yn, ColorKit as tn, RxDisposable as jr, InterceptorEffectEnum as Kr, FOCUSING_DOC as Yr, FOCUSING_UNIVER_EDITOR as Zr, DisposableCollection as Be, RANGE_TYPE as Se, getBodySlice as Rn, EDITOR_ACTIVATED as En, createInternalEditorID as zr, BuildTextUtils as Gr, IConfigService as Zn, RichTextBuilder as Xr, DependentOn as Qr, Plugin as Jr, merge as eo, registerDependencies as to, touchDependencies as no } from "@univerjs/core";
import { SheetPasteCommand as ro, PREDEFINED_HOOK_NAME as Je, IEditorBridgeService as nn, SetCellEditVisibleOperation as zn, HoverManagerService as oo, CellAlertManagerService as so, CellAlertType as io, IAutoFillService as co, APPLY_TYPE as lo, DATA_TYPE as In, ISheetClipboardService as ao, COPY_TYPE as Gn, SheetSkeletonManagerService as $t, attachSelectionWithCoord as Yt, SelectionControl as Xn, SELECTION_SHAPE_DEPTH as uo, useActiveWorkbook as ho, getCurrentRangeDisable$ as fo, PASTE_SPECIAL_MENU_ID as go, whenFormulaEditorActivated as mt, whenSheetEditorFocused as mo, SheetsUIPart as po, BaseSelectionRenderService as So, getCoordByOffset as bn, checkInHeaderRanges as yn, getAllSelection as vo, genNormalSelectionStyle as Qn, getSheetObject as Co, MoveSelectionCommand as Tn, JumpOver as xn, ExpandSelectionCommand as On, EMBEDDING_FORMULA_EDITOR as _o, IMarkSelectionService as Ro, RANGE_SELECTOR_COMPONENT_KEY as Eo, EMBEDDING_FORMULA_EDITOR_COMPONENT_KEY as Io } from "@univerjs/sheets-ui";
import { sequenceNodeType as Q, serializeRange as ve, FormulaDataModel as pt, LexerTreeBuilder as Ce, ErrorType as ue, extractFormulaError as Jn, SetFormulaCalculationResultMutation as bo, SetArrayFormulaDataMutation as yo, SetFormulaCalculationStopMutation as To, FunctionType as er, matchToken as qe, deserializeRangeWithSheetWithCache as xo, matchRefDrawToken as Oo, isFormulaLexerToken as No, deserializeRangeWithSheet as st, serializeRangeToRefString as Mo, serializeRangeWithSheet as ht, serializeRangeWithSpreadsheet as wo, generateStringWithSequence as ko, operatorToken as tr, UniverFormulaEnginePlugin as Ao } from "@univerjs/engine-formula";
import { Subject as ct, debounceTime as rn, combineLatestWith as Do, map as nr, switchMap as Fo, of as Nn, Observable as Lo, BehaviorSubject as rr, throttleTime as $o, filter as or, distinctUntilChanged as Po, merge as Uo } from "rxjs";
import { IEditorService as je, DocSelectionRenderService as sr, ReplaceTextRunsCommand as Mn, MoveSelectionOperation as Wo, MoveCursorOperation as Vo, useKeyboardEvent as Ho, useResize as Bo, DocBackScrollRenderController as qo, RichTextEditor as jo } from "@univerjs/docs-ui";
import { DeviceInputEventType as xe, IRenderManagerService as Oe, ScrollTimerType as Bt, SHEET_VIEWPORT_KEY as wn, Vector2 as kn } from "@univerjs/engine-render";
import { SheetsSelectionsService as on, getSheetCommandTarget as ir, getCellAtRowCol as Ko, SetSelectionsOperation as sn, SheetInterceptorService as cn, ReorderRangeCommand as Yo, SetRangeValuesMutation as Tt, SetRangeValuesUndoMutationFactory as Zo, BEFORE_CELL_EDIT as zo, SetWorksheetRowAutoHeightMutation as Go, INTERCEPTOR_POINT as Xo, WorksheetSetCellValuePermission as Qo, WorksheetEditPermission as Jo, RangeProtectionPermissionEditPoint as es, WorkbookEditablePermission as ts, IRefSelectionsService as Pt, SelectionMoveType as qt, convertSelectionDataToRange as ns, setEndForRange as rs, REF_SELECTIONS_ENABLED as An, SetWorksheetActiveOperation as Dn } from "@univerjs/sheets";
import { InsertFunctionCommand as os, TriggerCalculationController as ss, IDescriptionService as Ut, QuickSumCommand as is, UniverSheetsFormulaPlugin as cs } from "@univerjs/sheets-formula";
import { ISidebarService as ln, IZenZoneService as ls, useDependency as T, useObservable as pe, ProgressBar as as, MenuItemType as us, IClipboardInterfaceService as Fn, KeyCode as L, MetaKeys as P, IMenuManagerService as hs, IShortcutService as Wt, IUIPartsService as cr, ComponentManager as lr, connectInjector as ar, useEvent as te, RectPopup as Zt, IContextMenuService as ds, useUpdateEffect as fs, BuiltInUIPart as gs } from "@univerjs/ui";
import { jsx as M, jsxs as V, Fragment as ms } from "react/jsx-runtime";
import { useCallback as xt, useState as W, useRef as G, createElement as Ke, forwardRef as Fe, useEffect as $, useMemo as ie, useLayoutEffect as Ot, useImperativeHandle as ps } from "react";
import { clsx as re, scrollbarClassName as nt, borderLeftClassName as Ss, Select as vs, Input as ur, borderClassName as an, Button as et, borderTopClassName as Cs, Tooltip as _s, Dialog as Rs } from "@univerjs/design";
import { DocSelectionManagerService as hr } from "@univerjs/docs";
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
    if (typeof n == "string" || n.nodeType !== Q.REFERENCE)
      return;
    const o = e.length - n.token.length, r = { ...n };
    r.token = e, r.endIndex += o, this._sequenceNodes[t] = r;
    for (let s = t + 1, i = this._sequenceNodes.length; s < i; s++) {
      const c = this._sequenceNodes[s];
      if (typeof c == "string")
        continue;
      const a = { ...c };
      a.startIndex += o, a.endIndex += o, this._sequenceNodes[s] = a;
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
      nodeType: Q.REFERENCE
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
      const a = { ...c };
      a.startIndex += r, a.endIndex += r, this._sequenceNodes[s] = a;
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
}, Os = {
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
    const { worksheet: i, unitId: c, subUnitId: a } = s, g = i.getCellMatrix(), { value: m } = e, d = t.get(me);
    t.get(nn);
    const h = [], u = [];
    let l = null, v = 0, p = 0, S = "";
    if (r.length === 1 && (ws(r[0].range) || ks(r[0].range) && $n(g, r[0].range))) {
      const { range: I, primary: f } = r[0], b = (C = f == null ? void 0 : f.actualRow) != null ? C : I.startRow, y = (_ = f == null ? void 0 : f.actualColumn) != null ? _ : I.startColumn;
      l = I, v = b, p = y;
      const A = Ln(g, b, y);
      A && (S = ve(A));
    } else
      r.some((I) => {
        var y, A;
        const { range: f, primary: b } = I;
        if ($n(g, f)) {
          const D = (y = b == null ? void 0 : b.actualRow) != null ? y : f.startRow, x = (A = b == null ? void 0 : b.actualColumn) != null ? A : f.startColumn, O = Ln(g, D, x);
          if (!O)
            return l = f, v = D, p = x, !0;
          const k = ve(O), E = `=${m}(${k})`;
          h.push({
            range: f,
            primary: {
              row: D,
              column: x
            },
            formula: E
          });
        } else {
          const { startRow: D, startColumn: x, endRow: O, endColumn: k } = f;
          if (D === O) {
            const E = As(g, D, k, i.getColumnCount() - 1), N = E === k ? k - 1 : k, R = ve({
              startRow: D,
              endRow: O,
              startColumn: x,
              endColumn: N
            }), w = `=${m}(${R})`;
            u.push({
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
              const w = Ds(g, R, O, i.getRowCount() - 1);
              E = Math.max(E, w);
            }
            const N = E === O ? O - 1 : O;
            for (let R = x; R <= k; R++) {
              const w = ve({
                startRow: D,
                endRow: N,
                startColumn: R,
                endColumn: R
              }), F = `=${m}(${w})`;
              u.push({
                range: f,
                primary: {
                  row: E,
                  column: R
                },
                formula: F
              });
            }
          }
        }
        return !1;
      });
    if (l) {
      const I = Ko(v, p, i), f = {
        range: qn.clone(l),
        primary: {
          startRow: I.startRow,
          startColumn: I.startColumn,
          endRow: I.endRow,
          endColumn: I.endColumn,
          actualRow: v,
          actualColumn: p,
          isMerged: I.isMerged,
          isMergedMainCell: I.startRow === v && I.startColumn === p
        }
      }, b = {
        unitId: c,
        subUnitId: a,
        selections: [f]
      };
      await d.executeCommand(sn.id, b);
      const y = o.getEditor(Jt), A = o.getEditor(jn);
      d.syncExecuteCommand(zn.id, {
        visible: !0,
        unitId: c,
        eventType: xe.Dblclick
      });
      const D = `=${m}(${S}`;
      y == null || y.replaceText(D), A == null || A.replaceText(D, !1);
    }
    return h.length === 0 && u.length === 0 ? !1 : d.executeCommand(os.id, {
      list: h,
      listOfRangeHasNumber: u
    });
  }
};
function Ln(t, e, n) {
  const o = Ns(t, e, n);
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
function Ns(t, e, n) {
  let o = !1;
  if (e === 0) return e;
  for (let r = e - 1; r >= 0; r--) {
    const s = t.getValue(r, n);
    if (dt(s) && !o) {
      if (r === 0) return 0;
      o = !0;
    } else {
      if (o && !dt(s))
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
    if (dt(s) && !o) {
      if (r === 0) return 0;
      o = !0;
    } else {
      if (o && !dt(s))
        return r + 1;
      if (o && r === 0)
        return 0;
    }
  }
  return n;
}
function dt(t) {
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
      if (dt(t.getValue(n, o)))
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
const dr = "SHEET_FORMULA_UI_PLUGIN", fr = `${dr}_MORE_FUNCTIONS_COMPONENT`, Fs = {
  id: "formula-ui.operation.more-functions",
  type: Te.OPERATION,
  handler: async (t) => (t.get(ln).open({
    header: { title: "formula.insert.tooltip" },
    children: { label: fr }
  }), !0)
}, gr = {
  id: "formula-ui.operation.change-ref-to-absolute",
  type: Te.OPERATION,
  handler: async (t) => !0
}, Ls = {
  id: "formula-ui.operation.search-function",
  type: Te.OPERATION,
  handler: async (t, e) => (t.get(Vt).search(e), !0)
};
var $s = Object.getOwnPropertyDescriptor, Ps = (t, e, n, o) => {
  for (var r = o > 1 ? void 0 : o ? $s(e, n) : e, s = t.length - 1, i; s >= 0; s--)
    (i = t[s]) && (r = i(r) || r);
  return r;
}, It = (t, e) => (n, o) => e(n, o, t);
let Nt = class extends rt {
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
    const e = [], n = [], { unitId: o, subUnitId: r, range: s, order: i } = t, c = this._univerInstanceService.getUniverSheetInstance(o), a = c == null ? void 0 : c.getSheetBySheetId(r);
    if (!a)
      return {
        redos: e,
        undos: n
      };
    const g = a.getCellMatrix(), m = new Ae(), d = new Ae();
    let h = !1;
    return Hr.foreach(s, (u, l) => {
      let v = u;
      i.hasOwnProperty(u) && (v = i[u]);
      const p = g.getValue(v, l);
      if (p != null && p.f || p != null && p.si) {
        h = !0;
        const S = this._formulaDataModel.getFormulaStringByCell(v, l, r, o), C = this._lexerTreeBuilder.moveFormulaRefOffset(
          S,
          0,
          u - v
        ), _ = en.deepClone(p);
        _.f = C, _.si = null, m.setValue(u, l, _);
      } else
        m.setValue(u, l, p);
      d.setValue(u, l, g.getValue(u, l));
    }), h ? (e.push({
      id: Tt.id,
      params: {
        unitId: o,
        subUnitId: r,
        cellValue: m.getMatrix()
      }
    }), n.push({
      id: Tt.id,
      params: {
        unitId: o,
        subUnitId: r,
        cellValue: d.getMatrix()
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
Nt = Ps([
  It(0, B(cn)),
  It(1, B(ne)),
  It(2, B(pt)),
  It(3, B(Ce))
], Nt);
const mr = "sheets-formula-ui.base.config", Pn = {};
var Us = Object.getOwnPropertyDescriptor, Ws = (t, e, n, o) => {
  for (var r = o > 1 ? void 0 : o ? Us(e, n) : e, s = t.length - 1, i; s >= 0; s--)
    (i = t[s]) && (r = i(r) || r);
  return r;
}, lt = (t, e) => (n, o) => e(n, o, t);
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
        const a = c.getCell(t.location.row, t.location.col), g = (r = (o = (n = (e = this._formulaDataModel.getArrayFormulaCellData()) == null ? void 0 : e[t.location.unitId]) == null ? void 0 : n[t.location.subUnitId]) == null ? void 0 : o[t.location.row]) == null ? void 0 : r[t.location.col];
        if (Br(a)) {
          const m = Jn(a, !!g);
          if (!m) {
            this._hideAlert();
            return;
          }
          const d = this._cellAlertManagerService.currentAlert.get(jt), h = (s = d == null ? void 0 : d.alert) == null ? void 0 : s.location;
          if (h && h.row === t.location.row && h.col === t.location.col && h.subUnitId === t.location.subUnitId && h.unitId === t.location.unitId)
            return;
          this._cellAlertManagerService.showAlert({
            type: io.ERROR,
            title: this._localeService.t("formula.error.title"),
            message: this._localeService.t(`formula.error.${Vs[m]}`),
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
  lt(1, B(oo)),
  lt(2, B(so)),
  lt(3, B(ot)),
  lt(4, B(pt)),
  lt(5, ls)
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
        [lo.COPY]: (e, n, o, r, s) => {
          const { data: i, index: c } = e;
          return this._fillCopyFormula(i, n, o, c, r, s);
        }
      }
    };
    this._autoFillService.registerRule(t);
  }
  _fillCopyFormula(t, e, n, o, r, s) {
    var g, m;
    const i = js(r), c = [], a = /* @__PURE__ */ new Map();
    for (let d = 1; d <= e; d++) {
      const h = (d - 1) % t.length, u = o[h], l = en.deepClone(t[h]);
      if (l) {
        const v = ((g = t[h]) == null ? void 0 : g.f) || "", p = ((m = t[h]) == null ? void 0 : m.si) || "", S = ke(v);
        if (tt(p))
          l.si = p, l.f = null, l.v = null, l.p = null, l.t = null, c.push(l);
        else if (S) {
          let _ = a.get(h);
          if (_)
            l.si = _, l.f = null, l.v = null, l.p = null, l.t = null;
          else {
            _ = Lt(6), a.set(h, _);
            const { offsetX: I, offsetY: f } = qs(i, e, n, s, u), b = this._lexerTreeBuilder.moveFormulaRefOffset(
              v,
              I,
              f
            );
            l.si = _, l.f = b, l.v = null, l.p = null, l.t = null;
          }
          c.push(l);
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
  const { source: s, target: i } = o, { rows: c } = i, { rows: a } = s;
  let g = 0, m = 0;
  switch (n) {
    case he.UP:
      m = c[r] - a[r];
      break;
    case he.RIGHT:
      g = t;
      break;
    case he.DOWN:
      m = c[r] - a[r];
      break;
    case he.LEFT:
      g = -t * e;
      break;
  }
  return { offsetX: g, offsetY: m };
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
}, at = (t, e) => (n, o) => e(n, o, t);
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
    var h;
    if ([
      Je.SPECIAL_PASTE_FORMAT,
      Je.SPECIAL_PASTE_COL_WIDTH
    ].includes(o.pasteType))
      return {
        undos: [],
        redos: []
      };
    const i = this._currentUniverSheet.getCurrentUnitForType(j.UNIVER_SHEET), c = e.unitId || i.getUnitId(), a = e.subUnitId || ((h = i.getActiveSheet()) == null ? void 0 : h.getSheetId());
    if (!c || !a)
      return {
        undos: [],
        redos: []
      };
    const g = e.range, m = n, d = {
      copyType: o.copyType || Gn.COPY,
      copyRange: t == null ? void 0 : t.range,
      pasteType: o.pasteType
    };
    return this._injector.invoke((u) => zs(
      c,
      a,
      g,
      m,
      u,
      d,
      this._lexerTreeBuilder,
      this._formulaDataModel,
      r,
      t
    ));
  }
};
wt = Ys([
  at(0, ne),
  at(1, B(Ce)),
  at(2, ao),
  at(3, B(ft)),
  at(4, B(pt))
], wt);
function zs(t, e, n, o, r, s, i, c, a = !1, g) {
  const m = [], d = [], h = Gs(t, e, n, o, s, i, c, g);
  if (!h.hasValue())
    return {
      undos: [],
      redos: []
    };
  const u = {
    unitId: t,
    subUnitId: e,
    cellValue: h.getData()
  };
  m.push({
    id: Tt.id,
    params: u
  });
  const l = Zo(
    r,
    u
  );
  return d.push({
    id: Tt.id,
    params: l
  }), {
    undos: d,
    redos: m
  };
}
function Gs(t, e, n, o, r, s, i, c) {
  return c ? r.pasteType === Je.SPECIAL_PASTE_VALUE ? Qs(t, e, n, o, i, c) : r.pasteType === Je.SPECIAL_PASTE_FORMULA ? Js(t, e, n, o, s, i, c) : ei(t, e, n, o, r.copyType, s, i, c) : Xs(t, e, n, o, i);
}
function Xs(t, e, n, o, r) {
  const s = new Ae(), i = r.getSheetFormulaData(t, e);
  return o.forValue((c, a, g) => {
    var u;
    const m = n.rows[c], d = n.cols[a], h = {};
    ke(g.v) ? (h.v = null, h.f = `${g.v}`, h.si = null, h.p = null, s.setValue(m, d, h)) : (u = i == null ? void 0 : i[m]) != null && u[d] && (h.v = g.v, h.f = null, h.si = null, h.p = null, s.setValue(m, d, h));
  }), s;
}
function Qs(t, e, n, o, r, s) {
  var g, m;
  const i = new Ae(), c = (m = (g = r.getArrayFormulaCellData()) == null ? void 0 : g[s.unitId]) == null ? void 0 : m[s.subUnitId], a = r.getSheetFormulaData(t, e);
  return o.forValue((d, h, u) => {
    var _, I;
    const l = s.range.rows[d % s.range.rows.length], v = s.range.cols[h % s.range.cols.length], p = n.rows[d], S = n.cols[h], C = {};
    if (ke(u.f) || tt(u.si))
      C.v = u.v, C.f = null, C.si = null, C.p = null, i.setValue(p, S, C);
    else if ((_ = c == null ? void 0 : c[l]) != null && _[v]) {
      const f = c[l][v];
      C.v = f.v, C.f = null, C.si = null, C.p = null, i.setValue(p, S, C);
    } else if ((I = a == null ? void 0 : a[p]) != null && I[S]) {
      if (C.v = u.v, C.f = null, C.si = null, C.p = null, u.p) {
        const f = pr(u);
        f && (C.v = f);
      }
      i.setValue(p, S, C);
    }
  }), i;
}
function Js(t, e, n, o, r, s, i) {
  const c = new Ae(), a = /* @__PURE__ */ new Map();
  return o.forValue((g, m, d) => {
    const h = n.rows[g], u = n.cols[m], l = {};
    if (tt(d.si)) {
      if (i.unitId !== t || i.subUnitId !== e) {
        const v = s.getFormulaStringByCell(
          i.range.rows[g % i.range.rows.length],
          i.range.cols[m % i.range.cols.length],
          i.subUnitId,
          i.unitId
        ), p = n.cols[m] - i.range.cols[m % i.range.cols.length], S = n.rows[g] - i.range.rows[g % i.range.rows.length], C = r.moveFormulaRefOffset(v || "", p, S);
        l.si = null, l.f = C;
      } else
        l.si = d.si, l.f = null;
      l.v = null, l.p = null, c.setValue(h, u, l);
    } else if (ke(d.f)) {
      const v = `${g % i.range.rows.length}_${m % i.range.cols.length}`;
      let p = a.get(v);
      if (p)
        l.si = p, l.f = null;
      else {
        p = Lt(6), a.set(v, p);
        const S = n.cols[m] - i.range.cols[m % i.range.cols.length], C = n.rows[g] - i.range.rows[g % i.range.rows.length], _ = r.moveFormulaRefOffset(d.f || "", S, C);
        l.si = p, l.f = _;
      }
      l.v = null, l.p = null, c.setValue(h, u, l);
    } else {
      if (l.v = d.v, l.f = null, l.si = null, l.p = null, d.p) {
        const v = pr(d);
        v && (l.v = v);
      }
      c.setValue(h, u, l);
    }
  }), c;
}
function ei(t, e, n, o, r, s, i, c) {
  const a = new Ae(), g = /* @__PURE__ */ new Map(), m = i.getSheetFormulaData(t, e), d = [];
  return r === Gn.CUT ? o.forValue((h, u, l) => {
    const v = n.rows[h], p = n.cols[u], S = {};
    if (tt(l.si)) {
      if (ke(l.f))
        d.push(l.si), S.f = l.f, S.si = l.si;
      else if (d.includes(l.si))
        S.f = null, S.si = l.si;
      else {
        const C = i.getFormulaStringByCell(
          c.range.rows[h % c.range.rows.length],
          c.range.cols[u % c.range.cols.length],
          c.subUnitId,
          c.unitId
        );
        S.f = C, S.si = null;
      }
      S.v = null, S.p = null, a.setValue(v, p, S);
    } else ke(l.f) && (S.f = l.f, S.si = null, S.v = null, S.p = null, a.setValue(v, p, S));
  }) : o.forValue((h, u, l) => {
    var C;
    const v = n.rows[h], p = n.cols[u], S = {};
    if (tt(l.si)) {
      if (c.unitId !== t || c.subUnitId !== e) {
        const _ = i.getFormulaStringByCell(
          c.range.rows[h % c.range.rows.length],
          c.range.cols[u % c.range.cols.length],
          c.subUnitId,
          c.unitId
        ), I = n.cols[u] - c.range.cols[u % c.range.cols.length], f = n.rows[h] - c.range.rows[h % c.range.rows.length], b = s.moveFormulaRefOffset(_ || "", I, f);
        S.si = null, S.f = b;
      } else
        S.si = l.si, S.f = null;
      S.v = null, S.p = null, a.setValue(v, p, S);
    } else if (ke(l.f)) {
      const _ = `${h % c.range.rows.length}_${u % c.range.cols.length}`;
      let I = g.get(_);
      if (I)
        S.si = I, S.f = null;
      else {
        I = Lt(6), g.set(_, I);
        const f = n.cols[u] - c.range.cols[u % c.range.cols.length], b = n.rows[h] - c.range.rows[h % c.range.rows.length], y = s.moveFormulaRefOffset(l.f || "", f, b);
        S.si = I, S.f = y;
      }
      S.v = null, S.p = null, a.setValue(v, p, S);
    } else (C = m == null ? void 0 : m[v]) != null && C[p] && (S.v = l.v, S.f = null, S.si = null, S.p = l.p, a.setValue(v, p, S));
  }), d.length > 0 && new Ae(m).forValue((h, u, l) => {
    if (!(c.range.rows.includes(h) && c.range.cols.includes(u)) && !(n.rows.includes(h) && n.cols.includes(u)) && d.includes(l == null ? void 0 : l.si)) {
      const v = i.getFormulaStringByCell(
        h,
        u,
        c.subUnitId,
        c.unitId
      );
      a.setValue(h, u, {
        f: v,
        si: null,
        v: null,
        p: null
      });
    }
  }), a;
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
  constructor(e, n, o, r, s, i, c, a) {
    super();
    q(this, "_previousShape");
    q(this, "_skeleton");
    this._context = e, this._sheetInterceptorService = n, this._formulaDataModel = o, this._themeService = r, this._renderManagerService = s, this._sheetSkeletonManagerService = i, this._commandService = c, this._logService = a, this._initSkeletonChangeListener(), this._initInterceptorEditorStart(), this._commandExecutedListener();
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
            var l, v, p, S;
            const { row: r, col: s, unitId: i, subUnitId: c, worksheet: a } = n, g = this._formulaDataModel.getArrayFormulaRange(), m = this._formulaDataModel.getArrayFormulaCellData();
            if (this._removeArrayFormulaRangeShape(), e == null)
              return o(e);
            let d = null;
            const h = this._formulaDataModel.getFormulaStringByCell(r, s, c, i);
            if (h !== null && (d = { f: h }), e.v != null && e.v !== "" && ((p = (v = (l = m[i]) == null ? void 0 : l[c]) == null ? void 0 : v[r]) == null ? void 0 : p[s]) == null)
              return d ? { ...e, ...d } : o(e);
            const u = (S = g == null ? void 0 : g[i]) == null ? void 0 : S[c];
            return u != null && (d = this._displayArrayFormulaRangeShape(u, r, s, i, c, a, d)), d ? { ...e, ...d } : o(e);
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
    const a = this._formulaDataModel.getSheetFormulaData(r, s);
    return new Ae(e).forValue((g, m, d) => {
      var p;
      if (d == null)
        return !0;
      const { startRow: h, startColumn: u, endRow: l, endColumn: v } = d;
      if (g === n && m === o)
        return this._createArrayFormulaRangeShape(d, r), !1;
      if (n >= h && n <= l && o >= u && o <= v) {
        const S = i.getCell(h, u);
        if ((S == null ? void 0 : S.v) === ue.SPILL)
          return;
        const C = (p = a == null ? void 0 : a[g]) == null ? void 0 : p[m];
        return C == null || C.f == null ? !0 : (c == null && (c = {
          f: C.f,
          isInArrayFormulaRange: !0
        }), this._createArrayFormulaRangeShape(d, r), !1);
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
    }, c = Yt(i, r), { rowHeaderWidth: a, columnHeaderHeight: g } = r, m = new Xn(s, uo.FORMULA_EDITOR_SHOW, this._themeService, {
      highlightHeader: !1,
      rowHeaderWidth: a,
      columnHeaderHeight: g
    });
    m.updateRangeBySelectionWithCoord(c), m.setEvent(!1), this._previousShape = m;
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
    for (let a = 0; a < o.length; a++) {
      const { row: g } = o[a];
      if (r >= g) {
        const m = {
          startRow: r,
          endRow: s,
          startColumn: i,
          endColumn: c
        };
        this._refreshArrayFormulaRangeShape(e, m);
        break;
      }
    }
  }
};
kt = ni([
  Ve(1, B(cn)),
  Ve(2, B(pt)),
  Ve(3, B(gt)),
  Ve(4, Oe),
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
          var c, a, g, m;
          const s = (m = (g = (a = (c = this._formulaDataModel.getArrayFormulaCellData()) == null ? void 0 : c[o.unitId]) == null ? void 0 : a[o.subUnitId]) == null ? void 0 : g[o.row]) == null ? void 0 : m[o.col];
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
  return /* @__PURE__ */ M(as, { progress: n, onTerminate: o, onClearProgress: r });
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
function li(t) {
  const { functionInfo: e, onChange: n } = t;
  if (!e) return null;
  const [o, r] = W([]), [s, i] = W(e.functionParameter), [c, a] = W(-1);
  return /* @__PURE__ */ V("div", { children: [
    /* @__PURE__ */ M("div", { className: re("univer-h-[364px] univer-overflow-y-auto", nt), children: s.map((g, m) => /* @__PURE__ */ V("div", { children: [
      /* @__PURE__ */ M("div", { className: "univer-text-sm", children: g.name }),
      /* @__PURE__ */ M("div", { className: "univer-mb-2 univer-mt-1" })
    ] }, m)) }),
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
  const { icon: n, id: o, className: r, extend: s, ...i } = e, c = `univerjs-icon univerjs-icon-${o} ${r || ""}`.trim(), a = G(`_${hi()}`);
  return Cr(n, `${o}`, {
    defIds: n.defIds,
    idSuffix: a.current
  }, {
    ref: t,
    className: c,
    ...i
  }, s);
}
function Cr(t, e, n, o, r) {
  return Ke(t.tag, {
    key: e,
    ...ai(t, n, r),
    ...o
  }, (ui(t, n).children || []).map((s, i) => Cr(s, `${e}-${t.tag}-${i}`, n, void 0, r)));
}
function ai(t, e, n) {
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
function hi() {
  return Math.random().toString(36).substring(2, 8);
}
Ye.displayName = "UniverIcon";
const di = {
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
}, _r = Fe(function(e, n) {
  return Ke(Ye, Object.assign({}, e, {
    id: "check-mark-icon",
    ref: n,
    icon: di
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
}, Rr = Fe(function(e, n) {
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
}, Er = Fe(function(e, n) {
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
}, Ir = Fe(function(e, n) {
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
}, br = Fe(function(e, n) {
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
}, yr = Fe(function(e, n) {
  return Ke(Ye, Object.assign({}, e, {
    id: "select-range-icon",
    ref: n,
    icon: Si
  }));
});
yr.displayName = "SelectRangeIcon";
function vi(t) {
  const { onChange: e } = t, n = "-1", [o, r] = W(""), [s, i] = W([]), [c, a] = W(0), [g, m] = W(n), [d, h] = W(0), [u, l] = W(null), v = T(Ut), p = T(ot), S = T(ln), C = pe(S.sidebarOptions$), _ = ci(er, p);
  _.unshift({
    label: p.t("formula.moreFunctions.allFunctions"),
    value: n
  });
  const I = p.t("formula.prompt.required"), f = p.t("formula.prompt.optional");
  $(() => {
    A(n);
  }, []), $(() => {
    y(0);
  }, [s]), $(() => {
    C != null && C.visible && (r(""), i([]), a(0), m(n), h(0), l(null), A(n));
  }, [C]);
  const b = (E) => {
    if (o.trim() === "") return E;
    const N = new RegExp(`(${o.toLocaleUpperCase()})`);
    return E.split(N).filter(Boolean).map((w, F) => w.match(N) ? /* @__PURE__ */ M("span", { className: "univer-text-red-500", children: w }, F) : w);
  }, y = (E) => {
    if (s.length === 0) {
      l(null);
      return;
    }
    h(E);
    const N = v.getFunctionInfo(s[E].name);
    if (!N) {
      l(null);
      return;
    }
    l(N), e(N);
  };
  function A(E) {
    m(E);
    const N = v.getSearchListByType(+E);
    i(N);
  }
  function D(E) {
    r(E);
    const N = v.getSearchListByName(E);
    i(N);
  }
  function x(E) {
    if (E.stopPropagation(), E.key === "ArrowDown") {
      const N = c + 1;
      a(N === s.length ? 0 : N);
    } else if (E.key === "ArrowUp") {
      const N = c - 1;
      a(N === -1 ? s.length - 1 : N);
    } else E.key === "Enter" && y(c);
  }
  const O = (E) => {
    a(E);
  }, k = () => {
    a(-1);
  };
  return /* @__PURE__ */ V("div", { children: [
    /* @__PURE__ */ V("div", { className: "univer-flex univer-items-center univer-justify-between univer-gap-2", children: [
      /* @__PURE__ */ M(vs, { value: g, options: _, onChange: A }),
      /* @__PURE__ */ M(
        ur,
        {
          placeholder: p.t("formula.moreFunctions.searchFunctionPlaceholder"),
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
        className: re("univer-mb-0 univer-mt-2 univer-box-border univer-max-h-72 univer-w-full univer-select-none univer-list-none univer-overflow-y-auto univer-rounded univer-p-3 univer-outline-none", an, nt),
        onKeyDown: x,
        tabIndex: -1,
        children: s.map(({ name: E }, N) => /* @__PURE__ */ V(
          "li",
          {
            className: re("univer-relative univer-box-border univer-cursor-pointer univer-rounded univer-px-7 univer-py-1 univer-text-sm univer-text-gray-900 univer-transition-colors dark:!univer-text-white", {
              "univer-bg-gray-200 dark:!univer-bg-gray-600": c === N
            }),
            onMouseEnter: () => O(N),
            onMouseLeave: k,
            onClick: () => y(N),
            children: [
              d === N && /* @__PURE__ */ M(
                _r,
                {
                  className: "univer-absolute univer-left-1.5 univer-top-1/2 univer-inline-flex -univer-translate-y-1/2 univer-text-base univer-text-primary-600"
                }
              ),
              /* @__PURE__ */ M("span", { className: "univer-block", children: b(E) })
            ]
          },
          N
        ))
      }
    ),
    u && /* @__PURE__ */ V("div", { className: re("univer-mx-0 univer-my-2 univer-overflow-y-auto", nt), children: [
      /* @__PURE__ */ M(ut, { title: u.functionName, value: u.description }),
      /* @__PURE__ */ M(
        ut,
        {
          title: p.t("formula.moreFunctions.syntax"),
          value: /* @__PURE__ */ M(vr, { prefix: u.functionName, value: u.functionParameter })
        }
      ),
      /* @__PURE__ */ M(
        ut,
        {
          title: p.t("formula.prompt.helpExample"),
          value: `${u.functionName}(${u.functionParameter.map((E) => E.example).join(",")})`
        }
      ),
      u.functionParameter && u.functionParameter.map((E) => /* @__PURE__ */ M(
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
  const c = T(ot), a = T(je), g = T(ne), m = T(me);
  function d() {
    n(!e), r(!o);
  }
  function h() {
    const u = ir(g);
    if (!u) return;
    m.executeCommand(zn.id, {
      visible: !0,
      unitId: u.unitId,
      eventType: xe.Dblclick
    });
    const l = a.getEditor(Jt), v = a.getEditor(jn), p = `=${s == null ? void 0 : s.functionName}(`;
    l == null || l.replaceText(p), v == null || v.replaceText(p, !1);
  }
  return /* @__PURE__ */ V(
    "div",
    {
      "data-u-comp": "sheets-formula-functions-panel",
      className: "univer-box-border univer-flex univer-h-full univer-flex-col univer-justify-between univer-py-2",
      children: [
        e && /* @__PURE__ */ M(vi, { onChange: i }),
        o && /* @__PURE__ */ M(li, { functionInfo: s, onChange: () => {
        } }),
        /* @__PURE__ */ V("div", { className: "univer-flex univer-justify-end", children: [
          o && /* @__PURE__ */ M(
            et,
            {
              variant: "primary",
              onClick: d,
              className: "univer-mb-5 univer-ml-4 univer-mr-0 univer-mt-0",
              children: c.t("formula.moreFunctions.next")
            }
          ),
          o && /* @__PURE__ */ M(et, { onClick: d, className: "univer-mb-5 univer-ml-4 univer-mr-0 univer-mt-0", children: c.t("formula.moreFunctions.prev") }),
          e && !!t && /* @__PURE__ */ M(
            et,
            {
              variant: "primary",
              onClick: h,
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
  return t.get(ne).getCurrentTypeOfUnit$(j.UNIVER_SHEET).pipe(
    Fo((o) => o ? t.get(Fn) ? new Lo((s) => s.next(!t.get(Fn).supportClipboard)) : Nn(!0) : Nn(!0))
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
  L.ARROW_DOWN,
  L.ARROW_UP,
  L.ARROW_LEFT,
  L.ARROW_RIGHT
], yi = [...Ht, L.ENTER, L.TAB, L.ESC];
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
function Oi() {
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
function Ni() {
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
  binding: L.F4,
  preconditions: (t) => mt(t)
};
function wi() {
  const t = [];
  for (const e of [L.ENTER, L.TAB, L.ARROW_DOWN, L.ARROW_UP])
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
  binding: P.ALT | L.EQUAL,
  preconditions: mo,
  mac: P.CTRL_COMMAND | P.ALT | L.EQUAL,
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
      Os,
      Fs,
      Ls,
      xs,
      it,
      gr
    ].forEach((t) => this.disposeWithMe(this._commandService.registerCommand(t)));
  }
  _registerShortcuts() {
    [
      ...Ti(),
      ...xi(),
      ...Oi(),
      ...Ni(),
      ...wi(),
      ki,
      Mi
    ].forEach((t) => {
      this.disposeWithMe(this._shortcutService.registerShortcut(t));
    });
  }
  _registerComponents() {
    this.disposeWithMe(this._uiPartsService.registerComponent(po.FORMULA_AUX, () => ar(ii, this._injector))), this._componentManager.register(fr, Ci);
  }
  _registerRenderModules() {
    this.disposeWithMe(this._renderManagerService.registerRenderModule(j.UNIVER_SHEET, [kt]));
  }
};
Dt = Di([
  He(0, B(ft)),
  He(1, hs),
  He(2, me),
  He(3, Wt),
  He(4, cr),
  He(5, Oe),
  He(6, B(lr))
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
var Fi = Object.getOwnPropertyDescriptor, Li = (t, e, n, o) => {
  for (var r = o > 1 ? void 0 : o ? Fi(e, n) : e, s = t.length - 1, i; s >= 0; s--)
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
    return c.add(r == null ? void 0 : r.onPointerDown$.subscribeEvent((a, g) => {
      this.inRefSelectionMode() && (this._onPointerDown(a, r.zIndex + 1, Se.NORMAL, this._getActiveViewport(a)), a.button !== 2 && g.stopPropagation());
    })), c.add(
      n == null ? void 0 : n.onPointerDown$.subscribeEvent((a, g) => {
        if (!this.inRefSelectionMode()) return;
        const m = this._sheetSkeletonManagerService.getCurrent().skeleton, { row: d } = bn(a.offsetX, a.offsetY, i, m);
        yn(this._workbookSelections.getCurrentSelections(), d, Se.ROW) || (this._onPointerDown(a, (r.zIndex || 1) + 1, Se.ROW, this._getActiveViewport(a), Bt.Y), a.button !== 2 && g.stopPropagation());
      })
    ), c.add(o == null ? void 0 : o.onPointerDown$.subscribeEvent((a, g) => {
      if (!this.inRefSelectionMode()) return;
      const m = this._sheetSkeletonManagerService.getCurrent().skeleton, { column: d } = bn(a.offsetX, a.offsetY, i, m);
      yn(this._workbookSelections.getCurrentSelections(), d, Se.COLUMN) || (this._onPointerDown(a, (r.zIndex || 1) + 1, Se.COLUMN, this._getActiveViewport(a), Bt.X), a.button !== 2 && g.stopPropagation());
    })), c.add(s == null ? void 0 : s.onPointerDown$.subscribeEvent((a, g) => {
      if (this._reset(), !this.inRefSelectionMode()) return;
      const m = this._sheetSkeletonManagerService.getCurrent().skeleton, d = vo(m);
      this._addSelectionControlByModelData(d), this._selectionMoveStart$.next(this.getSelectionDataWithStyle());
      const h = i.onPointerUp$.subscribeEvent(() => {
        h.unsubscribe(), this._selectionMoveEnd$.next(this.getSelectionDataWithStyle());
      });
      a.button !== 2 && g.stopPropagation();
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
    const { offsetX: a, offsetY: g } = e, m = c.getViewport(wn.VIEW_MAIN);
    if (!m) return;
    const d = c.getCoordRelativeToViewport(kn.FromArray([a, g])), { x: h, y: u } = d;
    this._startViewportPosX = h, this._startViewportPosY = u;
    const l = c.getScrollXYInfoByViewport(d), { scaleX: v, scaleY: p } = c.getAncestorScale(), S = this._skeleton.getCellByOffset(h, u, v, p, l);
    if (!S) return;
    switch (o) {
      case Se.NORMAL:
        break;
      case Se.ROW:
        S.startColumn = 0, S.endColumn = this._skeleton.getColumnCount() - 1;
        break;
      case Se.COLUMN:
        S.startRow = 0, S.endRow = this._skeleton.getRowCount() - 1;
        break;
      case Se.ALL:
        S.startRow = 0, S.startColumn = 0, S.endRow = this._skeleton.getRowCount() - 1, S.endColumn = this._skeleton.getColumnCount() - 1;
    }
    const C = { range: S, primary: S, style: null };
    C.range.rangeType = o;
    const _ = Yt(C, this._skeleton);
    this._startRangeWhenPointerDown = { ..._.rangeWithCoord };
    const I = { ..._.rangeWithCoord, rangeType: o };
    let f = this.getActiveSelectionControl();
    const b = this.getSelectionControls();
    for (const O of b) {
      if (e.button === 2 && qn.contains(O.model, I)) {
        f = O;
        return;
      }
      if (O.model.isEqual(I)) {
        f = O;
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
    for (let O = 0; O < this.getSelectionControls().length - 1; O++)
      this.getSelectionControls()[O].clearHighlight();
    this._selectionMoveStart$.next(this.getSelectionDataWithStyle()), c.disableObjectsEvent(), this._clearUpdatingListeners(), this._addEndingListeners(), (x = c.getTransformer()) == null || x.clearSelectedObjects(), this._setupPointerMoveListener(m, f, o, s, h, u), this._escapeShortcutDisposable = this._shortcutService.forceEscape(), this._scenePointerUpSub = c.onPointerUp$.subscribeEvent(() => {
      var O;
      this._clearUpdatingListeners(), this._selectionMoveEnd$.next(this.getSelectionDataWithStyle()), (O = this._escapeShortcutDisposable) == null || O.dispose(), this._escapeShortcutDisposable = null;
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
    }), a = Yt(o, n);
    return c.updateRangeBySelectionWithCoord(a), this._selectionControls.push(c), c.setControlExtension({
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
De = Li([
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
const hn = (t, e, n = !0) => {
  let o = -1;
  return t.reduce((r, s, i) => {
    if (r.isFinish)
      return r;
    const c = r.currentIndex;
    if (typeof s != "string")
      r.currentIndex += s.token.length;
    else {
      const a = s.length;
      r.currentIndex += a;
    }
    return (n ? r.currentIndex === e : e > c && e <= r.currentIndex) && (o = i, r.isFinish = !0), r;
  }, { currentIndex: 0, isFinish: !1 }), o;
}, xr = (t, e) => {
  const n = t[e];
  let o = -1;
  if (!n || typeof n == "string" || n.nodeType !== Q.REFERENCE) return -1;
  for (let r = 0; r <= e; r++) {
    const s = t[r];
    typeof s != "string" && s.nodeType === Q.REFERENCE && o++;
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
function Or(t, e, n) {
  const o = T(je), r = ie(() => new rr({ left: -999, top: -999, right: -999, bottom: -999 }), []), s = T(ln), i = T(ne), c = te(() => {
    var _;
    const a = o.getEditor(t);
    if (!a)
      return;
    const g = a.getBoundingClientRect(), { marginTop: m = 0, marginBottom: d = 0 } = a.getDocumentData().documentStyle, h = a.getSkeleton();
    if (!h) return;
    const u = (_ = h.getSkeletonData()) == null ? void 0 : _.pages[0].height;
    let { left: l, top: v, right: p, bottom: S } = g;
    v = v + m, S = u ? v + u : S - d;
    const C = r.getValue();
    if (!(C.left === l && C.top === v && C.right === p && C.bottom === S))
      return r.next({ left: l - 1, right: p + 1, top: v - 1, bottom: S + 1 }), g;
  });
  return $(() => {
    e && c();
  }, [t, o, i.unitAdded$, c, e, ...n != null ? n : []]), Pi(c), $(() => {
    const a = s.scrollEvent$.pipe($o(100)).subscribe(c);
    return () => {
      a.unsubscribe();
    };
  }, []), [r, c];
}
const Ze = (t) => {
  const e = G(t);
  return e.current = t, e;
}, Ui = (t, e, n) => {
  const o = T(Vt), r = T(Ut), s = T(Ce), [i, c] = W(), [a, g] = W(-1), [m, d] = W(!0), h = Ze(m), u = G(e);
  u.current = e;
  const l = () => {
    c(void 0), g(-1), d(!1);
  };
  return $(() => {
    const v = s.sequenceNodesBuilder(e.slice(1));
    o.setSequenceNodes(v != null ? v : []);
  }, [e]), $(() => {
    if (n && t) {
      const v = n.selectionChange$.pipe(rn(50)).subscribe((S) => {
        if (S.textRanges.length === 1) {
          const [C] = S.textRanges;
          if (C.collapsed && h.current) {
            const { startOffset: _ } = C, I = o.getCurrentSequenceNodeIndex(_ - 2), f = o.getCurrentSequenceNodeByIndex(I), b = o.getCurrentSequenceNodeByIndex(I + 1);
            if (f)
              if (typeof f != "string" && f.nodeType === 3 && !r.hasDefinedNameDescription(f.token.trim()) && b === qe.OPEN_BRACKET) {
                const y = r.getFunctionInfo(f.token);
                c(y), g(-1);
                return;
              } else {
                const y = s.getFunctionAndParameter(`${u.current}A`, _ - 1);
                if (y) {
                  const { functionName: A, paramIndex: D } = y, x = r.getFunctionInfo(A);
                  c(x), g(D);
                  return;
                }
              }
          }
        }
        c(void 0), g(-1);
      }), p = n.selectionChange$.pipe(
        or((S) => S.textRanges.length === 1),
        nr((S) => S.textRanges[0].startOffset),
        Po()
      ).subscribe(() => {
        d(!0);
      });
      return () => {
        v.unsubscribe(), p.unsubscribe();
      };
    }
  }, [n, t]), $(() => {
    t || l();
  }, [t]), {
    functionInfo: i,
    paramIndex: a,
    reset: l
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
  const { onParamsSwitch: e = Vn, onClose: n = Vn, isFocus: o, editor: r, formulaText: s } = t, { functionInfo: i, paramIndex: c, reset: a } = Ui(o, s, r), g = T(nn), m = !pe(g.helpFunctionVisible$), [d, h] = W(!0), u = T(ot), l = u.t("formula.prompt.required"), v = u.t("formula.prompt.optional"), p = r.getEditorId(), [S] = Or(p, !!i, [i, c]);
  function C(f) {
    e && e(f);
  }
  const _ = te((f) => {
    g.helpFunctionVisible$.next(!f);
  }), I = () => {
    _(!0), n();
  };
  return i ? m ? /* @__PURE__ */ M(Zt, { portal: !0, anchorRect$: S, direction: "left-center", children: /* @__PURE__ */ M(Wi, { onClick: () => _(!1) }) }, "hidden") : /* @__PURE__ */ M(Zt, { portal: !0, onClickOutside: () => a(), anchorRect$: S, direction: "vertical", children: /* @__PURE__ */ V(
    "div",
    {
      className: re("univer-m-0 univer-box-border univer-w-[250px] univer-select-none univer-list-none univer-rounded-lg univer-bg-white univer-leading-5 univer-shadow-md univer-outline-none dark:!univer-bg-gray-900", an),
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
                    style: { transform: d ? "rotateZ(-90deg)" : "rotateZ(90deg)" },
                    onClick: () => h(!d),
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
              height: d ? "unset" : 0,
              padding: d ? "revert-layer" : 0
            },
            children: /* @__PURE__ */ V("div", { className: "univer-mt-3", children: [
              /* @__PURE__ */ M(
                Kt,
                {
                  title: u.t("formula.prompt.helpExample"),
                  value: `${i.functionName}(${i.functionParameter.map((f) => f.example).join(",")})`
                }
              ),
              /* @__PURE__ */ M(
                Kt,
                {
                  title: u.t("formula.prompt.helpAbstract"),
                  value: i.description
                }
              ),
              i && i.functionParameter && i.functionParameter.map((f, b) => /* @__PURE__ */ M(
                Kt,
                {
                  className: c === b ? "univer-text-primary-500" : "",
                  title: f.name,
                  value: `${f.require ? l : v} ${f.detail}`
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
`, a = Math.max(c.length - 2, 0);
        t.setSelectionRanges([{ startOffset: a, endOffset: a }]);
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
  const { editorId: e, isFocus: n, disableOnClick: o, unitId: r, subUnitId: s } = t, i = T(Oe), c = T(ne), a = i.getRenderById(r), g = i.getRenderById(e), m = g == null ? void 0 : g.with(sr), d = T(hr), h = T(ft), [u, l] = W(
    0
    /* NOT_SELECT */
  ), v = T(Ce), p = G(!0), S = a == null ? void 0 : a.with(De), C = Ze(u), _ = c.getUnit(r, j.UNIVER_SHEET), I = _ == null ? void 0 : _.getSheetBySheetId(s), f = te((A) => {
    S && S.setSkipLastEnabled(
      A === 1 || A === 3 || A === 4
      /* EDIT_OTHER_WORKBOOK_REFERENCE */
    ), C.current = A, l(A);
  }), b = te(() => {
    var ce, le, de;
    const A = c.getCurrentUnitOfType(j.UNIVER_SHEET);
    if (!A) return;
    const D = A.getActiveSheet(), x = m == null ? void 0 : m.getActiveTextRange(), O = x != null && x.collapsed ? x.startOffset : -1, k = qi(h);
    if (!k) return;
    const E = (ce = k == null ? void 0 : k.dataStream) == null ? void 0 : ce.slice(0, -2), N = ((le = v.sequenceNodesBuilder(E)) != null ? le : []).map((K) => typeof K == "object" ? K.nodeType === Q.REFERENCE ? {
      ...K,
      range: xo(K.token)
    } : {
      ...K,
      range: void 0
    } : K), R = E[O - 1], w = E[O], F = N.find((K) => typeof K == "object" && K.nodeType === Q.REFERENCE && O === K.endIndex + 2), U = R && Oo(R) && (!w || No(w) && w !== qe.OPEN_BRACKET), H = !!F;
    if ((E == null ? void 0 : E.substring(0, 1)) === "=" && (U || H))
      if (H) {
        if (p.current)
          return;
        const { sheetName: K, unitId: J } = F.range, ze = (de = c.getCurrentUnitOfType(j.UNIVER_SHEET)) == null ? void 0 : de.getUnitId();
        J && J !== ze ? f(
          4
          /* EDIT_OTHER_WORKBOOK_REFERENCE */
        ) : !K && D.getSheetId() === (I == null ? void 0 : I.getSheetId()) || K === D.getName() ? f(
          2
          /* CAN_EDIT */
        ) : f(
          3
          /* EDIT_OTHER_SHEET_REFERENCE */
        );
      } else
        p.current = !1, f(
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
    const A = d.textSelection$.pipe(or((D) => D.unitId === e)).subscribe(() => {
      b();
    });
    return () => A.unsubscribe();
  }, [b, d.textSelection$, e]), $(() => {
    n || (f(
      0
      /* NOT_SELECT */
    ), p.current = !0);
  }, [n, f]), $(() => {
    var D;
    if (!o) return;
    const A = (D = g == null ? void 0 : g.mainComponent) == null ? void 0 : D.onPointerDown$.subscribeEvent(() => {
      f(
        0
        /* NOT_SELECT */
      ), p.current = !0;
    });
    return () => A == null ? void 0 : A.unsubscribe();
  }, [o, (y = g == null ? void 0 : g.mainComponent) == null ? void 0 : y.onPointerDown$, f]), $(() => {
    if (!n) return;
    const A = _ == null ? void 0 : _.activeSheet$.subscribe(() => {
      b();
    }), D = c.getCurrentTypeOfUnit$(j.UNIVER_SHEET).subscribe(() => {
      b();
    });
    return () => {
      A == null || A.unsubscribe(), D == null || D.unsubscribe();
    };
  }, [b, n, _ == null ? void 0 : _.activeSheet$, c.getCurrentTypeOfUnit$]), { isSelecting: u, isSelectingRef: C };
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
function Nr(t) {
  var _, I, f;
  const {
    unitId: e,
    subUnitId: n,
    currentWorkbook: o,
    refSelections: r,
    editor: s,
    refSelectionsService: i,
    refSelectionsRenderService: c,
    sheetSkeletonManagerService: a,
    themeService: g,
    univerInstanceService: m
  } = t, d = o.getUnitId(), h = m.getUnit(e, j.UNIVER_SHEET), u = h == null ? void 0 : h.getActiveSheet(), l = [];
  if (!h || !u) {
    i.setSelections(l);
    return;
  }
  const v = u.getSheetId(), p = (b) => {
    var y;
    return (y = h == null ? void 0 : h.getSheetBySheetName(b)) == null ? void 0 : y.getSheetId();
  };
  if (!((_ = a == null ? void 0 : a.getWorksheetSkeleton(v)) == null ? void 0 : _.skeleton)) return;
  const C = [];
  for (let b = 0, y = r.length; b < y; b++) {
    const A = r[b], { themeColor: D, token: x, refIndex: O, endIndex: k } = A, E = st(x), { unitId: N, sheetName: R, range: w } = E, F = p(R);
    if (!F && R || d !== e && N !== d || N && N !== d || F && F !== v || !F && v !== n)
      continue;
    const U = rs(w, u.getRowCount(), u.getColumnCount());
    U.unitId = e, U.sheetId = v, l.push({
      range: U,
      primary: null,
      style: Yi(g, D, O.toString())
    }), C.push(k);
  }
  if (s) {
    const b = (f = (I = s.getSelectionRanges()) == null ? void 0 : I[0]) == null ? void 0 : f.startOffset, y = C.findIndex((A) => A + 2 === b);
    y !== -1 ? c == null || c.setActiveSelectionIndex(y) : c == null || c.resetActiveSelectionIndex();
  }
  return l;
}
function Zi(t, e) {
  const n = T(ne), o = T(gt), r = T(Pt), s = T(Oe), i = pe(ie(() => n.getCurrentTypeOfUnit$(j.UNIVER_SHEET), [n])), c = i ? s.getRenderById(i.getUnitId()) : null, a = c == null ? void 0 : c.with(De), g = c == null ? void 0 : c.with($t), m = te((d, h) => {
    const u = n.getCurrentUnitOfType(j.UNIVER_SHEET);
    if (!u || a != null && a.selectionMoving) return;
    const l = Nr({
      unitId: t,
      subUnitId: e,
      currentWorkbook: u,
      refSelections: d,
      editor: h,
      refSelectionsService: r,
      refSelectionsRenderService: a,
      sheetSkeletonManagerService: g,
      themeService: o,
      univerInstanceService: n
    });
    if (!l) return;
    ((a == null ? void 0 : a.getSelectionControls()) || []).length === l.length ? a == null || a.resetSelectionsByModelData(l) : r.setSelections(l);
  });
  return $(() => () => {
    a == null || a.resetActiveSelectionIndex();
  }, [a]), m;
}
function Mr(t = "") {
  const e = T(Ut), n = zi(), o = T(me), r = ie(() => t.length, [t]);
  return te((i, c, a = !0, g) => {
    const m = i.getDocumentData(), d = i.getEditorId();
    if (!m)
      return [];
    const h = m.body;
    if (!h)
      return [];
    const u = h.dataStream.slice(0, h.dataStream.length - 2), l = { dataStream: "", ...m.body };
    if (!u.startsWith(t)) return [];
    if (c == null || c.length === 0)
      return l.textRuns = [], o.syncExecuteCommand(Mn.id, {
        unitId: d,
        body: Rn(l, 0, l.dataStream.length - 2)
      }), [];
    {
      const { textRuns: v, refSelections: p } = Gi(e, n, c);
      r && v.forEach((_) => {
        _.ed = _.ed + r, _.st = _.st + r;
      }), l.textRuns = [{ st: 0, ed: 1, ts: { fs: 11 } }, ...v];
      const S = c.reduce((_, I) => typeof I == "string" ? `${_}${I}` : `${_}${I.token}`, "");
      l.dataStream = `${t}${S}\r
`;
      let C;
      if (a) {
        C = i.getSelectionRanges();
        const _ = l.dataStream.length - 2 + r;
        C.forEach((I) => {
          I.startOffset = Math.max(0, Math.min(I.startOffset, _)), I.endOffset = Math.max(0, Math.min(I.endOffset, _));
        });
      }
      return o.syncExecuteCommand(Mn.id, {
        unitId: d,
        body: Rn(l, 0, l.dataStream.length - 2),
        textRanges: g != null ? g : C
      }), p;
    }
  });
}
function zi() {
  const t = T(gt), e = t.getCurrentTheme();
  return ie(() => {
    const o = [
      t.getColorFromTheme("loop-color.1"),
      t.getColorFromTheme("loop-color.2"),
      t.getColorFromTheme("loop-color.3"),
      t.getColorFromTheme("loop-color.4"),
      t.getColorFromTheme("loop-color.5"),
      t.getColorFromTheme("loop-color.6"),
      t.getColorFromTheme("loop-color.7"),
      t.getColorFromTheme("loop-color.8"),
      t.getColorFromTheme("loop-color.9"),
      t.getColorFromTheme("loop-color.10"),
      t.getColorFromTheme("loop-color.11"),
      t.getColorFromTheme("loop-color.12")
    ].map((c) => t.isValidThemeColor(c) ? t.getColorFromTheme(c) : c), r = t.getColorFromTheme("blue.700"), s = t.getColorFromTheme("jiqing.800"), i = t.getColorFromTheme("black");
    return { formulaRefColors: o, numberColor: r, stringColor: s, plainTextColor: i };
  }, [e]);
}
function Gi(t, e, n) {
  const { formulaRefColors: o, numberColor: r, stringColor: s, plainTextColor: i } = e, c = [], a = [], g = /* @__PURE__ */ new Map();
  let m = 0;
  for (let d = 0, h = n.length; d < h; d++) {
    const u = n[d];
    if (typeof u == "string") {
      const _ = c[c.length - 1], I = _ ? _.ed : 0, f = I + u.length;
      c.push({
        st: I,
        ed: f,
        ts: {
          cl: {
            rgb: i
          },
          fs: 11
        }
      });
      continue;
    }
    if (t.hasDefinedNameDescription(u.token.trim())) {
      c.push({
        st: u.startIndex,
        ed: u.endIndex + 1,
        ts: {
          cl: {
            rgb: i
          },
          fs: 11
        }
      });
      continue;
    }
    const { startIndex: l, endIndex: v, nodeType: p, token: S } = u;
    let C = "";
    if (p === Q.REFERENCE) {
      if (g.has(S))
        C = g.get(S);
      else {
        const _ = m % o.length;
        C = o[_], g.set(S, C), m++;
      }
      a.push({
        refIndex: d,
        themeColor: C,
        token: S,
        startIndex: u.startIndex,
        endIndex: u.endIndex,
        index: a.length
      });
    } else p === Q.NUMBER ? C = r : (p === Q.STRING || p === Q.ARRAY) && (C = s);
    C && C.length > 0 ? c.push({
      st: l,
      ed: v + 1,
      ts: {
        cl: {
          rgb: C
        },
        fs: 11
      }
    }) : c.push({
      st: l,
      ed: v + 1,
      ts: {
        cl: {
          rgb: i
        },
        fs: 11
      }
    });
  }
  return { textRuns: c, refSelections: a };
}
const Xi = (t, e, n, o) => {
  const r = T(me), s = T(Wt), i = G(e);
  i.current = e;
  const c = G(o);
  c.current = o, $(() => {
    if (!n || !t)
      return;
    const g = `sheet.formula-embedding-editor.${n.getEditorId()}`, m = new Be(), d = (l, v) => {
      if (c.current) {
        c.current(l, v);
        return;
      }
      let p = he.LEFT;
      l === L.ARROW_DOWN ? p = he.DOWN : l === L.ARROW_UP ? p = he.UP : l === L.ARROW_RIGHT && (p = he.RIGHT), v === P.SHIFT ? r.executeCommand(Wo.id, {
        direction: p
      }) : r.executeCommand(Vo.id, {
        direction: p
      });
    }, h = (l, v) => {
      let p = he.DOWN;
      l === L.ARROW_DOWN ? p = he.DOWN : l === L.ARROW_UP ? p = he.UP : l === L.ARROW_LEFT ? p = he.LEFT : l === L.ARROW_RIGHT && (p = he.RIGHT), i.current ? v === P.CTRL_COMMAND ? r.executeCommand(Tn.id, {
        direction: p,
        jumpOver: xn.moveGap,
        extra: "formula-editor",
        fromCurrentSelection: i.current === ye.NEED_ADD || i.current === ye.EDIT_OTHER_SHEET_REFERENCE
      }) : v === P.SHIFT ? r.executeCommand(On.id, {
        direction: p,
        extra: "formula-editor"
      }) : v === (P.CTRL_COMMAND | P.SHIFT) ? r.executeCommand(On.id, {
        direction: p,
        jumpOver: xn.moveGap,
        extra: "formula-editor"
      }) : r.executeCommand(Tn.id, {
        direction: p,
        extra: "formula-editor",
        fromCurrentSelection: i.current === ye.NEED_ADD || i.current === ye.EDIT_OTHER_SHEET_REFERENCE
      }) : d(l, v);
    };
    return m.add(r.registerCommand({
      id: g,
      type: Te.OPERATION,
      handler(l, v) {
        const { keyCode: p, metaKey: S } = v;
        h(p, S);
      }
    })), [
      { keyCode: L.ARROW_DOWN },
      { keyCode: L.ARROW_LEFT },
      { keyCode: L.ARROW_RIGHT },
      { keyCode: L.ARROW_UP },
      { keyCode: L.ARROW_DOWN, metaKey: P.SHIFT },
      { keyCode: L.ARROW_LEFT, metaKey: P.SHIFT },
      { keyCode: L.ARROW_RIGHT, metaKey: P.SHIFT },
      { keyCode: L.ARROW_UP, metaKey: P.SHIFT },
      { keyCode: L.ARROW_DOWN, metaKey: P.CTRL_COMMAND },
      { keyCode: L.ARROW_LEFT, metaKey: P.CTRL_COMMAND },
      { keyCode: L.ARROW_RIGHT, metaKey: P.CTRL_COMMAND },
      { keyCode: L.ARROW_UP, metaKey: P.CTRL_COMMAND },
      { keyCode: L.ARROW_DOWN, metaKey: P.CTRL_COMMAND | P.SHIFT },
      { keyCode: L.ARROW_LEFT, metaKey: P.CTRL_COMMAND | P.SHIFT },
      { keyCode: L.ARROW_RIGHT, metaKey: P.CTRL_COMMAND | P.SHIFT },
      { keyCode: L.ARROW_UP, metaKey: P.CTRL_COMMAND | P.SHIFT }
    ].map(({ keyCode: l, metaKey: v }) => ({
      id: g,
      binding: v ? l | v : l,
      preconditions: () => !0,
      priority: 900,
      staticParameters: {
        eventType: xe.Keyboard,
        keyCode: l,
        metaKey: v
      }
    })).forEach((l) => {
      m.add(s.registerShortcut(l));
    }), () => {
      m.dispose();
    };
  }, [r, n, t, s]);
}, Qi = (t, e, n, o, r = !0) => {
  var u;
  const s = T(Oe), i = T(Qt), c = T(ds), a = T(Pt), g = T(ne), m = pe(ie(() => g.getCurrentTypeOfUnit$(j.UNIVER_SHEET), [g])), d = s.getRenderById((u = m == null ? void 0 : m.getUnitId()) != null ? u : ""), h = d == null ? void 0 : d.with(De);
  Ot(() => {
    if (t)
      return i.setContextValue(En, !0), r && c.disable(), () => {
        const l = g.getCurrentUnitOfType(j.UNIVER_DOC);
        (l == null ? void 0 : l.getUnitId()) === o && i.setContextValue(En, !1), r && c.enable(), a.clear();
      };
  }, [i, t, a, r, o]), Ot(() => {
    if (t && e) {
      const l = h == null ? void 0 : h.enableSelectionChanging();
      return i.setContextValue(An, !0), () => {
        i.setContextValue(An, !1), l == null || l.dispose();
      };
    }
  }, [i, t, h, e]), $(() => {
    t && (h == null || h.setSkipLastEnabled(!1));
  }, [t, h]);
}, Ji = (t, e, n) => {
  const o = T(ne), r = T(on);
  return xt(() => {
    if (t) {
      const i = [...r.getWorkbookSelections(e).getSelectionsOfWorksheet(n)], c = o.getCurrentUnitForType(j.UNIVER_SHEET), a = c == null ? void 0 : c.getActiveSheet();
      (c == null ? void 0 : c.getUnitId()) !== e && o.setCurrentUnitForType(e), a && a.getSheetId() === n && r.setSelections(i);
    }
  }, [t, r, n, e, o]);
}, ec = (t) => t.reduce((e, n) => typeof n == "string" ? e + n.length : e + n.token.length, 0), Xt = (t) => t.map((e) => typeof e == "string" ? e : e.token).join(""), bt = (t, e = !1, n = "", o = !1) => !e && !o ? t.map((r) => ve(r.range)) : t.map((r) => o ? Mo(r) : r.sheetName !== "" && r.sheetName !== n ? ht(r.sheetName, r.range) : ve(r.range)), tc = (t) => {
  var m, d, h;
  const { editor: e, lexerTreeBuilder: n } = t, o = e == null ? void 0 : e.getSelectionRanges();
  if ((o == null ? void 0 : o.length) !== 1)
    return;
  const s = o[0].startOffset - 1, i = ((d = (m = e == null ? void 0 : e.getDocumentData().body) == null ? void 0 : m.dataStream) != null ? d : `\r
`).slice(0, -2), c = (h = n.sequenceNodesBuilder(i.slice(1))) != null ? h : [], a = hn(c, s, !1), g = xr(c, a);
  return {
    nodeIndex: a,
    updatingRefIndex: g,
    sequenceNodes: c,
    offset: s
  };
}, nc = (() => {
}), rc = (t, e, n, o, r, s, i, c, a, g = nc) => {
  var O;
  const m = T(Oe), d = T(ne), h = T(me), u = T(hr), l = T(gt), v = T(Ce), p = d.getUnit(o), S = te((k, E) => {
    var N, R, w;
    return (w = (R = (N = d.getUnit(k)) == null ? void 0 : N.getSheetBySheetId(E)) == null ? void 0 : R.getName()) != null ? w : "";
  }), C = ie(() => S(o, r), [S, r, o]), _ = pe(p == null ? void 0 : p.activeSheet$), I = Ze({ activeSheet: _, sheetName: C }), f = pe(ie(() => d.getCurrentTypeOfUnit$(j.UNIVER_SHEET), [d])), b = m.getRenderById((O = f == null ? void 0 : f.getUnitId()) != null ? O : ""), y = b == null ? void 0 : b.with(De), A = b == null ? void 0 : b.with($t), D = T(Pt), x = te((k, E) => {
    var H, ce, le, de, K, J, ze, Ge, St, _e;
    const N = tc({ editor: a, lexerTreeBuilder: v });
    if (!N) return;
    const { nodeIndex: R, updatingRefIndex: w, sequenceNodes: F, offset: U } = N;
    if (n.current === ye.NEED_ADD)
      if (U !== 0) {
        if (R === -1 && F.length)
          return;
        const Y = k[k.length - 1], X = F.splice(R + 1), Re = (H = Y.sheetId) != null ? H : r, fe = {
          range: Y,
          unitId: (ce = Y.unitId) != null ? ce : f.getUnitId(),
          sheetName: S((le = Y.unitId) != null ? le : f.getUnitId(), Re)
        }, ee = Re !== r, ae = (f == null ? void 0 : f.getUnitId()) !== o, Ee = bt([fe], i && (ee || ae), C, ae);
        F.push({ token: Ee[0], nodeType: Q.REFERENCE });
        const Le = [...F, ...X], Ne = Xt(Le);
        g(Ne, ec(F), E);
      } else {
        const Y = k[k.length - 1], X = (de = Y.sheetId) != null ? de : r, Re = {
          range: Y,
          unitId: (K = Y.unitId) != null ? K : f.getUnitId(),
          sheetName: S((J = Y.unitId) != null ? J : f.getUnitId(), X)
        }, fe = X !== r, ee = (f == null ? void 0 : f.getUnitId()) !== o, ae = bt([Re], i && (fe || ee), C, ee);
        F.unshift({ token: ae[0], nodeType: Q.REFERENCE });
        const Ee = Xt(F);
        g(Ee, ae[0].length, E);
      }
    else if (n.current === ye.EDIT_OTHER_SHEET_REFERENCE || n.current === ye.EDIT_OTHER_WORKBOOK_REFERENCE) {
      const Y = k.pop();
      if (!Y) return;
      const X = F[R];
      if (typeof X == "object" && X.nodeType === Q.REFERENCE) {
        const Re = X.token;
        (f == null ? void 0 : f.getUnitId()) !== o ? X.token = wo((ze = f == null ? void 0 : f.getUnitId()) != null ? ze : "", C, Y) : X.token = C === (_ == null ? void 0 : _.getName()) ? ve(Y) : ht(_.getName(), Y);
        const ee = U + (X.token.length - Re.length);
        g(ko(F), ee, E);
      }
    } else {
      const Y = [...k];
      if (w !== -1) {
        const Z = Y.pop();
        Z && Y.splice(w, 0, Z);
      }
      let X = 0;
      const Re = F.map((Z) => {
        var ge, Xe, $e, Pe;
        if (typeof Z == "string")
          return Z;
        if (Z.nodeType === Q.REFERENCE) {
          const Ue = st(Z.token);
          if (Ue.sheetName || (Ue.sheetName = C), (Ue.unitId || o) !== (f == null ? void 0 : f.getUnitId()) || i && ((ge = I.current.activeSheet) == null ? void 0 : ge.getName()) !== Ue.sheetName)
            return Z.token;
          const oe = Y[X];
          if (X++, !oe)
            return "";
          const Ie = (Xe = oe.sheetId) != null ? Xe : r, vt = {
            range: oe,
            unitId: ($e = oe.unitId) != null ? $e : f.getUnitId(),
            sheetName: S((Pe = oe.unitId) != null ? Pe : f.getUnitId(), Ie)
          }, Ct = (f == null ? void 0 : f.getUnitId()) !== o;
          return bt([vt], i && (Ie !== r || Ct), C, Ct)[0];
        }
        return Z.token;
      });
      let fe = "", ee;
      Re.forEach((Z, ge) => {
        fe += Z, ge === R && (ee = fe.length);
      });
      const ae = [];
      for (let Z = X; Z <= k.length - 1; Z++) {
        const ge = k[Z], Xe = (Ge = ge.sheetId) != null ? Ge : r, $e = {
          range: ge,
          unitId: (St = ge.unitId) != null ? St : f.getUnitId(),
          sheetName: S((_e = ge.unitId) != null ? _e : f.getUnitId(), Xe)
        }, Pe = (f == null ? void 0 : f.getUnitId()) !== o, oe = bt([$e], i && (Xe !== r || Pe), C, Pe);
        ae.push(oe[0]);
      }
      const Ee = F[F.length - 1], Le = Ee && (typeof Ee == "string" ? !1 : Ee.nodeType === Q.REFERENCE), Ne = `${fe}${ae.length && Le ? "," : ""}${ae.join(",")}`;
      g(Ne, !ae.length && ee ? ee : Ne.length, E);
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
        x(R.map((F) => F.rangeWithCoord), w);
      }, N = new Be();
      return N.add(y.selectionMoving$.subscribe((R) => {
        E(R, !1);
      })), N.add(y.selectionMoveEnd$.subscribe((R) => {
        E(R, !0);
      })), () => {
        N.dispose();
      };
    }
  }, [t, x, y]), $(() => {
    if (e && y && a) {
      const k = new Be(), E = () => {
        k.dispose(), y.getSelectionControls().forEach((w, F) => {
          k.add(
            w.selectionScaling$.subscribe((U) => {
              const H = y.getSelectionDataWithStyle().map((le) => le.rangeWithCoord), ce = H[F];
              U.sheetId = ce.sheetId, U.unitId = ce.unitId, H[F] = U, x(H, !1);
            })
          ), k.add(
            w.selectionMoving$.subscribe((U) => {
              const H = y.getSelectionDataWithStyle().map((le) => le.rangeWithCoord), ce = H[F];
              U.sheetId = ce.sheetId, U.unitId = ce.unitId, H[F] = U, x(H, !0);
            })
          );
        });
      }, N = Uo(
        a.input$,
        D.selectionSet$,
        y.selectionMoveEnd$
      ).pipe(
        Es(50)
      ).subscribe(() => {
        E();
      });
      return () => {
        N.unsubscribe(), k.dispose();
      };
    }
  }, [a, e, x, y, D.selectionSet$]), y == null || y.getSelectionDataWithStyle(), $(() => {
    if (c) {
      const k = h.onCommandExecuted((E) => {
        var R;
        if (E.id !== sn.id)
          return;
        const N = E.params;
        if (N.extra === "formula-editor" && N.selections.length) {
          const w = N.selections[N.selections.length - 1];
          if (w) {
            const F = n.current === ye.NEED_ADD, U = ((R = y == null ? void 0 : y.getSelectionDataWithStyle()) != null ? R : []).map((H) => H.rangeWithCoord);
            F ? U.push(w.range) : U[U.length - 1] = w.range, x(U, !0);
          }
        }
      });
      return () => {
        k.dispose();
      };
    }
  }, [h, a, n, v, c, x, y]), $(() => {
    if (!a)
      return;
    const k = u.textSelection$.subscribe((E) => {
      E.unitId === a.getEditorId() && Nr({
        unitId: o,
        subUnitId: r,
        refSelections: s.current,
        editor: a,
        refSelectionsService: D,
        refSelectionsRenderService: y,
        sheetSkeletonManagerService: A,
        themeService: l,
        univerInstanceService: d,
        currentWorkbook: f
      });
    });
    return () => k.unsubscribe();
  }, [u.textSelection$, a, s, y, D, A, r, l, o, d]);
}, oc = (t, e, n, o, r, s) => {
  const i = T(me), c = T(je), g = T(Oe).getRenderById(e), m = T(ne), d = g == null ? void 0 : g.with(De);
  $(() => {
    if (t && d)
      if (n) {
        const h = () => {
          const v = d.getSelectionControls().length;
          for (let p = 1; p <= v; p++)
            d.clearLastSelection();
          return setTimeout(() => {
            s();
          }, 30);
        }, u = i.onCommandExecuted((v) => {
          v.id === Dn.id && h();
        }), l = m.getCurrentTypeOfUnit$(j.UNIVER_SHEET).subscribe((v) => {
          h();
        });
        return () => {
          u.dispose(), l.unsubscribe();
        };
      } else {
        const h = i.beforeCommandExecuted((u) => {
          if (u.id === Dn.id) {
            o(!1), r(), s();
            const l = c.getEditor(Jt);
            l == null || l.focus();
          }
        });
        return () => {
          h.dispose();
        };
      }
  }, [t, d]);
}, sc = (t, e, n) => {
  const o = T(Ce), r = G(!0);
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
  const o = T(Ut), [r, s] = W([]), [i, c] = W(""), a = G(-1), g = Ze({ nodes: e }), m = () => {
    s([]), c(""), a.current = -1;
  };
  return $(() => {
    if (n && t) {
      const h = n.input$.pipe(rn(300)).subscribe(() => {
        const u = n.getSelectionRanges();
        if (u.length === 1) {
          const l = g.current.nodes, v = u[0];
          if (v.collapsed) {
            const p = hn(l, v.startOffset - 1, !1);
            a.current = p;
            const S = l[p];
            if (S && typeof S != "string" && S.nodeType === Q.FUNCTION) {
              a.current = p;
              const C = S.token, _ = o.getSearchListByNameFirstLetter(C);
              s(_), c(C);
              return;
            }
          }
        }
        a.current = -1, c(""), s((l) => l != null && l.length ? [] : l);
      });
      return () => {
        h.unsubscribe();
      };
    }
  }, [n, t]), $(() => {
    t || m();
  }, [t]), {
    searchList: r,
    searchText: i,
    handlerFormulaReplace: (h, u) => {
      const l = [...g.current.nodes];
      if (a.current !== -1) {
        const v = l.splice(a.current + 1), p = l.pop() || "";
        let S = (typeof p == "string" ? p.length : p.token.length) - h.length;
        return l.push(h), v[0] !== qe.OPEN_BRACKET && u !== er.DefinedName && (l.push(qe.OPEN_BRACKET), S--), { text: Xt([...l, ...v]), offset: S };
      }
    },
    reset: m
  };
}, cc = () => {
}, lc = Fe(ac);
function ac(t, e) {
  const { isFocus: n, sequenceNodes: o, onSelect: r, editor: s, onClose: i = cc } = t, c = s.getEditorId(), a = T(Wt), g = T(me), { searchList: m, searchText: d, handlerFormulaReplace: h, reset: u } = ic(n, o, s), l = ie(() => !!m.length, [m]), v = G(void 0), [p, S] = W(0), C = G(!1), [_] = Or(c, l, [d, m]), I = Ze({ searchList: m, active: p }), f = (x, O) => {
    const k = h(x, O);
    k && (u(), r(k));
  };
  function b(x) {
    C.current && S(x);
  }
  function y() {
    C.current && S(-1);
  }
  $(() => {
    if (!m.length)
      return;
    const x = `sheet.formula-embedding-editor.search_function.${c}`, O = new Be(), k = (E) => {
      const { searchList: N, active: R } = I.current;
      switch (E) {
        case L.ARROW_UP: {
          S((w) => {
            const F = Math.max(0, w - 1);
            return A(F), F;
          });
          break;
        }
        case L.ARROW_DOWN: {
          S((w) => {
            const F = Math.min(N.length - 1, w + 1);
            return A(F), F;
          });
          break;
        }
        case L.TAB:
        case L.ENTER: {
          const w = N[R];
          f(w.name, w.functionType);
          break;
        }
        case L.ESC: {
          u(), i();
          break;
        }
      }
    };
    return O.add(g.registerCommand({
      id: x,
      type: Te.OPERATION,
      handler(E, N) {
        const { keyCode: R } = N;
        k(R);
      }
    })), [L.ARROW_UP, L.ARROW_DOWN, L.ENTER, L.ESC, L.TAB].map((E) => ({
      id: x,
      binding: E,
      preconditions: () => !0,
      priority: 1e3,
      staticParameters: {
        eventType: xe.Keyboard,
        keyCode: E
      }
    })).forEach((E) => {
      O.add(a.registerShortcut(E));
    }), () => {
      O.dispose();
    };
  }, [m]);
  function A(x) {
    const O = v.current;
    if (!O) return;
    const k = O.children[x];
    if (!k) return;
    const N = O.getBoundingClientRect().top, R = O.offsetHeight, w = k.getBoundingClientRect(), F = w.top, U = w.height;
    if (F >= 0 && F > N && F - N + U <= R)
      return;
    const H = k.offsetTop - (R - U) / 2;
    O.scrollTo({
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
  return m.length > 0 && l && /* @__PURE__ */ M(Zt, { portal: !0, anchorRect$: _, direction: "vertical", children: /* @__PURE__ */ M(
    "ul",
    {
      ref: (x) => {
        v.current = x, e && (e.current = x);
      },
      "data-u-comp": "sheets-formula-editor",
      className: re("univer-m-0 univer-box-border univer-max-h-[400px] univer-w-[250px] univer-list-none univer-overflow-y-auto univer-rounded-lg univer-bg-white univer-p-2 univer-leading-5 univer-shadow-md univer-outline-none dark:!univer-bg-gray-900", an, nt),
      children: m.map((x, O) => /* @__PURE__ */ V(
        "li",
        {
          className: re("univer-box-border univer-cursor-pointer univer-rounded univer-px-2 univer-py-1 univer-text-gray-900 univer-transition-colors dark:!univer-text-white", {
            "univer-bg-gray-200 dark:!univer-bg-gray-600": p === O
          }),
          onMouseEnter: () => b(O),
          onMouseLeave: y,
          onMouseMove: D,
          onClick: () => {
            f(x.name, x.functionType), s && s.focus();
          },
          children: [
            /* @__PURE__ */ V("span", { className: "univer-block univer-overflow-x-hidden univer-text-ellipsis univer-text-xs", children: [
              /* @__PURE__ */ M("span", { className: "univer-text-red-500", children: x.name.substring(0, d.length) }),
              /* @__PURE__ */ M("span", { children: x.name.slice(d.length) })
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
}, hc = Fe((t, e) => {
  var gn, mn, pn, Sn;
  const {
    errorText: n,
    initValue: o,
    unitId: r,
    subUnitId: s,
    isFocus: i = !0,
    isSupportAcrossSheet: c = !1,
    onFocus: a = Hn,
    onBlur: g = Hn,
    onChange: m,
    onVerify: d,
    className: h,
    editorId: u,
    moveCursor: l = !0,
    onFormulaSelectingChange: v,
    keyboardEventConfig: p,
    onMoveInEditor: S,
    resetSelectionOnBlur: C = !0,
    autoScrollbar: _ = !0,
    isSingle: I = !0,
    disableSelectionOnClick: f = !1,
    autofocus: b = !0,
    disableContextMenu: y,
    style: A
  } = t, D = T(je), x = G(null), O = te(m);
  ps(e, () => ({
    isClickOutSide: (z) => x.current ? !x.current.contains(z.target) : !1
  }));
  const k = te(v), E = G(null), N = G(void 0), R = N.current, [w, F] = W(i), U = G(null), H = ie(() => u != null ? u : zr(`${_o}-${Lt(4)}`), []), ce = ie(() => n !== void 0, [n]), le = T(ne), de = le.getUnit(H);
  pe(de == null ? void 0 : de.change$);
  const K = Ki(), J = Gr.transform.getPlainText((mn = (gn = de == null ? void 0 : de.getBody()) == null ? void 0 : gn.dataStream) != null ? mn : ""), ze = Ze(J), Ge = ie(() => uc(J), [J]), St = ie(() => K(Ge), [Ge, K]), { isSelecting: _e, isSelectingRef: Y } = ji({ unitId: r, subUnitId: s, editorId: H, isFocus: w, disableOnClick: f }), X = G(""), fe = T(Oe).getRenderById(H), ee = fe == null ? void 0 : fe.with(sr), ae = ee == null ? void 0 : ee.isFocusing, Ee = ie(() => le.getCurrentTypeOfUnit$(j.UNIVER_DOC), [le]), Le = pe(Ee), Ne = (Le == null ? void 0 : Le.getUnitId()) === H, Z = G([]), ge = _e, $e = (Sn = (pn = T(Zn).getConfig(mr)) == null ? void 0 : pn.functionScreenTips) != null ? Sn : !0;
  fs(() => {
    O(J);
  }, [J, O]);
  const Pe = Mr("="), Ue = Zi(r, s), oe = te((z, se = !0, Me, be) => {
    if (!N.current) return;
    X.current = z;
    const _t = z[0] === "=" ? z.slice(1) : "", we = K(_t), Dr = we.reduce((We, Et) => typeof Et == "object" ? `${We}${Et.token}` : `${We}${Et}`, ""), Rt = Pe(
      N.current,
      Dr === _t ? we : [],
      se,
      be
    );
    if (Z.current = Rt, Me) {
      const We = be != null ? be : R == null ? void 0 : R.getSelectionRanges();
      if ((We == null ? void 0 : We.length) !== 1)
        return;
      const Fr = We[0].startOffset - 1, Lr = hn(we, Fr, !1), vn = xr(we, Lr);
      if (vn >= 0) {
        const Cn = Rt.splice(vn, 1)[0];
        Cn && Rt.push(Cn);
      }
      Ue(w ? Rt : [], N.current);
    }
  });
  $(() => {
    w && oe(J, !1, !0);
  }, [w]), $(() => {
    if (w) {
      if (X.current === J) return;
      oe(J, !1, !0);
    }
  }, [J]), sc(w, d, J);
  const Ie = Bi(R), vt = Ji(w, r, s);
  $(() => {
    var z;
    k(_e, (z = ee == null ? void 0 : ee.isFocusing) != null ? z : !0);
  }, [k, _e]), Ho(w, p, R), Ot(() => {
    let z;
    if (U.current) {
      z = D.register({
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
      N.current = se, oe(o, !1, !0);
    }
    return () => {
      z == null || z.dispose();
    };
  }, []), Ot(() => {
    i ? (F(i), Ie()) : (C && (R == null || R.blur(), vt()), F(i));
  }, [i, R, Ie, vt, C]);
  const { checkScrollBar: Ct } = Bo(R, I, _);
  Qi(w, !!(_e && Ne), r, H, y), Xi(!!(w && ae && l), ge, R, S);
  const dn = te((z, se, Me) => {
    if (!ae)
      return;
    const be = se !== -1 ? [{ startOffset: se + 1, endOffset: se + 1, collapsed: !0 }] : void 0;
    oe(`=${z}`, !0, Me, be), Me && (Ie(), se !== -1 && setTimeout(() => {
      const _t = { startOffset: se + 1, endOffset: se + 1 }, we = R == null ? void 0 : R.render.with(qo);
      we == null || we.scrollToRange({ ..._t, collapsed: !0 });
    }, 50), Ct());
  });
  rc(
    w && !!(_e && Ne),
    w,
    Y,
    r,
    s,
    Z,
    c,
    !!ge,
    R,
    dn
  ), oc(w && !!(_e && Ne), r, c, F, g, () => {
    oe(ze.current, !1, !0);
  });
  const fn = (z) => {
    if (z) {
      const se = R == null ? void 0 : R.getSelectionRanges();
      if (se && se.length === 1) {
        const Me = se[0];
        if (Me.collapsed) {
          const be = z.offset;
          setTimeout(() => {
            R == null || R.setSelectionRanges([{ startOffset: Me.startOffset - be, endOffset: Me.endOffset - be }]);
          }, 30);
        }
      }
      Ie(), oe(`=${z.text}`);
    }
  }, Ar = () => {
    F(!0), a(), Ie();
  };
  return /* @__PURE__ */ V("div", { className: h, children: [
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
      lc,
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
function dc(t, e, n, o) {
  const r = T(Ce), s = Mr(""), i = pe(t == null ? void 0 : t.getDocumentDataModel().change$), [c, a] = W([]), g = T(Ro), m = G(""), d = T(ne);
  return $(() => {
    if (!t) return;
    const h = t.getDocumentDataModel().getPlainText();
    if (m.current === h)
      return;
    m.current = h;
    const u = r.sequenceNodesBuilder(h);
    a(u != null ? u : []);
  }, [i, t, r]), $(() => {
    var l, v;
    if (!t) return;
    if (!e) {
      const p = t.getDocumentData();
      t.setDocumentData({
        ...p,
        body: {
          ...p.body,
          dataStream: (v = (l = p.body) == null ? void 0 : l.dataStream) != null ? v : "",
          textRuns: []
        }
      });
      return;
    }
    const h = s(t, c, !1), u = new Be();
    return h.forEach((p) => {
      const S = st(p.token), C = d.getCurrentUnitForType(j.UNIVER_SHEET), _ = C == null ? void 0 : C.getActiveSheet();
      if (!S.sheetName && o !== (_ == null ? void 0 : _.getSheetId()) || S.sheetName && (_ == null ? void 0 : _.getName()) !== S.sheetName)
        return;
      const I = new tn(p.themeColor).toRgb(), f = g.addShape({
        range: S.range,
        style: {
          stroke: p.themeColor,
          fill: `rgba(${I.r}, ${I.g}, ${I.b}, 0.1)`,
          strokeDash: 12
        },
        primary: null
      });
      f && u.add(() => g.removeShape(f));
    }), () => {
      u.dispose();
    };
  }, [t, e, s, g, c]), { sequenceNodes: c };
}
function fc(t) {
  const e = T(on), { supportAcrossSheet: n = !1, keepSheetReference: o = !1, unitId: r, subUnitId: s, onChange: i } = t, a = T(ne).getUnit(r, j.UNIVER_SHEET), g = te(i), m = te((d, h) => {
    const u = a == null ? void 0 : a.getActiveSheet();
    if (!u || !n && u.getSheetId() !== s || !(d != null && d.length)) return;
    const l = o ? u.getName() : u.getSheetId() === s ? "" : u.getName(), v = d.map((p) => ({
      range: p.range,
      unitId: r,
      sheetName: l
    }));
    g(v, h);
  });
  $(() => {
    const d = new Be();
    return d.add(e.selectionMoveStart$.subscribe((h) => {
      m(h, !0);
    })), d.add(e.selectionMoving$.subscribe((h) => {
      m(h, !1);
    })), d.add(e.selectionMoveEnd$.subscribe((h) => {
      m(h, !1);
    })), () => {
      d.dispose();
    };
  }, [m, e.selectionMoveEnd$, e.selectionMoveStart$, e.selectionMoving$]);
}
const Bn = (t) => !t.some((n) => {
  if (typeof n == "string") {
    if (n !== qe.COMMA)
      return !0;
  } else if (n.nodeType !== Q.REFERENCE)
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
    onConfirm: a,
    onClose: g,
    onShowBySelection: m
  } = t, d = T(ot), h = T(Ce), [u, l] = W([]), [v, p] = W(0), S = G(null);
  $(() => {
    if (e && n.length) {
      const f = n.map((b) => b.sheetName ? ht(b.sheetName, b.range) : ve(b.range));
      l(f), p(f.length - 1);
    } else
      l([""]), p(0);
  }, [e]);
  const C = (f, b) => {
    const y = [...u];
    y[f] = b, l(y);
  }, _ = () => {
    l([...u, ""]), p(u.length);
  }, I = (f) => {
    u.splice(f, 1), l([...u]);
  };
  return fc({
    unitId: o,
    subUnitId: r,
    supportAcrossSheet: i,
    keepSheetReference: c,
    onChange: (f, b) => {
      if (!e && m != null && m(f))
        return;
      const y = new Set(u), A = f.map((O) => O.sheetName ? ht(O.sheetName, O.range) : ve(O.range)), D = A.filter((O) => !y.has(O));
      if (!D.length) return;
      const x = [...u];
      if (A.length > 1) {
        b || x.splice(v, 1), x.push(...D);
        const O = x.slice(0, s);
        l(O), p(O.length - 1), requestAnimationFrame(() => {
          var k;
          (k = S.current) == null || k.scrollTo({ top: S.current.scrollHeight });
        });
      } else {
        x.splice(v, 1, ...D);
        const O = x.slice(0, s);
        l(O), p(v + D.length - 1);
      }
    }
  }), /* @__PURE__ */ M(
    Rs,
    {
      width: "328px",
      open: e,
      title: d.t("rangeSelector.title"),
      draggable: !0,
      mask: !1,
      maskClosable: !1,
      footer: /* @__PURE__ */ V("footer", { className: "univer-flex univer-gap-2", children: [
        /* @__PURE__ */ M(et, { onClick: g, children: d.t("rangeSelector.cancel") }),
        /* @__PURE__ */ M(
          et,
          {
            variant: "primary",
            onClick: () => {
              a(
                u.filter((f) => {
                  const b = h.sequenceNodesBuilder(f);
                  return b && b.length === 1 && typeof b[0] != "string" && b[0].nodeType === Q.REFERENCE;
                }).map((f) => st(f)).map((f) => ({ ...f, range: gc(f.range) }))
              );
            },
            children: d.t("rangeSelector.confirm")
          }
        )
      ] }),
      onClose: g,
      children: /* @__PURE__ */ V(
        "div",
        {
          ref: S,
          className: re("-univer-mx-6 univer-max-h-60 univer-overflow-y-auto univer-px-6", nt),
          children: [
            u.map((f, b) => /* @__PURE__ */ V(
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
                      placeholder: d.t("rangeSelector.placeHolder"),
                      onFocus: () => p(b),
                      value: f,
                      onChange: (y) => C(b, y)
                    }
                  ),
                  u.length > 1 && /* @__PURE__ */ M(
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
            u.length < s && /* @__PURE__ */ M("div", { children: /* @__PURE__ */ V(et, { variant: "link", onClick: _, children: [
              /* @__PURE__ */ M(Ir, {}),
              /* @__PURE__ */ M("span", { children: d.t("rangeSelector.addAnotherRange") })
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
  return t.map((e) => e.sheetName ? ht(e.sheetName, e.range) : ve(e.range)).join(qe.COMMA);
}
function wr(t) {
  const [e, n] = W(null), {
    onVerify: o,
    selectorRef: r,
    unitId: s,
    subUnitId: i,
    maxRangeCount: c,
    supportAcrossSheet: a,
    keepSheetReference: g,
    autoFocus: m,
    onChange: d,
    onRangeSelectorDialogVisibleChange: h,
    onClickOutside: u,
    onFocusChange: l,
    forceShowDialogWhenSelectionChanged: v,
    hideEditor: p,
    resetRange: S
  } = t, [C, _] = W(m != null ? m : !1), [I, f] = W(!1), [b, y] = W([]), A = T(ot), D = T(je), { sequenceNodes: x } = dc(e, C, s, i), O = Ze(x), k = T(me), E = te(() => {
    e == null || e.setSelectionRanges([]), e == null || e.blur(), D.blur();
  }), N = te(() => {
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
      verify: () => Bn(O.current),
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
  }, [E, e, D, r, O]), $(() => {
    var R;
    o == null || o(Bn(x), (R = e == null ? void 0 : e.getDocumentDataModel().getPlainText()) != null ? R : "");
  }, [x]), $(() => {
    h == null || h(I);
  }, [I]), $(() => {
    if (I && S)
      return () => {
        const R = {
          unitId: s,
          subUnitId: i,
          selections: S
        };
        k.executeCommand(sn.id, R);
      };
  }, [I]), /* @__PURE__ */ V(ms, { children: [
    p ? null : /* @__PURE__ */ M(
      jo,
      {
        isSingle: !0,
        ...t,
        onFocusChange: (R, w) => {
          _(R), l == null || l(R, w);
        },
        editorRef: n,
        onClickOutside: () => {
          _(!1), E(), u == null || u();
        },
        icon: /* @__PURE__ */ M(_s, { title: A.t("rangeSelector.buttonTooltip"), placement: "bottom", children: /* @__PURE__ */ M(
          yr,
          {
            className: "univer-cursor-pointer dark:!univer-text-gray-300",
            onClick: N
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
          const w = Sc(R), F = Xr.newEmptyData();
          F.body.dataStream = w, e == null || e.replaceText(w, !1), d == null || d(F, w), f(!1), y([]), requestAnimationFrame(() => {
            E();
          });
        },
        onClose: () => {
          f(!1), y([]);
        },
        supportAcrossSheet: a,
        keepSheetReference: g,
        onShowBySelection: (R) => C || v ? (y(R), f(!0), !1) : !0
      }
    )
  ] });
}
const vc = () => {
  var o, r;
  const t = T(Tr), e = pe(t.currentSelector$), n = G(null);
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
        e == null || e.callback((c = i == null ? void 0 : i.split(",").map((a) => st(a))) != null ? c : []);
      }
    }
  );
};
var Cc = Object.defineProperty, _c = Object.getOwnPropertyDescriptor, Rc = (t, e, n) => e in t ? Cc(t, e, { enumerable: !0, configurable: !0, writable: !0, value: n }) : t[e] = n, Ec = (t, e, n, o) => {
  for (var r = o > 1 ? void 0 : o ? _c(e, n) : e, s = t.length - 1, i; s >= 0; s--)
    (i = t[s]) && (r = i(r) || r);
  return r;
}, yt = (t, e) => (n, o) => e(n, o, t), kr = (t, e, n) => Rc(t, typeof e != "symbol" ? e + "" : e, n);
let Ft = class extends Jr {
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
      [Nt]
    ]), this._initUIPart();
  }
  onReady() {
    [
      [De]
    ].forEach((t) => {
      this.disposeWithMe(this._renderManagerService.registerRenderModule(j.UNIVER_SHEET, t));
    });
  }
  onRendered() {
    [
      [Gt]
    ].forEach((t) => {
      this.disposeWithMe(this._renderManagerService.registerRenderModule(j.UNIVER_SHEET, t));
    }), no(this._injector, [
      [Dt],
      // FormulaProgressBar relies on TriggerCalculationController, but it is necessary to ensure that the formula calculation is done after rendered.
      [wt],
      [At]
    ]);
  }
  onSteady() {
    this._injector.get(Mt), this._injector.get(Nt);
  }
  _initUIPart() {
    const t = this._injector.get(lr);
    this.disposeWithMe(t.register(Eo, wr)), this.disposeWithMe(t.register(Io, hc)), this.disposeWithMe(this._uiPartsService.registerComponent(gs.GLOBAL, () => ar(vc, this._injector)));
  }
};
kr(Ft, "pluginName", dr);
kr(Ft, "type", j.UNIVER_SHEET);
Ft = Ec([
  Qr(Ao, cs),
  yt(1, B(ft)),
  yt(2, Oe),
  yt(3, Zn),
  yt(4, cr)
], Ft);
export {
  Ts as FORMULA_PROMPT_ACTIVATED,
  hc as FormulaEditor,
  Nt as FormulaReorderController,
  Tr as GlobalRangeSelectorService,
  xs as HelpFunctionOperation,
  Os as InsertFunctionOperation,
  Fs as MoreFunctionsOperation,
  wr as RangeSelector,
  De as RefSelectionsRenderService,
  gr as ReferenceAbsoluteOperation,
  Ls as SearchFunctionOperation,
  it as SelectEditorFormulaOperation,
  un as SheetOnlyPasteFormulaCommand,
  Ft as UniverSheetsFormulaUIPlugin
};
