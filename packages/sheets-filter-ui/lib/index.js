var ar = Object.defineProperty;
var cr = (t, e, r) => e in t ? ar(t, e, { enumerable: !0, configurable: !0, writable: !0, value: r }) : t[e] = r;
var v = (t, e, r) => cr(t, typeof e != "symbol" ? e + "" : e, r);
import { BooleanNumber as dt, createIdentifier as wt, Inject as T, LocaleService as G, IUniverInstanceService as De, ILogService as ur, extractPureTextFromCell as hr, numfmt as mt, Disposable as he, Injector as re, Quantity as Pt, Tools as be, ColorKit as ft, ICommandService as j, CommandType as Be, IContextService as We, ThemeService as $t, RxDisposable as Lt, fromCallback as dr, VerticalAlign as Re, InterceptorEffectEnum as mr, UniverInstanceType as ie, DependentOn as Mt, IConfigService as xt, Plugin as rt, merge as Ut, Optional as fr, registerDependencies as pr, touchDependencies as pt } from "@univerjs/core";
import { CustomFilterOperator as u, FilterBy as P, SetSheetsFilterCriteriaCommand as V, SheetsFilterService as z, SmartToggleSheetsFilterCommand as Ne, FILTER_MUTATIONS as vr, SetSheetsFilterRangeMutation as _r, SetSheetsFilterCriteriaMutation as gr, RemoveSheetsFilterMutation as Sr, ReCalcSheetsFilterMutation as Cr, UniverSheetsFilterPlugin as kt, SheetsFilterSyncController as Tr, ReCalcSheetsFilterCommand as it, ClearSheetsFilterCriteriaCommand as st, RemoveSheetFilterCommand as Er, SetSheetFilterRangeCommand as Fr } from "@univerjs/sheets-filter";
import { IEditorBridgeService as vt, SetCellEditVisibleOperation as Nr, SheetSkeletonManagerService as yr, ISheetSelectionRenderService as Ir, SelectionControl as Or, attachSelectionWithCoord as br, getCoordByCell as Rr, SheetsRenderService as Ht, SheetsUIPart as Ar, getObservableWithExclusiveRange$ as wr, getCurrentRangeDisable$ as Pr, whenSheetEditorFocused as $r, SheetCanvasPopManagerService as Lr } from "@univerjs/sheets-ui";
import { ILayoutService as Mr, useDependency as D, useObservable as M, IMessageService as Dt, useComponentsOfPart as xr, ComponentContainer as Ur, getMenuHiddenObservable as nt, MenuItemType as ot, RibbonStartGroup as kr, KeyCode as Hr, MetaKeys as _t, ComponentManager as Dr, IShortcutService as Br, IMenuManagerService as Wr } from "@univerjs/ui";
import { COLOR_BLACK_RGB as gt, Rect as St, Shape as Vr, IRenderManagerService as Bt } from "@univerjs/engine-render";
import { RefRangeService as Qr, SheetPermissionCheckController as Gr, SheetsSelectionsService as jr, getSheetCommandTarget as Yr, WorksheetFilterPermission as ve, WorksheetViewPermission as _e, RangeProtectionPermissionViewPoint as ge, expandToContinuousRange as Zr, SheetInterceptorService as Kr, SetRangeValuesMutation as qr, INTERCEPTOR_POINT as Xr } from "@univerjs/sheets";
import { BehaviorSubject as Q, ReplaySubject as zr, Subject as Jr, merge as ei, combineLatest as ti, throttleTime as Wt, startWith as Vt, map as ae, shareReplay as ri, of as de, switchMap as lt, filter as ii, takeUntil as si, distinctUntilChanged as ni } from "rxjs";
import { IRPCChannelService as Qt, toModule as oi, fromModule as li } from "@univerjs/rpc";
import { clsx as Ae, borderClassName as at, Select as Ct, RadioGroup as ai, Radio as Tt, Input as Gt, Checkbox as ci, Tree as ui, Tooltip as hi, Switch as di, MessageType as jt, Segmented as mi, Button as Ye } from "@univerjs/design";
import { useRef as fi, createElement as me, forwardRef as ye, useCallback as k, useMemo as ct } from "react";
import { jsx as d, jsxs as N, Fragment as Et } from "react/jsx-runtime";
var I = /* @__PURE__ */ ((t) => (t[t.FIRST = 0] = "FIRST", t[t.SECOND = 1] = "SECOND", t))(I || {}), g = /* @__PURE__ */ ((t) => (t.NONE = "none", t.STARTS_WITH = "startsWith", t.DOES_NOT_START_WITH = "doesNotStartWith", t.ENDS_WITH = "endsWith", t.DOES_NOT_END_WITH = "doesNotEndWith", t.CONTAINS = "contains", t.DOES_NOT_CONTAIN = "doesNotContain", t.EQUALS = "equals", t.NOT_EQUALS = "notEquals", t.EMPTY = "empty", t.NOT_EMPTY = "notEmpty", t.BETWEEN = "between", t.NOT_BETWEEN = "notBetween", t.CUSTOM = "custom", t))(g || {}), f;
((t) => {
  t.NONE = {
    label: "sheets-filter.conditions.none",
    operator: g.NONE,
    order: I.SECOND,
    numOfParameters: 0,
    getDefaultFormParams: () => {
      throw new Error("[FilterConditionItems.NONE]: should not have initial form params!");
    },
    testMappingParams: (i) => i.operator1 === g.NONE,
    mapToFilterColumn: () => null,
    testMappingFilterColumn: (i) => !i.customFilters && !i.filters ? {} : !1
  }, t.EMPTY = {
    label: "sheets-filter.conditions.empty",
    operator: g.EMPTY,
    order: I.SECOND,
    numOfParameters: 0,
    getDefaultFormParams: () => {
      throw new Error("[FilterConditionItems.EMPTY]: should not have initial form params!");
    },
    testMappingParams: ({ operator1: i }) => i === g.EMPTY,
    mapToFilterColumn: () => ({ customFilters: { customFilters: [{ val: "" }] } }),
    testMappingFilterColumn: (i) => {
      var c;
      if (((c = i.customFilters) == null ? void 0 : c.customFilters.length) !== 1)
        return !1;
      const o = i.customFilters.customFilters[0];
      return o.val === "" && o.operator === void 0 ? { operator1: g.EMPTY } : !1;
    }
  }, t.NOT_EMPTY = {
    label: "sheets-filter.conditions.not-empty",
    operator: g.NOT_EMPTY,
    order: I.SECOND,
    numOfParameters: 0,
    getDefaultFormParams: () => {
      throw new Error("[FilterConditionItems.NOT_EMPTY]: should not have initial form params!");
    },
    testMappingParams: ({ operator1: i }) => i === g.NOT_EMPTY,
    mapToFilterColumn: () => ({ customFilters: { customFilters: [{ val: "", operator: u.NOT_EQUALS }] } }),
    testMappingFilterColumn: (i) => {
      var c;
      if (((c = i.customFilters) == null ? void 0 : c.customFilters.length) !== 1)
        return !1;
      const o = i.customFilters.customFilters[0];
      return o.val === " " && o.operator === u.NOT_EQUALS ? { operator1: g.NOT_EMPTY } : !1;
    }
  }, t.TEXT_CONTAINS = {
    label: "sheets-filter.conditions.text-contains",
    operator: g.CONTAINS,
    order: I.FIRST,
    numOfParameters: 1,
    getDefaultFormParams: () => ({ operator1: g.CONTAINS, val1: "" }),
    testMappingParams: (i) => {
      const [o] = U(i);
      return o === g.CONTAINS;
    },
    mapToFilterColumn: (i) => {
      const { val1: o } = i;
      return o === "" ? null : {
        customFilters: { customFilters: [{ val: `*${o}*` }] }
      };
    },
    testMappingFilterColumn: (i) => {
      var c;
      if (((c = i.customFilters) == null ? void 0 : c.customFilters.length) !== 1)
        return !1;
      const o = i.customFilters.customFilters[0], l = o.val.toString();
      return !o.operator && l.startsWith("*") && l.endsWith("*") ? { operator1: g.CONTAINS, val1: l.slice(1, -1) } : !1;
    }
  }, t.DOES_NOT_CONTAIN = {
    label: "sheets-filter.conditions.does-not-contain",
    operator: g.DOES_NOT_CONTAIN,
    order: I.FIRST,
    numOfParameters: 1,
    getDefaultFormParams: () => ({ operator1: g.DOES_NOT_CONTAIN, val1: "" }),
    mapToFilterColumn: (i) => ({
      customFilters: { customFilters: [{ val: `*${i.val1}*`, operator: u.NOT_EQUALS }] }
    }),
    testMappingParams: (i) => {
      const [o] = U(i);
      return o === g.DOES_NOT_CONTAIN;
    },
    testMappingFilterColumn: (i) => {
      var c;
      if (((c = i.customFilters) == null ? void 0 : c.customFilters.length) !== 1)
        return !1;
      const o = i.customFilters.customFilters[0], l = o.val.toString();
      return o.operator === u.NOT_EQUALS && l.startsWith("*") && l.endsWith("*") ? { operator1: g.DOES_NOT_CONTAIN, val1: l.slice(1, -1) } : !1;
    }
  }, t.STARTS_WITH = {
    label: "sheets-filter.conditions.starts-with",
    operator: g.STARTS_WITH,
    order: I.FIRST,
    numOfParameters: 1,
    getDefaultFormParams: () => ({ operator1: g.STARTS_WITH, val1: "" }),
    mapToFilterColumn: (i) => ({
      customFilters: { customFilters: [{ val: `${i.val1}*` }] }
    }),
    testMappingParams: (i) => {
      const [o] = U(i);
      return o === g.STARTS_WITH;
    },
    testMappingFilterColumn: (i) => {
      var c;
      if (((c = i.customFilters) == null ? void 0 : c.customFilters.length) !== 1)
        return !1;
      const o = i.customFilters.customFilters[0], l = o.val.toString();
      return !o.operator && l.endsWith("*") && !l.startsWith("*") ? { operator1: g.STARTS_WITH, val1: l.slice(0, -1) } : !1;
    }
  }, t.ENDS_WITH = {
    label: "sheets-filter.conditions.ends-with",
    operator: g.ENDS_WITH,
    order: I.FIRST,
    numOfParameters: 1,
    getDefaultFormParams: () => ({ operator1: g.ENDS_WITH, val1: "" }),
    mapToFilterColumn: (i) => ({
      customFilters: { customFilters: [{ val: `*${i.val1}` }] }
    }),
    testMappingParams: (i) => {
      const [o] = U(i);
      return o === g.ENDS_WITH;
    },
    testMappingFilterColumn: (i) => {
      var c;
      if (((c = i.customFilters) == null ? void 0 : c.customFilters.length) !== 1)
        return !1;
      const o = i.customFilters.customFilters[0], l = o.val.toString();
      return !o.operator && l.startsWith("*") && !l.endsWith("*") ? { operator1: g.ENDS_WITH, val1: l.slice(1) } : !1;
    }
  }, t.EQUALS = {
    label: "sheets-filter.conditions.equals",
    operator: g.EQUALS,
    order: I.FIRST,
    numOfParameters: 1,
    getDefaultFormParams: () => ({ operator1: g.EQUALS, val1: "" }),
    testMappingParams: (i) => {
      const [o] = U(i);
      return o === g.EQUALS;
    },
    mapToFilterColumn: (i) => {
      const { val1: o } = i;
      return o === "" ? null : {
        customFilters: { customFilters: [{ val: o }] }
      };
    },
    testMappingFilterColumn: (i) => {
      var o, l, c;
      return ((l = (o = i.filters) == null ? void 0 : o.filters) == null ? void 0 : l.length) === 1 ? { operator1: g.EQUALS, val1: "" } : ((c = i.customFilters) == null ? void 0 : c.customFilters.length) === 1 && !i.customFilters.customFilters[0].operator ? { operator1: g.EQUALS, val1: i.customFilters.customFilters[0].val.toString() } : !1;
    }
  }, t.GREATER_THAN = {
    label: "sheets-filter.conditions.greater-than",
    operator: u.GREATER_THAN,
    numOfParameters: 1,
    order: I.FIRST,
    getDefaultFormParams: () => ({ operator1: u.GREATER_THAN, val1: "" }),
    mapToFilterColumn: (i) => ({
      customFilters: { customFilters: [{ val: i.val1, operator: u.GREATER_THAN }] }
    }),
    testMappingParams: (i) => {
      const [o] = U(i);
      return o === u.GREATER_THAN;
    },
    testMappingFilterColumn: (i) => {
      var l;
      if (((l = i.customFilters) == null ? void 0 : l.customFilters.length) !== 1)
        return !1;
      const o = i.customFilters.customFilters[0];
      return o.operator !== u.GREATER_THAN ? !1 : { operator1: u.GREATER_THAN, val1: o.val.toString() };
    }
  }, t.GREATER_THAN_OR_EQUAL = {
    label: "sheets-filter.conditions.greater-than-or-equal",
    operator: u.GREATER_THAN_OR_EQUAL,
    numOfParameters: 1,
    order: I.FIRST,
    getDefaultFormParams: () => ({ operator1: u.GREATER_THAN_OR_EQUAL, val1: "" }),
    testMappingParams: (i) => {
      const [o] = U(i);
      return o === u.GREATER_THAN_OR_EQUAL;
    },
    mapToFilterColumn: (i) => ({
      customFilters: { customFilters: [{ val: i.val1, operator: u.GREATER_THAN_OR_EQUAL }] }
    }),
    testMappingFilterColumn: (i) => {
      var l;
      if (((l = i.customFilters) == null ? void 0 : l.customFilters.length) !== 1)
        return !1;
      const o = i.customFilters.customFilters[0];
      return o.operator !== u.GREATER_THAN_OR_EQUAL ? !1 : { operator1: u.GREATER_THAN_OR_EQUAL, val1: o.val.toString() };
    }
  }, t.LESS_THAN = {
    label: "sheets-filter.conditions.less-than",
    operator: u.LESS_THAN,
    numOfParameters: 1,
    order: I.FIRST,
    getDefaultFormParams: () => ({ operator1: u.LESS_THAN, val1: "" }),
    testMappingParams: (i) => {
      const [o] = U(i);
      return o === u.LESS_THAN;
    },
    mapToFilterColumn: (i) => ({
      customFilters: { customFilters: [{ val: i.val1, operator: u.LESS_THAN }] }
    }),
    testMappingFilterColumn: (i) => {
      var l;
      if (((l = i.customFilters) == null ? void 0 : l.customFilters.length) !== 1)
        return !1;
      const o = i.customFilters.customFilters[0];
      return o.operator !== u.LESS_THAN ? !1 : { operator1: u.LESS_THAN, val1: o.val.toString() };
    }
  }, t.LESS_THAN_OR_EQUAL = {
    label: "sheets-filter.conditions.less-than-or-equal",
    operator: u.LESS_THAN_OR_EQUAL,
    numOfParameters: 1,
    order: I.FIRST,
    getDefaultFormParams: () => ({ operator1: u.LESS_THAN_OR_EQUAL, val1: "" }),
    testMappingParams: (i) => {
      const [o] = U(i);
      return o === u.LESS_THAN_OR_EQUAL;
    },
    mapToFilterColumn: (i) => ({
      customFilters: { customFilters: [{ val: i.val1, operator: u.LESS_THAN_OR_EQUAL }] }
    }),
    testMappingFilterColumn: (i) => {
      var l;
      if (((l = i.customFilters) == null ? void 0 : l.customFilters.length) !== 1)
        return !1;
      const o = i.customFilters.customFilters[0];
      return o.operator !== u.LESS_THAN_OR_EQUAL ? !1 : { operator1: u.LESS_THAN_OR_EQUAL, val1: o.val.toString() };
    }
  }, t.EQUAL = {
    label: "sheets-filter.conditions.equal",
    operator: u.EQUAL,
    numOfParameters: 1,
    order: I.FIRST,
    getDefaultFormParams: () => ({ operator1: u.EQUAL, val1: "" }),
    testMappingParams: (i) => {
      const [o] = U(i);
      return o === u.EQUAL;
    },
    mapToFilterColumn: (i) => ({
      customFilters: { customFilters: [{ val: i.val1, operator: u.EQUAL }] }
    }),
    testMappingFilterColumn: (i) => {
      var l;
      if (((l = i.customFilters) == null ? void 0 : l.customFilters.length) !== 1)
        return !1;
      const o = i.customFilters.customFilters[0];
      return o.operator !== u.EQUAL ? !1 : { operator1: u.EQUAL, val1: o.val.toString() };
    }
  }, t.NOT_EQUAL = {
    label: "sheets-filter.conditions.not-equal",
    operator: u.NOT_EQUALS,
    numOfParameters: 1,
    order: I.FIRST,
    getDefaultFormParams: () => ({ operator1: u.NOT_EQUALS, val1: "" }),
    testMappingParams: (i) => {
      const [o] = U(i);
      return o === u.NOT_EQUALS;
    },
    mapToFilterColumn: (i) => ({
      customFilters: { customFilters: [{ val: i.val1, operator: u.NOT_EQUALS }] }
    }),
    testMappingFilterColumn: (i) => {
      var l;
      if (((l = i.customFilters) == null ? void 0 : l.customFilters.length) !== 1)
        return !1;
      const o = i.customFilters.customFilters[0];
      return o.operator !== u.NOT_EQUALS ? !1 : { operator1: u.NOT_EQUALS, val1: o.val.toString() };
    }
  }, t.BETWEEN = {
    label: "sheets-filter.conditions.between",
    operator: g.BETWEEN,
    order: I.SECOND,
    numOfParameters: 2,
    getDefaultFormParams: () => ({
      and: !0,
      operator1: u.GREATER_THAN_OR_EQUAL,
      val1: "",
      operator2: u.LESS_THAN_OR_EQUAL,
      val2: ""
    }),
    testMappingParams: (i) => {
      const { and: o, operator1: l, operator2: c } = i;
      if (!o) return !1;
      const h = [l, c];
      return h.includes(u.GREATER_THAN_OR_EQUAL) && h.includes(u.LESS_THAN_OR_EQUAL);
    },
    mapToFilterColumn: (i) => {
      const { val1: o, val2: l, operator1: c } = i, h = c === u.GREATER_THAN_OR_EQUAL;
      return {
        customFilters: {
          and: dt.TRUE,
          customFilters: [
            { val: h ? o : l, operator: u.GREATER_THAN_OR_EQUAL },
            { val: h ? l : o, operator: u.LESS_THAN_OR_EQUAL }
          ]
        }
      };
    },
    testMappingFilterColumn: (i) => {
      var c;
      if (((c = i.customFilters) == null ? void 0 : c.customFilters.length) !== 2)
        return !1;
      const [o, l] = i.customFilters.customFilters;
      return o.operator === u.GREATER_THAN_OR_EQUAL && l.operator === u.LESS_THAN_OR_EQUAL && i.customFilters.and ? {
        and: !0,
        operator1: u.GREATER_THAN_OR_EQUAL,
        val1: o.val.toString(),
        operator2: u.LESS_THAN_OR_EQUAL,
        val2: l.val.toString()
      } : l.operator === u.GREATER_THAN_OR_EQUAL && o.operator === u.LESS_THAN_OR_EQUAL && i.customFilters.and ? {
        and: !0,
        operator1: u.GREATER_THAN_OR_EQUAL,
        val1: l.val.toString(),
        operator2: u.LESS_THAN_OR_EQUAL,
        val2: o.val.toLocaleString()
      } : !1;
    }
  }, t.NOT_BETWEEN = {
    label: "sheets-filter.conditions.not-between",
    operator: g.NOT_BETWEEN,
    order: I.SECOND,
    numOfParameters: 2,
    getDefaultFormParams: () => ({
      operator1: u.LESS_THAN,
      val1: "",
      operator2: u.GREATER_THAN,
      val2: ""
    }),
    testMappingParams: (i) => {
      const { and: o, operator1: l, operator2: c } = i;
      if (o) return !1;
      const h = [l, c];
      return h.includes(u.GREATER_THAN) && h.includes(u.LESS_THAN);
    },
    mapToFilterColumn: (i) => {
      const { val1: o, val2: l, operator1: c } = i, h = c === u.GREATER_THAN;
      return {
        customFilters: {
          customFilters: [
            { val: h ? o : l, operator: u.GREATER_THAN },
            { val: h ? l : o, operator: u.LESS_THAN }
          ]
        }
      };
    },
    testMappingFilterColumn: (i) => {
      var c;
      if (((c = i.customFilters) == null ? void 0 : c.customFilters.length) !== 2)
        return !1;
      const [o, l] = i.customFilters.customFilters;
      return o.operator === u.LESS_THAN && l.operator === u.GREATER_THAN && !i.customFilters.and ? {
        operator1: u.LESS_THAN,
        val1: o.val.toString(),
        operator2: u.GREATER_THAN,
        val2: l.val.toString()
      } : l.operator === u.LESS_THAN && o.operator === u.GREATER_THAN && !i.customFilters.and ? {
        operator1: u.GREATER_THAN,
        val1: l.val.toString(),
        operator2: u.LESS_THAN,
        val2: o.val.toLocaleString()
      } : !1;
    }
  }, t.CUSTOM = {
    label: "sheets-filter.conditions.custom",
    operator: g.CUSTOM,
    order: I.SECOND,
    numOfParameters: 2,
    getDefaultFormParams: () => ({
      operator1: g.NONE,
      val1: "",
      operator2: g.NONE,
      val2: ""
    }),
    testMappingParams: () => !0,
    mapToFilterColumn: (i) => {
      const { and: o, val1: l, val2: c, operator1: h, operator2: m } = i;
      function C(b, y) {
        for (const R of t.ALL_CONDITIONS)
          if (R.operator === b)
            return R.mapToFilterColumn({ val1: y, operator1: b });
      }
      const p = !h || h === t.NONE.operator, _ = !m || m === t.NONE.operator;
      if (p && _)
        return t.NONE.mapToFilterColumn({});
      if (p)
        return C(m, c);
      if (_)
        return C(h, l);
      const S = C(h, l), E = C(m, c), F = {
        customFilters: [
          S.customFilters.customFilters[0],
          E.customFilters.customFilters[0]
        ]
      };
      return o && (F.and = dt.TRUE), { customFilters: F };
    },
    testMappingFilterColumn: (i) => {
      var c;
      if (((c = i.customFilters) == null ? void 0 : c.customFilters.length) !== 2)
        return !1;
      const o = i.customFilters.customFilters.map((h) => a({ customFilters: { customFilters: [h] } })), l = {
        operator1: o[0][0].operator,
        val1: o[0][1].val1,
        operator2: o[1][0].operator,
        val2: o[1][1].val1
      };
      return i.customFilters.and && (l.and = !0), l;
    }
  }, t.ALL_CONDITIONS = [
    // ------------------------------
    t.NONE,
    // ------------------------------
    t.EMPTY,
    t.NOT_EMPTY,
    // ------------------------------
    t.TEXT_CONTAINS,
    t.DOES_NOT_CONTAIN,
    t.STARTS_WITH,
    t.ENDS_WITH,
    t.EQUALS,
    // ------------------------------
    t.GREATER_THAN,
    t.GREATER_THAN_OR_EQUAL,
    t.LESS_THAN,
    t.LESS_THAN_OR_EQUAL,
    t.EQUAL,
    t.NOT_EQUAL,
    t.BETWEEN,
    t.NOT_BETWEEN,
    // ------------------------------
    t.CUSTOM
  ];
  function e(i) {
    const o = t.ALL_CONDITIONS.find((l) => l.operator === i);
    if (!o)
      throw new Error(`[SheetsFilter]: no condition item found for operator: ${i}`);
    return o;
  }
  t.getItemByOperator = e;
  function r(i, o) {
    for (const l of t.ALL_CONDITIONS.filter((c) => c.numOfParameters === o))
      if (l.numOfParameters !== 0 && l.testMappingParams(i))
        return l;
    for (const l of t.ALL_CONDITIONS)
      if (l.testMappingParams(i))
        return l;
    throw new Error("[SheetsFilter]: no condition item can be mapped from the filter map params!");
  }
  t.testMappingParams = r;
  function s(i) {
    const o = t.ALL_CONDITIONS.find((l) => l.operator === i);
    return (o == null ? void 0 : o.numOfParameters) === 0 ? { operator1: o.operator } : o.getDefaultFormParams();
  }
  t.getInitialFormParams = s;
  function n(i, o) {
    return i.mapToFilterColumn(o);
  }
  t.mapToFilterColumn = n;
  function a(i) {
    if (!i)
      return [t.NONE, {}];
    for (const o of t.ALL_CONDITIONS) {
      const l = o.testMappingFilterColumn(i);
      if (l)
        return [o, l];
    }
    return [t.NONE, {}];
  }
  t.testMappingFilterColumn = a;
})(f || (f = {}));
function U(t) {
  const { operator1: e, operator2: r, val1: s, val2: n } = t;
  if (e && r)
    throw new Error("Both operator1 and operator2 are set!");
  if (!e && !r)
    throw new Error("Neither operator1 and operator2 and both not set!");
  return e ? [e, s] : [r, n];
}
function Xe(t) {
  const e = [], r = [];
  let s = 0, n = 0;
  function a(i) {
    i.leaf && (i.checked ? (e.push(i), s += i.count) : (r.push(i), n += i.count)), i.children && i.children.forEach(a);
  }
  return t.forEach(a), {
    checkedItems: e,
    uncheckedItems: r,
    checked: s,
    unchecked: n
  };
}
var pi = Object.getOwnPropertyDescriptor, vi = (t, e, r, s) => {
  for (var n = s > 1 ? void 0 : s ? pi(e, r) : e, a = t.length - 1, i; a >= 0; a--)
    (i = t[a]) && (n = i(n) || n);
  return n;
}, Ze = (t, e) => (r, s) => e(r, s, t);
const ut = "sheets-filter.generate-filter-values.service", we = wt(ut), _i = ["yyyy-mm-dd", "yyyy-mm-dd;@", "yyyy/mm/dd;@", "yyyy/mm/dd hh:mm", "yyyy-m-d am/pm h:mm", "yyyy-MM-dd", "yyyy/MM/dd", "yyyy/mm/dd", 'yyyy"年"MM"月"dd"日"', "MM-dd", 'M"月"d"日"', "MM-dd A/P hh:mm"];
let ze = class extends he {
  constructor(t, e, r) {
    super(), this._localeService = t, this._univerInstanceService = e, this._logService = r;
  }
  async getFilterValues(t) {
    var m;
    const { unitId: e, subUnitId: r, filteredOutRowsByOtherColumns: s, filterColumn: n, filters: a, blankChecked: i, iterateRange: o, alreadyChecked: l } = t, c = this._univerInstanceService.getUnit(e), h = (m = this._univerInstanceService.getUnit(e)) == null ? void 0 : m.getSheetBySheetId(r);
    return !c || !h ? [] : (this._logService.debug("[SheetsGenerateFilterValuesService]", "getFilterValues for", { unitId: e, subUnitId: r }), Yt(
      a,
      this._localeService,
      o,
      h,
      new Set(s),
      n,
      new Set(l.map(String)),
      i,
      c.getStyles()
    ));
  }
};
ze = vi([
  Ze(0, T(G)),
  Ze(1, De),
  Ze(2, ur)
], ze);
function Yt(t, e, r, s, n, a, i, o, l) {
  var F, b, y, R, $, Y, Z, se, x, A;
  const c = /* @__PURE__ */ new Map(), h = /* @__PURE__ */ new Map(), m = "yyyy-mm-dd", C = new Set(_i), p = "empty", _ = !t && ((a == null ? void 0 : a.filterBy) === P.COLORS || (a == null ? void 0 : a.filterBy) === P.CONDITIONS) && ((F = a.filteredOutRows) == null ? void 0 : F.size);
  let S = 0;
  for (const w of s.iterateByColumn(r, !1, !1)) {
    const { row: Qe, rowSpan: ht = 1 } = w;
    let ne = 0;
    for (; ne < ht; ) {
      const or = Qe + ne;
      if (n.has(or)) {
        ne++;
        continue;
      }
      const J = w != null && w.value ? hr(w.value) : "";
      if (!J) {
        S += 1, ne += ht;
        continue;
      }
      const Ie = (b = w.value) != null && b.v && !w.value.p ? ($ = (R = l.get((y = w.value) == null ? void 0 : y.s)) == null ? void 0 : R.n) == null ? void 0 : $.pattern : "", lr = Ie && mt.getFormatInfo(Ie).isDate;
      if (Ie && lr && C.has(Ie)) {
        const K = (Y = s.getCellRaw(w.row, w.col)) == null ? void 0 : Y.v;
        if (!K) {
          ne++;
          continue;
        }
        const oe = mt.format(m, K), [O, H, pe] = oe.split("-").map(Number);
        let ee = c.get(`${O}`);
        ee || (ee = {
          title: `${O}`,
          key: `${O}`,
          children: [],
          count: 0,
          leaf: !1,
          checked: !1
        }, c.set(`${O}`, ee), h.set(`${O}`, [`${O}`]));
        let W = (Z = ee.children) == null ? void 0 : Z.find((je) => je.key === `${O}-${H}`);
        W || (W = {
          title: e.t(`sheets-filter.date.${H}`),
          key: `${O}-${H}`,
          children: [],
          count: 0,
          leaf: !1,
          checked: !1
        }, (se = ee.children) == null || se.push(W), h.set(`${O}-${H}`, [`${O}`, `${O}-${H}`]));
        const Ge = (x = W == null ? void 0 : W.children) == null ? void 0 : x.find((je) => je.key === `${O}-${H}-${pe}`);
        Ge ? (Ge.originValues.add(J), Ge.count++, W.count++, ee.count++) : ((A = W.children) == null || A.push({
          title: `${pe}`,
          key: `${O}-${H}-${pe}`,
          count: 1,
          originValues: /* @__PURE__ */ new Set([J]),
          leaf: !0,
          checked: _ ? !1 : i.size ? i.has(J) : !o
        }), W.count++, ee.count++, h.set(`${O}-${H}-${pe}`, [`${O}`, `${O}-${H}`, `${O}-${H}-${pe}`]));
      } else {
        const K = J;
        let oe = c.get(K);
        oe ? oe.count++ : (oe = {
          title: J,
          leaf: !0,
          checked: _ ? !1 : i.size ? i.has(J) : !o,
          key: K,
          count: 1
        }, c.set(K, oe), h.set(K, [K]));
      }
      ne++;
    }
  }
  const E = _ ? !1 : t ? o : !0;
  if (S > 0) {
    const w = {
      title: e.t("sheets-filter.panel.empty"),
      count: S,
      leaf: !0,
      checked: E,
      key: p
    };
    c.set("empty", w), h.set("empty", [p]);
  }
  return {
    filterTreeItems: gi(Array.from(c.values())),
    filterTreeMapCache: h
  };
}
function gi(t) {
  return Array.from(t).sort((e, r) => e.children && !r.children ? -1 : !e.children && r.children ? 1 : Si(e.title, r.title)).map((e) => (e.children && e.children.sort((r, s) => {
    const n = Number.parseInt(r.key.split("-")[1], 10), a = Number.parseInt(s.key.split("-")[1], 10);
    return n - a;
  }).forEach((r) => {
    r.children && r.children.sort((s, n) => {
      const a = Number.parseInt(s.key.split("-")[2], 10), i = Number.parseInt(n.key.split("-")[2], 10);
      return a - i;
    });
  }), e));
}
const Ft = (t) => !Number.isNaN(Number(t)) && !Number.isNaN(Number.parseFloat(t));
function Si(t, e) {
  const r = Ft(t), s = Ft(e);
  return r && s ? Number.parseFloat(t) - Number.parseFloat(e) : r && !s ? -1 : !r && s ? 1 : t.localeCompare(e);
}
function Je(t, e) {
  for (const r of t) {
    if (r.key === e)
      return r;
    if (r.children) {
      const s = Je(r.children, e);
      if (s)
        return s;
    }
  }
  return null;
}
function Zt(t) {
  return t.leaf ? t.checked : t.children ? t.children.every((e) => Zt(e)) : !0;
}
function Se(t, e) {
  t.leaf && (e !== void 0 ? t.checked = e : t.checked = !t.checked), t.children && t.children.forEach((r) => Se(r, e));
}
function Kt(t, e) {
  const r = [];
  return t.forEach((s) => {
    const n = s.originValues ? e.some(
      (o) => Array.from(s.originValues).some(
        (l) => l.toLowerCase().includes(o.toLowerCase())
      )
    ) : !1, a = !n && e.some(
      (o) => s.title.toLowerCase().includes(o.toLowerCase())
    );
    if (n || a)
      r.push({ ...s });
    else if (s.children) {
      const o = Kt(s.children, e);
      if (o.length > 0) {
        const l = o.reduce((c, h) => c + h.count, 0);
        r.push({ ...s, count: l, children: o });
      }
    }
  }), r;
}
var Ci = Object.getOwnPropertyDescriptor, Ve = (t, e, r, s) => {
  for (var n = s > 1 ? void 0 : s ? Ci(e, r) : e, a = t.length - 1, i; a >= 0; a--)
    (i = t[a]) && (n = i(n) || n);
  return n;
}, Ee = (t, e) => (r, s) => e(r, s, t);
wt("sheets-filter-ui.sheets-filter-panel.service");
let X = class extends he {
  constructor(e, r) {
    super();
    v(this, "_filterBy$", new Q(P.VALUES));
    v(this, "filterBy$", this._filterBy$.asObservable());
    v(this, "_filterByModel$", new zr(1));
    v(this, "filterByModel$", this._filterByModel$.asObservable());
    v(this, "_filterByModel", null);
    v(this, "_hasCriteria$", new Q(!1));
    v(this, "hasCriteria$", this._hasCriteria$.asObservable());
    v(this, "_filterModel", null);
    v(this, "_col$", new Q(-1));
    v(this, "col$", this._col$.asObservable());
    v(this, "_filterHeaderListener", null);
    this._injector = e, this._refRangeService = r;
  }
  get filterBy() {
    return this._filterBy$.getValue();
  }
  get filterByModel() {
    return this._filterByModel;
  }
  set filterByModel(e) {
    this._filterByModel = e, this._filterByModel$.next(e);
  }
  get filterModel() {
    return this._filterModel;
  }
  get col() {
    return this._col$.getValue();
  }
  dispose() {
    this._filterBy$.complete(), this._filterByModel$.complete(), this._hasCriteria$.complete();
  }
  setupCol(e, r) {
    this.terminate(), this._filterModel = e, this._col$.next(r);
    const s = e.getFilterColumn(r);
    if (s) {
      const n = s.getColumnData();
      if (n.customFilters) {
        this._hasCriteria$.next(!0), this._setupByConditions(e, r);
        return;
      }
      if (n.colorFilters) {
        this._hasCriteria$.next(!0), this._setupByColors(e, r);
        return;
      }
      if (n.filters) {
        this._hasCriteria$.next(!0), this._setupByValues(e, r);
        return;
      }
      this._hasCriteria$.next(!1), this._setupByValues(e, r);
      return;
    }
    this._hasCriteria$.next(!1), this._setupByValues(e, r);
  }
  changeFilterBy(e) {
    if (!this._filterModel || this.col === -1)
      return !1;
    switch (e) {
      case P.VALUES:
        this._setupByValues(this._filterModel, this.col);
        break;
      case P.COLORS:
        this._setupByColors(this._filterModel, this.col);
        break;
      case P.CONDITIONS:
        this._setupByConditions(this._filterModel, this.col);
        break;
    }
    return !0;
  }
  terminate() {
    return this._filterModel = null, this._col$.next(-1), this._disposeFilterHeaderChangeListener(), !0;
  }
  _disposeFilterHeaderChangeListener() {
    var e;
    (e = this._filterHeaderListener) == null || e.dispose(), this._filterHeaderListener = null;
  }
  _listenToFilterHeaderChange(e, r) {
    this._disposeFilterHeaderChangeListener();
    const s = e.unitId, n = e.subUnitId, a = e.getRange(), i = {
      startColumn: r,
      startRow: a.startRow,
      endRow: a.startRow,
      endColumn: r
    };
    this._filterHeaderListener = this._refRangeService.watchRange(s, n, i, (o, l) => {
      if (!l)
        this.terminate();
      else {
        const c = l.startColumn - o.startColumn;
        c !== 0 && this._filterByModel.deltaCol(c);
      }
    });
  }
  async _setupByValues(e, r) {
    this._disposePreviousModel();
    const s = e.getRange();
    if (s.startRow === s.endRow) return !1;
    const n = await $e.fromFilterColumn(
      this._injector,
      e,
      r
    );
    return this.filterByModel = n, this._filterBy$.next(P.VALUES), this._listenToFilterHeaderChange(e, r), !0;
  }
  async _setupByColors(e, r) {
    this._disposePreviousModel();
    const s = e.getRange();
    if (s.startRow === s.endRow) return !1;
    const n = await Le.fromFilterColumn(
      this._injector,
      e,
      r
    );
    return this.filterByModel = n, this._filterBy$.next(P.COLORS), this._listenToFilterHeaderChange(e, r), !0;
  }
  _setupByConditions(e, r) {
    this._disposePreviousModel();
    const s = e.getRange();
    if (s.startRow === s.endRow) return !1;
    const n = Pe.fromFilterColumn(
      this._injector,
      e,
      r,
      e.getFilterColumn(r)
    );
    return this.filterByModel = n, this._filterBy$.next(P.CONDITIONS), this._listenToFilterHeaderChange(e, r), !0;
  }
  _disposePreviousModel() {
    var e;
    (e = this._filterByModel) == null || e.dispose(), this.filterByModel = null;
  }
};
X = Ve([
  Ee(0, T(re)),
  Ee(1, T(Qr))
], X);
let Pe = class extends he {
  constructor(e, r, s, n, a) {
    super();
    v(this, "canApply$", de(!0));
    v(this, "_conditionItem$");
    v(this, "conditionItem$");
    v(this, "_filterConditionFormParams$");
    v(this, "filterConditionFormParams$");
    this._filterModel = e, this.col = r, this._commandService = a, this._conditionItem$ = new Q(s), this.conditionItem$ = this._conditionItem$.asObservable(), this._filterConditionFormParams$ = new Q(n), this.filterConditionFormParams$ = this._filterConditionFormParams$.asObservable();
  }
  /**
   * Create a model with targeting filter column. If there is not a filter column, the model would be created with
   * default values.
   *
   * @param injector
   * @param filterModel
   * @param col
   * @param filterColumn
   *
   * @returns the model to control the panel's state
   */
  static fromFilterColumn(e, r, s, n) {
    const [a, i] = f.testMappingFilterColumn(n == null ? void 0 : n.getColumnData());
    return e.createInstance(Pe, r, s, a, i);
  }
  get conditionItem() {
    return this._conditionItem$.getValue();
  }
  get filterConditionFormParams() {
    return this._filterConditionFormParams$.getValue();
  }
  dispose() {
    super.dispose(), this._conditionItem$.complete(), this._filterConditionFormParams$.complete();
  }
  deltaCol(e) {
    this.col += e;
  }
  clear() {
    return this._disposed ? Promise.resolve(!1) : this._commandService.executeCommand(V.id, {
      unitId: this._filterModel.unitId,
      subUnitId: this._filterModel.subUnitId,
      col: this.col,
      criteria: null
    });
  }
  /**
   * Apply the filter condition to the target filter column.
   */
  async apply() {
    if (this._disposed) return !1;
    const e = f.mapToFilterColumn(this.conditionItem, this.filterConditionFormParams);
    return this._commandService.executeCommand(V.id, {
      unitId: this._filterModel.unitId,
      subUnitId: this._filterModel.subUnitId,
      col: this.col,
      criteria: e
    });
  }
  /**
   * This method would be called when user changes the primary condition. The model would load the corresponding
   * `IFilterConditionFormParams` and load default condition form params.
   */
  onPrimaryConditionChange(e) {
    const r = f.ALL_CONDITIONS.find((s) => s.operator === e);
    if (!r)
      throw new Error(`[ByConditionsModel]: condition item not found for operator: ${e}!`);
    this._conditionItem$.next(r), this._filterConditionFormParams$.next(f.getInitialFormParams(e));
  }
  /**
   * This method would be called when user changes the primary conditions, the input values or "AND" "OR" ratio.
   * If the primary conditions or the ratio is changed, the method would load the corresponding `IFilterCondition`.
   *
   * When the panel call this method, it only has to pass the changed keys.
   *
   * @param params
   */
  onConditionFormChange(e) {
    const r = { ...this.filterConditionFormParams, ...e };
    if (r.and !== !0 && delete r.and, typeof e.and < "u" || typeof e.operator1 < "u" || typeof e.operator2 < "u") {
      const s = f.testMappingParams(r, this.conditionItem.numOfParameters);
      this._conditionItem$.next(s);
    }
    this._filterConditionFormParams$.next(r);
  }
};
Pe = Ve([
  Ee(4, j)
], Pe);
let $e = class extends he {
  constructor(e, r, s, n, a) {
    super();
    v(this, "_rawFilterItems$");
    v(this, "rawFilterItems$");
    v(this, "filterItems$");
    v(this, "_filterItems", []);
    v(this, "_treeMapCache");
    v(this, "canApply$");
    v(this, "_manuallyUpdateFilterItems$");
    v(this, "_searchString$");
    v(this, "searchString$");
    this._filterModel = e, this.col = r, this._commandService = a, this._treeMapCache = n, this._searchString$ = new Q(""), this.searchString$ = this._searchString$.asObservable(), this._rawFilterItems$ = new Q(s), this.rawFilterItems$ = this._rawFilterItems$.asObservable(), this._manuallyUpdateFilterItems$ = new Jr(), this.filterItems$ = ei(
      ti([
        this._searchString$.pipe(
          Wt(500, void 0, { leading: !0, trailing: !0 }),
          Vt(void 0)
        ),
        this._rawFilterItems$
      ]).pipe(
        ae(([i, o]) => {
          if (!i) return o;
          const c = i.toLowerCase().split(/\s+/).filter((h) => !!h);
          return Kt(o, c);
        })
      ),
      this._manuallyUpdateFilterItems$
    ).pipe(ri(1)), this.canApply$ = this.filterItems$.pipe(ae((i) => Xe(i).checked > 0)), this.disposeWithMe(this.filterItems$.subscribe((i) => this._filterItems = i));
  }
  /**
   * Create a model with targeting filter column. If there is not a filter column, the model would be created with
   * default values.
   *
   * @param injector
   * @param filterModel
   * @param col
   *
   * @returns the model to control the panel's state
   */
  static async fromFilterColumn(e, r, s) {
    const n = e.get(De), a = e.get(G), i = e.get(we, Pt.OPTIONAL), { unitId: o, subUnitId: l } = r, c = n.getUniverSheetInstance(o);
    if (!c) throw new Error(`[ByValuesModel]: Workbook not found for filter model with unitId: ${o}!`);
    const h = c == null ? void 0 : c.getSheetBySheetId(l);
    if (!h) throw new Error(`[ByValuesModel]: Worksheet not found for filter model with unitId: ${o} and subUnitId: ${l}!`);
    const m = r.getRange(), C = s, p = r.getFilterColumn(s), _ = p == null ? void 0 : p.getColumnData().filters, S = new Set(_ == null ? void 0 : _.filters), E = !!(_ && _.blank), F = r.getFilteredOutRowsExceptCol(s), b = { ...m, startRow: m.startRow + 1, startColumn: C, endColumn: C };
    let y, R;
    if (i) {
      const $ = await i.getFilterValues({
        unitId: o,
        subUnitId: l,
        filteredOutRowsByOtherColumns: Array.from(F),
        filterColumn: p,
        filters: !!_,
        blankChecked: E,
        iterateRange: b,
        alreadyChecked: Array.from(S)
      });
      y = $.filterTreeItems, R = $.filterTreeMapCache;
    } else {
      const $ = Yt(
        !!_,
        a,
        b,
        h,
        F,
        p,
        S,
        E,
        c.getStyles()
      );
      y = $.filterTreeItems, R = $.filterTreeMapCache;
    }
    return e.createInstance($e, r, s, y, R);
  }
  get rawFilterItems() {
    return this._rawFilterItems$.getValue();
  }
  get filterItems() {
    return this._filterItems;
  }
  get treeMapCache() {
    return this._treeMapCache;
  }
  dispose() {
    this._rawFilterItems$.complete(), this._searchString$.complete();
  }
  deltaCol(e) {
    this.col += e;
  }
  setSearchString(e) {
    this._searchString$.next(e);
  }
  onCheckAllToggled(e) {
    const r = be.deepClone(this._filterItems);
    r.forEach((s) => Se(s, e)), this._manuallyUpdateFilterItems(r);
  }
  /**
   * Toggle a filter item.
   */
  onFilterCheckToggled(e) {
    const r = be.deepClone(this._filterItems), s = Je(r, e.key);
    if (!s)
      return;
    const n = Zt(s);
    Se(s, !n), this._manuallyUpdateFilterItems(r);
  }
  onFilterOnly(e) {
    const r = be.deepClone(this._filterItems);
    r.forEach((s) => Se(s, !1)), e.forEach((s) => {
      const n = Je(r, s);
      n && Se(n, !0);
    }), this._manuallyUpdateFilterItems(r);
  }
  _manuallyUpdateFilterItems(e) {
    this._manuallyUpdateFilterItems$.next(e);
  }
  // expose method here to let the panel change filter items
  // #region ByValuesModel apply methods
  clear() {
    return this._disposed ? Promise.resolve(!1) : this._commandService.executeCommand(V.id, {
      unitId: this._filterModel.unitId,
      subUnitId: this._filterModel.subUnitId,
      col: this.col,
      criteria: null
    });
  }
  /**
   * Apply the filter condition to the target filter column.
   */
  async apply() {
    if (this._disposed)
      return !1;
    const e = Xe(this._filterItems), { checked: r, checkedItems: s } = e, n = this.rawFilterItems;
    let a = 0;
    for (const c of n)
      a += c.count;
    const i = r === 0, o = e.checked === a, l = { colId: this.col };
    if (i)
      throw new Error("[ByValuesModel]: no checked items!");
    if (o)
      return this._commandService.executeCommand(V.id, {
        unitId: this._filterModel.unitId,
        subUnitId: this._filterModel.subUnitId,
        col: this.col,
        criteria: null
      });
    {
      l.filters = {};
      const c = s.filter((m) => m.key !== "empty");
      c.length > 0 && (l.filters = {
        filters: c.flatMap((m) => m.originValues ? Array.from(m.originValues) : [m.title])
      }), c.length !== s.length && (l.filters.blank = !0);
    }
    return this._commandService.executeCommand(V.id, {
      unitId: this._filterModel.unitId,
      subUnitId: this._filterModel.subUnitId,
      col: this.col,
      criteria: l
    });
  }
  // #endregion
};
$e = Ve([
  Ee(4, j)
], $e);
let Le = class extends he {
  constructor(e, r, s, n, a) {
    super();
    v(this, "canApply$", de(!0));
    v(this, "_cellFillColors$");
    v(this, "cellFillColors$");
    v(this, "_cellTextColors$");
    v(this, "cellTextColors$");
    this._filterModel = e, this.col = r, this._commandService = a, this._cellFillColors$ = new Q(Array.from(s.values())), this.cellFillColors$ = this._cellFillColors$.asObservable(), this._cellTextColors$ = new Q(Array.from(n.values())), this.cellTextColors$ = this._cellTextColors$.asObservable();
  }
  /**
   * Create a model with targeting filter column. If there is not a filter column, the model would be created with
   * default values.
   *
   * @param injector
   * @param filterModel
   * @param col
   *
   * @returns the model to control the panel's state
   */
  static async fromFilterColumn(e, r, s) {
    var b, y, R;
    const n = e.get(De), { unitId: a, subUnitId: i } = r, o = n.getUniverSheetInstance(a);
    if (!o) throw new Error(`[ByColorsModel]: Workbook not found for filter model with unitId: ${a}!`);
    const l = o == null ? void 0 : o.getSheetBySheetId(i);
    if (!l) throw new Error(`[ByColorsModel]: Worksheet not found for filter model with unitId: ${a} and subUnitId: ${i}!`);
    const c = r.getRange(), h = s, m = (b = r.getFilterColumn(s)) == null ? void 0 : b.getColumnData().colorFilters, C = r.getFilteredOutRowsExceptCol(s), p = { ...c, startRow: c.startRow + 1, startColumn: h, endColumn: h }, _ = /* @__PURE__ */ new Map(), S = new Set((y = m == null ? void 0 : m.cellFillColors) != null ? y : []), E = /* @__PURE__ */ new Map(), F = new Set((R = m == null ? void 0 : m.cellTextColors) != null ? R : []);
    for (const $ of l.iterateByColumn(p, !1, !0)) {
      const { row: Y, col: Z, value: se } = $;
      if (C.has(Y))
        continue;
      const x = l.getComposedCellStyleByCellData(Y, Z, se);
      if (x.bg && x.bg.rgb) {
        const A = new ft(x.bg.rgb).toRgbString();
        _.has(A) || _.set(A, { color: A, checked: S.has(A) });
      } else
        _.set("default-fill-color", { color: null, checked: S.has(null) });
      if (x.cl && x.cl.rgb) {
        const A = new ft(x.cl.rgb).toRgbString();
        E.has(A) || E.set(A, { color: A, checked: F.has(A) });
      } else
        E.set("default-font-color", { color: gt, checked: F.has(gt) });
    }
    return e.createInstance(Le, r, s, _, E);
  }
  get cellFillColors() {
    return this._cellFillColors$.getValue();
  }
  get cellTextColors() {
    return this._cellTextColors$.getValue();
  }
  dispose() {
    super.dispose(), this._cellFillColors$.complete();
  }
  deltaCol(e) {
    this.col += e;
  }
  // expose method here to let the panel change filter items
  // #region ByColorsModel apply methods
  clear() {
    return this._disposed ? Promise.resolve(!1) : this._commandService.executeCommand(V.id, {
      unitId: this._filterModel.unitId,
      subUnitId: this._filterModel.subUnitId,
      col: this.col,
      criteria: null
    });
  }
  onFilterCheckToggled(e, r = !0) {
    const s = r ? this.cellFillColors : this.cellTextColors, n = [];
    let a = !1;
    for (let i = 0; i < s.length; i++) {
      const o = s[i];
      if (o.color === e.color) {
        a = !0, n.push({
          color: o.color,
          checked: !o.checked
        });
        continue;
      }
      n.push({
        color: o.color,
        checked: o.checked
      });
    }
    a && (this._resetColorsCheckedStatus(!r), r ? this._cellFillColors$.next([...n]) : this._cellTextColors$.next([...n]));
  }
  _resetColorsCheckedStatus(e = !0) {
    const r = e ? this.cellFillColors : this.cellTextColors, s = [];
    for (let n = 0; n < r.length; n++)
      s.push({
        color: r[n].color,
        checked: !1
      });
    e ? this._cellFillColors$.next([...s]) : this._cellTextColors$.next([...s]);
  }
  /**
   * Apply the filter condition to the target filter column.
   */
  async apply() {
    if (this._disposed)
      return !1;
    const e = this.cellFillColors.filter((n) => n.checked).map((n) => n.color), r = this.cellTextColors.filter((n) => n.checked).map((n) => n.color);
    if (e.length === 0 && r.length === 0)
      return this._commandService.executeCommand(V.id, {
        unitId: this._filterModel.unitId,
        subUnitId: this._filterModel.subUnitId,
        col: this.col,
        criteria: null
      });
    const s = { colId: this.col };
    return e.length > 0 ? s.colorFilters = {
      cellFillColors: e
    } : r.length > 0 && (s.colorFilters = {
      cellTextColors: r
    }), this._commandService.executeCommand(V.id, {
      unitId: this._filterModel.unitId,
      subUnitId: this._filterModel.subUnitId,
      col: this.col,
      criteria: s
    });
  }
  // #endregion
};
Le = Ve([
  Ee(4, j)
], Le);
const ce = "FILTER_PANEL_OPENED", Me = {
  id: "sheet.operation.open-filter-panel",
  type: Be.OPERATION,
  handler: (t, e) => {
    const r = t.get(We), s = t.get(z), n = t.get(X), a = t.get(j), i = t.has(vt) ? t.get(vt) : null;
    i != null && i.isVisible().visible && a.syncExecuteCommand(Nr.id, { visible: !1 });
    const { unitId: o, subUnitId: l, col: c } = e, h = s.getFilterModel(o, l);
    return h ? (n.setupCol(h, c), r.getContextValue(ce) || r.setContextValue(ce, !0), !0) : !1;
  }
}, Ce = {
  id: "sheet.operation.close-filter-panel",
  type: Be.OPERATION,
  handler: (t) => {
    const e = t.get(We), r = t.get(X), s = t.get(Mr, Pt.OPTIONAL);
    return e.getContextValue(ce) ? (e.setContextValue(ce, !1), s == null || s.focus(), r.terminate()) : !1;
  }
}, qt = {
  id: "sheet.operation.apply-filter",
  type: Be.OPERATION,
  handler: (t, e) => {
    const { filterBy: r } = e;
    return t.get(X).changeFilterBy(r);
  }
}, Xt = "sheets-filter-ui.config", xe = {};
var Ti = Object.getOwnPropertyDescriptor, Ei = (t, e, r, s) => {
  for (var n = s > 1 ? void 0 : s ? Ti(e, r) : e, a = t.length - 1, i; a >= 0; a--)
    (i = t[a]) && (n = i(n) || n);
  return n;
}, le = (t, e) => (r, s) => e(r, s, t);
let ue = class extends he {
  constructor(t, e, r, s, n, a) {
    super(), this._sheetsFilterService = t, this._localeService = e, this._commandService = r, this._sheetPermissionCheckPermission = s, this._injector = n, this._sheetsSelectionService = a, this._commandExecutedListener();
  }
  _commandExecutedListener() {
    this.disposeWithMe(
      this._commandService.beforeCommandExecuted((t) => {
        var e, r, s;
        if (t.id === Ne.id) {
          const n = this._injector.get(De), a = Yr(n);
          if (!a) return;
          const { unitId: i, subUnitId: o, worksheet: l } = a, c = (e = this._sheetsFilterService.getFilterModel(i, o)) == null ? void 0 : e.getRange();
          let h;
          if (c)
            h = this._sheetPermissionCheckPermission.permissionCheckWithRanges({
              rangeTypes: [ge],
              worksheetTypes: [ve, _e]
            }, [c]);
          else {
            const m = (r = this._sheetsSelectionService.getCurrentLastSelection()) == null ? void 0 : r.range;
            if (m) {
              let C = { ...m };
              C = m.startColumn === m.endColumn && m.startRow === m.endRow ? Zr(C, { left: !0, right: !0, up: !0, down: !0 }, l) : C, h = this._sheetPermissionCheckPermission.permissionCheckWithRanges({
                rangeTypes: [ge],
                worksheetTypes: [_e, ve]
              }, [C], i, o);
            } else
              h = this._sheetPermissionCheckPermission.permissionCheckWithoutRange({
                rangeTypes: [ge],
                worksheetTypes: [_e, ve]
              });
          }
          h || this._sheetPermissionCheckPermission.blockExecuteWithoutPermission(this._localeService.t("permission.dialog.filterErr"));
        }
        if (t.id === Me.id) {
          const n = t.params, { unitId: a, subUnitId: i } = n, o = (s = this._sheetsFilterService.getFilterModel(a, i)) == null ? void 0 : s.getRange(), l = be.deepClone(o);
          l && (l.startColumn = n.col, l.endColumn = n.col, this._sheetPermissionCheckPermission.permissionCheckWithRanges({
            rangeTypes: [ge],
            worksheetTypes: [ve, _e]
          }, [l]) || this._sheetPermissionCheckPermission.blockExecuteWithoutPermission(this._localeService.t("permission.dialog.filterErr")));
        }
      })
    );
  }
};
ue = Ei([
  le(0, T(z)),
  le(1, T(G)),
  le(2, j),
  le(3, T(Gr)),
  le(4, T(re)),
  le(5, T(jr))
], ue);
const q = 16, Fi = new Path2D("M3.30363 3C2.79117 3 2.51457 3.60097 2.84788 3.99024L6.8 8.60593V12.5662C6.8 12.7184 6.8864 12.8575 7.02289 12.9249L8.76717 13.7863C8.96655 13.8847 9.2 13.7396 9.2 13.5173V8.60593L13.1521 3.99024C13.4854 3.60097 13.2088 3 12.6964 3H3.30363Z");
class Nt {
  static drawNoCriteria(e, r, s, n) {
    e.save(), St.drawWith(e, {
      radius: 2,
      width: q,
      height: q,
      fill: n
    }), e.lineCap = "square", e.strokeStyle = s, e.scale(r / q, r / q), e.beginPath(), e.lineWidth = 1, e.lineCap = "round", e.moveTo(3, 4), e.lineTo(13, 4), e.moveTo(4.5, 8), e.lineTo(11.5, 8), e.moveTo(6, 12), e.lineTo(10, 12), e.stroke(), e.restore();
  }
  static drawHasCriteria(e, r, s, n) {
    e.save(), St.drawWith(e, {
      radius: 2,
      width: q,
      height: q,
      fill: n
    }), e.scale(r / q, r / q), e.fillStyle = s, e.fill(Fi), e.restore();
  }
}
var Ni = Object.getOwnPropertyDescriptor, yi = (t, e, r, s) => {
  for (var n = s > 1 ? void 0 : s ? Ni(e, r) : e, a = t.length - 1, i; a >= 0; a--)
    (i = t[a]) && (n = i(n) || n);
  return n;
}, Ke = (t, e) => (r, s) => e(r, s, t);
const B = 16, Te = 1;
let et = class extends Vr {
  constructor(e, r, s, n, a) {
    super(e, r);
    v(this, "_cellWidth", 0);
    v(this, "_cellHeight", 0);
    v(this, "_filterParams");
    v(this, "_hovered", !1);
    this._contextService = s, this._commandService = n, this._themeService = a, this.setShapeProps(r), this.onPointerDown$.subscribeEvent((i) => this.onPointerDown(i)), this.onPointerEnter$.subscribeEvent(() => this.onPointerEnter()), this.onPointerLeave$.subscribeEvent(() => this.onPointerLeave());
  }
  setShapeProps(e) {
    typeof e.cellHeight < "u" && (this._cellHeight = e.cellHeight), typeof e.cellWidth < "u" && (this._cellWidth = e.cellWidth), typeof e.filterParams < "u" && (this._filterParams = e.filterParams), this.transformByState({
      width: e.width,
      height: e.height
    });
  }
  _draw(e) {
    const r = this._cellHeight, s = this._cellWidth, n = B - s, a = B - r;
    e.save();
    const i = new Path2D();
    i.rect(n, a, s, r), e.clip(i);
    const { hasCriteria: o } = this._filterParams, l = this._themeService.getColorFromTheme("primary.600"), c = this._hovered ? this._themeService.getColorFromTheme("gray.50") : "rgba(255, 255, 255, 1.0)";
    o ? Nt.drawHasCriteria(e, B, l, c) : Nt.drawNoCriteria(e, B, l, c), e.restore();
  }
  onPointerDown(e) {
    if (e.button === 2)
      return;
    const { col: r, unitId: s, subUnitId: n } = this._filterParams;
    this._contextService.getContextValue(ce) || !this._commandService.hasCommand(Me.id) || setTimeout(() => {
      this._commandService.executeCommand(Me.id, {
        unitId: s,
        subUnitId: n,
        col: r
      });
    }, 200);
  }
  onPointerEnter() {
    this._hovered = !0, this.makeDirty(!0);
  }
  onPointerLeave() {
    this._hovered = !1, this.makeDirty(!0);
  }
};
et = yi([
  Ke(2, We),
  Ke(3, j),
  Ke(4, T($t))
], et);
var Ii = Object.getOwnPropertyDescriptor, Oi = (t, e, r, s) => {
  for (var n = s > 1 ? void 0 : s ? Ii(e, r) : e, a = t.length - 1, i; a >= 0; a--)
    (i = t[a]) && (n = i(n) || n);
  return n;
}, te = (t, e) => (r, s) => e(r, s, t);
const bi = 1e3, Ri = 5e3;
function Ai(t, e, r, s) {
  switch (s) {
    case Re.TOP:
      return t + Te;
    case Re.MIDDLE:
      return t + Math.max(0, (r - B) / 2);
    case Re.BOTTOM:
    default:
      return e - B - Te;
  }
}
let tt = class extends Lt {
  constructor(e, r, s, n, a, i, o, l) {
    super();
    v(this, "_filterRangeShape", null);
    v(this, "_buttonRenderDisposable", null);
    v(this, "_filterButtonShapes", []);
    this._context = e, this._injector = r, this._sheetSkeletonManagerService = s, this._sheetsFilterService = n, this._themeService = a, this._sheetInterceptorService = i, this._commandService = o, this._selectionRenderService = l, this._initRenderer();
  }
  dispose() {
    super.dispose(), this._disposeRendering();
  }
  _initRenderer() {
    this._sheetSkeletonManagerService.currentSkeleton$.pipe(
      lt((e) => {
        var o, l;
        if (!e) return de(null);
        const { unit: r, unitId: s } = this._context, n = ((o = r.getActiveSheet()) == null ? void 0 : o.getSheetId()) || "", a = (l = this._sheetsFilterService.getFilterModel(s, n)) != null ? l : void 0, i = () => ({
          unitId: s,
          worksheetId: n,
          filterModel: a,
          range: a == null ? void 0 : a.getRange(),
          skeleton: e.skeleton
        });
        return dr(this._commandService.onCommandExecuted.bind(this._commandService)).pipe(
          ii(
            ([c]) => {
              var h;
              return c.type === Be.MUTATION && ((h = c.params) == null ? void 0 : h.unitId) === r.getUnitId() && (vr.has(c.id) || c.id === qr.id);
            }
          ),
          Wt(20, void 0, { leading: !1, trailing: !0 }),
          ae(i),
          Vt(i())
          // must trigger once
        );
      }),
      si(this.dispose$)
    ).subscribe((e) => {
      this._disposeRendering(), !(!e || !e.range) && (this._renderRange(e.range, e.skeleton), this._renderButtons(e));
    });
  }
  _renderRange(e, r) {
    const { scene: s } = this._context, { rowHeaderWidth: n, columnHeaderHeight: a } = r, i = this._filterRangeShape = new Or(
      s,
      bi,
      this._themeService,
      {
        rowHeaderWidth: n,
        columnHeaderHeight: a,
        enableAutoFill: !1,
        highlightHeader: !1
      }
    ), l = br({
      range: e,
      primary: null,
      style: { fill: "rgba(0, 0, 0, 0.0)" }
    }, r);
    i.updateRangeBySelectionWithCoord(l), i.setEvent(!1), s.makeDirty(!0);
  }
  _renderButtons(e) {
    const { range: r, filterModel: s, unitId: n, skeleton: a, worksheetId: i } = e, { unit: o, scene: l } = this._context, c = o.getSheetBySheetId(i);
    if (!c)
      return;
    this._interceptCellContent(n, i, e.range);
    const { startColumn: h, endColumn: m, startRow: C } = r;
    for (let p = h; p <= m; p++) {
      const _ = `sheets-filter-button-${p}`, S = Rr(C, p, l, a), E = c.getComposedCellStyle(C, p), F = (E == null ? void 0 : E.vt) || Re.BOTTOM, { startX: b, startY: y, endX: R, endY: $ } = S, Y = R - b, Z = $ - y;
      if (Z <= Te || Y <= Te)
        continue;
      const se = !!s.getFilterColumn(p), x = R - B - Te, A = Ai(y, $, Z, F), w = {
        left: x,
        top: A,
        height: B,
        width: B,
        zIndex: Ri,
        cellHeight: Z,
        cellWidth: Y,
        filterParams: { unitId: n, subUnitId: i, col: p, hasCriteria: se }
      }, Qe = this._injector.createInstance(et, _, w);
      this._filterButtonShapes.push(Qe);
    }
    l.addObjects(this._filterButtonShapes), l.makeDirty();
  }
  _interceptCellContent(e, r, s) {
    const { startRow: n, startColumn: a, endColumn: i } = s;
    this._buttonRenderDisposable = this._sheetInterceptorService.intercept(Xr.CELL_CONTENT, {
      effect: mr.Style,
      handler: (o, l, c) => {
        const { row: h, col: m, unitId: C, subUnitId: p } = l;
        return C !== e || p !== r || h !== n || m < a || m > i || ((!o || o === l.rawData) && (o = { ...l.rawData }), o.fontRenderExtension = {
          ...o == null ? void 0 : o.fontRenderExtension,
          rightOffset: B
        }), c(o);
      },
      priority: 10
    });
  }
  _disposeRendering() {
    var e, r;
    (e = this._filterRangeShape) == null || e.dispose(), this._filterButtonShapes.forEach((s) => s.dispose()), (r = this._buttonRenderDisposable) == null || r.dispose(), this._filterRangeShape = null, this._buttonRenderDisposable = null, this._filterButtonShapes = [];
  }
};
tt = Oi([
  te(1, T(re)),
  te(2, T(yr)),
  te(3, T(z)),
  te(4, T($t)),
  te(5, T(Kr)),
  te(6, j),
  te(7, Ir)
], tt);
var wi = Object.getOwnPropertyDescriptor, Pi = (t, e, r, s) => {
  for (var n = s > 1 ? void 0 : s ? wi(e, r) : e, a = t.length - 1, i; a >= 0; a--)
    (i = t[a]) && (n = i(n) || n);
  return n;
}, yt = (t, e) => (r, s) => e(r, s, t);
let Fe = class extends Lt {
  constructor(t, e) {
    super(), this._renderManagerService = t, this._sheetsRenderService = e, [
      _r,
      gr,
      Sr,
      Cr
    ].forEach((r) => this.disposeWithMe(this._sheetsRenderService.registerSkeletonChangingMutations(r.id))), this.disposeWithMe(this._renderManagerService.registerRenderModule(
      ie.UNIVER_SHEET,
      [tt]
    ));
  }
};
Fe = Pi([
  yt(0, Bt),
  yt(1, T(Ht))
], Fe);
var $i = Object.defineProperty, Li = Object.getOwnPropertyDescriptor, Mi = (t, e, r) => e in t ? $i(t, e, { enumerable: !0, configurable: !0, writable: !0, value: r }) : t[e] = r, xi = (t, e, r, s) => {
  for (var n = s > 1 ? void 0 : s ? Li(e, r) : e, a = t.length - 1, i; a >= 0; a--)
    (i = t[a]) && (n = i(n) || n);
  return n;
}, It = (t, e) => (r, s) => e(r, s, t), zt = (t, e, r) => Mi(t, typeof e != "symbol" ? e + "" : e, r);
const Ui = "SHEET_FILTER_UI_PLUGIN";
let Ue = class extends rt {
  constructor(t = xe, e, r) {
    super(), this._config = t, this._injector = e, this._configService = r;
    const { menu: s, ...n } = Ut(
      {},
      xe,
      this._config
    );
    s && this._configService.setConfig("menu", s, { merge: !0 }), this._configService.setConfig(Xt, n);
  }
  onStarting() {
    [
      [ue],
      [Fe]
    ].forEach((t) => this._injector.add(t));
  }
  onReady() {
    this._injector.get(ue);
  }
  onRendered() {
    this._injector.get(Fe);
  }
};
zt(Ue, "type", ie.UNIVER_SHEET);
zt(Ue, "pluginName", Ui);
Ue = xi([
  Mt(kt),
  It(1, T(re)),
  It(2, xt)
], Ue);
function fe({ ref: t, ...e }) {
  const { icon: r, id: s, className: n, extend: a, ...i } = e, o = `univerjs-icon univerjs-icon-${s} ${n || ""}`.trim(), l = fi(`_${Di()}`);
  return Jt(r, `${s}`, {
    defIds: r.defIds,
    idSuffix: l.current
  }, {
    ref: t,
    className: o,
    ...i
  }, a);
}
function Jt(t, e, r, s, n) {
  return me(t.tag, {
    key: e,
    ...ki(t, r, n),
    ...s
  }, (Hi(t, r).children || []).map((a, i) => Jt(a, `${e}-${t.tag}-${i}`, r, void 0, n)));
}
function ki(t, e, r) {
  const s = { ...t.attrs };
  r != null && r.colorChannel1 && s.fill === "colorChannel1" && (s.fill = r.colorChannel1), t.tag === "mask" && s.id && (s.id = s.id + e.idSuffix), Object.entries(s).forEach(([a, i]) => {
    a === "mask" && typeof i == "string" && (s[a] = i.replace(/url\(#(.*)\)/, `url(#$1${e.idSuffix})`));
  });
  const { defIds: n } = e;
  return !n || n.length === 0 || (t.tag === "use" && s["xlink:href"] && (s["xlink:href"] = s["xlink:href"] + e.idSuffix), Object.entries(s).forEach(([a, i]) => {
    typeof i == "string" && (s[a] = i.replace(/url\(#(.*)\)/, `url(#$1${e.idSuffix})`));
  })), s;
}
function Hi(t, e) {
  var s;
  const { defIds: r } = e;
  return !r || r.length === 0 ? t : t.tag === "defs" && ((s = t.children) != null && s.length) ? {
    ...t,
    children: t.children.map((n) => typeof n.attrs.id == "string" && r && r.includes(n.attrs.id) ? {
      ...n,
      attrs: {
        ...n.attrs,
        id: n.attrs.id + e.idSuffix
      }
    } : n)
  } : t;
}
function Di() {
  return Math.random().toString(36).substring(2, 8);
}
fe.displayName = "UniverIcon";
const Bi = {
  tag: "svg",
  attrs: {
    xmlns: "http://www.w3.org/2000/svg",
    fill: "none",
    viewBox: "0 0 20 20",
    width: "1em",
    height: "1em"
  },
  children: [{
    tag: "path",
    attrs: {
      fill: "currentColor",
      d: "M10 1.05957C10.356 1.05957 10.6816 1.26162 10.8408 1.58008L18.8408 17.5801L18.8799 17.668C19.0486 18.1134 18.8551 18.6232 18.4199 18.8408C17.9557 19.0727 17.3913 18.8841 17.1592 18.4199L10 4.10156L2.84082 18.4199C2.60871 18.8841 2.04434 19.0727 1.58008 18.8408C1.11587 18.6087 0.92731 18.0443 1.15918 17.5801L9.15918 1.58008C9.31841 1.26162 9.64395 1.05957 10 1.05957Z"
    }
  }, {
    tag: "path",
    attrs: {
      fill: "currentColor",
      d: "M15.3337 11.7261L15.4294 11.731C15.9035 11.779 16.2732 12.1798 16.2732 12.6665C16.2732 13.1532 15.9035 13.554 15.4294 13.602L15.3337 13.6069H4.66675C4.1476 13.6069 3.72632 13.1856 3.72632 12.6665C3.72632 12.1474 4.1476 11.7261 4.66675 11.7261H15.3337Z"
    }
  }]
}, er = ye(function(e, r) {
  return me(fe, Object.assign({}, e, {
    id: "a-icon",
    ref: r,
    icon: Bi
  }));
});
er.displayName = "AIcon";
const Wi = {
  tag: "svg",
  attrs: {
    xmlns: "http://www.w3.org/2000/svg",
    fill: "none",
    viewBox: "0 0 20 20",
    width: "1em",
    height: "1em"
  },
  children: [{
    tag: "path",
    attrs: {
      fill: "currentColor",
      d: "M17.0596 10C17.0596 6.10087 13.8992 2.94043 10 2.94043C6.10087 2.94043 2.94043 6.10087 2.94043 10C2.94043 13.8992 6.10087 17.0596 10 17.0596C13.8992 17.0596 17.0596 13.8992 17.0596 10ZM18.9404 10C18.9404 14.9374 14.9374 18.9404 10 18.9404C5.06257 18.9404 1.05957 14.9374 1.05957 10C1.05957 5.06257 5.06257 1.05957 10 1.05957C14.9374 1.05957 18.9404 5.06257 18.9404 10Z"
    }
  }, {
    tag: "path",
    attrs: {
      fill: "currentColor",
      d: "M4.29492 4.13476C4.63911 3.79057 5.1845 3.76906 5.55371 4.07031L5.625 4.13476L16.0244 14.5352L16.0889 14.6064C16.3902 14.9757 16.3686 15.52 16.0244 15.8643C15.6573 16.2313 15.0624 16.2313 14.6953 15.8643L4.29492 5.46484L4.23047 5.39355C3.92922 5.02434 3.95073 4.47895 4.29492 4.13476Z"
    }
  }]
}, tr = ye(function(e, r) {
  return me(fe, Object.assign({}, e, {
    id: "ban-icon",
    ref: r,
    icon: Wi
  }));
});
tr.displayName = "BanIcon";
const Vi = {
  tag: "svg",
  attrs: {
    xmlns: "http://www.w3.org/2000/svg",
    fill: "none",
    viewBox: "0 0 17 16",
    width: "1em",
    height: "1em"
  },
  children: [{
    tag: "path",
    attrs: {
      fill: "currentColor",
      d: "M3.32182 2.60967C2.98161 2.60967 2.79671 3.0074 3.01601 3.2675L6.85819 7.8246C6.94943 7.93282 6.99947 8.06981 6.99947 8.21136V12.7338C6.99947 12.898 7.0998 13.0455 7.2525 13.1058L8.73833 13.6928C9.00085 13.7965 9.28531 13.6031 9.28531 13.3208V8.21136C9.28531 8.06981 9.33535 7.93282 9.42659 7.8246L13.2688 3.2675C13.4881 3.0074 13.3032 2.60967 12.963 2.60967H3.32182ZM2.09858 4.04101C1.22139 3.0006 1.96097 1.40967 3.32182 1.40967H12.963C14.3238 1.40967 15.0634 3.0006 14.1862 4.04101L10.4853 8.43054V13.3208C10.4853 14.4498 9.34747 15.2237 8.29742 14.8089L6.81158 14.2219C6.20078 13.9806 5.79947 13.3905 5.79947 12.7338V8.43054L2.09858 4.04101Z",
      fillRule: "evenodd",
      clipRule: "evenodd"
    }
  }]
}, rr = ye(function(e, r) {
  return me(fe, Object.assign({}, e, {
    id: "filter-icon",
    ref: r,
    icon: Vi
  }));
});
rr.displayName = "FilterIcon";
const Qi = {
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
      d: "M8.00016 1.33203C6.68162 1.33203 5.39269 1.72302 4.29636 2.45557C3.20004 3.18811 2.34555 4.2293 1.84097 5.44747C1.33638 6.66565 1.20436 8.00609 1.4616 9.2993C1.71883 10.5925 2.35377 11.7804 3.28612 12.7127C4.21847 13.6451 5.40636 14.28 6.69956 14.5373C7.99277 14.7945 9.33321 14.6625 10.5514 14.1579C11.7696 13.6533 12.8108 12.7988 13.5433 11.7025C14.2758 10.6062 14.6668 9.31724 14.6668 7.9987C14.6649 6.23118 13.9619 4.53662 12.7121 3.2868C11.4622 2.03697 9.76768 1.33397 8.00016 1.33203ZM7.66683 3.9987C7.86461 3.9987 8.05795 4.05735 8.2224 4.16723C8.38685 4.27711 8.51502 4.43329 8.59071 4.61601C8.6664 4.79874 8.6862 4.99981 8.64762 5.19379C8.60903 5.38777 8.51379 5.56595 8.37394 5.7058C8.23409 5.84566 8.0559 5.9409 7.86192 5.97948C7.66794 6.01807 7.46687 5.99826 7.28415 5.92258C7.10142 5.84689 6.94524 5.71872 6.83536 5.55427C6.72548 5.38982 6.66683 5.19648 6.66683 4.9987C6.66683 4.73348 6.77219 4.47913 6.95972 4.29159C7.14726 4.10405 7.40162 3.9987 7.66683 3.9987ZM9.3335 11.332H6.66683C6.49002 11.332 6.32045 11.2618 6.19543 11.1368C6.0704 11.0117 6.00016 10.8422 6.00016 10.6654C6.00016 10.4886 6.0704 10.319 6.19543 10.194C6.32045 10.0689 6.49002 9.9987 6.66683 9.9987H7.3335V7.9987H6.66683C6.49002 7.9987 6.32045 7.92846 6.19543 7.80343C6.0704 7.67841 6.00016 7.50884 6.00016 7.33203C6.00016 7.15522 6.0704 6.98565 6.19543 6.86063C6.32045 6.7356 6.49002 6.66536 6.66683 6.66536H8.00016C8.17698 6.66536 8.34655 6.7356 8.47157 6.86063C8.59659 6.98565 8.66683 7.15522 8.66683 7.33203V9.9987H9.3335C9.51031 9.9987 9.67988 10.0689 9.8049 10.194C9.92993 10.319 10.0002 10.4886 10.0002 10.6654C10.0002 10.8422 9.92993 11.0117 9.8049 11.1368C9.67988 11.2618 9.51031 11.332 9.3335 11.332Z"
    }
  }]
}, ir = ye(function(e, r) {
  return me(fe, Object.assign({}, e, {
    id: "info-icon",
    ref: r,
    icon: Qi
  }));
});
ir.displayName = "InfoIcon";
const Gi = {
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
      d: "M8 15C11.866 15 15 11.866 15 8C15 4.13401 11.866 1 8 1C4.13401 1 1 4.13401 1 8C1 11.866 4.13401 15 8 15ZM11.7245 6.42417C11.9588 6.18985 11.9588 5.80995 11.7245 5.57564C11.4901 5.34132 11.1102 5.34132 10.8759 5.57564L7.3002 9.15137L5.72446 7.57564C5.49014 7.34132 5.11025 7.34132 4.87593 7.57564C4.64162 7.80995 4.64162 8.18985 4.87593 8.42417L6.87593 10.4242C7.11025 10.6585 7.49014 10.6585 7.72446 10.4242L11.7245 6.42417Z",
      fillRule: "evenodd",
      clipRule: "evenodd"
    }
  }]
}, sr = ye(function(e, r) {
  return me(fe, Object.assign({}, e, {
    id: "success-icon",
    ref: r,
    icon: Gi
  }));
});
sr.displayName = "SuccessIcon";
function ji(t) {
  const { model: e } = t, r = D(G), s = M(e.cellFillColors$, [], !0), n = M(e.cellTextColors$, [], !0), a = k((o) => {
    e.onFilterCheckToggled(o);
  }, [e]), i = k((o) => {
    e.onFilterCheckToggled(o, !1);
  }, [e]);
  return /* @__PURE__ */ d(
    "div",
    {
      "data-u-comp": "sheets-filter-panel-colors-container",
      className: "univer-flex univer-h-full univer-min-h-[300px] univer-flex-col",
      children: /* @__PURE__ */ N(
        "div",
        {
          "data-u-comp": "sheets-filter-panel",
          className: Ae("univer-mt-2 univer-box-border univer-flex univer-h-[300px] univer-flex-grow univer-flex-col univer-gap-4 univer-overflow-auto univer-rounded-md univer-px-2 univer-py-2.5", at),
          children: [
            s.length > 1 && /* @__PURE__ */ N("div", { children: [
              /* @__PURE__ */ d(
                "div",
                {
                  className: "univer-mb-2 univer-text-sm univer-text-gray-900 dark:!univer-text-white",
                  children: r.t("sheets-filter.panel.filter-by-cell-fill-color")
                }
              ),
              /* @__PURE__ */ d(
                "div",
                {
                  className: "univer-grid univer-grid-cols-8 univer-items-center univer-justify-start univer-gap-2",
                  children: s.map((o, l) => /* @__PURE__ */ N(
                    "div",
                    {
                      className: "univer-relative univer-h-6 univer-w-6",
                      onClick: () => a(o),
                      children: [
                        o.color ? /* @__PURE__ */ d(
                          "button",
                          {
                            type: "button",
                            className: Ae("univer-box-border univer-h-6 univer-w-6 univer-cursor-pointer univer-rounded-full univer-border univer-border-solid univer-border-transparent univer-bg-gray-300 univer-transition-shadow hover:univer-ring-2 hover:univer-ring-offset-2 hover:univer-ring-offset-white"),
                            style: { backgroundColor: o.color }
                          }
                        ) : /* @__PURE__ */ d(
                          tr,
                          {
                            className: "univer-h-6 univer-w-6 univer-cursor-pointer univer-rounded-full hover:univer-ring-2 hover:univer-ring-offset-2 hover:univer-ring-offset-white"
                          }
                        ),
                        o.checked && /* @__PURE__ */ d(Ot, {})
                      ]
                    },
                    `sheets-filter-cell-fill-color-${l}`
                  ))
                }
              )
            ] }),
            n.length > 1 && /* @__PURE__ */ N("div", { children: [
              /* @__PURE__ */ d(
                "div",
                {
                  className: "univer-mb-2 univer-text-sm univer-text-gray-900 dark:!univer-text-white",
                  children: r.t("sheets-filter.panel.filter-by-cell-text-color")
                }
              ),
              /* @__PURE__ */ d(
                "div",
                {
                  className: "univer-grid univer-grid-cols-8 univer-items-center univer-justify-start univer-gap-2",
                  children: n.map((o, l) => /* @__PURE__ */ N(
                    "div",
                    {
                      className: "univer-relative univer-h-6 univer-w-6",
                      onClick: () => i(o),
                      children: [
                        /* @__PURE__ */ d(
                          "div",
                          {
                            className: "univer-box-border univer-flex univer-h-full univer-w-full univer-cursor-pointer univer-items-center univer-justify-center univer-rounded-full univer-border univer-border-solid univer-border-[rgba(13,13,13,0.06)] univer-p-0.5 hover:univer-ring-2 hover:univer-ring-offset-2 hover:univer-ring-offset-white dark:!univer-border-[rgba(255,255,255,0.06)]",
                            children: /* @__PURE__ */ d(er, { style: { color: o.color } })
                          }
                        ),
                        o.checked && /* @__PURE__ */ d(Ot, {})
                      ]
                    },
                    `sheets-filter-cell-text-color-${l}`
                  ))
                }
              )
            ] }),
            s.length <= 1 && n.length <= 1 && /* @__PURE__ */ d(
              "div",
              {
                className: "univer-flex univer-h-full univer-w-full univer-items-center univer-justify-center univer-text-sm univer-text-gray-900 dark:!univer-text-gray-200",
                children: r.t("sheets-filter.panel.filter-by-color-none")
              }
            )
          ]
        }
      )
    }
  );
}
function Ot() {
  return /* @__PURE__ */ d(
    "div",
    {
      className: "univer-absolute -univer-bottom-0.5 -univer-right-0.5 univer-flex univer-h-3 univer-w-3 univer-cursor-pointer univer-items-center univer-justify-center univer-rounded-full univer-bg-white",
      children: /* @__PURE__ */ d(
        sr,
        {
          className: "univer-h-full univer-w-full univer-font-bold univer-text-[#418F1F]"
        }
      )
    }
  );
}
function Yi(t) {
  var p, _;
  const { model: e } = t, r = D(G), s = M(e.conditionItem$, void 0), n = M(e.filterConditionFormParams$, void 0), a = n != null && n.and ? "AND" : "OR", i = k((S) => {
    e.onConditionFormChange({ and: S === "AND" });
  }, [e]), o = Zi(r), l = k((S) => {
    e.onPrimaryConditionChange(S);
  }, [e]), c = Ki(r), h = k((S) => {
    e.onConditionFormChange(S);
  }, [e]), m = r.t("sheets-filter.panel.input-values-placeholder");
  function C(S, E, F) {
    const b = f.getItemByOperator(S).numOfParameters === 1;
    return /* @__PURE__ */ N(Et, { children: [
      F === "operator2" && /* @__PURE__ */ N(ai, { value: a, onChange: i, children: [
        /* @__PURE__ */ d(Tt, { value: "AND", children: r.t("sheets-filter.panel.and") }),
        /* @__PURE__ */ d(Tt, { value: "OR", children: r.t("sheets-filter.panel.or") })
      ] }),
      /* @__PURE__ */ d(
        Ct,
        {
          value: S,
          options: c,
          onChange: (y) => h({ [F]: y })
        }
      ),
      b && /* @__PURE__ */ d("div", { children: /* @__PURE__ */ d(
        Gt,
        {
          className: "univer-mt-2",
          value: E,
          placeholder: m,
          onChange: (y) => h({ [F === "operator1" ? "val1" : "val2"]: y })
        }
      ) })
    ] });
  }
  return /* @__PURE__ */ d(
    "div",
    {
      "data-u-comp": "sheets-filter-panel-conditions-container",
      className: "univer-flex univer-h-full univer-min-h-[300px] univer-flex-col",
      children: s && n && /* @__PURE__ */ N(Et, { children: [
        /* @__PURE__ */ d(Ct, { value: s.operator, options: o, onChange: l }),
        f.getItemByOperator(s.operator).numOfParameters !== 0 ? /* @__PURE__ */ N(
          "div",
          {
            "data-u-comp": "sheets-filter-panel-conditions-container-inner",
            className: Ae("univer-mt-2 univer-flex-grow univer-overflow-hidden univer-rounded-md univer-p-2", at),
            children: [
              s.numOfParameters >= 1 && C(n.operator1, (p = n.val1) != null ? p : "", "operator1"),
              s.numOfParameters >= 2 && C(n.operator2, (_ = n.val2) != null ? _ : "", "operator2"),
              /* @__PURE__ */ N(
                "div",
                {
                  "data-u-comp": "sheets-filter-panel-conditions-desc",
                  className: "univer-mt-2 univer-text-xs univer-text-gray-500",
                  children: [
                    r.t("sheets-filter.panel.?"),
                    /* @__PURE__ */ d("br", {}),
                    r.t("sheets-filter.panel.*")
                  ]
                }
              )
            ]
          }
        ) : null
      ] })
    }
  );
}
function Zi(t) {
  const e = t.getCurrentLocale();
  return ct(() => [
    {
      options: [
        { label: t.t(f.NONE.label), value: f.NONE.operator }
      ]
    },
    {
      options: [
        { label: t.t(f.EMPTY.label), value: f.EMPTY.operator },
        { label: t.t(f.NOT_EMPTY.label), value: f.NOT_EMPTY.operator }
      ]
    },
    {
      options: [
        { label: t.t(f.TEXT_CONTAINS.label), value: f.TEXT_CONTAINS.operator },
        { label: t.t(f.DOES_NOT_CONTAIN.label), value: f.DOES_NOT_CONTAIN.operator },
        { label: t.t(f.STARTS_WITH.label), value: f.STARTS_WITH.operator },
        { label: t.t(f.ENDS_WITH.label), value: f.ENDS_WITH.operator },
        { label: t.t(f.EQUALS.label), value: f.EQUALS.operator }
      ]
    },
    {
      options: [
        { label: t.t(f.GREATER_THAN.label), value: f.GREATER_THAN.operator },
        { label: t.t(f.GREATER_THAN_OR_EQUAL.label), value: f.GREATER_THAN_OR_EQUAL.operator },
        { label: t.t(f.LESS_THAN.label), value: f.LESS_THAN.operator },
        { label: t.t(f.LESS_THAN_OR_EQUAL.label), value: f.LESS_THAN_OR_EQUAL.operator },
        { label: t.t(f.EQUAL.label), value: f.EQUAL.operator },
        { label: t.t(f.NOT_EQUAL.label), value: f.NOT_EQUAL.operator },
        { label: t.t(f.BETWEEN.label), value: f.BETWEEN.operator },
        { label: t.t(f.NOT_BETWEEN.label), value: f.NOT_BETWEEN.operator }
      ]
    },
    {
      options: [
        { label: t.t(f.CUSTOM.label), value: f.CUSTOM.operator }
      ]
    }
  ], [e, t]);
}
function Ki(t) {
  const e = t.getCurrentLocale();
  return ct(() => f.ALL_CONDITIONS.filter((r) => r.numOfParameters !== 2).map((r) => ({ label: t.t(r.label), value: r.operator })), [e, t]);
}
function qi(t) {
  const { model: e } = t, r = D(G), s = M(e.searchString$, "", !0), n = M(e.filterItems$, void 0, !0), a = r.t("sheets-filter.panel.filter-only"), i = Xe(n), o = i.checked > 0 && i.unchecked === 0, l = i.checked > 0 && i.unchecked > 0, c = e.treeMapCache, h = k(() => {
    e.onCheckAllToggled(!o);
  }, [e, o]), m = k((p) => {
    e.setSearchString(p);
  }, [e]);
  function C(p) {
    let _ = [];
    return p.forEach((S) => {
      S.checked && _.push(S.key), S.children && (_ = _.concat(C(S.children)));
    }), _;
  }
  return /* @__PURE__ */ N(
    "div",
    {
      "data-u-comp": "sheets-filter-panel-values-container",
      className: "univer-flex univer-h-full univer-min-h-[300px] univer-flex-col",
      children: [
        /* @__PURE__ */ d(
          Gt,
          {
            autoFocus: !0,
            value: s,
            placeholder: r.t("sheets-filter.panel.search-placeholder"),
            onChange: m
          }
        ),
        /* @__PURE__ */ N(
          "div",
          {
            "data-u-comp": "sheets-filter-panel",
            className: Ae("univer-mt-2 univer-box-border univer-flex univer-flex-grow univer-flex-col univer-overflow-hidden univer-rounded-md univer-px-2 univer-py-2.5", at),
            children: [
              /* @__PURE__ */ d(
                "div",
                {
                  "data-u-comp": "sheets-filter-panel-values-item",
                  className: "univer-box-border univer-h-8 univer-w-full univer-py-0.5",
                  children: /* @__PURE__ */ N(
                    "div",
                    {
                      "data-u-comp": "sheets-filter-panel-values-item-inner",
                      className: "univer-box-border univer-flex univer-h-7 univer-items-center univer-rounded-md univer-pb-0 univer-pl-5 univer-pr-0.5 univer-pt-0 univer-text-sm",
                      children: [
                        /* @__PURE__ */ d(
                          ci,
                          {
                            indeterminate: l,
                            disabled: n.length === 0,
                            checked: o,
                            onChange: h
                          }
                        ),
                        /* @__PURE__ */ d(
                          "span",
                          {
                            "data-u-comp": "sheets-filter-panel-values-item-text",
                            className: "univer-mx-1 univer-inline-block univer-flex-shrink univer-overflow-hidden univer-text-ellipsis univer-whitespace-nowrap univer-text-gray-900 dark:!univer-text-white",
                            children: `${r.t("sheets-filter.panel.select-all")}`
                          }
                        ),
                        /* @__PURE__ */ d(
                          "span",
                          {
                            "data-u-comp": "sheets-filter-panel-values-item-count",
                            className: "univer-text-gray-400 dark:!univer-text-gray-500",
                            children: `(${i.checked}/${i.checked + i.unchecked})`
                          }
                        )
                      ]
                    }
                  )
                }
              ),
              /* @__PURE__ */ d("div", { "data-u-comp": "sheets-filter-panel-values-virtual", className: "univer-flex-grow", children: /* @__PURE__ */ d(
                ui,
                {
                  data: n,
                  defaultExpandAll: !1,
                  valueGroup: C(n),
                  onChange: (p) => {
                    e.onFilterCheckToggled(p);
                  },
                  defaultCache: c,
                  itemHeight: 28,
                  treeNodeClassName: `
                          univer-pr-2 univer-border-box univer-rounded-md
                          [&:hover_a]:univer-inline-block
                          hover:univer-bg-gray-50 univer-h-full
                          univer-text-gray-900 dark:hover:!univer-bg-gray-900
                          dark:!univer-text-white
                        `,
                  attachRender: (p) => /* @__PURE__ */ N(
                    "div",
                    {
                      className: "univer-ml-1 univer-flex univer-h-5 univer-flex-1 univer-cursor-pointer univer-items-center univer-justify-between univer-text-sm univer-text-primary-500",
                      children: [
                        /* @__PURE__ */ d(
                          "span",
                          {
                            "data-u-comp": "sheets-filter-panel-values-item-count",
                            className: "univer-text-gray-400 dark:!univer-text-gray-500",
                            children: `(${p.count})`
                          }
                        ),
                        /* @__PURE__ */ d(
                          "a",
                          {
                            className: "univer-box-border univer-hidden univer-h-4 univer-whitespace-nowrap univer-px-1.5",
                            onClick: () => {
                              const _ = [];
                              p.children ? p.children.forEach((S) => {
                                S.children ? S.children.forEach((E) => {
                                  _.push(E.key);
                                }) : _.push(S.key);
                              }) : _.push(p.key), e.onFilterOnly(_);
                            },
                            children: a
                          }
                        )
                      ]
                    }
                  )
                }
              ) })
            ]
          }
        )
      ]
    }
  );
}
function Xi() {
  const t = D(Tr);
  if (!M(t.visible$, void 0, !0)) return null;
  const r = D(G), s = D(Dt), n = M(t.enabled$, void 0, !0);
  return /* @__PURE__ */ N(
    "div",
    {
      className: "univer-mt-2 univer-flex univer-items-center univer-justify-between univer-text-sm univer-text-gray-900 dark:!univer-text-gray-200",
      children: [
        /* @__PURE__ */ N("div", { className: "univer-flex univer-items-center univer-gap-1", children: [
          /* @__PURE__ */ d("span", { children: r.t("sheets-filter.sync.title") }),
          /* @__PURE__ */ d(
            hi,
            {
              title: n ? r.t("sheets-filter.sync.statusTips.off") : r.t("sheets-filter.sync.statusTips.on"),
              asChild: !0,
              children: /* @__PURE__ */ d(ir, { className: "univer-block" })
            }
          )
        ] }),
        /* @__PURE__ */ d(
          di,
          {
            defaultChecked: n,
            onChange: (a) => {
              const i = a ? r.t("sheets-filter.sync.switchTips.on") : r.t("sheets-filter.sync.switchTips.off");
              t.setEnabled(a), s.show({
                content: i,
                type: jt.Success,
                duration: 2e3
              });
            }
          }
        )
      ]
    }
  );
}
function zi() {
  var E;
  const t = D(X), e = D(G), r = D(j), s = M(t.filterBy$, void 0, !0), n = M(t.filterByModel$, void 0, !1), a = M(() => (n == null ? void 0 : n.canApply$) || de(!1), void 0, !1, [n]), i = Ji(e), o = !M(t.hasCriteria$), l = k((F) => {
    r.executeCommand(qt.id, { filterBy: F });
  }, [r]), c = k(async () => {
    await (n == null ? void 0 : n.clear()), r.executeCommand(Ce.id);
  }, [n, r]), h = k(() => {
    r.executeCommand(Ce.id);
  }, [r]), m = k(async () => {
    await (n == null ? void 0 : n.apply()), r.executeCommand(Ce.id);
  }, [n, r]), p = (E = D(z).activeFilterModel) == null ? void 0 : E.getRange(), _ = t.col, S = xr(Ar.FILTER_PANEL_EMBED_POINT);
  return /* @__PURE__ */ N(
    "div",
    {
      "data-u-comp": "sheets-filter-panel",
      className: "univer-box-border univer-flex univer-max-h-[500px] univer-w-[400px] univer-flex-col univer-rounded-lg univer-bg-white univer-p-4 univer-shadow-lg dark:!univer-border-gray-600 dark:!univer-bg-gray-700",
      children: [
        /* @__PURE__ */ d(
          Ur,
          {
            components: S,
            sharedProps: { range: p, colIndex: _, onClose: h }
          }
        ),
        /* @__PURE__ */ d("div", { className: "univer-mb-1 univer-flex-shrink-0 univer-flex-grow-0", children: /* @__PURE__ */ d(
          mi,
          {
            value: s,
            items: i,
            onChange: (F) => l(F)
          }
        ) }),
        n ? /* @__PURE__ */ d(
          "div",
          {
            "data-u-comp": "sheets-filter-panel-content",
            className: "univer-flex-shrink univer-flex-grow univer-pt-2",
            children: s === P.VALUES ? /* @__PURE__ */ d(qi, { model: n }) : s === P.COLORS ? /* @__PURE__ */ d(ji, { model: n }) : /* @__PURE__ */ d(Yi, { model: n })
          }
        ) : /* @__PURE__ */ d("div", { className: "univer-flex-1" }),
        /* @__PURE__ */ d(Xi, {}),
        /* @__PURE__ */ N(
          "div",
          {
            "data-u-comp": "sheets-filter-panel-footer",
            className: "univer-mt-4 univer-inline-flex univer-flex-shrink-0 univer-flex-grow-0 univer-flex-nowrap univer-justify-between univer-overflow-hidden",
            children: [
              /* @__PURE__ */ d(Ye, { variant: "link", onClick: c, disabled: o, children: e.t("sheets-filter.panel.clear-filter") }),
              /* @__PURE__ */ N("span", { className: "univer-flex univer-gap-2", children: [
                /* @__PURE__ */ d(Ye, { variant: "default", onClick: h, children: e.t("sheets-filter.panel.cancel") }),
                /* @__PURE__ */ d(Ye, { disabled: !a, variant: "primary", onClick: m, children: e.t("sheets-filter.panel.confirm") })
              ] })
            ]
          }
        )
      ]
    }
  );
}
function Ji(t) {
  const e = t.getCurrentLocale();
  return ct(() => [
    { label: t.t("sheets-filter.panel.by-values"), value: P.VALUES },
    { label: t.t("sheets-filter.panel.by-colors"), value: P.COLORS },
    { label: t.t("sheets-filter.panel.by-conditions"), value: P.CONDITIONS }
  ], [e, t]);
}
function es(t) {
  const e = t.get(z);
  return {
    id: Ne.id,
    type: ot.BUTTON_SELECTOR,
    icon: "FilterIcon",
    tooltip: "sheets-filter.toolbar.smart-toggle-filter-tooltip",
    hidden$: nt(t, ie.UNIVER_SHEET),
    activated$: e.activeFilterModel$.pipe(ae((r) => !!r)),
    disabled$: wr(
      t,
      Pr(
        t,
        {
          worksheetTypes: [ve, _e],
          rangeTypes: [ge]
        }
      )
    )
  };
}
function ts(t) {
  const e = t.get(z);
  return {
    id: st.id,
    type: ot.BUTTON,
    title: "sheets-filter.toolbar.clear-filter-criteria",
    hidden$: nt(t, ie.UNIVER_SHEET),
    disabled$: e.activeFilterModel$.pipe(lt((r) => {
      var s;
      return (s = r == null ? void 0 : r.hasCriteria$.pipe(ae((n) => !n))) != null ? s : de(!0);
    }))
  };
}
function rs(t) {
  const e = t.get(z);
  return {
    id: it.id,
    type: ot.BUTTON,
    title: "sheets-filter.toolbar.re-calc-filter-conditions",
    hidden$: nt(t, ie.UNIVER_SHEET),
    disabled$: e.activeFilterModel$.pipe(lt((r) => {
      var s;
      return (s = r == null ? void 0 : r.hasCriteria$.pipe(ae((n) => !n))) != null ? s : de(!0);
    }))
  };
}
const is = {
  [kr.OTHERS]: {
    [Ne.id]: {
      order: 2,
      menuItemFactory: es,
      [st.id]: {
        order: 0,
        menuItemFactory: ts
      },
      [it.id]: {
        order: 1,
        menuItemFactory: rs
      }
    }
  }
}, ss = {
  id: Ne.id,
  binding: Hr.L | _t.CTRL_COMMAND | _t.SHIFT,
  description: "sheets-filter.shortcut.smart-toggle-filter",
  preconditions: $r,
  group: "4_sheet-edit"
};
var ns = Object.getOwnPropertyDescriptor, os = (t, e, r, s) => {
  for (var n = s > 1 ? void 0 : s ? ns(e, r) : e, a = t.length - 1, i; a >= 0; a--)
    (i = t[a]) && (n = i(n) || n);
  return n;
}, L = (t, e) => (r, s) => e(r, s, t);
const bt = "FILTER_PANEL_POPUP";
let ke = class extends Fe {
  constructor(e, r, s, n, a, i, o, l, c, h, m, C, p) {
    super(p, C);
    v(this, "_popupDisposable");
    this._injector = e, this._componentManager = r, this._sheetsFilterPanelService = s, this._sheetCanvasPopupService = n, this._sheetsFilterService = a, this._localeService = i, this._shortcutService = o, this._commandService = l, this._menuManagerService = c, this._contextService = h, this._messageService = m, this._initCommands(), this._initShortcuts(), this._initMenuItems(), this._initUI();
  }
  dispose() {
    super.dispose(), this._closeFilterPopup();
  }
  _initShortcuts() {
    [
      ss
    ].forEach((e) => {
      this.disposeWithMe(this._shortcutService.registerShortcut(e));
    });
  }
  _initCommands() {
    [
      Ne,
      Er,
      Fr,
      V,
      st,
      it,
      qt,
      Me,
      Ce
    ].forEach((e) => {
      this.disposeWithMe(this._commandService.registerCommand(e));
    });
  }
  _initMenuItems() {
    this._menuManagerService.mergeMenu(is);
  }
  _initUI() {
    [
      [bt, zi],
      ["FilterIcon", rr]
    ].forEach(([e, r]) => {
      this.disposeWithMe(
        this._componentManager.register(e, r)
      );
    }), this.disposeWithMe(this._contextService.subscribeContextValue$(ce).pipe(ni()).subscribe((e) => {
      e ? this._openFilterPopup() : this._closeFilterPopup();
    })), this.disposeWithMe(this._sheetsFilterService.errorMsg$.subscribe((e) => {
      e && this._messageService.show({
        type: jt.Error,
        content: this._localeService.t(e)
      });
    }));
  }
  _openFilterPopup() {
    const e = this._sheetsFilterPanelService.filterModel;
    if (!e)
      throw new Error("[SheetsFilterUIController]: no filter model when opening filter popup!");
    const r = e.getRange(), s = this._sheetsFilterPanelService.col, { startRow: n } = r;
    this._popupDisposable = this._sheetCanvasPopupService.attachPopupToCell(n, s, {
      componentKey: bt,
      direction: "horizontal",
      onClickOutside: () => this._commandService.syncExecuteCommand(Ce.id),
      offset: [5, 0]
    });
  }
  _closeFilterPopup() {
    var e;
    (e = this._popupDisposable) == null || e.dispose(), this._popupDisposable = null;
  }
};
ke = os([
  L(0, T(re)),
  L(1, T(Dr)),
  L(2, T(X)),
  L(3, T(Lr)),
  L(4, T(z)),
  L(5, T(G)),
  L(6, Br),
  L(7, j),
  L(8, Wr),
  L(9, We),
  L(10, Dt),
  L(11, T(Ht)),
  L(12, Bt)
], ke);
var ls = Object.defineProperty, as = Object.getOwnPropertyDescriptor, cs = (t, e, r) => e in t ? ls(t, e, { enumerable: !0, configurable: !0, writable: !0, value: r }) : t[e] = r, us = (t, e, r, s) => {
  for (var n = s > 1 ? void 0 : s ? as(e, r) : e, a = t.length - 1, i; a >= 0; a--)
    (i = t[a]) && (n = i(n) || n);
  return n;
}, qe = (t, e) => (r, s) => e(r, s, t), nr = (t, e, r) => cs(t, typeof e != "symbol" ? e + "" : e, r);
const hs = "SHEET_FILTER_UI_PLUGIN";
let He = class extends rt {
  constructor(t = xe, e, r, s) {
    super(), this._config = t, this._injector = e, this._configService = r, this._rpcChannelService = s;
    const { menu: n, ...a } = Ut(
      {},
      xe,
      this._config
    );
    n && this._configService.setConfig("menu", n, { merge: !0 }), this._configService.setConfig(Xt, a);
  }
  onStarting() {
    pr(this._injector, [
      [X],
      [ue],
      [ke]
    ]), this._config.useRemoteFilterValuesGenerator && this._rpcChannelService && this._injector.add([we, {
      useFactory: () => oi(
        this._rpcChannelService.requestChannel(ut)
      )
    }]);
  }
  onReady() {
    pt(this._injector, [
      [ue]
    ]);
  }
  onRendered() {
    pt(this._injector, [
      [ke]
    ]);
  }
};
nr(He, "type", ie.UNIVER_SHEET);
nr(He, "pluginName", hs);
He = us([
  Mt(kt),
  qe(1, T(re)),
  qe(2, xt),
  qe(3, fr(Qt))
], He);
var ds = Object.getOwnPropertyDescriptor, ms = (t, e, r, s) => {
  for (var n = s > 1 ? void 0 : s ? ds(e, r) : e, a = t.length - 1, i; a >= 0; a--)
    (i = t[a]) && (n = i(n) || n);
  return n;
}, Rt = (t, e) => (r, s) => e(r, s, t), Oe;
let At = (Oe = class extends rt {
  constructor(t, e, r) {
    super(), this._config = t, this._injector = e, this._rpcChannelService = r;
  }
  onStarting() {
    [
      [we, { useClass: ze }]
    ].forEach((t) => this._injector.add(t));
  }
  onReady() {
    this._rpcChannelService.registerChannel(
      ut,
      li(this._injector.get(we))
    );
  }
}, v(Oe, "type", ie.UNIVER_SHEET), v(Oe, "pluginName", "SHEET_FILTER_UI_WORKER_PLUGIN"), Oe);
At = ms([
  Rt(1, T(re)),
  Rt(2, Qt)
], At);
export {
  qt as ChangeFilterByOperation,
  Ce as CloseFilterPanelOperation,
  Me as OpenFilterPanelOperation,
  Ue as UniverSheetsFilterMobileUIPlugin,
  He as UniverSheetsFilterUIPlugin,
  At as UniverSheetsFilterUIWorkerPlugin
};
