import { createIdentifier as l, CommandType as P, Inject as m, IUniverInstanceService as C, ICommandService as M, IResourceManagerService as j, Disposable as G, UniverInstanceType as _, generateRandomId as J, DependentOn as V, Injector as H, IConfigService as L, Plugin as T, merge as W } from "@univerjs/core";
import { UnitDrawingService as F, IDrawingManagerService as O, UniverDrawingPlugin as x } from "@univerjs/drawing";
import { SheetInterceptorService as B, RemoveSheetCommand as $, CopySheetCommand as K } from "@univerjs/sheets";
var Y = /* @__PURE__ */ ((e) => (e.Position = "0", e.Both = "1", e.None = "2", e))(Y || {});
class k extends F {
}
const f = l("sheets-drawing.sheet-drawing.service");
var h = /* @__PURE__ */ ((e) => (e[e.INSERT = 0] = "INSERT", e[e.REMOVE = 1] = "REMOVE", e[e.UPDATE = 2] = "UPDATE", e[e.ARRANGE = 3] = "ARRANGE", e[e.GROUP = 4] = "GROUP", e[e.UNGROUP = 5] = "UNGROUP", e))(h || {});
const g = {
  id: "sheet.mutation.set-drawing-apply",
  type: P.MUTATION,
  handler: (e, r) => {
    const t = e.get(O), i = e.get(f), { op: n, unitId: s, subUnitId: o, type: c, objects: a } = r;
    switch (t.applyJson1(s, o, n), i.applyJson1(s, o, n), c) {
      case 0:
        t.addNotification(a), i.addNotification(a);
        break;
      case 1:
        t.removeNotification(a), i.removeNotification(a);
        break;
      case 2:
        t.updateNotification(a), i.updateNotification(a);
        break;
      case 3:
        t.orderNotification(a), i.orderNotification(a);
        break;
      case 4:
        t.groupUpdateNotification(a);
        break;
      case 5:
        t.ungroupUpdateNotification(a);
        break;
    }
    return !0;
  }
};
var q = Object.getOwnPropertyDescriptor, z = (e, r, t, i) => {
  for (var n = i > 1 ? void 0 : i ? q(r, t) : r, s = e.length - 1, o; s >= 0; s--)
    (o = e[s]) && (n = o(n) || n);
  return n;
}, d = (e, r) => (t, i) => r(t, i, e);
const D = "SHEET_DRAWING_PLUGIN";
let I = class extends G {
  constructor(e, r, t, i, n, s) {
    super(), this._sheetInterceptorService = e, this._univerInstanceService = r, this._commandService = t, this._sheetDrawingService = i, this._drawingManagerService = n, this._resourceManagerService = s, this._initSnapshot(), this._initSheetChange(), this.disposeWithMe(this._commandService.registerCommand(g));
  }
  _initSnapshot() {
    const e = (t, i) => {
      const n = i || this._sheetDrawingService.getDrawingDataForUnit(t);
      return n ? JSON.stringify(n) : "";
    }, r = (t) => {
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
        pluginName: D,
        businesses: [_.UNIVER_SHEET],
        toJson: (t, i) => e(t, i),
        parseJson: (t) => r(t),
        onUnLoad: (t) => {
          this._sheetDrawingService.removeDrawingDataForUnit(t), this._drawingManagerService.removeDrawingDataForUnit(t);
        },
        onLoad: (t, i) => {
          this._sheetDrawingService.registerDrawingData(t, i), this._drawingManagerService.registerDrawingData(t, i);
        }
      })
    );
  }
  // eslint-disable-next-line max-lines-per-function
  _initSheetChange() {
    this.disposeWithMe(
      this._sheetInterceptorService.interceptCommand({
        // eslint-disable-next-line max-lines-per-function
        getMutations: (e) => {
          var r;
          if (e.id === $.id) {
            const t = e.params, i = t.unitId || this._univerInstanceService.getCurrentUnitOfType(_.UNIVER_SHEET).getUnitId(), n = t.subUnitId || ((r = this._univerInstanceService.getCurrentUnitOfType(_.UNIVER_SHEET).getActiveSheet()) == null ? void 0 : r.getSheetId());
            if (!i || !n)
              return { redos: [], undos: [] };
            const s = this._sheetDrawingService.getDrawingData(i, n), o = Object.values(s);
            if (o.length === 0)
              return { redos: [], undos: [] };
            const c = this._sheetDrawingService.getBatchRemoveOp(o), { unitId: a, subUnitId: u, undo: v, redo: U, objects: S } = c;
            return {
              redos: [
                {
                  id: g.id,
                  params: {
                    op: U,
                    unitId: a,
                    subUnitId: u,
                    objects: S,
                    type: h.REMOVE
                  }
                }
              ],
              undos: [
                {
                  id: g.id,
                  params: {
                    op: v,
                    unitId: a,
                    subUnitId: u,
                    objects: S,
                    type: h.INSERT
                  }
                }
              ]
            };
          } else if (e.id === K.id) {
            const t = e.params, { unitId: i, subUnitId: n, targetSubUnitId: s } = t;
            if (!i || !n || !s)
              return { redos: [], undos: [] };
            const o = this._sheetDrawingService.getDrawingData(i, n), c = Object.values(o).map((R) => ({
              ...R,
              subUnitId: s,
              drawingId: J(6)
            }));
            if (c.length === 0)
              return { redos: [], undos: [] };
            const a = this._sheetDrawingService.getBatchAddOp(c), { unitId: u, subUnitId: v, undo: U, redo: S, objects: E } = a;
            return {
              redos: [
                {
                  id: g.id,
                  params: {
                    op: S,
                    unitId: u,
                    subUnitId: v,
                    objects: E,
                    type: h.INSERT
                  }
                }
              ],
              undos: [
                {
                  id: g.id,
                  params: {
                    op: U,
                    unitId: u,
                    subUnitId: v,
                    objects: E,
                    type: h.REMOVE
                  }
                }
              ]
            };
          }
          return { redos: [], undos: [] };
        }
      })
    );
  }
};
I = z([
  d(0, m(B)),
  d(1, m(C)),
  d(2, M),
  d(3, f),
  d(4, O),
  d(5, j)
], I);
const A = "sheets-drawing.config", N = {};
var Q = Object.defineProperty, X = Object.getOwnPropertyDescriptor, Z = (e, r, t) => r in e ? Q(e, r, { enumerable: !0, configurable: !0, writable: !0, value: t }) : e[r] = t, y = (e, r, t, i) => {
  for (var n = i > 1 ? void 0 : i ? X(r, t) : r, s = e.length - 1, o; s >= 0; s--)
    (o = e[s]) && (n = o(n) || n);
  return n;
}, b = (e, r) => (t, i) => r(t, i, e), w = (e, r, t) => Z(e, typeof r != "symbol" ? r + "" : r, t);
let p = class extends T {
  constructor(e = N, r, t) {
    super(), this._config = e, this._injector = r, this._configService = t;
    const { ...i } = W(
      {},
      N,
      this._config
    );
    this._configService.setConfig(A, i);
  }
  onStarting() {
    [
      [I],
      [f, { useClass: k }]
    ].forEach((e) => this._injector.add(e)), this._injector.get(I);
  }
};
w(p, "pluginName", D);
w(p, "type", _.UNIVER_SHEET);
p = y([
  V(x),
  b(1, m(H)),
  b(2, L)
], p);
export {
  h as DrawingApplyType,
  f as ISheetDrawingService,
  D as SHEET_DRAWING_PLUGIN,
  g as SetDrawingApplyMutation,
  Y as SheetDrawingAnchorType,
  p as UniverSheetsDrawingPlugin
};
