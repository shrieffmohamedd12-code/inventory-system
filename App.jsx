import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import ... from "./db";
import { supabase, SUPABASE_READY } from "./lib/supabaseClient";
import {
  Plus, ArrowDownToLine, ArrowUpFromLine, Search, Package, History, X,
  AlertTriangle, Users, Upload, LogOut, Shield, UserCircle, FileSpreadsheet,
  Check, Trash2, ArrowRight, Download, Building2, Palette, Boxes,
  Factory, ClipboardList, Lock, Unlock, Pencil, RefreshCw,
} from "lucide-react";

const STORAGE_KEY = "inventory-data-v4";

const WAREHOUSES = [
  { code: "1", name: "dice5,Horaya wh" },
  { code: "2", name: "Helwan wh" },
];

const DEFAULT_FACTORIES = ["Kafr eldawar", "Helwan", "Syria hall", "Joki", "Impact"];

const ORDER_RECEIPT_SIZE = "استلام أوردر";

const DEFAULT_SIZES = ["S2/3", "M4/5", "L6/7", "XL8/9", "XXL10/11", "XXXL12/13", "14/15"];

const emptyState = { users: [], items: [], transactions: [], factories: DEFAULT_FACTORIES, orders: [] };

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}
function todayISO() {
  return new Date().toISOString();
}
function todayDateInput() {
  return new Date().toISOString().slice(0, 10);
}
function fmtDate(iso) {
  const d = new Date(iso);
  return d.toLocaleDateString("ar-EG", { day: "2-digit", month: "2-digit", year: "numeric" }) +
    " - " + d.toLocaleTimeString("ar-EG", { hour: "2-digit", minute: "2-digit" });
}
function whName(code) {
  return WAREHOUSES.find((w) => w.code === code)?.name || code || "—";
}
function colorTotalPieces(color) {
  return (color.sizes || []).reduce((s, x) => s + (Number(x.qty) || 0), 0);
}
function itemTotalPieces(item) {
  return (item.colors || []).reduce((s, c) => s + colorTotalPieces(c), 0);
}
function itemTotalPack(item) {
  return (item.colors || []).reduce((s, c) => s + (Number(c.packQty) || 0), 0);
}

// upgrades older saved data: remaps old 3-warehouse codes to the new merged pair,
// seeds factories/orders, and makes sure every item has a minQty field
function migrateData(base) {
  const remapWarehouse = (code) => {
    if (code === "1" || code === "2") return code;
    if (code === "01" || code === "02") return "1";
    if (code === "03") return "2";
    return "1";
  };
  const items = (base.items || []).map((it) => ({
    ...it,
    warehouseCode: remapWarehouse(it.warehouseCode),
    minQty: it.minQty == null ? null : it.minQty,
  }));
  const transactions = (base.transactions || []).map((t) => ({
    ...t,
    warehouseCode: t.warehouseCode ? remapWarehouse(t.warehouseCode) : t.warehouseCode,
  }));
  const factories = base.factories && base.factories.length ? base.factories : DEFAULT_FACTORIES;
  const orders = (base.orders || []).map((o) => ({ ...o, warehouseCode: remapWarehouse(o.warehouseCode) }));
  return { ...base, items, transactions, factories, orders };
}

function downloadWorkbook(wb, filenamePrefix) {
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${filenamePrefix}-${new Date().toISOString().slice(0, 10)}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function buildWarehouseRows(items, warehouseCode) {
  const whItems = items.filter((i) => i.warehouseCode === warehouseCode);
  const sizeNames = [];
  whItems.forEach((it) =>
    (it.colors || []).forEach((c) =>
      (c.sizes || []).forEach((s) => {
        if (s.size && !sizeNames.includes(s.size)) sizeNames.push(s.size);
      })
    )
  );
  const rows = [];
  whItems.slice().sort((a, b) => a.code.localeCompare(b.code)).forEach((it) => {
    (it.colors || []).forEach((c) => {
      const row = { "الكود": it.code, "اللون": c.color };
      sizeNames.forEach((sn) => {
        const found = (c.sizes || []).find((s) => s.size === sn);
        row[sn] = found ? found.qty : 0;
      });
      row["اجمالي قطعة"] = colorTotalPieces(c);
      row["اجمالي باكنج"] = c.packQty || 0;
      rows.push(row);
    });
  });
  return rows;
}

function exportFullReport(items, transactions) {
  const wb = XLSX.utils.book_new();
  WAREHOUSES.forEach((w) => {
    const rows = buildWarehouseRows(items, w.code);
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, `${w.code} - ${w.name}`);
  });
  const txData = transactions.map((t) => ({
    "التاريخ": fmtDate(t.date),
    "المخزن": whName(t.warehouseCode),
    "الكود": t.code,
    "اللون": t.color || "",
    "المقاس/باكنج": t.kind === "pack" ? "باكنج" : (t.size || ""),
    "النوع": t.type === "in" ? "استلام" : "صرف",
    "الكمية": t.qty,
    "ملاحظة": t.note || "",
    "بواسطة": t.by || "",
  }));
  const wsTx = XLSX.utils.json_to_sheet(txData);
  XLSX.utils.book_append_sheet(wb, wsTx, "سجل الحركات");
  downloadWorkbook(wb, "تقرير-المخازن");
}

function exportItemHistory(item, transactions) {
  const wb = XLSX.utils.book_new();
  const txData = transactions.map((t) => ({
    "التاريخ": fmtDate(t.date),
    "اللون": t.color || "",
    "المقاس/باكنج": t.kind === "pack" ? "باكنج" : (t.size || ""),
    "النوع": t.type === "in" ? "استلام" : "صرف",
    "الكمية": t.qty,
    "ملاحظة": t.note || "",
    "بواسطة": t.by || "",
  }));
  const ws = XLSX.utils.json_to_sheet(txData);
  XLSX.utils.book_append_sheet(wb, ws, "حركات الصنف");
  downloadWorkbook(wb, `حركات-${item.code}`);
}

/* ---------- Orders: Excel helpers ---------- */

function parseOrdersExcelRows(rows, warehouseCode) {
  if (!rows.length) return { rows: [], error: "الملف فارغ" };
  const headerRow = rows[0].map((c) => String(c ?? "").trim());

  const factoryIdx = headerRow.findIndex((h) => /مصنع|factory/i.test(h));
  const dateIdx = headerRow.findIndex((h) => /تاريخ|date/i.test(h));
  const codeIdx = headerRow.findIndex((h) => /كود|code/i.test(h));
  const internalPoIdx = headerRow.findIndex((h) => /internal|داخلي/i.test(h));
  const poIdx = headerRow.findIndex((h, i) => /po/i.test(h) && i !== internalPoIdx);
  const colorIdx = headerRow.findIndex((h) => /لون|color/i.test(h));
  const remainingIdx = headerRow.findIndex((h) => /متبق|باق|remain/i.test(h));
  const requestedIdx = headerRow.findIndex((h) => /مطلوب|طلب|request/i.test(h));

  if (codeIdx === -1 || colorIdx === -1 || factoryIdx === -1) {
    return { rows: [], error: "تعذر العثور على أعمدة المصنع أو الكود أو اللون في الملف" };
  }

  const usedIdxs = new Set([factoryIdx, dateIdx, codeIdx, poIdx, internalPoIdx, colorIdx, remainingIdx, requestedIdx].filter((i) => i !== -1));
  const sizeIdxs = warehouseCode === "2" ? [] : headerRow.map((h, i) => i).filter((i) => !usedIdxs.has(i) && headerRow[i]);

  const dataRows = rows.slice(1).filter((r) => String(r[codeIdx] ?? "").trim());

  const parsed = dataRows.map((r) => {
    const sizes = sizeIdxs.map((i) => ({ size: headerRow[i], qty: Number(r[i]) || 0 })).filter((s) => s.size);
    const sizesSum = sizes.reduce((s, x) => s + x.qty, 0);
    const orderedQty = requestedIdx !== -1 ? (Number(r[requestedIdx]) || 0) : sizesSum;
    const remainingAtFactory = remainingIdx !== -1 ? (Number(r[remainingIdx]) || 0) : orderedQty;
    return {
      factory: String(r[factoryIdx] ?? "").trim(),
      date: String(r[dateIdx] ?? "").trim() || todayDateInput(),
      code: String(r[codeIdx] ?? "").trim(),
      po: poIdx !== -1 ? String(r[poIdx] ?? "").trim() : "",
      internalPo: internalPoIdx !== -1 ? String(r[internalPoIdx] ?? "").trim() : "",
      color: String(r[colorIdx] ?? "").trim() || "بدون لون",
      sizes, orderedQty, remainingAtFactory,
    };
  }).filter((r) => r.code && r.factory);

  return { rows: parsed, error: parsed.length === 0 ? "لم يتم العثور على بيانات صالحة" : null };
}

function exportOrdersReport(orders, factories) {
  const wb = XLSX.utils.book_new();
  factories.forEach((f) => {
    const fOrders = orders.filter((o) => o.factory === f);
    if (fOrders.length === 0) return;
    const rows = fOrders.map((o) => {
      const row = {
        "المصنع": o.factory, "التاريخ": o.date, "المخزن": whName(o.warehouseCode),
        "الكود": o.code, "PO": o.po, "Internal PO": o.internalPo, "اللون": o.color,
      };
      if (o.warehouseCode === "1" && (o.sizes || []).length) {
        o.sizes.forEach((s) => { row[s.size] = s.qty; });
      }
      row["الكمية المطلوبة"] = o.orderedQty;
      row["المستلم حتى الآن"] = o.receivedQty;
      row["المتبقي في المصنع"] = o.remainingAtFactory;
      row["الحالة"] = o.status === "open" ? "مفتوح" : "مغلق";
      return row;
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, f.slice(0, 28) || "مصنع");
  });

  const totalsMap = {};
  orders.forEach((o) => { totalsMap[o.code] = (totalsMap[o.code] || 0) + o.orderedQty; });
  const totalsRows = Object.entries(totalsMap)
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([code, total]) => ({ "الكود": code, "إجمالي المطلوب من كل المصانع": total }));
  const wsTotals = XLSX.utils.json_to_sheet(totalsRows);
  XLSX.utils.book_append_sheet(wb, wsTotals, "إجمالي حسب الكود");

  downloadWorkbook(wb, "تقرير-الأوردرات");
}

export default function InventorySystem() {
  const [data, setData] = useState(emptyState);
  const [loaded, setLoaded] = useState(false);
  const [authChecked, setAuthChecked] = useState(false);
  const [saveError, setSaveError] = useState(false);
  const [busy, setBusy] = useState(false);
  const [view, setView] = useState("items");
  const [search, setSearch] = useState("");
  const [warehouseFilter, setWarehouseFilter] = useState("");
  const [currentUser, setCurrentUser] = useState(null);

  const [showAddItem, setShowAddItem] = useState(false);
  const [showImport, setShowImport] = useState(false);
  const [showDailyRefresh, setShowDailyRefresh] = useState(false);
  const [addColorFor, setAddColorFor] = useState(null);
  const [movementFor, setMovementFor] = useState(null);
  const [detailItemId, setDetailItemId] = useState(null);
  const [editMinQtyFor, setEditMinQtyFor] = useState(null);

  const [orderFactoryFilter, setOrderFactoryFilter] = useState("");
  const [orderSearch, setOrderSearch] = useState("");
  const [showAddOrder, setShowAddOrder] = useState(false);
  const [showImportOrders, setShowImportOrders] = useState(false);
  const [showFactories, setShowFactories] = useState(false);
  const [receiveFor, setReceiveFor] = useState(null);
  const [detailOrderId, setDetailOrderId] = useState(null);

  const isAdmin = currentUser?.role === "admin";

  // reloads everything from Supabase — called after every mutation, and once after login
  const reload = async () => {
    try {
      const fresh = await db.fetchAll();
      setData(fresh);
      setSaveError(false);
    } catch (e) {
      console.error(e);
      setSaveError(true);
    }
  };

  // runs a mutation against Supabase, then refreshes local state from the server
  const run = async (fn) => {
    setBusy(true);
    try {
      await fn();
      await reload();
    } catch (e) {
      console.error(e);
      setSaveError(true);
    } finally {
      setBusy(false);
    }
  };

  // on mount: check if a Supabase session already exists (page refresh case), and load data if so
  useEffect(() => {
    (async () => {
      const profile = await db.getCurrentProfile();
      if (profile) {
        setCurrentUser(profile);
        await reload();
      }
      setAuthChecked(true);
      setLoaded(true);
    })();
  }, []);

  const byName = () => currentUser?.name || "—";

  const addUser = async (name, pin, role) => { await db.createUser(name, pin, role); await reload(); };
  const deleteUserAction = (userId) => run(() => db.deleteUser(userId));

  const addItem = (payload) => run(() => db.addItem(payload, byName()));
  const addColorToItem = (itemId, payload) => run(() => db.addColorToItem(itemId, payload));
  const applyMovement = (item, colorId, kind, sizeId, type, qty, note) =>
    run(() => db.applyMovement(item, colorId, kind, sizeId, type, qty, note, byName()));
  const importFromExcel = (rows, warehouseCode) => run(() => db.importItemsFromExcel(rows, warehouseCode, byName()));
  const setItemMinQty = (itemId, minQty) => run(() => db.setItemMinQty(itemId, minQty));
  const replaceWarehouseFromExcel = (rows, warehouseCode) => run(() => db.replaceWarehouseFromExcel(rows, warehouseCode, byName()));

  const addFactory = (name) => run(() => db.addFactory(name));
  const removeFactory = (name) => run(() => db.removeFactory(name));

  const addOrder = (payload) => run(() => db.addOrder(payload, byName()));
  const importOrders = (rows, warehouseCode) => run(() => db.importOrders(rows, warehouseCode, byName()));
  const setOrderStatus = (orderId, status) => run(() => db.setOrderStatus(orderId, status));
  const receiveOrder = (order, internalPo, sizeBreakdown, totalQty, note) =>
    run(() => db.receiveOrder(order, internalPo, sizeBreakdown, totalQty, note, byName()));

  const filteredItems = useMemo(() => {
    const q = search.trim().toLowerCase();
    return data.items.filter((it) => {
      if (warehouseFilter && it.warehouseCode !== warehouseFilter) return false;
      if (!q) return true;
      if (it.code.toLowerCase().includes(q)) return true;
      if ((it.name || "").toLowerCase().includes(q)) return true;
      return (it.colors || []).some((c) => (c.color || "").toLowerCase().includes(q));
    });
  }, [data.items, search, warehouseFilter]);

  const totalCodes = filteredItems.length;
  const totalPieces = filteredItems.reduce((s, it) => s + itemTotalPieces(it), 0);
  const totalPack = filteredItems.reduce((s, it) => s + itemTotalPack(it), 0);
  const lowStockItems = filteredItems.filter((it) => it.minQty != null && itemTotalPieces(it) <= it.minQty);

  const filteredOrders = useMemo(() => {
    const q = orderSearch.trim().toLowerCase();
    return (data.orders || []).filter((o) => {
      if (orderFactoryFilter && o.factory !== orderFactoryFilter) return false;
      if (!q) return true;
      return o.code.toLowerCase().includes(q) || (o.color || "").toLowerCase().includes(q) || (o.po || "").toLowerCase().includes(q) || (o.internalPo || "").toLowerCase().includes(q);
    });
  }, [data.orders, orderSearch, orderFactoryFilter]);

  if (!SUPABASE_READY) {
    return (
      <div dir="rtl" style={styles.page}>
        <GlobalStyle />
        <div style={styles.loginWrap}>
          <div style={styles.loginCard}>
            <div style={styles.loginMark}>⚠️</div>
            <div style={styles.loginTitle}>الإعداد غير مكتمل</div>
            <div style={styles.loginSub}>
              لم يتم ضبط بيانات الاتصال بـ Supabase. أنشئ ملف <b>.env</b> في جذر المشروع
              (انسخه من .env.example) وضع فيه VITE_SUPABASE_URL و VITE_SUPABASE_ANON_KEY الخاصين بمشروعك،
              ثم أعد تشغيل المشروع.
            </div>
          </div>
        </div>
      </div>
    );
  }

  if (!loaded || !authChecked) {
    return <div dir="rtl" style={styles.page}><GlobalStyle /><div style={styles.loading}>جاري التحميل...</div></div>;
  }

  const handleLogout = async () => { await db.signOut(); setCurrentUser(null); setData(emptyState); };

  if (!currentUser) {
    return (
      <div dir="rtl" style={styles.page}>
        <GlobalStyle />
        <LoginScreen
          fetchUsers={db.fetchProfiles}
          onLogin={async (email, pin) => { await db.signInByName(email, pin); const p = await db.getCurrentProfile(); setCurrentUser(p); await reload(); }}
          onCreateFirstAdmin={async (name, pin) => { const u = await db.signUpFirstAdmin(name, pin); setCurrentUser(u); await reload(); }}
        />
      </div>
    );
  }

  const detailItem = detailItemId ? data.items.find((i) => i.id === detailItemId) : null;

  return (
    <div dir="rtl" style={styles.page}>
      <GlobalStyle />
      <header style={styles.header}>
        <div style={styles.headerInner}>
          <div style={styles.brand}>
            <div style={styles.brandMark}>م</div>
            <div>
              <div style={styles.brandTitle}>نظام المخازن</div>
              <div style={styles.brandSub}>إدارة الأصناف والحركات</div>
            </div>
          </div>
          <div style={styles.headerRight}>
            <nav style={styles.nav}>
              <button onClick={() => { setView("items"); setDetailItemId(null); }} style={{ ...styles.navBtn, ...(view === "items" ? styles.navBtnActive : {}) }}>
                <Package size={16} /><span className="hide-narrow">الأصناف</span>
              </button>
              <button onClick={() => setView("log")} style={{ ...styles.navBtn, ...(view === "log" ? styles.navBtnActive : {}) }}>
                <History size={16} /><span className="hide-narrow">سجل الحركات</span>
              </button>
              {isAdmin && (
                <button onClick={() => setView("orders")} style={{ ...styles.navBtn, ...(view === "orders" ? styles.navBtnActive : {}) }}>
                  <ClipboardList size={16} /><span className="hide-narrow">الأوردرات</span>
                </button>
              )}
              {isAdmin && (
                <button onClick={() => setView("users")} style={{ ...styles.navBtn, ...(view === "users" ? styles.navBtnActive : {}) }}>
                  <Users size={16} /><span className="hide-narrow">المستخدمين</span>
                </button>
              )}
            </nav>
            <div style={styles.userBadge}>
              <div style={{ ...styles.roleDot, background: isAdmin ? "#D9A441" : "#7C8591" }} />
              <span style={styles.userName}>{currentUser.name}</span>
              <button style={styles.logoutBtn} onClick={handleLogout}><LogOut size={15} /></button>
            </div>
          </div>
        </div>
      </header>

      <main style={styles.main}>
        {view === "items" && (
          detailItem ? (
            <ItemDetailView
              item={detailItem}
              transactions={data.transactions.filter((t) => t.itemId === detailItem.id)}
              onBack={() => setDetailItemId(null)}
              onMove={(type, presetColorId, presetKind, presetSizeId) => setMovementFor({ item: detailItem, type, presetColorId, presetKind, presetSizeId })}
              onAddColor={() => setAddColorFor(detailItem)}
              onEditMinQty={isAdmin ? () => setEditMinQtyFor(detailItem) : null}
            />
          ) : (
            <>
              <div style={styles.statsRow}>
                <StatCard label="عدد الأصناف" value={totalCodes} accent="#D9A441" />
                <StatCard label="إجمالي القطع" value={totalPieces} accent="#3D7A4E" />
                <StatCard label="إجمالي الباكنج" value={totalPack} accent="#4A6FA5" />
                <StatCard label="تحت الحد الأدنى" value={lowStockItems.length} accent="#B5493A" warn={lowStockItems.length > 0} />
              </div>

              <div style={styles.toolbar}>
                <select style={styles.whSelect} value={warehouseFilter} onChange={(e) => setWarehouseFilter(e.target.value)}>
                  <option value="">كل المخازن</option>
                  {WAREHOUSES.map((w) => <option key={w.code} value={w.code}>{w.code} — {w.name}</option>)}
                </select>
                <div style={styles.searchBox}>
                  <Search size={17} color="#8A8F98" style={{ flexShrink: 0 }} />
                  <input placeholder="ابحث بالكود أو اللون..." value={search} onChange={(e) => setSearch(e.target.value)} style={styles.searchInput} />
                </div>
                {isAdmin && (
                  <>
                    <button style={styles.secondaryBtn} onClick={() => setShowImport(true)}><FileSpreadsheet size={17} />استيراد Excel</button>
                    <button style={styles.secondaryBtn} onClick={() => setShowDailyRefresh(true)}><RefreshCw size={17} />تحديث يومي للرصيد</button>
                    <button style={styles.primaryBtn} onClick={() => setShowAddItem(true)}><Plus size={17} />إضافة صنف</button>
                  </>
                )}
                <button style={styles.secondaryBtn} onClick={() => exportFullReport(data.items, data.transactions)}><Download size={17} />تصدير Excel</button>
              </div>

              {filteredItems.length === 0 ? (
                <EmptyState hasSearch={!!search || !!warehouseFilter} onAdd={isAdmin ? () => setShowAddItem(true) : null} />
              ) : (
                <div style={styles.itemsGrid}>
                  {filteredItems.slice().sort((a, b) => a.code.localeCompare(b.code)).map((item) => (
                    <ItemCard
                      key={item.id}
                      item={item}
                      showWarehouse={!warehouseFilter}
                      onOpen={() => setDetailItemId(item.id)}
                      onIn={() => setMovementFor({ item, type: "in" })}
                      onOut={() => setMovementFor({ item, type: "out" })}
                    />
                  ))}
                </div>
              )}
            </>
          )
        )}

        {view === "log" && <TransactionLog transactions={data.transactions} showWarehouse />}

        {view === "orders" && isAdmin && (
          <OrdersScreen
            orders={filteredOrders}
            factories={data.factories || []}
            factoryFilter={orderFactoryFilter}
            onFactoryFilter={setOrderFactoryFilter}
            search={orderSearch}
            onSearch={setOrderSearch}
            onAdd={() => setShowAddOrder(true)}
            onImport={() => setShowImportOrders(true)}
            onExport={() => exportOrdersReport(data.orders || [], data.factories || [])}
            onManageFactories={() => setShowFactories(true)}
            onOpen={(id) => setDetailOrderId(id)}
            detailOrder={detailOrderId ? (data.orders || []).find((o) => o.id === detailOrderId) : null}
            onBack={() => setDetailOrderId(null)}
            onReceive={(order) => setReceiveFor(order)}
            onSetStatus={setOrderStatus}
          />
        )}

        {view === "users" && isAdmin && (
          <UsersPanel users={data.users} currentUser={currentUser} onAdd={addUser} onDelete={deleteUserAction} />
        )}
      </main>

      {busy && <div style={styles.busyBanner}>جاري الحفظ...</div>}
      {saveError && <div style={styles.saveError}><AlertTriangle size={16} />تعذر حفظ البيانات، حاول مرة أخرى</div>}

      {showAddItem && (
        <AddItemModal
          onClose={() => setShowAddItem(false)}
          onSave={(payload) => { addItem(payload); setShowAddItem(false); }}
          existingItems={data.items}
          defaultWarehouse={warehouseFilter}
        />
      )}

      {addColorFor && (
        <AddColorModal
          item={addColorFor}
          onClose={() => setAddColorFor(null)}
          onSave={(payload) => { addColorToItem(addColorFor.id, payload); setAddColorFor(null); }}
        />
      )}

      {movementFor && (
        <MovementModal
          item={movementFor.item}
          type={movementFor.type}
          presetColorId={movementFor.presetColorId}
          presetKind={movementFor.presetKind}
          presetSizeId={movementFor.presetSizeId}
          onClose={() => setMovementFor(null)}
          onConfirm={(colorId, kind, sizeId, qty, note) => {
            applyMovement(movementFor.item, colorId, kind, sizeId, movementFor.type, qty, note);
            setMovementFor(null);
          }}
        />
      )}

      {showImport && (
        <ImportExcelModal onClose={() => setShowImport(false)} onConfirm={(rows, wh) => { importFromExcel(rows, wh); setShowImport(false); }} defaultWarehouse={warehouseFilter} />
      )}

      {editMinQtyFor && (
        <EditMinQtyModal
          item={editMinQtyFor}
          onClose={() => setEditMinQtyFor(null)}
          onSave={(val) => { setItemMinQty(editMinQtyFor.id, val); setEditMinQtyFor(null); }}
        />
      )}

      {showDailyRefresh && (
        <DailyRefreshModal
          onClose={() => setShowDailyRefresh(false)}
          onConfirm={(rows, wh) => { replaceWarehouseFromExcel(rows, wh); setShowDailyRefresh(false); }}
        />
      )}

      {showAddOrder && (
        <OrderFormModal
          factories={data.factories || []}
          onClose={() => setShowAddOrder(false)}
          onSave={(payload) => { addOrder(payload); setShowAddOrder(false); }}
        />
      )}

      {showImportOrders && (
        <ImportOrdersModal
          onClose={() => setShowImportOrders(false)}
          onConfirm={(rows, wh) => { importOrders(rows, wh); setShowImportOrders(false); }}
        />
      )}

      {showFactories && (
        <FactoriesModal
          factories={data.factories || []}
          onAdd={addFactory}
          onRemove={removeFactory}
          onClose={() => setShowFactories(false)}
        />
      )}

      {receiveFor && (
        <ReceiveOrderModal
          order={receiveFor}
          onClose={() => setReceiveFor(null)}
          onConfirm={(internalPo, sizeBreakdown, totalQty, note) => {
            receiveOrder(receiveFor, internalPo, sizeBreakdown, totalQty, note);
            setReceiveFor(null);
          }}
        />
      )}
    </div>
  );
}

function GlobalStyle() {
  return (
    <style>{`
      * { box-sizing: border-box; }
      body { margin: 0; }
      ::selection { background: #D9A441; color: #1B1F24; }
      button { font-family: inherit; }
      input, textarea, select { font-family: inherit; }
      @media (max-width: 640px) { .hide-narrow { display: none !important; } }
    `}</style>
  );
}

/* ---------- Login (wired to Supabase Auth) ---------- */

function LoginScreen({ fetchUsers, onLogin, onCreateFirstAdmin }) {
  const [users, setUsers] = useState(null); // null = still loading
  const [loadError, setLoadError] = useState("");

  useEffect(() => {
    (async () => {
      try { setUsers(await fetchUsers()); }
      catch (e) { setLoadError("تعذر الاتصال بقاعدة البيانات، تأكد من إعدادات Supabase"); }
    })();
  }, []);

  if (loadError) {
    return (
      <div style={styles.loginWrap}>
        <div style={styles.loginCard}>
          <div style={styles.loginMark}>⚠️</div>
          <div style={styles.loginTitle}>خطأ في الاتصال</div>
          <div style={styles.loginSub}>{loadError}</div>
        </div>
      </div>
    );
  }
  if (users === null) {
    return <div style={styles.loading}>جاري التحميل...</div>;
  }
  if (users.length === 0) return <FirstAdminSetup onCreate={onCreateFirstAdmin} />;

  const [picked, setPicked] = useState(null);
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const submit = async () => {
    if (!picked || busy) return;
    setBusy(true);
    setError("");
    try {
      await onLogin(picked.email, pin);
    } catch (e) {
      setError("الرقم السري غير صحيح");
    } finally {
      setBusy(false);
    }
  };

  if (picked) {
    return (
      <div style={styles.loginWrap}>
        <div style={styles.loginCard}>
          <div style={styles.loginMark}>م</div>
          <div style={styles.loginTitle}>أهلاً {picked.name}</div>
          <div style={styles.loginSub}>أدخل الرقم السري الخاص بك</div>
          <input style={styles.pinInput} type="password" inputMode="numeric" maxLength={6} value={pin}
            onChange={(e) => { setPin(e.target.value.replace(/\D/g, "")); setError(""); }}
            onKeyDown={(e) => e.key === "Enter" && submit()} placeholder="••••••" autoFocus />
          {error && <div style={styles.errorText}>{error}</div>}
          <button style={{ ...styles.primaryBtn, width: "100%", justifyContent: "center", marginTop: 6, opacity: busy ? 0.6 : 1 }} onClick={submit} disabled={busy}>
            {busy ? "جاري الدخول..." : "دخول"}
          </button>
          <button style={styles.linkBtn} onClick={() => { setPicked(null); setPin(""); setError(""); }}>العودة لاختيار المستخدم</button>
        </div>
      </div>
    );
  }

  return (
    <div style={styles.loginWrap}>
      <div style={styles.loginCard}>
        <div style={styles.loginMark}>م</div>
        <div style={styles.loginTitle}>نظام المخازن</div>
        <div style={styles.loginSub}>اختر اسمك لتسجيل الدخول</div>
        <div style={styles.userList}>
          {users.map((u) => (
            <button key={u.id} style={styles.userListItem} onClick={() => setPicked(u)}>
              <div style={{ ...styles.roleDot, background: u.role === "admin" ? "#D9A441" : "#7C8591" }} />
              <span style={styles.userListName}>{u.name}</span>
              <span style={styles.userListRole}>{u.role === "admin" ? "مسؤول" : "موظف مخزن"}</span>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

function FirstAdminSetup({ onCreate }) {
  const [name, setName] = useState("");
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const submit = async () => {
    if (!name.trim()) return setError("أدخل اسمك");
    if (pin.length < 6) return setError("يجب أن يتكوّن الرقم السري من 6 أرقام على الأقل (يشترطه نظام الحماية في Supabase)");
    setBusy(true);
    setError("");
    try {
      await onCreate(name.trim(), pin);
    } catch (e) {
      setError(e.message || "حدث خطأ، تأكد من إعدادات Supabase (خصوصًا تعطيل تأكيد الإيميل)");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div style={styles.loginWrap}>
      <div style={styles.loginCard}>
        <div style={styles.loginMark}>م</div>
        <div style={styles.loginTitle}>ابدأ نظام المخازن</div>
        <div style={styles.loginSub}>أنشئ حساب المسؤول الأول لتتمكن من إضافة باقي الموظفين لاحقًا</div>
        <Field label="اسمك"><input style={styles.input} value={name} onChange={(e) => setName(e.target.value)} placeholder="مثال: أحمد" autoFocus /></Field>
        <Field label="رقم سري (6 أرقام على الأقل)"><input style={styles.input} type="password" inputMode="numeric" value={pin} onChange={(e) => setPin(e.target.value.replace(/\D/g, ""))} placeholder="••••••" /></Field>
        {error && <div style={styles.errorText}>{error}</div>}
        <button style={{ ...styles.primaryBtn, width: "100%", justifyContent: "center", marginTop: 8, opacity: busy ? 0.6 : 1 }} onClick={submit} disabled={busy}>
          {busy ? "جاري الإنشاء..." : "إنشاء الحساب والدخول"}
        </button>
      </div>
    </div>
  );
}

/* ---------- Users panel (wired to Supabase Auth admin actions) ---------- */

function UsersPanel({ users, currentUser, onAdd, onDelete }) {
  const [showAdd, setShowAdd] = useState(false);
  const adminCount = users.filter((u) => u.role === "admin").length;

  return (
    <div>
      <div style={styles.toolbar}>
        <div style={{ flex: 1 }} />
        <button style={styles.primaryBtn} onClick={() => setShowAdd(true)}><Plus size={17} />إضافة مستخدم</button>
      </div>
      <div style={styles.logList}>
        {users.map((u) => {
          const isLastAdmin = u.role === "admin" && adminCount <= 1;
          return (
            <div key={u.id} style={styles.logRow}>
              <div style={{ ...styles.logIcon, background: u.role === "admin" ? "#FBF2DF" : "#EEF0F3", color: u.role === "admin" ? "#B8862F" : "#5C6470" }}>
                {u.role === "admin" ? <Shield size={15} /> : <UserCircle size={15} />}
              </div>
              <div style={styles.logMid}>
                <div style={styles.logTitle}>{u.name} {u.id === currentUser.id && <span style={styles.youTag}>(انت)</span>}</div>
                <div style={styles.logNote}>{u.role === "admin" ? "مسؤول — صلاحيات كاملة" : "موظف مخزن — صرف واستلام فقط"}</div>
              </div>
              {u.id !== currentUser.id && !isLastAdmin && (
                <button style={styles.deleteBtn} onClick={() => onDelete(u.id)}><Trash2 size={15} /></button>
              )}
            </div>
          );
        })}
      </div>
      {showAdd && (
        <AddUserModal existingNames={users.map((u) => u.name.toLowerCase())} onClose={() => setShowAdd(false)} onSave={onAdd} />
      )}
    </div>
  );
}

function AddUserModal({ onClose, onSave, existingNames }) {
  const [name, setName] = useState("");
  const [pin, setPin] = useState("");
  const [role, setRole] = useState("staff");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const submit = async () => {
    if (!name.trim()) return setError("أدخل اسم المستخدم");
    if (existingNames.includes(name.trim().toLowerCase())) return setError("هذا الاسم مستخدم بالفعل");
    if (pin.length < 6) return setError("يجب أن يتكوّن الرقم السري من 6 أرقام على الأقل");
    setBusy(true);
    setError("");
    try {
      await onSave(name.trim(), pin, role);
      onClose();
    } catch (e) {
      setError(e.message || "تعذر إنشاء المستخدم");
    } finally {
      setBusy(false);
    }
  };

  return (
    <Modal onClose={onClose} title="إضافة مستخدم جديد">
      <Field label="الاسم"><input style={styles.input} value={name} onChange={(e) => setName(e.target.value)} placeholder="اسم الموظف" autoFocus /></Field>
      <Field label="رقم سري (6 أرقام على الأقل)"><input style={styles.input} type="password" inputMode="numeric" value={pin} onChange={(e) => setPin(e.target.value.replace(/\D/g, ""))} placeholder="••••••" /></Field>
      <Field label="الصلاحية">
        <div style={styles.roleToggle}>
          <button style={{ ...styles.roleOption, ...(role === "staff" ? styles.roleOptionActive : {}) }} onClick={() => setRole("staff")}><UserCircle size={15} />موظف مخزن</button>
          <button style={{ ...styles.roleOption, ...(role === "admin" ? styles.roleOptionActive : {}) }} onClick={() => setRole("admin")}><Shield size={15} />مسؤول</button>
        </div>
        <div style={styles.roleHint}>{role === "admin" ? "يمكنه إضافة الأصناف والمستخدمين واستيراد ملفات Excel" : "يمكنه فقط صرف واستلام الأصناف الموجودة"}</div>
      </Field>
      {error && <div style={styles.errorText}>{error}</div>}
      <button style={{ ...styles.primaryBtn, width: "100%", justifyContent: "center", marginTop: 8, opacity: busy ? 0.6 : 1 }} onClick={submit} disabled={busy}>
        {busy ? "جاري الحفظ..." : "حفظ المستخدم"}
      </button>
    </Modal>
  );
}

/* ---------- Items ---------- */

function StatCard({ label, value, accent }) {
  return (
    <div style={{ ...styles.statCard, borderInlineStart: `3px solid ${accent}` }}>
      <div style={styles.statValue}>{value}</div>
      <div style={styles.statLabel}>{label}</div>
    </div>
  );
}

function ItemCard({ item, showWarehouse, onOpen, onIn, onOut }) {
  const pieces = itemTotalPieces(item);
  const pack = itemTotalPack(item);
  const low = item.minQty != null && pieces <= item.minQty;
  return (
    <div style={styles.itemCard} onClick={onOpen} role="button" tabIndex={0}>
      <div style={styles.itemTop}>
        <div>
          <div style={styles.itemCode}>{item.code}</div>
          <div style={styles.itemName}>{item.name}</div>
        </div>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 4 }}>
          {showWarehouse && (
            <div style={styles.whBadge}><Building2 size={11} />{item.warehouseCode} - {whName(item.warehouseCode)}</div>
          )}
          {low && <div style={styles.lowBadge}><AlertTriangle size={11} />تحت الحد الأدنى</div>}
        </div>
      </div>
      <div style={styles.colorChipsRow}>
        <Palette size={13} color="#8A8F98" />
        {item.colors.map((c) => <span key={c.id} style={styles.colorChip}>{c.color}</span>)}
      </div>
      <div style={styles.itemQtyRow}>
        <div><span style={styles.itemQty}>{pieces}</span><span style={styles.itemUnit}> قطعة</span></div>
        <div><span style={styles.itemQtyPack}>{pack}</span><span style={styles.itemUnit}> باكنج</span></div>
      </div>
      <div style={styles.itemActions}>
        <button style={styles.inBtn} onClick={(e) => { e.stopPropagation(); onIn(); }}><ArrowDownToLine size={15} />استلام</button>
        <button style={styles.outBtn} onClick={(e) => { e.stopPropagation(); onOut(); }}><ArrowUpFromLine size={15} />صرف</button>
      </div>
    </div>
  );
}

function ItemDetailView({ item, transactions, onBack, onMove, onAddColor, onEditMinQty }) {
  const pieces = itemTotalPieces(item);
  const pack = itemTotalPack(item);
  const low = item.minQty != null && pieces <= item.minQty;

  return (
    <div>
      <button style={styles.backBtn} onClick={onBack}><ArrowRight size={16} />العودة إلى جميع الأصناف</button>

      <div style={styles.detailHeaderCard}>
        <div style={styles.detailTopRow}>
          <div>
            <div style={styles.itemCode}>{item.code}</div>
            <div style={styles.detailName}>{item.name}</div>
          </div>
          <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 6 }}>
            <div style={styles.whBadge}><Building2 size={11} />{item.warehouseCode} - {whName(item.warehouseCode)}</div>
            {low && <div style={styles.lowBadge}><AlertTriangle size={11} />تحت الحد الأدنى</div>}
          </div>
        </div>
        <div style={styles.detailQtyRow}>
          <span style={styles.detailQty}>{pieces}</span><span style={styles.itemUnit}>قطعة</span>
          <span style={{ marginInlineStart: 18, color: "#4A6FA5", fontSize: 24, fontWeight: 800 }}>{pack}</span>
          <span style={styles.itemUnit}>باكنج</span>
        </div>
        <div style={styles.minQtyRow}>
          حد التنبيه: <b>{item.minQty == null ? "غير محدد" : item.minQty}</b> قطعة
          {onEditMinQty && (
            <button style={styles.editMinQtyBtn} onClick={onEditMinQty}><Pencil size={12} />تعديل</button>
          )}
        </div>
        <div style={styles.detailActions}>
          <button style={styles.inBtn} onClick={() => onMove("in")}><ArrowDownToLine size={15} />استلام</button>
          <button style={styles.outBtn} onClick={() => onMove("out")}><ArrowUpFromLine size={15} />صرف</button>
          <button style={styles.secondaryBtn} onClick={onAddColor}><Plus size={15} />إضافة لون</button>
        </div>
      </div>

      {item.colors.map((c) => (
        <div key={c.id} style={styles.colorSection}>
          <div style={styles.colorSectionHeader}>
            <div style={styles.colorSectionTitle}><Palette size={14} />{c.color}</div>
            <div style={styles.colorSectionPack}>
              <Boxes size={13} /> باكنج: <b>{c.packQty}</b>
              <button style={styles.miniBtn} onClick={() => onMove("in", c.id, "pack")}>+</button>
              <button style={styles.miniBtn} onClick={() => onMove("out", c.id, "pack")}>−</button>
            </div>
          </div>
          <div style={styles.sizeTable}>
            {c.sizes.map((s) => (
              <div key={s.id} style={styles.sizeRow}>
                <span style={styles.sizeName}>{s.size}</span>
                <span style={styles.sizeQty}>{s.qty}</span>
                <button style={styles.miniBtnIn} onClick={() => onMove("in", c.id, "size", s.id)}><ArrowDownToLine size={13} /></button>
                <button style={styles.miniBtnOut} onClick={() => onMove("out", c.id, "size", s.id)}><ArrowUpFromLine size={13} /></button>
              </div>
            ))}
          </div>
        </div>
      ))}

      <div style={styles.detailLogRow}>
        <div style={styles.detailLogTitle}>حركات هذا الصنف</div>
        <button style={styles.exportSmallBtn} onClick={() => exportItemHistory(item, transactions)}><Download size={14} />تصدير Excel</button>
      </div>
      <TransactionLog transactions={transactions} />
    </div>
  );
}

function TransactionLog({ transactions, showWarehouse }) {
  if (transactions.length === 0) return <EmptyState hasSearch={false} onAdd={null} message="لا توجد حركات بعد" />;
  return (
    <div style={styles.logList}>
      {transactions.map((tx) => (
        <div key={tx.id} style={styles.logRow}>
          <div style={{ ...styles.logIcon, background: tx.type === "in" ? "#EAF4EC" : "#FBEEEA", color: tx.type === "in" ? "#3D7A4E" : "#B5493A" }}>
            {tx.type === "in" ? <ArrowDownToLine size={15} /> : <ArrowUpFromLine size={15} />}
          </div>
          <div style={styles.logMid}>
            <div style={styles.logTitle}>
              {tx.code} {tx.color && `· ${tx.color}`} {tx.kind === "pack" ? "· باكنج" : (tx.size && tx.size !== "—" ? `· مقاس ${tx.size}` : "")}
              {showWarehouse && <span style={styles.logWh}> ({whName(tx.warehouseCode)})</span>}
            </div>
            <div style={styles.logNote}>{tx.note || (tx.type === "in" ? "استلام" : "صرف")} — بواسطة {tx.by || "—"}</div>
          </div>
          <div style={styles.logRight}>
            <div style={{ ...styles.logQty, color: tx.type === "in" ? "#3D7A4E" : "#B5493A" }}>{tx.type === "in" ? "+" : "−"}{tx.qty}</div>
            <div style={styles.logDate}>{fmtDate(tx.date)}</div>
          </div>
        </div>
      ))}
    </div>
  );
}

function EmptyState({ hasSearch, onAdd, message }) {
  return (
    <div style={styles.empty}>
      <div style={styles.emptyIcon}><Package size={28} color="#B8BCC4" /></div>
      <div style={styles.emptyTitle}>{message || (hasSearch ? "لا توجد نتائج مطابقة" : "لم تتم إضافة أي أصناف بعد")}</div>
      {!hasSearch && onAdd && <button style={styles.primaryBtn} onClick={onAdd}><Plus size={17} />إضافة أول صنف</button>}
    </div>
  );
}

/* ---------- Size rows editor (shared by Add Item / Add Color) ---------- */

function SizeRowsEditor({ rows, setRows }) {
  const updateRow = (i, field, val) => setRows(rows.map((r, idx) => (idx === i ? { ...r, [field]: val } : r)));
  const removeRow = (i) => setRows(rows.filter((_, idx) => idx !== i));
  const addRow = () => setRows([...rows, { size: "", qty: "" }]);

  return (
    <div style={styles.field}>
      <label style={styles.fieldLabel}>المقاسات والكميات (بالقطعة)</label>
      {rows.map((r, i) => (
        <div key={i} style={styles.sizeEditRow}>
          <input style={{ ...styles.input, flex: 2 }} value={r.size} onChange={(e) => updateRow(i, "size", e.target.value)} placeholder="المقاس" />
          <input style={{ ...styles.input, flex: 1 }} type="number" value={r.qty} onChange={(e) => updateRow(i, "qty", e.target.value)} placeholder="الكمية" />
          <button style={styles.removeRowBtn} onClick={() => removeRow(i)}><X size={14} /></button>
        </div>
      ))}
      <button style={styles.addRowBtn} onClick={addRow}><Plus size={13} />إضافة مقاس</button>
    </div>
  );
}

function AddItemModal({ onClose, onSave, existingItems, defaultWarehouse }) {
  const [warehouseCode, setWarehouseCode] = useState(defaultWarehouse || WAREHOUSES[0].code);
  const [code, setCode] = useState("");
  const [name, setName] = useState("");
  const [color, setColor] = useState("");
  const [rows, setRows] = useState(DEFAULT_SIZES.map((s) => ({ size: s, qty: "" })));
  const [packQty, setPackQty] = useState("");
  const [error, setError] = useState("");

  const submit = () => {
    if (!code.trim() || !color.trim()) return setError("أدخل الكود واللون على الأقل");
    const dup = existingItems.some((it) => it.warehouseCode === warehouseCode && it.code === code.trim());
    if (dup) return setError("هذا الكود موجود بالفعل في هذا المخزن — استخدم زر (إضافة لون) من داخل الصنف بدل ذلك");
    const sizes = rows.filter((r) => r.size.trim()).map((r) => ({ size: r.size.trim(), qty: r.qty === "" ? 0 : Number(r.qty) || 0 }));
    onSave({
      warehouseCode, code: code.trim(), name: name.trim(), color: color.trim(),
      sizes, packQty: packQty === "" ? 0 : Number(packQty) || 0,
    });
  };

  return (
    <Modal onClose={onClose} title="إضافة صنف جديد">
      <Field label="المخزن">
        <select style={styles.input} value={warehouseCode} onChange={(e) => setWarehouseCode(e.target.value)}>
          {WAREHOUSES.map((w) => <option key={w.code} value={w.code}>{w.code} — {w.name}</option>)}
        </select>
      </Field>
      <div style={styles.fieldRow}>
        <Field label="كود الصنف"><input style={styles.input} value={code} onChange={(e) => setCode(e.target.value)} placeholder="مثال: DM235" autoFocus /></Field>
        <Field label="اللون"><input style={styles.input} value={color} onChange={(e) => setColor(e.target.value)} placeholder="مثال: أبيض" /></Field>
      </div>
      <Field label="اسم الصنف (اختياري)"><input style={styles.input} value={name} onChange={(e) => setName(e.target.value)} placeholder="وصف الصنف" /></Field>
      <SizeRowsEditor rows={rows} setRows={setRows} />
      <Field label="الكمية بالباكنج"><input style={styles.input} type="number" value={packQty} onChange={(e) => setPackQty(e.target.value)} placeholder="0" /></Field>
      {error && <div style={styles.errorText}>{error}</div>}
      <button style={{ ...styles.primaryBtn, width: "100%", justifyContent: "center", marginTop: 8 }} onClick={submit}>حفظ الصنف</button>
    </Modal>
  );
}

function AddColorModal({ item, onClose, onSave }) {
  const [color, setColor] = useState("");
  const [rows, setRows] = useState(DEFAULT_SIZES.map((s) => ({ size: s, qty: "" })));
  const [packQty, setPackQty] = useState("");
  const [error, setError] = useState("");

  const submit = () => {
    if (!color.trim()) return setError("أدخل اللون");
    if (item.colors.some((c) => c.color.trim() === color.trim())) return setError("هذا اللون موجود بالفعل لهذا الصنف");
    const sizes = rows.filter((r) => r.size.trim()).map((r) => ({ size: r.size.trim(), qty: r.qty === "" ? 0 : Number(r.qty) || 0 }));
    onSave({ color: color.trim(), sizes, packQty: packQty === "" ? 0 : Number(packQty) || 0 });
  };

  return (
    <Modal onClose={onClose} title={`إضافة لون جديد — ${item.code}`}>
      <Field label="اللون"><input style={styles.input} value={color} onChange={(e) => setColor(e.target.value)} placeholder="مثال: كحلي" autoFocus /></Field>
      <SizeRowsEditor rows={rows} setRows={setRows} />
      <Field label="الكمية بالباكنج"><input style={styles.input} type="number" value={packQty} onChange={(e) => setPackQty(e.target.value)} placeholder="0" /></Field>
      {error && <div style={styles.errorText}>{error}</div>}
      <button style={{ ...styles.primaryBtn, width: "100%", justifyContent: "center", marginTop: 8 }} onClick={submit}>حفظ اللون</button>
    </Modal>
  );
}

/* ---------- Movement ---------- */

function MovementModal({ item, type, presetColorId, presetKind, presetSizeId, onClose, onConfirm }) {
  const [colorId, setColorId] = useState(presetColorId || item.colors[0]?.id);
  const [kind, setKind] = useState(presetKind || "size");
  const selectedColor = item.colors.find((c) => c.id === colorId) || item.colors[0];
  const [sizeId, setSizeId] = useState(presetSizeId || selectedColor?.sizes[0]?.id);
  const [qty, setQty] = useState("");
  const [note, setNote] = useState("");
  const [error, setError] = useState("");
  const isOut = type === "out";

  const currentSize = selectedColor?.sizes.find((s) => s.id === sizeId);
  const currentAvailable = kind === "pack" ? (selectedColor?.packQty || 0) : (currentSize?.qty || 0);

  const changeColor = (id) => {
    setColorId(id);
    const c = item.colors.find((x) => x.id === id);
    setSizeId(c?.sizes[0]?.id);
  };

  const submit = () => {
    const q = Number(qty);
    if (!qty || isNaN(q) || q <= 0) return setError("أدخل كمية صحيحة أكبر من صفر");
    if (kind === "size" && !sizeId) return setError("اختر المقاس");
    if (isOut && q > currentAvailable) return setError(`الرصيد الحالي ${currentAvailable} فقط، لا يمكن صرف كمية أكبر منه`);
    onConfirm(colorId, kind, kind === "size" ? sizeId : null, q, note.trim());
  };

  return (
    <Modal onClose={onClose} title={`${isOut ? "صرف" : "استلام"} — ${item.code}`}>
      {item.colors.length > 1 && !presetColorId && (
        <Field label="اللون">
          <select style={styles.input} value={colorId} onChange={(e) => changeColor(e.target.value)}>
            {item.colors.map((c) => <option key={c.id} value={c.id}>{c.color}</option>)}
          </select>
        </Field>
      )}
      {(item.colors.length <= 1 || presetColorId) && (
        <div style={styles.movementItemName}>اللون: {selectedColor?.color}</div>
      )}

      {!presetKind && (
        <Field label="نوع الحركة">
          <div style={styles.roleToggle}>
            <button style={{ ...styles.roleOption, ...(kind === "size" ? styles.roleOptionActive : {}) }} onClick={() => setKind("size")}>بالمقاس (قطعة)</button>
            <button style={{ ...styles.roleOption, ...(kind === "pack" ? styles.roleOptionActive : {}) }} onClick={() => setKind("pack")}>باكنج</button>
          </div>
        </Field>
      )}

      {kind === "size" && selectedColor?.sizes.length > 1 && !presetSizeId && (
        <Field label="المقاس">
          <select style={styles.input} value={sizeId} onChange={(e) => setSizeId(e.target.value)}>
            {selectedColor.sizes.map((s) => <option key={s.id} value={s.id}>{s.size}</option>)}
          </select>
        </Field>
      )}
      {kind === "size" && (selectedColor?.sizes.length <= 1 || presetSizeId) && (
        <div style={styles.movementItemName}>المقاس: {currentSize?.size}</div>
      )}

      <div style={styles.movementCurrent}>الرصيد الحالي: <b>{currentAvailable}</b> {kind === "pack" ? "باكنج" : "قطعة"}</div>

      <Field label={`الكمية ${isOut ? "المصروفة" : "المستلمة"}`}>
        <input style={styles.input} type="number" value={qty} onChange={(e) => setQty(e.target.value)} placeholder="0" autoFocus />
      </Field>
      <Field label="ملاحظة (اختياري)">
        <input style={styles.input} value={note} onChange={(e) => setNote(e.target.value)} placeholder={isOut ? "الجهة المستلمة، سبب الصرف..." : "المورد، رقم الفاتورة..."} />
      </Field>
      {error && <div style={styles.errorText}>{error}</div>}
      <button style={{ ...styles.primaryBtn, width: "100%", justifyContent: "center", marginTop: 8, background: isOut ? "#B5493A" : "#3D7A4E" }} onClick={submit}>
        {isOut ? "تأكيد الصرف" : "تأكيد الاستلام"}
      </button>
    </Modal>
  );
}

/* ---------- Import ---------- */

function parseSheetRows(rows) {
  let headerIdx = -1, headerRow = null;
  for (let i = 0; i < Math.min(rows.length, 8); i++) {
    const r = rows[i].map((c) => String(c ?? "").trim());
    if (r.some((h) => /^code$|^كود$/i.test(h))) { headerIdx = i; headerRow = r; break; }
  }
  if (headerIdx === -1) return { rows: [], error: "تعذر العثور على عمود CODE في الملف" };

  const codeIdx = headerRow.findIndex((h) => /code|كود/i.test(h));
  const colorIdx = headerRow.findIndex((h) => /color|colour|لون/i.test(h));
  const pakIdx = headerRow.findIndex((h) => /pak|باكنج/i.test(h));
  const pcsIdx = headerRow.findIndex((h) => /pcs|piece|قطع/i.test(h));
  if (colorIdx === -1) return { rows: [], error: "تعذر العثور على عمود COLOR في الملف" };

  const sizeIdxs = headerRow
    .map((h, i) => i)
    .filter((i) => i > colorIdx && i !== pakIdx && i !== pcsIdx && headerRow[i] && headerRow[i] !== "");

  const dataRows = rows.slice(headerIdx + 1).filter((r) => String(r[codeIdx] ?? "").trim());

  const parsed = dataRows.map((r) => {
    const sizes = sizeIdxs.map((i) => ({ size: headerRow[i], qty: Number(r[i]) || 0 }));
    const pak = pakIdx !== -1 ? (Number(r[pakIdx]) || 0) : sizes.reduce((s, x) => s + x.qty, 0);
    return { code: String(r[codeIdx] ?? "").trim(), color: String(r[colorIdx] ?? "").trim() || "بدون لون", sizes, packQty: pak };
  }).filter((r) => r.code);

  return { rows: parsed, error: null };
}

function ImportExcelModal({ onClose, onConfirm, defaultWarehouse }) {
  const [warehouseCode, setWarehouseCode] = useState(defaultWarehouse || WAREHOUSES[0].code);
  const [parsed, setParsed] = useState(null);
  const [error, setError] = useState("");
  const [fileName, setFileName] = useState("");
  const inputRef = useRef(null);

  const handleFile = async (file) => {
    setError(""); setFileName(file.name);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      const { rows: result, error: parseError } = parseSheetRows(rows);
      if (parseError || result.length === 0) { setError(parseError || "لم يتم العثور على بيانات صالحة"); setParsed(null); return; }
      setParsed(result);
    } catch (e) { setError("تعذرت قراءة الملف، تأكد من أنه ملف Excel صحيح (.xlsx أو .xls)"); setParsed(null); }
  };

  return (
    <Modal onClose={onClose} title="استيراد أصناف من Excel">
      <Field label="المخزن المستهدف">
        <select style={styles.input} value={warehouseCode} onChange={(e) => setWarehouseCode(e.target.value)}>
          {WAREHOUSES.map((w) => <option key={w.code} value={w.code}>{w.code} — {w.name}</option>)}
        </select>
      </Field>

      {!parsed ? (
        <>
          <div style={styles.importHint}>الملف لازم يحتوي على عمود CODE وعمود COLOR وأعمدة المقاسات، وعمود إجمالي الباكنج (PAK) اختياري.</div>
          <div style={styles.dropZone} onClick={() => inputRef.current?.click()}>
            <Upload size={26} color="#8A8F98" />
            <div style={styles.dropText}>{fileName || "اضغط لاختيار ملف Excel"}</div>
            <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])} />
          </div>
          {error && <div style={styles.errorText}>{error}</div>}
        </>
      ) : (
        <>
          <div style={styles.importHint}>سيتم استيراد <b>{parsed.length}</b> صف (كود + لون) إلى مخزن <b>{whName(warehouseCode)}</b>. أي كود+لون موجود بالفعل سيتم تحديث أرصدته.</div>
          <div style={styles.previewTable}>
            {parsed.slice(0, 50).map((r, i) => (
              <div key={i} style={styles.previewRow}>
                <span style={styles.previewCode}>{r.code}</span>
                <span style={styles.previewName}>{r.color}</span>
                <span style={styles.previewQty}>{r.packQty}</span>
              </div>
            ))}
            {parsed.length > 50 && <div style={styles.previewMore}>و {parsed.length - 50} صف آخر...</div>}
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
            <button style={{ ...styles.secondaryBtn, flex: 1, justifyContent: "center" }} onClick={() => { setParsed(null); setFileName(""); }}>اختيار ملف آخر</button>
            <button style={{ ...styles.primaryBtn, flex: 1, justifyContent: "center" }} onClick={() => onConfirm(parsed, warehouseCode)}><Check size={17} />تأكيد الاستيراد</button>
          </div>
        </>
      )}
    </Modal>
  );
}

/* ---------- Min-qty alert threshold ---------- */

function EditMinQtyModal({ item, onClose, onSave }) {
  const [val, setVal] = useState(item.minQty == null ? "" : String(item.minQty));
  return (
    <Modal onClose={onClose} title={`حد التنبيه — ${item.code}`}>
      <div style={styles.importHint}>عند وصول إجمالي القطع لهذا الصنف (كل الألوان والمقاسات مع بعض) إلى هذا الرقم أو أقل، سيظهر تنبيه "تحت الحد الأدنى".</div>
      <Field label="حد التنبيه (بالقطعة) — اتركه فارغًا لإلغاء التنبيه">
        <input style={styles.input} type="number" value={val} onChange={(e) => setVal(e.target.value)} placeholder="مثال: 100" autoFocus />
      </Field>
      <button style={{ ...styles.primaryBtn, width: "100%", justifyContent: "center", marginTop: 8 }} onClick={() => onSave(val === "" ? null : Number(val))}>
        حفظ
      </button>
    </Modal>
  );
}

/* ---------- Daily balance refresh (full replace) ---------- */

function DailyRefreshModal({ onClose, onConfirm }) {
  const [warehouseCode, setWarehouseCode] = useState(WAREHOUSES[0].code);
  const [parsed, setParsed] = useState(null);
  const [error, setError] = useState("");
  const [fileName, setFileName] = useState("");
  const [confirmed, setConfirmed] = useState(false);
  const inputRef = useRef(null);

  const handleFile = async (file) => {
    setError(""); setFileName(file.name); setConfirmed(false);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      const { rows: result, error: parseError } = parseSheetRows(rows);
      if (parseError || result.length === 0) { setError(parseError || "لم يتم العثور على بيانات صالحة"); setParsed(null); return; }
      setParsed(result);
    } catch (e) { setError("تعذرت قراءة الملف، تأكد من أنه ملف Excel صحيح (.xlsx أو .xls)"); setParsed(null); }
  };

  return (
    <Modal onClose={onClose} title="تحديث يومي للرصيد">
      <div style={{ ...styles.importHint, background: "#FBEEEA", color: "#8A3A2E" }}>
        ⚠️ هذا سيمسح <b>كل</b> الأصناف الحالية في المخزن المختار ويستبدلها بالكامل بمحتوى الملف. أي صنف كان موجودًا ومش موجود في الملف الجديد سيُحذف نهائيًا. سجل الحركات السابق لن يتأثر.
      </div>
      <Field label="المخزن">
        <select style={styles.input} value={warehouseCode} onChange={(e) => { setWarehouseCode(e.target.value); setConfirmed(false); }}>
          {WAREHOUSES.map((w) => <option key={w.code} value={w.code}>{w.code} — {w.name}</option>)}
        </select>
      </Field>

      {!parsed ? (
        <>
          <div style={styles.dropZone} onClick={() => inputRef.current?.click()}>
            <Upload size={26} color="#8A8F98" />
            <div style={styles.dropText}>{fileName || "اضغط لاختيار ملف Excel"}</div>
            <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])} />
          </div>
          {error && <div style={styles.errorText}>{error}</div>}
        </>
      ) : (
        <>
          <div style={styles.importHint}>الملف يحتوي على <b>{parsed.length}</b> صف (كود + لون). سيصبح هذا هو الرصيد الكامل لمخزن <b>{whName(warehouseCode)}</b>.</div>
          <label style={styles.confirmCheckRow}>
            <input type="checkbox" checked={confirmed} onChange={(e) => setConfirmed(e.target.checked)} />
            أفهم أن هذا سيمسح الرصيد الحالي بالكامل ويستبدله بمحتوى الملف
          </label>
          <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
            <button style={{ ...styles.secondaryBtn, flex: 1, justifyContent: "center" }} onClick={() => { setParsed(null); setFileName(""); setConfirmed(false); }}>اختيار ملف آخر</button>
            <button
              style={{ ...styles.primaryBtn, flex: 1, justifyContent: "center", opacity: confirmed ? 1 : 0.5, cursor: confirmed ? "pointer" : "not-allowed", background: "#B5493A" }}
              onClick={() => confirmed && onConfirm(parsed, warehouseCode)}
              disabled={!confirmed}
            >
              <RefreshCw size={17} />استبدال الرصيد بالكامل
            </button>
          </div>
        </>
      )}
    </Modal>
  );
}

/* ---------- Orders ---------- */

function OrdersScreen({ orders, factories, factoryFilter, onFactoryFilter, search, onSearch, onAdd, onImport, onExport, onManageFactories, onOpen, detailOrder, onBack, onReceive, onSetStatus }) {
  if (detailOrder) {
    return <OrderDetailView order={detailOrder} onBack={onBack} onReceive={() => onReceive(detailOrder)} onSetStatus={onSetStatus} />;
  }

  const openCount = orders.filter((o) => o.status === "open").length;

  return (
    <div>
      <div style={styles.statsRow}>
        <StatCard label="عدد الأوردرات" value={orders.length} accent="#D9A441" />
        <StatCard label="أوردرات مفتوحة" value={openCount} accent="#4A6FA5" />
        <StatCard label="إجمالي المطلوب" value={orders.reduce((s, o) => s + o.orderedQty, 0)} accent="#3D7A4E" />
      </div>

      <div style={styles.toolbar}>
        <select style={styles.whSelect} value={factoryFilter} onChange={(e) => onFactoryFilter(e.target.value)}>
          <option value="">كل المصانع</option>
          {factories.map((f) => <option key={f} value={f}>{f}</option>)}
        </select>
        <div style={styles.searchBox}>
          <Search size={17} color="#8A8F98" style={{ flexShrink: 0 }} />
          <input placeholder="ابحث بالكود أو PO..." value={search} onChange={(e) => onSearch(e.target.value)} style={styles.searchInput} />
        </div>
        <button style={styles.secondaryBtn} onClick={onManageFactories}><Factory size={17} />إدارة المصانع</button>
        <button style={styles.secondaryBtn} onClick={onImport}><FileSpreadsheet size={17} />استيراد أوردرات</button>
        <button style={styles.secondaryBtn} onClick={onExport}><Download size={17} />تصدير الأوردرات</button>
        <button style={styles.primaryBtn} onClick={onAdd}><Plus size={17} />أوردر جديد</button>
      </div>

      {orders.length === 0 ? (
        <EmptyState hasSearch={!!search || !!factoryFilter} onAdd={onAdd} message="لا توجد أوردرات مسجلة بعد" />
      ) : (
        <div style={styles.logList}>
          {orders.slice().sort((a, b) => (b.createdAt || "").localeCompare(a.createdAt || "")).map((o) => (
            <OrderRow key={o.id} order={o} onOpen={() => onOpen(o.id)} />
          ))}
        </div>
      )}
    </div>
  );
}

function OrderRow({ order, onOpen }) {
  const closed = order.status === "closed";
  return (
    <div style={styles.orderRow} onClick={onOpen} role="button" tabIndex={0}>
      <div style={{ ...styles.statusDot, background: closed ? "#8A8F98" : "#3D7A4E" }} />
      <div style={styles.logMid}>
        <div style={styles.logTitle}>{order.code} · {order.color} <span style={styles.logWh}>({order.factory})</span></div>
        <div style={styles.logNote}>PO: {order.po || "—"} {order.internalPo && `· Internal PO: ${order.internalPo}`} · {order.date}</div>
      </div>
      <div style={styles.orderProgress}>
        <div>مطلوب <b>{order.orderedQty}</b></div>
        <div style={{ color: "#3D7A4E" }}>وصل <b>{order.receivedQty}</b></div>
        <div style={{ color: "#B5493A" }}>متبقي <b>{order.remainingAtFactory}</b></div>
      </div>
    </div>
  );
}

function OrderDetailView({ order, onBack, onReceive, onSetStatus }) {
  const closed = order.status === "closed";
  return (
    <div>
      <button style={styles.backBtn} onClick={onBack}><ArrowRight size={16} />العودة لكل الأوردرات</button>

      <div style={styles.detailHeaderCard}>
        <div style={styles.detailTopRow}>
          <div>
            <div style={styles.itemCode}>{order.code} · {order.color}</div>
            <div style={styles.detailName}>{order.factory}</div>
          </div>
          <div style={{ ...styles.statusChip, background: closed ? "#EEF0F3" : "#EAF4EC", color: closed ? "#5C6470" : "#3D7A4E" }}>
            {closed ? <Lock size={12} /> : <Unlock size={12} />}
            {closed ? "مغلق" : "مفتوح"}
          </div>
        </div>

        <div style={styles.orderInfoGrid}>
          <div><span style={styles.orderInfoLabel}>PO</span><div>{order.po || "—"}</div></div>
          <div><span style={styles.orderInfoLabel}>Internal PO</span><div>{order.internalPo || "لم يُسجَّل بعد"}</div></div>
          <div><span style={styles.orderInfoLabel}>التاريخ</span><div>{order.date}</div></div>
          <div><span style={styles.orderInfoLabel}>المخزن</span><div>{whName(order.warehouseCode)}</div></div>
        </div>

        <div style={styles.orderTotalsRow}>
          <div style={styles.orderTotalBox}><div style={styles.orderTotalVal}>{order.orderedQty}</div><div style={styles.orderTotalLabel}>مطلوب</div></div>
          <div style={styles.orderTotalBox}><div style={{ ...styles.orderTotalVal, color: "#3D7A4E" }}>{order.receivedQty}</div><div style={styles.orderTotalLabel}>وصل</div></div>
          <div style={styles.orderTotalBox}><div style={{ ...styles.orderTotalVal, color: "#B5493A" }}>{order.remainingAtFactory}</div><div style={styles.orderTotalLabel}>متبقي في المصنع</div></div>
        </div>

        {(order.sizes || []).length > 0 && (
          <div style={styles.sizeTable}>
            {order.sizes.map((s, i) => (
              <div key={i} style={styles.sizeRow}><span style={styles.sizeName}>{s.size}</span><span style={styles.sizeQty}>{s.qty}</span></div>
            ))}
          </div>
        )}

        <div style={styles.detailActions}>
          {!closed && <button style={styles.inBtn} onClick={onReceive}><ArrowDownToLine size={15} />تسجيل استلام</button>}
          <button style={styles.secondaryBtn} onClick={() => onSetStatus(order.id, closed ? "open" : "closed")}>
            {closed ? <><Unlock size={15} />إعادة فتح الأوردر</> : <><Lock size={15} />إغلاق الأوردر</>}
          </button>
        </div>
      </div>

      <div style={styles.detailLogTitle}>سجل الاستلامات</div>
      {(order.receipts || []).length === 0 ? (
        <EmptyState hasSearch={false} onAdd={null} message="لم يتم تسجيل أي استلام بعد" />
      ) : (
        <div style={styles.logList}>
          {order.receipts.slice().reverse().map((r) => (
            <div key={r.id} style={styles.logRow}>
              <div style={{ ...styles.logIcon, background: "#EAF4EC", color: "#3D7A4E" }}><ArrowDownToLine size={15} /></div>
              <div style={styles.logMid}>
                <div style={styles.logTitle}>استلام {r.qty} قطعة</div>
                <div style={styles.logNote}>{r.note || "—"} — بواسطة {r.by || "—"}</div>
              </div>
              <div style={styles.logRight}><div style={styles.logDate}>{fmtDate(r.date)}</div></div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function OrderFormModal({ factories, onClose, onSave }) {
  const [factory, setFactory] = useState(factories[0] || "");
  const [date, setDate] = useState(todayDateInput());
  const [warehouseCode, setWarehouseCode] = useState(WAREHOUSES[0].code);
  const [code, setCode] = useState("");
  const [po, setPo] = useState("");
  const [color, setColor] = useState("");
  const [bySize, setBySize] = useState(false);
  const [rows, setRows] = useState(DEFAULT_SIZES.map((s) => ({ size: s, qty: "" })));
  const [totalQty, setTotalQty] = useState("");
  const [error, setError] = useState("");

  const canUseSizes = warehouseCode === "1";

  const submit = () => {
    if (!factory) return setError("اختر المصنع");
    if (!code.trim() || !color.trim()) return setError("أدخل الكود واللون على الأقل");
    let sizes = [], orderedQty = 0;
    if (canUseSizes && bySize) {
      sizes = rows.filter((r) => r.size.trim()).map((r) => ({ size: r.size.trim(), qty: r.qty === "" ? 0 : Number(r.qty) || 0 }));
      orderedQty = sizes.reduce((s, x) => s + x.qty, 0);
      if (orderedQty <= 0) return setError("أدخل كميات المقاسات");
    } else {
      orderedQty = Number(totalQty);
      if (!totalQty || isNaN(orderedQty) || orderedQty <= 0) return setError("أدخل الكمية المطلوبة");
    }
    onSave({ factory, date, warehouseCode, code: code.trim(), po: po.trim(), color: color.trim(), sizes, orderedQty });
  };

  return (
    <Modal onClose={onClose} title="أوردر جديد">
      <Field label="المصنع">
        <select style={styles.input} value={factory} onChange={(e) => setFactory(e.target.value)}>
          {factories.map((f) => <option key={f} value={f}>{f}</option>)}
        </select>
      </Field>
      <div style={styles.fieldRow}>
        <Field label="تاريخ الأوردر"><input style={styles.input} type="date" value={date} onChange={(e) => setDate(e.target.value)} /></Field>
        <Field label="المخزن">
          <select style={styles.input} value={warehouseCode} onChange={(e) => { setWarehouseCode(e.target.value); setBySize(false); }}>
            {WAREHOUSES.map((w) => <option key={w.code} value={w.code}>{w.code} — {w.name}</option>)}
          </select>
        </Field>
      </div>
      <div style={styles.fieldRow}>
        <Field label="كود الصنف"><input style={styles.input} value={code} onChange={(e) => setCode(e.target.value)} placeholder="مثال: DM235" autoFocus /></Field>
        <Field label="اللون"><input style={styles.input} value={color} onChange={(e) => setColor(e.target.value)} placeholder="مثال: أبيض" /></Field>
      </div>
      <Field label="PO (رقمك الخاص)"><input style={styles.input} value={po} onChange={(e) => setPo(e.target.value)} placeholder="مثال: PO-1024" /></Field>

      {canUseSizes && (
        <Field label="طريقة تسجيل الكمية">
          <div style={styles.roleToggle}>
            <button style={{ ...styles.roleOption, ...(!bySize ? styles.roleOptionActive : {}) }} onClick={() => setBySize(false)}>إجمالي فقط</button>
            <button style={{ ...styles.roleOption, ...(bySize ? styles.roleOptionActive : {}) }} onClick={() => setBySize(true)}>بالمقاسات</button>
          </div>
        </Field>
      )}

      {canUseSizes && bySize ? (
        <SizeRowsEditor rows={rows} setRows={setRows} />
      ) : (
        <Field label="الكمية المطلوبة من المصنع (بالقطعة)">
          <input style={styles.input} type="number" value={totalQty} onChange={(e) => setTotalQty(e.target.value)} placeholder="0" />
        </Field>
      )}

      {error && <div style={styles.errorText}>{error}</div>}
      <button style={{ ...styles.primaryBtn, width: "100%", justifyContent: "center", marginTop: 8 }} onClick={submit}>حفظ الأوردر</button>
    </Modal>
  );
}

function ReceiveOrderModal({ order, onClose, onConfirm }) {
  const [internalPo, setInternalPo] = useState(order.internalPo || "");
  const [rows, setRows] = useState((order.sizes || []).map((s) => ({ size: s.size, qty: "" })));
  const [totalQty, setTotalQty] = useState("");
  const [note, setNote] = useState("");
  const [error, setError] = useState("");
  const bySize = (order.sizes || []).length > 0 && order.warehouseCode === "1";

  const submit = () => {
    if (!order.internalPo && !internalPo.trim()) return setError("أدخل رقم الـ Internal PO الخاص بالمصنع");
    let sizeBreakdown = null, qty = 0;
    if (bySize) {
      sizeBreakdown = rows.map((r) => ({ size: r.size, qty: r.qty === "" ? 0 : Number(r.qty) || 0 })).filter((r) => r.qty > 0);
      qty = sizeBreakdown.reduce((s, x) => s + x.qty, 0);
    } else {
      qty = Number(totalQty);
    }
    if (!qty || qty <= 0) return setError("أدخل كمية صحيحة أكبر من صفر");
    if (qty > order.remainingAtFactory) return setError(`المتبقي في المصنع ${order.remainingAtFactory} قطعة فقط، لا يمكن استلام كمية أكبر منه`);
    onConfirm(internalPo.trim(), sizeBreakdown, qty, note.trim());
  };

  return (
    <Modal onClose={onClose} title={`تسجيل استلام — ${order.code}`}>
      <div style={styles.movementCurrent}>المتبقي في المصنع حاليًا: <b>{order.remainingAtFactory}</b> قطعة</div>

      {order.internalPo ? (
        <div style={styles.movementItemName}>Internal PO لهذا الأوردر: <b>{order.internalPo}</b> (ثابت لكل استلامات هذا الأوردر)</div>
      ) : (
        <Field label="Internal PO (رقم أوردر المصنع)">
          <input style={styles.input} value={internalPo} onChange={(e) => setInternalPo(e.target.value)} placeholder="مثال: INT-778" autoFocus />
        </Field>
      )}

      {bySize ? (
        <div style={styles.field}>
          <label style={styles.fieldLabel}>الكمية المستلمة الآن لكل مقاس</label>
          {rows.map((r, i) => (
            <div key={i} style={styles.sizeEditRow}>
              <span style={{ ...styles.sizeName, flex: 2 }}>{r.size}</span>
              <input style={{ ...styles.input, flex: 1 }} type="number" value={r.qty} onChange={(e) => setRows(rows.map((x, idx) => (idx === i ? { ...x, qty: e.target.value } : x)))} placeholder="0" />
            </div>
          ))}
        </div>
      ) : (
        <Field label="الكمية المستلمة الآن (بالقطعة)">
          <input style={styles.input} type="number" value={totalQty} onChange={(e) => setTotalQty(e.target.value)} placeholder="0" />
        </Field>
      )}

      <Field label="ملاحظة (اختياري)"><input style={styles.input} value={note} onChange={(e) => setNote(e.target.value)} placeholder="أي تفاصيل إضافية..." /></Field>
      {error && <div style={styles.errorText}>{error}</div>}
      <button style={{ ...styles.primaryBtn, width: "100%", justifyContent: "center", marginTop: 8, background: "#3D7A4E" }} onClick={submit}>
        تأكيد الاستلام
      </button>
    </Modal>
  );
}

function ImportOrdersModal({ onClose, onConfirm }) {
  const [warehouseCode, setWarehouseCode] = useState(WAREHOUSES[0].code);
  const [parsed, setParsed] = useState(null);
  const [error, setError] = useState("");
  const [fileName, setFileName] = useState("");
  const inputRef = useRef(null);

  const handleFile = async (file) => {
    setError(""); setFileName(file.name);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      const { rows: result, error: parseError } = parseOrdersExcelRows(rows, warehouseCode);
      if (parseError || result.length === 0) { setError(parseError || "لم يتم العثور على بيانات صالحة"); setParsed(null); return; }
      setParsed(result);
    } catch (e) { setError("تعذرت قراءة الملف، تأكد من أنه ملف Excel صحيح (.xlsx أو .xls)"); setParsed(null); }
  };

  return (
    <Modal onClose={onClose} title="استيراد أوردرات من Excel">
      <Field label="المخزن">
        <select style={styles.input} value={warehouseCode} onChange={(e) => { setWarehouseCode(e.target.value); setParsed(null); setFileName(""); }}>
          {WAREHOUSES.map((w) => <option key={w.code} value={w.code}>{w.code} — {w.name}</option>)}
        </select>
      </Field>

      {!parsed ? (
        <>
          <div style={styles.importHint}>
            الأعمدة المتوقعة: المصنع، التاريخ، الكود، PO، Internal PO، اللون، الكمية المتبقية في المصنع، الكمية المطلوبة
            {warehouseCode === "1" && "، وعمود لكل مقاس"}.
          </div>
          <div style={styles.dropZone} onClick={() => inputRef.current?.click()}>
            <Upload size={26} color="#8A8F98" />
            <div style={styles.dropText}>{fileName || "اضغط لاختيار ملف Excel"}</div>
            <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])} />
          </div>
          {error && <div style={styles.errorText}>{error}</div>}
        </>
      ) : (
        <>
          <div style={styles.importHint}>سيتم استيراد <b>{parsed.length}</b> أوردر إلى مخزن <b>{whName(warehouseCode)}</b>.</div>
          <div style={styles.previewTable}>
            {parsed.slice(0, 50).map((r, i) => (
              <div key={i} style={styles.previewRow}>
                <span style={styles.previewCode}>{r.code}</span>
                <span style={styles.previewName}>{r.factory} · {r.color}</span>
                <span style={styles.previewQty}>{r.orderedQty}</span>
              </div>
            ))}
            {parsed.length > 50 && <div style={styles.previewMore}>و {parsed.length - 50} أوردر آخر...</div>}
          </div>
          <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
            <button style={{ ...styles.secondaryBtn, flex: 1, justifyContent: "center" }} onClick={() => { setParsed(null); setFileName(""); }}>اختيار ملف آخر</button>
            <button style={{ ...styles.primaryBtn, flex: 1, justifyContent: "center" }} onClick={() => onConfirm(parsed, warehouseCode)}><Check size={17} />تأكيد الاستيراد</button>
          </div>
        </>
      )}
    </Modal>
  );
}

function FactoriesModal({ factories, onAdd, onRemove, onClose }) {
  const [name, setName] = useState("");
  return (
    <Modal onClose={onClose} title="إدارة المصانع">
      <div style={styles.field}>
        {factories.map((f) => (
          <div key={f} style={styles.factoryRow}>
            <span>{f}</span>
            <button style={styles.deleteBtn} onClick={() => onRemove(f)}><Trash2 size={14} /></button>
          </div>
        ))}
      </div>
      <div style={styles.sizeEditRow}>
        <input style={{ ...styles.input, flex: 1 }} value={name} onChange={(e) => setName(e.target.value)} placeholder="اسم مصنع جديد" />
        <button style={styles.primaryBtn} onClick={() => { if (name.trim()) { onAdd(name.trim()); setName(""); } }}><Plus size={16} /></button>
      </div>
    </Modal>
  );
}

/* ---------- Shared UI ---------- */

function Modal({ title, children, onClose }) {
  return (
    <div style={styles.overlay} onClick={onClose}>
      <div style={styles.modal} onClick={(e) => e.stopPropagation()}>
        <div style={styles.modalHeader}>
          <span style={styles.modalTitle}>{title}</span>
          <button style={styles.closeBtn} onClick={onClose}><X size={18} /></button>
        </div>
        <div style={styles.modalBody}>{children}</div>
      </div>
    </div>
  );
}

function Field({ label, children }) {
  return <div style={styles.field}><label style={styles.fieldLabel}>{label}</label>{children}</div>;
}

const styles = {
  page: { minHeight: "100vh", background: "#F4F1EB", fontFamily: "'Segoe UI', Tahoma, 'Cairo', sans-serif", color: "#1B1F24" },
  header: { background: "#1B1F24", borderBottom: "3px solid #D9A441", position: "sticky", top: 0, zIndex: 10 },
  headerInner: { maxWidth: 1040, margin: "0 auto", padding: "14px 18px", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 },
  headerRight: { display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" },
  brand: { display: "flex", alignItems: "center", gap: 10 },
  brandMark: { width: 36, height: 36, borderRadius: 8, background: "#D9A441", color: "#1B1F24", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, fontSize: 17, flexShrink: 0 },
  brandTitle: { color: "#F4F1EB", fontWeight: 700, fontSize: 16, lineHeight: 1.3 },
  brandSub: { color: "#9AA0AB", fontSize: 12, lineHeight: 1.3 },
  nav: { display: "flex", gap: 6, background: "#262B33", padding: 4, borderRadius: 10 },
  navBtn: { display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", borderRadius: 7, border: "none", background: "transparent", color: "#9AA0AB", fontSize: 13.5, fontWeight: 600, cursor: "pointer" },
  navBtnActive: { background: "#D9A441", color: "#1B1F24" },
  userBadge: { display: "flex", alignItems: "center", gap: 8, background: "#262B33", borderRadius: 20, padding: "6px 8px 6px 12px" },
  roleDot: { width: 8, height: 8, borderRadius: "50%", flexShrink: 0 },
  userName: { color: "#F4F1EB", fontSize: 13, fontWeight: 700 },
  logoutBtn: { background: "#3A404A", border: "none", width: 26, height: 26, borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", color: "#D9A441", cursor: "pointer" },
  main: { maxWidth: 1040, margin: "0 auto", padding: "22px 18px 60px" },
  loading: { textAlign: "center", padding: 60, color: "#8A8F98" },
  statsRow: { display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 20 },
  statCard: { background: "#FFFFFF", borderRadius: 10, padding: "14px 16px", boxShadow: "0 1px 2px rgba(27,31,36,0.06)" },
  statValue: { fontSize: 24, fontWeight: 800, lineHeight: 1.2 },
  statLabel: { fontSize: 12.5, color: "#6B7280", marginTop: 2 },
  toolbar: { display: "flex", gap: 10, marginBottom: 18, flexWrap: "wrap" },
  whSelect: { border: "1px solid #E3E0D8", borderRadius: 9, padding: "10px 12px", fontSize: 13.5, fontWeight: 700, background: "#FFFFFF", color: "#1B1F24" },
  searchBox: { flex: 1, minWidth: 180, display: "flex", alignItems: "center", gap: 8, background: "#FFFFFF", border: "1px solid #E3E0D8", borderRadius: 9, padding: "10px 14px" },
  searchInput: { border: "none", outline: "none", background: "transparent", flex: 1, fontSize: 14, color: "#1B1F24" },
  primaryBtn: { display: "flex", alignItems: "center", gap: 7, background: "#1B1F24", color: "#FFFFFF", border: "none", borderRadius: 9, padding: "10px 18px", fontSize: 14, fontWeight: 700, cursor: "pointer", whiteSpace: "nowrap" },
  secondaryBtn: { display: "flex", alignItems: "center", gap: 7, background: "#FFFFFF", color: "#1B1F24", border: "1px solid #E3E0D8", borderRadius: 9, padding: "10px 18px", fontSize: 14, fontWeight: 700, cursor: "pointer", whiteSpace: "nowrap" },
  itemsGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(240px, 1fr))", gap: 12 },
  itemCard: { background: "#FFFFFF", borderRadius: 12, padding: 16, border: "1px solid #EDEAE2", cursor: "pointer" },
  whBadge: { display: "flex", alignItems: "center", gap: 4, background: "#EEF2F7", color: "#4A6FA5", fontSize: 10.5, fontWeight: 700, padding: "3px 8px", borderRadius: 20, flexShrink: 0, whiteSpace: "nowrap" },
  colorChipsRow: { display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap", margin: "10px 0" },
  colorChip: { background: "#F4F1EB", color: "#4B5563", fontSize: 11, fontWeight: 600, padding: "2px 8px", borderRadius: 20 },
  backBtn: { display: "flex", alignItems: "center", gap: 6, background: "none", border: "none", color: "#6B7280", fontSize: 13.5, fontWeight: 700, cursor: "pointer", padding: "6px 0", marginBottom: 14 },
  detailHeaderCard: { background: "#FFFFFF", borderRadius: 14, border: "1px solid #EDEAE2", padding: 20, marginBottom: 18 },
  detailTopRow: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 },
  detailName: { fontSize: 19, fontWeight: 800, marginTop: 3, color: "#1B1F24" },
  detailQtyRow: { display: "flex", alignItems: "baseline", gap: 8, margin: "18px 0" },
  detailQty: { fontSize: 42, fontWeight: 800, color: "#1B1F24", lineHeight: 1 },
  detailActions: { display: "flex", gap: 10, flexWrap: "wrap" },
  detailLogTitle: { fontSize: 14, fontWeight: 700, color: "#1B1F24", margin: "4px 0 10px" },
  detailLogRow: { display: "flex", alignItems: "center", justifyContent: "space-between", margin: "18px 0 10px" },
  exportSmallBtn: { display: "flex", alignItems: "center", gap: 5, background: "#FFFFFF", color: "#1B1F24", border: "1px solid #E3E0D8", borderRadius: 7, padding: "6px 11px", fontSize: 12, fontWeight: 700, cursor: "pointer" },
  itemTop: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 },
  itemCode: { fontSize: 12, color: "#8A8F98", fontWeight: 700, letterSpacing: 0.3 },
  itemName: { fontSize: 15.5, fontWeight: 700, marginTop: 2, color: "#1B1F24" },
  itemQtyRow: { display: "flex", alignItems: "baseline", gap: 18, margin: "10px 0 14px" },
  itemQty: { fontSize: 26, fontWeight: 800, color: "#1B1F24", lineHeight: 1 },
  itemQtyPack: { fontSize: 26, fontWeight: 800, color: "#4A6FA5", lineHeight: 1 },
  itemUnit: { fontSize: 12.5, color: "#8A8F98" },
  itemActions: { display: "flex", gap: 8 },
  inBtn: { flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 6, background: "#EAF4EC", color: "#3D7A4E", border: "none", borderRadius: 8, padding: "9px 10px", fontSize: 13, fontWeight: 700, cursor: "pointer" },
  outBtn: { flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 6, background: "#FBEEEA", color: "#B5493A", border: "none", borderRadius: 8, padding: "9px 10px", fontSize: 13, fontWeight: 700, cursor: "pointer" },
  colorSection: { background: "#FFFFFF", borderRadius: 12, border: "1px solid #EDEAE2", padding: 16, marginBottom: 12 },
  colorSectionHeader: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, flexWrap: "wrap", gap: 8 },
  colorSectionTitle: { display: "flex", alignItems: "center", gap: 6, fontSize: 14.5, fontWeight: 800, color: "#1B1F24" },
  colorSectionPack: { display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, color: "#4A6FA5", fontWeight: 600 },
  sizeTable: { display: "flex", flexDirection: "column", gap: 6 },
  sizeRow: { display: "flex", alignItems: "center", gap: 10, background: "#F9F8F5", borderRadius: 8, padding: "7px 10px" },
  sizeName: { flex: 1, fontSize: 13, fontWeight: 700, color: "#4B5563" },
  sizeQty: { fontSize: 14, fontWeight: 800, color: "#1B1F24", minWidth: 34, textAlign: "center" },
  miniBtn: { background: "#EEF2F7", color: "#4A6FA5", border: "none", width: 22, height: 22, borderRadius: 6, fontWeight: 800, cursor: "pointer" },
  miniBtnIn: { background: "#EAF4EC", color: "#3D7A4E", border: "none", width: 26, height: 26, borderRadius: 6, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer" },
  miniBtnOut: { background: "#FBEEEA", color: "#B5493A", border: "none", width: 26, height: 26, borderRadius: 6, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer" },
  logList: { display: "flex", flexDirection: "column", gap: 8 },
  logRow: { display: "flex", alignItems: "center", gap: 12, background: "#FFFFFF", borderRadius: 10, padding: "12px 14px", border: "1px solid #EDEAE2" },
  logIcon: { width: 32, height: 32, borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 },
  logMid: { flex: 1, minWidth: 0 },
  logTitle: { fontSize: 14, fontWeight: 700, color: "#1B1F24" },
  logWh: { color: "#4A6FA5", fontWeight: 600 },
  logNote: { fontSize: 12.5, color: "#8A8F98", marginTop: 1 },
  logRight: { textAlign: "left", flexShrink: 0 },
  logQty: { fontSize: 16, fontWeight: 800 },
  logDate: { fontSize: 11, color: "#8A8F98", marginTop: 2, whiteSpace: "nowrap" },
  youTag: { color: "#B8862F", fontWeight: 700, fontSize: 12 },
  deleteBtn: { background: "#FBEEEA", border: "none", width: 32, height: 32, borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", color: "#B5493A", cursor: "pointer", flexShrink: 0 },
  empty: { textAlign: "center", padding: "60px 20px", background: "#FFFFFF", borderRadius: 12, border: "1px dashed #D8D4C8" },
  emptyIcon: { width: 56, height: 56, borderRadius: 14, background: "#F4F1EB", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 14px" },
  emptyTitle: { fontSize: 14.5, color: "#6B7280", marginBottom: 16, fontWeight: 600 },
  busyBanner: { position: "fixed", top: 10, left: "50%", transform: "translateX(-50%)", background: "#1B1F24", color: "#F4F1EB", padding: "6px 16px", borderRadius: 20, fontSize: 12.5, fontWeight: 700, zIndex: 100, boxShadow: "0 4px 14px rgba(0,0,0,0.25)" },
  saveError: { position: "fixed", bottom: 18, right: 18, left: 18, maxWidth: 400, margin: "0 auto", background: "#B5493A", color: "#FFF", padding: "10px 14px", borderRadius: 9, display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontWeight: 600, boxShadow: "0 4px 14px rgba(0,0,0,0.2)" },
  overlay: { position: "fixed", inset: 0, background: "rgba(27,31,36,0.5)", display: "flex", alignItems: "center", justifyContent: "center", padding: 16, zIndex: 50 },
  modal: { background: "#FFFFFF", borderRadius: 14, width: "100%", maxWidth: 460, maxHeight: "90vh", overflowY: "auto" },
  modalHeader: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "16px 18px", borderBottom: "1px solid #EDEAE2" },
  modalTitle: { fontSize: 16, fontWeight: 800, color: "#1B1F24" },
  closeBtn: { background: "#F4F1EB", border: "none", width: 30, height: 30, borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", color: "#6B7280" },
  modalBody: { padding: 18 },
  movementItemName: { fontSize: 13.5, fontWeight: 700, marginBottom: 10, color: "#4B5563" },
  movementCurrent: { fontSize: 13, color: "#6B7280", marginBottom: 16 },
  field: { marginBottom: 14 },
  fieldRow: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 },
  fieldLabel: { display: "block", fontSize: 12.5, fontWeight: 700, color: "#4B5563", marginBottom: 6 },
  input: { width: "100%", border: "1px solid #E3E0D8", borderRadius: 8, padding: "10px 12px", fontSize: 14, outline: "none", color: "#1B1F24" },
  errorText: { color: "#B5493A", fontSize: 12.5, fontWeight: 600, marginTop: -4, marginBottom: 10 },
  sizeEditRow: { display: "flex", gap: 6, marginBottom: 6, alignItems: "center" },
  removeRowBtn: { background: "#FBEEEA", color: "#B5493A", border: "none", width: 32, height: 32, borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", flexShrink: 0 },
  addRowBtn: { display: "flex", alignItems: "center", gap: 5, background: "none", border: "1px dashed #D8D4C8", color: "#6B7280", borderRadius: 8, padding: "7px 12px", fontSize: 12.5, fontWeight: 700, cursor: "pointer", marginTop: 2 },

  loginWrap: { minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", padding: 20 },
  loginCard: { background: "#FFFFFF", borderRadius: 16, padding: "32px 28px", width: "100%", maxWidth: 380, boxShadow: "0 8px 30px rgba(27,31,36,0.08)", textAlign: "center" },
  loginMark: { width: 48, height: 48, borderRadius: 12, background: "#D9A441", color: "#1B1F24", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, fontSize: 21, margin: "0 auto 16px" },
  loginTitle: { fontSize: 19, fontWeight: 800, color: "#1B1F24" },
  loginSub: { fontSize: 13.5, color: "#8A8F98", marginTop: 4, marginBottom: 22, lineHeight: 1.5 },
  pinInput: { width: "100%", border: "1px solid #E3E0D8", borderRadius: 10, padding: "14px", fontSize: 22, letterSpacing: 6, textAlign: "center", outline: "none", color: "#1B1F24" },
  linkBtn: { background: "none", border: "none", color: "#8A8F98", fontSize: 12.5, marginTop: 14, cursor: "pointer", textDecoration: "underline" },
  userList: { display: "flex", flexDirection: "column", gap: 8, textAlign: "right" },
  userListItem: { display: "flex", alignItems: "center", gap: 10, background: "#F4F1EB", border: "none", borderRadius: 10, padding: "12px 14px", cursor: "pointer", width: "100%" },
  userListName: { flex: 1, fontSize: 14.5, fontWeight: 700, color: "#1B1F24" },
  userListRole: { fontSize: 11.5, color: "#8A8F98", fontWeight: 600 },

  roleToggle: { display: "flex", gap: 8 },
  roleOption: { flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 6, padding: "10px 8px", borderRadius: 8, border: "1px solid #E3E0D8", background: "#FFFFFF", color: "#6B7280", fontSize: 13, fontWeight: 700, cursor: "pointer" },
  roleOptionActive: { background: "#1B1F24", color: "#FFFFFF", borderColor: "#1B1F24" },
  roleHint: { fontSize: 12, color: "#8A8F98", marginTop: 6, lineHeight: 1.4 },

  importHint: { fontSize: 13, color: "#6B7280", lineHeight: 1.6, marginBottom: 14, background: "#F4F1EB", padding: "10px 12px", borderRadius: 8 },
  dropZone: { border: "1.5px dashed #D8D4C8", borderRadius: 12, padding: "32px 16px", textAlign: "center", cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 8 },
  dropText: { fontSize: 13.5, color: "#6B7280", fontWeight: 600 },
  previewTable: { maxHeight: 300, overflowY: "auto", border: "1px solid #EDEAE2", borderRadius: 10 },
  previewRow: { display: "flex", alignItems: "center", gap: 10, padding: "8px 12px", borderBottom: "1px solid #EDEAE2", fontSize: 12.5 },
  previewCode: { fontWeight: 700, width: 90, flexShrink: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  previewName: { flex: 1, color: "#6B7280", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  previewQty: { fontWeight: 700, width: 50, textAlign: "center", flexShrink: 0 },
  previewMore: { textAlign: "center", padding: 10, fontSize: 12, color: "#8A8F98" },
};
