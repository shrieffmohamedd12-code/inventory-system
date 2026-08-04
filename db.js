import { supabase } from "./supabaseClient";

function uid() {
  return crypto.randomUUID();
}
function todayISO() {
  return new Date().toISOString();
}

/* ---------- mapping: DB rows (snake_case) <-> app shape (camelCase) ---------- */

const itemFromRow = (r) => ({
  id: r.id,
  warehouseCode: r.warehouse_code,
  code: r.code,
  name: r.name,
  minQty: r.min_qty,
  colors: r.colors || [],
});

const txFromRow = (r) => ({
  id: r.id,
  itemId: r.item_id,
  warehouseCode: r.warehouse_code,
  code: r.code,
  color: r.color,
  size: r.size,
  kind: r.kind,
  type: r.type,
  qty: r.qty,
  note: r.note,
  by: r.by_name,
  date: r.created_at,
});

const orderFromRow = (r) => ({
  id: r.id,
  factory: r.factory,
  date: r.order_date,
  warehouseCode: r.warehouse_code,
  code: r.code,
  po: r.po,
  internalPo: r.internal_po,
  color: r.color,
  sizes: r.sizes || [],
  orderedQty: r.ordered_qty,
  receivedQty: r.received_qty,
  remainingAtFactory: r.remaining_at_factory,
  status: r.status,
  receipts: r.receipts || [],
  createdBy: r.created_by,
  createdAt: r.created_at,
});

const userFromRow = (r) => ({ id: r.id, name: r.name, email: r.email, role: r.role });

/* ---------- full data load ---------- */

export async function fetchProfiles() {
  const { data, error } = await supabase.from("profiles").select("*").order("created_at", { ascending: true });
  if (error) throw error;
  return (data || []).map(userFromRow);
}

export async function fetchAll() {
  const [usersRes, factoriesRes, itemsRes, txRes, ordersRes] = await Promise.all([
    supabase.from("profiles").select("*").order("created_at", { ascending: true }),
    supabase.from("factories").select("*").order("name", { ascending: true }),
    supabase.from("items").select("*").order("code", { ascending: true }),
    supabase.from("transactions").select("*").order("created_at", { ascending: false }).limit(500),
    supabase.from("orders").select("*").order("created_at", { ascending: false }),
  ]);

  for (const res of [usersRes, factoriesRes, itemsRes, txRes, ordersRes]) {
    if (res.error) throw res.error;
  }

  return {
    users: (usersRes.data || []).map(userFromRow),
    factories: (factoriesRes.data || []).map((f) => f.name),
    items: (itemsRes.data || []).map(itemFromRow),
    transactions: (txRes.data || []).map(txFromRow),
    orders: (ordersRes.data || []).map(orderFromRow),
  };
}

/* ---------- auth ---------- */

export async function signInByName(email, pin) {
  const { error } = await supabase.auth.signInWithPassword({ email, password: pin });
  if (error) throw error;
}

export async function signOut() {
  await supabase.auth.signOut();
}

export async function getCurrentProfile() {
  const { data: sessionData } = await supabase.auth.getSession();
  const authUser = sessionData?.session?.user;
  if (!authUser) return null;
  const { data, error } = await supabase.from("profiles").select("*").eq("id", authUser.id).single();
  if (error) return null;
  return userFromRow(data);
}

// used only once, for the very first admin account (bootstraps the system)
export async function signUpFirstAdmin(name, pin) {
  const email = `u-${uid()}@warehouse.local`;
  const { data, error } = await supabase.auth.signUp({ email, password: pin });
  if (error) throw error;
  const userId = data.user.id;
  const { error: profileErr } = await supabase.from("profiles").insert({ id: userId, name, email, role: "admin" });
  if (profileErr) throw profileErr;
  return { id: userId, name, email, role: "admin" };
}

// admin-only: creates another user via the create-user Edge Function (keeps admin's own session intact)
export async function createUser(name, pin, role) {
  const { data, error } = await supabase.functions.invoke("create-user", { body: { name, pin, role } });
  if (error) throw error;
  if (data?.error) throw new Error(data.error);
  return data;
}

// admin-only: deletes a user via the delete-user Edge Function
export async function deleteUser(userId) {
  const { data, error } = await supabase.functions.invoke("delete-user", { body: { userId } });
  if (error) throw error;
  if (data?.error) throw new Error(data.error);
}

/* ---------- factories ---------- */

export async function addFactory(name) {
  const { error } = await supabase.from("factories").insert({ name });
  if (error) throw error;
}
export async function removeFactory(name) {
  const { error } = await supabase.from("factories").delete().eq("name", name);
  if (error) throw error;
}

/* ---------- items & stock movements ---------- */

async function fetchItemFresh(id) {
  const { data, error } = await supabase.from("items").select("*").eq("id", id).single();
  if (error) throw error;
  return itemFromRow(data);
}

export async function addItem(payload, byName) {
  const item = {
    warehouse_code: payload.warehouseCode,
    code: payload.code,
    name: payload.name || payload.code,
    min_qty: null,
    colors: [{
      id: uid(),
      color: payload.color,
      sizes: payload.sizes.map((s) => ({ id: uid(), size: s.size, qty: s.qty })),
      packQty: payload.packQty,
    }],
  };
  const { data, error } = await supabase.from("items").insert(item).select().single();
  if (error) throw error;
  const saved = itemFromRow(data);

  const total = saved.colors.reduce((s, c) => s + (c.sizes || []).reduce((s2, x) => s2 + (Number(x.qty) || 0), 0), 0) +
    saved.colors.reduce((s, c) => s + (Number(c.packQty) || 0), 0);
  if (total > 0) {
    await supabase.from("transactions").insert({
      item_id: saved.id, warehouse_code: saved.warehouseCode, code: saved.code,
      color: payload.color, size: "—", kind: "size", type: "in", qty: total,
      note: "رصيد افتتاحي", by_name: byName, created_at: todayISO(),
    });
  }
  return saved;
}

export async function addColorToItem(itemId, payload) {
  const fresh = await fetchItemFresh(itemId);
  const newColor = {
    id: uid(), color: payload.color,
    sizes: payload.sizes.map((s) => ({ id: uid(), size: s.size, qty: s.qty })),
    packQty: payload.packQty,
  };
  const colors = [...fresh.colors, newColor];
  const { error } = await supabase.from("items").update({ colors, updated_at: todayISO() }).eq("id", itemId);
  if (error) throw error;
}

export async function setItemMinQty(itemId, minQty) {
  const { error } = await supabase.from("items").update({ min_qty: minQty, updated_at: todayISO() }).eq("id", itemId);
  if (error) throw error;
}

export async function applyMovement(item, colorId, kind, sizeId, type, qty, note, byName) {
  const fresh = await fetchItemFresh(item.id);
  const delta = type === "in" ? qty : -qty;
  let colorLabel = "", sizeLabel = "";
  const colors = fresh.colors.map((c) => {
    if (c.id !== colorId) return c;
    colorLabel = c.color;
    if (kind === "pack") {
      return { ...c, packQty: (c.packQty || 0) + delta };
    }
    const sizes = c.sizes.map((s) => {
      if (s.id !== sizeId) return s;
      sizeLabel = s.size;
      return { ...s, qty: (s.qty || 0) + delta };
    });
    return { ...c, sizes };
  });

  const { error } = await supabase.from("items").update({ colors, updated_at: todayISO() }).eq("id", item.id);
  if (error) throw error;

  await supabase.from("transactions").insert({
    item_id: item.id, warehouse_code: item.warehouseCode, code: item.code,
    color: colorLabel, size: sizeLabel, kind, type, qty,
    note: note || "", by_name: byName, created_at: todayISO(),
  });
}

export async function importItemsFromExcel(rows, warehouseCode, byName) {
  const { data: existingRows, error: fetchErr } = await supabase.from("items").select("*").eq("warehouse_code", warehouseCode);
  if (fetchErr) throw fetchErr;
  const existing = (existingRows || []).map(itemFromRow);

  const grouped = {};
  rows.forEach((r) => { (grouped[r.code] ||= []).push(r); });

  const toInsert = [];
  const toUpdate = [];

  Object.keys(grouped).forEach((code) => {
    const found = existing.find((i) => i.code === code);
    const colors = found ? [...found.colors] : [];
    grouped[code].forEach((row) => {
      const idx = colors.findIndex((c) => c.color === row.color);
      const newColor = {
        id: idx !== -1 ? colors[idx].id : uid(),
        color: row.color,
        sizes: row.sizes.map((s) => ({ id: uid(), size: s.size, qty: s.qty })),
        packQty: row.packQty,
      };
      if (idx !== -1) colors[idx] = newColor; else colors.push(newColor);
    });
    if (found) toUpdate.push({ id: found.id, colors });
    else toInsert.push({ warehouse_code: warehouseCode, code, name: code, min_qty: null, colors });
  });

  if (toInsert.length) {
    const { error } = await supabase.from("items").insert(toInsert);
    if (error) throw error;
  }
  for (const u of toUpdate) {
    const { error } = await supabase.from("items").update({ colors: u.colors, updated_at: todayISO() }).eq("id", u.id);
    if (error) throw error;
  }

  await supabase.from("transactions").insert({
    item_id: null, warehouse_code: warehouseCode, code: "—", color: "", size: "",
    kind: "size", type: "in", qty: rows.length,
    note: `استيراد ${rows.length} صف من Excel`, by_name: byName, created_at: todayISO(),
  });
}

// daily balance refresh: wipes every item in a warehouse and replaces it from a fresh sheet
export async function replaceWarehouseFromExcel(rows, warehouseCode, byName) {
  const { error: delErr } = await supabase.from("items").delete().eq("warehouse_code", warehouseCode);
  if (delErr) throw delErr;

  const grouped = {};
  rows.forEach((r) => { (grouped[r.code] ||= []).push(r); });
  const newItems = Object.keys(grouped).map((code) => ({
    warehouse_code: warehouseCode, code, name: code, min_qty: null,
    colors: grouped[code].map((row) => ({
      id: uid(), color: row.color,
      sizes: row.sizes.map((s) => ({ id: uid(), size: s.size, qty: s.qty })),
      packQty: row.packQty,
    })),
  }));
  if (newItems.length) {
    const { error } = await supabase.from("items").insert(newItems);
    if (error) throw error;
  }

  await supabase.from("transactions").insert({
    item_id: null, warehouse_code: warehouseCode, code: "—", color: "", size: "", kind: "size", type: "in",
    qty: newItems.length,
    note: `تحديث يومي — تم استبدال الرصيد بالكامل (${newItems.length} كود)`,
    by_name: byName, created_at: todayISO(),
  });
}

/* ---------- orders ---------- */

export async function addOrder(payload, byName) {
  const { error } = await supabase.from("orders").insert({
    factory: payload.factory, order_date: payload.date, warehouse_code: payload.warehouseCode,
    code: payload.code, po: payload.po, internal_po: "", color: payload.color,
    sizes: payload.sizes, ordered_qty: payload.orderedQty, received_qty: 0,
    remaining_at_factory: payload.orderedQty, status: "open", receipts: [], created_by: byName,
  });
  if (error) throw error;
}

export async function importOrders(rows, warehouseCode, byName) {
  const newOrders = rows.map((r) => ({
    factory: r.factory, order_date: r.date, warehouse_code: warehouseCode,
    code: r.code, po: r.po, internal_po: r.internalPo || "", color: r.color,
    sizes: r.sizes || [], ordered_qty: r.orderedQty,
    received_qty: Math.max(0, r.orderedQty - r.remainingAtFactory),
    remaining_at_factory: r.remainingAtFactory,
    status: r.remainingAtFactory > 0 ? "open" : "closed",
    receipts: [], created_by: byName,
  }));
  if (newOrders.length) {
    const { error } = await supabase.from("orders").insert(newOrders);
    if (error) throw error;
  }
  const factoryNames = Array.from(new Set(rows.map((r) => r.factory).filter(Boolean)));
  if (factoryNames.length) {
    await supabase.from("factories").upsert(factoryNames.map((name) => ({ name })), { onConflict: "name", ignoreDuplicates: true });
  }
}

export async function setOrderStatus(orderId, status) {
  const { error } = await supabase.from("orders").update({ status }).eq("id", orderId);
  if (error) throw error;
}

// records a receipt against an order: updates the order's progress AND the real warehouse stock, atomically enough for this scale
export async function receiveOrder(order, internalPo, sizeBreakdown, totalQty, note, byName) {
  const { data: freshOrderRow, error: fetchErr } = await supabase.from("orders").select("*").eq("id", order.id).single();
  if (fetchErr) throw fetchErr;
  const freshOrder = orderFromRow(freshOrderRow);

  const receipt = { id: uid(), qty: totalQty, sizeBreakdown, date: todayISO(), by: byName, note: note || "" };
  const { error: orderUpdateErr } = await supabase.from("orders").update({
    internal_po: internalPo || freshOrder.internalPo,
    received_qty: freshOrder.receivedQty + totalQty,
    remaining_at_factory: Math.max(0, freshOrder.remainingAtFactory - totalQty),
    receipts: [...freshOrder.receipts, receipt],
  }).eq("id", order.id);
  if (orderUpdateErr) throw orderUpdateErr;

  const { data: existingItemRow } = await supabase
    .from("items").select("*")
    .eq("warehouse_code", order.warehouseCode).eq("code", order.code)
    .maybeSingle();

  let item = existingItemRow ? itemFromRow(existingItemRow) : null;
  if (!item) {
    const { data: created, error: createErr } = await supabase.from("items").insert({
      warehouse_code: order.warehouseCode, code: order.code, name: order.code, min_qty: null, colors: [],
    }).select().single();
    if (createErr) throw createErr;
    item = itemFromRow(created);
  }

  let colors = [...item.colors];
  let colorIdx = colors.findIndex((c) => c.color === order.color);
  if (colorIdx === -1) {
    colors.push({ id: uid(), color: order.color, sizes: [], packQty: 0 });
    colorIdx = colors.length - 1;
  }
  const color = { ...colors[colorIdx], sizes: [...colors[colorIdx].sizes] };

  const noteLabel = `استلام أوردر — PO:${order.po || "—"} / Internal PO:${internalPo || freshOrder.internalPo || "—"}${note ? " — " + note : ""}`;
  const txRows = [];

  if (sizeBreakdown && sizeBreakdown.length) {
    sizeBreakdown.forEach(({ size, qty }) => {
      if (!qty) return;
      let sIdx = color.sizes.findIndex((s) => s.size === size);
      if (sIdx === -1) { color.sizes.push({ id: uid(), size, qty: 0 }); sIdx = color.sizes.length - 1; }
      color.sizes[sIdx] = { ...color.sizes[sIdx], qty: color.sizes[sIdx].qty + qty };
      txRows.push({
        item_id: item.id, warehouse_code: item.warehouseCode, code: item.code,
        color: order.color, size, kind: "size", type: "in", qty, note: noteLabel, by_name: byName, created_at: todayISO(),
      });
    });
  } else {
    const RECEIPT_SIZE = "استلام أوردر";
    let sIdx = color.sizes.findIndex((s) => s.size === RECEIPT_SIZE);
    if (sIdx === -1) { color.sizes.push({ id: uid(), size: RECEIPT_SIZE, qty: 0 }); sIdx = color.sizes.length - 1; }
    color.sizes[sIdx] = { ...color.sizes[sIdx], qty: color.sizes[sIdx].qty + totalQty };
    txRows.push({
      item_id: item.id, warehouse_code: item.warehouseCode, code: item.code,
      color: order.color, size: RECEIPT_SIZE, kind: "size", type: "in", qty: totalQty, note: noteLabel, by_name: byName, created_at: todayISO(),
    });
  }

  colors[colorIdx] = color;
  const { error: itemUpdateErr } = await supabase.from("items").update({ colors, updated_at: todayISO() }).eq("id", item.id);
  if (itemUpdateErr) throw itemUpdateErr;

  if (txRows.length) {
    const { error: txErr } = await supabase.from("transactions").insert(txRows);
    if (txErr) throw txErr;
  }
}
const db = {
  fetchAll,
  fetchProfiles,
  signInByName,
  signOut,
  getCurrentProfile,
  signUpFirstAdmin,
  createUser,
  deleteUser,
  addFactory,
  removeFactory,
  addItem,
  addColorToItem,
  setItemMinQty,
  applyMovement,
  importItemsFromExcel,
  replaceWarehouseFromExcel,
  addOrder,
  importOrders,
  setOrderStatus,
  receiveOrder,
};

export default db;
