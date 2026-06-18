// =============================================================================
// app.js  —  Space Global Incoming Stock (v2)
// Adds: DD-MM-YYYY dates, full descriptions, country flags, brand logos,
// product images (auto-fetched from space.ae, cached), DG / Non-DG + stock
// filters, and Excel export. Static site; no build step.
// =============================================================================

import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
// mapping.js is imported lazily inside readFile() so an issue there can never block sign-in.

const cfg = window.APP_CONFIG || {};
if (!cfg.SUPABASE_URL || cfg.SUPABASE_URL.includes("YOUR-PROJECT")) {
  const m = document.getElementById("loginMsg");
  if (m) m.textContent = "Open config.js and add your Supabase URL and key, then reload.";
}
const sb = createClient(cfg.SUPABASE_URL, cfg.SUPABASE_ANON_KEY);

// ---- helpers ---------------------------------------------------------------
const $ = (id) => document.getElementById(id);
const el = (h) => { const t = document.createElement("template"); t.innerHTML = h.trim(); return t.content.firstChild; };
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const fmt = (n) => (n === null || n === undefined || n === "") ? "—" : Number(n).toLocaleString(undefined, { maximumFractionDigits: 0 });
const money = (n) => (n === null || n === undefined || n === "" || isNaN(Number(n))) ? "—" : Number(n).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const curSym = (c) => c === "USD" || !c ? "$" : c === "AED" ? "AED " : c + " ";
// dates from the database are "YYYY-MM-DD" -> show "DD-MM-YYYY"
const fmtDMY = (iso) => { if (!iso) return "—"; const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? `${m[3]}-${m[2]}-${m[1]}` : iso; };

const FLAGS = { "UAE": "🇦🇪", "Qatar": "🇶🇦", "Bahrain": "🇧🇭", "Kuwait": "🇰🇼", "Oman": "🇴🇲", "Saudi Arabia": "🇸🇦" };
const flag = (region) => FLAGS[region] ? `<span class="flag">${FLAGS[region]}</span>` : "";
const BRAND_TINT = { AT: "#00A651", Momax: "#00AEEF", Tangem: "#111111" };
const brandColor = (b) => BRAND_TINT[b] || "#5a6b86";

const COMPANY_ID = "a0000000-0000-4000-8000-000000000001";
let state = { profile: null, charts: {}, canUpload: false };
const imgCache = {};   // sku -> url | null (null = tried, none found)

function brandCell(brand) {
  if (!brand) return "—";
  const c = brandColor(brand);
  const file = "logo-" + String(brand).toLowerCase().replace(/[^a-z0-9]/g, "") + ".png";
  return `<span class="brandcell" style="color:${c}"><img class="blogo" src="${file}" alt="" onerror="this.style.display='none'"><span class="bdot" style="background:${c}"></span>${esc(brand)}</span>`;
}
function thumb(i, row) {
  const tint = brandColor(row.brand);
  const initial = esc((row.brand || row.sku || "?").slice(0, 1).toUpperCase());
  const inner = row.image_url
    ? `<img class="thumb" src="${esc(row.image_url)}" alt="" loading="lazy" onerror="this.outerHTML='<div class=&quot;thumb ph&quot; style=&quot;background:${tint}&quot;>${initial}</div>'">`
    : `<div class="thumb ph" id="thumb-${i}" style="background:${tint}">${initial}</div>`;
  const badge = row.is_new ? `<span class="newbadge">NEW</span>` : "";
  const can = state.canUpload ? " can" : "";
  const title = state.canUpload ? ' title="Click to mark / unmark as New"' : "";
  const up = state.canUpload ? `<button class="upbtn" data-sku="${esc(row.sku)}" title="Upload an image">↑</button>` : "";
  return `<div class="thumbwrap${can}" data-id="${row.id}" data-new="${row.is_new ? 1 : 0}"${title}>${inner}${badge}${up}</div>`;
}

// ---- auth ------------------------------------------------------------------
$("loginForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const btn = $("loginBtn"), msg = $("loginMsg");
  msg.className = "msg"; msg.textContent = "";
  btn.disabled = true; btn.innerHTML = '<span class="spin"></span> Signing in…';
  const { error } = await sb.auth.signInWithPassword({ email: $("email").value.trim(), password: $("password").value });
  btn.disabled = false; btn.textContent = "Sign in";
  if (error) { msg.className = "msg err"; msg.textContent = error.message; }
});
$("signOut").addEventListener("click", async () => { await sb.auth.signOut(); location.reload(); });
sb.auth.onAuthStateChange((_e, s) => { if (s) enterApp(); });
(async () => { const { data } = await sb.auth.getSession(); if (data.session) enterApp(); })();

async function enterApp() {
  try {
    if (!$("app").classList.contains("hidden")) return;
    const { data: u } = await sb.auth.getUser();
    if (!u || !u.user) return;
    let prof = null;
    try { prof = (await sb.from("profiles").select("full_name, role, email").eq("id", u.user.id).maybeSingle()).data; } catch (_) {}
    state.profile = prof || { role: "sales", email: u.user.email };
    state.canUpload = ["admin", "operations"].includes(state.profile.role);
    const setText = (id, t) => { const e = $(id); if (e) e.textContent = t; };
    setText("whoEmail", state.profile.email || u.user.email);
    setText("whoRole", state.profile.role);
    const nu = $("navUpload"); if (nu) nu.classList.toggle("hidden", !state.canUpload);
    $("login").classList.add("hidden");
    $("app").classList.remove("hidden");
    let last = "dashboard";
    try { const v = localStorage.getItem("lastView"); if (["dashboard", "incoming", "upload"].includes(v)) last = v; } catch (_) {}
    if (last === "upload" && !state.canUpload) last = "dashboard";
    go(last);
  } catch (err) {
    const m = $("loginMsg");
    if (m) { m.className = "msg err"; m.textContent = "Signed in, but the page needs a hard refresh to finish loading."; }
  }
}

// ---- nav -------------------------------------------------------------------
document.querySelectorAll(".nav button[data-view]").forEach((b) => b.addEventListener("click", () => go(b.dataset.view)));
function go(view) {
  try { localStorage.setItem("lastView", view); } catch (_) {}
  document.querySelectorAll(".nav button[data-view]").forEach((b) => b.classList.toggle("active", b.dataset.view === view));
  ["dashboard", "incoming", "upload"].forEach((v) => { const el = $("view-" + v); if (el) el.classList.toggle("hidden", v !== view); });
  if (view === "dashboard") loadDashboard();
  if (view === "incoming") loadIncoming();
  if (view === "upload") loadUpload();
}

// ---- dashboard -------------------------------------------------------------
async function loadDashboard() {
  const [{ data: k }, { data: oos }, { data: brand }, { data: eta }, { data: delayed }] = await Promise.all([
    sb.from("v_dashboard_kpis").select("*").maybeSingle(),
    sb.from("v_factory_out_of_stock").select("sku,product_name,brand,region"),
    sb.from("v_incoming_by_brand").select("brand,quantity,lines"),
    sb.from("v_eta_buckets").select("bucket"),
    sb.from("v_incoming_stock").select("sku,product_name,brand,region,po_number,eta,po_deadline,is_new").eq("is_delayed", true).order("eta", { nullsFirst: true }).limit(40),
  ]);
  const cards = [
    ["Total incoming", k?.total_incoming_qty, ""], ["Open POs", k?.open_purchase_orders, ""],
    ["Delayed", k?.delayed_shipments, "attn"], ["Out of stock", k?.factory_out_of_stock_items, "bad"],
    ["Arriving ≤7 days", k?.arriving_this_week, ""], ["Arriving this month", k?.arriving_this_month, ""],
  ];
  $("kpis").innerHTML = cards.map(([l, v, c]) => `<div class="kpi glass ${c}"><span class="edge"></span><div class="k">${l}</div><div class="v">${fmt(v)}</div></div>`).join("");
  $("dashMeta").textContent = "updated " + new Date().toLocaleString();

  $("delayCount").textContent = (delayed?.length || 0) + (delayed?.length === 40 ? "+" : "");
  $("delayTable").innerHTML = delayed?.length ? table(["SKU", "Product", "Brand", "Region", "PO", "ETA"],
    delayed.map((r) => [
      `<td class="code">${esc(r.sku)}</td>`, `<td class="desc">${r.is_new ? '<span class="chip newon" style="margin-right:6px">NEW</span>' : ""}${esc(r.product_name)}</td>`,
      `<td>${brandCell(r.brand)}</td>`, `<td>${flag(r.region)}${esc(r.region || "")}</td>`,
      `<td class="code">${esc(r.po_number)}</td>`, `<td class="code">${fmtDMY(r.eta || r.po_deadline)}</td>`,
    ]), { sticky: true }) : emptyState("Nothing delayed", "Everything on track.");

  $("oosCount").textContent = oos?.length || 0;
  $("oosTable").innerHTML = oos?.length ? table(["SKU", "Product", "Brand", "Region"],
    oos.map((r) => [
      `<td class="code">${esc(r.sku)}</td>`, `<td class="desc">${esc(r.product_name)}</td>`,
      `<td>${brandCell(r.brand)}</td>`, `<td>${flag(r.region)}${esc(r.region || "")}</td>`,
    ]), { sticky: true }) : emptyState("Nothing out of stock", "No factory shortages right now.");

  drawEtaChart(eta || []); drawBrandChart(brand || []);
}
function drawEtaChart(rows) {
  const order = ["delayed", "this_week", "next_week", "this_month", "later"];
  const labels = { delayed: "Delayed", this_week: "≤7 days", next_week: "8–14 days", this_month: "This month", later: "Later" };
  const counts = Object.fromEntries(order.map((o) => [o, 0]));
  rows.forEach((r) => { if (r.bucket in counts) counts[r.bucket]++; });
  const colors = { delayed: "#b9751a", this_week: "#2b5f8e", next_week: "#3a78ad", this_month: "#2f7d5b", later: "#9aa3b2" };
  chart("etaChart", "bar", { labels: order.map((o) => labels[o]), datasets: [{ data: order.map((o) => counts[o]), backgroundColor: order.map((o) => colors[o]), borderRadius: 6 }] },
    { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { precision: 0 } } } });
}
function drawBrandChart(rows) {
  const labels = rows.map((r) => r.brand || "—");
  const values = rows.map((r) => Number(r.quantity || 0));
  const colors = rows.map((r) => brandColor(r.brand));
  // inline plugin: print each slice's quantity on the arc
  const sliceNumbers = {
    id: "sliceNumbers",
    afterDraw(c) {
      const { ctx } = c; const meta = c.getDatasetMeta(0);
      ctx.save(); ctx.font = "700 12px Inter, sans-serif"; ctx.textAlign = "center"; ctx.textBaseline = "middle";
      meta.data.forEach((arc, idx) => {
        const v = values[idx]; if (!v) return;
        const p = arc.getCenterPoint(); ctx.fillStyle = "#fff";
        ctx.fillText(v.toLocaleString(), p.x, p.y);
      });
      ctx.restore();
    },
  };
  chart("brandChart", "doughnut",
    { labels, datasets: [{ data: values, backgroundColor: colors, borderWidth: 2, borderColor: "rgba(255,255,255,.85)" }] },
    { plugins: { legend: { position: "bottom", labels: { boxWidth: 12, font: { size: 12, weight: "600" } } } }, cutout: "58%" },
    [sliceNumbers]);
}
function chart(id, type, data, options, plugins = []) {
  if (typeof Chart === "undefined") return;
  if (state.charts[id]) state.charts[id].destroy();
  state.charts[id] = new Chart($(id), { type, data, options: { responsive: true, maintainAspectRatio: false, ...options }, plugins });
}

// ---- incoming --------------------------------------------------------------
let incInit = false, lastRows = [];
async function loadIncoming() {
  if (!incInit) {
    incInit = true;
    const [{ data: brands }, { data: regions }] = await Promise.all([
      sb.from("brands").select("name").order("name"), sb.from("regions").select("name").order("name"),
    ]);
    (brands || []).forEach((b) => $("incBrand").appendChild(el(`<option>${esc(b.name)}</option>`)));
    (regions || []).forEach((r) => $("incRegion").appendChild(el(`<option>${esc(r.name)}</option>`)));
    ["incSearch", "incBrand", "incRegion", "incFactory", "incMode", "incDelayed"].forEach((id) => $(id).addEventListener("input", debounce(runIncoming, 250)));
    $("incExportView").addEventListener("click", () => exportRows(lastRows, "incoming-view"));
    $("incExportAll").addEventListener("click", exportAll);
    $("incFetchImg").addEventListener("click", fetchImages);
    $("incFetchImg").classList.toggle("hidden", !state.canUpload);
    $("incTable").addEventListener("click", onIncTableClick);
    bindImageUpload();
  }
  runIncoming();
}

function rowById(id) { return lastRows.find((r) => String(r.id) === String(id)); }

async function onIncTableClick(e) {
  if (!state.canUpload) return;
  const upb = e.target.closest(".upbtn");
  if (upb) {
    e.stopPropagation();
    pendingImgSku = upb.dataset.sku;
    $("imgFileInput").click();
    return;
  }
  const actNew = e.target.closest(".act.new");
  if (actNew) {
    const id = actNew.dataset.id; const turnOn = actNew.dataset.new !== "1";
    const { error } = await sb.rpc("app_set_line_tags", { p_id: id, p_is_new: turnOn });
    if (error) { alert(error.message); return; }
    const row = rowById(id); if (row) row.is_new = turnOn;
    runIncoming();
    return;
  }
  const actEdit = e.target.closest(".act.edit");
  if (actEdit) { openEditModal(actEdit.dataset.id); return; }
  const actDel = e.target.closest(".act.del");
  if (actDel) {
    const id = actDel.dataset.id; const row = rowById(id);
    if (!confirm(`Delete this line${row ? ` (${row.sku} · ${row.region})` : ""}? It will be removed from the dashboard.`)) return;
    const { error } = await sb.rpc("app_delete_line", { p_id: id });
    if (error) { alert(error.message); return; }
    runIncoming();
    return;
  }
  const tw = e.target.closest(".thumbwrap.can");
  if (tw) {
    const id = tw.dataset.id;
    const turnOn = tw.dataset.new !== "1";
    const { error } = await sb.rpc("app_set_line_tags", { p_id: id, p_is_new: turnOn });
    if (error) { alert(error.message); return; }
    const row = rowById(id); if (row) row.is_new = turnOn;
    tw.dataset.new = turnOn ? "1" : "0";
    const existing = tw.querySelector(".newbadge");
    if (turnOn && !existing) tw.insertAdjacentHTML("beforeend", '<span class="newbadge">NEW</span>');
    if (!turnOn && existing) existing.remove();
    return;
  }
  const cell = e.target.closest(".remcell");
  if (cell && !cell.querySelector("input")) startRemarkEdit(cell);
}

// ---- edit modal ------------------------------------------------------------
let editId = null, editBound = false;
function openEditModal(id) {
  const r = rowById(id); if (!r) return;
  editId = id;
  $("edIdent").textContent = `${r.sku} · ${r.brand || ""} · ${r.region || ""} · ${r.po_number || ""}`;
  $("edName").value = r.product_name || "";
  $("edEan").value = r.ean || "";
  $("edQty").value = r.ordered_quantity ?? "";
  $("edDeadline").value = (r.po_deadline || r.eta || "").slice(0, 10);
  $("edMode").value = r.ship_mode || "";
  $("edFactory").value = r.factory_status || "";
  $("edRemarks").value = r.remarks || "";
  if (!editBound) {
    editBound = true;
    $("edCancel").addEventListener("click", closeEditModal);
    $("editModal").addEventListener("click", (ev) => { if (ev.target.id === "editModal") closeEditModal(); });
    $("edSave").addEventListener("click", saveEditModal);
  }
  $("editModal").classList.remove("hidden");
}
function closeEditModal() { $("editModal").classList.add("hidden"); editId = null; }
async function saveEditModal() {
  if (!editId) return;
  const btn = $("edSave"); btn.disabled = true; btn.innerHTML = '<span class="spin"></span> Saving…';
  const { error } = await sb.rpc("app_update_line", {
    p_id: editId, p_qty: $("edQty").value === "" ? null : Number($("edQty").value),
    p_deadline: $("edDeadline").value || null, p_ship_mode: $("edMode").value,
    p_factory_status: $("edFactory").value, p_remarks: $("edRemarks").value,
    p_name: $("edName").value, p_ean: $("edEan").value });
  btn.disabled = false; btn.textContent = "Save changes";
  if (error) { alert(error.message); return; }
  closeEditModal();
  runIncoming();
}

function startRemarkEdit(cell) {
  const id = cell.dataset.id;
  const current = (rowById(id)?.remarks) || "";
  cell.innerHTML = `<input class="reminput" value="${esc(current)}" placeholder="Type a note…" />`;
  const input = cell.querySelector("input");
  input.focus(); input.select();
  let done = false;
  const save = async () => {
    if (done) return; done = true;
    const val = input.value;
    cell.innerHTML = '<span class="remhint">saving…</span>';
    const { error } = await sb.rpc("app_set_line_tags", { p_id: id, p_remarks: val });
    if (error) { cell.innerHTML = remCell(current); alert(error.message); return; }
    const row = rowById(id); if (row) row.remarks = val;
    cell.innerHTML = remCell(val);
  };
  input.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter") { ev.preventDefault(); save(); }
    else if (ev.key === "Escape") { done = true; cell.innerHTML = remCell(current); }
  });
  input.addEventListener("blur", save);
}
function incomingQuery() {
  let q = sb.from("v_incoming_stock").select("id,sku,product_name,ean,hs_code,brand,region,po_number,voucher_type,ship_mode,ordered_quantity,pending_quantity,unit_cost,total_value,currency,po_deadline,eta,order_date,factory_status,shipment_status,is_delayed,is_new,remarks,image_url").order("order_date", { ascending: false, nullsFirst: false }).order("po_number", { ascending: false });
  const term = $("incSearch").value.trim();
  if (term) { const t = term.replace(/[,%]/g, " "); q = q.or(`sku.ilike.%${t}%,product_name.ilike.%${t}%,po_number.ilike.%${t}%,ean.ilike.%${t}%`); }
  if ($("incBrand").value) q = q.eq("brand", $("incBrand").value);
  if ($("incRegion").value) q = q.eq("region", $("incRegion").value);
  if ($("incFactory").value) q = q.eq("factory_status", $("incFactory").value);
  if ($("incMode").value) q = q.eq("ship_mode", $("incMode").value);
  if ($("incDelayed").checked) q = q.eq("is_delayed", true);
  return q;
}
async function runIncoming() {
  const { data, error } = await incomingQuery().limit(1000);
  if (error) { $("incTable").innerHTML = emptyState("Couldn’t load", error.message); return; }
  lastRows = data;
  $("incMeta").textContent = `${data.length} line${data.length === 1 ? "" : "s"}${data.length === 1000 ? " (first 1000)" : ""}`;
  if (!data.length) { $("incTable").innerHTML = emptyState("No matching lines", "Try clearing the filters."); return; }
  const showActions = state.canUpload;
  const sym = curSym((data.find((r) => r.currency) || {}).currency);
  const sumFob = data.reduce((a, r) => a + (Number(r.unit_cost) || 0), 0);
  const sumVal = data.reduce((a, r) => a + (Number(r.total_value) || 0), 0);
  const headers = ["", "SKU", "Product", "Brand", "Region", "PO", "Qty", "FOB", "Total value", "ETA", "Mode", "Remarks"];
  const aligns = ["", "", "", "", "", "c", "c", "r", "r", "c", "", ""];
  if (showActions) { headers.push("Actions"); aligns.push("r"); }
  const restCols = showActions ? 4 : 3;
  const foot =
    `<td colspan="7" class="ftot-label">Totals · ${data.length} line${data.length === 1 ? "" : "s"}</td>` +
    `<td class="ftot">${sym}${money(sumFob)}</td>` +
    `<td class="ftot">${sym}${money(sumVal)}</td>` +
    `<td colspan="${restCols}"></td>`;
  $("incTable").innerHTML = table(
    headers,
    data.map((r, i) => [
      `<td>${thumb(i, r)}</td>`,
      `<td class="code">${esc(r.sku)}</td>`,
      `<td class="desc" title="${esc(r.product_name)}">${esc(r.product_name)}</td>`,
      `<td>${brandCell(r.brand)}</td>`,
      `<td>${flag(r.region)}${esc(r.region || "")}</td>`,
      `<td class="code c">${esc(r.po_number)}</td>`,
      `<td class="code c">${fmt(r.ordered_quantity)}</td>`,
      `<td class="code money">${r.unit_cost == null ? "—" : sym + money(r.unit_cost)}</td>`,
      `<td class="code money">${r.total_value == null ? "—" : sym + money(r.total_value)}</td>`,
      `<td class="code c">${r.is_delayed ? `<span class="chip delay">${fmtDMY(r.eta || r.po_deadline)}</span>` : fmtDMY(r.eta || r.po_deadline)}</td>`,
      `<td>${modeChip(r.ship_mode)}</td>`,
      `<td class="remcell" data-id="${r.id}">${remCell(r.remarks)}</td>`,
      showActions
        ? `<td class="actions">
             <button class="act new ${r.is_new ? "on" : ""}" data-id="${r.id}" data-new="${r.is_new ? 1 : 0}" title="${r.is_new ? "Unmark New" : "Mark as New"}">${r.is_new ? "★" : "☆"}</button>
             <button class="act edit" data-id="${r.id}" title="Edit">✎</button>
             <button class="act del" data-id="${r.id}" title="Delete">🗑</button>
           </td>`
        : "",
    ]),
    { classes: data.map((r) => (r.is_delayed ? "delayed" : "")), aligns, foot }
  );
}
function remCell(text) {
  if (text) return esc(text);
  return state.canUpload ? `<span class="remhint">add note…</span>` : "";
}
function factoryChip(s) {
  if (!s) return "";
  const m = { out_of_stock: ["oos", "out of stock"], available: ["avail", "available"], limited_stock: ["limited", "limited"], discontinued: ["neutral", "discont."] };
  const [cls, label] = m[s] || ["neutral", s];
  return `<span class="chip ${cls}">${esc(label)}</span>`;
}
function modeChip(s) {
  if (!s) return "";
  const up = String(s).toUpperCase();
  if (up === "DG") return `<span class="chip dg">DG</span>`;
  if (up.includes("NON")) return `<span class="chip nondg">NON-DG</span>`;
  return `<span class="chip neutral">${esc(s)}</span>`;
}

// ---- product images -------------------------------------------------------
// Images are fetched server-side (space.ae blocks direct browser calls) and
// cached in products.image_url. This button fills in any that are missing.
async function fetchImages() {
  const btn = $("incFetchImg"); if (!btn) return;
  const original = btn.textContent;
  btn.disabled = true;
  let updated = 0, iterations = 0;
  try {
    while (iterations++ < 50) {
      btn.innerHTML = `<span class="spin"></span> ${updated} found…`;
      const { data, error } = await sb.rpc("app_fetch_images", { p_limit: 5 });
      if (error) { alert(error.message); break; }
      updated += data.updated || 0;
      if ((data.tried || 0) < 5) break; // no more products left without an image
    }
  } finally {
    btn.disabled = false; btn.textContent = original;
  }
  await runIncoming();
  alert(`Done — found ${updated} image${updated === 1 ? "" : "s"}. Products without a match on space.ae keep their colored tile.`);
}

// ---- export ----------------------------------------------------------------
function rowsToSheet(rows, columns) {
  return rows.map((r) => {
    const o = {};
    for (const [key, label] of columns) {
      let v = r[key];
      if (["eta", "po_deadline", "order_date", "original_eta"].includes(key)) v = v ? fmtDMY(v) : "";
      o[label] = v ?? "";
    }
    return o;
  });
}
const VIEW_COLS = [["sku", "SKU"], ["product_name", "Product"], ["ean", "Barcode"], ["brand", "Brand"], ["region", "Region"], ["po_number", "PO"], ["ordered_quantity", "Qty"], ["unit_cost", "Unit cost"], ["total_value", "Total value"], ["po_deadline", "Deadline"], ["eta", "ETA"], ["factory_status", "Factory status"], ["ship_mode", "Ship mode"], ["remarks", "Remarks"]];
function downloadSheet(data, columns, filename) {
  if (typeof XLSX === "undefined") { alert("Spreadsheet library not loaded."); return; }
  if (!data.length) { alert("Nothing to export with the current filters."); return; }
  const ws = XLSX.utils.json_to_sheet(rowsToSheet(data, columns));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Incoming");
  XLSX.writeFile(wb, `${filename}_${new Date().toISOString().slice(0, 10)}.xlsx`);
}
function exportRows(rows, filename) { downloadSheet(rows, VIEW_COLS, filename); }
async function exportAll() {
  const { data, error } = await sb.from("v_incoming_stock").select("*").limit(100000);
  if (error) { alert(error.message); return; }
  const ALL = Object.keys(data[0] || {}).filter((k) => !["id", "company_id"].includes(k)).map((k) => [k, k]);
  downloadSheet(data || [], ALL.length ? ALL : VIEW_COLS, "incoming-all");
}

// ---- upload ----------------------------------------------------------------
let upState = { mappings: [], parsed: null };
async function loadUpload() {
  if (!upState.mappings.length) {
    const { data } = await sb.from("column_mappings").select("id,name,mapping").order("name");
    upState.mappings = data || [];
    $("upMapping").innerHTML = upState.mappings.map((m, i) => `<option value="${i}">${esc(m.name)}</option>`).join("");
    $("upRead").addEventListener("click", readFile);
  }
  setupAddForm();
  loadUploadLog();
}
async function readFile() {
  const f = $("upFile").files[0], prev = $("upPreview");
  if (!f) { prev.innerHTML = banner("warn", "Choose an Excel file first."); return; }
  const mapping = upState.mappings[$("upMapping").value]?.mapping;
  if (!mapping) { prev.innerHTML = banner("err", "Pick a template."); return; }
  prev.innerHTML = '<div class="empty"><span class="spin" style="border-color:#16233a40;border-top-color:#16233a"></span> Reading…</div>';
  try {
    const wb = XLSX.read(await f.arrayBuffer(), { cellDates: true });
    const { mapWorkbook } = await import("./mapping.js?v=14");
    const { rows, report } = mapWorkbook(wb, mapping, XLSX);
    upState.parsed = { fileName: f.name, rows };
    if (!rows.length) {
      prev.innerHTML = banner("warn", `No rows found. Template expects sheet(s): <b>${esc(Object.keys(mapping.sheets || {}).join(", "))}</b>. This file has: <b>${esc(report.sheetsSeen.join(", "))}</b>.`) + `<div class="note">If the sheet names differ, tell me and I’ll adjust the template.</div>`;
      return;
    }
    const per = Object.entries(report.perSheet).map(([s, n]) => `${esc(s)}: <b>${n}</b>`).join(" &nbsp;·&nbsp; ");
    const oos = rows.filter((r) => r.factory_status === "out_of_stock").length;
    prev.innerHTML = `<div class="stat-row"><div class="s"><b>${rows.length}</b>rows ready</div><div class="s"><b>${oos}</b>out of stock</div></div><div class="note">From — ${per}</div><button class="btn full" id="upImport" style="margin-top:14px">Import ${rows.length} rows</button><div id="upResult"></div>`;
    $("upImport").addEventListener("click", doImport);
  } catch (e) { prev.innerHTML = banner("err", "Could not read this file: " + esc(e.message)); }
}
async function doImport() {
  const btn = $("upImport"), res = $("upResult");
  btn.disabled = true; btn.innerHTML = '<span class="spin"></span> Importing…';
  const { data, error } = await sb.rpc("app_import_rows", { p_file_name: upState.parsed.fileName, p_rows: upState.parsed.rows });
  btn.disabled = false; btn.textContent = "Import again";
  if (error) { res.innerHTML = banner("err", error.message); return; }
  res.innerHTML = banner("ok", `Done — <b>${data.created}</b> new, <b>${data.updated}</b> updated, <b>${data.unchanged}</b> unchanged${data.removed ? `, <b>${data.removed}</b> removed (no longer in sheet)` : ""} of <b>${data.total}</b> rows.`);
  loadUploadLog();
}
async function loadUploadLog() {
  const { data } = await sb.from("upload_logs").select("file_name,status,total_rows,created_count,updated_count,unchanged_count,started_at").order("started_at", { ascending: false }).limit(10);
  $("upLog").innerHTML = data?.length ? table(["File", "When", "Rows", "New", "Updated", "Same", "Status"],
    data.map((r) => [
      `<td>${esc(r.file_name)}</td>`, `<td class="code">${esc(new Date(r.started_at).toLocaleString())}</td>`,
      `<td class="code c">${fmt(r.total_rows)}</td>`, `<td class="code c">${fmt(r.created_count)}</td>`,
      `<td class="code c">${fmt(r.updated_count)}</td>`, `<td class="code c">${fmt(r.unchanged_count)}</td>`,
      `<td><span class="chip ${r.status === "completed" ? "avail" : "neutral"}">${esc(r.status)}</span></td>`,
    ]), { aligns: ["", "", "c", "c", "c", "c", ""] }) : emptyState("No uploads yet", "Your first import will appear here.");
}

// ---- manual add ------------------------------------------------------------
let addInit = false;
async function setupAddForm() {
  if (addInit || !$("adBrand")) return;
  addInit = true;
  const [{ data: brands }, { data: regions }] = await Promise.all([
    sb.from("brands").select("name").order("name"), sb.from("regions").select("name").order("name"),
  ]);
  $("adBrand").innerHTML = '<option value="">Select…</option>' + (brands || []).map((b) => `<option>${esc(b.name)}</option>`).join("");
  $("adRegion").innerHTML = '<option value="">Select…</option>' + (regions || []).map((r) => `<option>${esc(r.name)}</option>`).join("");
  $("adSave").addEventListener("click", saveAdd);
}
async function saveAdd() {
  const v = (id) => $(id).value.trim();
  const f = { brand: v("adBrand"), region: v("adRegion"), po: v("adPo"), sku: v("adSku"), name: v("adName"),
    ean: v("adEan"), qty: v("adQty"), deadline: v("adDeadline"), mode: v("adMode"), factory: v("adFactory"), remarks: v("adRemarks") };
  const missing = Object.entries(f).filter(([, val]) => val === "").map(([k]) => k);
  if (missing.length) { $("adResult").innerHTML = banner("warn", "Please fill in every field before saving."); return; }
  const btn = $("adSave"); btn.disabled = true; btn.innerHTML = '<span class="spin"></span> Saving…';
  const { error } = await sb.rpc("app_add_line", {
    p_brand: f.brand, p_region: f.region, p_po: f.po, p_sku: f.sku, p_name: f.name, p_ean: f.ean,
    p_qty: Number(f.qty), p_deadline: f.deadline, p_ship_mode: f.mode, p_factory_status: f.factory, p_remarks: f.remarks });
  btn.disabled = false; btn.textContent = "Add to incoming stock";
  if (error) { $("adResult").innerHTML = banner("err", error.message); return; }
  $("adResult").innerHTML = banner("ok", `Added <b>${esc(f.sku)}</b> (${esc(f.region)}) to incoming stock.`);
  ["adPo", "adSku", "adName", "adEan", "adQty", "adDeadline", "adRemarks"].forEach((id) => ($(id).value = ""));
  incInit && runIncoming();
}

// ---- manual image upload ---------------------------------------------------
let pendingImgSku = null;
function bindImageUpload() {
  const input = $("imgFileInput");
  if (!input || input.dataset.bound) return;
  input.dataset.bound = "1";
  input.addEventListener("change", async () => {
    const file = input.files[0]; const sku = pendingImgSku; input.value = ""; pendingImgSku = null;
    if (!file || !sku) return;
    try {
      const ext = (file.name.split(".").pop() || "jpg").toLowerCase().replace(/[^a-z0-9]/g, "");
      const path = `${sku.replace(/[^A-Za-z0-9_.-]/g, "_")}-${Date.now()}.${ext}`;
      const up = await sb.storage.from("product-images").upload(path, file, { upsert: true, contentType: file.type });
      if (up.error) throw up.error;
      const { data: pub } = sb.storage.from("product-images").getPublicUrl(path);
      const url = pub.publicUrl;
      const { error } = await sb.rpc("app_set_product_image", { p_sku: sku, p_url: url });
      if (error) throw error;
      lastRows.forEach((r) => { if (r.sku === sku) r.image_url = url; });
      runIncoming();
    } catch (e) { alert("Image upload failed: " + (e.message || e)); }
  });
}


function table(headers, rows, opts = {}) {
  const { sticky = false, classes = [], aligns = [], foot = null } = opts;
  const head = headers.map((h, i) => `<th${aligns[i] === "c" ? ' style="text-align:center"' : aligns[i] === "r" ? ' style="text-align:right"' : ""}>${esc(h)}</th>`).join("");
  const body = rows.map((cells, i) => `<tr class="${classes[i] || ""}">${cells.join("")}</tr>`).join("");
  const footer = foot ? `<tfoot><tr>${foot}</tr></tfoot>` : "";
  return `<div class="tablewrap" style="${sticky ? "max-height:360px" : ""}"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody>${footer}</table></div>`;
}
function emptyState(t, s) { return `<div class="empty"><b>${esc(t)}</b>${esc(s)}</div>`; }
function banner(kind, html) { return `<div class="banner ${kind}">${html}</div>`; }
function debounce(fn, ms) { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; }
