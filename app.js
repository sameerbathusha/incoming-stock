// =============================================================================
// app.js  —  Space Global Incoming Stock (lean build)
// Talks to Supabase for login + data, parses Excel in the browser, and calls
// the app_import_rows() database function to load it. No build step required.
// =============================================================================

import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
import { mapWorkbook } from "./mapping.js";

const cfg = window.APP_CONFIG || {};
if (!cfg.SUPABASE_URL || cfg.SUPABASE_URL.includes("YOUR-PROJECT")) {
  document.getElementById("loginMsg").textContent =
    "Open config.js and add your Supabase URL and key, then reload.";
}
const sb = createClient(cfg.SUPABASE_URL, cfg.SUPABASE_ANON_KEY);

// ---- tiny helpers ----------------------------------------------------------
const $ = (id) => document.getElementById(id);
const el = (html) => { const t = document.createElement("template"); t.innerHTML = html.trim(); return t.content.firstChild; };
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const fmt = (n) => (n === null || n === undefined || n === "") ? "—" : Number(n).toLocaleString(undefined, { maximumFractionDigits: 0 });
const fmtDate = (d) => d ? d : "—";
let state = { profile: null, charts: {} };

// ---- auth ------------------------------------------------------------------
$("loginForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const btn = $("loginBtn"); const msg = $("loginMsg");
  msg.className = "msg"; msg.textContent = "";
  btn.disabled = true; btn.innerHTML = '<span class="spin"></span> Signing in…';
  const { error } = await sb.auth.signInWithPassword({ email: $("email").value.trim(), password: $("password").value });
  btn.disabled = false; btn.textContent = "Sign in";
  if (error) { msg.className = "msg err"; msg.textContent = error.message; }
});

$("signOut").addEventListener("click", async () => { await sb.auth.signOut(); location.reload(); });

sb.auth.onAuthStateChange((_evt, session) => { if (session) enterApp(); });

(async () => {
  const { data } = await sb.auth.getSession();
  if (data.session) enterApp();
})();

async function enterApp() {
  if (!$("app").classList.contains("hidden")) return; // already in
  const { data: u } = await sb.auth.getUser();
  const { data: prof } = await sb.from("profiles").select("full_name, role, email").eq("id", u.user.id).maybeSingle();
  state.profile = prof || { role: "sales", email: u.user.email };
  $("whoEmail").textContent = state.profile.email || u.user.email;
  $("whoRole").textContent = state.profile.role;
  const canUpload = ["admin", "operations"].includes(state.profile.role);
  $("navUpload").classList.toggle("hidden", !canUpload);
  $("login").classList.add("hidden");
  $("app").classList.remove("hidden");
  go("dashboard");
}

// ---- navigation ------------------------------------------------------------
document.querySelectorAll(".nav button[data-view]").forEach((b) =>
  b.addEventListener("click", () => go(b.dataset.view))
);
function go(view) {
  document.querySelectorAll(".nav button[data-view]").forEach((b) => b.classList.toggle("active", b.dataset.view === view));
  ["dashboard", "incoming", "upload"].forEach((v) => $("view-" + v).classList.toggle("hidden", v !== view));
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
    sb.from("v_incoming_stock").select("sku,product_name,brand,region,po_number,eta,po_deadline,is_delayed").eq("is_delayed", true).order("eta", { nullsFirst: true }).limit(40),
  ]);

  const cards = [
    ["Total incoming", k?.total_incoming_qty, ""],
    ["Open POs", k?.open_purchase_orders, ""],
    ["Delayed", k?.delayed_shipments, "attn"],
    ["Out of stock", k?.factory_out_of_stock_items, "bad"],
    ["Arriving ≤7 days", k?.arriving_this_week, ""],
    ["Arriving this month", k?.arriving_this_month, ""],
  ];
  $("kpis").innerHTML = cards.map(([label, val, cls]) =>
    `<div class="kpi ${cls}"><span class="edge"></span><div class="k">${label}</div><div class="v">${fmt(val)}</div></div>`
  ).join("");
  $("dashMeta").textContent = "updated " + new Date().toLocaleString();

  // delayed table
  $("delayCount").textContent = (delayed?.length || 0) + (delayed?.length === 40 ? "+" : "");
  $("delayTable").innerHTML = delayed?.length ? table(
    ["SKU", "Product", "Brand", "PO", "ETA"],
    delayed.map((r) => [
      `<td class="code">${esc(r.sku)}</td>`,
      `<td>${esc(trim(r.product_name, 46))}</td>`,
      `<td>${esc(r.brand)}</td>`,
      `<td class="code">${esc(r.po_number)}</td>`,
      `<td class="code">${esc(fmtDate(r.eta || r.po_deadline))}</td>`,
    ]), { sticky: true }
  ) : emptyState("Nothing delayed", "Everything on track.");

  // oos table
  $("oosCount").textContent = oos?.length || 0;
  $("oosTable").innerHTML = oos?.length ? table(
    ["SKU", "Product", "Brand", "Region"],
    oos.map((r) => [
      `<td class="code">${esc(r.sku)}</td>`,
      `<td>${esc(trim(r.product_name, 40))}</td>`,
      `<td>${esc(r.brand)}</td>`,
      `<td>${esc(r.region)}</td>`,
    ])
  ) : emptyState("Nothing out of stock", "No factory shortages right now.");

  drawEtaChart(eta || []);
  drawBrandChart(brand || []);
}

function drawEtaChart(rows) {
  const order = ["delayed", "this_week", "next_week", "this_month", "later"];
  const labels = { delayed: "Delayed", this_week: "≤7 days", next_week: "8–14 days", this_month: "This month", later: "Later" };
  const counts = Object.fromEntries(order.map((o) => [o, 0]));
  rows.forEach((r) => { if (r.bucket in counts) counts[r.bucket]++; });
  const colors = { delayed: "#b9751a", this_week: "#2b5f8e", next_week: "#3a78ad", this_month: "#2f7d5b", later: "#9aa3b2" };
  chart("etaChart", "bar", {
    labels: order.map((o) => labels[o]),
    datasets: [{ data: order.map((o) => counts[o]), backgroundColor: order.map((o) => colors[o]), borderRadius: 5 }],
  }, { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { precision: 0 } } } });
}

function drawBrandChart(rows) {
  chart("brandChart", "doughnut", {
    labels: rows.map((r) => r.brand || "—"),
    datasets: [{ data: rows.map((r) => Number(r.quantity || 0)), backgroundColor: ["#16233a", "#2b5f8e", "#2f7d5b", "#b9751a", "#7b5ea7", "#b23029"], borderWidth: 0 }],
  }, { plugins: { legend: { position: "bottom", labels: { boxWidth: 12, font: { size: 11 } } } }, cutout: "62%" });
}

function chart(id, type, data, options) {
  if (typeof Chart === "undefined") return;
  if (state.charts[id]) state.charts[id].destroy();
  state.charts[id] = new Chart($(id), { type, data, options: { responsive: true, maintainAspectRatio: false, ...options } });
}

// ---- incoming --------------------------------------------------------------
let incFiltersInit = false;
async function loadIncoming() {
  if (!incFiltersInit) {
    incFiltersInit = true;
    const [{ data: brands }, { data: regions }] = await Promise.all([
      sb.from("brands").select("name").order("name"),
      sb.from("regions").select("name").order("name"),
    ]);
    (brands || []).forEach((b) => $("incBrand").appendChild(el(`<option>${esc(b.name)}</option>`)));
    (regions || []).forEach((r) => $("incRegion").appendChild(el(`<option>${esc(r.name)}</option>`)));
    ["incSearch", "incBrand", "incRegion", "incDelayed"].forEach((id) =>
      $(id).addEventListener("input", debounce(runIncoming, 250)));
    $("incRefresh").addEventListener("click", runIncoming);
  }
  runIncoming();
}

async function runIncoming() {
  let q = sb.from("v_incoming_stock")
    .select("sku,product_name,ean,brand,region,po_number,ship_mode,ordered_quantity,unit_cost,total_value,po_deadline,eta,factory_status,shipment_status,is_delayed,remarks")
    .order("eta", { nullsFirst: false }).limit(1000);
  const term = $("incSearch").value.trim();
  if (term) {
    const t = term.replace(/[,%]/g, " ");
    q = q.or(`sku.ilike.%${t}%,product_name.ilike.%${t}%,po_number.ilike.%${t}%,ean.ilike.%${t}%`);
  }
  if ($("incBrand").value) q = q.eq("brand", $("incBrand").value);
  if ($("incRegion").value) q = q.eq("region", $("incRegion").value);
  if ($("incDelayed").checked) q = q.eq("is_delayed", true);

  const { data, error } = await q;
  if (error) { $("incTable").innerHTML = emptyState("Couldn’t load", error.message); return; }
  $("incMeta").textContent = `${data.length} line${data.length === 1 ? "" : "s"}${data.length === 1000 ? " (showing first 1000)" : ""}`;
  if (!data.length) { $("incTable").innerHTML = emptyState("No matching lines", "Try clearing the filters."); return; }

  $("incTable").innerHTML = table(
    ["SKU", "Product", "Brand", "Region", "PO", "Qty", "ETA", "Factory", "Mode"],
    data.map((r) => [
      `<td class="code">${esc(r.sku)}</td>`,
      `<td>${esc(trim(r.product_name, 48))}</td>`,
      `<td>${esc(r.brand)}</td>`,
      `<td>${esc(r.region)}</td>`,
      `<td class="code">${esc(r.po_number)}</td>`,
      `<td class="n">${fmt(r.ordered_quantity)}</td>`,
      `<td class="code">${r.is_delayed ? `<span class="chip delay">${esc(fmtDate(r.eta || r.po_deadline))}</span>` : esc(fmtDate(r.eta || r.po_deadline))}</td>`,
      `<td>${factoryChip(r.factory_status)}</td>`,
      `<td class="code">${esc(r.ship_mode || "")}</td>`,
    ]),
    { classes: data.map((r) => (r.is_delayed ? "delayed" : "")) }
  );
}

function factoryChip(s) {
  if (!s) return "";
  const map = { out_of_stock: ["oos", "out of stock"], available: ["avail", "available"], limited_stock: ["limited", "limited"], discontinued: ["neutral", "discont."] };
  const [cls, label] = map[s] || ["neutral", s];
  return `<span class="chip ${cls}">${esc(label)}</span>`;
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
  loadUploadLog();
}

async function readFile() {
  const f = $("upFile").files[0];
  const prev = $("upPreview");
  if (!f) { prev.innerHTML = banner("warn", "Choose an Excel file first."); return; }
  const mapping = upState.mappings[$("upMapping").value]?.mapping;
  if (!mapping) { prev.innerHTML = banner("err", "Pick a template."); return; }
  prev.innerHTML = '<div class="empty"><span class="spin" style="border-top-color:#16233a;border-color:#16233a40"></span> Reading…</div>';
  try {
    const buf = await f.arrayBuffer();
    const wb = XLSX.read(buf, { cellDates: true });
    const { rows, report } = mapWorkbook(wb, mapping, XLSX);
    upState.parsed = { fileName: f.name, rows };

    if (!rows.length) {
      prev.innerHTML = banner("warn",
        `No rows found. The template expects sheet(s): <b>${esc(Object.keys(mapping.sheets || {}).join(", "))}</b>. This file has: <b>${esc(report.sheetsSeen.join(", "))}</b>.`)
        + `<div class="note">If the sheet names differ, tell me and I’ll adjust the template.</div>`;
      return;
    }
    const per = Object.entries(report.perSheet).map(([s, n]) => `${esc(s)}: <b>${n}</b>`).join(" &nbsp;·&nbsp; ");
    const oos = rows.filter((r) => r.factory_status === "out_of_stock").length;
    prev.innerHTML = `
      <div class="stat-row">
        <div class="s"><b>${rows.length}</b>rows ready</div>
        <div class="s"><b>${oos}</b>out of stock</div>
      </div>
      <div class="note">From — ${per}</div>
      <button class="btn full" id="upImport" style="margin-top:14px">Import ${rows.length} rows</button>
      <div id="upResult"></div>`;
    $("upImport").addEventListener("click", doImport);
  } catch (e) {
    prev.innerHTML = banner("err", "Could not read this file: " + esc(e.message));
  }
}

async function doImport() {
  const btn = $("upImport"); const res = $("upResult");
  btn.disabled = true; btn.innerHTML = '<span class="spin"></span> Importing…';
  const { data, error } = await sb.rpc("app_import_rows", {
    p_file_name: upState.parsed.fileName, p_rows: upState.parsed.rows,
  });
  btn.disabled = false; btn.textContent = "Import again";
  if (error) { res.innerHTML = banner("err", error.message); return; }
  res.innerHTML = banner("ok",
    `Done — <b>${data.created}</b> new, <b>${data.updated}</b> updated, <b>${data.unchanged}</b> unchanged of <b>${data.total}</b> rows.`);
  loadUploadLog();
}

async function loadUploadLog() {
  const { data } = await sb.from("upload_logs").select("file_name,status,total_rows,created_count,updated_count,unchanged_count,started_at").order("started_at", { ascending: false }).limit(10);
  $("upLog").innerHTML = data?.length ? table(
    ["File", "When", "Rows", "New", "Updated", "Same", "Status"],
    data.map((r) => [
      `<td>${esc(r.file_name)}</td>`,
      `<td class="code">${esc(new Date(r.started_at).toLocaleString())}</td>`,
      `<td class="n">${fmt(r.total_rows)}</td>`,
      `<td class="n">${fmt(r.created_count)}</td>`,
      `<td class="n">${fmt(r.updated_count)}</td>`,
      `<td class="n">${fmt(r.unchanged_count)}</td>`,
      `<td><span class="chip ${r.status === "completed" ? "avail" : "neutral"}">${esc(r.status)}</span></td>`,
    ])
  ) : emptyState("No uploads yet", "Your first import will appear here.");
}

// ---- shared render helpers -------------------------------------------------
function table(headers, rows, opts = {}) {
  const { sticky = false, classes = [] } = opts;
  const head = headers.map((h) => `<th>${esc(h)}</th>`).join("");
  const body = rows.map((cells, i) => `<tr class="${classes[i] || ""}">${cells.join("")}</tr>`).join("");
  return `<div style="${sticky ? "max-height:340px;overflow:auto" : "overflow:auto"}"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
}
function emptyState(title, sub) { return `<div class="empty"><b>${esc(title)}</b>${esc(sub)}</div>`; }
function banner(kind, html) { return `<div class="banner ${kind}">${html}</div>`; }
function trim(s, n) { s = s || ""; return s.length > n ? s.slice(0, n - 1) + "…" : s; }
function debounce(fn, ms) { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; }
