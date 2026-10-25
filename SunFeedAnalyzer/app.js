"use strict";

const TYPES = ["Production", "Sales", "Sales Return", "Transfer Out", "Re-Process", "Stock Correction"];
// Stock hisab er sign. Finished Goods Report er sathe milie nin. Dorkar hole 1 / -1 / 0 bodlan.
const SIGN = {
  "Production": 1, "Sales Return": 1, "Stock Correction": 1,
  "Sales": -1, "Transfer Out": -1, "Re-Process": -1
};

const CANON = {
  "date": "Date", "feed name": "Feed Name", "quantity": "Quantity", "package": "Package",
  "quantity (mt)": "MT", "type": "Type", "feed category": "Category",
  "sub-category": "SubCategory", "invoice no.": "Invoice", "customer id": "CustomerID",
  "customer name": "Customer", "rate": "Rate", "debit amount": "Amount",
  "discount": "Discount", "officer name": "Officer", "area": "Area",
  "vehicle no.": "Vehicle", "location": "Location", "remarks": "Remarks"
};
const NUM_FIELDS = ["Quantity", "Package", "MT", "Rate", "Amount", "Discount"];
const TEXT_FIELDS = ["Feed Name", "Type", "Category", "SubCategory", "Invoice", "CustomerID",
  "Customer", "Officer", "Area", "Vehicle", "Location", "Remarks"];

let ALL = [];
let SHEETS = {};
let DATA_SHEET = "";
const charts = {};

const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"']/g,
  c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const nf = (v, d = 3) => v == null || isNaN(v) ? "—"
  : Number(v).toLocaleString(undefined, { maximumFractionDigits: d });

/* ---------- Parsing ---------- */

function normType(t) {
  if (!t) return "Unknown";
  const s = String(t).toLowerCase();
  if (s.includes("return")) return "Sales Return";
  if (s.includes("transfer")) return "Transfer Out";
  if (s.includes("process")) return "Re-Process";
  if (s.includes("correct")) return "Stock Correction";
  if (s.includes("sale")) return "Sales";
  if (s.includes("produc")) return "Production";
  return String(t).trim();
}

function toNum(v) {
  if (v === "" || v == null) return null;
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(/,/g, ""));
  return isNaN(n) ? null : n;
}

function toDate(v) {
  if (typeof v === "number") {
    if (v < 20000 || v > 80000) return null;
    return new Date(Math.round((v - 25569) * 86400000)).toISOString().slice(0, 10);
  }
  if (typeof v === "string") {
    const m = v.trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return m[0];
  }
  return null;
}

function parseWorkbook(buf) {
  const wb = XLSX.read(buf, { type: "array" });
  SHEETS = {};
  let best = null;

  wb.SheetNames.forEach(name => {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: "" });
    SHEETS[name] = rows;
    for (let i = 0; i < Math.min(40, rows.length); i++) {
      const set = new Set(rows[i].map(v => String(v).trim().toLowerCase()));
      if (set.has("date") && set.has("feed name") && set.has("type")) {
        if (!best || rows.length > best.rows.length) best = { name, rows, h: i };
        break;
      }
    }
  });

  if (!best) throw new Error("Kono sheet e 'Date', 'Feed Name', 'Type' header paoa jayni");

  DATA_SHEET = best.name;
  const idx = {};
  best.rows[best.h].forEach((h, i) => {
    const key = CANON[String(h).trim().toLowerCase()];
    if (key && !(key in idx)) idx[key] = i;
  });

  const out = [];
  for (let i = best.h + 1; i < best.rows.length; i++) {
    const r = best.rows[i];
    const rec = {};
    for (const k of Object.keys(CANON).map(x => CANON[x])) {
      const v = k in idx ? r[idx[k]] : "";
      if (k === "Date") rec.Date = toDate(v);
      else if (NUM_FIELDS.includes(k)) rec[k] = toNum(v);
      else rec[k] = String(v ?? "").trim();
    }
    if (!rec.Date || !rec["Feed Name"]) continue;
    if (rec.MT == null) rec.MT = (rec.Quantity || 0) * (rec.Package || 0) / 1000;
    rec.TType = normType(rec.Type);
    out.push(rec);
  }
  return out;
}

/* ---------- Filtering ---------- */

function filtered() {
  const start = $("start").value, end = $("end").value;
  const cat = $("category").value, loc = $("location").value, off = $("officer").value;
  const noopen = $("noopen").checked;
  return ALL.filter(r =>
    (!start || r.Date >= start) &&
    (!end || r.Date <= end) &&
    (!cat || r.Category === cat) &&
    (!loc || r.Location === loc) &&
    (!off || r.Officer === off) &&
    (!noopen || !/opening stock/i.test(r.Remarks))
  );
}

function periodKey(date, freq) {
  if (freq === "D") return date;
  if (freq === "M") return date.slice(0, 7);
  const d = new Date(date + "T00:00:00Z");
  d.setUTCDate(d.getUTCDate() - ((d.getUTCDay() + 6) % 7));
  return d.toISOString().slice(0, 10);
}

/* ---------- Helpers ---------- */

function pivot(rows, keyFn) {
  const m = {};
  rows.forEach(r => {
    const k = keyFn(r) || "(blank)";
    const o = m[k] ??= Object.fromEntries(TYPES.map(t => [t, 0]));
    o[r.TType] = (o[r.TType] || 0) + r.MT;
  });
  return m;
}

function salesBy(rows, field, limit) {
  const m = {};
  rows.filter(r => r.TType === "Sales").forEach(r => {
    const k = r[field] || "(blank)";
    const o = m[k] ??= { name: k, Bags: 0, MT: 0, Amount: 0, inv: new Set(), cust: new Set(), Last: "" };
    o.Bags += r.Quantity || 0;
    o.MT += r.MT;
    o.Amount += r.Amount || 0;
    if (r.Invoice) o.inv.add(r.Invoice);
    if (r.Customer) o.cust.add(r.Customer);
    if (r.Date > o.Last) o.Last = r.Date;
  });
  const arr = Object.values(m).sort((a, b) => b.MT - a.MT).map(o => ({
    ...o, Invoices: o.inv.size, Customers: o.cust.size
  }));
  return limit ? arr.slice(0, limit) : arr;
}

function tbl(el, cols, rows) {
  $(el).innerHTML = "<tr>" + cols.map(c => `<th>${esc(c[1])}</th>`).join("") + "</tr>" +
    rows.map(r => "<tr>" + cols.map(c => {
      const v = r[c[0]];
      if (c[2]) return `<td class="num ${v < 0 ? "neg" : ""}">${nf(v)}</td>`;
      return `<td>${esc(v)}</td>`;
    }).join("") + "</tr>").join("");
}

function mk(id, cfg) {
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart($(id), cfg);
}

function fill(id, arr) {
  $(id).innerHTML = `<option value="">All</option>` +
    arr.map(x => `<option>${esc(x)}</option>`).join("");
}

const uniq = f => [...new Set(ALL.map(r => r[f]).filter(Boolean))].sort();

/* ---------- Render ---------- */

function render() {
  const rows = filtered();
  if (!rows.length) {
    $("kpis").innerHTML = "<div class='card'>Ei filter e kono data nai.</div>";
    return;
  }

  // KPI
  const sum = t => rows.filter(r => r.TType === t).reduce((a, r) => a + r.MT, 0);
  const salesAmt = rows.filter(r => r.TType === "Sales").reduce((a, r) => a + (r.Amount || 0), 0);

  // Product table
  const prodMap = {};
  rows.forEach(r => {
    const p = prodMap[r["Feed Name"]] ??= {
      "Feed Name": r["Feed Name"], Category: r.Category,
      ...Object.fromEntries(TYPES.map(t => [t, 0])), Stock: 0
    };
    p[r.TType] = (p[r.TType] || 0) + r.MT;
    p.Stock += (SIGN[r.TType] || 0) * r.MT;
  });
  const prod = Object.values(prodMap).sort((a, b) => b.Sales - a.Sales);
  const totalStock = prod.reduce((a, p) => a + p.Stock, 0);

  const cards = [
    ["Production (MT)", sum("Production")], ["Sales (MT)", sum("Sales")],
    ["Sales Return (MT)", sum("Sales Return")], ["Net Sales (MT)", sum("Sales") - sum("Sales Return")],
    ["Transfer Out (MT)", sum("Transfer Out")], ["Sales Amount", salesAmt],
    ["Estimated Stock (MT)", totalStock], ["Rows", rows.length]
  ];
  $("kpis").innerHTML = cards.map(c =>
    `<div class="kpi"><b>${nf(c[1], 2)}</b><span>${c[0]}</span></div>`).join("");

  // Trend chart
  const freq = $("freq").value;
  const trend = pivot(rows, r => periodKey(r.Date, freq));
  const periods = Object.keys(trend).sort();
  const colors = { "Production": "#16a34a", "Sales": "#2563eb", "Sales Return": "#dc2626", "Transfer Out": "#f59e0b" };
  mk("c_trend", {
    type: "bar",
    data: {
      labels: periods,
      datasets: Object.keys(colors).map(t => ({
        label: t, data: periods.map(p => +trend[p][t].toFixed(3)), backgroundColor: colors[t]
      }))
    },
    options: { responsive: true, maintainAspectRatio: false }
  });

  // Category chart
  const cat = pivot(rows, r => r.Category);
  const catNames = Object.keys(cat);
  mk("c_cat", {
    type: "doughnut",
    data: {
      labels: catNames,
      datasets: [{
        data: catNames.map(c => +cat[c].Sales.toFixed(3)),
        backgroundColor: ["#2563eb", "#16a34a", "#f59e0b", "#dc2626", "#7c3aed", "#0891b2", "#6b7280"]
      }]
    },
    options: { responsive: true, maintainAspectRatio: false }
  });

  // Top 10 charts
  const topChart = (id, labels, data, color) => mk(id, {
    type: "bar",
    data: { labels, datasets: [{ label: "MT", data, backgroundColor: color }] },
    options: { indexAxis: "y", responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
  });
  const top10p = prod.slice(0, 10);
  topChart("c_prod", top10p.map(p => p["Feed Name"]), top10p.map(p => +p.Sales.toFixed(3)), "#2563eb");
  const cust = salesBy(rows, "Customer", 20);
  topChart("c_cust", cust.slice(0, 10).map(c => c.name), cust.slice(0, 10).map(c => +c.MT.toFixed(3)), "#16a34a");

  // Tables
  tbl("t_prod", [["Feed Name", "Product"], ["Category", "Category"], ["Production", "Production", 1],
    ["Sales", "Sales", 1], ["Sales Return", "Return", 1], ["Transfer Out", "Transfer", 1],
    ["Re-Process", "Re-Process", 1], ["Stock Correction", "Correction", 1], ["Stock", "Est. Stock", 1]], prod);
  tbl("t_cust", [["name", "Customer"], ["Bags", "Bags", 1], ["MT", "MT", 1], ["Amount", "Amount", 1],
    ["Invoices", "Invoices", 1], ["Last", "Last date"]], cust);
  tbl("t_off", [["name", "Officer"], ["MT", "MT", 1], ["Amount", "Amount", 1], ["Customers", "Customers", 1]],
    salesBy(rows, "Officer"));
  tbl("t_area", [["name", "Area"], ["MT", "MT", 1], ["Amount", "Amount", 1], ["Customers", "Customers", 1]],
    salesBy(rows, "Area", 20));

  renderQuality(rows, prod);
  renderRows();
}

function renderQuality(rows, prod) {
  const seen = new Set();
  let dup = 0;
  rows.forEach(r => {
    const k = JSON.stringify(r);
    if (seen.has(k)) dup++; else seen.add(k);
  });
  const mismatch = rows.filter(r => Math.abs((r.Quantity || 0) * (r.Package || 0) / 1000 - r.MT) > 0.001).length;
  const sales = rows.filter(r => r.TType === "Sales");
  const noCust = sales.filter(r => !r.Customer).length;
  const zeroAmt = sales.filter(r => !r.Amount).length;
  const opening = rows.filter(r => /opening stock/i.test(r.Remarks)).length;
  const unknown = [...new Set(rows.map(r => r.TType))].filter(t => !TYPES.includes(t));
  const neg = prod.filter(p => p.Stock < -0.0005);

  const li = [
    `Opening Stock remark wala row: <b>${opening}</b> (production total e dhora hoy, checkbox diye bad din)`,
    `Hubohu ek (identical) row: <b>${dup}</b> - duplicate entry hote pare, check korun`,
    `Quantity x Package theke MT er sathe na milche emon row: <b>${mismatch}</b>`,
    `Sales row jeta customer chara: <b>${noCust}</b>, amount 0/blank: <b>${zeroAmt}</b>`
  ];
  if (unknown.length) li.push(`Ochena Type: <b>${unknown.map(esc).join(", ")}</b>`);
  li.push(neg.length
    ? `Negative estimated stock: <b class="neg">${neg.map(p => esc(p["Feed Name"]) + " (" + nf(p.Stock) + ")").join(", ")}</b>`
    : "Kono product e negative stock nai.");
  $("quality").innerHTML = "<div>" + li.join("</div><div>") + "</div>";
}

function renderRows() {
  const q = $("q").value.trim().toLowerCase();
  let rows = filtered();
  if (q) {
    rows = rows.filter(r => [r["Feed Name"], r.Customer, r.Remarks, r.Invoice, r.Officer,
      r.Area, r.Location, r.Category, r.TType].join(" ").toLowerCase().includes(q));
  }
  const shown = [...rows].sort((a, b) => b.Date.localeCompare(a.Date)).slice(0, 300);
  $("rowinfo").textContent = `${rows.length} row match korche (notun 300 ta dekhano hocche)`;
  tbl("t_rows", [["Date", "Date"], ["Feed Name", "Product"], ["Quantity", "Bags", 1], ["Package", "Kg/bag", 1],
    ["MT", "MT", 1], ["TType", "Type"], ["Category", "Category"], ["Invoice", "Invoice"],
    ["Customer", "Customer"], ["Rate", "Rate", 1], ["Amount", "Amount", 1], ["Officer", "Officer"],
    ["Area", "Area"], ["Location", "Location"], ["Remarks", "Remarks"]], shown);
}

function renderSheet() {
  const rows = (SHEETS[$("sheetSel").value] || []).slice(0, 150);
  $("t_sheet").innerHTML = rows.map(r =>
    "<tr>" + r.map(c => `<td>${esc(c)}</td>`).join("") + "</tr>").join("");
}

/* ---------- Events ---------- */

$("file").addEventListener("change", async e => {
  const f = e.target.files[0];
  if (!f) return;
  $("empty").textContent = "File read hocche, ektu opekkha korun...";
  try {
    const buf = await f.arrayBuffer();
    ALL = parseWorkbook(buf);
  } catch (err) {
    $("empty").textContent = "File read kora jayni: " + err.message;
    return;
  }
  const dates = ALL.map(r => r.Date).sort();
  $("fname").textContent = f.name;
  $("empty").style.display = "none";
  $("app").style.display = "block";
  $("start").value = dates[0];
  $("end").value = dates[dates.length - 1];
  fill("category", uniq("Category"));
  fill("location", uniq("Location"));
  fill("officer", uniq("Officer"));
  $("sheetSel").innerHTML = Object.keys(SHEETS).map(n =>
    `<option value="${esc(n)}">${esc(n)} (${SHEETS[n].length} rows)</option>`).join("");
  $("info").textContent = `Data sheet: "${DATA_SHEET}" | ${ALL.length} transaction row | ${dates[0]} theke ${dates[dates.length - 1]}`;
  render();
  renderSheet();
});

$("applyBtn").addEventListener("click", render);
$("searchBtn").addEventListener("click", renderRows);
$("q").addEventListener("keydown", e => { if (e.key === "Enter") renderRows(); });
$("sheetSel").addEventListener("change", renderSheet);

window.addEventListener("load", () => {
  if (typeof XLSX === "undefined" || typeof Chart === "undefined") {
    $("empty").classList.add("warn");
    $("empty").innerHTML = "<b>lib folder e file nai.</b> <code>lib/xlsx.full.min.js</code> ar " +
      "<code>lib/chart.umd.js</code> download kore rakhun (Step 1 dekhun), tarpor page ta refresh korun.";
  }
});
