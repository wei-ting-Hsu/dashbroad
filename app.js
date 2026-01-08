const SHEET_NAME = "平均值樞紐";

/** ✅ 公司白名單（只會出現這些） */
const COMPANY_WHITELIST = [
  "台氯","台達","台聚","亞聚","華夏","華運","華聚","越峯","順昶","管顧","寰靖"
];

// 預設優先選台氯
const DEFAULT_COMPANY = "台氯";

// 「全體公司」選項的 value
const GROUP_ALL_VALUE = "__GROUP_ALL__";

const METRICS = [
  { id: "主管", excel: "平均值 - 主管" },
  { id: "薪酬", excel: "平均值 - 薪酬" },
  { id: "同事", excel: "平均值 - 同事" },
  { id: "工作", excel: "平均值 - 工作" },
  { id: "發展", excel: "平均值 - 發展" },
  { id: "企業文化", excel: "平均值 - 企業文化" },
  { id: "永續經營", excel: "平均值 - 永續經營" },
  { id: "組織承諾", excel: "平均值 - 組織承諾" },
  { id: "整體滿意度", excel: "平均值 - 整體滿意度" }
];

// DOM
const grid = document.getElementById("grid");
const fileInput = document.getElementById("fileInput");
const loadBtn = document.getElementById("loadBtn");
const companySelect = document.getElementById("companySelect");
const deptSelect = document.getElementById("deptSelect");
const roleSelect = document.getElementById("roleSelect");
const errorBox = document.getElementById("errorBox");
const noteBox = document.getElementById("noteBox");
const badgeCompany = document.getElementById("badgeCompany");

// ===== utils =====

// 吃掉：半形/全形空白、NBSP、tab、換行、樞紐縮排等
function normalize(s){
  return String(s ?? "")
    .replace(/\u00A0/g, " ")  // NBSP
    .replace(/\u3000/g, " ")  // 全形空白
    .replace(/\s+/g, "")      // 所有空白
    .trim();
}

function fmt(v){
  const n = (typeof v === "number") ? v : Number(String(v).trim());
  return Number.isFinite(n) ? n.toFixed(2) : "—";
}

function showError(msg){
  errorBox.style.display = "block";
  errorBox.textContent = msg;
}
function clearError(){
  errorBox.style.display = "none";
  errorBox.textContent = "";
}

function render(values){
  grid.innerHTML = "";
  for (const m of METRICS){
    const div = document.createElement("div");
    div.className = "card";
    div.innerHTML = `
      <div class="k">${m.id}</div>
      <div class="v">${fmt(values?.[m.id])}</div>
      <div class="sub">平均值</div>
    `;
    grid.appendChild(div);
  }
}

function findHeaderKey(rowObj, expectedName){
  const expected = normalize(expectedName);
  for (const key of Object.keys(rowObj)){
    if (normalize(key) === expected) return key;
  }
  return null;
}

function extractMetrics(row){
  const out = {};
  for (const m of METRICS){
    const key = findHeaderKey(row, m.excel);
    out[m.id] = key ? row[key] : null;
  }
  return out;
}

// 工種映射：Excel 的「工具」視為「工員」
function mapRole(label){
  const t = normalize(label);
  if (t === "工具") return "工員";
  if (t === "工員") return "工員";
  if (t === "職員") return "職員";
  return null;
}

// ===== ✅ 公司辨識：用 normalize 對白名單比對（解決樞紐縮排） =====
const COMPANY_NORM_TO_CANON = new Map(
  COMPANY_WHITELIST.map(name => [normalize(name), name])
);
function getCompanyName(label){
  return COMPANY_NORM_TO_CANON.get(normalize(label)) || null;
}

// 取得「計數 - ID」欄位的數值，用來辨識最底的總計（1241那列）
function getCountId(row){
  const key = findHeaderKey(row, "計數 - ID");
  if (!key) return null;
  const v = row[key];
  const n = (typeof v === "number") ? v : Number(String(v).trim());
  return Number.isFinite(n) ? n : null;
}

/**
 * 解析：
 * parsed.companies[公司] = { companyAll, departments }
 * departments[部門] = { ALL, 工員?, 職員? }
 * parsed.grandTotal = 全體公司（最底「總計」那列）
 */
function parseAllCompanies(rows){
  const colLabel = "列標籤";
  const labelKeyGuess = rows.length
    ? (findHeaderKey(rows[0], colLabel) || Object.keys(rows[0])[0])
    : colLabel;

  const companies = {};
  let currentCompany = null;
  let currentDept = null;

  // 全體公司總計（取所有「總計」中 count 最大的那筆，通常就是 1241）
  let grandTotal = null;
  let grandTotalCount = -Infinity;

  for (const r of rows){
    const labelKey = findHeaderKey(r, colLabel) || labelKeyGuess || Object.keys(r)[0];
    const labelRaw = r[labelKey];
    const label = String(labelRaw ?? "");
    const normLabel = normalize(label);
    if (!normLabel) continue;

    // 1) 公司行：只接受白名單（用 normalize 辨識）
    const companyName = getCompanyName(label);
    if (companyName){
      currentCompany = companyName;
      currentDept = null;
      if (!companies[currentCompany]) companies[currentCompany] = { companyAll: null, departments: {} };
      companies[currentCompany].companyAll = extractMetrics(r);
      continue;
    }

    // 2) 「總計」行：用 count 最大的那筆當全體公司（通常是最底 1241）
    if (normLabel === "總計"){
      const cnt = getCountId(r);
      if (cnt != null && cnt > grandTotalCount){
        grandTotalCount = cnt;
        grandTotal = extractMetrics(r);
      }
      // 同時結束部門（不結束公司）
      currentDept = null;
      continue;
    }

    // 3) 還沒進公司就略過（避免把表頭或其他東西當部門）
    if (!currentCompany) continue;

    // 4) 工種（工具/工員/職員）
    const role = mapRole(label);
    if (role){
      if (!currentDept) continue;
      companies[currentCompany].departments[currentDept][role] = extractMetrics(r);
      continue;
    }

    // 5) 其他一律視為部門（✅ 華運的 技術類/管理類 也會走這裡）
    currentDept = String(label)
      .replace(/\u00A0/g, " ")
      .replace(/\u3000/g, " ")
      .trim();

    const depts = companies[currentCompany].departments;
    if (!depts[currentDept]) depts[currentDept] = {};
    depts[currentDept].ALL = extractMetrics(r);
  }

  return { companies, grandTotal };
}

// ===== UI state =====
let parsed = null;

function buildCompanyOptions(){
  companySelect.innerHTML = "";

  // ✅ 第12個選項：全體公司（抓最底總計）
  const optGroup = document.createElement("option");
  optGroup.value = GROUP_ALL_VALUE;
  optGroup.textContent = "全體公司";
  companySelect.appendChild(optGroup);

  // 只顯示「有解析到資料」且在白名單內的公司，並按白名單順序排列
  const available = COMPANY_WHITELIST.filter(name => parsed.companies[name]);

  for (const name of available){
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    companySelect.appendChild(opt);
  }

  if (!available.length){
    throw new Error("白名單公司在 Excel 內都沒有找到（請確認樞紐是否含公司列）。");
  }

  companySelect.disabled = false;

  // 預設選台氯；如果要預設先看全體公司，把這行改成 GROUP_ALL_VALUE
  companySelect.value = available.includes(DEFAULT_COMPANY) ? DEFAULT_COMPANY : available[0];
}

function setDeptRoleForGroupAll(){
  deptSelect.innerHTML = `<option value="__ALL__">全體公司（不分部門）</option>`;
  deptSelect.value = "__ALL__";
  roleSelect.value = "ALL";
  deptSelect.disabled = true;
  roleSelect.disabled = true;
}

function buildDeptOptions(company){
  deptSelect.innerHTML = "";

  const optAll = document.createElement("option");
  optAll.value = "__ALL__";
  optAll.textContent = "全體員工（公司總覽）";
  deptSelect.appendChild(optAll);

  const depts = parsed.companies[company]?.departments || {};
  const deptNames = Object.keys(depts);

  for (const d of deptNames){
    const opt = document.createElement("option");
    opt.value = d;
    opt.textContent = d;
    deptSelect.appendChild(opt);
  }

  deptSelect.disabled = false;
  roleSelect.disabled = false;
  deptSelect.value = "__ALL__";
  roleSelect.value = "ALL";
}

function updateRoleOptions(company, dept){
  const optALL = Array.from(roleSelect.options).find(o => o.value === "ALL");
  const optWorker = Array.from(roleSelect.options).find(o => o.value === "工員");
  const optStaff = Array.from(roleSelect.options).find(o => o.value === "職員");

  if (dept === "__ALL__"){
    optALL.disabled = false;
    optWorker.disabled = true;
    optStaff.disabled = true;
    roleSelect.value = "ALL";
    return;
  }

  const d = parsed.companies[company]?.departments?.[dept] || {};
  const hasWorker = !!d["工員"];
  const hasStaff = !!d["職員"];

  optALL.disabled = false;
  optWorker.disabled = !hasWorker;
  optStaff.disabled = !hasStaff;

  if (roleSelect.value === "工員" && !hasWorker) roleSelect.value = "ALL";
  if (roleSelect.value === "職員" && !hasStaff) roleSelect.value = "ALL";
}

function getCurrentValues(){
  const company = companySelect.value;

  // ✅ 全體公司：抓最底總計
  if (company === GROUP_ALL_VALUE){
    return parsed.grandTotal || {};
  }

  const dept = deptSelect.value;
  const role = roleSelect.value;

  const c = parsed.companies[company];
  if (!c) return {};

  if (dept === "__ALL__"){
    return c.companyAll || {};
  }

  const d = c.departments[dept] || {};
  if (role === "ALL") return d.ALL || {};
  if (!d[role]) { roleSelect.value = "ALL"; return d.ALL || {}; }
  return d[role] || {};
}

function refreshView(){
  const company = companySelect.value;

  // ✅ 全體公司顯示
  if (company === GROUP_ALL_VALUE){
    render(parsed.grandTotal || {});
    badgeCompany.textContent = `公司：全體公司`;
    noteBox.textContent = `顯示：全體公司｜總計（樞紐最底列）`;
    return;
  }

  const dept = deptSelect.value;

  updateRoleOptions(company, dept);

  const role = roleSelect.value;
  const values = getCurrentValues();
  render(values);

  badgeCompany.textContent = `公司：${company}`;

  if (dept === "__ALL__"){
    noteBox.textContent = `顯示：${company}｜全體員工（公司總覽）`;
  } else {
    const d = parsed.companies[company]?.departments?.[dept] || {};
    const flags = `工員：${d["工員"] ? "有" : "無"}；職員：${d["職員"] ? "有" : "無"}`;
    noteBox.textContent = `顯示：${company}｜${dept}｜${role === "ALL" ? "全部（部門總覽）" : role}（${flags}）`;
  }
}

// load
async function loadExcel(file){
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type:"array" });

  const ws = wb.Sheets[SHEET_NAME];
  if (!ws) throw new Error(`找不到分頁：${SHEET_NAME}`);

  const rows = XLSX.utils.sheet_to_json(ws, { defval:"" });
  const p = parseAllCompanies(rows);

  // 至少要解析到白名單其中一家
  const any = COMPANY_WHITELIST.some(name => p.companies[name]);
  if (!any) throw new Error("沒有解析到任何白名單公司。");

  // 沒抓到全體公司總計就提醒（但不阻擋使用）
  if (!p.grandTotal){
    console.warn("⚠️ 沒有抓到全體公司『總計』列，『全體公司』會顯示空值。請確認樞紐最底是否有『總計』。");
  }

  return p;
}

// events
companySelect.addEventListener("change", () => {
  if (companySelect.value === GROUP_ALL_VALUE){
    setDeptRoleForGroupAll();
    refreshView();
    return;
  }

  buildDeptOptions(companySelect.value);
  refreshView();
});

deptSelect.addEventListener("change", refreshView);
roleSelect.addEventListener("change", refreshView);

loadBtn.addEventListener("click", async () => {
  clearError();
  const file = fileInput.files?.[0];
  if (!file){
    showError("請先選擇 Excel 檔案（.xlsx）。");
    return;
  }

  try{
    parsed = await loadExcel(file);
    buildCompanyOptions();

    // 如果你想預設先看「全體公司」，把下面兩行改成：
    // companySelect.value = GROUP_ALL_VALUE; setDeptRoleForGroupAll();
    buildDeptOptions(companySelect.value);

    refreshView();
  }catch(e){
    showError(
      `讀取失敗：${e.message}\n\n` +
      `檢查：\n` +
      `1) 分頁名稱「平均值樞紐」\n` +
      `2) 欄名是否包含：\n` +
      METRICS.map(m => `- ${m.excel}`).join("\n")
    );
    render({});
    noteBox.textContent = "讀取失敗，請確認 Excel 結構/欄名。";
    companySelect.disabled = true;
    deptSelect.disabled = true;
    roleSelect.disabled = true;
  }
});

render({});
