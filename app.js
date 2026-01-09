const SHEET_NAME = "平均值樞紐";

const COMPANY_WHITELIST = [
  "台氯","台達","台聚","亞聚","華夏","華運","華聚","越峯","順昶","管顧","寰靖"
];

const DEFAULT_COMPANY = "台氯";
const GROUP_ALL_VALUE = "__GROUP_ALL__";
const GROUP_ALL_LABEL = "全體";
const MAX_COMPARE = 6;

const ENGAGEMENT_SCORE = {
  "寰靖": 3.40,
  "台氯": 4.52,
  "台達": 4.50,
  "台聚": 4.10,
  "華夏": 4.50,
  "華運": 4.58,
  "公司全體": 4.36,
  "亞聚": 4.28,
  "華聚": 4.35,
  "越峯": 4.67,
  "順昶": 3.89,
  "管顧": 4.36,
};

const RESPONSE_RATE = {
  "寰靖": 92,
  "台氯": 68,
  "台達": 53,
  "台聚": 46,
  "華夏": 54,
  "華運": 59,
  "公司全體": 58,
  "亞聚": 48,
  "華聚": 95,
  "越峯": 72,
  "順昶": 62,
  "管顧": 73,
};

const PERCENT_METRIC_IDS = new Set(["總填答率"]);

const METRICS = [
  { id: "主管", excel: "平均值 - 主管" },
  { id: "薪酬", excel: "平均值 - 薪酬" },
  { id: "同事", excel: "平均值 - 同事" },
  { id: "工作", excel: "平均值 - 工作" },
  { id: "發展", excel: "平均值 - 發展" },
  { id: "企業文化", excel: "平均值 - 企業文化" },
  { id: "永續經營", excel: "平均值 - 永續經營" },
  { id: "組織承諾", excel: "平均值 - 組織承諾" },
  { id: "整體滿意度", excel: "平均值 - 整體滿意度" },
  { id: "員工敬業度", manual: true },
  { id: "總填答率", manual: true },
];

const CHART_METRIC_IDS = [
  "主管","薪酬","同事","工作","發展","企業文化","永續經營","組織承諾","整體滿意度","員工敬業度"
];

// ✅ AI 解讀預設主指標（你可改成「整體滿意度」或做下拉選單）
const AI_PRIMARY_METRIC = "主管";

// DOM
const grid = document.getElementById("grid");
const fileInput = document.getElementById("fileInput");
const loadBtn = document.getElementById("loadBtn");
const companyList = document.getElementById("companyList");
const selectedCount = document.getElementById("selectedCount");
const deptSelect = document.getElementById("deptSelect");
const roleSelect = document.getElementById("roleSelect");
const errorBox = document.getElementById("errorBox");
const noteBox = document.getElementById("noteBox");
const badgeCompany = document.getElementById("badgeCompany");

const kpiResponseRate = document.getElementById("kpiResponseRate");
const kpiOverallSat = document.getElementById("kpiOverallSat");
const kpiEngagement = document.getElementById("kpiEngagement");

const singlePanel = document.getElementById("singlePanel");
const barCanvas = document.getElementById("barChart");
const radarCanvas = document.getElementById("radarChart");

const aiNote = document.getElementById("aiNote");
const aiText = document.getElementById("aiText");

const exportCsvBtn = document.getElementById("exportCsvBtn");
const exportPdfBtn = document.getElementById("exportPdfBtn");

// state
let parsed = null;
let selectedCompanies = [];
let barChart = null;
let radarChart = null;

// utils
function normalize(s){
  return String(s ?? "")
    .replace(/\u00A0/g, " ")
    .replace(/\u3000/g, " ")
    .replace(/\s+/g, "")
    .trim();
}
function toNum(v){
  const n = (typeof v === "number") ? v : Number(String(v).trim());
  return Number.isFinite(n) ? n : null;
}
function fmt(v, metricId){
  if (v === null || v === undefined || String(v).trim() === "") return "—";
  if (PERCENT_METRIC_IDS.has(metricId)){
    if (typeof v === "string" && v.trim().endsWith("%")) return v.trim();
    const n = Number(String(v).replace("%","").trim());
    return Number.isFinite(n) ? `${Math.round(n)}%` : "—";
  }
  const n = toNum(v);
  return (n != null) ? n.toFixed(2) : "—";
}
function showError(msg){
  errorBox.style.display = "block";
  errorBox.textContent = msg;
}
function clearError(){
  errorBox.style.display = "none";
  errorBox.textContent = "";
}
function clamp01(x){ return Math.max(0, Math.min(1, x)); }

// excel helpers
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
    if (m.manual){
      out[m.id] = null;
      continue;
    }
    const key = findHeaderKey(row, m.excel);
    out[m.id] = key ? row[key] : null;
  }
  return out;
}
function mapRole(label){
  const t = normalize(label);
  if (t === "工具") return "工員";
  if (t === "工員") return "工員";
  if (t === "職員") return "職員";
  return null;
}
const COMPANY_NORM_TO_CANON = new Map(
  COMPANY_WHITELIST.map(name => [normalize(name), name])
);
function getCompanyName(label){
  return COMPANY_NORM_TO_CANON.get(normalize(label)) || null;
}
function getCountId(row){
  const key = findHeaderKey(row, "計數 - ID");
  if (!key) return null;
  const v = row[key];
  const n = (typeof v === "number") ? v : Number(String(v).trim());
  return Number.isFinite(n) ? n : null;
}

/**
 * parsed.companies[公司] = { companyAll, departments }
 * parsed.grandTotal = 集團全體（最底「總計」列，count 最大）
 */
function parseAllCompanies(rows){
  const colLabel = "列標籤";
  const labelKeyGuess = rows.length
    ? (findHeaderKey(rows[0], colLabel) || Object.keys(rows[0])[0])
    : colLabel;

  const companies = {};
  let currentCompany = null;
  let currentDept = null;

  let grandTotal = null;
  let grandTotalCount = -Infinity;

  for (const r of rows){
    const labelKey = findHeaderKey(r, colLabel) || labelKeyGuess || Object.keys(r)[0];
    const label = String(r[labelKey] ?? "");
    const normLabel = normalize(label);
    if (!normLabel) continue;

    const companyName = getCompanyName(label);
    if (companyName){
      currentCompany = companyName;
      currentDept = null;
      if (!companies[currentCompany]) companies[currentCompany] = { companyAll: null, departments: {} };
      companies[currentCompany].companyAll = extractMetrics(r);
      continue;
    }

    if (normLabel === "總計"){
      const cnt = getCountId(r);
      if (cnt != null && cnt > grandTotalCount){
        grandTotalCount = cnt;
        grandTotal = extractMetrics(r);
      }
      currentDept = null;
      continue;
    }

    if (!currentCompany) continue;

    const role = mapRole(label);
    if (role){
      if (!currentDept) continue;
      companies[currentCompany].departments[currentDept][role] = extractMetrics(r);
      continue;
    }

    currentDept = String(label).replace(/\u00A0/g," ").replace(/\u3000/g," ").trim();
    const depts = companies[currentCompany].departments;
    if (!depts[currentDept]) depts[currentDept] = {};
    depts[currentDept].ALL = extractMetrics(r);
  }

  return { companies, grandTotal };
}

// manual metrics
function applyCompanyManualMetrics(values, companyValue, dept){
  if (companyValue === GROUP_ALL_VALUE){
    return {
      ...values,
      "員工敬業度": ENGAGEMENT_SCORE["公司全體"] ?? null,
      "總填答率": RESPONSE_RATE["公司全體"] ?? null,
    };
  }
  if (dept === "__ALL__"){
    return {
      ...values,
      "員工敬業度": ENGAGEMENT_SCORE[companyValue] ?? null,
      "總填答率": RESPONSE_RATE[companyValue] ?? null,
    };
  }
  return { ...values, "員工敬業度": null, "總填答率": null };
}

// checklist UI
function updateSelectedCount(){
  selectedCount.textContent = `${selectedCompanies.length}/${MAX_COMPARE}`;
}
function makeCompanyCheckbox(value, label){
  const row = document.createElement("label");
  row.className = "navItem";

  const cb = document.createElement("input");
  cb.type = "checkbox";
  cb.value = value;

  cb.addEventListener("change", () => {
    clearError();

    const wantChecked = cb.checked;
    const already = selectedCompanies.includes(value);

    if (wantChecked && !already){
      if (selectedCompanies.length >= MAX_COMPARE){
        cb.checked = false;
        showError(`最多只能勾選 ${MAX_COMPARE} 家（含「${GROUP_ALL_LABEL}」）。`);
        return;
      }
      selectedCompanies.push(value);
    }

    if (!wantChecked && already){
      selectedCompanies = selectedCompanies.filter(x => x !== value);
      if (selectedCompanies.length === 0){
        selectedCompanies = [value];
        cb.checked = true;
      }
    }

    syncChecklistCheckedState();
    updateSelectedCount();
    syncDeptRoleAvailability();
    refreshView();
  });

  const name = document.createElement("div");
  name.className = "navName";
  name.textContent = label;

  row.appendChild(cb);
  row.appendChild(name);
  return row;
}
function syncChecklistCheckedState(){
  const items = companyList.querySelectorAll("label.navItem");
  items.forEach(item => {
    const cb = item.querySelector("input[type=checkbox]");
    const checked = selectedCompanies.includes(cb.value);
    cb.checked = checked;
    item.classList.toggle("isChecked", checked);
  });
}
function buildCompanyChecklist(){
  companyList.innerHTML = "";
  companyList.appendChild(makeCompanyCheckbox(GROUP_ALL_VALUE, GROUP_ALL_LABEL));

  const available = COMPANY_WHITELIST.filter(name => parsed.companies[name]);
  for (const name of available){
    companyList.appendChild(makeCompanyCheckbox(name, name));
  }

  if (!selectedCompanies.length){
    const defaultPick = available.includes(DEFAULT_COMPANY) ? DEFAULT_COMPANY : (available[0] || GROUP_ALL_VALUE);
    selectedCompanies = [defaultPick];
  }

  syncChecklistCheckedState();
  updateSelectedCount();
}

// dept/role
function buildDeptOptions(company){
  deptSelect.innerHTML = "";

  const optAll = document.createElement("option");
  optAll.value = "__ALL__";
  optAll.textContent = "全體員工（公司總覽）";
  deptSelect.appendChild(optAll);

  const depts = parsed.companies[company]?.departments || {};
  for (const d of Object.keys(depts)){
    const opt = document.createElement("option");
    opt.value = d;
    opt.textContent = d;
    deptSelect.appendChild(opt);
  }

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
function syncDeptRoleAvailability(){
  const single = (selectedCompanies.length === 1);
  const only = selectedCompanies[0];

  if (!single || only === GROUP_ALL_VALUE){
    deptSelect.disabled = true;
    roleSelect.disabled = true;
    deptSelect.innerHTML = `<option value="__ALL__">（比較模式：固定公司總覽）</option>`;
    deptSelect.value = "__ALL__";
    roleSelect.value = "ALL";
    return;
  }

  deptSelect.disabled = false;
  roleSelect.disabled = false;
  buildDeptOptions(only);
  updateRoleOptions(only, deptSelect.value);
}

// getters
function getCompanyOverallValues(companyValue){
  if (companyValue === GROUP_ALL_VALUE){
    const base = parsed.grandTotal || {};
    return applyCompanyManualMetrics(base, GROUP_ALL_VALUE, "__ALL__");
  }
  const c = parsed.companies[companyValue];
  const base = c?.companyAll || {};
  return applyCompanyManualMetrics(base, companyValue, "__ALL__");
}
function getSingleModeValues(){
  const companyValue = selectedCompanies[0];

  if (companyValue === GROUP_ALL_VALUE){
    const base = parsed.grandTotal || {};
    return applyCompanyManualMetrics(base, GROUP_ALL_VALUE, "__ALL__");
  }

  const dept = deptSelect.value;
  const role = roleSelect.value;

  const c = parsed.companies[companyValue];
  if (!c) return {};

  let values = {};

  if (dept === "__ALL__"){
    values = c.companyAll || {};
    return applyCompanyManualMetrics(values, companyValue, "__ALL__");
  }

  const d = c.departments[dept] || {};
  if (role === "ALL") values = d.ALL || {};
  else if (!d[role]) { roleSelect.value = "ALL"; values = d.ALL || {}; }
  else values = d[role] || {};

  return applyCompanyManualMetrics(values, companyValue, dept);
}

// KPI
function setKpisFromValue(values){
  kpiResponseRate.textContent = fmt(values?.["總填答率"], "總填答率");
  kpiOverallSat.textContent = fmt(values?.["整體滿意度"], "整體滿意度");
  kpiEngagement.textContent = fmt(values?.["員工敬業度"], "員工敬業度");
}
function setKpisFromAverages(valuesList){
  const avg = (metricId) => {
    const nums = valuesList.map(v => toNum(v?.[metricId])).filter(x => x != null);
    if (!nums.length) return null;
    return nums.reduce((a,b)=>a+b,0) / nums.length;
  };

  const avgRate = (() => {
    const nums = valuesList.map(v => {
      const raw = v?.["總填答率"];
      if (raw == null) return null;
      if (typeof raw === "string") return toNum(raw.replace("%",""));
      return toNum(raw);
    }).filter(x => x != null);
    if (!nums.length) return null;
    return nums.reduce((a,b)=>a+b,0) / nums.length;
  })();

  kpiResponseRate.textContent = (avgRate == null) ? "—" : `${Math.round(avgRate)}%`;
  const sat = avg("整體滿意度");
  const eng = avg("員工敬業度");
  kpiOverallSat.textContent = (sat == null) ? "—" : sat.toFixed(2);
  kpiEngagement.textContent = (eng == null) ? "—" : eng.toFixed(2);
}

// single cards
function renderSingleCards(values){
  grid.innerHTML = "";
  for (const m of METRICS){
    const div = document.createElement("div");
    div.className = "mCard";
    const sub = PERCENT_METRIC_IDS.has(m.id) ? "比率" : "平均值";
    div.innerHTML = `
      <div class="mK">${m.id}</div>
      <div class="mV">${fmt(values?.[m.id], m.id)}</div>
      <div class="mSub">${sub}</div>
    `;
    grid.appendChild(div);
  }
}

// charts
function destroyCharts(){
  if (barChart){ barChart.destroy(); barChart = null; }
  if (radarChart){ radarChart.destroy(); radarChart = null; }
}
function makePalette(n){
  const base = [
    "rgba(47,95,191,0.85)",
    "rgba(79,70,229,0.80)",
    "rgba(11,42,87,0.80)",
    "rgba(59,130,246,0.75)",
    "rgba(99,102,241,0.70)",
    "rgba(147,197,253,0.70)"
  ];
  const out = [];
  for (let i=0;i<n;i++) out.push(base[i % base.length]);
  return out;
}
function buildCharts(compareLabels, compareValuesMap){
  const axes = CHART_METRIC_IDS;
  const palette = makePalette(compareLabels.length);

  const datasets = compareLabels.map((label, idx) => {
    const v = compareValuesMap[label] || {};
    const data = axes.map(mid => {
      const n = toNum(v[mid]);
      return (n == null) ? null : Number(n.toFixed(2));
    });
    return { label, data, color: palette[idx] };
  });

  destroyCharts();

  barChart = new Chart(barCanvas, {
    type: "bar",
    data: {
      labels: axes,
      datasets: datasets.map(ds => ({
        label: ds.label,
        data: ds.data,
        backgroundColor: ds.color,
        borderRadius: 10,
        barPercentage: 0.78,
        categoryPercentage: 0.70
      })),
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      resizeDelay: 60,
      animation: false,
      plugins: {
        legend: { labels: { color: "rgba(15,23,42,.72)", font: { weight: "700" } } },
        tooltip: {
          callbacks: { label: (ctx) => `${ctx.dataset.label}：${(ctx.raw == null ? "—" : ctx.raw)}` }
        }
      },
      scales: {
        y: {
          suggestedMin: 0,
          suggestedMax: 5,
          grid: { color: "rgba(15,23,42,.06)" },
          ticks: { color: "rgba(15,23,42,.60)" }
        },
        x: {
          grid: { display: false },
          ticks: { color: "rgba(15,23,42,.62)", maxRotation: 0, minRotation: 0 }
        }
      }
    }
  });

  radarChart = new Chart(radarCanvas, {
    type: "radar",
    data: {
      labels: axes,
      datasets: datasets.map(ds => ({
        label: ds.label,
        data: ds.data,
        borderColor: ds.color,
        backgroundColor: ds.color.replace(/0\.\d+\)/, "0.12)"),
        borderWidth: 2.5,
        pointRadius: 2
      })),
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      resizeDelay: 60,
      animation: false,
      plugins: { legend: { labels: { color: "rgba(15,23,42,.72)", font: { weight: "700" } } } },
      scales: {
        r: {
          suggestedMin: 0,
          suggestedMax: 5,
          grid: { color: "rgba(15,23,42,.07)" },
          angleLines: { color: "rgba(15,23,42,.07)" },
          pointLabels: { color: "rgba(15,23,42,.65)", font: { weight: "700" } },
          ticks: { color: "rgba(15,23,42,.45)" }
        }
      }
    }
  });
}


function buildAIComment(compareLabels, compareValuesMap, primaryMetricId){
  const metric = primaryMetricId || AI_PRIMARY_METRIC;

  // 收集 (label, score)
  const items = compareLabels
    .map(label => ({ label, score: toNum(compareValuesMap[label]?.[metric]) }))
    .filter(x => x.score != null);

  if (items.length < 2){
    return `目前可比較的公司數不足，無法產出跨公司對比評語。`;
  }

  // best / worst
  items.sort((a,b)=>b.score - a.score);
  const best = items[0];
  const worst = items[items.length - 1];
  const gap = best.score - worst.score;

  
  let suggestion = "";
  if (gap >= 0.60) suggestion = "差距明顯，建議針對高分單位進行跨單位標竿學習，並針對低分構面制定改善方案。";
  else if (gap >= 0.30) suggestion = "存在一定差距，建議優先聚焦低分單位的管理作法與流程，導入可複製的改善措施。";
  else suggestion = "差距不大，建議以「共通弱項構面」做集團層級優化，提升整體一致性。";

  
  const satItems = compareLabels
    .map(label => ({ label, score: toNum(compareValuesMap[label]?.["整體滿意度"]) }))
    .filter(x => x.score != null)
    .sort((a,b)=>b.score - a.score);

  let satHint = "";
  if (satItems.length >= 2){
    const satBest = satItems[0];
    const satWorst = satItems[satItems.length - 1];
    const satGap = satBest.score - satWorst.score;
    satHint = `同時，「整體滿意度」以 ${satBest.label}(${satBest.score.toFixed(2)}) 較佳，${satWorst.label}(${satWorst.score.toFixed(2)}) 較低，差距 ${satGap.toFixed(2)} 分。`;
  }

  
  const engNums = compareLabels.map(l => toNum(compareValuesMap[l]?.["員工敬業度"])).filter(x=>x!=null);
  const rrNums = compareLabels.map(l => {
    const raw = compareValuesMap[l]?.["總填答率"];
    if (raw == null) return null;
    if (typeof raw === "string") return toNum(raw.replace("%",""));
    return toNum(raw);
  }).filter(x=>x!=null);

  let qualityHint = "";
  if (rrNums.length){
    const rrAvg = rrNums.reduce((a,b)=>a+b,0)/rrNums.length;
    if (rrAvg < 55) qualityHint += `（提醒：平均填答覆蓋率約 ${Math.round(rrAvg)}%，若偏低可能影響代表性）`;
  }
  if (engNums.length){
    const engAvg = engNums.reduce((a,b)=>a+b,0)/engNums.length;
    if (engAvg < 4.0) qualityHint += `（敬業度平均約 ${engAvg.toFixed(2)}，建議同步關注留才與投入度）`;
  }

  
  return `2025年 ${metric}：${best.label}（${best.score.toFixed(2)}）表現最佳；${worst.label}（${worst.score.toFixed(2)}）暫居末位；差距 ${gap.toFixed(2)} 分，${suggestion}${satHint}${qualityHint}`;
}

function showAI(text){
  aiNote.style.display = "flex";
  aiText.textContent = text;
}
function hideAI(){
  aiNote.style.display = "none";
  aiText.textContent = "—";
}

// refresh
function refreshView(){
  const displayNames = selectedCompanies.map(v => v === GROUP_ALL_VALUE ? GROUP_ALL_LABEL : v);
  badgeCompany.textContent = `公司：${displayNames.join("、")}`;

  // 多公司比較
  if (selectedCompanies.length > 1){
    singlePanel.style.display = "none";

    const map = {};
    const valuesList = [];
    for (const v of selectedCompanies){
      const label = (v === GROUP_ALL_VALUE) ? GROUP_ALL_LABEL : v;
      const val = getCompanyOverallValues(v);
      map[label] = val;
      valuesList.push(val);
    }

    setKpisFromAverages(valuesList);
    buildCharts(displayNames, map);

   
    const comment = buildAIComment(displayNames, map, AI_PRIMARY_METRIC);
    showAI(comment);

    noteBox.textContent = `比較模式：已選 ${selectedCompanies.length} 家｜顯示各公司「全體部門（公司總覽）」`;
    return;
  }


  singlePanel.style.display = "block";
  hideAI(); 

  const only = selectedCompanies[0];
  const values = getSingleModeValues();
  setKpisFromValue(values);

  const label = (only === GROUP_ALL_VALUE) ? GROUP_ALL_LABEL : only;
  buildCharts([label], { [label]: values });

  renderSingleCards(values);

  if (only === GROUP_ALL_VALUE){
    noteBox.textContent = `顯示：${GROUP_ALL_LABEL}｜總計（樞紐最底列）`;
  } else {
    const dept = deptSelect.value;
    const role = roleSelect.value;
    if (dept === "__ALL__"){
      noteBox.textContent = `顯示：${only}｜全體員工（公司總覽）`;
    } else {
      const d = parsed.companies[only]?.departments?.[dept] || {};
      const flags = `工員：${d["工員"] ? "有" : "無"}；職員：${d["職員"] ? "有" : "無"}`;
      noteBox.textContent = `顯示：${only}｜${dept}｜${role === "ALL" ? "全部（部門總覽）" : role}（${flags}）`;
    }
  }
}

// load excel
async function loadExcel(file){
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type:"array" });

  const ws = wb.Sheets[SHEET_NAME];
  if (!ws) throw new Error(`找不到分頁：${SHEET_NAME}`);

  const rows = XLSX.utils.sheet_to_json(ws, { defval:"" });
  const p = parseAllCompanies(rows);

  const any = COMPANY_WHITELIST.some(name => p.companies[name]);
  if (!any) throw new Error("沒有解析到任何白名單公司。");

  return p;
}

// events
deptSelect.addEventListener("change", () => {
  const only = selectedCompanies[0];
  if (selectedCompanies.length === 1 && only !== GROUP_ALL_VALUE){
    updateRoleOptions(only, deptSelect.value);
  }
  refreshView();
});
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

    selectedCompanies = [];
    buildCompanyChecklist();

    const available = COMPANY_WHITELIST.filter(name => parsed.companies[name]);
    if (!selectedCompanies.length){
      selectedCompanies = [available.includes(DEFAULT_COMPANY) ? DEFAULT_COMPANY : (available[0] || GROUP_ALL_VALUE)];
    }

    syncChecklistCheckedState();
    updateSelectedCount();
    syncDeptRoleAvailability();

    refreshView();
  }catch(e){
    showError(
      `讀取失敗：${e.message}\n\n` +
      `檢查：\n` +
      `1) 分頁名稱「平均值樞紐」\n` +
      `2) 欄名是否包含：\n` +
      METRICS.filter(m => !m.manual).map(m => `- ${m.excel}`).join("\n")
    );

    if (barChart){ barChart.destroy(); barChart = null; }
    if (radarChart){ radarChart.destroy(); radarChart = null; }
    grid.innerHTML = "";
    noteBox.textContent = "讀取失敗，請確認 Excel 結構/欄名。";
    deptSelect.disabled = true;
    roleSelect.disabled = true;
    companyList.innerHTML = `<div class="placeholder">讀取失敗</div>`;
    selectedCompanies = [];
    updateSelectedCount();

    kpiResponseRate.textContent = "—";
    kpiOverallSat.textContent = "—";
    kpiEngagement.textContent = "—";
    hideAI();
  }
});

// 目前先保留按鈕（要真的匯出我再幫你接）
exportCsvBtn.addEventListener("click", () => alert("CSV 匯出：尚未接出檔（需要你確認匯出欄位）"));
exportPdfBtn.addEventListener("click", () => alert("PDF 匯出：尚未接出檔（可用 jsPDF 產報告）"));
