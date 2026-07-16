let currentSession = null;
let currentHistoryId = null;
let currentPreview = null;
let historySearchPage = 1;
let historySearchHasRun = false;

const statusEl = document.querySelector("#status");
const configForm = document.querySelector("#configForm");
const configResult = document.querySelector("#configResult");
const activeConfig = document.querySelector("#activeConfig");
const uploadForm = document.querySelector("#uploadForm");
const documentsInput = document.querySelector("#documents");
const documentsName = document.querySelector("#documentsName");
const fileCount = document.querySelector("#fileCount");
const summary = document.querySelector("#summary");
const previewPanel = document.querySelector("#previewPanel");
const previewBody = document.querySelector("#previewBody");
const previewToggle = document.querySelector("#previewToggle");
const previewMeta = document.querySelector("#previewMeta");
const previewRows = document.querySelector("#previewRows");
const warningsEl = document.querySelector("#warnings");
const previewLimit = document.querySelector("#previewLimit");
const confirmBox = document.querySelector("#confirmBox");
const warningConfirm = document.querySelector("#warningConfirm");
const generateBtn = document.querySelector("#generateBtn");
const historyList = document.querySelector("#historyList");
const refreshHistoryBtn = document.querySelector("#refreshHistoryBtn");
const historySearchForm = document.querySelector("#historySearchForm");
const historyStart = document.querySelector("#historyStart");
const historyEnd = document.querySelector("#historyEnd");
const historySearchSummary = document.querySelector("#historySearchSummary");
const historySearchList = document.querySelector("#historySearchList");
const historyPager = document.querySelector("#historyPager");

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, char => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    "\"": "&quot;",
    "'": "&#39;"
  }[char]));
}

function fmt(value, digits = 2) {
  if (value === null || value === undefined || value === "") return "-";
  if (typeof value === "number") {
    return value.toLocaleString("zh-CN", { maximumFractionDigits: digits });
  }
  return value;
}

function setField(name, value) {
  const element = document.querySelector(`[data-field="${name}"]`);
  if (element) element.textContent = fmt(value, 6);
}

function status(text, tone = "") {
  statusEl.className = `status ${tone}`.trim();
  statusEl.textContent = text;
}

async function responseJson(response) {
  const data = await response.json();
  if (!response.ok || !data.ok) {
    throw new Error(data.error || `请求失败（${response.status}）`);
  }
  return data;
}

for (const id of ["template", "rules"]) {
  const input = document.querySelector(`#${id}`);
  const label = document.querySelector(`#${id}Name`);
  input.addEventListener("change", event => {
    label.textContent = event.target.files[0]?.name || `选择 .xlsx ${id === "template" ? "模板" : "规则"}`;
  });
}

documentsInput.addEventListener("change", event => {
  const files = Array.from(event.target.files || []);
  fileCount.textContent = `${files.length} / 30`;
  documentsName.textContent = files.length
    ? files.map(file => file.name).join(" / ")
    : "一次选择 1～30 份 .xls / .xlsx";
  if (files.length > 30) {
    fileCount.classList.add("over-limit");
  } else {
    fileCount.classList.remove("over-limit");
  }
});

function renderConfig(config) {
  const template = config.template || {};
  const rules = config.rules || {};
  const templateUnknown = template.unknownHeaders?.length
    ? ` · ${template.unknownHeaders.length} 个未支持列`
    : "";
  activeConfig.innerHTML = `
    <div class="config-card">
      <span>当前输出模板</span>
      <strong>${escapeHtml(template.filename || "-")}</strong>
      <small>${escapeHtml(template.sheet || "-")} · 表头第 ${escapeHtml(template.headerRow || "-")} 行 · ${escapeHtml(template.supportedColumns || 0)} 个支持字段${escapeHtml(templateUnknown)}</small>
    </div>
    <div class="config-card">
      <span>当前匹配规则</span>
      <strong>${escapeHtml(rules.filename || "-")}</strong>
      <small>${escapeHtml(rules.sheet || "-")} · ${escapeHtml(rules.recordCount || 0)} 条 QAD PN 映射${rules.ambiguousGroupCount ? ` · ${escapeHtml(rules.ambiguousGroupCount)} 组归一后冲突` : ""}</small>
    </div>
  `;
}

async function loadConfig() {
  try {
    const response = await fetch("/api/config");
    const data = await responseJson(response);
    renderConfig(data.active);
  } catch (error) {
    activeConfig.innerHTML = `<div class="empty-state">配置读取失败：${escapeHtml(error.message)}</div>`;
  }
}

configForm.addEventListener("submit", async event => {
  event.preventDefault();
  const templateFile = document.querySelector("#template").files[0];
  const rulesFile = document.querySelector("#rules").files[0];
  if (!templateFile && !rulesFile) {
    alert("请选择要上传的模板或规则文件");
    return;
  }
  status("验证配置中");
  configResult.hidden = true;
  try {
    const response = await fetch("/api/config", {
      method: "POST",
      body: new FormData(configForm)
    });
    const data = await responseJson(response);
    renderConfig(data.active);
    configResult.hidden = false;
    const updated = Object.entries(data.updated || {}).map(([kind, value]) =>
      `${kind === "template" ? "模板" : "规则"}：${value.filename}`
    );
    configResult.textContent = `已启用 ${updated.join("；")}`;
    configForm.reset();
    document.querySelector("#templateName").textContent = "选择 .xlsx 模板";
    document.querySelector("#rulesName").textContent = "选择 .xlsx 规则";
    status("配置已更新", "ok");
  } catch (error) {
    status("配置更新失败", "error");
    alert(error.message);
  }
});

function matchLabel(row) {
  const type = row._rule_match_type;
  if (type === "naming_rule") {
    const applied = row._applied_part_rules?.length
      ? `（${row._applied_part_rules.join("、")}）`
      : "";
    return `命名规则匹配 → ${row._matched_qad_pn || "-"}${applied}`;
  }
  if (type === "ambiguous") return "规则冲突，未自动匹配";
  if (type === "missing") return "未匹配";
  return "精确匹配";
}

function renderPreview(preview) {
  currentPreview = preview;
  const info = preview.summary || {};
  summary.hidden = false;
  previewPanel.hidden = false;
  setField("sourceCount", info.sourceCount);
  setField("rowCount", info.rowCount);
  setField("totalQuantity", info.totalQuantity);
  setField("totalAmount", info.totalAmount);
  setField("exactCount", info.exactCount);
  setField("namingRuleCount", info.namingRuleCount);
  setField("missingCount", info.missingCount);
  previewMeta.textContent = `模板工作表：${preview.templateSheet} · 规则工作表：${preview.ruleSheet}`;

  const warningItems = preview.warnings || [];
  warningsEl.hidden = !warningItems.length;
  warningsEl.innerHTML = warningItems.map(item => `
    <div class="warning-item ${item.severity === "error" ? "warning-error" : ""}">
      <strong>${item.severity === "error" ? "红色异常" : "黄色提醒"}</strong>
      <span>${escapeHtml(item.message)}</span>
      ${item.sourceFile ? `<small>${escapeHtml(item.sourceFile)}${item.outputRow ? ` · 输出第 ${escapeHtml(item.outputRow)} 行` : ""}</small>` : ""}
    </div>
  `).join("");

  previewRows.innerHTML = (preview.rows || []).map(row => {
    const rowClass = row._rule_match_type === "naming_rule"
      ? "row-fallback"
      : ["ambiguous", "missing"].includes(row._rule_match_type)
        ? "row-missing"
        : "";
    return `
      <tr class="${rowClass}">
        <td>${escapeHtml(fmt(row.source_file))}</td>
        <td>${escapeHtml(fmt(row.serial_no))}</td>
        <td>${escapeHtml(fmt(row.item_no))}</td>
        <td>${escapeHtml(fmt(row.qad_pn))}</td>
        <td>${escapeHtml(fmt(row.quantity, 6))}</td>
        <td>${escapeHtml(fmt(row.amount, 2))}</td>
        <td>${escapeHtml(fmt(row.currency))}</td>
        <td>${escapeHtml(fmt(row.pickup_date))}</td>
        <td>${escapeHtml(fmt(row.ship_to))}</td>
        <td>${escapeHtml(fmt(row.goods_name))}</td>
        <td>${escapeHtml(fmt(row.hs_code))}</td>
        <td><span class="match-tag ${rowClass}">${escapeHtml(matchLabel(row))}</span></td>
      </tr>
    `;
  }).join("");

  previewLimit.hidden = !preview.rowsTruncated;
  previewLimit.textContent = preview.rowsTruncated
    ? `为保证页面性能，只显示前 ${preview.rowDisplayLimit} 条；最终 Excel 会包含全部 ${info.rowCount} 条。`
    : "";

  warningConfirm.checked = false;
  confirmBox.hidden = !info.requiresConfirmation;
  generateBtn.disabled = Boolean(info.requiresConfirmation);
  previewBody.hidden = false;
  previewToggle.textContent = "收起预览";
  previewToggle.setAttribute("aria-expanded", "true");
}

warningConfirm.addEventListener("change", () => {
  generateBtn.disabled = !warningConfirm.checked;
});

uploadForm.addEventListener("submit", async event => {
  event.preventDefault();
  const count = documentsInput.files.length;
  if (count < 1 || count > 30) {
    alert(`请一次选择 1～30 份源报表；当前为 ${count} 份`);
    return;
  }
  currentSession = null;
  currentHistoryId = null;
  currentPreview = null;
  previewPanel.hidden = true;
  summary.hidden = true;
  previewRows.innerHTML = "";
  status("解析中");
  try {
    const response = await fetch("/api/preview", {
      method: "POST",
      body: new FormData(uploadForm)
    });
    const data = await responseJson(response);
    currentSession = data.sessionId;
    currentHistoryId = data.historyId;
    renderPreview(data.preview);
    const needsReview = data.preview.summary?.requiresConfirmation;
    status(needsReview ? "需要复核" : "预览完成", needsReview ? "warning" : "ok");
    loadHistory();
  } catch (error) {
    status("解析失败", "error");
    alert(error.message);
  }
});

generateBtn.addEventListener("click", async () => {
  if (!currentSession || !currentHistoryId || !currentPreview) return;
  if (currentPreview.summary?.requiresConfirmation && !warningConfirm.checked) {
    alert("请先勾选确认黄色和红色异常");
    return;
  }
  generateBtn.disabled = true;
  status("生成中");
  try {
    const response = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        sessionId: currentSession,
        historyId: currentHistoryId,
        confirmed: true
      })
    });
    const data = await responseJson(response);
    status("已生成，正在下载", "ok");
    loadHistory();
    window.location.href = data.downloadUrl;
  } catch (error) {
    status("生成失败", "error");
    alert(error.message);
  } finally {
    generateBtn.disabled = Boolean(
      currentPreview?.summary?.requiresConfirmation && !warningConfirm.checked
    );
  }
});

previewToggle.addEventListener("click", () => {
  const collapse = !previewBody.hidden;
  previewBody.hidden = collapse;
  previewToggle.textContent = collapse ? "展开预览" : "收起预览";
  previewToggle.setAttribute("aria-expanded", String(!collapse));
});

function downloadLink(record, kind, label, disabled = false) {
  if (disabled) return `<span class="history-disabled">${escapeHtml(label)}</span>`;
  return `<a href="/api/history/${encodeURIComponent(record.id)}/download?kind=${encodeURIComponent(kind)}">${escapeHtml(label)}</a>`;
}

function historyTitle(record) {
  if (record.serials?.length === 1) return record.serials[0];
  if (record.serials?.length) return `${record.serials[0]} 等 ${record.sourceCount} 份`;
  return `${record.sourceCount} 份报表`;
}

function historyStatus(record) {
  return record.status === "generated" ? "已生成" : "已预览";
}

function historyItem(record, detailed = false) {
  const warningBadge = record.warningCount
    ? `<span class="history-warning">${
        record.missingCount
          ? `${record.missingCount} 条未匹配`
          : record.namingRuleCount
            ? `${record.namingRuleCount} 条命名规则匹配`
            : `${record.warningCount} 条提醒`
      }</span>`
    : `<span class="history-ok">全部精确匹配</span>`;
  return `
    <article class="history-item ${detailed ? "history-query-item" : ""}">
      <div class="history-main">
        <strong>${escapeHtml(historyTitle(record))}</strong>
        <span>${escapeHtml(record.createdAt)} · ${escapeHtml(historyStatus(record))}</span>
        ${detailed ? `<small>${escapeHtml((record.sourceNames || []).join(" / "))}</small>` : ""}
      </div>
      <div class="history-meta">
        <span>${escapeHtml(record.rowCount)} 条 · 数量 ${escapeHtml(fmt(record.totalQuantity, 6))} · 金额 ${escapeHtml(fmt(record.totalAmount, 2))}</span>
        ${warningBadge}
      </div>
      <div class="history-actions">
        ${downloadLink(record, "sources", "源报表")}
        ${downloadLink(record, "template", "当时模板")}
        ${downloadLink(record, "rules", "当时规则")}
        ${downloadLink(record, "preview", "预览 JSON")}
        ${downloadLink(record, "output", "合并结果", !record.outputName)}
        <button class="text-danger" type="button" data-delete="${escapeHtml(record.id)}">删除</button>
      </div>
    </article>
  `;
}

function renderRecentHistory(records) {
  historyList.innerHTML = records.length
    ? records.map(record => historyItem(record)).join("")
    : `<div class="empty-state">暂无历史记录</div>`;
}

async function loadHistory() {
  try {
    const response = await fetch("/api/history?limit=5");
    const data = await responseJson(response);
    renderRecentHistory(data.history || []);
  } catch (error) {
    historyList.innerHTML = `<div class="empty-state">${escapeHtml(error.message)}</div>`;
  }
}

function renderHistoryPager(pagination) {
  if (!pagination || pagination.total <= pagination.limit) {
    historyPager.hidden = true;
    historyPager.innerHTML = "";
    return;
  }
  historyPager.hidden = false;
  historyPager.innerHTML = `
    <button class="secondary" type="button" data-history-page="${pagination.page - 1}" ${pagination.page <= 1 ? "disabled" : ""}>上一页</button>
    <span>第 ${escapeHtml(pagination.page)} / ${escapeHtml(pagination.totalPages)} 页</span>
    <button class="secondary" type="button" data-history-page="${pagination.page + 1}" ${pagination.page >= pagination.totalPages ? "disabled" : ""}>下一页</button>
  `;
}

async function runHistorySearch(page = 1) {
  const params = new URLSearchParams({ limit: "50", page: String(page) });
  if (historyStart.value) params.set("start", historyStart.value);
  if (historyEnd.value) params.set("end", historyEnd.value);
  historySearchHasRun = true;
  historySearchPage = page;
  historySearchSummary.textContent = "查询中";
  historySearchList.innerHTML = `<div class="empty-state">查询中</div>`;
  try {
    const response = await fetch(`/api/history?${params.toString()}`);
    const data = await responseJson(response);
    const pagination = data.pagination || {};
    historySearchPage = pagination.page || page;
    historySearchSummary.textContent = `${data.filters?.start || "不限开始时间"} 至 ${data.filters?.end || "不限结束时间"}，共 ${pagination.total || 0} 条`;
    historySearchList.innerHTML = data.history?.length
      ? data.history.map(record => historyItem(record, true)).join("")
      : `<div class="empty-state">该时间段内没有记录</div>`;
    renderHistoryPager(pagination);
  } catch (error) {
    historySearchSummary.textContent = "查询失败";
    historySearchList.innerHTML = `<div class="empty-state">${escapeHtml(error.message)}</div>`;
  }
}

historySearchForm.addEventListener("submit", event => {
  event.preventDefault();
  runHistorySearch(1);
});

historyPager.addEventListener("click", event => {
  const button = event.target.closest("[data-history-page]");
  if (!button || button.disabled) return;
  runHistorySearch(Number(button.dataset.historyPage));
});

async function handleHistoryDelete(event) {
  const button = event.target.closest("[data-delete]");
  if (!button) return;
  if (!confirm("确定删除这条历史记录及其全部本地文件吗？")) return;
  status("删除历史中");
  try {
    const response = await fetch(
      `/api/history/${encodeURIComponent(button.dataset.delete)}`,
      { method: "DELETE" }
    );
    await responseJson(response);
    status("历史已删除", "ok");
    loadHistory();
    if (historySearchHasRun) runHistorySearch(historySearchPage);
  } catch (error) {
    status("删除失败", "error");
    alert(error.message);
  }
}

historyList.addEventListener("click", handleHistoryDelete);
historySearchList.addEventListener("click", handleHistoryDelete);
refreshHistoryBtn.addEventListener("click", loadHistory);

loadConfig();
loadHistory();
