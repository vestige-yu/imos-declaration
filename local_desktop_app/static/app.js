const state = {
  activeModule: "",
  declaration: {
    sessionId: "",
    historyId: "",
    preview: null,
    historyPage: 1,
    historySearched: false
  },
  merge: {
    sessionId: "",
    historyId: "",
    preview: null,
    historyPage: 1,
    historySearched: false
  }
};

const $ = selector => document.querySelector(selector);
const $$ = selector => Array.from(document.querySelectorAll(selector));
const statusEl = $("#status");
const breadcrumbEl = $("#breadcrumb");
const homeView = $("#homeView");

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, character => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    "\"": "&quot;",
    "'": "&#39;"
  }[character]));
}

function fmt(value, digits = 2) {
  if (value === null || value === undefined || value === "") return "-";
  if (typeof value === "number") {
    return value.toLocaleString("zh-CN", { maximumFractionDigits: digits });
  }
  return value;
}

function setStatus(text, tone = "") {
  statusEl.className = `status ${tone}`.trim();
  statusEl.textContent = text;
}

async function responseJson(response) {
  let data;
  try {
    data = await response.json();
  } catch (_) {
    throw new Error(`请求失败（${response.status}）`);
  }
  if (!response.ok || !data.ok) {
    throw new Error(data.error || `请求失败（${response.status}）`);
  }
  return data;
}

function bindFileLabel(inputSelector, labelSelector, fallback) {
  const input = $(inputSelector);
  const label = $(labelSelector);
  input.addEventListener("change", () => {
    label.textContent = input.files[0]?.name || fallback;
  });
}

function showHome() {
  state.activeModule = "";
  homeView.hidden = false;
  $$("[data-module-view]").forEach(view => {
    view.hidden = true;
  });
  breadcrumbEl.textContent = "功能首页";
  setStatus("就绪");
}

function showModule(moduleName) {
  state.activeModule = moduleName;
  homeView.hidden = true;
  $$("[data-module-view]").forEach(view => {
    view.hidden = view.dataset.moduleView !== moduleName;
  });
  breadcrumbEl.textContent = moduleName === "declaration"
    ? "功能首页 / 报关单生成"
    : "功能首页 / 报表合并";
  setStatus("就绪");
  activateTab(moduleName, "work");
  if (moduleName === "declaration") {
    loadDeclarationConfig();
    loadDeclarationHistory();
  } else {
    loadMergeConfig();
    loadMergeHistory();
  }
}

function activateTab(moduleName, tabName) {
  $$(`.tab[data-module="${moduleName}"]`).forEach(button => {
    button.classList.toggle("active", button.dataset.tab === tabName);
  });
  $$(`.tab-panel[data-module="${moduleName}"]`).forEach(panel => {
    panel.hidden = panel.dataset.panel !== tabName;
  });
  const labels = {
    declaration: {
      work: "生成报关单",
      config: "规则与模板",
      recent: "最近记录",
      history: "历史查询"
    },
    merge: {
      work: "合并报表",
      config: "规则与模板",
      recent: "最近记录",
      history: "历史查询"
    }
  };
  breadcrumbEl.textContent = `功能首页 / ${moduleName === "declaration" ? "报关单生成" : "报表合并"} / ${labels[moduleName][tabName]}`;
  if (tabName === "config") {
    moduleName === "declaration" ? loadDeclarationConfig() : loadMergeConfig();
  }
  if (tabName === "recent") {
    moduleName === "declaration" ? loadDeclarationHistory() : loadMergeHistory();
  }
}

$("#homeButton").addEventListener("click", showHome);
$$("[data-go-home]").forEach(button => button.addEventListener("click", showHome));
$$("[data-open-module]").forEach(button => {
  button.addEventListener("click", () => showModule(button.dataset.openModule));
});
$$(".tab").forEach(button => {
  button.addEventListener("click", () => activateTab(button.dataset.module, button.dataset.tab));
});

function togglePreview(body, button) {
  const collapse = !body.hidden;
  body.hidden = collapse;
  button.textContent = collapse ? "展开预览" : "收起预览";
  button.setAttribute("aria-expanded", String(!collapse));
}

function historyPagerHtml(pagination, prefix) {
  if (!pagination || pagination.total <= pagination.limit) return "";
  return `
    <button class="secondary compact-button" type="button" data-${prefix}-history-page="${pagination.page - 1}" ${pagination.page <= 1 ? "disabled" : ""}>上一页</button>
    <span>第 ${escapeHtml(pagination.page)} / ${escapeHtml(pagination.totalPages)} 页</span>
    <button class="secondary compact-button" type="button" data-${prefix}-history-page="${pagination.page + 1}" ${pagination.page >= pagination.totalPages ? "disabled" : ""}>下一页</button>
  `;
}

function downloadLink(url, label, disabled = false) {
  if (disabled) return `<span class="history-disabled">${escapeHtml(label)}</span>`;
  return `<a href="${escapeHtml(url)}">${escapeHtml(label)}</a>`;
}

// 报关单生成
const d = {
  uploadForm: $("#d-upload-form"),
  documents: $("#d-documents"),
  documentsName: $("#d-documents-name"),
  summary: $("#d-summary"),
  previewPanel: $("#d-preview-panel"),
  previewBody: $("#d-preview-body"),
  previewToggle: $("#d-preview-toggle"),
  generate: $("#d-generate"),
  lines: $("#d-lines"),
  warnings: $("#d-warnings"),
  configForm: $("#d-config-form"),
  activeConfig: $("#d-active-config"),
  configResult: $("#d-config-result"),
  historyList: $("#d-history-list"),
  historySearchForm: $("#d-history-search-form"),
  historyStart: $("#d-history-start"),
  historyEnd: $("#d-history-end"),
  historySummary: $("#d-history-search-summary"),
  historyPager: $("#d-history-pager"),
  historySearchList: $("#d-history-search-list")
};

d.documents.addEventListener("change", () => {
  const files = Array.from(d.documents.files || []);
  d.documentsName.textContent = files.length
    ? files.map(file => file.name).join(" / ")
    : "选择两个 .xls / .xlsx 文件";
});
bindFileLabel("#d-template", "#d-template-name", "选择 .xlsx 模板");
bindFileLabel("#d-rules", "#d-rules-name", "选择 .xlsx 规则表");

function renderDeclarationConfig(config) {
  const template = config.template || {};
  const rules = config.rules || {};
  d.activeConfig.innerHTML = `
    <div class="config-item">
      <span>当前报关单模板</span>
      <strong>${escapeHtml(template.filename || "-")}</strong>
      <small>${escapeHtml(template.updatedAt || "-")} · ${template.source === "uploaded" ? "客户上传" : "内置默认"}</small>
    </div>
    <div class="config-item">
      <span>当前匹配规则</span>
      <strong>${escapeHtml(rules.filename || "-")}</strong>
      <small>${escapeHtml(rules.updatedAt || "-")} · ${escapeHtml(rules.recordCount || 0)} 条料号记录${rules.conflictCount ? ` · ${escapeHtml(rules.conflictCount)} 组冲突` : ""}</small>
    </div>
  `;
}

async function loadDeclarationConfig() {
  try {
    const data = await responseJson(await fetch("/api/declaration/config"));
    renderDeclarationConfig(data.active);
  } catch (error) {
    d.activeConfig.innerHTML = `<div class="empty-state">配置读取失败：${escapeHtml(error.message)}</div>`;
  }
}

d.configForm.addEventListener("submit", async event => {
  event.preventDefault();
  if (!$("#d-template").files[0] && !$("#d-rules").files[0]) {
    alert("请选择要上传的报关单模板或匹配规则");
    return;
  }
  setStatus("验证配置中");
  d.configResult.hidden = true;
  try {
    const data = await responseJson(await fetch("/api/admin/rules", {
      method: "POST",
      body: new FormData(d.configForm)
    }));
    renderDeclarationConfig(data.active);
    const updated = Object.entries(data.updated || {}).map(([kind, item]) =>
      `${kind === "template" ? "模板" : "规则"}：${item.filename}`
    );
    d.configResult.textContent = `已启用 ${updated.join("；")}`;
    d.configResult.hidden = false;
    d.configForm.reset();
    $("#d-template-name").textContent = "选择 .xlsx 模板";
    $("#d-rules-name").textContent = "选择 .xlsx 规则表";
    setStatus("配置已更新", "ok");
  } catch (error) {
    setStatus("配置更新失败", "error");
    alert(error.message);
  }
});

function setDeclarationField(name, value) {
  const element = $(`[data-d-field="${name}"]`);
  if (element) element.textContent = fmt(value);
}

function renderDeclarationPreview(preview) {
  state.declaration.preview = preview;
  d.summary.hidden = false;
  d.previewPanel.hidden = false;
  setDeclarationField("contractNo", preview.contractNo);
  setDeclarationField("packageCount", preview.packageCount);
  setDeclarationField("grossWeight", preview.grossWeight);
  setDeclarationField("netWeight", preview.netWeight);
  setDeclarationField("totalQuantity", preview.totals?.quantity);
  setDeclarationField("totalAmount", `${fmt(preview.totals?.amount)} ${preview.totals?.currency || ""}`.trim());
  $("#d-consignee").textContent = fmt(preview.consignee);
  $("#d-trade-term").textContent = fmt(preview.tradeTerm);
  $("#d-currency").textContent = fmt(preview.currency);

  const warningMessages = [...(preview.warnings || [])];
  if (preview.sourceAnomalies?.length) {
    warningMessages.push(`上传文件中发现 ${preview.sourceAnomalies.length} 处数据异常，请重点复核。`);
    preview.sourceAnomalies.slice(0, 10).forEach(item => {
      warningMessages.push(item.message || `${item.file || "文件"} ${item.cell || ""} 数据异常`);
    });
  }
  d.warnings.hidden = !warningMessages.length;
  d.warnings.innerHTML = warningMessages.map(message => `<div>${escapeHtml(message)}</div>`).join("");
  d.lines.innerHTML = (preview.commodityLines || []).map(line => `
    <tr>
      <td>${escapeHtml(fmt(line.itemNo))}</td>
      <td>${escapeHtml(fmt(line.hsCode))}</td>
      <td>${escapeHtml(fmt(line.goodsName))}</td>
      <td>${escapeHtml(fmt(line.brand))}</td>
      <td>${escapeHtml(fmt(line.quantity, 6))}</td>
      <td>${escapeHtml(fmt(line.netWeight, 6))}</td>
      <td>${escapeHtml(fmt(line.amount, 2))}</td>
      <td>${escapeHtml(fmt(line.currency))}</td>
    </tr>
  `).join("");
  d.previewBody.hidden = false;
  d.previewToggle.textContent = "收起预览";
}

d.uploadForm.addEventListener("submit", async event => {
  event.preventDefault();
  if (d.documents.files.length !== 2) {
    alert(`请一次选择 Invoice 和 Packing list 两个文件；当前为 ${d.documents.files.length} 个`);
    return;
  }
  state.declaration.sessionId = "";
  state.declaration.historyId = "";
  state.declaration.preview = null;
  d.summary.hidden = true;
  d.previewPanel.hidden = true;
  d.lines.innerHTML = "";
  setStatus("解析中");
  try {
    const data = await responseJson(await fetch("/api/parse", {
      method: "POST",
      body: new FormData(d.uploadForm)
    }));
    state.declaration.sessionId = data.sessionId;
    state.declaration.historyId = data.historyId;
    renderDeclarationPreview(data.preview);
    const needsReview = Boolean(data.preview.warnings?.length || data.preview.sourceAnomalies?.length);
    setStatus(needsReview ? "需要复核" : "预览完成", needsReview ? "warning" : "ok");
    loadDeclarationHistory();
  } catch (error) {
    setStatus("解析失败", "error");
    alert(error.message);
  }
});

d.generate.addEventListener("click", async () => {
  if (!state.declaration.sessionId || !state.declaration.preview) return;
  d.generate.disabled = true;
  setStatus("生成中");
  try {
    const data = await responseJson(await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        sessionId: state.declaration.sessionId,
        historyId: state.declaration.historyId,
        preview: state.declaration.preview
      })
    }));
    state.declaration.historyId = data.historyId || state.declaration.historyId;
    setStatus("已生成，正在下载", "ok");
    loadDeclarationHistory();
    window.location.href = data.downloadUrl;
  } catch (error) {
    setStatus("生成失败", "error");
    alert(error.message);
  } finally {
    d.generate.disabled = false;
  }
});

d.previewToggle.addEventListener("click", () => togglePreview(d.previewBody, d.previewToggle));

function declarationHistoryItem(record, detailed = false) {
  const warningBadge = record.warnings?.length
    ? `<span class="history-warning">需复核</span>`
    : `<span class="history-ok">无警告</span>`;
  const base = `/api/history/${encodeURIComponent(record.id)}/download`;
  return `
    <article class="history-item">
      <div class="history-main">
        <strong>${escapeHtml(record.contractNo || "未命名记录")}</strong>
        <span>${escapeHtml(record.createdAt)} · ${escapeHtml(record.invoiceName || "-")} / ${escapeHtml(record.packingName || "-")}</span>
        ${detailed && record.outputName ? `<small>报关单：${escapeHtml(record.outputName)}</small>` : ""}
      </div>
      <div class="history-meta">
        <span>${escapeHtml(fmt(record.totals?.amount))} ${escapeHtml(record.totals?.currency || "")}</span>
        ${warningBadge}
      </div>
      <div class="history-actions">
        ${downloadLink(`${base}?kind=invoice`, "Invoice")}
        ${downloadLink(`${base}?kind=packing`, "Packing")}
        ${downloadLink(`${base}?kind=preview`, "预览")}
        ${downloadLink(`${base}?kind=output`, "报关单", !record.outputName)}
        <button class="text-danger" type="button" data-d-delete="${escapeHtml(record.id)}">删除</button>
      </div>
    </article>
  `;
}

async function loadDeclarationHistory() {
  try {
    const data = await responseJson(await fetch("/api/history?limit=5"));
    d.historyList.innerHTML = data.history?.length
      ? data.history.map(record => declarationHistoryItem(record)).join("")
      : `<div class="empty-state">暂无历史记录</div>`;
  } catch (error) {
    d.historyList.innerHTML = `<div class="empty-state">${escapeHtml(error.message)}</div>`;
  }
}

async function searchDeclarationHistory(page = 1) {
  const params = new URLSearchParams({ limit: "50", page: String(page) });
  if (d.historyStart.value) params.set("start", d.historyStart.value);
  if (d.historyEnd.value) params.set("end", d.historyEnd.value);
  state.declaration.historySearched = true;
  d.historySummary.textContent = "查询中";
  d.historySearchList.innerHTML = `<div class="empty-state">查询中</div>`;
  try {
    const data = await responseJson(await fetch(`/api/history?${params}`));
    const pagination = data.pagination || {};
    state.declaration.historyPage = pagination.page || page;
    d.historySummary.textContent = `${data.filters?.start || "不限开始时间"} 至 ${data.filters?.end || "不限结束时间"}，共 ${pagination.total || 0} 条；本页最多 50 条`;
    d.historySearchList.innerHTML = data.history?.length
      ? data.history.map(record => declarationHistoryItem(record, true)).join("")
      : `<div class="empty-state">该时间段内没有记录</div>`;
    d.historyPager.innerHTML = historyPagerHtml(pagination, "d");
    d.historyPager.hidden = !d.historyPager.innerHTML;
  } catch (error) {
    d.historySummary.textContent = "查询失败";
    d.historySearchList.innerHTML = `<div class="empty-state">${escapeHtml(error.message)}</div>`;
    d.historyPager.hidden = true;
  }
}

d.historySearchForm.addEventListener("submit", event => {
  event.preventDefault();
  searchDeclarationHistory(1);
});
d.historyPager.addEventListener("click", event => {
  const button = event.target.closest("[data-d-history-page]");
  if (!button || button.disabled) return;
  searchDeclarationHistory(Number(button.dataset.dHistoryPage));
});

async function handleDeclarationDelete(event) {
  const button = event.target.closest("[data-d-delete]");
  if (!button) return;
  if (!confirm("确定删除这条报关单历史记录和全部本地文件吗？")) return;
  setStatus("删除历史中");
  try {
    await responseJson(await fetch(`/api/history/${encodeURIComponent(button.dataset.dDelete)}`, {
      method: "DELETE"
    }));
    setStatus("历史已删除", "ok");
    loadDeclarationHistory();
    if (state.declaration.historySearched) {
      searchDeclarationHistory(state.declaration.historyPage);
    }
  } catch (error) {
    setStatus("删除失败", "error");
    alert(error.message);
  }
}

d.historyList.addEventListener("click", handleDeclarationDelete);
d.historySearchList.addEventListener("click", handleDeclarationDelete);
$("#d-refresh-history").addEventListener("click", loadDeclarationHistory);

// 报表合并
const m = {
  uploadForm: $("#m-upload-form"),
  documents: $("#m-documents"),
  documentsName: $("#m-documents-name"),
  fileCount: $("#m-file-count"),
  summary: $("#m-summary"),
  previewPanel: $("#m-preview-panel"),
  previewBody: $("#m-preview-body"),
  previewToggle: $("#m-preview-toggle"),
  previewMeta: $("#m-preview-meta"),
  previewRows: $("#m-preview-rows"),
  warnings: $("#m-warnings"),
  previewLimit: $("#m-preview-limit"),
  confirmBox: $("#m-confirm-box"),
  warningConfirm: $("#m-warning-confirm"),
  generate: $("#m-generate"),
  configForm: $("#m-config-form"),
  activeConfig: $("#m-active-config"),
  configResult: $("#m-config-result"),
  historyList: $("#m-history-list"),
  historySearchForm: $("#m-history-search-form"),
  historyStart: $("#m-history-start"),
  historyEnd: $("#m-history-end"),
  historySummary: $("#m-history-search-summary"),
  historyPager: $("#m-history-pager"),
  historySearchList: $("#m-history-search-list")
};

m.documents.addEventListener("change", () => {
  const files = Array.from(m.documents.files || []);
  m.fileCount.textContent = `${files.length} / 30`;
  m.fileCount.classList.toggle("status-error", files.length > 30);
  m.documentsName.textContent = files.length
    ? files.map(file => file.name).join(" / ")
    : "一次选择 1～30 份 .xls / .xlsx";
});
bindFileLabel("#m-template", "#m-template-name", "选择 .xlsx 模板");
bindFileLabel("#m-rules", "#m-rules-name", "选择 .xlsx 规则");

function renderMergeConfig(config) {
  const template = config.template || {};
  const rules = config.rules || {};
  const unknown = template.unknownHeaders?.length
    ? ` · ${template.unknownHeaders.length} 个未支持列`
    : "";
  m.activeConfig.innerHTML = `
    <div class="config-item">
      <span>当前合并生成模板</span>
      <strong>${escapeHtml(template.filename || "-")}</strong>
      <small>${escapeHtml(template.sheet || "-")} · 表头第 ${escapeHtml(template.headerRow || "-")} 行 · ${escapeHtml(template.supportedColumns || 0)} 个支持字段${escapeHtml(unknown)}</small>
    </div>
    <div class="config-item">
      <span>当前合并规则</span>
      <strong>${escapeHtml(rules.filename || "-")}</strong>
      <small>${escapeHtml(rules.sheet || "-")} · ${escapeHtml(rules.recordCount || 0)} 条 QAD PN 映射${rules.ambiguousGroupCount ? ` · ${escapeHtml(rules.ambiguousGroupCount)} 组冲突` : ""}</small>
    </div>
  `;
}

async function loadMergeConfig() {
  try {
    const data = await responseJson(await fetch("/api/merge/config"));
    renderMergeConfig(data.active);
  } catch (error) {
    m.activeConfig.innerHTML = `<div class="empty-state">配置读取失败：${escapeHtml(error.message)}</div>`;
  }
}

m.configForm.addEventListener("submit", async event => {
  event.preventDefault();
  if (!$("#m-template").files[0] && !$("#m-rules").files[0]) {
    alert("请选择要上传的合并生成模板或合并规则");
    return;
  }
  setStatus("验证配置中");
  m.configResult.hidden = true;
  try {
    const data = await responseJson(await fetch("/api/merge/config", {
      method: "POST",
      body: new FormData(m.configForm)
    }));
    renderMergeConfig(data.active);
    const updated = Object.entries(data.updated || {}).map(([kind, item]) =>
      `${kind === "template" ? "模板" : "规则"}：${item.filename}`
    );
    m.configResult.textContent = `已启用 ${updated.join("；")}`;
    m.configResult.hidden = false;
    m.configForm.reset();
    $("#m-template-name").textContent = "选择 .xlsx 模板";
    $("#m-rules-name").textContent = "选择 .xlsx 规则";
    setStatus("配置已更新", "ok");
  } catch (error) {
    setStatus("配置更新失败", "error");
    alert(error.message);
  }
});

function setMergeField(name, value) {
  const element = $(`[data-m-field="${name}"]`);
  if (element) element.textContent = fmt(value, 6);
}

function mergeMatchLabel(row) {
  if (row._rule_match_type === "naming_rule") {
    const rules = row._applied_part_rules?.length ? `（${row._applied_part_rules.join("、")}）` : "";
    return `命名规则匹配 → ${row._matched_qad_pn || "-"}${rules}`;
  }
  if (row._rule_match_type === "ambiguous") return "规则冲突，未自动匹配";
  if (row._rule_match_type === "missing") return "未匹配";
  return "精确匹配";
}

function renderMergePreview(preview) {
  state.merge.preview = preview;
  const info = preview.summary || {};
  m.summary.hidden = false;
  m.previewPanel.hidden = false;
  setMergeField("sourceCount", info.sourceCount);
  setMergeField("rowCount", info.rowCount);
  setMergeField("totalQuantity", info.totalQuantity);
  setMergeField("totalAmount", info.totalAmount);
  setMergeField("exactCount", info.exactCount);
  setMergeField("namingRuleCount", info.namingRuleCount);
  setMergeField("missingCount", info.missingCount);
  m.previewMeta.textContent = `模板工作表：${preview.templateSheet} · 规则工作表：${preview.ruleSheet}`;

  const warnings = preview.warnings || [];
  m.warnings.hidden = !warnings.length;
  m.warnings.innerHTML = warnings.map(item => `
    <div class="warning-item ${item.severity === "error" ? "warning-error" : ""}">
      <strong>${item.severity === "error" ? "红色异常" : "黄色提醒"}</strong>
      <span>${escapeHtml(item.message)}</span>
      ${item.sourceFile ? `<small>${escapeHtml(item.sourceFile)}${item.outputRow ? ` · 输出第 ${escapeHtml(item.outputRow)} 行` : ""}</small>` : ""}
    </div>
  `).join("");

  m.previewRows.innerHTML = (preview.rows || []).map(row => {
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
        <td><span class="match-tag ${rowClass}">${escapeHtml(mergeMatchLabel(row))}</span></td>
      </tr>
    `;
  }).join("");
  m.previewLimit.hidden = !preview.rowsTruncated;
  m.previewLimit.textContent = preview.rowsTruncated
    ? `页面只显示前 ${preview.rowDisplayLimit} 条；最终 Excel 会包含全部 ${info.rowCount} 条。`
    : "";
  m.warningConfirm.checked = false;
  m.confirmBox.hidden = !info.requiresConfirmation;
  m.generate.disabled = Boolean(info.requiresConfirmation);
  m.previewBody.hidden = false;
  m.previewToggle.textContent = "收起预览";
}

m.warningConfirm.addEventListener("change", () => {
  m.generate.disabled = !m.warningConfirm.checked;
});

m.uploadForm.addEventListener("submit", async event => {
  event.preventDefault();
  const count = m.documents.files.length;
  if (count < 1 || count > 30) {
    alert(`请一次选择 1～30 份源报表；当前为 ${count} 份`);
    return;
  }
  state.merge.sessionId = "";
  state.merge.historyId = "";
  state.merge.preview = null;
  m.summary.hidden = true;
  m.previewPanel.hidden = true;
  m.previewRows.innerHTML = "";
  setStatus("解析中");
  try {
    const data = await responseJson(await fetch("/api/merge/preview", {
      method: "POST",
      body: new FormData(m.uploadForm)
    }));
    state.merge.sessionId = data.sessionId;
    state.merge.historyId = data.historyId;
    renderMergePreview(data.preview);
    const needsReview = Boolean(data.preview.summary?.requiresConfirmation);
    setStatus(needsReview ? "需要复核" : "预览完成", needsReview ? "warning" : "ok");
    loadMergeHistory();
  } catch (error) {
    setStatus("解析失败", "error");
    alert(error.message);
  }
});

m.generate.addEventListener("click", async () => {
  if (!state.merge.sessionId || !state.merge.historyId || !state.merge.preview) return;
  if (state.merge.preview.summary?.requiresConfirmation && !m.warningConfirm.checked) {
    alert("请先勾选确认黄色和红色异常");
    return;
  }
  m.generate.disabled = true;
  setStatus("生成中");
  try {
    const data = await responseJson(await fetch("/api/merge/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        sessionId: state.merge.sessionId,
        historyId: state.merge.historyId,
        confirmed: true
      })
    }));
    setStatus("已生成，正在下载", "ok");
    loadMergeHistory();
    window.location.href = data.downloadUrl;
  } catch (error) {
    setStatus("生成失败", "error");
    alert(error.message);
  } finally {
    m.generate.disabled = Boolean(
      state.merge.preview?.summary?.requiresConfirmation && !m.warningConfirm.checked
    );
  }
});

m.previewToggle.addEventListener("click", () => togglePreview(m.previewBody, m.previewToggle));

function mergeHistoryTitle(record) {
  if (record.serials?.length === 1) return record.serials[0];
  if (record.serials?.length) return `${record.serials[0]} 等 ${record.sourceCount} 份`;
  return `${record.sourceCount} 份报表`;
}

function mergeHistoryItem(record, detailed = false) {
  const warningBadge = record.warningCount
    ? `<span class="history-warning">${record.missingCount ? `${record.missingCount} 条未匹配` : `${record.warningCount} 条提醒`}</span>`
    : `<span class="history-ok">全部精确匹配</span>`;
  const base = `/api/merge/history/${encodeURIComponent(record.id)}/download`;
  return `
    <article class="history-item">
      <div class="history-main">
        <strong>${escapeHtml(mergeHistoryTitle(record))}</strong>
        <span>${escapeHtml(record.createdAt)} · ${record.status === "generated" ? "已生成" : "已预览"}</span>
        ${detailed ? `<small>${escapeHtml((record.sourceNames || []).join(" / "))}</small>` : ""}
      </div>
      <div class="history-meta">
        <span>${escapeHtml(record.rowCount)} 条 · 数量 ${escapeHtml(fmt(record.totalQuantity, 6))} · 金额 ${escapeHtml(fmt(record.totalAmount, 2))}</span>
        ${warningBadge}
      </div>
      <div class="history-actions">
        ${downloadLink(`${base}?kind=sources`, "源报表")}
        ${downloadLink(`${base}?kind=template`, "当时模板")}
        ${downloadLink(`${base}?kind=rules`, "当时规则")}
        ${downloadLink(`${base}?kind=preview`, "预览")}
        ${downloadLink(`${base}?kind=output`, "合并结果", !record.outputName)}
        <button class="text-danger" type="button" data-m-delete="${escapeHtml(record.id)}">删除</button>
      </div>
    </article>
  `;
}

async function loadMergeHistory() {
  try {
    const data = await responseJson(await fetch("/api/merge/history?limit=5"));
    m.historyList.innerHTML = data.history?.length
      ? data.history.map(record => mergeHistoryItem(record)).join("")
      : `<div class="empty-state">暂无历史记录</div>`;
  } catch (error) {
    m.historyList.innerHTML = `<div class="empty-state">${escapeHtml(error.message)}</div>`;
  }
}

async function searchMergeHistory(page = 1) {
  const params = new URLSearchParams({ limit: "50", page: String(page) });
  if (m.historyStart.value) params.set("start", m.historyStart.value);
  if (m.historyEnd.value) params.set("end", m.historyEnd.value);
  state.merge.historySearched = true;
  m.historySummary.textContent = "查询中";
  m.historySearchList.innerHTML = `<div class="empty-state">查询中</div>`;
  try {
    const data = await responseJson(await fetch(`/api/merge/history?${params}`));
    const pagination = data.pagination || {};
    state.merge.historyPage = pagination.page || page;
    m.historySummary.textContent = `${data.filters?.start || "不限开始时间"} 至 ${data.filters?.end || "不限结束时间"}，共 ${pagination.total || 0} 条；本页最多 50 条`;
    m.historySearchList.innerHTML = data.history?.length
      ? data.history.map(record => mergeHistoryItem(record, true)).join("")
      : `<div class="empty-state">该时间段内没有记录</div>`;
    m.historyPager.innerHTML = historyPagerHtml(pagination, "m");
    m.historyPager.hidden = !m.historyPager.innerHTML;
  } catch (error) {
    m.historySummary.textContent = "查询失败";
    m.historySearchList.innerHTML = `<div class="empty-state">${escapeHtml(error.message)}</div>`;
    m.historyPager.hidden = true;
  }
}

m.historySearchForm.addEventListener("submit", event => {
  event.preventDefault();
  searchMergeHistory(1);
});
m.historyPager.addEventListener("click", event => {
  const button = event.target.closest("[data-m-history-page]");
  if (!button || button.disabled) return;
  searchMergeHistory(Number(button.dataset.mHistoryPage));
});

async function handleMergeDelete(event) {
  const button = event.target.closest("[data-m-delete]");
  if (!button) return;
  if (!confirm("确定删除这条合并历史记录和全部本地文件吗？")) return;
  setStatus("删除历史中");
  try {
    await responseJson(await fetch(`/api/merge/history/${encodeURIComponent(button.dataset.mDelete)}`, {
      method: "DELETE"
    }));
    setStatus("历史已删除", "ok");
    loadMergeHistory();
    if (state.merge.historySearched) {
      searchMergeHistory(state.merge.historyPage);
    }
  } catch (error) {
    setStatus("删除失败", "error");
    alert(error.message);
  }
}

m.historyList.addEventListener("click", handleMergeDelete);
m.historySearchList.addEventListener("click", handleMergeDelete);
$("#m-refresh-history").addEventListener("click", loadMergeHistory);

showHome();
