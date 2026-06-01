let currentSession = null;
let currentPreview = null;
let currentHistoryId = null;

const uploadForm = document.querySelector("#uploadForm");
const statusEl = document.querySelector("#status");
const summary = document.querySelector("#summary");
const previewPanel = document.querySelector("#previewPanel");
const generateBtn = document.querySelector("#generateBtn");
const previewToggle = document.querySelector("#previewToggle");
const previewBody = document.querySelector("#previewBody");
const linesEl = document.querySelector("#lines");
const warningsEl = document.querySelector("#warnings");
const documentsInput = document.querySelector("#documents");
const documentsName = document.querySelector("#documentsName");
const adminForm = document.querySelector("#adminForm");
const adminResult = document.querySelector("#adminResult");
const historyList = document.querySelector("#historyList");
const refreshHistoryBtn = document.querySelector("#refreshHistoryBtn");

documentsInput.addEventListener("change", event => {
  const files = Array.from(event.target.files || []);
  documentsName.textContent = files.length
    ? files.map(file => file.name).join(" / ")
    : "一次选择两个文件";
});

for (const id of ["template", "rules"]) {
  const input = document.querySelector(`#${id}`);
  const label = document.querySelector(`#${id}Name`);
  input.addEventListener("change", event => {
    label.textContent = event.target.files[0]?.name || "未选择";
  });
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, char => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    "\"": "&quot;",
    "'": "&#39;"
  }[char]));
}

function fmt(value) {
  if (value === null || value === undefined || value === "") return "-";
  if (typeof value === "number") return value.toLocaleString("zh-CN", { maximumFractionDigits: 2 });
  return value;
}

function setField(name, value) {
  const el = document.querySelector(`[data-field="${name}"]`);
  if (el) el.textContent = fmt(value);
}

function downloadLink(record, kind, label, disabled = false) {
  if (disabled) return `<span class="history-disabled">${escapeHtml(label)}</span>`;
  return `<a href="/api/history/${encodeURIComponent(record.id)}/download?kind=${encodeURIComponent(kind)}">${escapeHtml(label)}</a>`;
}

function renderHistory(records) {
  if (!records.length) {
    historyList.innerHTML = `<div class="empty-state">暂无历史记录</div>`;
    return;
  }
  historyList.innerHTML = records.map(record => {
    const warnings = record.warnings?.length ? `<span class="history-warning">需复核</span>` : "";
    const amount = record.totals?.amount ? `${fmt(record.totals.amount)} ${fmt(record.totals.currency)}` : "-";
    return `
      <article class="history-item" data-id="${escapeHtml(record.id)}">
        <div class="history-main">
          <strong>${escapeHtml(record.contractNo || "未命名记录")}</strong>
          <span>${escapeHtml(record.createdAt)} · ${escapeHtml(record.invoiceName)} / ${escapeHtml(record.packingName)}</span>
        </div>
        <div class="history-meta">
          <span>${escapeHtml(amount)}</span>
          ${warnings}
        </div>
        <div class="history-actions">
          ${downloadLink(record, "invoice", "Invoice")}
          ${downloadLink(record, "packing", "Packing")}
          ${downloadLink(record, "preview", "预览")}
          ${downloadLink(record, "output", "报关单", !record.outputName)}
          <button class="text-danger" type="button" data-delete="${escapeHtml(record.id)}">删除</button>
        </div>
      </article>
    `;
  }).join("");
}

async function loadHistory() {
  try {
    const response = await fetch("/api/history");
    const data = await response.json();
    if (!data.ok) throw new Error(data.error || "历史记录读取失败");
    renderHistory(data.history || []);
  } catch (error) {
    historyList.innerHTML = `<div class="empty-state">${escapeHtml(error.message)}</div>`;
  }
}

function renderPreview(preview) {
  summary.hidden = false;
  previewPanel.hidden = false;
  setField("contractNo", preview.contractNo);
  setField("packageCount", preview.packageCount);
  setField("grossWeight", preview.grossWeight);
  setField("netWeight", preview.netWeight);
  setField("totalQuantity", preview.totals.quantity);
  setField("totalAmount", `${fmt(preview.totals.amount)} ${preview.totals.currency}`);
  document.querySelector("#consignee").textContent = fmt(preview.consignee);
  document.querySelector("#tradeTerm").textContent = fmt(preview.tradeTerm);
  document.querySelector("#currency").textContent = fmt(preview.currency);

  warningsEl.hidden = !preview.warnings.length;
  warningsEl.innerHTML = preview.warnings.map(item => `<div>${escapeHtml(item)}</div>`).join("");

  linesEl.innerHTML = preview.commodityLines.map(line => `
    <tr>
      <td>${escapeHtml(fmt(line.itemNo))}</td>
      <td>${escapeHtml(fmt(line.hsCode))}</td>
      <td>${escapeHtml(fmt(line.goodsName))}</td>
      <td>${escapeHtml(fmt(line.brand))}</td>
      <td>${escapeHtml(fmt(line.quantity))}</td>
      <td>${escapeHtml(fmt(line.netWeight))}</td>
      <td>${escapeHtml(fmt(line.amount))}</td>
      <td>${escapeHtml(fmt(line.currency))}</td>
    </tr>
  `).join("");

  previewBody.hidden = false;
  previewToggle.textContent = "收起预览";
  previewToggle.setAttribute("aria-expanded", "true");
}

uploadForm.addEventListener("submit", async event => {
  event.preventDefault();
  statusEl.textContent = "解析中";
  currentSession = null;
  currentHistoryId = null;
  currentPreview = null;
  previewPanel.hidden = true;
  summary.hidden = true;
  linesEl.innerHTML = "";

  try {
    if (documentsInput.files.length < 2) {
      throw new Error("请同时选择 Invoice 和 Packing list 两个文件");
    }
    const response = await fetch("/api/parse", { method: "POST", body: new FormData(uploadForm) });
    const data = await response.json();
    if (!data.ok) throw new Error(data.error || "解析失败");
    currentSession = data.sessionId;
    currentHistoryId = data.historyId;
    currentPreview = data.preview;
    renderPreview(currentPreview);
    statusEl.textContent = data.preview.warnings.length ? "需要复核" : "预览完成";
  } catch (error) {
    statusEl.textContent = "解析失败";
    alert(error.message);
  }
});

generateBtn.addEventListener("click", async () => {
  if (!currentSession || !currentPreview) return;
  statusEl.textContent = "生成中";
  generateBtn.disabled = true;
  try {
    const response = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId: currentSession, historyId: currentHistoryId, preview: currentPreview })
    });
    const data = await response.json();
    if (!data.ok) throw new Error(data.error || "生成失败");
    statusEl.textContent = "已生成";
    currentHistoryId = data.historyId || currentHistoryId;
    loadHistory();
    window.location.href = data.downloadUrl;
  } catch (error) {
    statusEl.textContent = "生成失败";
    alert(error.message);
  } finally {
    generateBtn.disabled = false;
  }
});

previewToggle.addEventListener("click", () => {
  const collapsed = !previewBody.hidden;
  previewBody.hidden = collapsed;
  previewToggle.textContent = collapsed ? "展开预览" : "收起预览";
  previewToggle.setAttribute("aria-expanded", String(!collapsed));
});

refreshHistoryBtn.addEventListener("click", loadHistory);

historyList.addEventListener("click", async event => {
  const button = event.target.closest("[data-delete]");
  if (!button) return;
  if (!confirm("确定删除这条历史记录和本地文件吗？")) return;
  statusEl.textContent = "删除历史中";
  try {
    const response = await fetch(`/api/history/${encodeURIComponent(button.dataset.delete)}`, { method: "DELETE" });
    const data = await response.json();
    if (!data.ok) throw new Error(data.error || "删除失败");
    statusEl.textContent = "历史已删除";
    loadHistory();
  } catch (error) {
    statusEl.textContent = "删除失败";
    alert(error.message);
  }
});

adminForm.addEventListener("submit", async event => {
  event.preventDefault();
  statusEl.textContent = "上传维护文件中";
  adminResult.hidden = true;
  try {
    const response = await fetch("/api/admin/rules", {
      method: "POST",
      credentials: "same-origin",
      body: new FormData(adminForm)
    });
    const data = await response.json();
    if (!data.ok) throw new Error(data.error || "上传失败");
    adminResult.hidden = false;
    adminResult.textContent = JSON.stringify(data.updated, null, 2);
    statusEl.textContent = "维护文件已启用";
  } catch (error) {
    statusEl.textContent = "上传失败";
    alert(error.message);
  }
});

loadHistory();
