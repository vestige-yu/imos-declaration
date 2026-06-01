let currentSession = null;
let currentPreview = null;

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
const maintenanceForms = {
  template: document.querySelector("#templateForm"),
  rules: document.querySelector("#rulesForm")
};
const maintenanceTokens = { template: null, rules: null };

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
    maintenanceTokens[id] = null;
    document.querySelector(`[data-enable="${id}"]`).disabled = true;
    document.querySelector(`#${id}PreviewWrap`).hidden = true;
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
      body: JSON.stringify({ sessionId: currentSession, preview: currentPreview })
    });
    const data = await response.json();
    if (!data.ok) throw new Error(data.error || "生成失败");
    statusEl.textContent = "已生成";
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

function renderSheetSummary(sheets) {
  return `
    <div class="sheet-list">
      ${sheets.map(sheet => `
        <div class="sheet-item">
          <strong>${escapeHtml(sheet.name)}</strong>
          <span>${escapeHtml(sheet.rows)} 行 / ${escapeHtml(sheet.columns)} 列</span>
        </div>
      `).join("")}
    </div>
  `;
}

function renderMaintenancePreview(kind, preview) {
  const wrap = document.querySelector(`#${kind}PreviewWrap`);
  const target = document.querySelector(`#${kind}Preview`);
  const samples = kind === "rules" && preview.samples?.length
    ? `
      <div class="sample-table">
        <table>
          <thead>
            <tr>
              <th>Part No</th>
              <th>HS Code</th>
              <th>商品名称</th>
              <th>品牌</th>
            </tr>
          </thead>
          <tbody>
            ${preview.samples.map(item => `
              <tr>
                <td>${escapeHtml(item.partNo)}</td>
                <td>${escapeHtml(item.hsCode || "-")}</td>
                <td>${escapeHtml(item.goodsName || "-")}</td>
                <td>${escapeHtml(item.brand || "-")}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    `
    : "";
  target.innerHTML = `
    <div class="preview-summary">
      <strong>${escapeHtml(preview.filename)}</strong>
      <span>${escapeHtml(preview.summary)}</span>
    </div>
    ${renderSheetSummary(preview.sheets || [])}
    ${samples}
  `;
  wrap.hidden = false;
  wrap.open = true;
}

async function previewMaintenanceFile(kind) {
  const form = maintenanceForms[kind];
  const input = document.querySelector(`#${kind}`);
  if (!input.files.length) throw new Error("请先选择文件");
  const body = new FormData(form);
  body.append("kind", kind);
  statusEl.textContent = kind === "template" ? "预览模板中" : "预览规则表中";
  const response = await fetch("/api/admin/preview", { method: "POST", body });
  const data = await response.json();
  if (!data.ok) throw new Error(data.error || "预览失败");
  maintenanceTokens[kind] = data.token;
  document.querySelector(`[data-enable="${kind}"]`).disabled = false;
  renderMaintenancePreview(kind, data.preview);
  statusEl.textContent = "预览完成";
}

async function enableMaintenanceFile(kind) {
  const token = maintenanceTokens[kind];
  if (!token) throw new Error("请先预览文件");
  statusEl.textContent = kind === "template" ? "启用模板中" : "启用规则表中";
  const response = await fetch("/api/admin/enable", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ token })
  });
  const data = await response.json();
  if (!data.ok) throw new Error(data.error || "启用失败");
  maintenanceTokens[kind] = null;
  document.querySelector(`[data-enable="${kind}"]`).disabled = true;
  statusEl.textContent = "维护文件已启用";
}

for (const [kind, form] of Object.entries(maintenanceForms)) {
  form.addEventListener("submit", async event => {
    event.preventDefault();
    try {
      await previewMaintenanceFile(kind);
    } catch (error) {
      statusEl.textContent = "预览失败";
      alert(error.message);
    }
  });
}

for (const button of document.querySelectorAll("[data-enable]")) {
  button.addEventListener("click", async event => {
    try {
      await enableMaintenanceFile(event.currentTarget.dataset.enable);
    } catch (error) {
      statusEl.textContent = "启用失败";
      alert(error.message);
    }
  });
}
