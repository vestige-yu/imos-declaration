let currentSession = null;
let currentPreview = null;

const uploadForm = document.querySelector("#uploadForm");
const statusEl = document.querySelector("#status");
const summary = document.querySelector("#summary");
const previewPanel = document.querySelector("#previewPanel");
const generateBtn = document.querySelector("#generateBtn");
const linesEl = document.querySelector("#lines");
const warningsEl = document.querySelector("#warnings");
const documentsInput = document.querySelector("#documents");
const documentsName = document.querySelector("#documentsName");

documentsInput.addEventListener("change", event => {
  const files = Array.from(event.target.files || []);
  documentsName.textContent = files.length
    ? files.map(file => file.name).join(" / ")
    : "一次选择两个文件";
});

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
  warningsEl.innerHTML = preview.warnings.map(item => `<div>${item}</div>`).join("");

  linesEl.innerHTML = preview.commodityLines.map(line => `
    <tr>
      <td>${fmt(line.itemNo)}</td>
      <td>${fmt(line.hsCode)}</td>
      <td>${fmt(line.goodsName)}</td>
      <td>${fmt(line.brand)}</td>
      <td>${fmt(line.quantity)}</td>
      <td>${fmt(line.netWeight)}</td>
      <td>${fmt(line.amount)}</td>
      <td>${fmt(line.currency)}</td>
    </tr>
  `).join("");
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
