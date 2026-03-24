import * as pdfjsLib from "/vendor/pdfjs/pdf.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc = "/vendor/pdfjs/pdf.worker.mjs";

const form = document.getElementById("generator-form");
const templateSelect = document.getElementById("templateId");
const templateLibraryEl = document.getElementById("templateLibrary");
const templateCountEl = document.getElementById("templateCount");
const excelInput = document.getElementById("excel");
const fileStatusEl = document.getElementById("fileStatus");
const statusEl = document.getElementById("status");
const footerSummaryEl = document.getElementById("footerSummary");
const submitBtn = document.getElementById("submitBtn");
const refreshPreviewBtn = document.getElementById("refreshPreviewBtn");
const sendFeishuBtn = document.getElementById("sendFeishuBtn");
const downloadBtn = document.getElementById("downloadBtn");
const loadingOverlayEl = document.getElementById("loadingOverlay");
const loadingMessageEl = document.getElementById("loadingMessage");
const workspacePanel = document.getElementById("workspacePanel");
const employeeCountEl = document.getElementById("employeeCount");
const employeeListEl = document.getElementById("employeeList");
const editorForm = document.getElementById("editorForm");
const cardStage = document.getElementById("cardStage");
const cardStageWrap = document.getElementById("cardStageWrap");
const selectedFieldHint = document.getElementById("selectedFieldHint");
const previewPanel = document.getElementById("previewPanel");
const previewFrame = document.getElementById("previewFrame");
const previewCanvas = document.getElementById("previewCanvas");
const previewPlaceholder = document.getElementById("previewPlaceholder");

let currentPdfUrl = "";
let currentPreviewBytes = null;
let currentPreviewDocumentTask = null;
let currentPreviewRenderTask = null;
let previewResizeTimer = null;
let templateCatalog = [];
let currentEmployees = [];
let currentColumns = [];
let currentFileNameField = "";
let selectedEmployeeId = "";
let selectedFieldKey = "";
let currentRenderModel = null;
let inlineEditorEl = null;
let dragState = null;
let suppressClickUntil = 0;
let previewRefreshTimer = null;
let lastStatusMessage = "待开始";
let lastStatusType = "";
let hoveredTemplateId = "";
const DRAG_SNAP_THRESHOLD = 1.2;
const AUTO_PREVIEW_DELAY_MS = 500;
const TEMPLATE_META = {
  "fjd-japanese-v2026.3": { name: "FJDynamics 日文名片", note: "适用于日本业务", badge: "JP", theme: "jp" },
  "fjd-hk-traditional-v2026.3": { name: "FJDynamics 香港繁体名片", note: "适用于港澳业务", badge: "HK", theme: "hk" },
  "fjd-english-v2025.6": { name: "FJDynamics 英文名片", note: "适用于海外业务", badge: "EN", theme: "en" },
  "fjd-chinese-v2026.3": { name: "FJDynamics 中文名片", note: "适用于中国大陆业务", badge: "CN", theme: "cn" },
  "svea-japanese-v2026.3": { name: "SVEA 日文名片", note: "适用于日本团队", badge: "JP", theme: "jp" },
  "svea-english-v2026.3": { name: "SVEA 英文名片", note: "适用于海外团队", badge: "EN", theme: "en" }
};
const TEMPLATE_ORDER = [
  "fjd-english-v2025.6",
  "fjd-chinese-v2026.3",
  "fjd-japanese-v2026.3",
  "fjd-hk-traditional-v2026.3",
  "svea-english-v2026.3",
  "svea-japanese-v2026.3"
];

function getElementKey(elementModel) {
  return elementModel.positionKey || elementModel.fieldKey || "";
}

function setStatus(message, type = "") {
  lastStatusMessage = message || "待开始";
  lastStatusType = type;
  statusEl.textContent = message;
  statusEl.className = `status-banner ${type}`.trim();
  updateFooterSummary();
}

function setLoadingState(visible, message = "PDF 正在生成，请稍候...") {
  loadingOverlayEl.classList.toggle("hidden", !visible);
  loadingOverlayEl.setAttribute("aria-hidden", String(!visible));
  loadingMessageEl.textContent = message;
}

function getTemplateMeta(templateLike) {
  const templateId =
    typeof templateLike === "string"
      ? templateLike
      : templateLike?.id || templateLike?.value || templateSelect.value || "";
  const fallbackName =
    typeof templateLike === "object" && templateLike
      ? templateLike.name || templateLike.textContent || templateLike.label || templateId
      : templateSelect.selectedOptions[0]?.textContent || templateId;
  const predefined = TEMPLATE_META[templateId] || {};

  return {
    id: templateId,
    name: predefined.name || fallbackName || "未命名模板",
    note: predefined.note || "支持批量 PDF 导出",
    badge: predefined.badge || "BC",
    theme: predefined.theme || "cn"
  };
}

function getCurrentTemplateMeta() {
  return getTemplateMeta(
    templateCatalog.find((template) => template.id === templateSelect.value) ||
      templateSelect.selectedOptions[0] ||
      templateSelect.value
  );
}

function updateFooterSummary() {
  const templateMeta = getCurrentTemplateMeta();
  const employeeCount = currentEmployees.length;
  const statusText =
    lastStatusType === "error"
      ? `异常：${lastStatusMessage}`
      : lastStatusType === "success"
        ? lastStatusMessage || "完成"
        : lastStatusMessage || "待开始";

  footerSummaryEl.textContent =
    `当前模板：${templateMeta.name} | 已加载 ${employeeCount} 条员工信息 | 生成状态：${statusText}`;
}

function updateFileStatus() {
  const fileCount = excelInput.files?.length || 0;
  if (fileCount === 0) {
    fileStatusEl.textContent = "未选择文件";
    return;
  }

  if (currentEmployees.length > 0) {
    fileStatusEl.textContent = `已上传 ${currentEmployees.length} 条数据`;
    return;
  }

  fileStatusEl.textContent = `已选择 ${fileCount} 个文件`;
}

function updateActionAvailability() {
  const disabled = currentEmployees.length === 0;
  refreshPreviewBtn.disabled = disabled;
  sendFeishuBtn.disabled = disabled;
  downloadBtn.disabled = disabled;
}

function renderTemplateShowcase(templateMeta, message = "上传员工信息后自动生成对应名片内容") {
  if (!templateMeta?.id) {
    cardStage.innerHTML = "";
    return;
  }

  cardStage.innerHTML = `
    <div class="template-showcase">
      <div class="template-showcase-brand">FJDynamics</div>
      <div class="template-showcase-lang">${templateMeta.name}</div>
      <div class="template-showcase-note">${message}</div>
    </div>
  `;
}

function renderTemplateLibrary() {
  templateLibraryEl.innerHTML = "";
  templateCountEl.textContent = String(templateCatalog.length);

  templateCatalog.forEach((template) => {
    const templateMeta = getTemplateMeta(template);
    const button = document.createElement("button");
    button.type = "button";
    button.className = `template-card ${template.id === templateSelect.value ? "active" : ""}`.trim();
    button.dataset.templateId = template.id;
    button.innerHTML = `
      <div class="template-thumb theme-${templateMeta.theme}">
        <span>${templateMeta.badge}</span>
      </div>
      <div class="template-content">
        <strong>${templateMeta.name}</strong>
        <span>${templateMeta.note}</span>
      </div>
    `;

    button.addEventListener("mouseenter", () => {
      hoveredTemplateId = template.id;
      if (!currentEmployees.length) {
        renderTemplateShowcase(templateMeta, "悬停模板可先查看空白样式");
      }
    });

    button.addEventListener("mouseleave", () => {
      hoveredTemplateId = "";
      if (!currentEmployees.length) {
        renderTemplateShowcase(getCurrentTemplateMeta());
      }
    });

    button.addEventListener("click", () => {
      templateSelect.value = template.id;
      currentColumns = template.columns || [];
      currentFileNameField = template.fileNameField || "";
      renderTemplateLibrary();
      clearWorkspace();
      setStatus(`已切换模板：${templateMeta.name}`, "success");
    });

    templateLibraryEl.appendChild(button);
  });
}

function decodeContentDispositionFileName(value) {
  if (!value) {
    return "";
  }

  const utf8Match = value.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf8Match?.[1]) {
    try {
      return decodeURIComponent(utf8Match[1]);
    } catch (_error) {
      return utf8Match[1];
    }
  }

  const quotedMatch = value.match(/filename="([^"]+)"/i);
  if (quotedMatch?.[1]) {
    return quotedMatch[1];
  }

  const plainMatch = value.match(/filename=([^;]+)/i);
  return plainMatch?.[1]?.trim() || "";
}

function triggerBlobDownload(blob, fileName = "") {
  const downloadUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = downloadUrl;
  if (fileName) {
    link.download = fileName;
  }
  link.click();
  URL.revokeObjectURL(downloadUrl);
}

function cancelPreviewRender() {
  currentPreviewRenderTask?.cancel?.();
  currentPreviewRenderTask = null;
  currentPreviewDocumentTask?.destroy?.();
  currentPreviewDocumentTask = null;
}

function resetPreviewCanvas() {
  const context = previewCanvas.getContext("2d");
  context.clearRect(0, 0, previewCanvas.width, previewCanvas.height);
  previewCanvas.width = 0;
  previewCanvas.height = 0;
  previewCanvas.style.width = "0";
  previewCanvas.style.height = "0";
}

async function renderPreviewImageFromBytes() {
  if (!currentPreviewBytes?.length) {
    return;
  }

  cancelPreviewRender();
  previewFrame.classList.add("loading");
  previewPlaceholder.textContent = "正在渲染 PDF 预览...";

  const loadingTask = pdfjsLib.getDocument({ data: currentPreviewBytes.slice() });
  currentPreviewDocumentTask = loadingTask;

  try {
    const pdfDocument = await loadingTask.promise;
    const page = await pdfDocument.getPage(1);
    const baseViewport = page.getViewport({ scale: 1 });
    const availableWidth = Math.max(previewFrame.clientWidth - 28, 180);
    const availableHeight = Math.max(previewFrame.clientHeight - 28, 180);
    const cssScale = Math.min(
      availableWidth / baseViewport.width,
      availableHeight / baseViewport.height
    );
    const pixelRatio = window.devicePixelRatio || 1;
    const renderScale = cssScale * pixelRatio;
    const renderViewport = page.getViewport({ scale: renderScale });
    const cssWidth = Math.floor(baseViewport.width * cssScale);
    const cssHeight = Math.floor(baseViewport.height * cssScale);
    const context = previewCanvas.getContext("2d");

    previewCanvas.width = Math.ceil(renderViewport.width);
    previewCanvas.height = Math.ceil(renderViewport.height);
    previewCanvas.style.width = `${cssWidth}px`;
    previewCanvas.style.height = `${cssHeight}px`;

    context.setTransform(1, 0, 0, 1, 0, 0);
    context.clearRect(0, 0, previewCanvas.width, previewCanvas.height);

    const renderTask = page.render({
      canvasContext: context,
      viewport: renderViewport
    });
    currentPreviewRenderTask = renderTask;
    await renderTask.promise;
    currentPreviewRenderTask = null;
    previewFrame.classList.remove("loading");
  } catch (error) {
    if (error?.name === "RenderingCancelledException") {
      return;
    }

    resetPreviewCanvas();
    previewFrame.classList.remove("loading");
    previewPlaceholder.textContent = "预览渲染失败";
    throw error;
  } finally {
    currentPreviewDocumentTask = null;
  }
}

function clearPreview() {
  if (currentPdfUrl) {
    URL.revokeObjectURL(currentPdfUrl);
  }

  clearTimeout(previewResizeTimer);
  currentPdfUrl = "";
  currentPreviewBytes = null;
  cancelPreviewRender();
  resetPreviewCanvas();
  previewPlaceholder.textContent = "等待生成 PDF 预览";
  previewFrame.classList.add("loading");
}

function cancelScheduledPreviewRefresh() {
  if (previewRefreshTimer) {
    clearTimeout(previewRefreshTimer);
    previewRefreshTimer = null;
  }
}

function clearInlineEditor() {
  inlineEditorEl?.remove();
  inlineEditorEl = null;
}

function clearWorkspace() {
  clearInlineEditor();
  clearPreview();
  currentEmployees = [];
  currentColumns = [];
  currentFileNameField = "";
  selectedEmployeeId = "";
  selectedFieldKey = "";
  currentRenderModel = null;
  dragState = null;
  suppressClickUntil = 0;
  cancelScheduledPreviewRefresh();
  employeeListEl.innerHTML = "";
  editorForm.innerHTML = "";
  employeeCountEl.textContent = "0 条";
  selectedFieldHint.textContent = "当前未选中字段";
  cardStage.innerHTML = "";
  cardStage.removeAttribute("style");
  renderEmployeeList();
  renderEditorForm();
  renderTemplateShowcase(getCurrentTemplateMeta(), "上传员工信息后自动生成对应名片内容");
  updateActionAvailability();
  updateFileStatus();
  updateFooterSummary();
}

function getSelectedEmployee() {
  return currentEmployees.find((employee) => employee.id === selectedEmployeeId) || null;
}

function getEmployeeDisplayName(employee) {
  const rawName = employee?.row?.[currentFileNameField]?.trim();
  return rawName || employee?.displayName || "未命名员工";
}

function setSelectedField(fieldKey) {
  selectedFieldKey = fieldKey || "";
  selectedFieldHint.textContent = selectedFieldKey
    ? `当前字段：${selectedFieldKey}`
    : "当前未选中字段";

  cardStage.querySelectorAll(".card-element").forEach((node) => {
    node.classList.toggle("selected", node.dataset.fieldKey === selectedFieldKey);
  });

  editorForm.querySelectorAll("input, textarea").forEach((node) => {
    const isMatch = node.name === selectedFieldKey;
    node.closest(".field")?.classList.toggle("selected-field", isMatch);
  });
}

function createEmployeeListItem(employee) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `employee-item ${employee.id === selectedEmployeeId ? "active" : ""}`.trim();

  const nameEl = document.createElement("strong");
  nameEl.textContent = getEmployeeDisplayName(employee);
  button.appendChild(nameEl);

  if (employee.sourceFileName) {
    const metaEl = document.createElement("span");
    metaEl.textContent = employee.sourceFileName;
    button.appendChild(metaEl);
  }

  button.addEventListener("click", async () => {
    if (selectedEmployeeId === employee.id) {
      return;
    }

    clearInlineEditor();
    selectedEmployeeId = employee.id;
    selectedFieldKey = "";
    renderEmployeeList();
    renderEditorForm();
    await refreshVisualEditor();
    await refreshPdfPreview();
  });

  return button;
}

function renderEmployeeList() {
  employeeListEl.innerHTML = "";

  if (!currentEmployees.length) {
    const empty = document.createElement("div");
    empty.className = "empty-card";
    empty.textContent = "上传员工信息表后，这里会显示待生成的员工名片列表。";
    employeeListEl.appendChild(empty);
    employeeCountEl.textContent = "0 条";
    return;
  }

  currentEmployees.forEach((employee) => {
    employeeListEl.appendChild(createEmployeeListItem(employee));
  });

  employeeCountEl.textContent = `${currentEmployees.length} 条`;
}

function updateEmployeeField(fieldKey, value) {
  const employee = getSelectedEmployee();
  if (!employee) {
    return;
  }

  employee.row[fieldKey] = value;
  renderEmployeeList();
}

function shouldIgnoreKeyboardShortcut(target) {
  if (!target) {
    return false;
  }

  return ["INPUT", "TEXTAREA", "SELECT"].includes(target.tagName) || target.isContentEditable;
}

function createFieldControl(column, value) {
  const field = document.createElement("label");
  field.className = "field";

  const label = document.createElement("span");
  label.textContent = column;
  field.appendChild(label);

  const isLongText = /address|mobile|whats|鍦板潃|鎵嬫満/i.test(column);
  const control = document.createElement(isLongText ? "textarea" : "input");
  control.name = column;
  control.value = value || "";

  if (isLongText) {
    control.rows = 4;
  } else {
    control.type = "text";
  }

  control.addEventListener("focus", () => {
    setSelectedField(column);
  });

  control.addEventListener("input", (event) => {
    updateEmployeeField(column, event.target.value);
  });

  control.addEventListener("change", async () => {
    await refreshVisualEditor();
    schedulePreviewRefresh();
  });

  field.appendChild(control);
  return field;
}

function renderEditorForm() {
  const employee = getSelectedEmployee();
  editorForm.innerHTML = "";

  if (!employee) {
    const empty = document.createElement("div");
    empty.className = "empty-form";
    empty.textContent = "加载数据后，可在此修改字段内容，并与左侧可视化画布实时联动。";
    editorForm.appendChild(empty);
    return;
  }

  currentColumns.forEach((column) => {
    editorForm.appendChild(createFieldControl(column, employee.row[column]));
  });

  setSelectedField(selectedFieldKey);
}

function syncEditorFieldValue(fieldKey, value) {
  const control = editorForm.querySelector(`[name="${CSS.escape(fieldKey)}"]`);
  if (control) {
    control.value = value;
  }
}

function pageToStageScale(page) {
  return cardStage.clientWidth / page.width;
}

function pageYToTop(page, y) {
  return page.height - y;
}

function baselineToTop(page, y, size, scale) {
  return pageYToTop(page, y) * scale - size * scale * 0.84;
}

function topToBaseline(page, top, size, scale) {
  return page.height - top / scale - size * 0.84;
}

function applyPositionOverrideToEmployee(positionKey, x, y) {
  const employee = getSelectedEmployee();
  if (!employee || !positionKey) {
    return;
  }

  employee.fieldPositions ||= {};
  employee.fieldPositions[positionKey] = {
    x: Number(x.toFixed(2)),
    y: Number(y.toFixed(2))
  };
}

function applyPositionOverrideToRenderModel(positionKey, x, y) {
  const element = currentRenderModel?.elements?.find((item) => getElementKey(item) === positionKey);
  if (!element) {
    return;
  }

  if (element.type === "multiline-text") {
    const deltaY = y - element.y;
    element.x = x;
    element.y = y;
    element.lines = element.lines.map((line) => ({
      ...line,
      x,
      y: line.y + deltaY
    }));
    return;
  }

  element.x = x;
  element.y = y;
}

function schedulePreviewRefresh(delay = AUTO_PREVIEW_DELAY_MS) {
  if (!currentEmployees.length) {
    return;
  }

  cancelScheduledPreviewRefresh();
  previewRefreshTimer = window.setTimeout(async () => {
    previewRefreshTimer = null;
    await refreshPdfPreview({ silent: true });
  }, delay);
}

function getElementGuideOrigin(elementModel) {
  if (elementModel.type === "multiline-text") {
    return {
      x: elementModel.x ?? elementModel.lines[0]?.x ?? 0,
      y: elementModel.y ?? elementModel.lines[0]?.y ?? 0
    };
  }

  return {
    x: elementModel.x,
    y: elementModel.y
  };
}

function getGuideCandidates(positionKey) {
  if (!currentRenderModel) {
    return { vertical: [], horizontal: [] };
  }

  const vertical = [];
  const horizontal = [];

  currentRenderModel.elements.forEach((elementModel) => {
    if (getElementKey(elementModel) === positionKey) {
      return;
    }

    const origin = getElementGuideOrigin(elementModel);
    if (Number.isFinite(origin.x)) {
      vertical.push(origin.x);
    }
    if (Number.isFinite(origin.y)) {
      horizontal.push(origin.y);
    }
  });

  currentRenderModel.decorations.forEach((decoration) => {
    if (decoration.type === "line") {
      if (Number.isFinite(decoration.x)) {
        vertical.push(decoration.x, decoration.x + (decoration.width || 0));
      }
      if (Number.isFinite(decoration.y)) {
        horizontal.push(decoration.y);
      }
    }
  });

  return { vertical, horizontal };
}

function snapValue(value, candidates) {
  let snapped = value;
  let matched = null;
  let minDistance = Number.POSITIVE_INFINITY;

  candidates.forEach((candidate) => {
    const distance = Math.abs(candidate - value);
    if (distance <= DRAG_SNAP_THRESHOLD && distance < minDistance) {
      minDistance = distance;
      snapped = candidate;
      matched = candidate;
    }
  });

  return { value: snapped, guide: matched };
}

function renderDragGuides(guides, page, scale) {
  cardStage.querySelectorAll(".card-guide").forEach((guide) => guide.remove());

  if (guides.vertical !== null && guides.vertical !== undefined) {
    const guide = document.createElement("div");
    guide.className = "card-guide vertical";
    guide.style.left = `${guides.vertical * scale}px`;
    guide.style.top = "0";
    guide.style.height = `${page.height * scale}px`;
    cardStage.appendChild(guide);
  }

  if (guides.horizontal !== null && guides.horizontal !== undefined) {
    const guide = document.createElement("div");
    guide.className = "card-guide horizontal";
    guide.style.left = "0";
    guide.style.top = `${pageYToTop(page, guides.horizontal) * scale}px`;
    guide.style.width = `${page.width * scale}px`;
    cardStage.appendChild(guide);
  }
}

function clearDragGuides() {
  cardStage.querySelectorAll(".card-guide").forEach((guide) => guide.remove());
}

function createDecoration(decoration, page, scale) {
  const node = document.createElement("div");
  node.className = `card-decoration ${decoration.type}`.trim();

  if (decoration.type === "line") {
    node.style.left = `${decoration.x * scale}px`;
    node.style.top = `${pageYToTop(page, decoration.y + (decoration.thickness || 0)) * scale}px`;
    node.style.width = `${decoration.width * scale}px`;
    node.style.height = `${(decoration.thickness || 0.8) * scale}px`;
    node.style.background = decoration.color || "#111";
    return node;
  }

  if (decoration.type === "rect") {
    node.style.left = `${decoration.x * scale}px`;
    node.style.top = `${pageYToTop(page, decoration.y + decoration.height) * scale}px`;
    node.style.width = `${decoration.width * scale}px`;
    node.style.height = `${decoration.height * scale}px`;
    node.style.background = decoration.background || "#fff";
    node.style.border = `${(decoration.borderWidth || 1) * scale}px solid ${decoration.borderColor || "#111"}`;
    node.textContent = decoration.label || "";
    return node;
  }

  if (decoration.type === "text") {
    node.style.left = `${decoration.x * scale}px`;
    node.style.top = `${baselineToTop(page, decoration.y, decoration.size, scale)}px`;
    node.style.fontSize = `${decoration.size * scale}px`;
    node.style.fontWeight = decoration.fontWeight === "bold" ? "700" : "400";
    if (decoration.fontFamily) {
      node.style.fontFamily = decoration.fontFamily;
    }
    node.style.color = decoration.color || "#111";
    node.textContent = decoration.text || "";
  }

  return node;
}

function beginInlineEdit(elementModel, targetNode, page, scale) {
  if (!elementModel.textEditable) {
    return;
  }

  clearInlineEditor();

  const employee = getSelectedEmployee();
  if (!employee) {
    return;
  }

  const currentValue = employee.row[elementModel.fieldKey] || "";
  const isMultiline =
    elementModel.type === "multiline-text" ||
    elementModel.fieldKey === "Mobile Number" ||
    elementModel.fieldKey === "WhatsApp Number" ||
    elementModel.fieldKey === "鎵嬫満" ||
    elementModel.fieldKey === "Mobile";
  const editor = document.createElement(isMultiline ? "textarea" : "input");
  editor.className = "card-inline-editor";
  editor.value = currentValue;

  if (!isMultiline) {
    editor.type = "text";
  }

  const left = parseFloat(targetNode.style.left || "0");
  const top = parseFloat(targetNode.style.top || "0");
  const width = Math.max(targetNode.offsetWidth + 24, elementModel.maxWidth ? elementModel.maxWidth * scale : 160);
  const height = isMultiline
    ? Math.max(targetNode.offsetHeight + 24, 110)
    : Math.max(targetNode.offsetHeight + 18, 44);

  editor.style.left = `${left - 8}px`;
  editor.style.top = `${top - 8}px`;
  editor.style.width = `${width}px`;
  editor.style.height = `${height}px`;
  editor.style.fontSize = `${Math.max(elementModel.size * scale, 14)}px`;
  editor.style.fontWeight = elementModel.fontWeight === "bold" ? "700" : "400";
  if (elementModel.fontFamily) {
    editor.style.fontFamily = elementModel.fontFamily;
  }

  const commit = async () => {
    const value = editor.value.trim();
    updateEmployeeField(elementModel.fieldKey, value);
    syncEditorFieldValue(elementModel.fieldKey, value);
    clearInlineEditor();
    await refreshVisualEditor();
    schedulePreviewRefresh();
    setSelectedField(getElementKey(elementModel));
  };

  editor.addEventListener("blur", async () => {
    await commit();
  });

  editor.addEventListener("keydown", async (event) => {
    if (!isMultiline && event.key === "Enter") {
      event.preventDefault();
      editor.blur();
    }

    if (event.key === "Escape") {
      clearInlineEditor();
      await refreshVisualEditor();
      setSelectedField(getElementKey(elementModel));
    }
  });

  cardStage.appendChild(editor);
  inlineEditorEl = editor;
  editor.focus();
  editor.select();
}

function startDrag(event, elementModel, node, page, scale) {
  if (!elementModel.editable || event.button !== 0 || inlineEditorEl) {
    return;
  }

  event.preventDefault();
  setSelectedField(getElementKey(elementModel));

  const startLeft = parseFloat(node.style.left || "0");
  const startTop = parseFloat(node.style.top || "0");

  dragState = {
    fieldKey: getElementKey(elementModel),
    node,
    page,
    scale,
    size: elementModel.size,
    maxWidth: elementModel.maxWidth ?? null,
    startX: event.clientX,
    startY: event.clientY,
    originLeft: startLeft,
    originTop: startTop,
    dragged: false,
    guides: getGuideCandidates(getElementKey(elementModel))
  };

  node.classList.add("dragging");

  const handleMove = (moveEvent) => {
    if (!dragState) {
      return;
    }

    const deltaX = moveEvent.clientX - dragState.startX;
    const deltaY = moveEvent.clientY - dragState.startY;
    if (!dragState.dragged && (Math.abs(deltaX) > 2 || Math.abs(deltaY) > 2)) {
      dragState.dragged = true;
    }

    const nextLeft = Math.max(0, dragState.originLeft + deltaX);
    const nextTop = Math.max(0, dragState.originTop + deltaY);
    const proposedX = nextLeft / dragState.scale;
    const proposedY = topToBaseline(dragState.page, nextTop, dragState.size, dragState.scale);
    const snappedX = snapValue(proposedX, dragState.guides.vertical);
    const snappedY = snapValue(proposedY, dragState.guides.horizontal);
    const finalLeft = snappedX.value * dragState.scale;
    const finalTop = baselineToTop(dragState.page, snappedY.value, dragState.size, dragState.scale);

    dragState.node.style.left = `${finalLeft}px`;
    dragState.node.style.top = `${finalTop}px`;
    renderDragGuides(
      {
        vertical: snappedX.guide,
        horizontal: snappedY.guide
      },
      dragState.page,
      dragState.scale
    );
  };

  const handleUp = () => {
    if (!dragState) {
      return;
    }

    const { node: draggedNode, dragged, fieldKey, page: dragPage, scale: dragScale, size } = dragState;
    draggedNode.classList.remove("dragging");
    clearDragGuides();

    if (dragged) {
      const left = parseFloat(draggedNode.style.left || "0");
      const top = parseFloat(draggedNode.style.top || "0");
      const x = left / dragScale;
      const y = topToBaseline(dragPage, top, size, dragScale);
      applyPositionOverrideToEmployee(fieldKey, x, y);
      applyPositionOverrideToRenderModel(fieldKey, x, y);
      suppressClickUntil = Date.now() + 200;
      schedulePreviewRefresh();
      setStatus(`已调整 ${fieldKey} 的位置，右侧 PDF 正在自动刷新。`, "success");
    }

    dragState = null;
    window.removeEventListener("mousemove", handleMove);
    window.removeEventListener("mouseup", handleUp);
  };

  window.addEventListener("mousemove", handleMove);
  window.addEventListener("mouseup", handleUp);
}

function createTextElement(elementModel, page, scale) {
  const node = document.createElement("div");
  node.className = `card-element ${elementModel.editable ? "editable" : ""}`.trim();
  node.dataset.fieldKey = getElementKey(elementModel);
  node.style.left = `${elementModel.x * scale}px`;
  node.style.top = `${baselineToTop(page, elementModel.y, elementModel.size, scale)}px`;
  node.style.fontSize = `${elementModel.size * scale}px`;
  node.style.fontWeight = elementModel.fontWeight === "bold" ? "700" : "400";
  if (elementModel.fontFamily) {
    node.style.fontFamily = elementModel.fontFamily;
  }
  node.style.color = elementModel.color;
  node.textContent = elementModel.rawValue;

  if (elementModel.editable) {
    node.addEventListener("click", () => {
      if (Date.now() < suppressClickUntil) {
        return;
      }
      setSelectedField(getElementKey(elementModel));
      if (elementModel.fieldKey) {
        editorForm.querySelector(`[name="${CSS.escape(elementModel.fieldKey)}"]`)?.focus();
      }
    });
    node.addEventListener("dblclick", () => {
      if (Date.now() < suppressClickUntil) {
        return;
      }
      if (elementModel.textEditable) {
        beginInlineEdit(elementModel, node, page, scale);
      }
    });
    node.addEventListener("mousedown", (event) => {
      startDrag(event, elementModel, node, page, scale);
    });
  }

  return node;
}

function createMultilineElement(elementModel, page, scale) {
  const node = document.createElement("div");
  node.className = `card-element ${elementModel.editable ? "editable" : ""}`.trim();
  node.dataset.fieldKey = getElementKey(elementModel);
  node.style.left = `${elementModel.lines[0]?.x * scale || 0}px`;
  node.style.top = `${baselineToTop(page, elementModel.lines[0]?.y || 0, elementModel.size, scale)}px`;
  if (elementModel.maxWidth) {
    node.style.width = `${elementModel.maxWidth * scale}px`;
  }
  node.style.fontSize = `${elementModel.size * scale}px`;
  node.style.lineHeight = `${elementModel.lineGap * scale}px`;
  node.style.fontWeight = elementModel.fontWeight === "bold" ? "700" : "400";
  if (elementModel.fontFamily) {
    node.style.fontFamily = elementModel.fontFamily;
  }
  node.style.color = elementModel.color;

  elementModel.lines.forEach((line) => {
    const lineNode = document.createElement("span");
    lineNode.className = `card-line ${line.justify ? "justify" : ""}`.trim();
    lineNode.textContent = line.text;
    node.appendChild(lineNode);
  });

  if (elementModel.editable) {
    node.addEventListener("click", () => {
      if (Date.now() < suppressClickUntil) {
        return;
      }
      setSelectedField(getElementKey(elementModel));
      if (elementModel.fieldKey) {
        editorForm.querySelector(`[name="${CSS.escape(elementModel.fieldKey)}"]`)?.focus();
      }
    });
    node.addEventListener("dblclick", () => {
      if (Date.now() < suppressClickUntil) {
        return;
      }
      if (elementModel.textEditable) {
        beginInlineEdit(elementModel, node, page, scale);
      }
    });
    node.addEventListener("mousedown", (event) => {
      startDrag(event, elementModel, node, page, scale);
    });
  }

  return node;
}

function renderCardStage() {
  clearInlineEditor();
  cardStage.innerHTML = "";

  if (!currentRenderModel) {
    return;
  }

  const { page, decorations, elements } = currentRenderModel;
  cardStage.style.aspectRatio = `${page.width} / ${page.height}`;

  const scale = pageToStageScale(page);

  decorations.forEach((decoration) => {
    cardStage.appendChild(createDecoration(decoration, page, scale));
  });

  elements.forEach((elementModel) => {
    if (elementModel.type === "multiline-text") {
      cardStage.appendChild(createMultilineElement(elementModel, page, scale));
      return;
    }

    cardStage.appendChild(createTextElement(elementModel, page, scale));
  });

  setSelectedField(selectedFieldKey);
}

async function loadTemplates() {
  const response = await fetch("/api/templates");
  if (!response.ok) {
    throw new Error("模板列表加载失败。");
  }

  const data = await response.json();
  templateCatalog = [...(data.templates || [])].sort((left, right) => {
    const leftIndex = TEMPLATE_ORDER.indexOf(left.id);
    const rightIndex = TEMPLATE_ORDER.indexOf(right.id);
    const normalizedLeftIndex = leftIndex === -1 ? Number.MAX_SAFE_INTEGER : leftIndex;
    const normalizedRightIndex = rightIndex === -1 ? Number.MAX_SAFE_INTEGER : rightIndex;

    if (normalizedLeftIndex !== normalizedRightIndex) {
      return normalizedLeftIndex - normalizedRightIndex;
    }

    return getTemplateMeta(left).name.localeCompare(getTemplateMeta(right).name, "zh-CN");
  });
  templateSelect.innerHTML = "";

  templateCatalog.forEach((template) => {
    const templateMeta = getTemplateMeta(template);
    const option = document.createElement("option");
    option.value = template.id;
    option.textContent = templateMeta.name;
    option.dataset.columns = JSON.stringify(template.columns || []);
    option.dataset.fileNameField = template.fileNameField || "";
    templateSelect.appendChild(option);
  });

  if (templateCatalog.length) {
    const firstTemplate = templateCatalog[0];
    currentColumns = firstTemplate.columns || [];
    currentFileNameField = firstTemplate.fileNameField || "";
  }

  renderTemplateLibrary();
  renderTemplateShowcase(getCurrentTemplateMeta(), "悬停模板卡片可提前查看空白样式");
  renderEmployeeList();
  renderEditorForm();
  updateFooterSummary();
  updateFileStatus();
}

async function parseEmployees() {
  const files = Array.from(excelInput.files || []);
  if (!files.length) {
    throw new Error("请先选择至少一个 Excel 文件。");
  }

  const formData = new FormData();
  formData.append("templateId", templateSelect.value);
  files.forEach((file) => formData.append("excel", file));

  const response = await fetch("/api/parse", {
    method: "POST",
    body: formData
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(errorData.error || "解析 Excel 失败。");
  }

  return response.json();
}

async function fetchRenderModel() {
  const employee = getSelectedEmployee();
  if (!employee) {
    return null;
  }

  const response = await fetch("/api/render-model", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      templateId: templateSelect.value,
      employee
    })
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(errorData.error || "生成画布模型失败。");
  }

  return response.json();
}

async function refreshVisualEditor() {
  const employee = getSelectedEmployee();
  if (!employee) {
    return;
  }

  setStatus(`正在刷新 ${getEmployeeDisplayName(employee)} 的画布...`);

  try {
    currentRenderModel = await fetchRenderModel();
    renderCardStage();
    setStatus(`当前正在编辑：${getEmployeeDisplayName(employee)}`, "success");
  } catch (error) {
    renderTemplateShowcase(getCurrentTemplateMeta(), "画布刷新失败，请检查当前数据");
    setStatus(error.message, "error");
  }
}

async function refreshPdfPreview(options = {}) {
  const employee = getSelectedEmployee();
  if (!employee) {
    return;
  }

  cancelScheduledPreviewRefresh();
  refreshPreviewBtn.disabled = true;

  try {
    const response = await fetch("/api/preview", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        templateId: templateSelect.value,
        employee
      })
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      throw new Error(errorData.error || "生成预览失败。");
    }

    const blob = await response.blob();
    clearPreview();
    currentPreviewBytes = new Uint8Array(await blob.arrayBuffer());
    await renderPreviewImageFromBytes();
    if (!options.silent) {
      setStatus(`当前预览已更新：${getEmployeeDisplayName(employee)}`, "success");
    }
  } catch (error) {
    clearPreview();
    previewPlaceholder.textContent = "预览生成失败，请检查当前数据";
    setStatus(error.message, "error");
  } finally {
    updateActionAvailability();
  }
}

async function downloadAllPdfs() {
  if (!currentEmployees.length) {
    return;
  }

  downloadBtn.disabled = true;
  const isSingleEmployee = currentEmployees.length === 1;
  setLoadingState(true, "PDF 正在生成，请稍候...");
  setStatus(
    isSingleEmployee
      ? `正在导出 ${getEmployeeDisplayName(currentEmployees[0])} 的 PDF...`
      : `正在打包 ${currentEmployees.length} 份 PDF，请稍候...`
  );

  try {
    const response = await fetch(isSingleEmployee ? "/api/preview" : "/api/download-all", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        templateId: templateSelect.value,
        ...(isSingleEmployee
          ? { employee: currentEmployees[0] }
          : { employees: currentEmployees })
      })
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      throw new Error(errorData.error || (isSingleEmployee ? "导出 PDF 失败。" : "批量导出失败。"));
    }

    const blob = await response.blob();
    const fileName = decodeContentDispositionFileName(response.headers.get("Content-Disposition"));
    triggerBlobDownload(blob, fileName);
    setStatus(
      isSingleEmployee
        ? `已导出 ${fileName || `${getEmployeeDisplayName(currentEmployees[0])}.pdf`}`
        : `已导出 ${currentEmployees.length} 位员工的 PDF 压缩包。`,
      "success"
    );
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    setLoadingState(false);
    updateActionAvailability();
  }
}

async function sendCurrentCardToFeishu() {
  const employee = getSelectedEmployee();
  if (!employee) {
    return;
  }

  sendFeishuBtn.disabled = true;
  setStatus(`正在发送 ${getEmployeeDisplayName(employee)} 的名片到飞书...`);

  try {
    const response = await fetch("/api/feishu/send-current", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        templateId: templateSelect.value,
        employee
      })
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      throw new Error(errorData.error || "发送飞书确认消息失败。");
    }

    const data = await response.json();
    setStatus(`已发送到飞书：${data.employeeName}，请对方确认名片内容是否有问题。`, "success");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    updateActionAvailability();
  }
}

templateSelect.addEventListener("change", () => {
  const selected = templateSelect.selectedOptions[0];
  currentColumns = JSON.parse(selected?.dataset.columns || "[]");
  currentFileNameField = selected?.dataset.fileNameField || "";
  renderTemplateLibrary();
  clearWorkspace();
  setStatus("");
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  submitBtn.disabled = true;
  clearWorkspace();
  updateFileStatus();
  setStatus("正在解析 Excel 并加载员工数据...");

  try {
    const data = await parseEmployees();
    currentEmployees = data.employees || [];
    currentColumns = data.columns || [];
    currentFileNameField = data.fileNameField || "";
    selectedEmployeeId = currentEmployees[0]?.id || "";

    renderEmployeeList();
    renderEditorForm();
    updateActionAvailability();
    updateFileStatus();
    await refreshVisualEditor();
    await refreshPdfPreview();
  } catch (error) {
    clearWorkspace();
    setStatus(error.message, "error");
  } finally {
    submitBtn.disabled = false;
  }
});

refreshPreviewBtn.addEventListener("click", async () => {
  await refreshPdfPreview();
});

sendFeishuBtn.addEventListener("click", async () => {
  await sendCurrentCardToFeishu();
});

downloadBtn.addEventListener("click", async () => {
  await downloadAllPdfs();
});

excelInput.addEventListener("change", () => {
  updateFileStatus();
});

document.querySelectorAll(".tool-button").forEach((button) => {
  button.addEventListener("click", () => {
    setStatus(button.dataset.toolMessage || "功能即将开放", "success");
  });
});

cardStageWrap.addEventListener("click", (event) => {
  if (event.target === cardStageWrap || event.target === cardStage) {
    setSelectedField("");
  }
});

window.addEventListener("resize", () => {
  if (currentRenderModel) {
    renderCardStage();
  }

  if (currentPreviewBytes?.length) {
    clearTimeout(previewResizeTimer);
    previewResizeTimer = window.setTimeout(() => {
      renderPreviewImageFromBytes().catch((error) => {
        setStatus(error.message, "error");
      });
    }, 120);
  }
});

window.addEventListener("keydown", async (event) => {
  if (!selectedFieldKey || !currentRenderModel || shouldIgnoreKeyboardShortcut(event.target) || inlineEditorEl) {
    return;
  }

  const directionMap = {
    ArrowLeft: { dx: -1, dy: 0 },
    ArrowRight: { dx: 1, dy: 0 },
    ArrowUp: { dx: 0, dy: 1 },
    ArrowDown: { dx: 0, dy: -1 }
  };

  const movement = directionMap[event.key];
  if (!movement) {
    return;
  }

  event.preventDefault();
  const step = event.shiftKey ? 2 : 0.5;
  const elementModel = currentRenderModel.elements.find((item) => getElementKey(item) === selectedFieldKey);
  if (!elementModel) {
    return;
  }

  const origin = getElementGuideOrigin(elementModel);
  const nextX = origin.x + movement.dx * step;
  const nextY = origin.y + movement.dy * step;
  const positionKey = getElementKey(elementModel);
  applyPositionOverrideToEmployee(positionKey, nextX, nextY);
  applyPositionOverrideToRenderModel(positionKey, nextX, nextY);
  renderCardStage();
  setSelectedField(selectedFieldKey);
  schedulePreviewRefresh();
  setStatus(`已微调 ${selectedFieldKey} 的位置（${event.shiftKey ? "快速" : "精细"}模式），右侧 PDF 正在自动刷新。`, "success");
});

loadTemplates().catch((error) => {
  setStatus(error.message, "error");
});





