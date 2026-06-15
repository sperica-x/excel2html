const pasteZone = document.querySelector("#pasteZone");
const convertButton = document.querySelector("#convertButton");
const copyButton = document.querySelector("#copyButton");
const clearButton = document.querySelector("#clearButton");
const htmlOutput = document.querySelector("#htmlOutput");
const preview = document.querySelector("#preview");
const statusNode = document.querySelector("#status");
const liveVisitorsNode = document.querySelector("#liveVisitors");
const totalVisitorsNode = document.querySelector("#totalVisitors");
const undoButton = document.querySelector("#undoButton");
const redoButton = document.querySelector("#redoButton");
const addParagraphButton = document.querySelector("#addParagraphButton");
const fontFamilySelect = document.querySelector("#fontFamilySelect");
const fontSizeSelect = document.querySelector("#fontSizeSelect");
const boldButton = document.querySelector("#boldButton");
const italicButton = document.querySelector("#italicButton");
const underlineButton = document.querySelector("#underlineButton");
const alignLeftButton = document.querySelector("#alignLeftButton");
const alignCenterButton = document.querySelector("#alignCenterButton");
const alignRightButton = document.querySelector("#alignRightButton");
const lineHeightSelect = document.querySelector("#lineHeightSelect");
const backgroundColorInput = document.querySelector("#backgroundColorInput");
const textColorInput = document.querySelector("#textColorInput");
const borderColorInput = document.querySelector("#borderColorInput");
const borderWidthSelect = document.querySelector("#borderWidthSelect");
const applyBorderButton = document.querySelector("#applyBorderButton");
const clearFormatButton = document.querySelector("#clearFormatButton");
const addRowButton = document.querySelector("#addRowButton");
const deleteRowButton = document.querySelector("#deleteRowButton");
const addColumnButton = document.querySelector("#addColumnButton");
const deleteColumnButton = document.querySelector("#deleteColumnButton");
const mergeCellsButton = document.querySelector("#mergeCellsButton");
const splitCellButton = document.querySelector("#splitCellButton");

let lastTableHtml = "";
let selectedCells = [];
let selectionAnchor = null;
let selectionAnchorPoint = null;
let selectionAnchorInfo = null;
let selectionMode = "range";
let isSelectingCells = false;
let resizeState = null;
let historyStack = [];
let historyIndex = -1;
let isRestoringHistory = false;
let pasteBlocks = [];
let heartbeatTimer = null;
let statsSessionId = "";
let statsVisitorId = "";
const HEARTBEAT_MS = 45000;
const VISITOR_STORAGE_KEY = "board-insight-n-visitor-id";
const SESSION_STORAGE_KEY = "board-insight-n-session-id";
const EDITOR_BUTTONS = [
  undoButton,
  redoButton,
  addParagraphButton,
  fontFamilySelect,
  fontSizeSelect,
  boldButton,
  italicButton,
  underlineButton,
  alignLeftButton,
  alignCenterButton,
  alignRightButton,
  lineHeightSelect,
  backgroundColorInput,
  textColorInput,
  borderColorInput,
  borderWidthSelect,
  applyBorderButton,
  clearFormatButton,
  addRowButton,
  deleteRowButton,
  addColumnButton,
  deleteColumnButton,
  mergeCellsButton,
  splitCellButton,
];

function setStatus(message, isError = false) {
  statusNode.textContent = message;
  statusNode.style.color = isError ? "#a11d1d" : "";
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function cleanupCell(node, tagName, attrs = "") {
  const clone = node.cloneNode(true);
  clone.querySelectorAll("br").forEach((br) => br.replaceWith("\n"));

  let content = clone.textContent ?? "";
  content = content.replace(/\u00a0/g, " ").replace(/\r\n/g, "\n").trim();

  return `<${tagName}${attrs}>${escapeHtml(content).replaceAll("\n", "<br />")}</${tagName}>`;
}

function isTransparent(color) {
  return color === "rgba(0, 0, 0, 0)" || color === "transparent";
}

function shouldKeepColor(color) {
  return color && color !== "rgb(0, 0, 0)" && color !== "canvastext";
}

function buildStyleString(styleMap) {
  const entries = Object.entries(styleMap).filter(([, value]) => value);

  if (!entries.length) {
    return "";
  }

  return entries.map(([key, value]) => `${key}:${value}`).join(";");
}

function readCellStyle(cell) {
  const computed = window.getComputedStyle(cell);
  const styleMap = {};

  ["top", "right", "bottom", "left"].forEach((side) => {
    const width = computed.getPropertyValue(`border-${side}-width`);
    const borderStyle = computed.getPropertyValue(`border-${side}-style`);
    const color = computed.getPropertyValue(`border-${side}-color`);

    if (width !== "0px" && borderStyle !== "none") {
      styleMap[`border-${side}`] = `${width} ${borderStyle} ${color}`;
    }
  });

  if (!isTransparent(computed.backgroundColor)) {
    styleMap["background-color"] = computed.backgroundColor;
  }

  if (shouldKeepColor(computed.color)) {
    styleMap.color = computed.color;
  }

  if (computed.fontWeight && computed.fontWeight !== "400") {
    styleMap["font-weight"] = computed.fontWeight;
  }

  if (computed.fontStyle && computed.fontStyle !== "normal") {
    styleMap["font-style"] = computed.fontStyle;
  }

  if (computed.textAlign && computed.textAlign !== "start") {
    styleMap["text-align"] = computed.textAlign;
  }

  if (computed.verticalAlign && computed.verticalAlign !== "baseline") {
    styleMap["vertical-align"] = computed.verticalAlign;
  }

  if (computed.whiteSpace && computed.whiteSpace !== "normal") {
    styleMap["white-space"] = computed.whiteSpace;
  }

  ["top", "right", "bottom", "left"].forEach((side) => {
    const value = computed.getPropertyValue(`padding-${side}`);

    if (value && value !== "0px") {
      styleMap[`padding-${side}`] = value;
    }
  });

  if (computed.width && computed.width !== "auto" && computed.width !== "0px") {
    styleMap.width = computed.width;
  }

  if (computed.height && computed.height !== "auto" && computed.height !== "0px") {
    styleMap.height = computed.height;
  }

  return buildStyleString(styleMap);
}

function readTableStyle(table) {
  const computed = window.getComputedStyle(table);
  const styleMap = {
    "border-collapse": computed.borderCollapse || "collapse",
  };

  if (computed.borderSpacing && computed.borderSpacing !== "0px") {
    styleMap["border-spacing"] = computed.borderSpacing;
  }

  if (computed.width && computed.width !== "auto" && computed.width !== "0px") {
    styleMap.width = computed.width;
  }

  return buildStyleString(styleMap);
}

function createMeasurementRoot(rawHtml) {
  const root = document.createElement("div");
  root.style.position = "fixed";
  root.style.left = "-99999px";
  root.style.top = "0";
  root.style.visibility = "hidden";
  root.style.pointerEvents = "none";
  root.style.background = "white";
  root.innerHTML = rawHtml;
  document.body.append(root);
  return root;
}

function normalizeSingleTable(sourceTable) {
  const rows = [...sourceTable.querySelectorAll("tr")];

  if (!rows.length) {
    return "";
  }

  const normalizedRows = rows
    .map((row) => {
      const cells = [...row.children].filter((cell) => /^(TD|TH)$/i.test(cell.tagName));

      if (!cells.length) {
        return "";
      }

      const normalizedCells = cells
        .map((cell) => {
          const tagName = cell.tagName.toLowerCase() === "th" ? "th" : "td";
          const attrs = [];

          if (cell.colSpan > 1) {
            attrs.push(` colspan="${cell.colSpan}"`);
          }

          if (cell.rowSpan > 1) {
            attrs.push(` rowspan="${cell.rowSpan}"`);
          }

          const styleText = readCellStyle(cell);

          if (styleText) {
            attrs.push(` style="${escapeHtml(styleText)}"`);
          }

          return cleanupCell(cell, tagName, attrs.join(""));
        })
        .join("");

      return `<tr>${normalizedCells}</tr>`;
    })
    .filter(Boolean)
    .join("");

  if (!normalizedRows) {
    return "";
  }

  const tableStyle = readTableStyle(sourceTable);
  const tableAttrs = [
    'border="1"',
    'cellpadding="0"',
    'cellspacing="0"',
  ];

  if (tableStyle) {
    tableAttrs.push(` style="${escapeHtml(tableStyle)}"`);
  }

  return `<table ${tableAttrs.join(" ")}>${normalizedRows}</table>`;
}

const BLOCK_LEVEL_TAGS =
  /^(P|DIV|LI|UL|OL|H[1-6]|TR|TABLE|SECTION|ARTICLE|HEADER|FOOTER|BLOCKQUOTE|PRE|ADDRESS)$/;

// 표가 아닌 노드에서 줄바꿈 구조를 유지한 텍스트를 추출한다.
function extractTextWithBreaks(node) {
  if (node.nodeType === Node.TEXT_NODE) {
    return (node.textContent ?? "").replace(/ /g, " ");
  }

  if (node.nodeType !== Node.ELEMENT_NODE) {
    return "";
  }

  const tag = node.tagName.toUpperCase();

  if (tag === "BR") {
    return "\n";
  }

  if (tag === "TABLE") {
    return "";
  }

  let text = "";
  node.childNodes.forEach((child) => {
    text += extractTextWithBreaks(child);
  });

  // 블록 요소는 문단 경계로 취급해 빈 줄을 넣는다.
  if (BLOCK_LEVEL_TAGS.test(tag)) {
    text += "\n\n";
  }

  return text;
}

function paragraphsFromNodes(nodes) {
  if (!nodes.length) {
    return "";
  }

  const text = nodes
    .map((node) => extractTextWithBreaks(node))
    .join("");

  return normalizePlainTextParagraphs(text);
}

// 붙여넣은 HTML을 문서 순서대로 훑어 텍스트와 표를 모두 보존한다.
function collectHtmlBlocks(container, parts) {
  let pendingTextNodes = [];

  const flushText = () => {
    const html = paragraphsFromNodes(pendingTextNodes);

    if (html) {
      parts.push(html);
    }

    pendingTextNodes = [];
  };

  [...container.childNodes].forEach((child) => {
    if (child.nodeType === Node.ELEMENT_NODE && child.tagName.toUpperCase() === "TABLE") {
      flushText();
      const table = normalizeSingleTable(child);

      if (table) {
        parts.push(table);
      }

      return;
    }

    // 표를 품고 있는 컨테이너는 순서를 지키기 위해 안으로 들어가 처리한다.
    if (
      child.nodeType === Node.ELEMENT_NODE &&
      typeof child.querySelector === "function" &&
      child.querySelector("table")
    ) {
      flushText();
      collectHtmlBlocks(child, parts);
      return;
    }

    pendingTextNodes.push(child);
  });

  flushText();
}

function normalizeHtmlContent(rawHtml) {
  const measurementRoot = createMeasurementRoot(rawHtml);
  const parts = [];

  collectHtmlBlocks(measurementRoot, parts);
  measurementRoot.remove();

  return parts.filter(Boolean).join("");
}

function normalizePlainTextTable(rawText) {
  const lines = rawText
    .replace(/\r\n/g, "\n")
    .split("\n")
    .filter((line) => line.trim() !== "");

  if (!lines.length) {
    return "";
  }

  const rows = lines.map((line) => line.split("\t"));
  const tableRows = rows
    .map((cells) => {
      const inner = cells
        .map((cell) => `<td>${escapeHtml(cell.trim())}</td>`)
        .join("");
      return `<tr>${inner}</tr>`;
    })
    .join("");

  return `<table>${tableRows}</table>`;
}

function normalizePlainTextParagraphs(rawText) {
  const paragraphs = rawText
    .replace(/\r\n/g, "\n")
    .split(/\n{2,}/)
    .map((paragraph) => paragraph.trim())
    .filter(Boolean);

  if (!paragraphs.length) {
    return "";
  }

  return paragraphs
    .map((paragraph) => `<p>${escapeHtml(paragraph).replaceAll("\n", "<br />")}</p>`)
    .join("");
}

function normalizePlainText(rawText) {
  const lines = rawText.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  const segments = [];
  let current = null;

  // 탭이 있는 줄(표)과 없는 줄(텍스트)을 연속 구간으로 묶어 둘 다 살린다.
  lines.forEach((line) => {
    const kind = line.includes("\t") ? "table" : "text";

    if (!current || current.kind !== kind) {
      current = { kind, lines: [] };
      segments.push(current);
    }

    current.lines.push(line);
  });

  return segments
    .map((segment) => {
      const chunk = segment.lines.join("\n");

      return segment.kind === "table"
        ? normalizePlainTextTable(chunk)
        : normalizePlainTextParagraphs(chunk);
    })
    .filter(Boolean)
    .join("");
}

function resetPasteZoneIfPlaceholder() {
  if (!pasteBlocks.length) {
    pasteZone.innerHTML = "";
  }
}

function renderPasteBlocks() {
  if (!pasteBlocks.length) {
    pasteZone.innerHTML = "여기에 Ctrl+V로 엑셀 표를 붙여넣으세요.";
    return;
  }

  pasteZone.innerHTML = pasteBlocks
    .map((block, index) => {
      const label = block.kind === "html"
        ? "엑셀 HTML"
        : block.text.includes("\t")
          ? "탭 구분 표"
          : "일반 텍스트";
      const content = block.kind === "html"
        ? block.html
        : `<pre>${escapeHtml(block.text)}</pre>`;

      return `<section class="paste-block" data-index="${index}"><p class="paste-block-label">${index + 1}. ${label}</p>${content}</section>`;
    })
    .join("");
}

function addPasteBlock(block) {
  resetPasteZoneIfPlaceholder();
  pasteBlocks.push(block);
  renderPasteBlocks();
}

function buildDocumentFromPasteBlocks() {
  if (pasteBlocks.length) {
    return pasteBlocks
      .map((block) => {
        if (block.kind === "html") {
          return normalizeHtmlContent(block.html);
        }

        return normalizePlainText(block.text);
      })
      .filter(Boolean)
      .join("<p><br /></p>");
  }

  const liveHtml = pasteZone.innerHTML;
  const liveText = pasteZone.innerText;

  if (liveHtml.includes("<table")) {
    return normalizeHtmlContent(liveHtml);
  }

  return normalizePlainText(liveText);
}

function enableEditorControls(enabled) {
  EDITOR_BUTTONS.forEach((control) => {
    control.disabled = !enabled;
  });

  updateHistoryButtons();
}

function getCellFromTarget(target) {
  return target.closest?.("td, th") || null;
}

function clearCellSelection() {
  selectedCells.forEach((cell) => cell.classList.remove("selected-cell"));
  selectedCells = [];
}

function setSelectedCells(cells) {
  clearCellSelection();
  selectedCells = cells.filter(Boolean);
  selectedCells.forEach((cell) => cell.classList.add("selected-cell"));
  mergeCellsButton.disabled = selectedCells.length < 2;
}

function getCellPosition(cell) {
  return {
    rowIndex: cell.parentElement.rowIndex,
    cellIndex: cell.cellIndex,
  };
}

function buildTableGrid(table) {
  const grid = [];
  const cellInfo = new Map();

  [...table.rows].forEach((row, rowIndex) => {
    grid[rowIndex] = grid[rowIndex] || [];
    let columnIndex = 0;

    [...row.cells].forEach((cell) => {
      while (grid[rowIndex][columnIndex]) {
        columnIndex += 1;
      }

      const rowSpan = Math.max(1, cell.rowSpan || 1);
      const colSpan = Math.max(1, cell.colSpan || 1);
      const info = {
        rowIndex,
        columnIndex,
        rowSpan,
        colSpan,
      };

      cellInfo.set(cell, info);

      for (let rowOffset = 0; rowOffset < rowSpan; rowOffset += 1) {
        const targetRow = rowIndex + rowOffset;
        grid[targetRow] = grid[targetRow] || [];

        for (let colOffset = 0; colOffset < colSpan; colOffset += 1) {
          grid[targetRow][columnIndex + colOffset] = cell;
        }
      }

      columnIndex += colSpan;
    });
  });

  return {
    grid,
    cellInfo,
  };
}

function uniqueCells(cells) {
  return [...new Set(cells.filter(Boolean))];
}

function selectColumnRange(fromCell, toCell) {
  const table = fromCell.closest("table");

  if (!table || table !== toCell.closest("table")) {
    setSelectedCells([toCell]);
    return;
  }

  const { grid, cellInfo } = buildTableGrid(table);
  const from = selectionAnchorInfo || cellInfo.get(fromCell);
  const to = cellInfo.get(toCell);

  if (!from || !to) {
    setSelectedCells([toCell]);
    return;
  }

  const minRow = Math.min(from.rowIndex, to.rowIndex);
  const maxRow = Math.max(from.rowIndex, to.rowIndex);
  const columnIndex = from.columnIndex;
  const cells = [];

  for (let rowIndex = minRow; rowIndex <= maxRow; rowIndex += 1) {
    const row = grid[rowIndex] || [];
    cells.push(row[columnIndex]);
  }

  setSelectedCells(uniqueCells(cells));
}

function selectCellRange(fromCell, toCell) {
  const table = fromCell.closest("table");

  if (!table || table !== toCell.closest("table")) {
    setSelectedCells([toCell]);
    return;
  }

  const { grid, cellInfo } = buildTableGrid(table);
  const from = getCellPosition(fromCell);
  const to = getCellPosition(toCell);
  const fromInfo = selectionAnchorInfo || cellInfo.get(fromCell);
  const toInfo = cellInfo.get(toCell);

  if (!fromInfo || !toInfo) {
    setSelectedCells([toCell]);
    return;
  }

  const minRow = Math.min(from.rowIndex, to.rowIndex, fromInfo.rowIndex, toInfo.rowIndex);
  const maxRow = Math.max(from.rowIndex, to.rowIndex, fromInfo.rowIndex, toInfo.rowIndex);
  const minCell = Math.min(fromInfo.columnIndex, toInfo.columnIndex);
  const maxCell = Math.max(
    fromInfo.columnIndex + fromInfo.colSpan - 1,
    toInfo.columnIndex + toInfo.colSpan - 1,
  );
  const cells = [];

  for (let rowIndex = minRow; rowIndex <= maxRow; rowIndex += 1) {
    const row = grid[rowIndex] || [];

    for (let cellIndex = minCell; cellIndex <= maxCell; cellIndex += 1) {
      cells.push(row[cellIndex]);
    }
  }

  setSelectedCells(uniqueCells(cells));
}

function updateDragSelection(cell, event) {
  if (!selectionAnchor) {
    return;
  }

  if (selectionMode === "range" && selectionAnchorPoint) {
    const deltaX = Math.abs(event.clientX - selectionAnchorPoint.x);
    const deltaY = Math.abs(event.clientY - selectionAnchorPoint.y);

    if (deltaY > 8 && deltaY > deltaX * 1.4) {
      selectionMode = "column";
    }
  }

  if (selectionMode === "column") {
    selectColumnRange(selectionAnchor, cell);
    return;
  }

  selectCellRange(selectionAnchor, cell);
}

function getActiveEditable() {
  return document.activeElement?.closest?.("#preview td, #preview th, #preview p") || null;
}

function getPrimaryCell() {
  const activeCell = document.activeElement?.closest?.("#preview td, #preview th");
  return selectedCells[0] || activeCell || preview.querySelector("td, th");
}

function getTargetsForFormatting() {
  const active = getActiveEditable();

  if (selectedCells.length) {
    return selectedCells;
  }

  return active ? [active] : [];
}

function cleanEditorHtml() {
  const clone = preview.cloneNode(true);

  clone.querySelectorAll(".resize-handle, .row-resize-handle").forEach((node) => node.remove());
  clone.querySelectorAll(".selected-cell").forEach((node) => {
    node.classList.remove("selected-cell");
  });
  clone.querySelectorAll("[contenteditable]").forEach((node) => {
    node.removeAttribute("contenteditable");
  });
  clone.querySelectorAll("[data-editor-ready]").forEach((node) => {
    node.removeAttribute("data-editor-ready");
  });
  clone.querySelectorAll("td, th").forEach((cell) => {
    cell.classList.remove("selected-cell");

    if (!cell.getAttribute("class")) {
      cell.removeAttribute("class");
    }
  });

  return clone.innerHTML.trim();
}

function updateHistoryButtons() {
  if (preview.classList.contains("empty")) {
    undoButton.disabled = true;
    redoButton.disabled = true;
    return;
  }

  undoButton.disabled = historyIndex <= 0;
  redoButton.disabled = historyIndex < 0 || historyIndex >= historyStack.length - 1;
}

function captureHistory() {
  if (isRestoringHistory || preview.classList.contains("empty")) {
    return;
  }

  const snapshot = cleanEditorHtml();

  if (historyStack[historyIndex] === snapshot) {
    updateHistoryButtons();
    return;
  }

  historyStack = historyStack.slice(0, historyIndex + 1);
  historyStack.push(snapshot);

  if (historyStack.length > 50) {
    historyStack.shift();
  }

  historyIndex = historyStack.length - 1;
  updateHistoryButtons();
}

function restoreHistory(index) {
  if (index < 0 || index >= historyStack.length) {
    return;
  }

  isRestoringHistory = true;
  historyIndex = index;
  preview.innerHTML = historyStack[historyIndex];
  preview.classList.remove("empty");
  clearCellSelection();
  makeEditorReady(false);
  syncHtmlOutput();
  isRestoringHistory = false;
  updateHistoryButtons();
}

function syncHtmlOutput() {
  if (preview.classList.contains("empty")) {
    lastTableHtml = "";
    htmlOutput.value = "";
    return;
  }

  lastTableHtml = cleanEditorHtml();
  htmlOutput.value = lastTableHtml;
}

function makeEditorReady(recordHistory = false) {
  preview.querySelectorAll("td, th").forEach((cell) => {
    cell.contentEditable = "true";
  });

  preview.querySelectorAll("p").forEach((paragraph) => {
    paragraph.contentEditable = "true";
  });

  preview.querySelectorAll("table").forEach((table) => {
    table.style.borderCollapse = table.style.borderCollapse || "collapse";
    table.style.minWidth = "auto";
    table.style.width = table.style.width || "auto";
    table.querySelectorAll(".resize-handle, .row-resize-handle").forEach((handle) => handle.remove());

    const firstRow = table.rows[0];

    if (!firstRow) {
      return;
    }

    [...firstRow.cells].forEach((cell) => {
      const handle = document.createElement("span");
      handle.className = "resize-handle";
      handle.contentEditable = "false";
      cell.append(handle);
    });

    [...table.rows].forEach((row) => {
      const firstCell = row.cells[0];

      if (!firstCell) {
        return;
      }

      const handle = document.createElement("span");
      handle.className = "row-resize-handle";
      handle.contentEditable = "false";
      firstCell.append(handle);
    });
  });

  enableEditorControls(true);
  syncHtmlOutput();

  if (recordHistory) {
    captureHistory();
  }
}

function applyToTargets(callback) {
  const targets = getTargetsForFormatting();

  targets.forEach(callback);
  syncHtmlOutput();
  captureHistory();
}

function applyAlignment(value) {
  applyToTargets((target) => {
    target.style.textAlign = value;
  });
}

function insertParagraph() {
  const paragraph = document.createElement("p");
  paragraph.contentEditable = "true";
  paragraph.textContent = "새 문단";

  const active = getActiveEditable();
  const activeTable = active?.closest("table");
  const insertionTarget = activeTable || active;

  if (insertionTarget?.parentNode === preview) {
    insertionTarget.after(paragraph);
  } else {
    preview.append(paragraph);
  }

  paragraph.focus();
  syncHtmlOutput();
  captureHistory();
}

function cloneRow(row) {
  const clone = row.cloneNode(true);

  clone.querySelectorAll(".resize-handle").forEach((handle) => handle.remove());
  [...clone.cells].forEach((cell) => {
    cell.textContent = "";
    cell.removeAttribute("rowspan");
    cell.removeAttribute("colspan");
    cell.contentEditable = "true";
  });

  return clone;
}

function addRowAfterSelection() {
  const cell = getPrimaryCell();
  const row = cell?.parentElement;

  if (!row) {
    return;
  }

  const newRow = cloneRow(row);
  row.after(newRow);
  makeEditorReady();
  setSelectedCells([...newRow.cells]);
  captureHistory();
}

function deleteSelectedRows() {
  const rows = [...new Set(selectedCells.map((cell) => cell.parentElement))];
  const targetRows = rows.length ? rows : [getPrimaryCell()?.parentElement].filter(Boolean);

  targetRows.forEach((row) => {
    const table = row.closest("table");

    if (table.rows.length > 1) {
      row.remove();
    }
  });

  clearCellSelection();
  makeEditorReady();
  captureHistory();
}

function addColumnAfterSelection() {
  const cell = getPrimaryCell();
  const table = cell?.closest("table");

  if (!table) {
    return;
  }

  const insertAfter = cell.cellIndex;

  [...table.rows].forEach((row) => {
    const source = row.cells[insertAfter] || row.cells[row.cells.length - 1];
    const newCell = source.cloneNode(false);
    newCell.textContent = "";
    newCell.contentEditable = "true";
    newCell.removeAttribute("rowspan");
    newCell.removeAttribute("colspan");

    if (source.nextSibling) {
      row.insertBefore(newCell, source.nextSibling);
    } else {
      row.append(newCell);
    }
  });

  makeEditorReady();
  captureHistory();
}

function deleteSelectedColumn() {
  const cell = getPrimaryCell();
  const table = cell?.closest("table");

  if (!table) {
    return;
  }

  const columnIndex = cell.cellIndex;

  [...table.rows].forEach((row) => {
    if (row.cells.length > 1 && row.cells[columnIndex]) {
      row.cells[columnIndex].remove();
    }
  });

  clearCellSelection();
  makeEditorReady();
  captureHistory();
}

function mergeSelectedCells() {
  if (selectedCells.length < 2) {
    return;
  }

  const table = selectedCells[0].closest("table");

  if (!table || selectedCells.some((cell) => cell.closest("table") !== table)) {
    return;
  }

  const positions = selectedCells.map(getCellPosition);
  const minRow = Math.min(...positions.map((position) => position.rowIndex));
  const maxRow = Math.max(...positions.map((position) => position.rowIndex));
  const minCell = Math.min(...positions.map((position) => position.cellIndex));
  const maxCell = Math.max(...positions.map((position) => position.cellIndex));
  const master = table.rows[minRow]?.cells[minCell];

  if (!master) {
    return;
  }

  const mergedText = selectedCells
    .map((cell) => cell.innerText.trim())
    .filter(Boolean)
    .join("\n");

  selectedCells.forEach((cell) => {
    if (cell !== master) {
      cell.remove();
    }
  });

  master.rowSpan = maxRow - minRow + 1;
  master.colSpan = maxCell - minCell + 1;
  master.textContent = mergedText;
  makeEditorReady();
  setSelectedCells([master]);
  captureHistory();
}

function splitSelectedCell() {
  const cell = getPrimaryCell();

  if (!cell) {
    return;
  }

  cell.removeAttribute("rowspan");
  cell.removeAttribute("colspan");
  makeEditorReady();
  setSelectedCells([cell]);
  captureHistory();
}

function setColumnWidth(table, columnIndex, width) {
  const normalizedWidth = Math.max(24, Math.round(width));

  [...table.rows].forEach((row) => {
    const cell = row.cells[columnIndex];

    if (cell) {
      cell.style.width = `${normalizedWidth}px`;
      cell.style.minWidth = "0";
      cell.style.maxWidth = `${normalizedWidth}px`;
      cell.style.overflowWrap = "break-word";
    }
  });

  table.style.width = "auto";
  table.style.minWidth = "auto";
  syncHtmlOutput();
}

function setRowHeight(row, height) {
  [...row.cells].forEach((cell) => {
    cell.style.height = `${Math.max(24, Math.round(height))}px`;
  });

  syncHtmlOutput();
}

function clearFormatting() {
  applyToTargets((target) => {
    const content = target.innerHTML;
    target.removeAttribute("style");
    target.innerHTML = content;
  });
}

function readPastedTable() {
  return buildDocumentFromPasteBlocks();
}

function renderTable(tableHtml) {
  preview.innerHTML = tableHtml;
  preview.classList.remove("empty");
  copyButton.disabled = false;
  historyStack = [];
  historyIndex = -1;
  makeEditorReady(true);
  setStatus("문서 편집 영역에서 표를 수정한 뒤 HTML 코드를 복사하세요.");
}

function resetOutput(message) {
  lastTableHtml = "";
  htmlOutput.value = "";
  preview.textContent = "변환 결과가 여기에 표시됩니다.";
  preview.classList.add("empty");
  copyButton.disabled = true;
  enableEditorControls(false);
  clearCellSelection();
  historyStack = [];
  historyIndex = -1;
  updateHistoryButtons();
  setStatus(message);
}

pasteZone.addEventListener("paste", (event) => {
  const html = event.clipboardData?.getData("text/html") ?? "";
  const text = event.clipboardData?.getData("text/plain") ?? "";

  if (html) {
    event.preventDefault();
    addPasteBlock({
      kind: "html",
      html,
      text,
    });
    setStatus(`붙여넣기 ${pasteBlocks.length}개를 모았습니다. 계속 붙여넣거나 변환하세요.`);
    return;
  }

  if (text) {
    event.preventDefault();
    addPasteBlock({
      kind: "text",
      text,
    });
    setStatus(`붙여넣기 ${pasteBlocks.length}개를 모았습니다. 계속 붙여넣거나 변환하세요.`);
  }
});

convertButton.addEventListener("click", () => {
  const tableHtml = readPastedTable();

  if (!tableHtml) {
    resetOutput("붙여넣은 내용에서 표나 텍스트를 찾지 못했습니다.");
    statusNode.style.color = "#a11d1d";
    return;
  }

  renderTable(tableHtml);
});

copyButton.addEventListener("click", async () => {
  syncHtmlOutput();

  if (!lastTableHtml) {
    return;
  }

  try {
    await navigator.clipboard.writeText(lastTableHtml);
    setStatus("HTML 코드를 클립보드에 복사했습니다.");
  } catch (error) {
    setStatus("브라우저 복사 권한이 없어 수동 복사가 필요합니다.", true);
  }
});

clearButton.addEventListener("click", () => {
  pasteBlocks = [];
  renderPasteBlocks();
  resetOutput("초기화했습니다. 새로 붙여넣으세요.");
});

preview.addEventListener("mousedown", (event) => {
  const resizeHandle = event.target.closest(".resize-handle");
  const rowResizeHandle = event.target.closest(".row-resize-handle");

  if (resizeHandle) {
    const cell = resizeHandle.closest("td, th");
    resizeState = {
      type: "column",
      table: cell.closest("table"),
      columnIndex: cell.cellIndex,
      startX: event.clientX,
      startWidth: cell.getBoundingClientRect().width,
    };
    event.preventDefault();
    return;
  }

  if (rowResizeHandle) {
    const cell = rowResizeHandle.closest("td, th");
    const row = cell.parentElement;
    resizeState = {
      type: "row",
      row,
      startY: event.clientY,
      startHeight: row.getBoundingClientRect().height,
    };
    event.preventDefault();
    return;
  }

  const cell = getCellFromTarget(event.target);

  if (!cell || !preview.contains(cell)) {
    return;
  }

  selectionAnchor = cell;
  selectionAnchorPoint = {
    x: event.clientX,
    y: event.clientY,
  };
  selectionAnchorInfo = buildTableGrid(cell.closest("table")).cellInfo.get(cell);
  selectionMode = "range";
  isSelectingCells = true;
  selectCellRange(selectionAnchor, cell);
  event.preventDefault();
});

preview.addEventListener("mouseover", (event) => {
  if (!isSelectingCells || !selectionAnchor) {
    return;
  }

  const cell = getCellFromTarget(event.target);

  if (cell) {
    updateDragSelection(cell, event);
  }
});

preview.addEventListener("dblclick", (event) => {
  const editable = event.target.closest("td, th, p");

  if (editable) {
    editable.focus();
  }
});

preview.addEventListener("input", () => {
  syncHtmlOutput();
  captureHistory();
});

preview.addEventListener("keydown", (event) => {
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "z" && !event.shiftKey) {
    event.preventDefault();
    restoreHistory(historyIndex - 1);
  }

  if (
    (event.ctrlKey || event.metaKey) &&
    (event.key.toLowerCase() === "y" || (event.shiftKey && event.key.toLowerCase() === "z"))
  ) {
    event.preventDefault();
    restoreHistory(historyIndex + 1);
  }
});

document.addEventListener("mouseup", () => {
  if (isSelectingCells) {
    isSelectingCells = false;
    selectionAnchorPoint = null;
    selectionAnchorInfo = null;
    selectionMode = "range";

    if (selectedCells.length === 1) {
      selectedCells[0].focus();
    }
  }

  if (resizeState) {
    captureHistory();
    resizeState = null;
  }
});

document.addEventListener("mousemove", (event) => {
  if (!resizeState) {
    return;
  }

  if (resizeState.type === "column") {
    const width = resizeState.startWidth + event.clientX - resizeState.startX;
    setColumnWidth(resizeState.table, resizeState.columnIndex, width);
    return;
  }

  if (resizeState.type === "row") {
    const height = resizeState.startHeight + event.clientY - resizeState.startY;
    setRowHeight(resizeState.row, height);
  }
});

undoButton.addEventListener("click", () => restoreHistory(historyIndex - 1));
redoButton.addEventListener("click", () => restoreHistory(historyIndex + 1));
addParagraphButton.addEventListener("click", insertParagraph);

fontFamilySelect.addEventListener("change", (event) => {
  if (!event.target.value) {
    return;
  }

  applyToTargets((target) => {
    target.style.fontFamily = event.target.value;
  });
});

fontSizeSelect.addEventListener("change", (event) => {
  if (!event.target.value) {
    return;
  }

  applyToTargets((target) => {
    target.style.fontSize = event.target.value;
  });
});

boldButton.addEventListener("click", () => {
  applyToTargets((target) => {
    target.style.fontWeight = target.style.fontWeight === "700" ? "" : "700";
  });
});

italicButton.addEventListener("click", () => {
  applyToTargets((target) => {
    target.style.fontStyle = target.style.fontStyle === "italic" ? "" : "italic";
  });
});

underlineButton.addEventListener("click", () => {
  applyToTargets((target) => {
    target.style.textDecoration = target.style.textDecoration.includes("underline")
      ? ""
      : "underline";
  });
});

alignLeftButton.addEventListener("click", () => applyAlignment("left"));
alignCenterButton.addEventListener("click", () => applyAlignment("center"));
alignRightButton.addEventListener("click", () => applyAlignment("right"));

lineHeightSelect.addEventListener("change", (event) => {
  if (!event.target.value) {
    return;
  }

  applyToTargets((target) => {
    target.style.lineHeight = event.target.value;
  });
});

backgroundColorInput.addEventListener("input", (event) => {
  applyToTargets((target) => {
    target.style.backgroundColor = event.target.value;
  });
});

textColorInput.addEventListener("input", (event) => {
  applyToTargets((target) => {
    target.style.color = event.target.value;
  });
});

applyBorderButton.addEventListener("click", () => {
  const width = borderWidthSelect.value || "1px";
  const color = borderColorInput.value || "#222222";

  applyToTargets((target) => {
    target.style.setProperty("border", `${width} solid ${color}`, "important");
    target.style.setProperty("border-width", width, "important");
    target.style.setProperty("border-style", "solid", "important");
    target.style.setProperty("border-color", color, "important");
  });
});

clearFormatButton.addEventListener("click", clearFormatting);

addRowButton.addEventListener("click", addRowAfterSelection);
deleteRowButton.addEventListener("click", deleteSelectedRows);
addColumnButton.addEventListener("click", addColumnAfterSelection);
deleteColumnButton.addEventListener("click", deleteSelectedColumn);
mergeCellsButton.addEventListener("click", mergeSelectedCells);
splitCellButton.addEventListener("click", splitSelectedCell);

function formatCount(value) {
  if (typeof value !== "number" || Number.isNaN(value)) {
    return "-";
  }

  return new Intl.NumberFormat("ko-KR").format(Math.max(0, value));
}

function getOrCreateStorageValue(storage, key) {
  const existing = storage.getItem(key);

  if (existing) {
    return existing;
  }

  const created = crypto.randomUUID();
  storage.setItem(key, created);
  return created;
}

async function fetchJson(url, options = {}) {
  const response = await fetch(url, {
    cache: "no-store",
    ...options,
    headers: {
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });

  if (!response.ok) {
    throw new Error(`Request failed: ${response.status}`);
  }

  return response.json();
}

function renderStats(stats) {
  liveVisitorsNode.textContent = formatCount(stats.liveVisitors ?? 0);
  totalVisitorsNode.textContent = formatCount(stats.totalVisitors ?? 0);
}

async function refreshStats() {
  const stats = await fetchJson("/api/stats/summary");
  renderStats(stats);
}

async function sendHeartbeat() {
  const stats = await fetchJson("/api/stats/session", {
    method: "POST",
    body: JSON.stringify({
      visitorId: statsVisitorId,
      sessionId: statsSessionId,
      pathname: window.location.pathname,
      referrer: document.referrer || "",
    }),
  });

  renderStats(stats);
}

function startHeartbeat() {
  if (heartbeatTimer) {
    clearInterval(heartbeatTimer);
  }

  heartbeatTimer = window.setInterval(() => {
    sendHeartbeat().catch(() => {
      liveVisitorsNode.textContent = "-";
    });
  }, HEARTBEAT_MS);
}

async function initializeStats() {
  try {
    statsVisitorId = getOrCreateStorageValue(window.localStorage, VISITOR_STORAGE_KEY);
    statsSessionId = getOrCreateStorageValue(window.sessionStorage, SESSION_STORAGE_KEY);
    await sendHeartbeat();
    startHeartbeat();
    await refreshStats();
  } catch (error) {
    liveVisitorsNode.textContent = "-";
    totalVisitorsNode.textContent = "-";
  }
}

document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible") {
    sendHeartbeat().catch(() => {
      liveVisitorsNode.textContent = "-";
    });
  }
});

resetOutput("엑셀에서 셀 범위를 복사한 뒤 여기에 붙여넣으세요.");
initializeStats();
