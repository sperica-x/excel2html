const pasteZone = document.querySelector("#pasteZone");
const convertButton = document.querySelector("#convertButton");
const copyButton = document.querySelector("#copyButton");
const clearButton = document.querySelector("#clearButton");
const htmlOutput = document.querySelector("#htmlOutput");
const preview = document.querySelector("#preview");
const statusNode = document.querySelector("#status");
const liveVisitorsNode = document.querySelector("#liveVisitors");
const totalVisitorsNode = document.querySelector("#totalVisitors");

let lastTableHtml = "";
let heartbeatTimer = null;
let statsSessionId = "";
let statsVisitorId = "";
const HEARTBEAT_MS = 45000;
const VISITOR_STORAGE_KEY = "board-insight-n-visitor-id";
const SESSION_STORAGE_KEY = "board-insight-n-session-id";

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

function normalizeHtmlTable(rawHtml) {
  const measurementRoot = createMeasurementRoot(rawHtml);
  const sourceTable = measurementRoot.querySelector("table");

  if (!sourceTable) {
    measurementRoot.remove();
    return "";
  }

  const rows = [...sourceTable.querySelectorAll("tr")];

  if (!rows.length) {
    measurementRoot.remove();
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

  const tableStyle = readTableStyle(sourceTable);
  measurementRoot.remove();

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

function readPastedTable() {
  const explicitHtml = pasteZone.dataset.excelHtml || "";
  const explicitText = pasteZone.dataset.plainText || "";

  if (explicitHtml) {
    return normalizeHtmlTable(explicitHtml);
  }

  const liveHtml = pasteZone.innerHTML;
  const liveText = pasteZone.innerText;

  if (liveHtml.includes("<table")) {
    return normalizeHtmlTable(liveHtml);
  }

  return normalizePlainTextTable(explicitText || liveText);
}

function renderTable(tableHtml) {
  lastTableHtml = tableHtml;
  htmlOutput.value = tableHtml;
  preview.innerHTML = tableHtml;
  preview.classList.remove("empty");
  copyButton.disabled = false;
  setStatus("HTML 테이블 코드로 변환했습니다.");
}

function resetOutput(message) {
  lastTableHtml = "";
  htmlOutput.value = "";
  preview.textContent = "변환 결과가 여기에 표시됩니다.";
  preview.classList.add("empty");
  copyButton.disabled = true;
  setStatus(message);
}

pasteZone.addEventListener("paste", (event) => {
  const html = event.clipboardData?.getData("text/html") ?? "";
  const text = event.clipboardData?.getData("text/plain") ?? "";

  pasteZone.dataset.excelHtml = html;
  pasteZone.dataset.plainText = text;

  if (html) {
    event.preventDefault();
    pasteZone.innerHTML = html;
    setStatus("엑셀 HTML 클립보드를 감지했습니다. 변환 버튼을 누르세요.");
    return;
  }

  if (text) {
    event.preventDefault();
    pasteZone.textContent = text;
    setStatus("탭 구분 텍스트를 감지했습니다. 변환 버튼을 누르세요.");
  }
});

convertButton.addEventListener("click", () => {
  const tableHtml = readPastedTable();

  if (!tableHtml) {
    resetOutput("붙여넣은 내용에서 표 데이터를 찾지 못했습니다.");
    statusNode.style.color = "#a11d1d";
    return;
  }

  renderTable(tableHtml);
});

copyButton.addEventListener("click", async () => {
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
  pasteZone.innerHTML = "여기에 Ctrl+V로 엑셀 표를 붙여넣으세요.";
  pasteZone.dataset.excelHtml = "";
  pasteZone.dataset.plainText = "";
  resetOutput("초기화했습니다. 새로 붙여넣으세요.");
});

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
