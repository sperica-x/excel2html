const pasteZone = document.querySelector("#pasteZone");
const convertButton = document.querySelector("#convertButton");
const copyButton = document.querySelector("#copyButton");
const clearButton = document.querySelector("#clearButton");
const htmlOutput = document.querySelector("#htmlOutput");
const preview = document.querySelector("#preview");
const statusNode = document.querySelector("#status");

let lastTableHtml = "";

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

function cleanupCell(node, tagName) {
  const clone = node.cloneNode(true);
  clone.querySelectorAll("br").forEach((br) => br.replaceWith("\n"));

  let content = clone.textContent ?? "";
  content = content.replace(/\u00a0/g, " ").replace(/\r\n/g, "\n").trim();

  return `<${tagName}>${escapeHtml(content).replaceAll("\n", "<br />")}</${tagName}>`;
}

function normalizeHtmlTable(rawHtml) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(rawHtml, "text/html");
  const sourceTable = doc.querySelector("table");

  if (!sourceTable) {
    return "";
  }

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
          const spanAttrs = [];

          if (cell.colSpan > 1) {
            spanAttrs.push(` colspan="${cell.colSpan}"`);
          }

          if (cell.rowSpan > 1) {
            spanAttrs.push(` rowspan="${cell.rowSpan}"`);
          }

          const normalizedCell = cleanupCell(cell, tagName);
          return normalizedCell.replace(`<${tagName}>`, `<${tagName}${spanAttrs.join("")}>`);
        })
        .join("");

      return `<tr>${normalizedCells}</tr>`;
    })
    .filter(Boolean)
    .join("");

  return `<table>${normalizedRows}</table>`;
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

resetOutput("엑셀에서 셀 범위를 복사한 뒤 여기에 붙여넣으세요.");
