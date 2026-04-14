const analyzeForm = document.querySelector("#analyze-form");
const processButton = document.querySelector("#process-button");
const downloadButton = document.querySelector("#download-button");
const summarySection = document.querySelector("#summary");
const statusBox = document.querySelector("#status");
const isLocalPreviewHost = ["localhost", "127.0.0.1"].includes(window.location.hostname);
const API_BASE_URL = isLocalPreviewHost && window.location.port !== "3000" ? "http://localhost:3000" : "";

const summaryTargets = {
  templateName: document.querySelector("#template-name"),
  dataName: document.querySelector("#data-name"),
  recordCount: document.querySelector("#record-count"),
  slideCount: document.querySelector("#slide-count"),
  generationMode: document.querySelector("#generation-mode"),
  columnsList: document.querySelector("#columns-list"),
  placeholdersList: document.querySelector("#placeholders-list"),
  matchedList: document.querySelector("#matched-list"),
  missingList: document.querySelector("#missing-list"),
  unusedList: document.querySelector("#unused-list"),
  slidesList: document.querySelector("#slides-list"),
  warningsList: document.querySelector("#warnings-list")
};

const state = {
  jobId: null,
  canGenerate: false
};

function setStatus(message, tone = "default") {
  statusBox.textContent = message;
  statusBox.className = "status";

  if (tone === "loading") {
    statusBox.classList.add("is-loading");
  } else if (tone === "success") {
    statusBox.classList.add("is-success");
  } else if (tone === "error") {
    statusBox.classList.add("is-error");
  }
}

function clearNode(node) {
  while (node.firstChild) {
    node.removeChild(node.firstChild);
  }
}

function renderTagList(node, items, emptyLabel) {
  clearNode(node);

  if (!items.length) {
    const span = document.createElement("span");
    span.className = "tag";
    span.textContent = emptyLabel;
    node.appendChild(span);
    return;
  }

  items.forEach((item) => {
    const span = document.createElement("span");
    span.className = "tag";
    span.textContent = item;
    node.appendChild(span);
  });
}

function renderList(node, items, emptyLabel) {
  clearNode(node);

  if (!items.length) {
    const listItem = document.createElement("li");
    listItem.textContent = emptyLabel;
    node.appendChild(listItem);
    return;
  }

  items.forEach((item) => {
    const listItem = document.createElement("li");
    listItem.textContent = item;
    node.appendChild(listItem);
  });
}

function renderSummary(data) {
  summaryTargets.templateName.textContent = data.templateFileName;
  summaryTargets.dataName.textContent = data.dataFileName;
  summaryTargets.recordCount.textContent = String(data.recordCount);
  summaryTargets.slideCount.textContent = String(data.slideCount);
  summaryTargets.generationMode.textContent = data.generationMode;

  renderTagList(summaryTargets.columnsList, data.columns, "Sin columnas");
  renderTagList(summaryTargets.placeholdersList, data.placeholders, "Sin placeholders");
  renderList(
    summaryTargets.matchedList,
    data.matchedFields.map((item) => `${item.placeholder} -> ${item.column}`),
    "No hubo coincidencias."
  );
  renderList(summaryTargets.missingList, data.missingPlaceholders, "No faltan placeholders.");
  renderList(summaryTargets.unusedList, data.unusedColumns, "No hay columnas extra.");
  renderList(
    summaryTargets.slidesList,
    data.templateSlides.map(
      (slide) =>
        `Slide ${slide.index}: ${slide.placeholders.join(", ")}${slide.previewText ? ` | ${slide.previewText}` : ""}`
    ),
    "No se detectaron slides template."
  );
  renderList(summaryTargets.warningsList, data.warnings, "Sin advertencias.");

  summarySection.classList.remove("hidden");
}

function parseDownloadFileName(response) {
  const header = response.headers.get("Content-Disposition");
  if (!header) {
    return "cenefas-procesado.pptx";
  }

  const match = header.match(/filename="([^"]+)"/i);
  return match?.[1] ?? "cenefas-procesado.pptx";
}

function buildApiUrl(pathname) {
  return `${API_BASE_URL}${pathname}`;
}

function getReadableError(error) {
  if (error instanceof TypeError) {
    return "No pude conectarme con el backend. Levanta `npm.cmd start` para abrir la API en http://localhost:3000.";
  }

  return error.message || "Ocurrio un error inesperado.";
}

async function requestOutput() {
  const response = await fetch(buildApiUrl("/api/generate"), {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      jobId: state.jobId,
      format: "pptx"
    })
  });

  if (!response.ok) {
    const payload = await response.json();
    throw new Error(payload.message || "No se pudo generar el PPTX.");
  }

  return response;
}

async function triggerDownload() {
  if (!state.jobId || !state.canGenerate) {
    throw new Error("Primero necesitas procesar un template valido.");
  }

  const response = await requestOutput();
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = parseDownloadFileName(response);
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

async function analyzeAndGenerate(event) {
  event.preventDefault();

  const formData = new FormData(analyzeForm);
  let analysisReady = false;

  if (!formData.get("template")?.name) {
    setStatus("Falta subir el archivo template .pptx.", "error");
    return;
  }

  if (!formData.get("data")?.name) {
    setStatus("Falta subir el archivo de datos .xlsx o .csv.", "error");
    return;
  }

  processButton.disabled = true;
  downloadButton.disabled = true;
  setStatus("Analizando archivos y preparando la exportacion...", "loading");

  try {
    const response = await fetch(buildApiUrl("/api/analyze"), {
      method: "POST",
      body: formData
    });

    const payload = await response.json();

    if (!response.ok) {
      throw new Error(payload.message || "No se pudo analizar la carga.");
    }

    state.jobId = payload.jobId;
    state.canGenerate = payload.canGenerate;
    analysisReady = true;

    renderSummary(payload);

    if (!payload.canGenerate) {
      throw new Error("El template no contiene placeholders validos para generar la salida.");
    }

    setStatus("Analisis listo. Generando y descargando el PPTX final...", "loading");
    await triggerDownload();

    downloadButton.disabled = false;

    const statusMessage =
      payload.warnings.length > 0
        ? "Se descargo el PPTX final. Revisa tambien las advertencias del resumen."
        : "Se descargo el PPTX final correctamente.";

    setStatus(statusMessage, "success");
  } catch (error) {
    if (!analysisReady) {
      state.jobId = null;
      state.canGenerate = false;
      downloadButton.disabled = true;
    } else {
      downloadButton.disabled = !state.canGenerate;
    }

    setStatus(getReadableError(error), "error");
  } finally {
    processButton.disabled = false;
  }
}

async function redownloadOutput() {
  if (!state.jobId || !state.canGenerate) {
    setStatus("Primero necesitas procesar ambos archivos para poder descargar.", "error");
    return;
  }

  downloadButton.disabled = true;
  setStatus("Generando nuevamente el PowerPoint final...", "loading");

  try {
    await triggerDownload();
    setStatus("La descarga se lanzo nuevamente.", "success");
  } catch (error) {
    setStatus(getReadableError(error), "error");
  } finally {
    downloadButton.disabled = false;
  }
}

analyzeForm.addEventListener("submit", analyzeAndGenerate);
downloadButton.addEventListener("click", redownloadOutput);
