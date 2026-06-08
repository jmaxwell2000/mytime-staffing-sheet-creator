const form = document.querySelector("#convert-form");
const fileInput = document.querySelector("#workbook");
const fileLabel = document.querySelector("#file-label");
const optionsPanel = document.querySelector("#options-panel");
const includeBackPages = document.querySelector("#include-back-pages");
const offsetControl = document.querySelector("#offset-control");
const offsetInput = document.querySelector("#back-page-offset");
const offsetValue = document.querySelector("#offset-value");
const submitButton = document.querySelector("#submit-button");
const submitButtonLabel = document.querySelector("#submit-button-label");
const statusText = document.querySelector("#status");

const INCLUDE_BACK_PAGES_KEY = "mytime.includeBackPages";
const OFFSET_KEY = "mytime.backPageOffset";
const BUTTON_LABELS = {
  idle: "Upload file",
  ready: "Create staffing sheets",
  processing: "Creating...",
  downloaded: "Downloaded",
  error: "Try again",
};

let currentState = "idle";

function updateOffsetLabel() {
  offsetValue.textContent = `${Number(offsetInput.value).toFixed(2)} in`;
}

function setStatus(message, variant = "") {
  statusText.textContent = message;
  statusText.classList.toggle("error", variant === "error");
  statusText.classList.toggle("success", variant === "success");
  statusText.classList.toggle("busy", variant === "busy");
}

function setState(state, message = "", statusVariant = "") {
  currentState = state;
  submitButton.dataset.state = state;
  submitButtonLabel.textContent = BUTTON_LABELS[state];
  submitButton.disabled = state === "processing";
  setStatus(message, statusVariant);
}

function hasSelectedFile() {
  return Boolean(fileInput.files?.length);
}

function syncOffsetState() {
  offsetControl.classList.toggle("is-disabled", !includeBackPages.checked);
  offsetInput.disabled = !includeBackPages.checked;
}

function restoreDefaults() {
  includeBackPages.checked = localStorage.getItem(INCLUDE_BACK_PAGES_KEY) === "true";
  if (includeBackPages.checked) {
    optionsPanel.open = true;
  }

  const savedOffset = localStorage.getItem(OFFSET_KEY);
  if (savedOffset !== null) {
    const value = Number(savedOffset);
    if (value >= -0.1 && value <= 0.3) {
      offsetInput.value = value.toFixed(2);
    }
  }

  updateOffsetLabel();
  syncOffsetState();
}

function resetForSelectedFile() {
  fileLabel.textContent = fileInput.files?.[0]?.name || "Choose an .xlsx workbook";
  setState(hasSelectedFile() ? "ready" : "idle");
}

function parseDownloadFilename(response) {
  const disposition = response.headers.get("Content-Disposition") || "";
  const encodedMatch = disposition.match(/filename\*=UTF-8''([^;]+)/i);
  if (encodedMatch) {
    return decodeURIComponent(encodedMatch[1].trim());
  }

  const quotedMatch = disposition.match(/filename="([^"]+)"/i);
  if (quotedMatch) {
    return quotedMatch[1].trim();
  }

  const plainMatch = disposition.match(/filename=([^;]+)/i);
  if (plainMatch) {
    return plainMatch[1].trim();
  }

  return "staffing-sheets.xlsx";
}

async function safeErrorMessage(response) {
  try {
    const payload = await response.json();
    if (typeof payload.detail === "string") {
      return payload.detail;
    }
  } catch {
    // Keep the message generic if the server did not return JSON.
  }
  return "Could not create staffing sheets. Please check the workbook and try again.";
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

restoreDefaults();
setState("idle");

submitButton.addEventListener("click", (event) => {
  if (currentState === "idle" || currentState === "downloaded") {
    event.preventDefault();
    fileInput.click();
  }
});

fileInput.addEventListener("change", resetForSelectedFile);

includeBackPages.addEventListener("change", () => {
  localStorage.setItem(INCLUDE_BACK_PAGES_KEY, includeBackPages.checked ? "true" : "false");
  syncOffsetState();
});

offsetInput.addEventListener("input", () => {
  updateOffsetLabel();
  localStorage.setItem(OFFSET_KEY, Number(offsetInput.value).toFixed(2));
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  if (!hasSelectedFile()) {
    fileInput.click();
    return;
  }

  setState("processing", "Creating staffing sheets...", "busy");

  try {
    const response = await fetch(form.action, {
      method: "POST",
      body: new FormData(form),
    });

    if (!response.ok) {
      throw new Error(await safeErrorMessage(response));
    }

    const filename = parseDownloadFilename(response);
    const blob = await response.blob();
    downloadBlob(blob, filename);
    setState("downloaded", "Downloaded.", "success");
  } catch (error) {
    setState("error", error.message, "error");
  }
});
