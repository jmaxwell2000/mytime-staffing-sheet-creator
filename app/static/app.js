const form = document.querySelector("#convert-form");
const fileInput = document.querySelector("#workbook");
const fileLabel = document.querySelector("#file-label");
const includeBackPages = document.querySelector("#include-back-pages");
const offsetControl = document.querySelector("#offset-control");
const offsetInput = document.querySelector("#back-page-offset");
const offsetValue = document.querySelector("#offset-value");
const submitButton = document.querySelector("#submit-button");
const statusText = document.querySelector("#status");

const OFFSET_KEY = "mytime.backPageOffset";

function updateOffsetLabel() {
  offsetValue.textContent = `${Number(offsetInput.value).toFixed(2)} in`;
}

function setStatus(message, isError = false) {
  statusText.textContent = message;
  statusText.classList.toggle("error", isError);
}

function syncOffsetState() {
  offsetControl.classList.toggle("is-disabled", !includeBackPages.checked);
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

fileInput.addEventListener("change", () => {
  fileLabel.textContent = fileInput.files?.[0]?.name || "Choose an .xlsx workbook";
});

includeBackPages.addEventListener("change", syncOffsetState);

offsetInput.addEventListener("input", () => {
  updateOffsetLabel();
  localStorage.setItem(OFFSET_KEY, Number(offsetInput.value).toFixed(2));
});

form.addEventListener("submit", async (event) => {
  setStatus("Converting...");
  submitButton.disabled = true;
});
