"use strict";

const STORAGE_KEY = "kesb-utm-pakej3-itp-draft-v1";
const DEFAULT_PROJECT =
  "UTM – Baikpulih Makmal P & P di Bangunan Akademik (RP1-RMKe-13) Kampus Universiti Teknologi Malaysia, Johor Bahru – Pakej 3: Fakulti Komputeran (FKOM) & Fakulti Kejuruteraan Mekanikal (FKM)";
const FIELD_DEFINITIONS = [
  { key: "aktiviti", label: "Aktiviti / pemeriksaan", kind: "text", required: true },
  { key: "kriteria", label: "Kriteria penerimaan", kind: "textarea", wide: true },
  { key: "rujukan", label: "Rujukan spesifikasi / lukisan", kind: "text" },
  { key: "kekerapan", label: "Kekerapan", kind: "text" },
  { key: "rekod", label: "Rekod / borang", kind: "text" },
  { key: "titik", label: "Titik pemeriksaan (H/W/R)", kind: "select" },
  { key: "kontraktor", label: "Kontraktor / wakil pemeriksa", kind: "text" },
  { key: "so_perunding", label: "S.O. / perunding", kind: "text" },
  { key: "pemeriksaan_bersama_so", label: "Pemeriksaan bersama S.O.", kind: "checkbox", wide: true },
  { key: "catatan", label: "Catatan", kind: "textarea", wide: true },
];

const tradeList = document.querySelector("#trade-list");
const projectName = document.querySelector("#project-name");
const statusMessage = document.querySelector("#save-status");
const filePicker = document.querySelector("#file-picker");
let data = { projek: DEFAULT_PROJECT, itp: [], ujian_bahan: [] };

function createElement(tag, className, text) {
  const element = document.createElement(tag);
  if (className) element.className = className;
  if (text !== undefined) element.textContent = text;
  return element;
}

function setStatus(message, isError) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("status-error", Boolean(isError));
}

function saveDraft(message) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    setStatus(message || "Draf disimpan pada pelayar ini.", false);
  } catch (error) {
    setStatus(`Draf tidak dapat disimpan dalam pelayar: ${error.message}`, true);
  }
}

function isValidData(value) {
  const isObject = (item) => item !== null && typeof item === "object" && !Array.isArray(item);
  return isObject(value) &&
    typeof value.projek === "string" &&
    Array.isArray(value.itp) &&
    value.itp.every((trade) => {
      if (!isObject(trade) || typeof trade.trade !== "string" || !Array.isArray(trade.aktiviti)) return false;
      return trade.aktiviti.every((activity) => {
        if (!isObject(activity)) return false;
        const textFields = [
          "aktiviti", "kriteria", "rujukan", "kekerapan", "rekod", "kontraktor", "so_perunding", "catatan",
        ];
        if (textFields.some((key) => key in activity && typeof activity[key] !== "string")) return false;
        if ("titik" in activity && (typeof activity.titik !== "string" || !/^[HWR]$/i.test(activity.titik))) {
          return false;
        }
        return !("pemeriksaan_bersama_so" in activity) ||
          typeof activity.pemeriksaan_bersama_so === "boolean";
      });
    }) &&
    (value.ujian_bahan === undefined ||
      (Array.isArray(value.ujian_bahan) && value.ujian_bahan.every(isObject)));
}

function loadSavedDraft() {
  let saved;
  try {
    saved = localStorage.getItem(STORAGE_KEY);
  } catch (error) {
    setStatus(`Storan pelayar tidak dapat dibaca: ${error.message}`, true);
    return;
  }
  if (!saved) return;

  try {
    const parsed = JSON.parse(saved);
    if (!isValidData(parsed)) throw new Error("Struktur draf tidak sah.");
    data = parsed;
    setStatus("Draf terdahulu dimuatkan daripada pelayar ini.", false);
  } catch (error) {
    setStatus(`Draf tersimpan tidak dapat dimuatkan: ${error.message}`, true);
  }
}

function makeLabeledField(definition, tradeIndex, activityIndex, activity) {
  const wrapper = createElement(
    "label",
    `field${definition.wide ? " field-wide" : ""}${definition.kind === "checkbox" ? " field-checkbox" : ""}`
  );
  const caption = createElement("span", "", definition.label);
  let control;

  if (definition.kind === "textarea") {
    control = document.createElement("textarea");
    control.rows = 2;
  } else if (definition.kind === "select") {
    control = document.createElement("select");
    [
      ["R", "R — Review"],
      ["H", "H — Hold point"],
      ["W", "W — Witness point"],
    ].forEach(([value, label]) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = label;
      control.append(option);
    });
  } else if (definition.kind === "checkbox") {
    control = document.createElement("input");
    control.type = "checkbox";
    control.checked = activity[definition.key] === true;
  } else {
    control = document.createElement("input");
    control.type = "text";
    if (definition.required) control.required = true;
  }

  if (definition.kind !== "checkbox") {
    control.value = activity[definition.key] ?? (definition.key === "titik" ? "R" : "");
  }
  control.dataset.tradeIndex = String(tradeIndex);
  control.dataset.activityIndex = String(activityIndex);
  control.dataset.field = definition.key;
  wrapper.append(caption, control);
  return wrapper;
}

function render() {
  projectName.value = data.projek;
  tradeList.replaceChildren();

  if (data.itp.length === 0) {
    tradeList.append(
      createElement("p", "empty-state", "Belum ada trade. Pilih “Tambah trade” untuk mula membina ITP.")
    );
    return;
  }

  data.itp.forEach((trade, tradeIndex) => {
    const card = createElement("article", "trade-card");
    const header = createElement("div", "trade-header");
    const title = createElement("h2", "", `Trade ${tradeIndex + 1}`);
    const tradeField = createElement("label", "field trade-name");
    const tradeCaption = createElement("span", "", "Nama trade");
    const tradeInput = document.createElement("input");
    tradeInput.type = "text";
    tradeInput.value = trade.trade;
    tradeInput.placeholder = "Contoh: Kerja seni bina";
    tradeInput.dataset.tradeName = String(tradeIndex);
    tradeField.append(tradeCaption, tradeInput);

    const removeTrade = createElement("button", "button button-danger", "Padam trade");
    removeTrade.type = "button";
    removeTrade.dataset.action = "remove-trade";
    removeTrade.dataset.tradeIndex = String(tradeIndex);
    header.append(title, tradeField, removeTrade);
    card.append(header, createElement("h3", "activities-heading", "Aktiviti pemeriksaan"));

    trade.aktiviti.forEach((activity, activityIndex) => {
      const activityCard = createElement("section", "activity-card");
      const activityHeader = createElement("div", "activity-header");
      activityHeader.append(createElement("h4", "", `Aktiviti ${tradeIndex + 1}.${activityIndex + 1}`));
      const removeActivity = createElement("button", "button button-link-danger", "Padam aktiviti");
      removeActivity.type = "button";
      removeActivity.dataset.action = "remove-activity";
      removeActivity.dataset.tradeIndex = String(tradeIndex);
      removeActivity.dataset.activityIndex = String(activityIndex);
      activityHeader.append(removeActivity);

      const fields = createElement("div", "field-grid");
      FIELD_DEFINITIONS.forEach((definition) => {
        fields.append(makeLabeledField(definition, tradeIndex, activityIndex, activity));
      });
      activityCard.append(activityHeader, fields);
      card.append(activityCard);
    });

    const addActivity = createElement("button", "button button-add-activity", "+ Tambah aktiviti");
    addActivity.type = "button";
    addActivity.dataset.action = "add-activity";
    addActivity.dataset.tradeIndex = String(tradeIndex);
    card.append(addActivity);
    tradeList.append(card);
  });
}

function newActivity() {
  return {
    aktiviti: "",
    kriteria: "",
    rujukan: "",
    kekerapan: "",
    rekod: "",
    titik: "R",
    kontraktor: "",
    so_perunding: "",
    pemeriksaan_bersama_so: false,
    catatan: "",
  };
}

document.querySelector("#add-trade").addEventListener("click", () => {
  data.itp.push({ trade: "", aktiviti: [] });
  render();
  saveDraft();
  tradeList.querySelector(".trade-card:last-child .trade-name input")?.focus();
});

document.querySelector("#save-draft").addEventListener("click", () => saveDraft("Draf disimpan pada pelayar ini."));

document.querySelector("#download-json").addEventListener("click", () => {
  const json = JSON.stringify(data, null, 2);
  const blob = new Blob([json], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `itp-${new Date().toISOString().slice(0, 10)}.json`;
  document.body.append(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  setStatus("Fail JSON ITP telah dimuat turun.", false);
});

document.querySelector("#import-json").addEventListener("click", () => filePicker.click());

filePicker.addEventListener("change", async () => {
  const file = filePicker.files && filePicker.files[0];
  if (!file) return;
  try {
    const parsed = JSON.parse(await file.text());
    if (!isValidData(parsed)) {
      throw new Error("Fail mesti mengandungi projek (teks) dan itp (senarai trade serta aktiviti).");
    }
    if (!window.confirm("Import akan menggantikan draf semasa. Teruskan?")) return;
    data = parsed;
    if (!Array.isArray(data.ujian_bahan)) data.ujian_bahan = [];
    render();
    saveDraft("Data JSON berjaya diimport dan draf disimpan.");
  } catch (error) {
    setStatus(`Import gagal: ${error.message}`, true);
  } finally {
    filePicker.value = "";
  }
});

projectName.addEventListener("input", () => {
  data.projek = projectName.value;
  saveDraft();
});

tradeList.addEventListener("input", (event) => {
  const control = event.target;
  if (control.type === "checkbox") return;
  if (control.dataset.tradeName !== undefined) {
    data.itp[Number(control.dataset.tradeName)].trade = control.value;
  } else if (control.dataset.field) {
    const activity = data.itp[Number(control.dataset.tradeIndex)].aktiviti[Number(control.dataset.activityIndex)];
    activity[control.dataset.field] = control.value;
  } else {
    return;
  }
  saveDraft();
});

tradeList.addEventListener("change", (event) => {
  const control = event.target;
  if (!control.dataset.field) return;
  const activity = data.itp[Number(control.dataset.tradeIndex)].aktiviti[Number(control.dataset.activityIndex)];
  activity[control.dataset.field] = control.type === "checkbox" ? control.checked : control.value;
  saveDraft();
});

tradeList.addEventListener("click", (event) => {
  const button = event.target.closest("button[data-action]");
  if (!button) return;
  const tradeIndex = Number(button.dataset.tradeIndex);
  if (button.dataset.action === "add-activity") {
    data.itp[tradeIndex].aktiviti.push(newActivity());
  } else if (button.dataset.action === "remove-activity") {
    if (!window.confirm("Padam aktiviti ini?")) return;
    data.itp[tradeIndex].aktiviti.splice(Number(button.dataset.activityIndex), 1);
  } else if (button.dataset.action === "remove-trade") {
    if (!window.confirm("Padam trade ini bersama semua aktiviti di dalamnya?")) return;
    data.itp.splice(tradeIndex, 1);
  } else {
    return;
  }
  render();
  saveDraft();
});

if ("serviceWorker" in navigator && window.location.protocol !== "file:") {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./service-worker.js").catch((error) => {
      setStatus(`Mod luar talian tidak dapat diaktifkan: ${error.message}`, true);
    });
  });
}

loadSavedDraft();
render();
