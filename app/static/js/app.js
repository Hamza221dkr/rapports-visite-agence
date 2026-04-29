let currentFile = null, downloadUrl = null, downloadName = null;

document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("dateVisite").value = new Date().toISOString().split("T")[0];

  // Restore saved API key from localStorage
  const savedKey = localStorage.getItem("openai_api_key");
  if (savedKey) {
    document.getElementById("apiKey").value = savedKey;
    document.getElementById("savedKeyHint").style.display = "block";
  }

  const zone = document.getElementById("dropZone");
  zone.addEventListener("dragover",  e => { e.preventDefault(); zone.classList.add("dragover"); });
  zone.addEventListener("dragleave", ()  => zone.classList.remove("dragover"));
  zone.addEventListener("drop",      e  => { e.preventDefault(); zone.classList.remove("dragover"); if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); });
  zone.addEventListener("click",     ()  => document.getElementById("fileInput").click());
  document.getElementById("fileInput").addEventListener("change", e => { if (e.target.files[0]) handleFile(e.target.files[0]); });
});

function handleFile(file) {
  if (!file.name.match(/\.(xlsx|xls)$/i)) { showError("Format non supporté. Utilisez un fichier .xlsx ou .xls"); return; }
  currentFile = file;
  document.getElementById("fileName").textContent = file.name;
  document.getElementById("fileInfo").style.display = "flex";
  document.getElementById("dropZone").style.display = "none";
  hideError(); hideResult();
  loadAgencies(file);
}

function resetFile() {
  currentFile = null;
  document.getElementById("fileInput").value = "";
  document.getElementById("fileInfo").style.display = "none";
  document.getElementById("dropZone").style.display = "flex";
  document.getElementById("optionsBlock").style.display = "none";
  document.getElementById("agenciesGrid").innerHTML = "";
  document.getElementById("gptCheck").checked = false;
  document.getElementById("gptKeyBlock").style.display = "none";
  document.getElementById("gptCard").classList.remove("active");
  document.getElementById("gptBadge").style.display = "none";
  hideError(); hideResult();
}
function resetAll() { resetFile(); }

async function loadAgencies(file) {
  const fd = new FormData(); fd.append("file", file);
  try {
    const res = await fetch("/api/agencies", { method: "POST", body: fd });
    const data = await res.json();
    if (data.error) { showError(data.error); return; }
    renderAgencies(data.agencies);
    document.getElementById("optionsBlock").style.display = "block";
    updateHint();
  } catch(e) { showError("Erreur lecture fichier : " + e.message); }
}

function renderAgencies(agencies) {
  const grid = document.getElementById("agenciesGrid"); grid.innerHTML = "";
  agencies.forEach(name => {
    const chip = document.createElement("label"); chip.className = "agency-chip selected";
    chip.innerHTML = `<input type="checkbox" value="${name}" checked onchange="onChipChange(this)"/> ${name}`;
    grid.appendChild(chip);
  }); updateHint();
}

function onChipChange(cb) { cb.closest(".agency-chip").classList.toggle("selected", cb.checked); updateHint(); }

function toggleAll(state) {
  document.querySelectorAll(".agency-chip input").forEach(cb => { cb.checked = state; cb.closest(".agency-chip").classList.toggle("selected", state); });
  updateHint();
}

function updateHint() {
  const n = document.querySelectorAll(".agency-chip input:checked").length;
  const consol = document.getElementById("consolidatedCheck").checked;
  const gpt = document.getElementById("gptCheck").checked;
  const h = document.getElementById("generateHint");
  const parts = [];
  if (n === 0) { h.textContent = "Aucune agence sélectionnée"; h.style.color = "#C00000"; return; }
  parts.push(n === 1 ? "1 rapport Word" : `${n} rapports Word`);
  if (consol && n > 1) parts.push("1 rapport consolidé");
  if (gpt) parts.push("analyse IA activée");
  h.textContent = parts.join(" · ");
  h.style.color = "#595959";
}

function toggleGptKey() {
  const checked = document.getElementById("gptCheck").checked;
  document.getElementById("gptKeyBlock").style.display = checked ? "block" : "none";
  document.getElementById("gptCard").classList.toggle("active", checked);
  document.getElementById("gptBadge").style.display = checked ? "flex" : "none";
  updateHint();
}

function toggleKeyVisibility() {
  const input = document.getElementById("apiKey");
  input.type = input.type === "password" ? "text" : "password";
}

function forgetApiKey() {
  localStorage.removeItem("openai_api_key");
  document.getElementById("apiKey").value = "";
  document.getElementById("savedKeyHint").style.display = "none";
}

async function generate() {
  const selected = [...document.querySelectorAll(".agency-chip input:checked")].map(cb => cb.value);
  if (!currentFile) { showError("Veuillez d'abord charger un fichier Excel."); return; }
  if (selected.length === 0) { showError("Veuillez sélectionner au moins une agence."); return; }

  const rzName    = document.getElementById("rzName").value.trim();
  const dateRaw   = document.getElementById("dateVisite").value;
  const dateVis   = dateRaw ? dateRaw.split("-").reverse().join("/") : "";
  const useGpt    = document.getElementById("gptCheck").checked;
  const apiKey    = document.getElementById("apiKey").value.trim();
  const consol    = document.getElementById("consolidatedCheck").checked;

  if (useGpt && !apiKey) { showError("Veuillez saisir votre clé API OpenAI pour activer l'analyse IA."); return; }

  // Save key to localStorage if provided
  if (useGpt && apiKey) {
    localStorage.setItem("openai_api_key", apiKey);
    document.getElementById("savedKeyHint").style.display = "block";
  }

  // Show loader
  document.getElementById("optionsBlock").style.display = "none";
  document.getElementById("loaderBlock").style.display  = "flex";
  hideError(); hideResult();

  setLoaderState(useGpt);

  const fd = new FormData();
  fd.append("file", currentFile);
  fd.append("rz_name", rzName);
  fd.append("date_visite", dateVis);
  fd.append("use_gpt", useGpt ? "true" : "false");
  fd.append("api_key", apiKey);
  fd.append("consolidated", consol ? "true" : "false");
  selected.forEach(a => fd.append("agencies", a));

  try {
    const res = await fetch("/api/generate", { method: "POST", body: fd });
    if (!res.ok) { const d = await res.json().catch(() => ({ error: "Erreur inconnue" })); throw new Error(d.error || "Erreur serveur"); }

    const disp = res.headers.get("Content-Disposition") || "";
    const match = disp.match(/filename="?([^"]+)"?/);
    downloadName = match ? match[1] : "Rapports_Visite_Agences.zip";

    const blob = await res.blob();
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
    downloadUrl = URL.createObjectURL(blob);

    const isZip = downloadName.endsWith(".zip");
    let msg = selected.length === 1 && !consol
      ? `Rapport agence ${selected[0]} généré avec succès`
      : `${selected.length} rapport(s) individuel(s)${consol && selected.length > 1 ? " + 1 rapport consolidé" : ""} — ZIP prêt`;
    if (useGpt) msg += " · Analyse IA incluse";

    document.getElementById("resultText").textContent = msg;
    document.getElementById("downloadBtn").onclick = () => {
      const a = document.createElement("a"); a.href = downloadUrl; a.download = downloadName; a.click();
    };

    document.getElementById("loaderBlock").style.display  = "none";
    document.getElementById("resultBlock").style.display  = "flex";
    document.getElementById("optionsBlock").style.display = "block";

  } catch(e) {
    document.getElementById("loaderBlock").style.display  = "none";
    document.getElementById("optionsBlock").style.display = "block";
    showError(e.message);
  }
}

function setLoaderState(useGpt) {
  const text = document.getElementById("loaderText");
  const sub  = document.getElementById("loaderSub");
  if (useGpt) {
    text.textContent = "Génération & analyse IA en cours…";
    sub.textContent  = "ChatGPT rédige les appréciations — cela peut prendre 15-30 secondes.";
  } else {
    text.textContent = "Génération en cours…";
    sub.textContent  = "";
  }
}

function showError(msg) { document.getElementById("errorText").textContent = msg; document.getElementById("errorBlock").style.display = "flex"; }
function hideError()    { document.getElementById("errorBlock").style.display = "none"; }
function hideResult()   { document.getElementById("resultBlock").style.display = "none"; }
