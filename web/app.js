const ROWS = 30;
const COLUMNS = 6;
const LOCAL_KEY = "health.sante.document.v1";
const LEGACY_ROWS_KEY = "health.sante.rows.v2";
const THEME_KEY = "health.sante.theme.v2";
const DEVICE_KEY = "health.sante.device.v1";
const TOKEN_KEY = "health.sante.onedrive.tokens.v1";
const LAST_SYNC_KEY = "health.sante.onedrive.lastSync.v1";
const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";
const GRAPH_SCOPES = "openid profile offline_access Files.ReadWrite.AppFolder";
const SYNC_CONFIG = window.SANTE_SYNC_CONFIG || {};
const COLORS = ["#0071e3", "#34c759", "#ff9500", "#af52de", "#ff2d55"];
const PLAY_ICON = '<svg aria-hidden="true" viewBox="0 0 24 24"><path d="M8 5v14l11-7-11-7Z"/></svg>';
const PAUSE_ICON = '<svg aria-hidden="true" viewBox="0 0 24 24"><path d="M8 5v14M16 5v14"/></svg>';
const MOON_ICON = '<svg aria-hidden="true" viewBox="0 0 24 24"><path d="M20.4 14.5A7.8 7.8 0 0 1 9.5 3.6 8.6 8.6 0 1 0 20.4 14.5Z"/></svg>';
const SUN_ICON = '<svg aria-hidden="true" viewBox="0 0 24 24"><path d="M12 4V2m0 20v-2m8-8h2M2 12h2m14.4-6.4 1.4-1.4M4.2 19.8l1.4-1.4m0-12.8L4.2 4.2m15.6 15.6-1.4-1.4M12 16a4 4 0 1 0 0-8 4 4 0 0 0 0 8Z"/></svg>';

let documentData = makeDocument(makeDefaultRows(), false);
let rows = documentData.rows;
let saveTimer = null;
let statusTimer = null;
let stopwatchRunning = false;
let stopwatchStart = 0;
let stopwatchElapsed = 0;
let timerRunning = false;
let timerRemaining = 0;
let timerLastTick = 0;
let hiddenSeries = new Set();

const els = {
  saveState: document.getElementById("saveState"),
  table: document.getElementById("healthTable"),
  tableDrawer: document.getElementById("tableDrawer"),
  chartDrawer: document.getElementById("chartDrawer"),
  chartSvg: document.getElementById("chartSvg"),
  clockTime: document.getElementById("clockTime"),
  clockDate: document.getElementById("clockDate"),
  stopwatchTime: document.getElementById("stopwatchTime"),
  timerTime: document.getElementById("timerTime"),
  minutesInput: document.getElementById("minutesInput"),
  secondsInput: document.getElementById("secondsInput"),
  fileInput: document.getElementById("fileInput"),
  actionMenu: document.getElementById("actionMenu"),
  moreToggle: document.getElementById("moreToggle"),
  syncButton: document.getElementById("syncButton"),
  syncState: document.getElementById("syncState"),
  themeToggle: document.getElementById("themeToggle"),
  stopwatchToggle: document.getElementById("stopwatchToggle"),
  timerToggle: document.getElementById("timerToggle")
};

document.addEventListener("DOMContentLoaded", init);

async function init() {
  applyStoredTheme();
  bindControls();
  await completeOneDriveSignIn();
  renderSyncButton();
  documentData = loadDocument();
  rows = documentData.rows;
  renderTable();
  updateChart();
  updateClock();
  updateStopwatch();
  updateTimer();
  setInterval(updateClock, 1000);

  if (isSyncConfigured() && readTokens() && navigator.onLine) {
    window.setTimeout(syncNow, 800);
  }

  if ("serviceWorker" in navigator && location.protocol.startsWith("http")) {
    navigator.serviceWorker.register("service-worker.js").catch(() => {});
  }
}

function bindControls() {
  document.querySelectorAll(".segment").forEach((button) => {
    button.addEventListener("click", () => setMode(button.dataset.mode));
  });

  els.themeToggle.addEventListener("click", toggleTheme);
  els.moreToggle.addEventListener("click", toggleActionMenu);
  els.actionMenu.addEventListener("click", closeActionMenuAfterAction);
  document.addEventListener("click", closeActionMenuFromOutside);

  document.getElementById("tableToggle").addEventListener("click", () => {
    toggleDrawer(els.tableDrawer, document.getElementById("tableToggle"), "Afficher le tableau", "Masquer le tableau");
  });

  document.getElementById("chartToggle").addEventListener("click", () => {
    toggleDrawer(els.chartDrawer, document.getElementById("chartToggle"), "Afficher les courbes", "Masquer les courbes");
    updateChart();
  });

  document.getElementById("saveButton").addEventListener("click", syncNow);
  document.getElementById("refreshChartButton").addEventListener("click", updateChart);
  document.getElementById("importButton").addEventListener("click", () => els.fileInput.click());
  document.getElementById("exportButton").addEventListener("click", () => download("sante-data.json", JSON.stringify(currentDocument(), null, 2), "application/json"));
  els.fileInput.addEventListener("change", importFile);
  els.syncButton.addEventListener("click", syncNow);

  els.stopwatchToggle.addEventListener("click", toggleStopwatch);
  document.getElementById("stopwatchReset").addEventListener("click", resetStopwatch);

  els.timerToggle.addEventListener("click", toggleTimer);
  document.getElementById("timerReset").addEventListener("click", resetTimer);
}

function setMode(mode) {
  document.querySelectorAll(".segment").forEach((button) => {
    button.classList.toggle("active", button.dataset.mode === mode);
  });
  document.querySelectorAll(".mode-panel").forEach((panel) => panel.classList.remove("active"));
  document.getElementById(`${mode}Panel`).classList.add("active");
}

function toggleDrawer(drawer, button, showLabel, hideLabel) {
  const isOpen = drawer.classList.toggle("open");
  button.setAttribute("aria-label", isOpen ? hideLabel : showLabel);
  button.classList.toggle("active", isOpen);
}

function toggleActionMenu(event) {
  event.stopPropagation();
  const isOpen = els.actionMenu.classList.toggle("open");
  els.moreToggle.classList.toggle("active", isOpen);
  els.moreToggle.setAttribute("aria-expanded", String(isOpen));
}

function closeActionMenuFromOutside(event) {
  if (!els.actionMenu.classList.contains("open")) return;
  if (event.target.closest(".menu-wrap")) return;

  closeActionMenu();
}

function closeActionMenuAfterAction(event) {
  if (event.target.closest("button")) {
    window.setTimeout(closeActionMenu, 0);
  }
}

function closeActionMenu() {
  els.actionMenu.classList.remove("open");
  els.moreToggle.classList.remove("active");
  els.moreToggle.setAttribute("aria-expanded", "false");
}

function applyStoredTheme() {
  const useNight = localStorage.getItem(THEME_KEY) === "night";
  document.body.classList.toggle("night", useNight);
  renderThemeButton();
}

function toggleTheme() {
  const useNight = !document.body.classList.contains("night");
  document.body.classList.toggle("night", useNight);
  localStorage.setItem(THEME_KEY, useNight ? "night" : "day");
  renderThemeButton();
}

function renderThemeButton() {
  const useNight = document.body.classList.contains("night");
  els.themeToggle.innerHTML = useNight ? SUN_ICON : MOON_ICON;
  els.themeToggle.setAttribute("aria-label", useNight ? "Désactiver le mode nuit" : "Activer le mode nuit");
  els.themeToggle.title = useNight ? "Mode jour" : "Mode nuit";
}

function makeDefaultRows() {
  const seed = [
    ["Date", "SYS", "DIA", "BPM", "Poids(kg)", "Sommeil(h)"],
    ["02/06/26", "121", "78", "75", "75", "8"],
    ["03/06/26", "130", "85", "83", "75", "10"]
  ];
  return normaliseRows(seed);
}

function loadDocument() {
  const stored = localStorage.getItem(LOCAL_KEY);
  if (stored) {
    try {
      const parsed = JSON.parse(stored);
      return normaliseDocument(parsed);
    } catch {
      showStatus("Stockage local illisible");
    }
  }

  const legacyRows = localStorage.getItem(LEGACY_ROWS_KEY);
  if (legacyRows) {
    try {
      const migrated = makeDocument(JSON.parse(legacyRows));
      saveDocument(migrated, false);
      localStorage.removeItem(LEGACY_ROWS_KEY);
      showStatus("Donnees migrees en JSON");
      return migrated;
    } catch {
      showStatus("Anciennes donnees illisibles");
    }
  }

  const initial = makeDocument(makeDefaultRows());
  saveDocument(initial, false);
  return initial;
}

function makeDocument(inputRows, touched = true) {
  const now = new Date().toISOString();
  return {
    version: 1,
    updatedAt: touched ? now : "2026-06-03T00:00:00.000Z",
    deviceId: getDeviceId(),
    rows: normaliseRows(inputRows)
  };
}

function normaliseDocument(input) {
  if (Array.isArray(input)) {
    return makeDocument(input);
  }

  const doc = input && typeof input === "object" ? input : {};
  return {
    version: 1,
    updatedAt: typeof doc.updatedAt === "string" ? doc.updatedAt : new Date().toISOString(),
    deviceId: typeof doc.deviceId === "string" ? doc.deviceId : getDeviceId(),
    rows: normaliseRows(doc.rows)
  };
}

function currentDocument(touch = false) {
  if (touch) {
    documentData.updatedAt = new Date().toISOString();
    documentData.deviceId = getDeviceId();
  }

  documentData.version = 1;
  documentData.rows = normaliseRows(rows);
  rows = documentData.rows;
  return documentData;
}

function saveDocument(doc = currentDocument(), notify = true) {
  documentData = normaliseDocument(doc);
  rows = documentData.rows;
  localStorage.setItem(LOCAL_KEY, JSON.stringify(documentData));

  if (notify) {
    setSaveState("Sauvegarde locale");
  }
}

function getDeviceId() {
  let id = localStorage.getItem(DEVICE_KEY);
  if (!id) {
    id = crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
    localStorage.setItem(DEVICE_KEY, id);
  }
  return id;
}

function normaliseRows(input) {
  const clean = Array.isArray(input) ? input.slice(0, ROWS).map((row) => {
    const values = Array.isArray(row) ? row.slice(0, COLUMNS).map((value) => value == null ? "" : String(value)) : [];
    while (values.length < COLUMNS) values.push("");
    return values;
  }) : [];

  while (clean.length < ROWS) clean.push(Array(COLUMNS).fill(""));
  return clean;
}

function renderTable() {
  els.table.textContent = "";
  const colgroup = document.createElement("colgroup");

  for (let index = 0; index < COLUMNS; index += 1) {
    const col = document.createElement("col");
    col.dataset.col = index;
    colgroup.appendChild(col);
  }

  els.table.appendChild(colgroup);

  rows.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");

    row.forEach((value, colIndex) => {
      const td = document.createElement("td");
      const input = document.createElement("input");
      input.value = value;
      input.autocomplete = "off";
      input.dataset.row = rowIndex;
      input.dataset.col = colIndex;
      input.setAttribute("aria-label", `Ligne ${rowIndex + 1}, colonne ${colIndex + 1}`);
      input.addEventListener("input", handleCellInput);
      td.appendChild(input);
      tr.appendChild(td);
    });

    els.table.appendChild(tr);
  });

  updateTableColumnWidths();
}

function handleCellInput(event) {
  const row = Number(event.target.dataset.row);
  const col = Number(event.target.dataset.col);
  rows[row][col] = event.target.value;
  setSaveState("Modifié");
  scheduleSave();
  updateTableColumnWidths();
  updateChart();
}

function updateTableColumnWidths() {
  const sizes = Array(COLUMNS).fill(4);

  rows.forEach((row) => {
    row.forEach((value, index) => {
      const length = Array.from(String(value || "")).length;
      sizes[index] = Math.max(sizes[index], Math.min(length + 1, 12));
    });
  });

  const weights = sizes.map((size) => Math.max(4, Math.min(size, 12)));
  const total = weights.reduce((sum, value) => sum + value, 0);

  els.table.querySelectorAll("col").forEach((col, index) => {
    col.style.width = `${(weights[index] / total) * 100}%`;
  });
}

function scheduleSave() {
  window.clearTimeout(saveTimer);
  saveTimer = window.setTimeout(saveRowsNow, 700);
}

function setSaveState(message) {
  els.saveState.textContent = message;
}

async function saveRowsNow() {
  saveDocument(currentDocument(true));
  showStatus("Sauvegarde locale");
}

function showStatus(message, sticky = false) {
  window.clearTimeout(statusTimer);
  els.syncState.textContent = message;
  els.syncState.classList.add("visible");

  if (!sticky) {
    statusTimer = window.setTimeout(() => {
      els.syncState.classList.remove("visible");
    }, 2400);
  }
}

function updateClock() {
  const now = new Date();
  els.clockTime.textContent = new Intl.DateTimeFormat("fr-FR", {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit"
  }).format(now);
  els.clockDate.textContent = new Intl.DateTimeFormat("fr-FR", {
    weekday: "long",
    day: "numeric",
    month: "long"
  }).format(now);
}

function startStopwatch() {
  if (stopwatchRunning) return;
  stopwatchRunning = true;
  stopwatchStart = performance.now() - stopwatchElapsed;
  renderStopwatchButton();
}

function pauseStopwatch() {
  if (!stopwatchRunning) return;
  stopwatchRunning = false;
  stopwatchElapsed = performance.now() - stopwatchStart;
  renderStopwatchButton();
}

function toggleStopwatch() {
  if (stopwatchRunning) {
    pauseStopwatch();
  } else {
    startStopwatch();
  }
}

function resetStopwatch() {
  stopwatchRunning = false;
  stopwatchElapsed = 0;
  els.stopwatchTime.textContent = "00:00:00.0";
  renderStopwatchButton();
}

function updateStopwatch() {
  if (stopwatchRunning) {
    stopwatchElapsed = performance.now() - stopwatchStart;
  }
  els.stopwatchTime.textContent = formatStopwatch(stopwatchElapsed / 1000);
  requestAnimationFrame(updateStopwatch);
}

function formatStopwatch(seconds) {
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  const sec = Math.floor(seconds % 60);
  const decimal = Math.floor((seconds % 1) * 10);
  return `${pad(hours)}:${pad(minutes)}:${pad(sec)}.${decimal}`;
}

function renderStopwatchButton() {
  els.stopwatchToggle.innerHTML = stopwatchRunning ? PAUSE_ICON : PLAY_ICON;
  els.stopwatchToggle.classList.toggle("running", stopwatchRunning);
  els.stopwatchToggle.setAttribute("aria-label", stopwatchRunning ? "Pause" : "Démarrer");
  els.stopwatchToggle.title = stopwatchRunning ? "Pause" : "Démarrer";
}

function startTimer() {
  if (timerRunning) return;

  if (timerRemaining <= 0) {
    const minutes = Math.max(0, Number.parseInt(els.minutesInput.value || "0", 10));
    const seconds = Math.max(0, Number.parseInt(els.secondsInput.value || "0", 10));
    timerRemaining = minutes * 60 + Math.min(seconds, 59);
  }

  if (timerRemaining > 0) {
    timerRunning = true;
    timerLastTick = performance.now();
    renderTimerButton();
  }
}

function pauseTimer() {
  timerRunning = false;
  renderTimerButton();
}

function toggleTimer() {
  if (timerRunning) {
    pauseTimer();
  } else {
    startTimer();
  }
}

function resetTimer() {
  timerRunning = false;
  timerRemaining = 0;
  els.timerTime.textContent = "00:00";
  renderTimerButton();
}

function updateTimer() {
  if (timerRunning) {
    const now = performance.now();
    const delta = Math.floor((now - timerLastTick) / 1000);

    if (delta > 0) {
      timerRemaining = Math.max(0, timerRemaining - delta);
      timerLastTick += delta * 1000;
    }

    if (timerRemaining === 0) {
      timerRunning = false;
      renderTimerButton();
      navigator.vibrate?.(140);
    }
  }

  els.timerTime.textContent = formatTimer(timerRemaining);
  requestAnimationFrame(updateTimer);
}

function formatTimer(seconds) {
  return `${pad(Math.floor(seconds / 60))}:${pad(seconds % 60)}`;
}

function renderTimerButton() {
  els.timerToggle.innerHTML = timerRunning ? PAUSE_ICON : PLAY_ICON;
  els.timerToggle.classList.toggle("running", timerRunning);
  els.timerToggle.setAttribute("aria-label", timerRunning ? "Pause" : "Démarrer");
  els.timerToggle.title = timerRunning ? "Pause" : "Démarrer";
}

function pad(value) {
  return String(value).padStart(2, "0");
}

function updateChart() {
  const data = collectSeries();
  const svg = els.chartSvg;
  svg.textContent = "";

  if (!data.series.length) {
    addText(svg, 380, 180, "Aucune donnee valide", "empty-chart", "middle");
    return;
  }

  renderLegend(svg, data.series);

  const visibleSeries = data.series.filter((series) => !hiddenSeries.has(series.key));
  const visiblePoints = visibleSeries.flatMap((series) => series.values.filter((value) => value != null));

  if (!visibleSeries.length || !visiblePoints.length) {
    addText(svg, 380, 190, "Aucune courbe affichée", "empty-chart", "middle");
    return;
  }

  const margin = { top: 34, right: 36, bottom: 64, left: 58 };
  const width = 760;
  const height = 360;
  const plotWidth = width - margin.left - margin.right;
  const plotHeight = height - margin.top - margin.bottom;
  const minY = Math.min(...visiblePoints);
  const maxY = Math.max(...visiblePoints);
  const yPad = maxY === minY ? 1 : (maxY - minY) * 0.12;
  const yMin = minY - yPad;
  const yMax = maxY + yPad;
  const xCount = Math.max(1, data.x.length - 1);
  const xFor = (index) => margin.left + (plotWidth * index) / xCount;
  const yFor = (value) => margin.top + plotHeight - ((value - yMin) / (yMax - yMin)) * plotHeight;

  for (let i = 0; i <= 4; i += 1) {
    const y = margin.top + (plotHeight * i) / 4;
    addLine(svg, margin.left, y, width - margin.right, y, "grid-line");
    const label = yMax - ((yMax - yMin) * i) / 4;
    addText(svg, margin.left - 12, y + 4, trimNumber(label), "chart-label", "end");
  }

  addLine(svg, margin.left, margin.top, margin.left, height - margin.bottom, "axis");
  addLine(svg, margin.left, height - margin.bottom, width - margin.right, height - margin.bottom, "axis");

  data.x.forEach((label, index) => {
    if (index === 0 || index === data.x.length - 1 || index % Math.ceil(data.x.length / 5) === 0) {
      addText(svg, xFor(index), height - 36, label, "chart-label", "middle");
    }
  });

  visibleSeries.forEach((series) => {
    const color = COLORS[series.colorIndex % COLORS.length];
    let path = "";
    series.values.forEach((value, index) => {
      if (value == null) return;
      const x = xFor(index);
      const y = yFor(value);
      path += path ? ` L ${x} ${y}` : `M ${x} ${y}`;
    });

    if (!path) return;
    addPath(svg, path, color);
    series.values.forEach((value, index) => {
      if (value != null) addDot(svg, xFor(index), yFor(value), color);
    });
  });
}

function renderLegend(svg, seriesList) {
  seriesList.forEach((series) => {
    const color = COLORS[series.colorIndex % COLORS.length];
    const legendX = 58 + series.colorIndex * 138;
    const hidden = hiddenSeries.has(series.key);
    const dot = addDot(svg, legendX, 22, hidden ? "transparent" : color, 5);
    dot.setAttribute("stroke", color);
    dot.setAttribute("class", "legend-dot");
    dot.setAttribute("tabindex", "0");
    dot.setAttribute("role", "button");
    dot.setAttribute("aria-label", hidden ? `Afficher ${series.name}` : `Masquer ${series.name}`);
    dot.addEventListener("click", () => toggleSeries(series.key));
    dot.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        toggleSeries(series.key);
      }
    });

    const label = addText(svg, legendX + 12, 26, series.name, hidden ? "chart-label legend-label muted" : "chart-label legend-label", "start");
    label.addEventListener("click", () => toggleSeries(series.key));
  });
}

function toggleSeries(key) {
  if (hiddenSeries.has(key)) {
    hiddenSeries.delete(key);
  } else {
    hiddenSeries.add(key);
  }
  updateChart();
}

function collectSeries() {
  const headers = rows[0].slice(1);
  const x = [];
  const series = headers.map((name, index) => {
    const cleanName = name.trim();
    return { key: `${index}:${cleanName}`, colorIndex: index, name: cleanName, values: [] };
  });

  rows.slice(1).forEach((row) => {
    const label = row[0].trim();
    if (!label) return;

    x.push(label);
    row.slice(1).forEach((raw, index) => {
      const value = parseNumber(raw);
      series[index].values.push(value);
    });
  });

  return {
    x,
    series: series.filter((item) => item.name && item.values.some((value) => value != null))
  };
}

function parseNumber(value) {
  const clean = String(value || "").trim().replace(",", ".");
  if (!clean) return null;
  const parsed = Number(clean);
  return Number.isFinite(parsed) ? parsed : null;
}

function addLine(svg, x1, y1, x2, y2, className) {
  const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
  line.setAttribute("x1", x1);
  line.setAttribute("y1", y1);
  line.setAttribute("x2", x2);
  line.setAttribute("y2", y2);
  line.setAttribute("class", className);
  svg.appendChild(line);
}

function addPath(svg, d, color) {
  const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
  path.setAttribute("d", d);
  path.setAttribute("class", "series-line");
  path.setAttribute("stroke", color);
  svg.appendChild(path);
}

function addDot(svg, cx, cy, color, r = 6) {
  const dot = document.createElementNS("http://www.w3.org/2000/svg", "circle");
  dot.setAttribute("cx", cx);
  dot.setAttribute("cy", cy);
  dot.setAttribute("r", r);
  dot.setAttribute("class", "series-dot");
  dot.setAttribute("fill", color);
  svg.appendChild(dot);
  return dot;
}

function addText(svg, x, y, content, className, anchor) {
  const text = document.createElementNS("http://www.w3.org/2000/svg", "text");
  text.setAttribute("x", x);
  text.setAttribute("y", y);
  text.setAttribute("class", className);
  text.setAttribute("text-anchor", anchor);
  text.textContent = content;
  svg.appendChild(text);
  return text;
}

function trimNumber(value) {
  return Number.isInteger(value) ? String(value) : value.toFixed(1);
}

async function syncNow() {
  if (!isSyncConfigured()) {
    showStatus("Ajoutez le client ID Microsoft dans sync-config.js", true);
    return;
  }

  if (!location.protocol.startsWith("https") && location.hostname !== "localhost" && location.hostname !== "127.0.0.1") {
    showStatus("OneDrive demande une page HTTPS");
    return;
  }

  try {
    showStatus("Connexion OneDrive", true);
    const token = await getAccessToken();
    if (!token) return;

    showStatus("Synchronisation", true);
    const localDoc = currentDocument();
    const remoteDoc = await downloadOneDriveDocument(token);

    if (!remoteDoc) {
      await uploadOneDriveDocument(token, localDoc);
      rememberSync(localDoc);
      showStatus("Envoye vers OneDrive");
      return;
    }

    const remote = normaliseDocument(remoteDoc);
    const lastSync = localStorage.getItem(LAST_SYNC_KEY);
    const localChanged = localDoc.updatedAt !== lastSync;
    const remoteChanged = remote.updatedAt !== lastSync;

    if (localChanged && remoteChanged && localDoc.updatedAt !== remote.updatedAt) {
      const useRemote = remote.updatedAt > localDoc.updatedAt
        ? window.confirm("OneDrive est plus recent. OK importe OneDrive, Annuler garde cet appareil.")
        : !window.confirm("Cet appareil est plus recent. OK garde cet appareil, Annuler importe OneDrive.");

      if (useRemote) {
        applyRemoteDocument(remote);
        rememberSync(remote);
        showStatus("OneDrive importe");
        return;
      }

      await uploadOneDriveDocument(token, localDoc);
      rememberSync(localDoc);
      showStatus("OneDrive mis a jour");
      return;
    }

    if (remote.updatedAt > localDoc.updatedAt) {
      applyRemoteDocument(remote);
      rememberSync(remote);
      showStatus("OneDrive importe");
      return;
    }

    if (localDoc.updatedAt > remote.updatedAt) {
      await uploadOneDriveDocument(token, localDoc);
      rememberSync(localDoc);
      showStatus("OneDrive mis a jour");
      return;
    }

    rememberSync(localDoc);
    showStatus("Deja synchronise");
  } catch (error) {
    showStatus(error.message || "Synchronisation impossible", true);
  }
}

function isSyncConfigured() {
  return Boolean(SYNC_CONFIG.microsoftClientId && SYNC_CONFIG.microsoftClientId !== "YOUR_CLIENT_ID");
}

function renderSyncButton() {
  els.syncButton.classList.toggle("active", Boolean(readTokens()));
}

async function completeOneDriveSignIn() {
  const params = new URLSearchParams(location.search);
  const authError = params.get("error_description") || params.get("error");
  const code = params.get("code");
  const state = params.get("state");

  if (authError) {
    showStatus(decodeURIComponent(authError), true);
    history.replaceState({}, document.title, getRedirectUri());
    return;
  }

  if (!code) return;

  const expectedState = sessionStorage.getItem("sante.oauth.state");
  const verifier = sessionStorage.getItem("sante.oauth.verifier");

  if (!expectedState || !verifier || state !== expectedState) {
    showStatus("Connexion OneDrive refusee", true);
    return;
  }

  try {
    const tokens = await requestTokens({
      grant_type: "authorization_code",
      code,
      code_verifier: verifier,
      redirect_uri: getRedirectUri()
    });
    writeTokens(tokens);
    sessionStorage.removeItem("sante.oauth.state");
    sessionStorage.removeItem("sante.oauth.verifier");
    history.replaceState({}, document.title, getRedirectUri());
    showStatus("Connecte a OneDrive");
    window.setTimeout(syncNow, 250);
  } catch {
    showStatus("Connexion OneDrive impossible", true);
  }
}

async function getAccessToken() {
  let tokens = readTokens();

  if (tokens && tokens.access_token && tokens.expiresAt > Date.now() + 60000) {
    return tokens.access_token;
  }

  if (tokens?.refresh_token) {
    try {
      tokens = await requestTokens({
        grant_type: "refresh_token",
        refresh_token: tokens.refresh_token,
        redirect_uri: getRedirectUri()
      });
      writeTokens(tokens);
      renderSyncButton();
      return tokens.access_token;
    } catch {
      localStorage.removeItem(TOKEN_KEY);
      renderSyncButton();
    }
  }

  await startOneDriveSignIn();
  return null;
}

async function startOneDriveSignIn() {
  if (!crypto.subtle) {
    showStatus("Connexion securisee requise");
    return;
  }

  const verifier = makeCodeVerifier();
  const state = makeCodeVerifier();
  const challenge = await makeCodeChallenge(verifier);
  sessionStorage.setItem("sante.oauth.verifier", verifier);
  sessionStorage.setItem("sante.oauth.state", state);

  const authUrl = new URL(`${getAuthority()}/authorize`);
  authUrl.search = new URLSearchParams({
    client_id: SYNC_CONFIG.microsoftClientId,
    response_type: "code",
    redirect_uri: getRedirectUri(),
    response_mode: "query",
    scope: GRAPH_SCOPES,
    state,
    code_challenge: challenge,
    code_challenge_method: "S256"
  }).toString();

  location.assign(authUrl.toString());
}

async function requestTokens(values) {
  const body = new URLSearchParams({
    client_id: SYNC_CONFIG.microsoftClientId,
    scope: GRAPH_SCOPES,
    ...values
  });

  const response = await fetch(`${getAuthority()}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body
  });

  if (!response.ok) {
    throw new Error("Token Microsoft refuse");
  }

  const tokens = await response.json();
  tokens.expiresAt = Date.now() + (tokens.expires_in || 3600) * 1000;
  return tokens;
}

function getAuthority() {
  const tenant = SYNC_CONFIG.authTenant || "consumers";
  return `https://login.microsoftonline.com/${tenant}/oauth2/v2.0`;
}

function readTokens() {
  const stored = localStorage.getItem(TOKEN_KEY);
  if (!stored) return null;

  try {
    return JSON.parse(stored);
  } catch {
    return null;
  }
}

function writeTokens(tokens) {
  localStorage.setItem(TOKEN_KEY, JSON.stringify(tokens));
  renderSyncButton();
}

async function downloadOneDriveDocument(token) {
  const response = await fetch(oneDriveContentUrl(), {
    headers: { Authorization: `Bearer ${token}` },
    cache: "no-store"
  });

  if (response.status === 404) return null;
  if (response.status === 401) throw new Error("Connexion OneDrive expiree");
  if (!response.ok) throw new Error("Lecture OneDrive impossible");

  return response.json();
}

async function uploadOneDriveDocument(token, doc) {
  const response = await fetch(oneDriveContentUrl(), {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(doc, null, 2)
  });

  if (!response.ok) {
    throw new Error("Ecriture OneDrive impossible");
  }
}

function oneDriveContentUrl() {
  const fileName = SYNC_CONFIG.dataFileName || "data.json";
  return `${GRAPH_ROOT}/me/drive/special/approot:/${encodeURIComponent(fileName)}:/content`;
}

function applyRemoteDocument(doc) {
  documentData = normaliseDocument(doc);
  rows = documentData.rows;
  saveDocument(documentData, false);
  renderTable();
  updateChart();
}

function rememberSync(doc) {
  localStorage.setItem(LAST_SYNC_KEY, doc.updatedAt);
}

function getRedirectUri() {
  return `${location.origin}${location.pathname}`;
}

function makeCodeVerifier() {
  const bytes = new Uint8Array(32);
  crypto.getRandomValues(bytes);
  return base64Url(bytes);
}

async function makeCodeChallenge(verifier) {
  const bytes = new TextEncoder().encode(verifier);
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  return base64Url(new Uint8Array(digest));
}

function base64Url(bytes) {
  let text = "";
  bytes.forEach((byte) => {
    text += String.fromCharCode(byte);
  });
  return btoa(text).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

async function importFile(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  const text = await file.text();
  if (file.name.toLowerCase().endsWith(".json")) {
    const payload = JSON.parse(text);
    rows = normaliseRows(payload.rows || payload);
  } else {
    rows = normaliseRows(parseCsv(text));
  }

  renderTable();
  updateChart();
  await saveRowsNow();
  els.fileInput.value = "";
}

function parseCsv(text) {
  const result = [];
  let row = [];
  let cell = "";
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];

    if (char === '"' && quoted && next === '"') {
      cell += '"';
      index += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (char === "," && !quoted) {
      row.push(cell);
      cell = "";
    } else if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(cell);
      result.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }

  if (cell || row.length) {
    row.push(cell);
    result.push(row);
  }

  return result;
}

function toCsv(data) {
  return data.map((row) => row.map((cell) => {
    const value = String(cell ?? "");
    return /[",\n\r]/.test(value) ? `"${value.replaceAll('"', '""')}"` : value;
  }).join(",")).join("\n");
}

function download(filename, content, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}
