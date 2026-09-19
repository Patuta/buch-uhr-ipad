const { PublicClientApplication, InteractionRequiredAuthError } = window.msal;

const PROJECT_FILE = "Buch-Uhr.project.json";
const FILES_DIR = "Uhr";
const LEGACY_FILES_DIR = "Dateien";
const REFERENCE_DIR = "Stehsatz";
const TRASH_DIR = "Papierkorb";
const CHARS_PER_PAGE = 1800;
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const SCOPES = ["User.Read", "Files.ReadWrite"];

const state = {
  config: {
    clientId: localStorage.getItem("buchuhr.clientId") || "",
    tenant: localStorage.getItem("buchuhr.tenant") || "common",
    root: localStorage.getItem("buchuhr.root") || "Buch-Uhr",
    folder: "",
  },
  msal: null,
  account: null,
  project: null,
  projectETag: "",
  projectItem: null,
  folderItems: [],
  expandedExplorerFolders: (() => {
    try {
      return new Set(JSON.parse(localStorage.getItem("buchuhr.explorerFolders") || "[]"));
    } catch {
      return new Set();
    }
  })(),
  history: [],
  future: [],
  selectedDocId: null,
  editorDoc: null,
  sidebarsVisible: localStorage.getItem("buchuhr.sidebars") !== "false",
  dragging: null,
  selectedDocId: null,
  selectedFileId: null,
  selectedReferencePath: null,
  actionMode: "",
  actionMinute: 0,
  viewZoom: Number(localStorage.getItem("buchuhr.viewZoom") || 1),
  viewPanX: Number(localStorage.getItem("buchuhr.viewPanX") || 0),
  viewPanY: Number(localStorage.getItem("buchuhr.viewPanY") || 0),
  panning: null,
  lastDocTap: { id: null, time: 0 },
  lastDocClick: { id: null, time: 0 },
  suppressClickUntil: 0,
  pointerAction: null,
  lastCanvasFileActivation: { id: null, time: 0 },
  lastRasterTap: { second: null, time: 0 },
  lastSidebarBlankTap: { area: null, time: 0 },
  pinch: null,
  lastTouchEndTime: 0,
  lastSidebarFileTap: { key: null, time: 0 },
  previewRequestId: 0,
  previewEditors: { left: null, right: null },
  lastError: null,
  swipeAction: null,
  separatorDrag: null,
  suppressReferenceClickUntil: 0,
};


function clearStaleMsalInteraction() {
  const preserved = {
    clientId: localStorage.getItem("buchuhr.clientId"),
    tenant: localStorage.getItem("buchuhr.tenant"),
    root: localStorage.getItem("buchuhr.root"),
    sidebars: localStorage.getItem("buchuhr.sidebars"),
  };

  for (let i = localStorage.length - 1; i >= 0; i--) {
    const key = localStorage.key(i);
    if (key && (key.startsWith("msal.") || key.includes("interaction.status"))) {
      localStorage.removeItem(key);
    }
  }

  if (preserved.clientId !== null) localStorage.setItem("buchuhr.clientId", preserved.clientId);
  if (preserved.tenant !== null) localStorage.setItem("buchuhr.tenant", preserved.tenant);
  if (preserved.root !== null) localStorage.setItem("buchuhr.root", preserved.root);
  if (preserved.sidebars !== null) localStorage.setItem("buchuhr.sidebars", preserved.sidebars);
}

function isInteractionInProgress(error) {
  const text = `${error?.errorCode || ""} ${error?.message || error || ""}`.toLowerCase();
  return text.includes("interaction_in_progress");
}

const el = Object.fromEntries([
  "appLayout", "sidebarToggle", "projectName", "syncStatus", "syncBtn", "settingsBtn",
  "settingsDialog", "settingsForm", "clientIdInput", "tenantInput", "folderInput",
  "authDiagDialog", "authDiagText", "authDiagCancelBtn", "authDiagContinueBtn",
  "projectsBtn", "projectDialog", "projectList", "newProjectBtn", "openProjectBtn", "closeProjectBtn",
  "saveSettingsBtn", "clockCanvas", "clockLayer", "progressLayer", "rasterTitleLayer",
  "documentLayer", "referenceList", "referenceAddBtn", "fileList", "refreshFilesBtn", "emptyHint",
  "newTextBtn", "renameBtn", "deleteBtn", "zoomOutBtn", "zoomInBtn", "fitBtn", "undoBtn", "redoBtn", "editorDialog", "editorTitle", "editorText", "editorMeta",
  "docxMessage", "saveEditorBtn", "openExternalBtn", "closeEditorBtn", "toast", "contextMenu", "actionDialog", "actionDialogTitle", "actionDialogLabel", "actionDialogInput", "actionDialogText", "newTextTitleLabel", "newTextTitleInput", "actionDialogSaveBtn", "actionColorFields", "backgroundColorInput", "clockColorInput", "progressColorInput",
  "titleColorDialog", "titleColorRows", "titleColorPlus", "titleColorSave",
  "leftPreview", "leftPreviewTitle", "leftPreviewBody", "leftPreviewClose",
  "rightPreview", "rightPreviewTitle", "rightPreviewBody", "rightPreviewClose",
  "errorDialog", "errorDetailsText", "errorDetailsClose"
].map(id => [id, document.getElementById(id)]));

function toast(message, timeout = 2800) {
  el.toast.textContent = message;
  el.toast.classList.remove("hidden");
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => el.toast.classList.add("hidden"), timeout);
}

function setStatus(text, kind = "offline") {
  el.syncStatus.textContent = text;
  el.syncStatus.className = `status ${kind}`;
  el.syncStatus.classList.toggle("clickable", kind === "error" && Boolean(state.lastError));
  el.syncStatus.title = kind === "error" && state.lastError
    ? "Fehlerdetails anzeigen"
    : "";
}

function normalizeConfiguredPath(path = "") {
  return String(path)
    .trim()
    .replace(/\\/g, "/")
    .replace(/\/{2,}/g, "/")
    .replace(/^\/+|\/+$/g, "");
}

function errorDetails(error, context = "") {
  const message = error?.message || String(error || "Unbekannter Fehler");
  const parts = [];
  if (context) parts.push(context);
  parts.push(message);
  if (error?.status) parts.push(`HTTP-Status: ${error.status}`);
  if (state.config.root) parts.push(`OneDrive-Projektstamm: ${state.config.root}`);
  return parts.join("\n\n");
}

function rememberError(error, context = "") {
  state.lastError = errorDetails(error, context);
  setStatus("Fehler", "error");
  el.errorDetailsText.textContent = state.lastError;
}

function showErrorDetails() {
  if (!state.lastError) return;
  el.errorDetailsText.textContent = state.lastError;
  if (!el.errorDialog.open) el.errorDialog.showModal();
}

function clearRememberedError() {
  state.lastError = null;
  el.syncStatus.classList.remove("clickable");
  el.syncStatus.title = "";
}

function encodeGraphPath(path) {
  return normalizeConfiguredPath(path).split("/").filter(Boolean).map(encodeURIComponent).join("/");
}

function projectPath(relative = "") {
  const base = normalizeConfiguredPath(state.config.folder);
  return [base, normalizeRelative(relative)].filter(Boolean).join("/");
}

function projectsRootPath(relative = "") {
  const base = normalizeConfiguredPath(state.config.root) || "Buch-Uhr";
  return [base, normalizeRelative(relative)].filter(Boolean).join("/");
}

function normalizeRelative(path = "") {
  return String(path).replace(/\\/g, "/").replace(/^\/+/, "");
}

function basenameAny(path = "") {
  return String(path).replace(/\\/g, "/").split("/").filter(Boolean).pop() || "";
}

function stem(name) {
  return name.replace(/\.[^.]+$/, "");
}

function suffix(name) {
  const m = name.match(/(\.[^.]+)$/);
  return m ? m[1].toLowerCase() : "";
}

function projectCopyRelative(doc) {
  const raw = normalizeRelative(doc.project_path);
  if (raw.toLocaleLowerCase("de").startsWith(`${FILES_DIR.toLocaleLowerCase("de")}/`)) return raw;
  const name = basenameAny(raw);
  return `${FILES_DIR}/${name}`;
}

function safeProjectStem(value = "Text") {
  const cleaned = String(value || "Text")
    .replace(/[<>:"/\\|?*\x00-\x1f]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/[. ]+$/g, "");
  return cleaned || "Text";
}

function stripClockPrefix(value = "") {
  return String(value)
    .replace(/^\d{2}\.\d{2}\s+/, "")
    .replace(/^\d{2}_\d{2}\s*[–-]\s*/, "")
    .trim() || "Text";
}

function clockPrefix(second = 0) {
  const value = ((Math.round(Number(second) || 0) % 3600) + 3600) % 3600;
  return `${String(Math.floor(value / 60)).padStart(2, "0")}.${String(value % 60).padStart(2, "0")}`;
}

function clockFilename(title, second = 0) {
  return `${clockPrefix(second)} ${safeProjectStem(title)}.txt`;
}

function plainFilename(title) {
  return `${safeProjectStem(stripClockPrefix(title))}.txt`;
}

async function listFolderByRelative(relative = "") {
  const absolute = encodeGraphPath(projectPath(relative));
  const result = await graph(`/me/drive/root:/${absolute}:/children?$select=id,name,webUrl,file,folder,parentReference,lastModifiedDateTime&$top=500`);
  return result.value || [];
}

async function uniqueNameInFolder(relative, desiredName, ignoreId = "") {
  const used = new Set(
    (await listFolderByRelative(relative))
      .filter(item => item.id !== ignoreId)
      .map(item => String(item.name || "").toLocaleLowerCase("de"))
  );
  if (!used.has(desiredName.toLocaleLowerCase("de"))) return desiredName;

  const dot = desiredName.lastIndexOf(".");
  const base = dot > 0 ? desiredName.slice(0, dot) : desiredName;
  const ext = dot > 0 ? desiredName.slice(dot) : "";
  let number = 2;
  while (used.has(`${base} (${number})${ext}`.toLocaleLowerCase("de"))) number += 1;
  return `${base} (${number})${ext}`;
}

async function folderItemForRelative(relative = "") {
  if (!relative) return graphItemByAbsolutePath(projectPath());
  await ensureFolderPath(projectPath(relative));
  return graphItemByAbsolutePath(projectPath(relative));
}

async function moveDriveItem(item, targetRelative, desiredName) {
  const parent = await folderItemForRelative(targetRelative);
  const name = await uniqueNameInFolder(targetRelative, desiredName, item.id);
  return graph(`/me/drive/items/${item.id}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ parentReference: { id: parent.id }, name }),
  });
}

async function deleteDriveItem(item) {
  const accessToken = await token();
  const response = await fetch(`${GRAPH_BASE}/me/drive/items/${item.id}`, {
    method: "DELETE",
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!response.ok && response.status !== 204) {
    const body = await response.text();
    throw new Error(`Datei konnte nicht gelöscht werden.\n${response.status} ${response.statusText}\n${body}`);
  }
}

async function moveClockDocToFolder(doc, targetRelative) {
  const item = await getItemByPath(projectCopyRelative(doc));
  const moved = await moveDriveItem(item, targetRelative, plainFilename(doc.title));

  state.project.documents = (state.project.documents || []).filter(entry => entry.id !== doc.id);
  if (state.selectedDocId === doc.id) state.selectedDocId = null;

  if (targetRelative === REFERENCE_DIR) {
    state.project.reference_files ||= [];
    const raw = `${REFERENCE_DIR}/${moved.name}`;
    if (!state.project.reference_files.includes(raw)) state.project.reference_files.push(raw);
  }

  await saveProject();
  await loadFolderFiles();
  renderAll();
  return moved;
}

async function renameClockDocForState(doc, title = doc.title, second = doc.start_second) {
  const item = await getItemByPath(projectCopyRelative(doc));
  const name = await uniqueNameInFolder(FILES_DIR, clockFilename(title, second), item.id);
  const parent = await folderItemForRelative(FILES_DIR);
  const updated = await graph(`/me/drive/items/${item.id}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ parentReference: { id: parent.id }, name }),
  });
  doc.title = stripClockPrefix(stem(updated.name));
  doc.project_path = `${FILES_DIR}/${updated.name}`;
  doc.text_cache_path = doc.project_path;
  doc.original_path = "";
  doc.source_type = "project_text";
  doc.suffix = ".txt";
  doc.is_on_clock = true;
  doc.original_mtime_ns = 0;
  return updated;
}

async function attachProjectTxtToClock(relative, second = 0) {
  const normalized = normalizeRelative(relative);
  const item = await getItemByPath(normalized);
  if (item.folder || suffix(item.name) !== ".txt") throw new Error("Auf die Uhr können nur Projekt-TXT-Dateien gelegt werden.");

  const title = stripClockPrefix(stem(item.name));
  const name = await uniqueNameInFolder(FILES_DIR, clockFilename(title, second), item.id);
  const moved = await moveDriveItem(item, FILES_DIR, name);

  const { response } = await downloadByPath(`${FILES_DIR}/${moved.name}`);
  const content = await response.text();
  const doc = {
    id: crypto.randomUUID().replaceAll("-", ""),
    title,
    source_type: "project_text",
    original_path: "",
    project_path: `${FILES_DIR}/${moved.name}`,
    start_second: ((Math.round(second) % 3600) + 3600) % 3600,
    character_count: content.length,
    text_cache_path: `${FILES_DIR}/${moved.name}`,
    suffix: ".txt",
    is_on_clock: true,
    original_mtime_ns: 0
  };
  state.project.documents ||= [];
  state.project.documents.push(doc);

  state.project.reference_files = (state.project.reference_files || []).filter(raw =>
    isReferenceSeparator(raw) || normalizeRelative(raw).toLocaleLowerCase("de") !== normalized.toLocaleLowerCase("de")
  );

  await saveProject();
  await loadFolderFiles();
  renderAll();
  return doc;
}

function magneticSecond(raw, threshold = 9) {
  raw = ((Math.round(raw) % 3600) + 3600) % 3600;
  const nearest = Math.round(raw / 60) * 60 % 3600;
  const diff = Math.min(Math.abs(raw - nearest), 3600 - Math.abs(raw - nearest));
  return diff <= threshold ? nearest : raw;
}

function deepCopy(value) {
  return JSON.parse(JSON.stringify(value));
}

function pushHistory() {
  if (!state.project) return;
  state.history.push(deepCopy(state.project));
  if (state.history.length > 50) state.history.shift();
  state.future = [];
  updateHistoryButtons();
  applyViewZoom();
}

function updateHistoryButtons() {
  el.undoBtn.disabled = state.history.length === 0;
  el.redoBtn.disabled = state.future.length === 0;
}

async function undo() {
  if (!state.history.length || !state.project) return;
  state.future.push(deepCopy(state.project));
  state.project = state.history.pop();
  renderAll();
  updateHistoryButtons();
  await saveProject();
}

async function redo() {
  if (!state.future.length || !state.project) return;
  state.history.push(deepCopy(state.project));
  state.project = state.future.pop();
  renderAll();
  updateHistoryButtons();
  await saveProject();
}

async function ensureMsal() {
  if (!state.config.clientId) throw new Error("Bitte zuerst die Client-ID in den Einstellungen eintragen.");
  if (state.msal) return state.msal;

  state.msal = new PublicClientApplication({
    auth: {
      clientId: state.config.clientId,
      authority: `https://login.microsoftonline.com/${state.config.tenant || "common"}`,
      redirectUri: location.origin + location.pathname,
      postLogoutRedirectUri: location.origin + location.pathname,
    },
    cache: {
      cacheLocation: "localStorage",
    },
  });

  await state.msal.initialize();

  const redirectResult = await state.msal.handleRedirectPromise();
  if (redirectResult?.account) {
    state.account = redirectResult.account;
    state.msal.setActiveAccount(redirectResult.account);
  } else {
    const accounts = state.msal.getAllAccounts();
    state.account = state.msal.getActiveAccount() || accounts[0] || null;
    if (state.account) state.msal.setActiveAccount(state.account);
  }

  return state.msal;
}


function showAuthDiagnostics() {
  const tenant = state.config.tenant || "common";
  const redirectUri = location.origin + location.pathname;
  const authority = `https://login.microsoftonline.com/${tenant}`;
  const authorizeEndpoint = `${authority}/oauth2/v2.0/authorize`;
  const lines = [
    `Client-ID: ${state.config.clientId}`,
    `Mandant: ${tenant}`,
    `Authority: ${authority}`,
    `Authorize-Endpunkt: ${authorizeEndpoint}`,
    `Redirect-URI: ${redirectUri}`,
    `Redirect-Startseite: ${redirectUri}`,
    `Scopes: ${SCOPES.join(" ")}`,
    `Aktuelle Seite: ${location.href}`,
    `Origin: ${location.origin}`,
    `Pfad: ${location.pathname}`,
    `Browser: ${navigator.userAgent}`,
  ];
  el.authDiagText.textContent = lines.join("\n");
  return new Promise(resolve => {
    const finish = value => {
      el.authDiagDialog.removeEventListener("close", onClose);
      resolve(value);
    };
    const onClose = () => finish(el.authDiagDialog.returnValue === "default");
    el.authDiagDialog.addEventListener("close", onClose, { once: true });
    el.authDiagDialog.showModal();
  });
}

async function signIn() {
  const client = await ensureMsal();
  if (!state.account) {
    const continueToMicrosoft = await showAuthDiagnostics();
    if (!continueToMicrosoft) return null;
    try {
      await client.loginRedirect({
        scopes: SCOPES,
        prompt: "select_account",
        redirectStartPage: location.origin + location.pathname,
      });
    } catch (error) {
      if (isInteractionInProgress(error)) {
        clearStaleMsalInteraction();
        state.msal = null;
        state.account = null;
        location.replace(location.origin + location.pathname);
        return null;
      }
      throw error;
    }
    return null;
  }
  return state.account;
}

async function token() {
  const client = await ensureMsal();
  const account = await signIn();
  if (!account) return null;

  try {
    const result = await client.acquireTokenSilent({
      scopes: SCOPES,
      account,
    });
    return result.accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError || error?.errorCode) {
      await client.acquireTokenRedirect({
        scopes: SCOPES,
        account,
        redirectStartPage: location.origin + location.pathname,
      });
      return null;
    }
    throw error;
  }
}

async function graph(url, options = {}) {
  const accessToken = await token();
  if (!accessToken) throw new Error("Die Anmeldung wird abgeschlossen. Bitte warten.");
  const response = await fetch(`${GRAPH_BASE}${url}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(options.headers || {}),
    },
  });
  if (!response.ok) {
    const body = await response.text();
    const error = new Error(`Microsoft Graph: ${response.status} ${response.statusText}\n${body}`);
    error.status = response.status;
    throw error;
  }
  if (response.status === 204) return null;
  const type = response.headers.get("content-type") || "";
  return type.includes("application/json") ? response.json() : response.text();
}

async function getItemByPath(relative) {
  const path = encodeGraphPath(projectPath(relative));
  return graph(`/me/drive/root:/${path}?$select=id,name,eTag,lastModifiedDateTime,webUrl,@microsoft.graph.downloadUrl,file,folder`);
}

async function downloadByPath(relative) {
  const item = await getItemByPath(relative);

  const accessToken = await token();
  if (!accessToken) throw new Error("Die Anmeldung wird abgeschlossen. Bitte warten.");

  const path = encodeGraphPath(projectPath(relative));
  const response = await fetch(
    `${GRAPH_BASE}/me/drive/root:/${path}:/content`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
      redirect: "follow",
    }
  );

  if (!response.ok) {
    const body = await response.text();
    throw new Error(
      `Datei konnte nicht geladen werden: ${relative}\n` +
      `${response.status} ${response.statusText}\n${body}`
    );
  }

  return { item, response };
}

async function uploadByPath(relative, body, contentType, etag = "") {
  const path = encodeGraphPath(projectPath(relative));
  const headers = { "Content-Type": contentType };
  if (etag) headers["If-Match"] = etag;
  return graph(`/me/drive/root:/${path}:/content`, {
    method: "PUT",
    headers,
    body,
  });
}

async function graphItemByAbsolutePath(path) {
  const encoded = encodeGraphPath(path);
  return graph(`/me/drive/root:/${encoded}?$select=id,name,eTag,lastModifiedDateTime,webUrl,file,folder`);
}

async function ensureFolderPath(path) {
  const parts = normalizeRelative(path).split("/").filter(Boolean);
  let current = "";
  for (const part of parts) {
    const target = [current, part].filter(Boolean).join("/");
    try {
      await graphItemByAbsolutePath(target);
    } catch (error) {
      if (error.status !== 404) throw error;
      const parent = current ? `/me/drive/root:/${encodeGraphPath(current)}:/children` : "/me/drive/root/children";
      await graph(parent, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: part, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }),
      });
    }
    current = target;
  }
}

async function listProjects() {
  await ensureFolderPath(projectsRootPath());
  const root = encodeGraphPath(projectsRootPath());
  const result = await graph(`/me/drive/root:/${root}:/children?$select=id,name,file,folder,lastModifiedDateTime&$top=200`);
  const children = result.value || [];

  if (children.some(item => item.file && item.name === PROJECT_FILE)) {
    throw new Error(
      "Der eingetragene OneDrive-Projektstamm zeigt bereits direkt auf ein Buch-Uhr-Projekt. " +
      "Trage den übergeordneten Ordner ein, in dem die einzelnen Buch-Uhr-Projekte liegen. " +
      "Beispiel: Liegt das Projekt in …/Thomas Buch/Thomas Uhr, lautet der Projektstamm …/Thomas Buch."
    );
  }

  return children
    .filter(item => item.folder)
    .sort((a, b) => a.name.localeCompare(b.name, "de"));
}

function rememberProjectFolder(folder) {
  const root = state.config.root.replace(/^\/+|\/+$/g, "");
  const prefix = root ? `${root}/` : "";
  const projectName = folder.startsWith(prefix) ? folder.slice(prefix.length) : folder;
  const recent = JSON.parse(localStorage.getItem("buchuhr.recentProjects") || "[]")
    .filter(name => name !== projectName);
  recent.unshift(projectName);
  localStorage.setItem("buchuhr.recentProjects", JSON.stringify(recent.slice(0, 12)));
}

async function showProjectChooser() {
  if (!state.config.clientId) {
    openSettings();
    return;
  }
  await ensureMsal();
  await signIn();
  const projects = await listProjects();
  const recent = JSON.parse(localStorage.getItem("buchuhr.recentProjects") || "[]");
  const rank = new Map(recent.map((name, index) => [name, index]));
  projects.sort((a, b) => {
    const ar = rank.has(a.name) ? rank.get(a.name) : 9999;
    const br = rank.has(b.name) ? rank.get(b.name) : 9999;
    return ar - br || a.name.localeCompare(b.name, "de");
  });

  el.projectList.replaceChildren();
  for (const item of projects) {
    const option = document.createElement("option");
    option.value = item.name;
    option.textContent = rank.has(item.name) ? `Zuletzt: ${item.name}` : item.name;
    el.projectList.append(option);
  }
  if (projects.length) el.projectList.selectedIndex = 0;
  el.projectDialog.showModal();
}

async function openChosenProject() {
  const name = el.projectList.value;
  if (!name) return;
  state.config.folder = projectsRootPath(name);
  rememberProjectFolder(state.config.folder);
  el.projectDialog.close();
  await loadProject();
}

async function createProject() {
  const raw = prompt("Name des neuen Buch-Uhr-Projekts:", "");
  if (raw === null) return;
  const name = raw.trim().replace(/[<>:"/\\|?*\x00-\x1f]/g, " ").replace(/\s+/g, " ").replace(/[. ]+$/g, "");
  if (!name) return;

  const folder = projectsRootPath(name);
  try {
    await graphItemByAbsolutePath(folder);
    toast("Ein Projekt mit diesem Namen existiert bereits.", 5000);
    return;
  } catch (error) {
    if (error.status !== 404) throw error;
  }

  await ensureFolderPath(folder);
  state.config.folder = folder;
  await ensureFolderPath(projectPath(FILES_DIR));
  await ensureFolderPath(projectPath(REFERENCE_DIR));
  await ensureFolderPath(projectPath(TRASH_DIR));
  state.project = {
    version: 2,
    norm_pages: 381,
    characters_per_page: CHARS_PER_PAGE,
    background_color: "#17191f",
    clock_color: "#f4f4f4",
    progress_color: "#2e7df6",
    zoom: 1,
    center_x: 0,
    center_y: 0,
    view_initialized: false,
    explorer_path: ".",
    explorer_visible: true,
    raster_titles: {},
    reference_files: [],
    title_color_rules: [],
    documents: [],
  };
  const uploaded = await uploadByPath(PROJECT_FILE, JSON.stringify(state.project, null, 2), "application/json; charset=utf-8");
  state.projectETag = uploaded?.eTag || "";
  rememberProjectFolder(folder);
  el.projectDialog.close();
  await loadProject();
}

async function tryDownloadText(relative) {
  try {
    const { response } = await downloadByPath(relative);
    return await response.text();
  } catch {
    return null;
  }
}

async function migrateProjectToTextClockModel() {
  if (!state.project) return { changed: false, unresolved: [] };

  let changed = false;
  const unresolved = [];

  await ensureFolderPath(projectPath(FILES_DIR));
  await ensureFolderPath(projectPath(REFERENCE_DIR));
  await ensureFolderPath(projectPath(TRASH_DIR));

  // Alte Uhrdokumente nach „Uhr/HH.MM Titel.txt“ überführen.
  for (const doc of state.project.documents || []) {
    const current = normalizeRelative(doc.project_path || "");
    const alreadyNew = current.toLocaleLowerCase("de").startsWith(`${FILES_DIR.toLocaleLowerCase("de")}/`)
      && suffix(current) === ".txt"
      && /^\d{2}\.\d{2}\s+/.test(basenameAny(current));

    if (alreadyNew) {
      doc.source_type = "project_text";
      doc.original_path = "";
      doc.text_cache_path = current;
      doc.suffix = ".txt";
      doc.is_on_clock = true;
      continue;
    }

    const candidates = [
      normalizeRelative(doc.text_cache_path || ""),
      current,
    ].filter(Boolean);

    let text = null;
    let sourceRelative = "";
    for (const candidate of candidates) {
      const ext = suffix(candidate);
      if (![".txt", ".md"].includes(ext)) continue;
      text = await tryDownloadText(candidate);
      if (text !== null) {
        sourceRelative = candidate;
        break;
      }
    }

    if (text === null) {
      unresolved.push(doc.title || basenameAny(current) || "Unbekannte Word-Datei");
      continue;
    }

    const title = stripClockPrefix(doc.title || stem(basenameAny(current)));
    const desired = await uniqueNameInFolder(FILES_DIR, clockFilename(title, doc.start_second || 0));
    await uploadByPath(`${FILES_DIR}/${desired}`, text, "text/plain; charset=utf-8");

    const oldPaths = new Set(candidates);
    for (const oldRelative of oldPaths) {
      if (!oldRelative || normalizeRelative(oldRelative).toLocaleLowerCase("de") === `${FILES_DIR}/${desired}`.toLocaleLowerCase("de")) continue;
      try {
        const oldItem = await getItemByPath(oldRelative);
        await deleteDriveItem(oldItem);
      } catch {}
    }

    doc.title = title;
    doc.source_type = "project_text";
    doc.original_path = "";
    doc.project_path = `${FILES_DIR}/${desired}`;
    doc.text_cache_path = doc.project_path;
    doc.suffix = ".txt";
    doc.is_on_clock = true;
    doc.original_mtime_ns = 0;
    doc.character_count = text.length;
    changed = true;
  }

  // Alte Stehsatz-TXT/MD nach Stehsatz/*.txt überführen.
  const migratedRefs = [];
  for (const raw of state.project.reference_files || []) {
    if (isReferenceSeparator(raw)) {
      migratedRefs.push(raw);
      continue;
    }
    const relative = normalizeRelative(raw);
    const ext = suffix(relative);
    if (ext === ".txt" && relative.toLocaleLowerCase("de").startsWith(`${REFERENCE_DIR.toLocaleLowerCase("de")}/`)) {
      migratedRefs.push(relative);
      continue;
    }
    if (![".txt", ".md"].includes(ext)) {
      unresolved.push(stem(basenameAny(relative)));
      migratedRefs.push(relative);
      continue;
    }

    const content = await tryDownloadText(relative);
    if (content === null) {
      migratedRefs.push(relative);
      continue;
    }
    const targetName = await uniqueNameInFolder(REFERENCE_DIR, `${safeProjectStem(stem(basenameAny(relative)))}.txt`);
    await uploadByPath(`${REFERENCE_DIR}/${targetName}`, content, "text/plain; charset=utf-8");
    try {
      const oldItem = await getItemByPath(relative);
      await deleteDriveItem(oldItem);
    } catch {}
    migratedRefs.push(`${REFERENCE_DIR}/${targetName}`);
    changed = true;
  }
  state.project.reference_files = migratedRefs;

  // Alte technische Cache-Dateien im Ordner „Dateien“ beseitigen.
  try {
    const legacyItems = await listFolderByRelative(LEGACY_FILES_DIR);
    for (const item of legacyItems) {
      if (item.file && item.name.startsWith(".") && item.name.endsWith(".cache.txt")) {
        await deleteDriveItem(item);
        changed = true;
      }
    }
  } catch {}

  return { changed, unresolved: [...new Set(unresolved)] };
}

async function loadProject() {
  if (!state.config.folder) throw new Error("Bitte zuerst ein Buch-Uhr-Projekt öffnen.");
  setStatus("Lade Projekt …", "syncing");
  const { item, response } = await downloadByPath(PROJECT_FILE);
  const project = await response.json();
  if (!Array.isArray(project.documents)) project.documents = [];
  if (!Array.isArray(project.reference_files)) project.reference_files = [];
  if (!Array.isArray(project.title_color_rules)) project.title_color_rules = [];
  if (!project.raster_titles) project.raster_titles = {};

  state.project = project;
  state.projectItem = item;
  state.projectETag = item.eTag || "";

  const migration = await migrateProjectToTextClockModel();
  if (migration.changed) await saveProject();
  if (migration.unresolved.length) {
    toast(
      "Einige alte DOCX-Dateien konnten auf dem iPad nicht in TXT umgewandelt werden. Bitte dieses Projekt einmal mit Buchuhr v70 unter Windows öffnen.",
      9000
    );
  }

  state.history = [];
  state.future = [];
  el.projectName.textContent = basenameAny(state.config.folder);
  el.emptyHint.classList.add("hidden");
  renderAll();
  await loadFolderFiles();
  setStatus("Synchron", "online");
}

async function saveProject() {
  if (!state.project) return;
  setStatus("Speichert …", "syncing");
  try {
    const uploaded = await uploadByPath(
      PROJECT_FILE,
      JSON.stringify(
        state.project,
        (key, value) => key.startsWith("_") ? undefined : value,
        2
      ),
      "application/json; charset=utf-8",
      state.projectETag
    );
    state.projectETag = uploaded?.eTag || "";
    state.projectItem = uploaded || state.projectItem;
    setStatus("Synchron", "online");
  } catch (error) {
    if (error.status === 412) {
      setStatus("Konflikt", "error");
      toast("Die PC-Version wurde zwischenzeitlich geändert. Projekt wird neu geladen.", 5000);
      await loadProject();
      return;
    }
    setStatus("Fehler", "error");
    throw error;
  }
}

async function syncNow() {
  try {
    await signIn();
    await loadProject();
    clearRememberedError();
  } catch (error) {
    console.error(error);
    rememberError(error, "Synchronisierung fehlgeschlagen");
    toast(error.message || String(error), 8000);
  }
}

async function loadFolderFiles() {
  const base = encodeGraphPath(projectPath());
  const rootResult = await graph(`/me/drive/root:/${base}:/children?$select=id,name,webUrl,file,folder,lastModifiedDateTime&$top=200`);
  let items = (rootResult.value || []).map(item => ({ ...item, _relative: item.name }));

  for (const directory of [FILES_DIR, REFERENCE_DIR, TRASH_DIR]) {
    try {
      const dirPath = encodeGraphPath(projectPath(directory));
      const dirResult = await graph(`/me/drive/root:/${dirPath}:/children?$select=id,name,webUrl,file,folder,lastModifiedDateTime&$top=500`);
      items = items.concat((dirResult.value || []).map(item => ({
        ...item,
        _inFiles: directory === FILES_DIR,
        _inReferences: directory === REFERENCE_DIR,
        _inTrash: directory === TRASH_DIR,
        _relative: `${directory}/${item.name}`,
      })));
    } catch (error) {
      console.warn(`${directory}-Unterordner nicht lesbar`, error);
    }
  }

  state.folderItems = items;
  renderFileList();
}


function clearSidePreviews() {
  state.previewRequestId += 1;
  state.previewEditors.left = null;
  state.previewEditors.right = null;
  el.leftPreview.classList.add("hidden");
  el.rightPreview.classList.add("hidden");
  el.leftPreviewBody.replaceChildren();
  el.rightPreviewBody.replaceChildren();
  updatePreviewHeaderButtons("left");
  updatePreviewHeaderButtons("right");
}

function previewElements(side) {
  return side === "left"
    ? { panel: el.leftPreview, title: el.leftPreviewTitle, body: el.leftPreviewBody }
    : { panel: el.rightPreview, title: el.rightPreviewTitle, body: el.rightPreviewBody };
}

function previewHeader(side) {
  return previewElements(side).panel.querySelector("header");
}

function updatePreviewHeaderButtons(side) {
  const header = previewHeader(side);
  if (!header) return;
  header.querySelectorAll(".preview-edit-action").forEach(node => node.remove());

  const editor = state.previewEditors[side];
  if (!editor?.editing) return;

  const cancel = document.createElement("button");
  cancel.type = "button";
  cancel.className = "preview-edit-action";
  cancel.textContent = "Abbrechen";
  cancel.addEventListener("click", () => cancelPreviewEdit(side));

  const save = document.createElement("button");
  save.type = "button";
  save.className = "primary preview-edit-action";
  save.textContent = "Speichern";
  save.addEventListener("click", () => void savePreviewEdit(side));

  const close = side === "left" ? el.leftPreviewClose : el.rightPreviewClose;
  header.insertBefore(cancel, close);
  header.insertBefore(save, close);
}

function setPreviewLoading(side, title) {
  state.previewEditors[side] = null;
  const target = previewElements(side);
  target.title.textContent = title;
  target.body.textContent = "Lädt …";
  target.panel.classList.remove("hidden");
  updatePreviewHeaderButtons(side);
}

function textOffsetAtPoint(root, clientX, clientY) {
  let node = null;
  let offset = 0;

  if (document.caretPositionFromPoint) {
    const position = document.caretPositionFromPoint(clientX, clientY);
    node = position?.offsetNode || null;
    offset = position?.offset || 0;
  } else if (document.caretRangeFromPoint) {
    const range = document.caretRangeFromPoint(clientX, clientY);
    node = range?.startContainer || null;
    offset = range?.startOffset || 0;
  }

  if (!node || !root.contains(node)) return root.textContent.length;

  const range = document.createRange();
  range.selectNodeContents(root);
  try {
    range.setEnd(node, offset);
    return range.toString().length;
  } catch {
    return root.textContent.length;
  }
}

function setPreviewText(side, title, text, source) {
  const target = previewElements(side);
  target.title.textContent = title;
  target.body.replaceChildren();

  const pre = document.createElement("pre");
  pre.className = "preview-text";
  pre.textContent = text;
  target.body.append(pre);
  target.panel.classList.remove("hidden");

  state.previewEditors[side] = {
    side,
    source,
    title,
    originalText: text,
    editing: false,
    textarea: null,
  };
  updatePreviewHeaderButtons(side);

  let lastTap = { time: 0, x: 0, y: 0 };

  pre.addEventListener("dblclick", event => {
    event.preventDefault();
    event.stopPropagation();
    const caretOffset = textOffsetAtPoint(pre, event.clientX, event.clientY);
    beginPreviewEdit(side, caretOffset);
  });

  pre.addEventListener("touchend", event => {
    if (event.changedTouches.length !== 1) return;

    const touch = event.changedTouches[0];
    const now = Date.now();
    const isDouble = now - lastTap.time <= 430
      && Math.hypot(touch.clientX - lastTap.x, touch.clientY - lastTap.y) <= 34;

    if (isDouble) {
      // Nur diesen zweiten Tipp abfangen. Dadurch zoomt Safari nicht,
      // während alle Dateilisten normal weiterarbeiten.
      event.preventDefault();
      event.stopPropagation();
      const caretOffset = textOffsetAtPoint(pre, touch.clientX, touch.clientY);
      beginPreviewEdit(side, caretOffset);
      lastTap = { time: 0, x: 0, y: 0 };
    } else {
      lastTap = { time: now, x: touch.clientX, y: touch.clientY };
    }
  }, { passive: false });
}

function setPreviewMessage(side, title, message) {
  state.previewEditors[side] = null;
  const target = previewElements(side);
  target.title.textContent = title;
  target.body.textContent = message;
  target.panel.classList.remove("hidden");
  updatePreviewHeaderButtons(side);
}

function beginPreviewEdit(side, caretOffset = 0) {
  const editor = state.previewEditors[side];
  if (!editor || editor.editing) return;

  const target = previewElements(side);
  const textarea = document.createElement("textarea");
  textarea.className = "side-preview-editor";
  textarea.value = editor.originalText;
  target.body.replaceChildren(textarea);

  editor.editing = true;
  editor.textarea = textarea;
  updatePreviewHeaderButtons(side);

  requestAnimationFrame(() => {
    textarea.focus({ preventScroll: true });
    const safeOffset = Math.max(0, Math.min(caretOffset, textarea.value.length));
    textarea.setSelectionRange(safeOffset, safeOffset);
    const lineHeight = parseFloat(getComputedStyle(textarea).lineHeight) || 20;
    const linesBefore = textarea.value.slice(0, safeOffset).split("\n").length - 1;
    textarea.scrollTop = Math.max(0, linesBefore * lineHeight - textarea.clientHeight * 0.42);
  });
}

function cancelPreviewEdit(side) {
  const editor = state.previewEditors[side];
  if (!editor) return;
  setPreviewText(side, editor.title, editor.originalText, editor.source);
}

async function savePreviewEdit(side) {
  const editor = state.previewEditors[side];
  if (!editor?.editing || !editor.textarea) return;

  const value = editor.textarea.value.replace(/\r\n?/g, "\n");

  try {
    if (editor.source.type === "project") {
      const doc = editor.source.doc;
      const uploaded = await uploadByPath(
        editor.source.relative || projectCopyRelative(doc),
        value,
        "text/plain; charset=utf-8",
        editor.source.item?.eTag || ""
      );
      doc.character_count = value.length;
      editor.source.item = uploaded;
      await saveProject();
      renderProgress();
      renderDocuments();
    } else {
      const item = editor.source.item;
      const accessToken = await token();
      const response = await fetch(`${GRAPH_BASE}/me/drive/items/${item.id}/content`, {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "text/plain; charset=utf-8",
        },
        body: value,
      });
      if (!response.ok) {
        const body = await response.text();
        throw new Error(`Datei konnte nicht gespeichert werden.\n${response.status} ${response.statusText}\n${body}`);
      }
    }

    setPreviewText(side, editor.title, value, editor.source);
    toast("Datei gespeichert");
  } catch (error) {
    toast(error.message || String(error), 6000);
  }
}

async function showProjectDocPreview(doc, side = "right") {
  const requestId = ++state.previewRequestId;
  const title = doc.title || stem(basenameAny(doc.project_path));
  setPreviewLoading(side, title);

  const editable = [".txt", ".md"].includes(String(doc.suffix).toLowerCase())
    || doc.source_type === "dragged_text";

  if (!editable) {
    setPreviewMessage(side, title, "Diese Datei kann in der Vorschau nicht als Text angezeigt werden.");
    return;
  }

  try {
    const content = await getDocumentContent(doc);
    if (requestId !== state.previewRequestId) return;
    setPreviewText(side, title, content.text, {
      type: "project",
      doc,
      item: content.item,
      relative: content.relative,
    });
  } catch (error) {
    if (requestId !== state.previewRequestId) return;
    setPreviewMessage(side, title, error.message || String(error));
  }
}

async function showReferencePreview(raw, side = "left") {
  const requestId = ++state.previewRequestId;
  const name = basenameAny(raw);
  const title = stem(name);
  setPreviewLoading(side, title);

  const ext = suffix(name);
  if (![".txt", ".md"].includes(ext)) {
    setPreviewMessage(side, title, "Diese Datei kann in der Vorschau nicht als Text angezeigt werden.");
    return;
  }

  try {
    const item = findReferenceItem(raw) || await searchExactFile(name);
    if (!item) throw new Error(`Datei in OneDrive nicht gefunden: ${name}`);

    let response;
    if (item["@microsoft.graph.downloadUrl"]) {
      response = await fetch(item["@microsoft.graph.downloadUrl"]);
    } else {
      const accessToken = await token();
      response = await fetch(`${GRAPH_BASE}/me/drive/items/${item.id}/content`, {
        headers: { Authorization: `Bearer ${accessToken}` },
      });
    }

    if (!response.ok) throw new Error(`Datei konnte nicht geladen werden: ${name}`);
    const text = await response.text();
    if (requestId !== state.previewRequestId) return;
    setPreviewText(side, title, text, {
      type: "reference",
      raw,
      item,
    });
  } catch (error) {
    if (requestId !== state.previewRequestId) return;
    setPreviewMessage(side, title, error.message || String(error));
  }
}

async function activateSidebarFile(key, openEditor, previewFn) {
  state.lastSidebarFileTap = { key, time: Date.now() };
  await previewFn();
}


function normalizedTitleColorRules() {
  return (state.project?.title_color_rules || [])
    .filter(rule => rule && String(rule.word || "").trim() && /^#[0-9a-f]{6}$/i.test(String(rule.color || "")))
    .map(rule => ({ word: String(rule.word).trim(), color: String(rule.color).toLowerCase() }));
}

function applyTitleColorMarkup(node, title) {
  const source = String(title || "");
  node.replaceChildren();
  if (!source) return;

  const assignments = Array(source.length).fill(null);
  const folded = source.toLocaleLowerCase("de");

  for (const rule of normalizedTitleColorRules()) {
    const needle = rule.word.toLocaleLowerCase("de");
    let start = 0;
    while (needle && start < folded.length) {
      const pos = folded.indexOf(needle, start);
      if (pos < 0) break;
      const end = Math.min(source.length, pos + rule.word.length);
      for (let i = pos; i < end; i += 1) {
        if (assignments[i] === null) assignments[i] = rule.color;
      }
      start = Math.max(pos + 1, end);
    }
  }

  let runStart = 0;
  let runColor = assignments[0] || null;
  for (let i = 1; i <= source.length; i += 1) {
    const color = i < source.length ? assignments[i] : Symbol("end");
    if (i === source.length || color !== runColor) {
      const span = document.createElement("span");
      span.textContent = source.slice(runStart, i);
      if (runColor) {
        span.style.color = runColor;
        span.style.fontWeight = "700";
      }
      node.append(span);
      runStart = i;
      runColor = i < source.length ? color : null;
    }
  }
}


function clockTitleColorAssignments(title) {
  const source = String(title || "").trim();
  const assignments = Array(source.length).fill(null);
  const folded = source.toLocaleLowerCase("de");

  for (const rule of normalizedTitleColorRules()) {
    const needle = rule.word.toLocaleLowerCase("de");
    let start = 0;
    while (needle && start < folded.length) {
      const pos = folded.indexOf(needle, start);
      if (pos < 0) break;
      const end = Math.min(source.length, pos + rule.word.length);
      for (let i = pos; i < end; i += 1) {
        if (assignments[i] === null) assignments[i] = rule.color;
      }
      start = Math.max(pos + 1, end);
    }
  }
  return { source, assignments };
}

function splitClockTitleLines(title, maxWidth = 250, maxLines = 3) {
  const source = String(title || "").trim();
  if (!source) return [];

  const words = [...source.matchAll(/\S+/g)].map(match => ({
    text: match[0],
    start: match.index,
    end: match.index + match[0].length,
  }));
  if (!words.length) return [];

  const canvas = splitClockTitleLines._canvas || (splitClockTitleLines._canvas = document.createElement("canvas"));
  const context = canvas.getContext("2d");
  context.font = '500 15px system-ui, -apple-system, "Segoe UI", sans-serif';
  const spaceWidth = context.measureText(" ").width;

  const lines = [];
  let current = [];
  let width = 0;

  for (const word of words) {
    const wordWidth = context.measureText(word.text).width;
    const nextWidth = current.length ? width + spaceWidth + wordWidth : wordWidth;
    if (current.length && nextWidth > maxWidth) {
      lines.push(current);
      current = [word];
      width = wordWidth;
    } else {
      current.push(word);
      width = nextWidth;
    }
  }
  if (current.length) lines.push(current);

  if (lines.length <= maxLines) return lines;

  const kept = lines.slice(0, maxLines);
  const overflowWords = lines.slice(maxLines - 1).flat();
  kept[maxLines - 1] = overflowWords;
  return kept;
}

function appendColoredClockWord(textNode, source, assignments, word, addSpace) {
  if (addSpace) {
    const space = svg("tspan", {}, " ");
    textNode.append(space);
  }

  let runStart = word.start;
  let runColor = assignments[word.start] || null;
  for (let i = word.start + 1; i <= word.end; i += 1) {
    const color = i < word.end ? assignments[i] : Symbol("end");
    if (i === word.end || color !== runColor) {
      const attrs = {};
      if (runColor) {
        attrs.fill = runColor;
        attrs["font-weight"] = "700";
      }
      textNode.append(svg("tspan", attrs, source.slice(runStart, i)));
      runStart = i;
      runColor = i < word.end ? color : null;
    }
  }
}

function createClockTitleGroup(title, horizontal, vertical, baseColor) {
  const { source, assignments } = clockTitleColorAssignments(title);
  const lines = splitClockTitleLines(source, 250, 3);
  const group = svg("g", { class: "doc-title-svg-group" });
  if (!lines.length) return group;

  const centered = Math.abs(horizontal) < 0.28;
  const lineHeight = 18;
  const totalHeight = (lines.length - 1) * lineHeight;
  let x = 0;
  let firstY = -totalHeight / 2 + 5;
  let anchor = "middle";

  if (centered) {
    // Oben und unten liegen Titel radial außerhalb des Symbols statt darüber.
    firstY = vertical < 0
      ? -34 - totalHeight
      : 45;
  } else if (horizontal > 0) {
    x = 30;
    anchor = "start";
  } else {
    x = -30;
    anchor = "end";
  }

  lines.forEach((words, lineIndex) => {
    const text = svg("text", {
      x,
      y: firstY + lineIndex * lineHeight,
      class: "doc-title-svg",
      fill: baseColor,
      "text-anchor": anchor,
      "dominant-baseline": "middle",
    });
    words.forEach((word, wordIndex) => {
      appendColoredClockWord(text, source, assignments, word, wordIndex > 0);
    });
    group.append(text);
  });

  return group;
}

function renderTitleColorRulesEditor() {
  el.titleColorRows.replaceChildren();
  const rules = normalizedTitleColorRules();
  if (!rules.length) rules.push({ word: "", color: "#f4f4f4" });
  for (const rule of rules) addTitleColorRuleRow(rule.word, rule.color);
}

function addTitleColorRuleRow(word = "", color = "#f4f4f4") {
  const row = document.createElement("div");
  row.className = "title-color-row";

  const input = document.createElement("input");
  input.type = "text";
  input.className = "title-color-word";
  input.placeholder = "Wort im Dateinamen";
  input.value = word;

  const picker = document.createElement("input");
  picker.type = "color";
  picker.className = "title-color-picker";
  picker.value = /^#[0-9a-f]{6}$/i.test(color) ? color : "#f4f4f4";

  row.append(input, picker);
  el.titleColorRows.append(row);
  input.focus();
}

function openTitleColorDialog() {
  if (!state.project) return;
  renderTitleColorRulesEditor();
  el.titleColorDialog.showModal();
}

async function saveTitleColorRules() {
  if (!state.project) return;
  const rules = [...el.titleColorRows.querySelectorAll(".title-color-row")]
    .map(row => ({
      word: row.querySelector(".title-color-word")?.value.trim() || "",
      color: row.querySelector(".title-color-picker")?.value || "#f4f4f4",
    }))
    .filter(rule => rule.word);

  pushHistory();
  state.project.title_color_rules = rules;
  await saveProject();
  el.titleColorDialog.close();
  renderDocuments();
  renderReferenceList();
  toast("Titelfarben gespeichert");
}

function fileIconClass(name) {
  return ["≡", "text"];
}

function explorerFolderForItem(item) {
  if (item._inFiles) return FILES_DIR;
  if (item._inReferences) return REFERENCE_DIR;
  if (item._inTrash) return TRASH_DIR;
  return "";
}

function isExplorerFolderExpanded(name) {
  return state.expandedExplorerFolders.has(name);
}

function setExplorerFolderExpanded(name, expanded) {
  if (expanded) state.expandedExplorerFolders.add(name);
  else state.expandedExplorerFolders.delete(name);
  localStorage.setItem(
    "buchuhr.explorerFolders",
    JSON.stringify([...state.expandedExplorerFolders])
  );
  renderFileList();
}

function makeFileRow(item, nested = false) {
  const row = document.createElement("div");
  row.className = `file-row ${item.id === state.selectedFileId ? "selected" : ""}`;
  if (nested) row.classList.add("explorer-child");
  row.dataset.fileItemId = item.id || "";

  const [label, cls] = fileIconClass(item.name);
  const expanded = item.folder && isExplorerFolderExpanded(item.name);
  const iconLabel = item.folder ? (expanded ? "▾" : "▸") : label;
  row.innerHTML = `<span class="file-icon ${cls}">${iconLabel}</span><span class="file-name"></span>`;
  const displayName = item.folder ? item.name : stem(item.name);
  row.querySelector(".file-name").textContent = displayName;
  if (item._inTrash) row.classList.add("in-trash");
  if (item.folder) row.classList.add("explorer-folder");
  row.dataset.fileRelative = item._relative || item.name;

  if (item.folder) {
    row.setAttribute("aria-expanded", expanded ? "true" : "false");
    row.addEventListener("click", event => {
      event.preventDefault();
      event.stopPropagation();
      setExplorerFolderExpanded(item.name, !isExplorerFolderExpanded(item.name));
    });
    row.addEventListener("dblclick", event => {
      event.preventDefault();
      event.stopPropagation();
    });
    return row;
  }

  const matchingDoc = () => findDocByName(item.name);

  row.addEventListener("click", async event => {
    event.preventDefault();
    event.stopPropagation();

    state.selectedFileId = item.id || null;
    state.selectedReferencePath = null;
    const doc = matchingDoc();
    state.selectedDocId = doc?.id || null;
    updateSelectionVisuals();

    if (!doc) return;

    await activateSidebarFile(
      `left:${doc.id}`,
      async () => {
        selectDoc(doc);
        await openDocumentEditor(doc);
      },
      async () => {
        await showProjectDocPreview(doc, "right");
      }
    );
  });

  row.addEventListener("dblclick", event => {
    event.preventDefault();
    event.stopPropagation();
  });

  row.addEventListener("contextmenu", event => {
    event.preventDefault();
    state.selectedFileId = item.id || null;
    state.selectedReferencePath = null;
    const doc = matchingDoc();
    state.selectedDocId = doc?.id || null;
    updateSelectionVisuals();
    showContextMenu(event.clientX, event.clientY, doc);
  });

  attachSwipeRemoval(row, "right", async () => {
    await deleteLeftFile(item);
  });

  return row;
}

function renderFileList() {
  el.fileList.replaceChildren();

  const visibleItems = [...state.folderItems]
    .filter(item => item.name !== PROJECT_FILE)
    .filter(item => item.folder || suffix(item.name) === ".txt");

  const rootItems = visibleItems
    .filter(item => !explorerFolderForItem(item))
    .sort((a, b) =>
      Number(Boolean(b.folder)) - Number(Boolean(a.folder))
      || a.name.localeCompare(b.name, "de")
    );

  for (const item of rootItems) {
    el.fileList.append(makeFileRow(item));

    if (!item.folder || !isExplorerFolderExpanded(item.name)) continue;

    const children = visibleItems
      .filter(child => explorerFolderForItem(child) === item.name)
      .sort((a, b) => a.name.localeCompare(b.name, "de"));

    for (const child of children) {
      el.fileList.append(makeFileRow(child, true));
    }
  }
}

function svg(tag, attrs = {}, text = "") {
  const node = document.createElementNS("http://www.w3.org/2000/svg", tag);
  for (const [key, value] of Object.entries(attrs)) node.setAttribute(key, value);
  if (text) node.textContent = text;
  return node;
}

function polar(second, radius) {
  const angle = second / 3600 * Math.PI * 2 - Math.PI / 2;
  return { x: Math.cos(angle) * radius, y: Math.sin(angle) * radius, angle };
}

function secondFromSvgEvent(event) {
  const pt = el.clockCanvas.createSVGPoint();
  pt.x = event.clientX;
  pt.y = event.clientY;
  const local = pt.matrixTransform(el.clockCanvas.getScreenCTM().inverse());
  const angle = Math.atan2(local.y, local.x) + Math.PI / 2;
  return ((angle / (Math.PI * 2) * 3600) + 3600) % 3600;
}

function renderClock() {
  el.clockLayer.replaceChildren();
  const project = state.project;
  if (!project) return;
  const clockColor = project.clock_color || "#f4f4f4";

  el.clockLayer.append(svg("circle", {
    cx: 0, cy: 0, r: 360, class: "clock-outline", stroke: clockColor
  }));

  for (let minute = 0; minute < 60; minute++) {
    const major = minute % 15 === 0;
    const five = minute % 5 === 0;
    const inner = major ? 330 : five ? 340 : 347;
    const outer = 358;
    const a = polar(minute * 60, inner);
    const b = polar(minute * 60, outer);
    const tickSecond = minute * 60;

    el.clockLayer.append(svg("line", {
      x1: a.x, y1: a.y, x2: b.x, y2: b.y,
      class: "tick", stroke: clockColor,
      "stroke-width": major ? 6 : five ? 3.2 : 2
    }));

    const hitInner = polar(tickSecond, major ? 314 : 324);
    const hitOuter = polar(tickSecond, 374);
    el.clockLayer.append(svg("line", {
      x1: hitInner.x,
      y1: hitInner.y,
      x2: hitOuter.x,
      y2: hitOuter.y,
      class: "raster-hit",
      "data-second": String(tickSecond)
    }));
  }
}

function progressSpan(doc) {
  const capacity = Math.max(1, Number(state.project.norm_pages || 381) * CHARS_PER_PAGE);
  return Math.max(1, Math.ceil(Number(doc.character_count || 0) / capacity * 3600));
}

function arcPath(startSecond, spanSecond, radius = 360) {
  const start = polar(startSecond, radius);
  const end = polar((startSecond + spanSecond) % 3600, radius);
  const large = spanSecond > 1800 ? 1 : 0;
  return `M ${start.x} ${start.y} A ${radius} ${radius} 0 ${large} 1 ${end.x} ${end.y}`;
}

function renderProgress() {
  el.progressLayer.replaceChildren();
  if (!state.project) return;
  const color = state.project.progress_color || "#2e7df6";

  for (const doc of state.project.documents.filter(d => d.is_on_clock !== false)) {
    const path = svg("path", {
      d: arcPath(Number(doc.start_second || 0), progressSpan(doc)),
      class: "progress-arc canvas-file-target",
      stroke: color,
      "data-doc-id": doc.id,
    });

    path.addEventListener("pointerdown", event => {
      beginCanvasFilePointer(event, doc, path, "progress");
    });

    path.addEventListener("contextmenu", event => {
      event.preventDefault();
      event.stopPropagation();
      selectDoc(doc);
      showContextMenu(event.clientX, event.clientY, doc);
    });

    el.progressLayer.append(path);
  }
}

function wrapSvgText(text, maxChars = 24) {
  const words = text.split(/\s+/);
  const lines = [];
  let line = "";
  for (const word of words) {
    if ((line + " " + word).trim().length > maxChars && line) {
      lines.push(line);
      line = word;
    } else {
      line = (line + " " + word).trim();
    }
  }
  if (line) lines.push(line);
  return lines.slice(0, 4);
}

function renderRasterTitles() {
  el.rasterTitleLayer.replaceChildren();
  if (!state.project) return;
  const titles = state.project.raster_titles || {};
  const circleRadius = 360;

  for (const [minuteText, title] of Object.entries(titles)) {
    const minute = Number(minuteText) % 60;
    const major = minute % 15 === 0;
    const second = minute * 60;
    const angle = second / 3600 * Math.PI * 2 - Math.PI / 2;
    const cos = Math.cos(angle);
    const sin = Math.sin(angle);

    const boxWidth = major ? 200 : 178;
    const boxHeight = major ? 60 : 50;
    const inset = 34;

    let centerX = cos * 235;
    let centerY = sin * 235;

    if (minute === 15) centerX = circleRadius - inset - boxWidth / 2;
    if (minute === 45) centerX = -circleRadius + inset + boxWidth / 2;
    if (minute === 0) centerY = -circleRadius + inset + boxHeight / 2;
    if (minute === 30) centerY = circleRadius - inset - boxHeight / 2;

    const foreign = svg("foreignObject", {
      x: centerX - boxWidth / 2,
      y: centerY - boxHeight / 2,
      width: boxWidth,
      height: boxHeight,
      class: `raster-title-box ${major ? "major" : ""}`
    });

    const div = document.createElement("div");
    div.className = `raster-title-html ${major ? "major" : ""}`;
    div.textContent = String(title);
    foreign.append(div);

    const lineStart = polar(second, major ? 325 : 337);
    const lineStopRadius = major ? 292 : 304;
    const lineStop = polar(second, lineStopRadius);

    el.rasterTitleLayer.append(svg("line", {
      x1: lineStart.x,
      y1: lineStart.y,
      x2: lineStop.x,
      y2: lineStop.y,
      class: "connector"
    }));
    el.rasterTitleLayer.append(foreign);
  }
}

function renderDocuments() {
  el.documentLayer.replaceChildren();
  if (!state.project) return;

  for (const doc of state.project.documents.filter(d => d.is_on_clock !== false)) {
    const p = polar(Number(doc.start_second || 0), 430);
    const horizontal = Math.cos(p.angle);
    const vertical = Math.sin(p.angle);
    const right = horizontal >= 0;
    const boxWidth = 250;
    const boxX = right ? 30 : -30 - boxWidth;

    const g = svg("g", {
      class: `doc-node canvas-file-target ${doc.id === state.selectedDocId ? "selected" : ""}`,
      "data-doc-id": doc.id,
      transform: `translate(${p.x} ${p.y})`,
    });

    // Unsichtbare, zusammenhängende Trefferfläche für Symbol und Titel.
    const hitX = right ? -25 : boxX - 5;
    const hitWidth = boxWidth + 60;
    g.append(svg("rect", {
      x: hitX,
      y: -36,
      width: hitWidth,
      height: 72,
      rx: 8,
      class: "doc-hit-area",
    }));

    const icon = svg("g", {
      class: `doc-icon ${doc.suffix === ".docx" ? "word" : doc.suffix === ".md" ? "md" : ""}`,
    });
    icon.append(svg("rect", { x: -19, y: -19, width: 38, height: 38, rx: 4 }));
    const label = doc.suffix === ".docx" ? "W" : doc.suffix === ".md" ? "MD" : "≡";
    icon.append(svg("text", { x: 0, y: 1 }, label));
    g.append(icon);

    // Native SVG-Titel statt foreignObject/HTML. Das verhindert die fehlerhafte
    // Wort- und Zeichenverteilung in Firefox/Safari bei skaliertem SVG.
    g.append(createClockTitleGroup(
      doc.title || stem(basenameAny(doc.project_path)),
      horizontal,
      vertical,
      state.project.clock_color || "#f4f4f4",
    ));

    g.addEventListener("pointerdown", event => {
      beginCanvasFilePointer(event, doc, g, "document");
    });

    g.addEventListener("contextmenu", event => {
      event.preventDefault();
      event.stopPropagation();
      selectDoc(doc);
      showContextMenu(event.clientX, event.clientY, doc);
    });

    el.documentLayer.append(g);
  }
}

function pointerSecondFromClient(clientX, clientY) {
  const point = el.clockCanvas.createSVGPoint();
  point.x = clientX;
  point.y = clientY;
  const local = point.matrixTransform(el.clockCanvas.getScreenCTM().inverse());

  // View transform rückwärts berücksichtigen.
  const x = (local.x - state.viewPanX) / state.viewZoom;
  const y = (local.y - state.viewPanY) / state.viewZoom;
  let angle = Math.atan2(y, x) + Math.PI / 2;
  if (angle < 0) angle += Math.PI * 2;
  return angle / (Math.PI * 2) * 3600;
}

function signedSecondDelta(current, initial) {
  let delta = current - initial;
  while (delta > 1800) delta -= 3600;
  while (delta < -1800) delta += 3600;
  return delta;
}

function localSvgPointFromClient(clientX, clientY) {
  const point = el.clockCanvas.createSVGPoint();
  point.x = clientX;
  point.y = clientY;
  return point.matrixTransform(el.clockCanvas.getScreenCTM().inverse());
}

function contentPointFromClient(clientX, clientY) {
  const local = localSvgPointFromClient(clientX, clientY);
  return {
    x: (local.x - state.viewPanX) / state.viewZoom,
    y: (local.y - state.viewPanY) / state.viewZoom,
    local,
  };
}

function touchDistance(t1, t2) {
  return Math.hypot(t2.clientX - t1.clientX, t2.clientY - t1.clientY);
}

function touchMidpoint(t1, t2) {
  return {
    x: (t1.clientX + t2.clientX) / 2,
    y: (t1.clientY + t2.clientY) / 2,
  };
}

function clampZoom(value) {
  return Math.max(.55, Math.min(2.2, value));
}

function beginCanvasPinch(touches) {
  if (touches.length < 2) return;
  clearPointerAction();
  const [t1, t2] = touches;
  const midpoint = touchMidpoint(t1, t2);
  const content = contentPointFromClient(midpoint.x, midpoint.y);
  state.pinch = {
    distance: Math.max(1, touchDistance(t1, t2)),
    startZoom: state.viewZoom,
    startPanX: state.viewPanX,
    startPanY: state.viewPanY,
    anchorX: content.x,
    anchorY: content.y,
  };
}

function updateCanvasPinch(touches) {
  const pinch = state.pinch;
  if (!pinch || touches.length < 2) return;
  const [t1, t2] = touches;
  const midpoint = touchMidpoint(t1, t2);
  const local = localSvgPointFromClient(midpoint.x, midpoint.y);
  state.viewZoom = clampZoom(pinch.startZoom * (touchDistance(t1, t2) / pinch.distance));
  state.viewPanX = local.x - pinch.anchorX * state.viewZoom;
  state.viewPanY = local.y - pinch.anchorY * state.viewZoom;
  applyViewZoom();
}

function endCanvasPinch() {
  if (!state.pinch) return;
  state.pinch = null;
  applyViewZoom();
}

function isEditableTarget(target) {
  return Boolean(target?.closest?.('input, textarea, select, button, [contenteditable="true"]'));
}

function rememberTouchEnd(event) {
  const now = Date.now();
  const target = event.target;
  const inApp = Boolean(target?.closest?.('.topbar, .sidebar, .canvas-wrap, dialog'));
  if (!inApp || isEditableTarget(target)) {
    state.lastTouchEndTime = now;
    return;
  }
  if (now - state.lastTouchEndTime < 360) event.preventDefault();
  state.lastTouchEndTime = now;
}

function registerRasterTap(second) {
  const now = Date.now();
  const rounded = Math.round(Number(second || 0) / 60) * 60 % 3600;
  const previous = state.lastRasterTap;
  const isSecondTap = previous.second === rounded && now - previous.time <= 430;

  if (isSecondTap) {
    state.lastRasterTap = { second: null, time: 0 };
    editRasterTitle(rounded);
    return true;
  }

  state.lastRasterTap = { second: rounded, time: now };
  return false;
}

function registerSidebarBlankTap(area) {
  const now = Date.now();
  const previous = state.lastSidebarBlankTap;
  const isSecondTap = previous.area === area && now - previous.time <= 430;

  if (isSecondTap) {
    state.lastSidebarBlankTap = { area: null, time: 0 };
    createNewText();
    return true;
  }

  state.lastSidebarBlankTap = { area, time: now };
  return false;
}

function clearPointerAction() {
  const action = state.pointerAction;
  if (!action) return null;

  clearTimeout(action.touchHoldTimer);
  action.node?.classList.remove("dragging");

  try {
    if (action.captureSet && el.clockCanvas.hasPointerCapture(action.pointerId)) {
      el.clockCanvas.releasePointerCapture(action.pointerId);
    }
  } catch {}

  state.pointerAction = null;
  return action;
}

function beginCanvasFilePointer(event, doc, node, source) {
  if (event.button !== 0 && event.pointerType !== "touch") return;

  event.preventDefault();
  event.stopPropagation();
  selectDoc(doc);

  const initialPointerSecond = pointerSecondFromClient(event.clientX, event.clientY);
  const action = {
    kind: "file",
    source,
    pointerId: event.pointerId,
    pointerType: event.pointerType,
    doc,
    node,
    startClientX: event.clientX,
    startClientY: event.clientY,
    initialPointerSecond,
    originalSecond: Number(doc.start_second || 0),
    moved: false,
    captureSet: false,
    historyPushed: false,
    touchHoldTimer: null,
  };

  // Langdruck gibt es nur auf Touch, nie mit der Maus.
  if (event.pointerType === "touch") {
    action.touchHoldTimer = setTimeout(() => {
      if (state.pointerAction === action && !action.moved) {
        clearPointerAction();
        showContextMenu(event.clientX, event.clientY, doc);
      }
    }, 650);
  }

  state.pointerAction = action;
}

function beginCanvasPan(event) {
  if (event.button !== 0 && event.pointerType !== "touch") return;

  const point = el.clockCanvas.createSVGPoint();
  point.x = event.clientX;
  point.y = event.clientY;
  const local = point.matrixTransform(el.clockCanvas.getScreenCTM().inverse());

  state.pointerAction = {
    kind: "pan",
    pointerId: event.pointerId,
    startClientX: event.clientX,
    startClientY: event.clientY,
    lastX: local.x,
    lastY: local.y,
    moved: false,
    captureSet: false,
  };
}

async function activateCanvasFile(doc) {
  const now = Date.now();
  const previous = state.lastCanvasFileActivation;
  const isSecondClick = previous.id === doc.id && now - previous.time <= 430;

  if (isSecondClick) {
    state.lastCanvasFileActivation = { id: null, time: 0 };
    await openDocumentEditor(doc);
  } else {
    state.lastCanvasFileActivation = { id: doc.id, time: now };
  }
}

// Rastertitel: ausschließlich direkter Doppelklick auf einen Rasterstrich.
el.clockLayer.addEventListener("dblclick", event => {
  const tick = event.composedPath().find(
    node => node?.classList?.contains?.("raster-hit")
  );
  if (!tick) return;

  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation();
  editRasterTitle(Number(tick.dataset.second || 0));
}, true);

el.clockLayer.addEventListener("pointerup", event => {
  if (event.pointerType !== "touch") return;
  const tick = event.composedPath().find(
    node => node?.classList?.contains?.("raster-hit")
  );
  if (!tick) return;
  event.preventDefault();
  event.stopPropagation();
  registerRasterTap(Number(tick.dataset.second || 0));
}, true);

// Doppelklick auf freien Canvas oder auf Dateien erzeugt niemals Rastertitel.
el.clockCanvas.addEventListener("dblclick", event => {
  const onTick = event.composedPath().some(
    node => node?.classList?.contains?.("raster-hit")
  );
  if (onTick) return;
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation();
}, true);

el.clockCanvas.addEventListener("touchstart", event => {
  if (event.touches.length >= 2) {
    event.preventDefault();
    beginCanvasPinch(event.touches);
  }
}, { passive: false });

el.clockCanvas.addEventListener("touchmove", event => {
  if (event.touches.length >= 2) {
    event.preventDefault();
    updateCanvasPinch(event.touches);
  }
}, { passive: false });

el.clockCanvas.addEventListener("touchend", event => {
  rememberTouchEnd(event);
  if (state.pinch && event.touches.length < 2) endCanvasPinch();
}, { passive: false });

el.clockCanvas.addEventListener("touchcancel", () => {
  endCanvasPinch();
}, { passive: false });

el.clockCanvas.addEventListener("wheel", event => {
  if (!event.ctrlKey) return;
  event.preventDefault();
  event.stopPropagation();
  zoomBy(event.deltaY < 0 ? 1.10 : 0.90);
}, { passive: false });

el.clockCanvas.addEventListener("contextmenu", event => {
  const onFile = event.composedPath().some(
    node => node?.classList?.contains?.("canvas-file-target")
  );
  if (onFile) return;
  event.preventDefault();
  showContextMenu(event.clientX, event.clientY);
});

el.clockCanvas.addEventListener("pointerdown", event => {
  const path = event.composedPath();
  const onFile = path.some(node => node?.classList?.contains?.("canvas-file-target"));
  const onTick = path.some(node => node?.classList?.contains?.("raster-hit"));
  if (onFile || onTick || event.button === 2) return;
  clearSidePreviews();
  beginCanvasPan(event);
});

window.addEventListener("pointermove", event => {
  const action = state.pointerAction;
  if (!action || action.pointerId !== event.pointerId) return;

  // Nach Loslassen darf kein Zustand weiterlaufen.
  if (event.pointerType !== "touch" && event.buttons === 0) {
    clearPointerAction();
    return;
  }

  const distance = Math.hypot(
    event.clientX - action.startClientX,
    event.clientY - action.startClientY
  );

  if (action.kind === "file") {
    if (!action.moved && distance <= 10) return;

    if (!action.moved) {
      action.moved = true;
      clearTimeout(action.touchHoldTimer);
      action.node?.classList.add("dragging");

      if (!action.historyPushed) {
        pushHistory();
        action.historyPushed = true;
      }

      try {
        el.clockCanvas.setPointerCapture(action.pointerId);
        action.captureSet = true;
      } catch {}
    }

    const currentPointerSecond = pointerSecondFromClient(event.clientX, event.clientY);
    const delta = signedSecondDelta(currentPointerSecond, action.initialPointerSecond);
    action.doc.start_second = magneticSecond(action.originalSecond + delta);

    renderProgress();
    renderDocuments();
    return;
  }

  if (action.kind === "pan") {
    if (!action.moved && distance <= 4) return;

    if (!action.moved) {
      action.moved = true;
      try {
        el.clockCanvas.setPointerCapture(action.pointerId);
        action.captureSet = true;
      } catch {}
    }

    const point = el.clockCanvas.createSVGPoint();
    point.x = event.clientX;
    point.y = event.clientY;
    const local = point.matrixTransform(el.clockCanvas.getScreenCTM().inverse());

    state.viewPanX += local.x - action.lastX;
    state.viewPanY += local.y - action.lastY;
    action.lastX = local.x;
    action.lastY = local.y;
    applyViewZoom();
  }
}, true);

function pointInside(element, x, y) {
  const rect = element.getBoundingClientRect();
  return x >= rect.left && x <= rect.right && y >= rect.top && y <= rect.bottom;
}

function canvasDropDestination(x, y) {
  if (pointInside(el.rightPanel, x, y)) return REFERENCE_DIR;
  if (pointInside(el.leftPanel, x, y)) {
    const target = document.elementFromPoint(x, y)?.closest?.(".file-row");
    if (target) {
      const item = state.folderItems.find(entry => entry.id === target.dataset.fileItemId);
      if (item?.folder && [REFERENCE_DIR, TRASH_DIR].includes(item.name)) return item.name;
      if (item?.folder && item.name === FILES_DIR) return FILES_DIR;
    }
    return "";
  }
  return null;
}

window.addEventListener("pointerup", async event => {
  const action = state.pointerAction;
  if (!action || action.pointerId !== event.pointerId) return;

  const finished = clearPointerAction();
  if (!finished) return;

  if (finished.kind === "file") {
    if (finished.moved) {
      const destination = canvasDropDestination(event.clientX, event.clientY);
      try {
        if (destination !== null && destination !== FILES_DIR) {
          finished.doc.start_second = finished.originalSecond;
          await moveClockDocToFolder(finished.doc, destination);
          toast(destination === REFERENCE_DIR ? "Datei in den Stehsatz verschoben" :
                destination === TRASH_DIR ? "Datei in den Papierkorb verschoben" :
                "Datei ins Projekt-Root verschoben");
        } else {
          await renameClockDocForState(finished.doc, finished.doc.title, finished.doc.start_second);
          await saveProject();
          await loadFolderFiles();
          renderAll();
        }
      } catch (error) {
        finished.doc.start_second = finished.originalSecond;
        renderProgress();
        renderDocuments();
        toast(error.message || String(error), 6000);
      }
    } else {
      finished.doc.start_second = finished.originalSecond;
      await activateCanvasFile(finished.doc);
    }
    return;
  }

  if (finished.kind === "pan") {
    applyViewZoom();
  }
}, true);

window.addEventListener("pointercancel", event => {
  const action = state.pointerAction;
  if (!action || action.pointerId !== event.pointerId) return;

  const cancelled = clearPointerAction();
  if (cancelled?.kind === "file" && cancelled.moved) {
    cancelled.doc.start_second = cancelled.originalSecond;
    renderProgress();
    renderDocuments();
  }
}, true);

window.addEventListener("blur", () => {
  const action = clearPointerAction();
  if (action?.kind === "file" && action.moved) {
    action.doc.start_second = action.originalSecond;
    renderProgress();
    renderDocuments();
  }
});



const SEPARATOR_PREFIX = "__BUCHUHR_SEPARATOR__:";

function isReferenceSeparator(raw) {
  return String(raw || "").startsWith(SEPARATOR_PREFIX);
}

function separatorTitle(raw) {
  return String(raw || "").slice(SEPARATOR_PREFIX.length);
}

async function createReferenceTextFile() {
  if (!state.project) return;
  try {
    pushHistory();
    const name = await uniqueNameInFolder(REFERENCE_DIR, "Neue Textdatei.txt");
    await uploadByPath(`${REFERENCE_DIR}/${name}`, "", "text/plain; charset=utf-8");
    state.project.reference_files ||= [];
    const raw = `${REFERENCE_DIR}/${name}`;
    state.project.reference_files.push(raw);
    state.selectedReferencePath = raw;
    state.selectedFileId = null;
    state.selectedDocId = null;
    await saveProject();
    await loadFolderFiles();
    renderReferenceList();
    openActionDialog("renameReference", "Stehsatzdatei umbenennen", stem(name));
  } catch (error) {
    toast(error.message || String(error), 6000);
  }
}

async function createReferenceSeparator() {
  if (!state.project) return;
  pushHistory();

  const separator = `${SEPARATOR_PREFIX}`;
  state.project.reference_files ||= [];
  state.project.reference_files.push(separator);
  state.selectedReferencePath = separator;
  state.selectedFileId = null;
  state.selectedDocId = null;

  await saveProject();
  renderReferenceList();
  updateSelectionVisuals();
}

function renameReferenceSeparator(raw) {
  state.selectedReferencePath = raw;
  state.selectedFileId = null;
  state.selectedDocId = null;
  updateSelectionVisuals();
  openActionDialog(
    "renameSeparator",
    "Trennlinie betiteln",
    separatorTitle(raw)
  );
}

function registerRightBlankTap() {
  const now = Date.now();
  const previous = state.lastSidebarBlankTap;
  const isSecondTap = previous.area === "right" && now - previous.time <= 430;

  if (isSecondTap) {
    state.lastSidebarBlankTap = { area: null, time: 0 };
    void createReferenceSeparator();
    return true;
  }

  state.lastSidebarBlankTap = { area: "right", time: now };
  return false;
}


function resetSwipeRow(row) {
  row.style.transform = "";
  row.classList.remove("swiping", "swipe-delete-ready");
}

function attachSwipeRemoval(row, direction, onRemove) {
  row.addEventListener("pointerdown", event => {
    if (event.pointerType !== "touch") return;
    if (state.separatorDrag) return;

    const action = {
      row,
      pointerId: event.pointerId,
      startX: event.clientX,
      startY: event.clientY,
      direction,
      active: false,
      cancelled: false,
    };
    state.swipeAction = action;
  }, true);

  row.addEventListener("pointermove", event => {
    const action = state.swipeAction;
    if (!action || action.row !== row || action.pointerId !== event.pointerId) return;
    if (state.separatorDrag) return;

    const dx = event.clientX - action.startX;
    const dy = event.clientY - action.startY;
    const directed = direction === "right" ? dx : -dx;

    if (!action.active) {
      if (Math.abs(dy) > 18 && Math.abs(dy) > Math.abs(dx)) {
        action.cancelled = true;
        return;
      }
      if (directed <= 10) return;
      action.active = true;
      row.classList.add("swiping");
    }

    event.preventDefault();
    event.stopPropagation();

    const distance = Math.max(0, Math.min(150, directed));
    row.style.transform = `translateX(${direction === "right" ? distance : -distance}px)`;
    row.classList.toggle("swipe-delete-ready", distance >= 92);
  }, true);

  row.addEventListener("pointerup", event => {
    const action = state.swipeAction;
    if (!action || action.row !== row || action.pointerId !== event.pointerId) return;
    state.swipeAction = null;

    const dx = event.clientX - action.startX;
    const directed = direction === "right" ? dx : -dx;
    if (action.active && directed >= 92) {
      event.preventDefault();
      event.stopPropagation();
      void onRemove();
    } else {
      resetSwipeRow(row);
    }
  }, true);

  row.addEventListener("pointercancel", () => {
    if (state.swipeAction?.row === row) state.swipeAction = null;
    resetSwipeRow(row);
  }, true);
}

async function removeReferenceEntry(raw) {
  if (!state.project) return;
  const index = state.project.reference_files.indexOf(raw);
  if (index < 0) return;

  pushHistory();

  if (!isReferenceSeparator(raw)) {
    try {
      const item = findReferenceItem(raw) || await getItemByPath(normalizeRelative(raw));
      await moveDriveItem(item, "", plainFilename(stem(item.name)));
    } catch (error) {
      toast(error.message || String(error), 6000);
      return;
    }
  }

  state.project.reference_files.splice(index, 1);
  if (state.selectedReferencePath === raw) state.selectedReferencePath = null;
  clearSidePreviews();
  await saveProject();
  await loadFolderFiles();
  renderReferenceList();
  updateSelectionVisuals();
  toast(isReferenceSeparator(raw) ? "Trennlinie entfernt" : "Datei ins Projekt-Root verschoben");
}

async function trashFolderItem() {
  await ensureFolderPath(projectPath(TRASH_DIR));
  return graphItemByAbsolutePath(projectPath(TRASH_DIR));
}

async function uniqueTrashName(name) {
  const base = encodeGraphPath(projectPath(TRASH_DIR));
  const result = await graph(`/me/drive/root:/${base}:/children?$select=name&$top=500`);
  const used = new Set((result.value || []).map(entry => String(entry.name || "").toLocaleLowerCase("de")));
  if (!used.has(name.toLocaleLowerCase("de"))) return name;

  const dot = name.lastIndexOf(".");
  const stemName = dot > 0 ? name.slice(0, dot) : name;
  const ext = dot > 0 ? name.slice(dot) : "";
  let number = 2;
  while (used.has(`${stemName} (${number})${ext}`.toLocaleLowerCase("de"))) number += 1;
  return `${stemName} (${number})${ext}`;
}

async function moveItemToProjectTrash(item) {
  if (!item?.id) throw new Error("Datei kann nicht in den Papierkorb verschoben werden.");
  if (item._inTrash) {
    const accessToken = await token();
    const response = await fetch(`${GRAPH_BASE}/me/drive/items/${item.id}`, {
      method: "DELETE",
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    if (!response.ok && response.status !== 204) {
      const body = await response.text();
      throw new Error(`Datei konnte nicht endgültig gelöscht werden.\n${response.status} ${response.statusText}\n${body}`);
    }
    return;
  }

  const trash = await trashFolderItem();
  const name = await uniqueTrashName(item.name);
  await graph(`/me/drive/items/${item.id}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ parentReference: { id: trash.id }, name }),
  });
}

async function deleteLeftFile(item) {
  if (!item?.id || item.folder) return;

  try {
    const docMatches = (state.project?.documents || []).filter(doc =>
      basenameAny(doc.project_path).toLocaleLowerCase("de") === item.name.toLocaleLowerCase("de")
    );
    const referenceMatches = (state.project?.reference_files || []).filter(raw =>
      !isReferenceSeparator(raw) && basenameAny(raw).toLocaleLowerCase("de") === item.name.toLocaleLowerCase("de")
    );

    if (docMatches.length && !item._inTrash) {
      for (const doc of docMatches) await moveClockDocToFolder(doc, TRASH_DIR);
    } else {
      await moveItemToProjectTrash(item);
    }

    if (state.project) {
      if (docMatches.length) {
        const ids = new Set(docMatches.map(doc => doc.id));
        state.project.documents = state.project.documents.filter(doc => !ids.has(doc.id));
        if (state.selectedDocId && ids.has(state.selectedDocId)) state.selectedDocId = null;
      }
      if (referenceMatches.length) {
        const removed = new Set(referenceMatches);
        state.project.reference_files = state.project.reference_files.filter(raw => !removed.has(raw));
        if (state.selectedReferencePath && removed.has(state.selectedReferencePath)) state.selectedReferencePath = null;
      }
      if (docMatches.length || referenceMatches.length) await saveProject();
    }

    if (state.selectedFileId === item.id) state.selectedFileId = null;
    clearSidePreviews();
    await loadFolderFiles();
    renderAll();
    toast(item._inTrash ? "Datei endgültig gelöscht" : "Datei in den Projekt-Papierkorb verschoben");
  } catch (error) {
    toast(error.message || String(error), 6000);
    renderFileList();
  }
}

function beginReferenceLongPress(event, row, raw) {
  if (event.pointerType !== "touch") return;
  if (state.separatorDrag) return;

  const action = {
    row,
    raw,
    pointerId: event.pointerId,
    startX: event.clientX,
    startY: event.clientY,
    active: false,
    timer: null,
    moveHandler: null,
    upHandler: null,
    cancelHandler: null,
  };

  action.timer = setTimeout(() => {
    if (state.swipeAction?.active) return;

    action.active = true;
    state.separatorDrag = action;
    state.swipeAction = null;
    row.classList.add("reference-dragging");

    try {
      row.setPointerCapture(event.pointerId);
    } catch {}
  }, 560);

  action.moveHandler = moveEvent => {
    if (moveEvent.pointerId !== action.pointerId) return;

    const distance = Math.hypot(
      moveEvent.clientX - action.startX,
      moveEvent.clientY - action.startY
    );

    if (!action.active) {
      if (distance > 12) clearTimeout(action.timer);
      return;
    }

    moveEvent.preventDefault();
    moveEvent.stopPropagation();

    const otherRows = [...el.referenceList.querySelectorAll(".reference-row")]
      .filter(candidate => candidate !== row);

    let before = null;
    for (const candidate of otherRows) {
      const rect = candidate.getBoundingClientRect();
      if (moveEvent.clientY < rect.top + rect.height / 2) {
        before = candidate;
        break;
      }
    }

    if (before) {
      el.referenceList.insertBefore(row, before);
    } else {
      el.referenceList.append(row);
    }

    const listRect = el.referenceList.getBoundingClientRect();
    const edge = 54;
    if (moveEvent.clientY < listRect.top + edge) {
      el.referenceList.scrollTop -= 12;
    } else if (moveEvent.clientY > listRect.bottom - edge) {
      el.referenceList.scrollTop += 12;
    }
  };

  const finish = async (finishEvent, cancelled = false) => {
    if (finishEvent.pointerId !== action.pointerId) return;
    clearTimeout(action.timer);

    window.removeEventListener("pointermove", action.moveHandler, true);
    window.removeEventListener("pointerup", action.upHandler, true);
    window.removeEventListener("pointercancel", action.cancelHandler, true);

    if (!action.active) return;

    finishEvent.preventDefault();
    finishEvent.stopPropagation();

    row.classList.remove("reference-dragging");
    state.separatorDrag = null;
    state.suppressReferenceClickUntil = Date.now() + 450;

    try {
      if (row.hasPointerCapture?.(action.pointerId)) {
        row.releasePointerCapture(action.pointerId);
      }
    } catch {}

    if (cancelled) {
      renderReferenceList();
      return;
    }

    const ordered = [...el.referenceList.querySelectorAll(".reference-row")]
      .map(node => node.dataset.referenceSeparator || node.dataset.referencePath)
      .filter(Boolean);

    state.project.reference_files = ordered;
    await saveProject();
    renderReferenceList();
    toast(isReferenceSeparator(raw) ? "Trennlinie verschoben" : "Datei verschoben");
  };

  action.upHandler = event => { void finish(event, false); };
  action.cancelHandler = event => { void finish(event, true); };

  window.addEventListener("pointermove", action.moveHandler, true);
  window.addEventListener("pointerup", action.upHandler, true);
  window.addEventListener("pointercancel", action.cancelHandler, true);
}


function renderReferenceList() {
  el.referenceList.replaceChildren();
  if (!state.project) return;

  for (const raw of state.project.reference_files || []) {
    const row = document.createElement("div");
    row.className = "reference-row";
    if (isReferenceSeparator(raw)) {
      row.classList.add("separator");
      row.dataset.referenceSeparator = raw;
      row.innerHTML = `
        <span class="separator-line" aria-hidden="true"></span>
        <span class="separator-title"></span>
        <span class="separator-line" aria-hidden="true"></span>
      `;
      row.querySelector(".separator-title").textContent = separatorTitle(raw) || "— —";

      if (state.selectedReferencePath === raw) row.classList.add("selected");

      row.addEventListener("click", event => {
        event.preventDefault();
        event.stopPropagation();
        if (Date.now() < state.suppressReferenceClickUntil) return;
        state.selectedReferencePath = raw;
        state.selectedFileId = null;
        state.selectedDocId = null;
        updateSelectionVisuals();
      });

      row.querySelector(".separator-title").addEventListener("dblclick", event => {
        event.preventDefault();
        event.stopPropagation();
        renameReferenceSeparator(raw);
      });

      row.querySelector(".separator-title").addEventListener("pointerup", event => {
        if (event.pointerType !== "touch" || state.separatorDrag) return;
        event.preventDefault();
        event.stopPropagation();

        const now = Date.now();
        const key = `separator:${raw}`;
        const previous = state.lastSidebarFileTap;
        const isSecondTap = previous.key === key && now - previous.time <= 430;

        if (isSecondTap) {
          state.lastSidebarFileTap = { key: null, time: 0 };
          renameReferenceSeparator(raw);
        } else {
          state.lastSidebarFileTap = { key, time: now };
          state.selectedReferencePath = raw;
          state.selectedFileId = null;
          state.selectedDocId = null;
          updateSelectionVisuals();
        }
      }, true);

      row.addEventListener("pointerdown", event => {
        beginReferenceLongPress(event, row, raw);
      }, true);

      attachSwipeRemoval(row, "left", async () => {
        await removeReferenceEntry(raw);
      });
    } else {
      const name = basenameAny(raw);
      const [label, cls] = fileIconClass(name);
      row.innerHTML = `<span class="file-icon ${cls}">${label}</span><span class="file-name"></span>`;
      applyTitleColorMarkup(row.querySelector(".file-name"), stem(name));
      row.dataset.referencePath = raw;

      if (state.selectedReferencePath === raw) row.classList.add("selected");

      row.addEventListener("click", async event => {
        event.preventDefault();
        event.stopPropagation();
        if (Date.now() < state.suppressReferenceClickUntil) return;

        state.selectedReferencePath = raw;
        state.selectedFileId = null;
        state.selectedDocId = null;
        updateSelectionVisuals();

        await activateSidebarFile(
          `right:${raw}`,
          async () => {
            try {
              await openReferenceEditor(raw);
            } catch (error) {
              toast(error.message, 6000);
            }
          },
          async () => {
            await showReferencePreview(raw, "left");
          }
        );
      });

      row.addEventListener("dblclick", event => {
        event.preventDefault();
        event.stopPropagation();
      });

      row.addEventListener("contextmenu", event => {
        event.preventDefault();
        state.selectedReferencePath = raw;
        state.selectedFileId = null;
        state.selectedDocId = null;
        updateSelectionVisuals();
        showContextMenu(event.clientX, event.clientY, null);
      });

      row.addEventListener("pointerdown", event => {
        beginReferenceLongPress(event, row, raw);
      }, true);

      attachSwipeRemoval(row, "left", async () => {
        await removeReferenceEntry(raw);
      });
    }
    el.referenceList.append(row);
  }
}

async function searchExactFile(name) {
  const escaped = name.replace(/'/g, "''");
  const result = await graph(`/me/drive/root/search(q='${encodeURIComponent(escaped)}')?$select=id,name,webUrl,@microsoft.graph.downloadUrl,file&$top=50`);
  return (result.value || []).find(item => item.name.toLocaleLowerCase("de") === name.toLocaleLowerCase("de")) || null;
}

async function openReference(raw) {
  try {
    const name = basenameAny(raw);
    const item = findReferenceItem(raw) || await searchExactFile(name);
    if (!item) throw new Error(`Datei in OneDrive nicht gefunden: ${name}`);
    if (item.webUrl) window.open(item.webUrl, "_blank", "noopener");
  } catch (error) {
    toast(error.message, 5000);
  }
}



function findDocByName(name) {
  const normalized = String(name || "").toLocaleLowerCase("de");
  return state.project?.documents?.find(
    doc => basenameAny(doc.project_path).toLocaleLowerCase("de") === normalized
  ) || null;
}

function findReferenceItem(raw) {
  const relative = normalizeRelative(raw).toLocaleLowerCase("de");
  const name = basenameAny(raw).toLocaleLowerCase("de");
  return state.folderItems.find(item =>
    String(item._relative || "").toLocaleLowerCase("de") === relative
  ) || state.folderItems.find(item =>
    item.name?.toLocaleLowerCase("de") === name
  ) || null;
}

async function openReferenceEditor(raw) {
  const name = basenameAny(raw);
  const item = findReferenceItem(raw) || await searchExactFile(name);
  if (!item) throw new Error(`Datei in OneDrive nicht gefunden: ${name}`);

  const ext = suffix(name);
  const editable = [".txt", ".md"].includes(ext);

  state.editorDoc = {
    id: `reference:${raw}`,
    title: stem(name),
    suffix: ext,
    source_type: "reference",
    project_path: name,
    _referenceRaw: raw,
    _webItem: item,
    _relative: null,
  };

  el.editorTitle.textContent = stem(name);
  el.editorMeta.textContent = "";
  el.editorText.classList.toggle("hidden", !editable);
  el.docxMessage.classList.toggle("hidden", editable);
  el.saveEditorBtn.classList.toggle("hidden", !editable);

  if (editable) {
    let response;
    if (item["@microsoft.graph.downloadUrl"]) {
      response = await fetch(item["@microsoft.graph.downloadUrl"]);
    } else {
      const accessToken = await token();
      response = await fetch(`${GRAPH_BASE}/me/drive/items/${item.id}/content`, {
        headers: { Authorization: `Bearer ${accessToken}` },
      });
    }
    if (!response.ok) throw new Error(`Datei konnte nicht geladen werden: ${name}`);
    el.editorText.value = await response.text();
  } else {
    el.editorText.value = "";
  }

  el.editorDialog.showModal();
}

async function saveReferenceEditor(doc) {
  const value = el.editorText.value.replace(/\r\n?/g, "\n");
  const accessToken = await token();
  const response = await fetch(`${GRAPH_BASE}/me/drive/items/${doc._webItem.id}/content`, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "text/plain; charset=utf-8",
    },
    body: value,
  });
  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Datei konnte nicht gespeichert werden.\n${response.status} ${response.statusText}\n${body}`);
  }
}

async function openSelectedItem() {
  const doc = selectedDoc();
  if (doc) { await openDocumentEditor(doc); return true; }
  if (state.selectedReferencePath) { await openReferenceEditor(state.selectedReferencePath); return true; }
  if (state.selectedFileId) {
    const item = state.folderItems.find(entry => entry.id === state.selectedFileId);
    const matching = item ? findDocByName(item.name) : null;
    if (matching) { selectDoc(matching); await openDocumentEditor(matching); return true; }
  }
  return false;
}

function eventDocumentNode(event) {
  return event.composedPath().find(node => node?.classList?.contains?.("doc-node")) || null;
}

function updateSelectionVisuals() {
  for (const node of el.documentLayer.querySelectorAll(".doc-node")) node.classList.toggle("selected", node.dataset.docId === state.selectedDocId);
  for (const row of el.fileList.querySelectorAll(".file-row")) row.classList.toggle("selected", row.dataset.fileItemId === state.selectedFileId);
  for (const row of el.referenceList.querySelectorAll(".reference-row")) row.classList.toggle("selected", row.dataset.referencePath === state.selectedReferencePath);
  const hasDoc = Boolean(selectedDoc());
  el.renameBtn.disabled = !(hasDoc || state.selectedReferencePath);
  el.deleteBtn.disabled = !hasDoc;
}


function endDocumentDrag(pointerId = null) {
  const drag = state.dragging;
  if (!drag) return null;
  if (pointerId !== null && drag.pointerId !== pointerId) return null;

  try {
    drag.node?.classList.remove("dragging");
  } catch {}

  try {
    if (el.clockCanvas.hasPointerCapture?.(drag.pointerId)) {
      el.clockCanvas.releasePointerCapture(drag.pointerId);
    }
  } catch {}

  state.dragging = null;
  return drag;
}

function selectedDoc() {
  return state.project?.documents?.find(doc => doc.id === state.selectedDocId) || null;
}

function selectDoc(doc) {
  state.selectedDocId = doc?.id || null;
  state.selectedFileId = null;
  state.selectedReferencePath = null;
  updateSelectionVisuals();
}

function applyViewZoom() {
  state.viewZoom = Math.max(.55, Math.min(2.2, state.viewZoom));
  localStorage.setItem("buchuhr.viewZoom", String(state.viewZoom));
  localStorage.setItem("buchuhr.viewPanX", String(state.viewPanX));
  localStorage.setItem("buchuhr.viewPanY", String(state.viewPanY));
  el.clockCanvas.setAttribute("viewBox", "-620 -620 1240 1240");
  const transform = `translate(${state.viewPanX} ${state.viewPanY}) scale(${state.viewZoom})`;
  for (const layer of [el.clockLayer, el.progressLayer, el.rasterTitleLayer, el.documentLayer]) {
    layer.setAttribute("transform", transform);
  }
}

function zoomBy(factor) {
  state.viewZoom *= factor;
  applyViewZoom();
}

function fitClock() {
  state.viewZoom = 1;
  state.viewPanX = 0;
  state.viewPanY = 0;
  applyViewZoom();
}

function hideContextMenu() {
  el.contextMenu.classList.add("hidden");
}

function showContextMenu(x, y, doc = null) {
  if (doc) selectDoc(doc);
  const hasDoc = Boolean(selectedDoc());
  const hasReference = Boolean(state.selectedReferencePath && !isReferenceSeparator(state.selectedReferencePath));
  for (const button of el.contextMenu.querySelectorAll("button")) {
    const action = button.dataset.action;
    let hidden = false;
    if (["open", "rename"].includes(action)) hidden = !(hasDoc || hasReference);
    if (action === "remove") hidden = !hasDoc;
    if (action === "titleColors") hidden = !(hasDoc || hasReference);
    button.classList.toggle("hidden", hidden);
  }
  el.contextMenu.style.left = `${Math.min(x, innerWidth-240)}px`;
  el.contextMenu.style.top = `${Math.min(y, innerHeight-340)}px`;
  el.contextMenu.classList.remove("hidden");
}

function openActionDialog(mode, title, value = "") {
  state.actionMode = mode;
  el.actionDialogTitle.textContent = title;
  el.actionDialogInput.value = value;
  el.actionDialogText.value = "";
  el.actionDialogInput.type = mode === "norm" ? "number" : "text";
  el.actionDialogInput.classList.toggle("hidden", mode === "newText" || mode === "colors");
  el.newTextTitleLabel.classList.toggle("hidden", mode !== "newText");
  el.newTextTitleInput.value = "";
  el.actionDialogText.classList.toggle("hidden", mode !== "newText");
  el.actionColorFields.classList.toggle("hidden", mode !== "colors");
  if (mode === "colors") {
    el.backgroundColorInput.value = state.project?.background_color || "#17191f";
    el.clockColorInput.value = state.project?.clock_color || "#f4f4f4";
    el.progressColorInput.value = state.project?.progress_color || "#2e7df6";
  }
  el.actionDialog.showModal();
}

function createNewText() {
  openActionDialog("newText", "Neue Textdatei");
}

function renameSelected() {
  const doc = selectedDoc();
  if (doc) {
    openActionDialog("rename", "Datei umbenennen", doc.title || "");
    return;
  }
  if (state.selectedReferencePath) {
    if (isReferenceSeparator(state.selectedReferencePath)) {
      renameReferenceSeparator(state.selectedReferencePath);
    } else {
      openActionDialog("renameReference", "Stehsatzdatei umbenennen", stem(basenameAny(state.selectedReferencePath)));
    }
  }
}

async function removeSelected() {
  const doc = selectedDoc();
  if (!doc || !confirm(`„${doc.title}“ in den Papierkorb verschieben?`)) return;
  pushHistory();
  try {
    await moveClockDocToFolder(doc, TRASH_DIR);
    clearSidePreviews();
    toast("Datei in den Projekt-Papierkorb verschoben");
  } catch (error) {
    toast(error.message || String(error), 6000);
  }
}

function editRasterTitle(second = 0) {
  const minute = Math.round(second / 60) % 60;
  state.actionMinute = minute;
  openActionDialog("raster", `Rastertitel ${minute === 0 ? 60 : minute}`, state.project?.raster_titles?.[String(minute)] || "");
}

function editNormPages() {
  openActionDialog("norm", "Normseiten", String(state.project?.norm_pages || 381));
}

function editColors() {
  openActionDialog("colors", "Farben");
}

async function saveActionDialog() {
  if (!state.project) return;
  pushHistory();

  if (state.actionMode === "rename") {
    const doc = selectedDoc();
    const title = el.actionDialogInput.value.trim();
    if (doc && title) {
      await renameClockDocForState(doc, title, doc.start_second);
    }
  } else if (state.actionMode === "renameReference") {
    const title = el.actionDialogInput.value.trim();
    if (title && state.selectedReferencePath) {
      const raw = normalizeRelative(state.selectedReferencePath);
      const item = findReferenceItem(raw) || await getItemByPath(raw);
      const newName = await uniqueNameInFolder(REFERENCE_DIR, `${safeProjectStem(title)}.txt`, item.id);
      const moved = await moveDriveItem(item, REFERENCE_DIR, newName);
      const updated = `${REFERENCE_DIR}/${moved.name}`;
      const index = state.project.reference_files.indexOf(state.selectedReferencePath);
      if (index >= 0) state.project.reference_files[index] = updated;
      state.selectedReferencePath = updated;
    }
  } else if (state.actionMode === "renameSeparator") {
    const title = el.actionDialogInput.value.trim();
    const raw = state.selectedReferencePath;
    if (raw && isReferenceSeparator(raw)) {
      const updated = `${SEPARATOR_PREFIX}${title}`;
      const index = state.project.reference_files.indexOf(raw);
      if (index >= 0) state.project.reference_files[index] = updated;
      state.selectedReferencePath = updated;
    }
  } else if (state.actionMode === "norm") {
    state.project.norm_pages = Math.max(1, Math.round(Number(el.actionDialogInput.value) || 1));
  } else if (state.actionMode === "raster") {
    state.project.raster_titles ||= {};
    const value = el.actionDialogInput.value.trim();
    if (value) state.project.raster_titles[String(state.actionMinute)] = value;
    else delete state.project.raster_titles[String(state.actionMinute)];
  } else if (state.actionMode === "colors") {
    state.project.background_color = el.backgroundColorInput.value;
    state.project.clock_color = el.clockColorInput.value;
    state.project.progress_color = el.progressColorInput.value;
  } else if (state.actionMode === "newText") {
    const text = el.actionDialogText.value.replace(/\r\n?/g, "\n").trim();
    if (!text) return toast("Bitte Text eingeben.");
    const typedTitle = el.newTextTitleInput.value.trim();
    const title = (typedTitle || text.split(/\n/)[0].trim() || "Neue Textdatei").slice(0, 80);
    const safe = title.replace(/[<>:"/\\|?*]/g, "_");
    const filename = `00.00 ${safe}.txt`;
    await uploadByPath(`${FILES_DIR}/${filename}`, text + "\n", "text/plain; charset=utf-8");
    state.project.documents.push({
      id: crypto.randomUUID().replaceAll("-",""),
      title,
      source_type: "project_text",
      original_path: "",
      project_path: `${FILES_DIR}/${filename}`,
      start_second: 0,
      character_count: text.length,
      text_cache_path: `${FILES_DIR}/${filename}`,
      suffix: ".txt",
      is_on_clock: true,
      original_mtime_ns: 0
    });
  }

  await saveProject();
  el.actionDialog.close();
  renderAll();
  await loadFolderFiles().catch(console.warn);
}

function renderAll() {
  document.documentElement.style.setProperty("--bg", state.project?.background_color || "#17191f");
  renderClock();
  renderProgress();
  renderRasterTitles();
  renderDocuments();
  renderReferenceList();
  updateHistoryButtons();
  applyViewZoom();
}

async function getDocumentContent(doc) {
  const relative = projectCopyRelative(doc);
  const { item, response } = await downloadByPath(relative);
  return { item, text: await response.text(), relative };
}

async function openDocumentEditor(doc) {
  state.selectedDocId = doc.id;
  state.editorDoc = doc;
  el.editorTitle.textContent = doc.title || stem(basenameAny(doc.project_path));
  el.editorMeta.textContent = `${Number(doc.character_count || 0).toLocaleString("de-DE")} Zeichen`;
  const editable = true;
  el.editorText.classList.toggle("hidden", !editable);
  el.docxMessage.classList.toggle("hidden", editable);
  el.saveEditorBtn.classList.toggle("hidden", !editable);

  try {
    if (editable) {
      const content = await getDocumentContent(doc);
      doc._webItem = content.item;
      doc._relative = content.relative;
      el.editorText.value = content.text;
    } else {
      el.editorText.value = "";
      doc._webItem = await getItemByPath(projectCopyRelative(doc));
    }
    el.editorDialog.showModal();
  } catch (error) {
    toast(error.message, 6000);
  }
}

async function saveEditor() {
  const doc = state.editorDoc;
  if (!doc) return;
  try {
    if (doc.source_type === "reference") {
      await saveReferenceEditor(doc);
      el.editorDialog.close();
      toast("Datei gespeichert");
      return;
    }

    pushHistory();
    const value = el.editorText.value.replace(/\r\n?/g, "\n");
    const uploaded = await uploadByPath(
      doc._relative || projectCopyRelative(doc),
      value,
      "text/plain; charset=utf-8",
      doc._webItem?.eTag || ""
    );
    doc._webItem = uploaded;
    doc.character_count = value.length;
    await saveProject();
    renderAll();
    el.editorDialog.close();
    toast("Datei gespeichert");
  } catch (error) {
    toast(error.message, 6000);
  }
}

async function openExternal() {
  const doc = state.editorDoc;
  if (!doc) return;
  try {
    const item = doc._webItem || await getItemByPath(projectCopyRelative(doc));
    if (item.webUrl) window.open(item.webUrl, "_blank", "noopener");
  } catch (error) {
    toast(error.message, 5000);
  }
}

function applySidebarState() {
  document.body.classList.toggle("sidebars-hidden", !state.sidebarsVisible);
  el.appLayout.classList.toggle("sidebars-visible", state.sidebarsVisible);
  requestAnimationFrame(() => {
    applyViewZoom();
  });
}

function toggleSidebars() {
  state.sidebarsVisible = !state.sidebarsVisible;
  localStorage.setItem("buchuhr.sidebars", String(state.sidebarsVisible));
  applySidebarState();
}

function openSettings() {
  el.clientIdInput.value = state.config.clientId;
  el.tenantInput.value = state.config.tenant;
  el.folderInput.value = state.config.root;
  el.settingsDialog.showModal();
}

async function saveSettings(event) {
  event.preventDefault();
  state.config.clientId = el.clientIdInput.value.trim();
  state.config.tenant = el.tenantInput.value.trim() || "common";
  state.config.root = normalizeConfiguredPath(el.folderInput.value) || "Buch-Uhr";
  state.config.folder = "";
  localStorage.setItem("buchuhr.clientId", state.config.clientId);
  localStorage.setItem("buchuhr.tenant", state.config.tenant);
  localStorage.setItem("buchuhr.root", state.config.root);
  state.msal = null;
  state.account = null;
  clearRememberedError();
  el.settingsDialog.close();
  try {
    await showProjectChooser();
  } catch (error) {
    console.error(error);
    rememberError(error, "Verbindung nach dem Speichern der Einstellungen fehlgeschlagen");
    toast("Verbindung fehlgeschlagen – auf „Fehler“ klicken für Details.", 8000);
  }
}

// iPad: Browser-Zoom außerhalb des Canvas unterbinden.
for (const block of document.querySelectorAll('.topbar, .sidebar')) {
  block.addEventListener('touchstart', event => {
    if (event.touches.length >= 2) event.preventDefault();
  }, { passive: false });
  block.addEventListener('touchmove', event => {
    if (event.touches.length >= 2) event.preventDefault();
  }, { passive: false });
  block.addEventListener('touchend', rememberTouchEnd, { passive: false });
}

document.addEventListener('gesturestart', event => {
  event.preventDefault();
}, { passive: false });
document.addEventListener('gesturechange', event => {
  event.preventDefault();
}, { passive: false });
document.addEventListener('gestureend', event => {
  event.preventDefault();
}, { passive: false });

el.fileList.addEventListener('dblclick', event => {
  if (event.target.closest('.file-row')) return;
  event.preventDefault();
  createNewText();
});

el.fileList.addEventListener('pointerup', event => {
  if (event.pointerType !== 'touch') return;
  if (event.target.closest('.file-row')) return;
  event.preventDefault();
  registerSidebarBlankTap('left');
}, true);

el.referenceAddBtn.addEventListener("click", () => createReferenceTextFile());

el.referenceList.addEventListener("dblclick", event => {
  if (event.target.closest(".reference-row")) return;
  event.preventDefault();
  void createReferenceSeparator();
});

el.referenceList.addEventListener("pointerup", event => {
  if (event.pointerType !== "touch") return;
  if (event.target.closest(".reference-row")) return;
  event.preventDefault();
  event.stopPropagation();
  registerRightBlankTap();
}, true);

el.leftPreviewClose.addEventListener("click", clearSidePreviews);
el.rightPreviewClose.addEventListener("click", clearSidePreviews);

el.newTextBtn.addEventListener("click", createNewText);
el.renameBtn.addEventListener("click", renameSelected);
el.deleteBtn.addEventListener("click", removeSelected);
el.zoomOutBtn.addEventListener("click", () => zoomBy(.86));
el.zoomInBtn.addEventListener("click", () => zoomBy(1.16));
el.fitBtn.addEventListener("click", fitClock);
el.actionDialogSaveBtn.addEventListener("click", saveActionDialog);
el.titleColorPlus.addEventListener("click", () => addTitleColorRuleRow("", "#f4f4f4"));
el.titleColorSave.addEventListener("click", () => saveTitleColorRules().catch(error => toast(error.message || String(error), 6000)));

el.contextMenu.addEventListener("click", async event => {
  const action = event.target.closest("button")?.dataset.action;
  if (!action) return;
  hideContextMenu();
  if (action === "open" && selectedDoc()) openDocumentEditor(selectedDoc());
  else if (action === "open" && state.selectedReferencePath) openReferenceEditor(state.selectedReferencePath);
  else if (action === "rename") renameSelected();
  else if (action === "remove") await removeSelected();
  else if (action === "newText") createNewText();
  else if (action === "raster") editRasterTitle(0);
  else if (action === "norm") editNormPages();
  else if (action === "colors") editColors();
  else if (action === "titleColors") openTitleColorDialog();
  else if (action === "fit") fitClock();
});
document.addEventListener("pointerdown", event => {
  if (!el.contextMenu.contains(event.target)) hideContextMenu();
});

el.sidebarToggle.addEventListener("click", toggleSidebars);
el.settingsBtn.addEventListener("click", openSettings);
el.syncStatus.addEventListener("click", showErrorDetails);
el.errorDetailsClose.addEventListener("click", () => el.errorDialog.close());
el.projectsBtn.addEventListener("click", () => showProjectChooser().catch(error => toast(error.message || String(error), 6000)));
el.openProjectBtn.addEventListener("click", () => openChosenProject().catch(error => toast(error.message || String(error), 6000)));
el.newProjectBtn.addEventListener("click", () => createProject().catch(error => toast(error.message || String(error), 6000)));
el.syncBtn.addEventListener("click", syncNow);
el.refreshFilesBtn.addEventListener("click", loadFolderFiles);
el.undoBtn.addEventListener("click", undo);
el.redoBtn.addEventListener("click", redo);
el.saveSettingsBtn.addEventListener("click", saveSettings);
el.saveEditorBtn.addEventListener("click", saveEditor);
el.openExternalBtn.addEventListener("click", openExternal);

window.addEventListener("keydown", event => {
  const editable = ["INPUT", "TEXTAREA"].includes(document.activeElement?.tagName);
  const toggleKey = event.key === "Tab" || event.key.toLowerCase() === "e";
  if (toggleKey && !editable) {
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
    toggleSidebars();
  }
}, true);

window.addEventListener("keydown", async event => {
  const editable = ["INPUT", "TEXTAREA"].includes(document.activeElement?.tagName);
  if (event.key === "Enter" && !editable) {
    const opened = await openSelectedItem();
    if (opened) { event.preventDefault(); event.stopPropagation(); event.stopImmediatePropagation(); }
  }
}, true);

document.addEventListener("keydown", event => {
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "z") {
    event.preventDefault();
    undo();
  }
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "y") {
    event.preventDefault();
    redo();
  }
  if (!["INPUT","TEXTAREA"].includes(document.activeElement?.tagName)) {
    if (event.key.toLowerCase() === "n") { event.preventDefault(); createNewText(); }
    else if (event.key === "F2") { event.preventDefault(); renameSelected(); }
    else if (event.key === "Delete" || event.key === "Backspace") { event.preventDefault(); removeSelected(); }
    else if (event.key === "+" || event.key === "=") { event.preventDefault(); zoomBy(1.16); }
    else if (event.key === "-") { event.preventDefault(); zoomBy(.86); }
    else if (event.key === "0") { event.preventDefault(); fitClock(); }
  }
});

applySidebarState();
updateHistoryButtons();
renderClock();
applyViewZoom();

if ("serviceWorker" in navigator) {
  navigator.serviceWorker.register("./sw.js?v=32.1").catch(console.error);
}

async function startApp() {
  if (state.config.clientId) {
    try {
      await ensureMsal();
      await signIn();
      await showProjectChooser();
    } catch (error) {
      console.error(error);
      if (isInteractionInProgress(error)) {
        clearStaleMsalInteraction();
        state.msal = null;
        state.account = null;
        toast("Alter Anmeldestatus wurde bereinigt. Die Seite wird neu geladen.", 2500);
        setTimeout(() => location.replace(location.origin + location.pathname), 600);
        return;
      }
      rememberError(error, "Start/OneDrive-Verbindung fehlgeschlagen");
      toast("Verbindung fehlgeschlagen – auf „Fehler“ klicken für Details.", 8000);
    }
  } else {
    el.emptyHint.classList.remove("hidden");
    setTimeout(openSettings, 200);
  }
}

startApp();
