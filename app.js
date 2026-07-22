const { PublicClientApplication, InteractionRequiredAuthError } = window.msal;

const PROJECT_FILE = "Buch-Uhr.project.json";
const FILES_DIR = "Dateien";
const CHARS_PER_PAGE = 1800;
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const SCOPES = ["User.Read", "Files.ReadWrite"];

const state = {
  config: {
    clientId: localStorage.getItem("buchuhr.clientId") || "",
    tenant: localStorage.getItem("buchuhr.tenant") || "common",
    folder: localStorage.getItem("buchuhr.folder") || "",
  },
  msal: null,
  account: null,
  project: null,
  projectETag: "",
  projectItem: null,
  folderItems: [],
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
};


function clearStaleMsalInteraction() {
  const preserved = {
    clientId: localStorage.getItem("buchuhr.clientId"),
    tenant: localStorage.getItem("buchuhr.tenant"),
    folder: localStorage.getItem("buchuhr.folder"),
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
  if (preserved.folder !== null) localStorage.setItem("buchuhr.folder", preserved.folder);
  if (preserved.sidebars !== null) localStorage.setItem("buchuhr.sidebars", preserved.sidebars);
}

function isInteractionInProgress(error) {
  const text = `${error?.errorCode || ""} ${error?.message || error || ""}`.toLowerCase();
  return text.includes("interaction_in_progress");
}

const el = Object.fromEntries([
  "appLayout", "sidebarToggle", "projectName", "syncStatus", "syncBtn", "settingsBtn",
  "settingsDialog", "settingsForm", "clientIdInput", "tenantInput", "folderInput",
  "saveSettingsBtn", "clockCanvas", "clockLayer", "progressLayer", "rasterTitleLayer",
  "documentLayer", "referenceList", "fileList", "refreshFilesBtn", "emptyHint",
  "newTextBtn", "renameBtn", "deleteBtn", "zoomOutBtn", "zoomInBtn", "fitBtn", "undoBtn", "redoBtn", "editorDialog", "editorTitle", "editorText", "editorMeta",
  "docxMessage", "saveEditorBtn", "openExternalBtn", "closeEditorBtn", "toast", "contextMenu", "actionDialog", "actionDialogTitle", "actionDialogLabel", "actionDialogInput", "actionDialogText", "newTextTitleLabel", "newTextTitleInput", "actionDialogSaveBtn", "actionColorFields", "backgroundColorInput", "clockColorInput", "progressColorInput"
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
}

function encodeGraphPath(path) {
  return path.split("/").filter(Boolean).map(encodeURIComponent).join("/");
}

function projectPath(relative = "") {
  const base = state.config.folder.replace(/^\/+|\/+$/g, "");
  return [base, relative].filter(Boolean).join("/");
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
  const name = basenameAny(doc.project_path);
  return `${FILES_DIR}/${name}`;
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

async function signIn() {
  const client = await ensureMsal();
  if (!state.account) {
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

async function loadProject() {
  if (!state.config.folder) throw new Error("Bitte den OneDrive-Projektordner eintragen.");
  setStatus("Lade Projekt …", "syncing");
  const { item, response } = await downloadByPath(PROJECT_FILE);
  const project = await response.json();
  if (!Array.isArray(project.documents)) project.documents = [];
  if (!Array.isArray(project.reference_files)) project.reference_files = [];
  if (!project.raster_titles) project.raster_titles = {};

  state.project = project;
  state.projectItem = item;
  state.projectETag = item.eTag || "";
  state.history = [];
  state.future = [];
  el.projectName.textContent = state.config.folder;
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
      JSON.stringify(state.project, null, 2),
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
  } catch (error) {
    console.error(error);
    setStatus("Fehler", "error");
    toast(error.message || String(error), 6000);
  }
}

async function loadFolderFiles() {
  const base = encodeGraphPath(projectPath());
  const rootResult = await graph(`/me/drive/root:/${base}:/children?$select=id,name,webUrl,file,folder,lastModifiedDateTime&$top=200`);
  let items = rootResult.value || [];

  try {
    const filesPath = encodeGraphPath(projectPath(FILES_DIR));
    const filesResult = await graph(`/me/drive/root:/${filesPath}:/children?$select=id,name,webUrl,file,folder,lastModifiedDateTime&$top=500`);
    items = items.concat((filesResult.value || []).map(item => ({ ...item, _inFiles: true })));
  } catch (error) {
    console.warn("Dateien-Unterordner nicht lesbar", error);
  }

  state.folderItems = items;
  renderFileList();
}

function fileIconClass(name) {
  const ext = suffix(name);
  if (ext === ".docx") return ["W", "word"];
  if (ext === ".md") return ["MD", "md"];
  return ["≡", "text"];
}

function makeFileRow(item) {
  const row = document.createElement("div");
  row.className = `file-row ${item.id === state.selectedFileId ? "selected" : ""}`;
  const [label, cls] = fileIconClass(item.name);
  row.innerHTML = `<span class="file-icon ${cls}">${label}</span><span class="file-name"></span>`;
  row.querySelector(".file-name").textContent = item.folder ? item.name : stem(item.name);

  const matchingDoc = () => findDocByName(item.name);

  row.addEventListener("click", event => {
    event.preventDefault();
    state.selectedFileId = item.id || null;
    state.selectedReferencePath = null;
    const doc = matchingDoc();
    state.selectedDocId = doc?.id || null;
    renderFileList();
    renderReferenceList();
    renderDocuments();
    el.renameBtn.disabled = !doc;
    el.deleteBtn.disabled = !doc;
  });

  row.addEventListener("dblclick", async event => {
    event.preventDefault();
    const doc = matchingDoc();
    if (doc) {
      selectDoc(doc);
      await openDocumentEditor(doc);
    }
  });

  row.addEventListener("contextmenu", event => {
    event.preventDefault();
    state.selectedFileId = item.id || null;
    state.selectedReferencePath = null;
    const doc = matchingDoc();
    state.selectedDocId = doc?.id || null;
    renderFileList();
    renderReferenceList();
    renderDocuments();
    showContextMenu(event.clientX, event.clientY, doc);
  });

  return row;
}

function renderFileList() {
  el.fileList.replaceChildren();
  const items = [...state.folderItems]
    .filter(item => item.name !== PROJECT_FILE)
    .sort((a, b) => Number(Boolean(b.folder)) - Number(Boolean(a.folder)) || a.name.localeCompare(b.name, "de"));

  for (const item of items) {
    el.fileList.append(makeFileRow(item));
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
    el.clockLayer.append(svg("line", {
      x1: a.x, y1: a.y, x2: b.x, y2: b.y,
      class: "tick", stroke: clockColor,
      "stroke-width": major ? 6 : five ? 3.2 : 2
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
      class: "progress-arc",
      stroke: color,
      "data-doc-id": doc.id,
    });
    path.addEventListener("pointerdown", event => beginDocDrag(event, doc, path));
    path.addEventListener("click", event => {
      if (!state.dragging?.moved) openDocumentEditor(doc);
      event.stopPropagation();
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
  for (const [minuteText, title] of Object.entries(titles)) {
    const minute = Number(minuteText) % 60;
    const major = minute % 15 === 0;
    const p = polar(minute * 60, major ? 255 : 275);
    const tick = polar(minute * 60, major ? 325 : 338);
    el.rasterTitleLayer.append(svg("line", {
      x1: tick.x, y1: tick.y, x2: p.x, y2: p.y, class: "connector"
    }));
    const textNode = svg("text", {
      x: p.x, y: p.y, class: `raster-title ${major ? "major" : ""}`
    });
    const lines = wrapSvgText(String(title), major ? 25 : 21);
    lines.forEach((line, index) => {
      const tspan = svg("tspan", {
        x: p.x,
        dy: index === 0 ? `${-(lines.length - 1) * 10}px` : "22px"
      }, line);
      textNode.append(tspan);
    });
    el.rasterTitleLayer.append(textNode);
  }
}

function renderDocuments() {
  el.documentLayer.replaceChildren();
  if (!state.project) return;
  const color = state.project.clock_color || "#f4f4f4";

  for (const doc of state.project.documents.filter(d => d.is_on_clock !== false)) {
    const p = polar(Number(doc.start_second || 0), 485);
    const right = Math.cos(p.angle) >= 0;
    const g = svg("g", {
      class: `doc-node ${doc.id === state.selectedDocId ? "selected" : ""}`,
      transform: `translate(${p.x} ${p.y})`,
      "data-doc-id": doc.id
    });

    const icon = svg("g", { class: `doc-icon ${doc.suffix === ".docx" ? "word" : doc.suffix === ".md" ? "md" : ""}` });
    icon.append(svg("rect", { x: -19, y: -19, width: 38, height: 38, rx: 4 }));
    const label = doc.suffix === ".docx" ? "W" : doc.suffix === ".md" ? "MD" : "≡";
    icon.append(svg("text", { x: 0, y: 1 }, label));
    g.append(icon);

    const title = svg("text", {
      x: right ? 29 : -29, y: 0,
      "text-anchor": right ? "start" : "end",
      class: "doc-title",
      fill: color
    }, doc.title || stem(basenameAny(doc.project_path)));
    g.append(title);

    g.addEventListener("pointerdown", event => {
      selectDoc(doc);
      beginDocDrag(event, doc, g);
      const hold = setTimeout(() => {
        if (state.dragging && !state.dragging.moved) showContextMenu(event.clientX, event.clientY, doc);
      }, 650);
      g.addEventListener("pointerup", () => clearTimeout(hold), {once:true});
      g.addEventListener("pointercancel", () => clearTimeout(hold), {once:true});
    });
    g.addEventListener("contextmenu", event => {
      event.preventDefault();
      showContextMenu(event.clientX, event.clientY, doc);
    });
    g.addEventListener("click", event => {
      selectDoc(doc);
      if (!state.dragging?.moved) openDocumentEditor(doc);
      event.stopPropagation();
    });
    el.documentLayer.append(g);
  }
}

function beginDocDrag(event, doc, node) {
  event.preventDefault();
  event.stopPropagation();
  pushHistory();
  state.dragging = {
    pointerId: event.pointerId,
    doc,
    node,
    moved: false,
    startX: event.clientX,
    startY: event.clientY,
  };
  node.classList.add("dragging");
  el.clockCanvas.setPointerCapture(event.pointerId);
}

el.clockCanvas.addEventListener("contextmenu", event => {
  event.preventDefault();
  showContextMenu(event.clientX, event.clientY);
});
el.clockCanvas.addEventListener("dblclick", event => editRasterTitle(secondFromSvgEvent(event)));

el.clockCanvas.addEventListener("pointermove", event => {
  const drag = state.dragging;
  if (!drag || drag.pointerId !== event.pointerId) return;
  if (Math.hypot(event.clientX - drag.startX, event.clientY - drag.startY) > 5) drag.moved = true;
  const raw = secondFromSvgEvent(event);
  drag.doc.start_second = magneticSecond(raw);
  renderProgress();
  renderDocuments();
});

el.clockCanvas.addEventListener("pointerup", async event => {
  const drag = state.dragging;
  if (!drag || drag.pointerId !== event.pointerId) return;
  state.dragging = { ...drag };
  drag.node?.classList.remove("dragging");
  el.clockCanvas.releasePointerCapture(event.pointerId);
  if (drag.moved) {
    await saveProject().catch(error => toast(error.message, 6000));
  }
  setTimeout(() => { state.dragging = null; }, 0);
});

function renderReferenceList() {
  el.referenceList.replaceChildren();
  if (!state.project) return;

  for (const raw of state.project.reference_files || []) {
    const row = document.createElement("div");
    row.className = "reference-row";
    if (String(raw).startsWith("__BUCHUHR_SEPARATOR__:")) {
      row.classList.add("separator");
      row.textContent = String(raw).slice("__BUCHUHR_SEPARATOR__:".length) || "— —";
    } else {
      const name = basenameAny(raw);
      const [label, cls] = fileIconClass(name);
      row.innerHTML = `<span class="file-icon ${cls}">${label}</span><span class="file-name"></span>`;
      row.querySelector(".file-name").textContent = stem(name);

      if (state.selectedReferencePath === raw) row.classList.add("selected");

      row.addEventListener("click", event => {
        event.preventDefault();
        state.selectedReferencePath = raw;
        state.selectedFileId = null;
        state.selectedDocId = null;
        renderReferenceList();
        renderFileList();
        renderDocuments();
        el.renameBtn.disabled = false;
        el.deleteBtn.disabled = true;
      });

      row.addEventListener("dblclick", async event => {
        event.preventDefault();
        state.selectedReferencePath = raw;
        try {
          await openReferenceEditor(raw);
        } catch (error) {
          toast(error.message, 6000);
        }
      });

      row.addEventListener("contextmenu", event => {
        event.preventDefault();
        state.selectedReferencePath = raw;
        state.selectedFileId = null;
        state.selectedDocId = null;
        renderReferenceList();
        renderFileList();
        renderDocuments();
        showContextMenu(event.clientX, event.clientY, null);
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
    const item = await searchExactFile(name);
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
  const name = basenameAny(raw);
  return state.folderItems.find(
    item => item.name?.toLocaleLowerCase("de") === name.toLocaleLowerCase("de")
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

function selectedDoc() {
  return state.project?.documents?.find(doc => doc.id === state.selectedDocId) || null;
}

function selectDoc(doc) {
  state.selectedDocId = doc?.id || null;
  el.renameBtn.disabled = !doc;
  el.deleteBtn.disabled = !doc;
  renderDocuments();
}

function applyViewZoom() {
  state.viewZoom = Math.max(.55, Math.min(2.2, state.viewZoom));
  localStorage.setItem("buchuhr.viewZoom", String(state.viewZoom));
  el.clockCanvas.setAttribute("viewBox", "-620 -620 1240 1240");
  for (const layer of [el.clockLayer, el.progressLayer, el.rasterTitleLayer, el.documentLayer]) {
    layer.setAttribute("transform", `scale(${state.viewZoom})`);
  }
}

function zoomBy(factor) {
  state.viewZoom *= factor;
  applyViewZoom();
}

function fitClock() {
  state.viewZoom = 1;
  applyViewZoom();
}

function hideContextMenu() {
  el.contextMenu.classList.add("hidden");
}

function showContextMenu(x, y, doc = null) {
  if (doc) selectDoc(doc);
  const hasDoc = Boolean(selectedDoc());
  for (const button of el.contextMenu.querySelectorAll("button")) {
    button.classList.toggle("hidden", ["open","rename","remove"].includes(button.dataset.action) && !hasDoc);
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
    openActionDialog("renameReference", "Stehsatzdatei umbenennen", stem(basenameAny(state.selectedReferencePath)));
  }
}

async function removeSelected() {
  const doc = selectedDoc();
  if (!doc || !confirm(`„${doc.title}“ von der Uhr entfernen?`)) return;
  pushHistory();
  doc.is_on_clock = false;
  state.selectedDocId = null;
  await saveProject();
  renderAll();
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
    if (doc && title) doc.title = title;
  } else if (state.actionMode === "renameReference") {
    const title = el.actionDialogInput.value.trim();
    if (title && state.selectedReferencePath) {
      const raw = state.selectedReferencePath;
      const ext = suffix(basenameAny(raw));
      const updated = raw.replace(/[^\\/]+$/, `${title}${ext}`);
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
    const filename = `00_00 – ${safe}.txt`;
    await uploadByPath(`${FILES_DIR}/${filename}`, text + "\n", "text/plain; charset=utf-8");
    state.project.documents.push({
      id: crypto.randomUUID().replaceAll("-",""),
      title,
      source_type: "dragged_text",
      original_path: "",
      project_path: filename,
      start_second: 0,
      character_count: text.length,
      text_cache_path: filename,
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
  const editable = [".txt", ".md"].includes(String(doc.suffix).toLowerCase()) || doc.source_type === "dragged_text";
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
  el.appLayout.classList.toggle("sidebars-visible", state.sidebarsVisible);
  el.appLayout.classList.toggle("sidebars-hidden", !state.sidebarsVisible);
}

function toggleSidebars() {
  state.sidebarsVisible = !state.sidebarsVisible;
  localStorage.setItem("buchuhr.sidebars", String(state.sidebarsVisible));
  applySidebarState();
}

function openSettings() {
  el.clientIdInput.value = state.config.clientId;
  el.tenantInput.value = state.config.tenant;
  el.folderInput.value = state.config.folder;
  el.settingsDialog.showModal();
}

async function saveSettings(event) {
  event.preventDefault();
  state.config.clientId = el.clientIdInput.value.trim();
  state.config.tenant = el.tenantInput.value.trim() || "common";
  state.config.folder = el.folderInput.value.trim().replace(/^\/+|\/+$/g, "");
  localStorage.setItem("buchuhr.clientId", state.config.clientId);
  localStorage.setItem("buchuhr.tenant", state.config.tenant);
  localStorage.setItem("buchuhr.folder", state.config.folder);
  state.msal = null;
  state.account = null;
  el.settingsDialog.close();
  await syncNow();
}

el.newTextBtn.addEventListener("click", createNewText);
el.renameBtn.addEventListener("click", renameSelected);
el.deleteBtn.addEventListener("click", removeSelected);
el.zoomOutBtn.addEventListener("click", () => zoomBy(.86));
el.zoomInBtn.addEventListener("click", () => zoomBy(1.16));
el.fitBtn.addEventListener("click", fitClock);
el.actionDialogSaveBtn.addEventListener("click", saveActionDialog);

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
  else if (action === "fit") fitClock();
});
document.addEventListener("pointerdown", event => {
  if (!el.contextMenu.contains(event.target)) hideContextMenu();
});

el.sidebarToggle.addEventListener("click", toggleSidebars);
el.settingsBtn.addEventListener("click", openSettings);
el.syncBtn.addEventListener("click", syncNow);
el.refreshFilesBtn.addEventListener("click", loadFolderFiles);
el.undoBtn.addEventListener("click", undo);
el.redoBtn.addEventListener("click", redo);
el.saveSettingsBtn.addEventListener("click", saveSettings);
el.saveEditorBtn.addEventListener("click", saveEditor);
el.openExternalBtn.addEventListener("click", openExternal);

window.addEventListener("keydown", event => {
  if (event.key === "Tab" && !["INPUT","TEXTAREA"].includes(document.activeElement?.tagName)) {
    event.preventDefault();
    event.stopImmediatePropagation();
    toggleSidebars();
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
  navigator.serviceWorker.register("./sw.js").catch(console.error);
}

async function startApp() {
  if (state.config.clientId && state.config.folder) {
    try {
      await ensureMsal();
      await syncNow();
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
      setStatus("Fehler", "error");
      toast(error.message || String(error), 8000);
    }
  } else {
    el.emptyHint.classList.remove("hidden");
    setTimeout(openSettings, 200);
  }
}

startApp();
