import { normalizePriority, normalizeRecordId } from "./state-utils.js";


const STORAGE_KEY = "gameCollectionManager.v2";
const LEGACY_STORAGE_KEY = "gameCollectionManager.v1";
const BACKUP_STORAGE_KEY = "gameCollectionManager.backup.v2";
const SCHEMA_VERSION = 2;
const PLATFORM_ORDER = ["PS1", "PS2", "PS4", "DS WII"];
const PLATFORM_LOGOS = {
  ps1: "assets/logos/ps1.svg",
  ps2: "assets/logos/ps2.svg",
  ps4: "assets/logos/ps4.svg",
  dswii: "assets/logos/dswii.svg",
};

const state = {
  collection: [],
  wishlist: [],
  activeCollectionPlatform: "ALL",
  editingCollectionId: "",
  editingWishlistId: "",
};

const elements = {
  heroStats: document.getElementById("hero-stats"),
  status: document.getElementById("status"),
  exportStateButton: document.getElementById("export-state-btn"),
  importStateButton: document.getElementById("import-state-btn"),
  importStateFile: document.getElementById("import-state-file"),
  exportStateLink: document.getElementById("export-state-link"),
  tabs: [...document.querySelectorAll(".tab")],
  panels: {
    collection: document.getElementById("collection-panel"),
    wishlist: document.getElementById("wishlist-panel"),
  },
  collectionSearch: document.getElementById("collection-search"),
  collectionPlatformTabs: document.getElementById("collection-platform-tabs"),
  wishlistSearch: document.getElementById("wishlist-search"),
  wishlistPlatformFilter: document.getElementById("wishlist-platform-filter"),
  wishlistStatusFilter: document.getElementById("wishlist-status-filter"),
  wishlistResultSummary: document.getElementById("wishlist-result-summary"),
  collectionGroups: document.getElementById("collection-groups"),
  wishlistTableBody: document.querySelector("#wishlist-table tbody"),
  collectionForm: document.getElementById("collection-form"),
  collectionFormTitle: document.getElementById("collection-form-title"),
  collectionSubmitButton: document.getElementById("collection-submit-btn"),
  collectionCancelEditButton: document.getElementById("collection-cancel-edit-btn"),
  wishlistForm: document.getElementById("wishlist-form"),
  wishlistFormTitle: document.getElementById("wishlist-form-title"),
  wishlistSubmitButton: document.getElementById("wishlist-submit-btn"),
  wishlistCancelEditButton: document.getElementById("wishlist-cancel-edit-btn"),
  statusMessage: document.getElementById("status-message"),
  statusUndoButton: document.getElementById("status-undo-btn"),
  receiveDialog: document.getElementById("receive-dialog"),
  receiveForm: document.getElementById("receive-form"),
  receiveGameTitle: document.getElementById("receive-game-title"),
  receiveCancelButton: document.getElementById("receive-cancel-btn"),
  deleteDialog: document.getElementById("delete-dialog"),
  deleteDialogTitle: document.getElementById("delete-dialog-title"),
  deleteDialogMessage: document.getElementById("delete-dialog-message"),
  deleteMoveOption: document.getElementById("delete-move-option"),
  deleteMoveCheckbox: document.getElementById("delete-move-checkbox"),
  deleteCancelButton: document.getElementById("delete-cancel-btn"),
  deleteConfirmButton: document.getElementById("delete-confirm-btn"),
};

let statusTimer = null;
let currentExportBlobUrl = null;
let undoAction = null;
let pendingDeletion = null;

function normalize(input) {
  return (input || "").toString().trim().toLowerCase();
}

function normalizePlatformKey(platform) {
  return normalize(platform).replace(/[^a-z0-9]+/g, "");
}

function getLogoPath(platform) {
  return PLATFORM_LOGOS[normalizePlatformKey(platform)] || "";
}

function getPlatformAbbreviation(platform) {
  const compact = String(platform || "")
    .replaceAll("/", " ")
    .trim()
    .split(/\s+/)
    .map((chunk) => chunk[0] || "")
    .join("")
    .toUpperCase();
  return compact || "PLT";
}

function renderPlatformBadge(platform, size = "sm") {
  const label = escapeHtml(platform);
  const logoPath = getLogoPath(platform);
  const logoKey = normalizePlatformKey(platform);
  if (logoPath) {
    return `
      <span class="platform-badge platform-${size}">
        <span class="platform-logo-crop platform-logo-crop-${logoKey}">
          <img class="platform-logo platform-logo-${logoKey}" src="${logoPath}" alt="${label} logo" loading="lazy" />
        </span>
        <span>${label}</span>
      </span>
    `;
  }
  return `
    <span class="platform-badge platform-${size}">
      <span class="platform-fallback">${escapeHtml(getPlatformAbbreviation(platform))}</span>
      <span>${label}</span>
    </span>
  `;
}

function getPlatforms() {
  const byKey = new Map();
  for (const platform of PLATFORM_ORDER) {
    byKey.set(normalize(platform), platform);
  }
  for (const game of state.collection) {
    if (game.platform) byKey.set(normalize(game.platform), game.platform);
  }
  for (const wish of state.wishlist) {
    if (wish.platform) byKey.set(normalize(wish.platform), wish.platform);
  }

  const ordered = [];
  const pending = new Map(byKey);
  for (const platform of PLATFORM_ORDER) {
    const key = normalize(platform);
    if (pending.has(key)) {
      ordered.push(pending.get(key));
      pending.delete(key);
    }
  }
  const extras = [...pending.values()].sort((a, b) => a.localeCompare(b));
  return [...ordered, ...extras];
}

function newId(prefix) {
  const list = prefix === "c" ? state.collection : state.wishlist;
  let max = 0;
  for (const item of list) {
    const match = String(item.id || "").match(/\d+/);
    if (match) max = Math.max(max, Number(match[0]));
  }
  return `${prefix}${max + 1}`;
}

function isDuplicateCollection(platform, title) {
  const platformKey = normalize(platform);
  const titleKey = normalize(title);
  return state.collection.some(
    (item) => normalize(item.platform) === platformKey && normalize(item.title) === titleKey
  );
}

function statePayload() {
  return {
    schemaVersion: SCHEMA_VERSION,
    updatedAt: new Date().toISOString(),
    collection: state.collection,
    wishlist: state.wishlist,
  };
}

function saveState() {
  try {
    const payload = statePayload();
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    return true;
  } catch (error) {
    console.error(error);
    flashStatus("Changes are in memory, but browser storage could not be updated. Export a snapshot now.", true);
    return false;
  }
}

function saveImportBackup() {
  try {
    localStorage.setItem(BACKUP_STORAGE_KEY, JSON.stringify(statePayload()));
  } catch (error) {
    console.warn("Could not save import backup", error);
  }
}

function cloneCurrentData() {
  return JSON.parse(JSON.stringify({
    collection: state.collection,
    wishlist: state.wishlist,
  }));
}

function toBool(value) {
  if (typeof value === "boolean") return value;
  const text = normalize(value);
  return text === "true" || text === "1" || text === "yes" || text === "x";
}

function normalizeCollectionRow(row) {
  return {
    id: normalizeRecordId(row?.id, "c"),
    platform: String(row?.platform ?? "").trim(),
    title: String(row?.title ?? "").trim(),
    version: String(row?.version ?? "").trim(),
    cdCondition: String(row?.cdCondition ?? row?.cd_condition ?? "").trim(),
    manualCondition: String(row?.manualCondition ?? row?.manual_condition ?? "").trim(),
    price: String(row?.price ?? "").trim(),
    extra: String(row?.extra ?? "").trim(),
    note: String(row?.note ?? "").trim(),
    acquiredDate: String(row?.acquiredDate ?? row?.acquired_date ?? "").trim(),
    source: String(row?.source ?? "").trim(),
  };
}

function normalizeWishlistRow(row) {
  return {
    id: normalizeRecordId(row?.id, "w"),
    platform: String(row?.platform ?? "").trim(),
    title: String(row?.title ?? "").trim(),
    note: String(row?.note ?? "").trim(),
    inTransit: toBool(row?.inTransit ?? row?.in_transit),
    received: toBool(row?.received),
    priority: normalizePriority(row?.priority),
    targetPrice: String(row?.targetPrice ?? row?.target_price ?? "").trim(),
    orderedDate: String(row?.orderedDate ?? row?.ordered_date ?? "").trim(),
    receivedDate: String(row?.receivedDate ?? row?.received_date ?? "").trim(),
    listingUrl: String(row?.listingUrl ?? row?.listing_url ?? "").trim(),
    replacement: toBool(row?.replacement),
  };
}

function buildCollectionKey(platform, title) {
  return `${normalize(platform)}::${normalize(title)}`;
}

function normalizeStatePayload(payload, { strict = false } = {}) {
  if (!payload || typeof payload !== "object") {
    throw new Error("Invalid state format.");
  }

  if (strict && (!Array.isArray(payload.collection) || !Array.isArray(payload.wishlist))) {
    throw new Error("Snapshot must contain collection and wishlist arrays.");
  }
  if (Number(payload.schemaVersion || 1) > SCHEMA_VERSION) {
    throw new Error(`Snapshot schema ${payload.schemaVersion} is newer than this app supports.`);
  }

  const rawCollection = Array.isArray(payload.collection) ? payload.collection : [];
  const rawWishlist = Array.isArray(payload.wishlist) ? payload.wishlist : [];
  const collection = [];
  const wishlist = [];
  const seenCollection = new Set();
  const seenWishlist = new Set();

  for (const row of rawCollection) {
    const normalizedRow = normalizeCollectionRow(row);
    if (!normalizedRow.platform || !normalizedRow.title) {
      if (strict) throw new Error("A collection row is missing its platform or title.");
      continue;
    }
    const key = buildCollectionKey(normalizedRow.platform, normalizedRow.title);
    if (seenCollection.has(key)) {
      if (strict) throw new Error(`Duplicate collection entry: ${normalizedRow.title}.`);
      continue;
    }
    seenCollection.add(key);
    collection.push(normalizedRow);
  }

  for (const row of rawWishlist) {
    const normalizedRow = normalizeWishlistRow(row);
    if (!normalizedRow.platform || !normalizedRow.title) {
      if (strict) throw new Error("A wishlist row is missing its platform or title.");
      continue;
    }
    const key = buildCollectionKey(normalizedRow.platform, normalizedRow.title);
    if (normalizedRow.received) {
      if (!seenCollection.has(key)) {
        seenCollection.add(key);
        collection.push({
          id: "",
          platform: normalizedRow.platform,
          title: normalizedRow.title,
          version: "",
          cdCondition: "",
          manualCondition: "",
          price: "",
          extra: "",
          note: normalizedRow.note || "",
          acquiredDate: normalizedRow.receivedDate || "",
          source: "",
        });
      }
      continue;
    }
    if (seenWishlist.has(key)) {
      if (strict) throw new Error(`Duplicate wishlist entry: ${normalizedRow.title}.`);
      continue;
    }
    seenWishlist.add(key);
    wishlist.push({
      id: normalizedRow.id,
      platform: normalizedRow.platform,
      title: normalizedRow.title,
      note: normalizedRow.note,
      inTransit: normalizedRow.inTransit,
      received: false,
      priority: normalizedRow.priority,
      targetPrice: normalizedRow.targetPrice,
      orderedDate: normalizedRow.orderedDate,
      receivedDate: normalizedRow.receivedDate,
      listingUrl: normalizedRow.listingUrl,
      replacement: normalizedRow.replacement,
    });
  }

  collection.sort((a, b) => {
    if (a.platform !== b.platform) return a.platform.localeCompare(b.platform);
    return a.title.localeCompare(b.title);
  });
  wishlist.sort((a, b) => {
    if (a.platform !== b.platform) return a.platform.localeCompare(b.platform);
    return a.title.localeCompare(b.title);
  });

  const assignIds = (items, prefix) => {
    const used = new Set();
    let next = 1;
    return items.map((item) => {
      let id = item.id;
      if (!id || used.has(id)) {
        while (used.has(`${prefix}${next}`)) next += 1;
        id = `${prefix}${next}`;
      }
      used.add(id);
      const normalized = { ...item, id };
      return normalized;
    });
  };

  const collectionWithIds = assignIds(collection, "c");
  const ownedKeys = new Set(collectionWithIds.map((item) => buildCollectionKey(item.platform, item.title)));
  const wishlistWithIds = assignIds(wishlist, "w").map((item) => ({
    ...item,
    replacement: item.replacement || ownedKeys.has(buildCollectionKey(item.platform, item.title)),
  }));

  return {
    collection: collectionWithIds,
    wishlist: wishlistWithIds,
  };
}

function triggerJsonDownload(filename, payload) {
  if (currentExportBlobUrl) {
    URL.revokeObjectURL(currentExportBlobUrl);
    currentExportBlobUrl = null;
  }
  const blob = new Blob([`${JSON.stringify(payload, null, 2)}\n`], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  currentExportBlobUrl = url;

  if (elements.exportStateLink) {
    elements.exportStateLink.href = url;
    elements.exportStateLink.download = filename;
    elements.exportStateLink.hidden = false;
    elements.exportStateLink.textContent = `Download manually: ${filename}`;
  }

  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.rel = "noopener";
  link.style.display = "none";
  document.body.append(link);
  link.click();
  link.remove();
}

function exportStateSnapshot() {
  try {
    const normalizedState = normalizeStatePayload({
      collection: state.collection,
      wishlist: state.wishlist,
    });
    const payload = {
      schemaVersion: SCHEMA_VERSION,
      exportedAt: new Date().toISOString(),
      collection: normalizedState.collection,
      wishlist: normalizedState.wishlist,
    };
    const stamp = new Date().toISOString().replaceAll(":", "-");
    const filename = `game-collection-state-${stamp}.json`;
    triggerJsonDownload(filename, payload);
    flashStatus(`Export created. If download did not start, use the manual link.`);
  } catch (error) {
    console.error(error);
    flashStatus("Export failed. Check browser download permissions.", true);
  }
}

async function importStateSnapshot(file) {
  const text = await file.text();
  const parsed = JSON.parse(text);
  const normalizedState = normalizeStatePayload(parsed, { strict: true });
  const inTransit = normalizedState.wishlist.filter((item) => item.inTransit).length;
  const accepted = window.confirm(
    `Import ${normalizedState.collection.length} collection games and ${normalizedState.wishlist.length} wishlist games (${inTransit} in transit)?\n\nYour current browser state will be saved as a local backup.`
  );
  if (!accepted) {
    flashStatus("Import cancelled.");
    return;
  }
  const previous = cloneCurrentData();
  saveImportBackup();
  state.collection = normalizedState.collection;
  state.wishlist = normalizedState.wishlist;
  state.editingCollectionId = "";
  state.editingWishlistId = "";
  elements.collectionForm.elements.editId.value = "";
  saveState();
  renderAll();
  flashStatus(
    `Imported ${state.collection.length} collection and ${state.wishlist.length} wishlist games.`,
    false,
    () => restoreData(previous, "Import undone.")
  );
}

function flashStatus(message, isError = false, onUndo = null) {
  clearTimeout(statusTimer);
  undoAction = onUndo;
  elements.statusMessage.textContent = message;
  elements.status.style.color = isError ? "#ff8aa3" : "#65f3a2";
  elements.status.classList.add("is-visible");
  elements.statusUndoButton.hidden = !onUndo;
  statusTimer = setTimeout(() => {
    elements.statusMessage.textContent = "";
    elements.statusUndoButton.hidden = true;
    elements.status.classList.remove("is-visible");
    undoAction = null;
  }, onUndo ? 8000 : 4000);
}

function restoreData(previous, message) {
  state.collection = previous.collection;
  state.wishlist = previous.wishlist;
  resetCollectionForm();
  resetWishlistForm();
  saveState();
  renderAll();
  flashStatus(message);
}

function fillPlatformSelects() {
  const platforms = getPlatforms();
  const selects = [
    elements.wishlistPlatformFilter,
    elements.collectionForm.elements.platform,
    elements.wishlistForm.elements.platform,
  ];

  for (const select of selects) {
    const previous = select.value;
    const isFilter = select === elements.wishlistPlatformFilter;
    select.innerHTML = "";
    if (isFilter) {
      const option = document.createElement("option");
      option.value = "";
      option.textContent = "All platforms";
      select.append(option);
    }
    for (const platform of platforms) {
      const option = document.createElement("option");
      option.value = platform;
      option.textContent = platform;
      select.append(option);
    }
    if ([...select.options].some((option) => option.value === previous)) {
      select.value = previous;
    }
  }
}

function renderCollectionPlatformTabs() {
  const platforms = getPlatforms();
  if (state.activeCollectionPlatform !== "ALL") {
    const exists = platforms.some((platform) => normalize(platform) === normalize(state.activeCollectionPlatform));
    if (!exists) {
      state.activeCollectionPlatform = "ALL";
    }
  }

  const countsByPlatform = new Map();
  for (const game of state.collection) {
    const key = normalize(game.platform);
    countsByPlatform.set(key, (countsByPlatform.get(key) || 0) + 1);
  }

  const allCount = state.collection.length;
  const tabs = [{ label: "All", value: "ALL", count: allCount }];
  for (const platform of platforms) {
    tabs.push({
      label: platform,
      value: platform,
      count: countsByPlatform.get(normalize(platform)) || 0,
    });
  }

  elements.collectionPlatformTabs.innerHTML = "";
  for (const tab of tabs) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `platform-tab ${normalize(state.activeCollectionPlatform) === normalize(tab.value) ? "is-active" : ""}`;
    button.dataset.platformTab = tab.value;
    button.setAttribute("role", "tab");
    button.setAttribute(
      "aria-selected",
      normalize(state.activeCollectionPlatform) === normalize(tab.value) ? "true" : "false"
    );
    button.tabIndex = normalize(state.activeCollectionPlatform) === normalize(tab.value) ? 0 : -1;
    if (tab.value === "ALL") {
      button.innerHTML = `<span>All</span><strong>${tab.count}</strong>`;
    } else {
      button.innerHTML = `${renderPlatformBadge(tab.label, "sm")}<strong>${tab.count}</strong>`;
    }
    elements.collectionPlatformTabs.append(button);
  }
}

function renderHeroStats() {
  const wishlistInTransit = state.wishlist.filter((item) => item.inTransit).length;
  const stats = [
    { label: "Collection Games", value: state.collection.length },
    { label: "Wishlist Games", value: state.wishlist.length },
    { label: "In Transit", value: wishlistInTransit },
    { label: "Platforms", value: getPlatforms().length },
  ];

  elements.heroStats.innerHTML = "";
  for (const stat of stats) {
    const card = document.createElement("article");
    card.className = "stat-card";
    card.innerHTML = `<strong>${stat.value}</strong><span>${stat.label}</span>`;
    elements.heroStats.append(card);
  }
}

function getCollectionFiltered() {
  const query = normalize(elements.collectionSearch.value);
  const platform = normalize(state.activeCollectionPlatform);

  return state.collection.filter((game) => {
    if (platform && platform !== "all" && normalize(game.platform) !== platform) return false;
    if (!query) return true;
    const haystack = [game.title, game.version, game.note, game.extra, game.platform].map(normalize).join(" ");
    return haystack.includes(query);
  });
}

function renderCollection() {
  const filtered = getCollectionFiltered();
  const grouped = new Map();
  for (const game of filtered) {
    if (!grouped.has(game.platform)) grouped.set(game.platform, []);
    grouped.get(game.platform).push(game);
  }

  elements.collectionGroups.innerHTML = "";

  const platforms = getPlatforms();
  for (const platform of platforms) {
    const items = grouped.get(platform) || [];
    if (items.length === 0) continue;

    const card = document.createElement("article");
    card.className = "card group-card";

    const header = document.createElement("div");
    header.className = "group-title";
    header.innerHTML = `
      <h3 class="platform-heading">${renderPlatformBadge(platform, "lg")}</h3>
      <span>${items.length} game${items.length === 1 ? "" : "s"}</span>
    `;
    card.append(header);

    const wrap = document.createElement("div");
    wrap.className = "table-wrap";

    const table = document.createElement("table");
    table.className = "collection-table";
    table.innerHTML = `
      <colgroup>
        <col style="width: 24%" />
        <col style="width: 10%" />
        <col style="width: 7%" />
        <col style="width: 7%" />
        <col style="width: 8%" />
        <col style="width: 11%" />
        <col style="width: 18%" />
        <col style="width: 15%" />
      </colgroup>
      <thead>
        <tr>
          <th>Title</th>
          <th>Version</th>
          <th>CD</th>
          <th>Manual</th>
          <th>Price</th>
          <th>Extra</th>
          <th>Note</th>
          <th>Action</th>
        </tr>
      </thead>
      <tbody></tbody>
    `;

    const tbody = table.querySelector("tbody");
    for (const game of items.sort((a, b) => a.title.localeCompare(b.title))) {
      const row = document.createElement("tr");
      row.innerHTML = `
        <td data-label="Title" title="${escapeHtml(game.title)}">${escapeHtml(game.title)}</td>
        <td data-label="Version" title="${escapeHtml(game.version || "")}">${escapeHtml(game.version || "")}</td>
        <td data-label="Disc" class="mono">${escapeHtml(game.cdCondition || "")}</td>
        <td data-label="Manual" class="mono">${escapeHtml(game.manualCondition || "")}</td>
        <td data-label="Price">${escapeHtml(game.price || "")}</td>
        <td data-label="Extra" title="${escapeHtml(game.extra || "")}">${escapeHtml(game.extra || "")}</td>
        <td data-label="Note" title="${escapeHtml(game.note || "")}">${escapeHtml(game.note || "")}</td>
        <td data-label="Actions">
          <div class="row-actions">
            <button class="secondary" data-edit-collection="${escapeHtml(game.id)}" type="button" aria-label="Edit ${escapeHtml(game.title)}">Edit</button>
            <button class="danger" data-remove-collection="${escapeHtml(game.id)}" type="button" aria-label="Remove ${escapeHtml(game.title)}">Remove</button>
          </div>
        </td>
      `;
      tbody.append(row);
    }

    wrap.append(table);
    card.append(wrap);
    elements.collectionGroups.append(card);
  }

  if (!elements.collectionGroups.children.length) {
    elements.collectionGroups.innerHTML = '<p class="empty">No collection items match this filter.</p>';
  }
}

function getWishlistFiltered() {
  const query = normalize(elements.wishlistSearch.value);
  const platform = normalize(elements.wishlistPlatformFilter.value);
  const status = elements.wishlistStatusFilter.value;

  return state.wishlist.filter((item) => {
    if (platform && normalize(item.platform) !== platform) return false;
    if (status === "pending" && item.inTransit) return false;
    if (status === "transit" && !item.inTransit) return false;
    if (status === "replacement" && !item.replacement) return false;
    if (!query) return true;
    const haystack = [item.title, item.note, item.platform, item.priority].map(normalize).join(" ");
    return haystack.includes(query);
  });
}

function safeExternalUrl(value) {
  try {
    const url = new URL(value);
    return ["http:", "https:"].includes(url.protocol) ? url.href : "";
  } catch {
    return "";
  }
}

function renderWishlist() {
  const rows = getWishlistFiltered().sort((a, b) => {
    if (a.platform !== b.platform) return a.platform.localeCompare(b.platform);
    return a.title.localeCompare(b.title);
  });

  elements.wishlistResultSummary.textContent = `Showing ${rows.length} of ${state.wishlist.length} games`;

  elements.wishlistTableBody.innerHTML = "";
  for (const item of rows) {
    const row = document.createElement("tr");
    const listingUrl = safeExternalUrl(item.listingUrl);
    const note = item.replacement
      ? `${escapeHtml(item.note || "")}<span class="tag replacement-tag">Replacement</span>`
      : escapeHtml(item.note || "");
    row.innerHTML = `
      <td data-label="Platform">${renderPlatformBadge(item.platform, "sm")}</td>
      <td data-label="Title">${escapeHtml(item.title)}${listingUrl ? ` <a class="external-link" href="${escapeHtml(listingUrl)}" target="_blank" rel="noopener noreferrer" aria-label="Open listing for ${escapeHtml(item.title)}">↗</a>` : ""}</td>
      <td data-label="Note">${note}</td>
      <td data-label="Priority"><span class="tag priority-${normalizePriority(item.priority).toLowerCase()}">${escapeHtml(normalizePriority(item.priority))}</span></td>
      <td data-label="Target">${escapeHtml(item.targetPrice || "")}</td>
      <td data-label="In transit"><input class="checkbox" data-in-transit="${escapeHtml(item.id)}" type="checkbox" ${item.inTransit ? "checked" : ""} aria-label="Mark ${escapeHtml(item.title)} as in transit" /></td>
      <td data-label="Actions">
        <div class="row-actions">
          <button class="secondary" data-edit-wishlist="${escapeHtml(item.id)}" type="button" aria-label="Edit ${escapeHtml(item.title)}">Edit</button>
          <button data-receive-wishlist="${escapeHtml(item.id)}" type="button" aria-label="Receive ${escapeHtml(item.title)}">Receive</button>
          <button class="danger" data-remove-wishlist="${escapeHtml(item.id)}" type="button" aria-label="Delete ${escapeHtml(item.title)}">Delete</button>
        </div>
      </td>
    `;
    elements.wishlistTableBody.append(row);
  }

  if (!rows.length) {
    const row = document.createElement("tr");
    row.innerHTML = '<td colspan="7" class="empty">Wishlist is empty for this filter.</td>';
    elements.wishlistTableBody.append(row);
  }
}

function renderAll() {
  fillPlatformSelects();
  renderCollectionPlatformTabs();
  renderHeroStats();
  renderCollection();
  renderWishlist();
  setCollectionFormMode();
  setWishlistFormMode();
}

function setCollectionFormMode() {
  const editing = Boolean(state.editingCollectionId);
  elements.collectionFormTitle.textContent = editing ? "Edit Collection Game" : "Add To Collection";
  elements.collectionSubmitButton.textContent = editing ? "Save Changes" : "Add Game";
  elements.collectionCancelEditButton.hidden = !editing;
}

function resetCollectionForm() {
  elements.collectionForm.reset();
  state.editingCollectionId = "";
  elements.collectionForm.elements.editId.value = "";
  setCollectionFormMode();
}

function startCollectionEdit(id) {
  const game = state.collection.find((item) => item.id === id);
  if (!game) return;
  state.editingCollectionId = game.id;
  elements.collectionForm.elements.editId.value = game.id;
  elements.collectionForm.elements.platform.value = game.platform;
  elements.collectionForm.elements.title.value = game.title;
  elements.collectionForm.elements.version.value = game.version || "";
  elements.collectionForm.elements.cdCondition.value = game.cdCondition || "";
  elements.collectionForm.elements.manualCondition.value = game.manualCondition || "";
  elements.collectionForm.elements.price.value = game.price || "";
  elements.collectionForm.elements.extra.value = game.extra || "";
  elements.collectionForm.elements.note.value = game.note || "";
  elements.collectionForm.elements.acquiredDate.value = game.acquiredDate || "";
  elements.collectionForm.elements.source.value = game.source || "";
  setCollectionFormMode();
  elements.collectionForm.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function removeCollectionItem(id, moveToWishlist = false) {
  const game = state.collection.find((item) => item.id === id);
  if (!game) return;

  state.collection = state.collection.filter((item) => item.id !== id);

  if (moveToWishlist) {
    const existingWish = state.wishlist.find(
      (item) =>
        normalize(item.platform) === normalize(game.platform) &&
        normalize(item.title) === normalize(game.title)
    );
    if (existingWish) {
      existingWish.replacement = false;
    } else {
      state.wishlist.push({
        id: newId("w"),
        platform: game.platform,
        title: game.title,
        note: game.note || "",
        inTransit: false,
        received: false,
        priority: "Medium",
        targetPrice: "",
        orderedDate: "",
        receivedDate: "",
        listingUrl: "",
        replacement: false,
      });
    }
  }

  if (state.editingCollectionId === id) {
    resetCollectionForm();
  }
  saveState();
  renderAll();
  flashStatus(moveToWishlist ? "Game moved back to the wishlist." : "Game removed from collection.");
}

function removeWishlistItem(id) {
  const before = state.wishlist.length;
  state.wishlist = state.wishlist.filter((item) => item.id !== id);
  if (state.wishlist.length !== before) {
    saveState();
    renderAll();
    flashStatus("Game removed from wishlist.");
  }
}

function openDeleteDialog(type, id) {
  const item = type === "collection"
    ? state.collection.find((game) => game.id === id)
    : state.wishlist.find((wish) => wish.id === id);
  if (!item) return;

  pendingDeletion = { type, id };
  const location = type === "collection" ? "collection" : "wishlist";
  const verb = type === "collection" ? "remove" : "delete";
  elements.deleteDialogTitle.textContent = type === "collection" ? "Remove game?" : "Delete wishlist game?";
  elements.deleteDialogMessage.textContent = `Are you sure you want to ${verb} “${item.title}” from your ${location}? This action cannot be undone.`;
  elements.deleteConfirmButton.textContent = type === "collection" ? "Remove" : "Delete";
  elements.deleteMoveCheckbox.checked = false;
  elements.deleteMoveOption.hidden = type !== "collection";
  elements.deleteDialog.showModal();
}

function setWishlistTransit(id, value) {
  const target = state.wishlist.find((item) => item.id === id);
  if (!target) return;
  target.inTransit = Boolean(value);
  saveState();
  renderAll();
}

function openReceiveDialog(id) {
  const wish = state.wishlist.find((item) => item.id === id);
  if (!wish) return;
  elements.receiveForm.reset();
  elements.receiveForm.elements.wishId.value = wish.id;
  elements.receiveForm.elements.note.value = wish.note || "";
  const today = new Date();
  today.setMinutes(today.getMinutes() - today.getTimezoneOffset());
  elements.receiveForm.elements.acquiredDate.value = today.toISOString().slice(0, 10);
  elements.receiveGameTitle.textContent = `${wish.title} · ${wish.platform}`;
  elements.receiveDialog.showModal();
}

function receiveWishlistItem(id, details) {
  const previous = cloneCurrentData();
  const index = state.wishlist.findIndex((item) => item.id === id);
  if (index === -1) return;
  const [wish] = state.wishlist.splice(index, 1);
  const existing = state.collection.find(
    (item) =>
      normalize(item.platform) === normalize(wish.platform) &&
      normalize(item.title) === normalize(wish.title)
  );
  if (!existing) {
    state.collection.push({
      id: newId("c"),
      platform: wish.platform,
      title: wish.title,
      version: details.version,
      cdCondition: details.cdCondition,
      manualCondition: details.manualCondition,
      price: details.price,
      extra: "",
      note: details.note || wish.note || "",
      acquiredDate: details.acquiredDate,
      source: details.source,
    });
  } else {
    existing.version = details.version || existing.version;
    existing.cdCondition = details.cdCondition || existing.cdCondition;
    existing.manualCondition = details.manualCondition || existing.manualCondition;
    existing.price = details.price || existing.price;
    existing.note = details.note || existing.note;
    existing.acquiredDate = details.acquiredDate || existing.acquiredDate;
    existing.source = details.source || existing.source;
  }
  saveState();
  renderAll();
  flashStatus(
    existing
      ? `${wish.title} replacement details updated in collection.`
      : `${wish.title} moved to collection.`,
    false,
    () => restoreData(previous, "Move undone.")
  );
}

function setWishlistFormMode() {
  const editing = Boolean(state.editingWishlistId);
  elements.wishlistFormTitle.textContent = editing ? "Edit Wishlist Game" : "Add To Wishlist";
  elements.wishlistSubmitButton.textContent = editing ? "Save Changes" : "Add Wish";
  elements.wishlistCancelEditButton.hidden = !editing;
}

function resetWishlistForm() {
  elements.wishlistForm.reset();
  elements.wishlistForm.elements.priority.value = "Medium";
  elements.wishlistForm.elements.editId.value = "";
  state.editingWishlistId = "";
  setWishlistFormMode();
}

function startWishlistEdit(id) {
  const wish = state.wishlist.find((item) => item.id === id);
  if (!wish) return;
  const form = elements.wishlistForm;
  state.editingWishlistId = wish.id;
  form.elements.editId.value = wish.id;
  form.elements.platform.value = wish.platform;
  form.elements.title.value = wish.title;
  form.elements.note.value = wish.note || "";
  form.elements.priority.value = wish.priority || "Medium";
  form.elements.targetPrice.value = wish.targetPrice || "";
  form.elements.orderedDate.value = wish.orderedDate || "";
  form.elements.listingUrl.value = wish.listingUrl || "";
  form.elements.inTransit.checked = Boolean(wish.inTransit);
  form.elements.replacement.checked = Boolean(wish.replacement);
  setWishlistFormMode();
  form.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function switchTab(name) {
  for (const tab of elements.tabs) {
    const active = tab.dataset.tab === name;
    tab.classList.toggle("is-active", active);
    tab.setAttribute("aria-selected", active ? "true" : "false");
    tab.tabIndex = active ? 0 : -1;
  }
  for (const [key, panel] of Object.entries(elements.panels)) {
    const active = key === name;
    panel.classList.toggle("is-active", active);
    panel.hidden = !active;
  }
}

function handleTabKeyboard(event, tabs, activate) {
  const current = tabs.indexOf(event.currentTarget);
  if (current === -1) return;
  let next = current;
  if (["ArrowRight", "ArrowDown"].includes(event.key)) next = (current + 1) % tabs.length;
  else if (["ArrowLeft", "ArrowUp"].includes(event.key)) next = (current - 1 + tabs.length) % tabs.length;
  else if (event.key === "Home") next = 0;
  else if (event.key === "End") next = tabs.length - 1;
  else return;
  event.preventDefault();
  tabs[next].focus();
  activate(tabs[next]);
}

function bindEvents() {
  elements.tabs.forEach((tab) => {
    tab.addEventListener("click", () => switchTab(tab.dataset.tab));
    tab.addEventListener("keydown", (event) => handleTabKeyboard(event, elements.tabs, (target) => switchTab(target.dataset.tab)));
  });

  elements.collectionSearch.addEventListener("input", renderCollection);
  elements.wishlistSearch.addEventListener("input", renderWishlist);
  elements.wishlistPlatformFilter.addEventListener("change", renderWishlist);
  elements.wishlistStatusFilter.addEventListener("change", renderWishlist);
  elements.statusUndoButton.addEventListener("click", () => {
    const action = undoAction;
    undoAction = null;
    if (action) action();
  });
  elements.exportStateButton.addEventListener("click", exportStateSnapshot);
  elements.importStateButton.addEventListener("click", () => {
    elements.importStateFile.value = "";
    elements.importStateFile.click();
  });
  elements.importStateFile.addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      await importStateSnapshot(file);
    } catch (error) {
      console.error(error);
      flashStatus(error instanceof Error ? error.message : "Invalid snapshot file.", true);
    }
  });
  elements.collectionPlatformTabs.addEventListener("click", (event) => {
    const button = event.target.closest("[data-platform-tab]");
    if (!button) return;
    state.activeCollectionPlatform = button.dataset.platformTab || "ALL";
    renderAll();
  });
  elements.collectionPlatformTabs.addEventListener("keydown", (event) => {
    const tabs = [...elements.collectionPlatformTabs.querySelectorAll('[role="tab"]')];
    if (!event.target.matches('[role="tab"]')) return;
    handleTabKeyboard(event, tabs, (target) => {
      state.activeCollectionPlatform = target.dataset.platformTab || "ALL";
      renderAll();
      const active = elements.collectionPlatformTabs.querySelector('[aria-selected="true"]');
      active?.focus();
    });
  });

  elements.collectionGroups.addEventListener("click", (event) => {
    const editButton = event.target.closest("[data-edit-collection]");
    if (editButton) {
      startCollectionEdit(editButton.dataset.editCollection);
      return;
    }
    const removeButton = event.target.closest("[data-remove-collection]");
    if (!removeButton) return;
    openDeleteDialog("collection", removeButton.dataset.removeCollection);
  });

  elements.wishlistTableBody.addEventListener("click", (event) => {
    const editButton = event.target.closest("[data-edit-wishlist]");
    if (editButton) {
      startWishlistEdit(editButton.dataset.editWishlist);
      return;
    }
    const receiveButton = event.target.closest("[data-receive-wishlist]");
    if (receiveButton) {
      openReceiveDialog(receiveButton.dataset.receiveWishlist);
      return;
    }
    const deleteButton = event.target.closest("[data-remove-wishlist]");
    if (deleteButton) {
      openDeleteDialog("wishlist", deleteButton.dataset.removeWishlist);
    }
  });

  elements.deleteCancelButton.addEventListener("click", () => {
    pendingDeletion = null;
    elements.deleteMoveCheckbox.checked = false;
    elements.deleteDialog.close();
  });
  elements.deleteDialog.addEventListener("cancel", () => {
    pendingDeletion = null;
  });
  elements.deleteConfirmButton.addEventListener("click", () => {
    const deletion = pendingDeletion;
    const moveToWishlist = deletion?.type === "collection" && elements.deleteMoveCheckbox.checked;
    pendingDeletion = null;
    elements.deleteDialog.close();
    if (!deletion) return;
    if (deletion.type === "collection") removeCollectionItem(deletion.id, moveToWishlist);
    else removeWishlistItem(deletion.id);
  });

  elements.wishlistTableBody.addEventListener("change", (event) => {
    const transit = event.target.closest("[data-in-transit]");
    if (transit) {
      setWishlistTransit(transit.dataset.inTransit, transit.checked);
      return;
    }

  });

  elements.receiveCancelButton.addEventListener("click", () => elements.receiveDialog.close());
  elements.receiveForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    const details = {
      version: form.elements.version.value.trim(),
      price: form.elements.price.value.trim(),
      cdCondition: form.elements.cdCondition.value.trim(),
      manualCondition: form.elements.manualCondition.value.trim(),
      acquiredDate: form.elements.acquiredDate.value,
      source: form.elements.source.value.trim(),
      note: form.elements.note.value.trim(),
    };
    const wishId = form.elements.wishId.value;
    elements.receiveDialog.close();
    receiveWishlistItem(wishId, details);
  });

  elements.collectionForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    const editId = state.editingCollectionId || form.elements.editId.value.trim();
    const payload = {
      id: editId || newId("c"),
      platform: form.elements.platform.value.trim(),
      title: form.elements.title.value.trim(),
      version: form.elements.version.value.trim(),
      cdCondition: form.elements.cdCondition.value.trim(),
      manualCondition: form.elements.manualCondition.value.trim(),
      price: form.elements.price.value.trim(),
      extra: form.elements.extra.value.trim(),
      note: form.elements.note.value.trim(),
      acquiredDate: form.elements.acquiredDate.value,
      source: form.elements.source.value.trim(),
    };

    if (!payload.title || !payload.platform) {
      flashStatus("Platform and title are required.", true);
      return;
    }
    const duplicate = state.collection.some(
      (item) =>
        normalize(item.platform) === normalize(payload.platform) &&
        normalize(item.title) === normalize(payload.title) &&
        item.id !== editId
    );
    if (duplicate) {
      flashStatus("This game is already in your collection.", true);
      return;
    }

    if (editId) {
      const target = state.collection.find((item) => item.id === editId);
      if (!target) {
        flashStatus("Could not find the selected game to edit.", true);
        return;
      }
      target.platform = payload.platform;
      target.title = payload.title;
      target.version = payload.version;
      target.cdCondition = payload.cdCondition;
      target.manualCondition = payload.manualCondition;
      target.price = payload.price;
      target.extra = payload.extra;
      target.note = payload.note;
      target.acquiredDate = payload.acquiredDate;
      target.source = payload.source;
      saveState();
      renderAll();
      resetCollectionForm();
      flashStatus(`${payload.title} updated.`);
      return;
    }

    const matchingWish = state.wishlist.find(
      (item) => normalize(item.platform) === normalize(payload.platform) && normalize(item.title) === normalize(payload.title)
    );
    if (matchingWish) {
      flashStatus("This game is on your wishlist. Use its Receive action to preserve purchase details.", true);
      return;
    }

    state.collection.push(payload);
    saveState();
    renderAll();
    resetCollectionForm();
    flashStatus(`${payload.title} added to collection.`);
  });

  elements.collectionCancelEditButton.addEventListener("click", () => {
    resetCollectionForm();
    flashStatus("Edit cancelled.");
  });

  elements.wishlistForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    const editId = state.editingWishlistId || form.elements.editId.value.trim();
    const payload = {
      id: editId || newId("w"),
      platform: form.elements.platform.value.trim(),
      title: form.elements.title.value.trim(),
      note: form.elements.note.value.trim(),
      inTransit: form.elements.inTransit.checked,
      received: false,
      priority: form.elements.priority.value,
      targetPrice: form.elements.targetPrice.value.trim(),
      orderedDate: form.elements.orderedDate.value,
      receivedDate: "",
      listingUrl: form.elements.listingUrl.value.trim(),
      replacement: form.elements.replacement.checked,
    };

    if (!payload.title || !payload.platform) {
      flashStatus("Platform and title are required.", true);
      return;
    }
    const duplicate = state.wishlist.some(
      (item) =>
        normalize(item.platform) === normalize(payload.platform) &&
        normalize(item.title) === normalize(payload.title) &&
        item.id !== editId
    );
    if (duplicate) {
      flashStatus("This game is already in your wishlist.", true);
      return;
    }
    const alreadyOwned = isDuplicateCollection(payload.platform, payload.title);
    if (alreadyOwned && !payload.replacement) {
      flashStatus("This game is already owned. Mark it as a replacement or upgrade copy to add it.", true);
      return;
    }

    if (editId) {
      const target = state.wishlist.find((item) => item.id === editId);
      if (!target) {
        flashStatus("Could not find the selected wishlist game.", true);
        return;
      }
      Object.assign(target, payload);
      saveState();
      renderAll();
      resetWishlistForm();
      flashStatus(`${payload.title} updated.`);
      return;
    }

    state.wishlist.push(payload);
    saveState();
    resetWishlistForm();
    renderAll();
    flashStatus(`${payload.title} added to wishlist.`);
  });

  elements.wishlistCancelEditButton.addEventListener("click", () => {
    resetWishlistForm();
    flashStatus("Edit cancelled.");
  });
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

async function loadInitialState() {
  const cacheCandidates = [
    { key: STORAGE_KEY, label: "current browser state" },
    { key: BACKUP_STORAGE_KEY, label: "local backup" },
    { key: LEGACY_STORAGE_KEY, label: "legacy browser state" },
  ];

  for (const candidate of cacheCandidates) {
    const cached = localStorage.getItem(candidate.key);
    if (!cached) continue;
    try {
      const parsed = JSON.parse(cached);
      const normalizedState = normalizeStatePayload(parsed, { strict: true });
      state.collection = normalizedState.collection;
      state.wishlist = normalizedState.wishlist;
      saveState();
      if (candidate.key !== STORAGE_KEY) {
        setTimeout(() => flashStatus(`Recovered ${candidate.label}.`), 0);
      }
      return;
    } catch (error) {
      console.warn(`Could not load ${candidate.label}`, error);
    }
  }

  try {
    const response = await fetch("data/seed.json");
    if (!response.ok) throw new Error(`Failed to load seed (${response.status})`);
    const seed = await response.json();
    const normalizedState = normalizeStatePayload(seed, { strict: true });
    state.collection = normalizedState.collection;
    state.wishlist = normalizedState.wishlist;
    saveState();
  } catch (error) {
    console.error(error);
    flashStatus("Could not load seed data. Run a local web server and refresh.", true);
  }
}

await loadInitialState();
bindEvents();
renderAll();
