const STORAGE_KEY = "gameCollectionManager.v1";
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
  collectionGroups: document.getElementById("collection-groups"),
  wishlistTableBody: document.querySelector("#wishlist-table tbody"),
  collectionForm: document.getElementById("collection-form"),
  wishlistForm: document.getElementById("wishlist-form"),
};

let statusTimer = null;
let currentExportBlobUrl = null;

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

function saveState() {
  localStorage.setItem(
    STORAGE_KEY,
    JSON.stringify({
      collection: state.collection,
      wishlist: state.wishlist,
    })
  );
}

function toBool(value) {
  if (typeof value === "boolean") return value;
  const text = normalize(value);
  return text === "true" || text === "1" || text === "yes" || text === "x";
}

function normalizeCollectionRow(row) {
  return {
    platform: String(row?.platform ?? "").trim(),
    title: String(row?.title ?? "").trim(),
    version: String(row?.version ?? "").trim(),
    cdCondition: String(row?.cdCondition ?? row?.cd_condition ?? "").trim(),
    manualCondition: String(row?.manualCondition ?? row?.manual_condition ?? "").trim(),
    price: String(row?.price ?? "").trim(),
    extra: String(row?.extra ?? "").trim(),
    note: String(row?.note ?? "").trim(),
  };
}

function normalizeWishlistRow(row) {
  return {
    platform: String(row?.platform ?? "").trim(),
    title: String(row?.title ?? "").trim(),
    note: String(row?.note ?? "").trim(),
    inTransit: toBool(row?.inTransit ?? row?.in_transit),
    received: toBool(row?.received),
  };
}

function buildCollectionKey(platform, title) {
  return `${normalize(platform)}::${normalize(title)}`;
}

function normalizeStatePayload(payload) {
  if (!payload || typeof payload !== "object") {
    throw new Error("Invalid state format.");
  }

  const rawCollection = Array.isArray(payload.collection) ? payload.collection : [];
  const rawWishlist = Array.isArray(payload.wishlist) ? payload.wishlist : [];
  const collection = [];
  const wishlist = [];
  const seenCollection = new Set();
  const seenWishlist = new Set();

  for (const row of rawCollection) {
    const normalizedRow = normalizeCollectionRow(row);
    if (!normalizedRow.platform || !normalizedRow.title) continue;
    const key = buildCollectionKey(normalizedRow.platform, normalizedRow.title);
    if (seenCollection.has(key)) continue;
    seenCollection.add(key);
    collection.push(normalizedRow);
  }

  for (const row of rawWishlist) {
    const normalizedRow = normalizeWishlistRow(row);
    if (!normalizedRow.platform || !normalizedRow.title) continue;
    const key = buildCollectionKey(normalizedRow.platform, normalizedRow.title);
    if (normalizedRow.received) {
      if (!seenCollection.has(key)) {
        seenCollection.add(key);
        collection.push({
          platform: normalizedRow.platform,
          title: normalizedRow.title,
          version: "",
          cdCondition: "",
          manualCondition: "",
          price: "",
          extra: "",
          note: normalizedRow.note || "",
        });
      }
      continue;
    }
    if (seenWishlist.has(key)) continue;
    seenWishlist.add(key);
    wishlist.push({
      platform: normalizedRow.platform,
      title: normalizedRow.title,
      note: normalizedRow.note,
      inTransit: normalizedRow.inTransit,
      received: false,
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

  return {
    collection: collection.map((item, index) => ({ id: `c${index + 1}`, ...item })),
    wishlist: wishlist.map((item, index) => ({ id: `w${index + 1}`, ...item })),
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
      schemaVersion: 1,
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
  const normalizedState = normalizeStatePayload(parsed);
  state.collection = normalizedState.collection;
  state.wishlist = normalizedState.wishlist;
  saveState();
  renderAll();
  flashStatus(`Imported ${state.collection.length} collection and ${state.wishlist.length} wishlist games.`);
}

function flashStatus(message, isError = false) {
  clearTimeout(statusTimer);
  elements.status.textContent = message;
  elements.status.style.color = isError ? "#9f3718" : "#2a7f62";
  statusTimer = setTimeout(() => {
    elements.status.textContent = "";
  }, 3000);
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
    table.innerHTML = `
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
        <td>${escapeHtml(game.title)}</td>
        <td>${escapeHtml(game.version || "")}</td>
        <td class="mono">${escapeHtml(game.cdCondition || "")}</td>
        <td class="mono">${escapeHtml(game.manualCondition || "")}</td>
        <td>${escapeHtml(game.price || "")}</td>
        <td>${escapeHtml(game.extra || "")}</td>
        <td>${escapeHtml(game.note || "")}</td>
        <td>
          <div class="row-actions">
            <button class="danger" data-remove-collection="${game.id}" type="button">Remove</button>
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

  return state.wishlist.filter((item) => {
    if (platform && normalize(item.platform) !== platform) return false;
    if (!query) return true;
    const haystack = [item.title, item.note, item.platform].map(normalize).join(" ");
    return haystack.includes(query);
  });
}

function renderWishlist() {
  const rows = getWishlistFiltered().sort((a, b) => {
    if (a.platform !== b.platform) return a.platform.localeCompare(b.platform);
    return a.title.localeCompare(b.title);
  });

  elements.wishlistTableBody.innerHTML = "";
  for (const item of rows) {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${renderPlatformBadge(item.platform, "sm")}</td>
      <td>${escapeHtml(item.title)}</td>
      <td>${escapeHtml(item.note || "")}</td>
      <td><input class="checkbox" data-in-transit="${item.id}" type="checkbox" ${item.inTransit ? "checked" : ""} /></td>
      <td><input class="checkbox" data-received="${item.id}" type="checkbox" /></td>
      <td>
        <div class="row-actions">
          <button class="danger" data-remove-wishlist="${item.id}" type="button">Delete</button>
        </div>
      </td>
    `;
    elements.wishlistTableBody.append(row);
  }

  if (!rows.length) {
    const row = document.createElement("tr");
    row.innerHTML = '<td colspan="6" class="empty">Wishlist is empty for this filter.</td>';
    elements.wishlistTableBody.append(row);
  }
}

function renderAll() {
  fillPlatformSelects();
  renderCollectionPlatformTabs();
  renderHeroStats();
  renderCollection();
  renderWishlist();
}

function removeCollectionItem(id) {
  const before = state.collection.length;
  state.collection = state.collection.filter((item) => item.id !== id);
  if (state.collection.length !== before) {
    saveState();
    renderAll();
    flashStatus("Game removed from collection.");
  }
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

function setWishlistTransit(id, value) {
  const target = state.wishlist.find((item) => item.id === id);
  if (!target) return;
  target.inTransit = Boolean(value);
  saveState();
  renderAll();
}

function receiveWishlistItem(id) {
  const index = state.wishlist.findIndex((item) => item.id === id);
  if (index === -1) return;
  const [wish] = state.wishlist.splice(index, 1);
  const duplicate = isDuplicateCollection(wish.platform, wish.title);
  if (!duplicate) {
    state.collection.push({
      id: newId("c"),
      platform: wish.platform,
      title: wish.title,
      version: "",
      cdCondition: "",
      manualCondition: "",
      price: "",
      extra: "",
      note: wish.note || "",
    });
  }
  saveState();
  renderAll();
  flashStatus(`${wish.title} moved to collection${duplicate ? " (already existed)." : "."}`);
}

function switchTab(name) {
  for (const tab of elements.tabs) {
    tab.classList.toggle("is-active", tab.dataset.tab === name);
  }
  for (const [key, panel] of Object.entries(elements.panels)) {
    panel.classList.toggle("is-active", key === name);
  }
}

function bindEvents() {
  elements.tabs.forEach((tab) => {
    tab.addEventListener("click", () => switchTab(tab.dataset.tab));
  });

  elements.collectionSearch.addEventListener("input", renderCollection);
  elements.wishlistSearch.addEventListener("input", renderWishlist);
  elements.wishlistPlatformFilter.addEventListener("change", renderWishlist);
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
      flashStatus("Invalid snapshot file.", true);
    }
  });
  elements.collectionPlatformTabs.addEventListener("click", (event) => {
    const button = event.target.closest("[data-platform-tab]");
    if (!button) return;
    state.activeCollectionPlatform = button.dataset.platformTab || "ALL";
    renderAll();
  });

  elements.collectionGroups.addEventListener("click", (event) => {
    const button = event.target.closest("[data-remove-collection]");
    if (!button) return;
    removeCollectionItem(button.dataset.removeCollection);
  });

  elements.wishlistTableBody.addEventListener("click", (event) => {
    const deleteButton = event.target.closest("[data-remove-wishlist]");
    if (deleteButton) {
      removeWishlistItem(deleteButton.dataset.removeWishlist);
    }
  });

  elements.wishlistTableBody.addEventListener("change", (event) => {
    const transit = event.target.closest("[data-in-transit]");
    if (transit) {
      setWishlistTransit(transit.dataset.inTransit, transit.checked);
      return;
    }

    const received = event.target.closest("[data-received]");
    if (received && received.checked) {
      receiveWishlistItem(received.dataset.received);
    }
  });

  elements.collectionForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    const payload = {
      id: newId("c"),
      platform: form.elements.platform.value.trim(),
      title: form.elements.title.value.trim(),
      version: form.elements.version.value.trim(),
      cdCondition: form.elements.cdCondition.value.trim(),
      manualCondition: form.elements.manualCondition.value.trim(),
      price: form.elements.price.value.trim(),
      extra: form.elements.extra.value.trim(),
      note: form.elements.note.value.trim(),
    };

    if (!payload.title || !payload.platform) {
      flashStatus("Platform and title are required.", true);
      return;
    }
    if (isDuplicateCollection(payload.platform, payload.title)) {
      flashStatus("This game is already in your collection.", true);
      return;
    }
    state.collection.push(payload);
    saveState();
    form.reset();
    renderAll();
    flashStatus(`${payload.title} added to collection.`);
  });

  elements.wishlistForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    const payload = {
      id: newId("w"),
      platform: form.elements.platform.value.trim(),
      title: form.elements.title.value.trim(),
      note: form.elements.note.value.trim(),
      inTransit: form.elements.inTransit.checked,
      received: false,
    };

    if (!payload.title || !payload.platform) {
      flashStatus("Platform and title are required.", true);
      return;
    }
    const duplicate = state.wishlist.some(
      (item) => normalize(item.platform) === normalize(payload.platform) && normalize(item.title) === normalize(payload.title)
    );
    if (duplicate) {
      flashStatus("This game is already in your wishlist.", true);
      return;
    }
    state.wishlist.push(payload);
    saveState();
    form.reset();
    renderAll();
    flashStatus(`${payload.title} added to wishlist.`);
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
  try {
    const cached = localStorage.getItem(STORAGE_KEY);
    if (cached) {
      const parsed = JSON.parse(cached);
      const normalizedState = normalizeStatePayload(parsed);
      state.collection = normalizedState.collection;
      state.wishlist = normalizedState.wishlist;
      saveState();
      return;
    }

    const response = await fetch("data/seed.json");
    if (!response.ok) throw new Error(`Failed to load seed (${response.status})`);
    const seed = await response.json();
    const normalizedState = normalizeStatePayload(seed);
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
