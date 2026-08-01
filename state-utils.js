const PRIORITY_BY_KEY = Object.freeze({
  low: "Low",
  medium: "Medium",
  high: "High",
});

export function normalizePriority(value) {
  const key = String(value ?? "").trim().toLowerCase();
  return PRIORITY_BY_KEY[key] || "Medium";
}

export function normalizeRecordId(value, prefix) {
  if (prefix !== "c" && prefix !== "w") return "";
  const id = String(value ?? "").trim().toLowerCase();
  const pattern = prefix === "c" ? /^c\d{1,12}$/ : /^w\d{1,12}$/;
  return pattern.test(id) && Number(id.slice(1)) > 0 ? id : "";
}

export function persistOrRestore(previous, next, persist) {
  try {
    if (persist(next)) return { data: next, persisted: true };
  } catch {
    // The caller owns user-facing error reporting; state still rolls back here.
  }
  return { data: previous, persisted: false };
}

function withoutIds(rows) {
  return rows.map(({ id: _id, ...row }) => row);
}

export function statesHaveSameContent(left, right) {
  if (!left || !right) return false;
  return JSON.stringify({
    collection: withoutIds(left.collection || []),
    wishlist: withoutIds(left.wishlist || []),
  }) === JSON.stringify({
    collection: withoutIds(right.collection || []),
    wishlist: withoutIds(right.wishlist || []),
  });
}

export function getRevisionStatus(browserRevision, repositoryRevision, sameContent) {
  if (sameContent) return "same-content";
  if (browserRevision && browserRevision === repositoryRevision) return "repository-unchanged";
  return "conflict";
}
