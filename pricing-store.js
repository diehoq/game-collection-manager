export const PRICING_SCHEMA_VERSION = 1;
export const PRICING_STORAGE_KEY = "gameCollectionManager.priceCharting.v1";
export const SUPPORTED_CURRENCIES = Object.freeze(["EUR", "USD", "GBP"]);

const PRICE_FIELDS = Object.freeze([
  "looseCents",
  "cibCents",
  "newCents",
  "boxOnlyCents",
  "manualOnlyCents",
]);

function normalizedText(value) {
  return String(value ?? "").trim();
}

export function buildPricingKey(platform, title) {
  const normalizedPlatform = normalizedText(platform).toLowerCase();
  const normalizedTitle = normalizedText(title).toLowerCase();
  return `${normalizedPlatform}::${normalizedTitle}`;
}

export function decimalToCents(value) {
  if (value === null || value === undefined || normalizedText(value) === "") return null;
  const normalized = normalizedText(value).replace(",", ".");
  if (!/^\d+(?:\.\d{1,2})?$/.test(normalized)) return null;
  const cents = Math.round(Number(normalized) * 100);
  return Number.isSafeInteger(cents) && cents >= 0 ? cents : null;
}

export function centsToInput(cents) {
  if (!Number.isSafeInteger(cents) || cents < 0) return "";
  return (cents / 100).toFixed(2);
}

export function normalizeCurrency(value) {
  const currency = normalizedText(value).toUpperCase();
  return SUPPORTED_CURRENCIES.includes(currency) ? currency : "EUR";
}

export function normalizePriceChartingUrl(value) {
  const raw = normalizedText(value);
  if (!raw) return "";
  try {
    const url = new URL(raw);
    const hostname = url.hostname.toLowerCase();
    if (url.protocol !== "https:" || !["pricecharting.com", "www.pricecharting.com"].includes(hostname)) {
      return "";
    }
    return url.href;
  } catch {
    return "";
  }
}

export function buildPriceChartingSearchUrl(platform, title) {
  const platformNames = {
    ps1: "Playstation",
    ps2: "Playstation 2",
    ps4: "Playstation 4",
    "ds wii": "Nintendo DS Wii",
  };
  const platformKey = normalizedText(platform).toLowerCase();
  const query = `${normalizedText(title)} ${platformNames[platformKey] || normalizedText(platform)}`.trim();
  const url = new URL("https://www.pricecharting.com/search-products");
  url.searchParams.set("type", "prices");
  url.searchParams.set("q", query);
  return url.href;
}

export function normalizePricingRecord(record) {
  const normalized = {
    currency: normalizeCurrency(record?.currency),
    priceChartingUrl: normalizePriceChartingUrl(record?.priceChartingUrl),
    updatedAt: normalizedText(record?.updatedAt),
  };
  for (const field of PRICE_FIELDS) {
    const value = record?.[field];
    normalized[field] = Number.isSafeInteger(value) && value >= 0 ? value : null;
  }
  return normalized;
}

export function normalizePricingPayload(payload, { strict = false } = {}) {
  if (!payload || typeof payload !== "object") {
    if (strict) throw new Error("Invalid pricing backup format.");
    return {};
  }
  if (Number(payload.schemaVersion || 1) > PRICING_SCHEMA_VERSION) {
    throw new Error(`Pricing schema ${payload.schemaVersion} is newer than this app supports.`);
  }
  const rawRecords = payload.records;
  if (!rawRecords || typeof rawRecords !== "object" || Array.isArray(rawRecords)) {
    if (strict) throw new Error("Pricing backup must contain a records object.");
    return {};
  }

  const records = {};
  for (const [key, record] of Object.entries(rawRecords)) {
    if (!key.includes("::") || typeof record !== "object" || !record) {
      if (strict) throw new Error("Pricing backup contains an invalid record.");
      continue;
    }
    records[key] = normalizePricingRecord(record);
  }
  return records;
}

export function hasPricing(record) {
  return Boolean(record) && PRICE_FIELDS.some((field) => Number.isSafeInteger(record[field]));
}

export function preferredPrice(record) {
  if (!record) return null;
  for (const field of ["cibCents", "looseCents", "newCents", "boxOnlyCents", "manualOnlyCents"]) {
    if (Number.isSafeInteger(record[field])) return { field, cents: record[field] };
  }
  return null;
}

export function pricingPayload(records) {
  return {
    schemaVersion: PRICING_SCHEMA_VERSION,
    exportedAt: new Date().toISOString(),
    source: "Manual values referenced from PriceCharting",
    records,
  };
}
