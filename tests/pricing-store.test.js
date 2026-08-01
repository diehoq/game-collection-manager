import assert from "node:assert/strict";
import test from "node:test";

import {
  buildPriceChartingSearchUrl,
  buildPricingKey,
  centsToInput,
  decimalToCents,
  hasPricing,
  normalizePriceChartingUrl,
  normalizePricingPayload,
  preferredPrice,
} from "../pricing-store.js";


test("pricing keys are stable across casing and whitespace", () => {
  assert.equal(buildPricingKey(" PS2 ", "Shadow Hearts"), "ps2::shadow hearts");
});

test("manual decimal prices convert safely to cents", () => {
  assert.equal(decimalToCents("19.99"), 1999);
  assert.equal(decimalToCents("19,50"), 1950);
  assert.equal(decimalToCents(""), null);
  assert.equal(decimalToCents("-1"), null);
  assert.equal(decimalToCents("12.345"), null);
  assert.equal(centsToInput(1999), "19.99");
});

test("only HTTPS PriceCharting product links are accepted", () => {
  assert.equal(
    normalizePriceChartingUrl("https://www.pricecharting.com/game/pal-playstation-2/example"),
    "https://www.pricecharting.com/game/pal-playstation-2/example"
  );
  assert.equal(normalizePriceChartingUrl("javascript:alert(1)"), "");
  assert.equal(normalizePriceChartingUrl("https://example.com/game"), "");
});

test("search links include the game and mapped platform", () => {
  const url = new URL(buildPriceChartingSearchUrl("PS4", "Gravity Rush"));
  assert.equal(url.hostname, "www.pricecharting.com");
  assert.equal(url.searchParams.get("type"), "prices");
  assert.equal(url.searchParams.get("q"), "Gravity Rush Playstation 4");
});

test("pricing backups normalize records and reject malformed roots", () => {
  const records = normalizePricingPayload({
    schemaVersion: 1,
    records: {
      "ps2::game": { currency: "usd", cibCents: 2500, looseCents: -1 },
    },
  }, { strict: true });

  assert.equal(records["ps2::game"].currency, "USD");
  assert.equal(records["ps2::game"].cibCents, 2500);
  assert.equal(records["ps2::game"].looseCents, null);
  assert.equal(hasPricing(records["ps2::game"]), true);
  assert.deepEqual(preferredPrice(records["ps2::game"]), { field: "cibCents", cents: 2500 });
  assert.throws(() => normalizePricingPayload({ records: [] }, { strict: true }));
});
