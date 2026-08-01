import assert from "node:assert/strict";
import test from "node:test";

import { normalizePriority, normalizeRecordId } from "../state-utils.js";


test("record IDs accept only the expected prefix and numeric suffix", () => {
  assert.equal(normalizeRecordId("c17", "c"), "c17");
  assert.equal(normalizeRecordId(" W0042 ", "w"), "w0042");
  assert.equal(normalizeRecordId("w17", "c"), "");
  assert.equal(normalizeRecordId('c1\" onmouseover=\"alert(1)', "c"), "");
  assert.equal(normalizeRecordId("c0", "c"), "");
  assert.equal(normalizeRecordId("c1234567890123", "c"), "");
});

test("priority is constrained to the supported enum", () => {
  assert.equal(normalizePriority("high"), "High");
  assert.equal(normalizePriority(" LOW "), "Low");
  assert.equal(normalizePriority("urgent"), "Medium");
  assert.equal(normalizePriority('\" onmouseover=\"alert(1)'), "Medium");
});
