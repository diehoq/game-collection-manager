import assert from "node:assert/strict";
import test from "node:test";

import {
  getRevisionStatus,
  normalizePriority,
  normalizeRecordId,
  persistOrRestore,
  statesHaveSameContent,
} from "../state-utils.js";


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

test("failed persistence selects the previous state", () => {
  const previous = { collection: [{ id: "c1" }], wishlist: [] };
  const next = { collection: [], wishlist: [] };

  assert.deepEqual(persistOrRestore(previous, next, () => false), {
    data: previous,
    persisted: false,
  });
  assert.deepEqual(persistOrRestore(previous, next, () => true), {
    data: next,
    persisted: true,
  });
  assert.deepEqual(persistOrRestore(previous, next, () => { throw new Error("quota"); }), {
    data: previous,
    persisted: false,
  });
});

test("state comparison ignores generated IDs but detects content changes", () => {
  const left = {
    collection: [{ id: "c1", platform: "PS2", title: "Game" }],
    wishlist: [{ id: "w1", platform: "PS4", title: "Wish", priority: "High" }],
  };
  const same = {
    collection: [{ id: "c99", platform: "PS2", title: "Game" }],
    wishlist: [{ id: "w42", platform: "PS4", title: "Wish", priority: "High" }],
  };
  const changed = {
    collection: [{ id: "c1", platform: "PS2", title: "Different" }],
    wishlist: left.wishlist,
  };

  assert.equal(statesHaveSameContent(left, same), true);
  assert.equal(statesHaveSameContent(left, changed), false);
});

test("revision status detects true two-sided divergence", () => {
  assert.equal(getRevisionStatus("old", "new", true), "same-content");
  assert.equal(getRevisionStatus("same", "same", false), "repository-unchanged");
  assert.equal(getRevisionStatus("old", "new", false), "conflict");
  assert.equal(getRevisionStatus("", "new", false), "conflict");
});
