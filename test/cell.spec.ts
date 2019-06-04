import "mocha";
import * as assert from "power-assert";

import { col2alpha } from "../src/cell";

describe("col2alpha", () => {
  it("returns 'A' when 0 is given", () => {
    assert.strictEqual(col2alpha(0), "A");
  });
  it("returns 'Z' when 25 is given", () => {
    assert.strictEqual(col2alpha(25), "Z");
  });
  it("returns 'AA' when 26 is given", () => {
    assert.strictEqual(col2alpha(26), "AA");
  });
  it("returns 'ALL' when 999 is given", () => {
    assert.strictEqual(col2alpha(999), "ALL");
  });
});
