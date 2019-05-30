import { WorkbookWriter, SheetWriter, col2alpha } from "../src/writer";

import "mocha";
import * as assert from "power-assert";
import * as fs from "fs";

describe("WorkbookWriter", () => {
  it("write", async () => {
    const file = fs.createWriteStream("./test.xlsx");
    const workbook = new WorkbookWriter(file);
    workbook.addStyle("title", {
      fill: { color: "orange" },
      border: {
        top: { style: "medium" },
        bottom: { style: "medium" }
      }
    });
    workbook.addStyle("title-first", {
      border: {
        left: { style: "medium" }
      }
    });
    workbook.addStyle("title-last", {
      border: {
        right: { style: "medium" }
      }
    });
    const sheet1 = workbook.sheet("シート1");
    sheet1.addStyleRule("title", (r, c) => r === 0);
    sheet1.addStyleRule("title-first", (r, c) => r === 0 && c === 0);
    sheet1.addStyleRule("title-last", (r, c) => r === 0 && c === 3);

    sheet1.writeRow(["#", "日付", "名称", "備考"]);
    sheet1.writeRow([
      1,
      new Date("2019-04-01"),
      "テスト1",
      { type: "string", formula: "=B1" }
    ]);
    sheet1.writeRow([
      2,
      new Date("2019-04-02"),
      "テスト2",
      { type: "string", formula: "=B1+C1" }
    ]);
    sheet1.writeRow([
      3,
      new Date("2019-04-03"),
      "テスト3",
      { type: "string", formula: "=B1+C1+D1" }
    ]);
    sheet1.end();

    await workbook.end();
  });
});

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
