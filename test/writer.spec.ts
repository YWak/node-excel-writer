import { WorkbookWriter, SheetWriter } from "../src/writer";

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
    workbook.addStyle("date", {
      format: 31,
    });
    const sheet1 = workbook.sheet("シート1");
    sheet1.addStyleRule("title", (r, c) => r === 0);
    sheet1.addStyleRule("title-first", (r, c) => r === 0 && c === 0);
    sheet1.addStyleRule("title-last", (r, c) => r === 0 && c === 3);

    sheet1.addCellValue("#");
    sheet1.addCellValue("日付");
    sheet1.addCellValue("名称");
    sheet1.addCellValue("備考");
    sheet1.nextRow();

    sheet1.addCellValue(1);
    sheet1.addCell({ type: "date", value: new Date("2019-04-01"), styles: ["date"] });
    sheet1.addCellValue("テスト1");
    sheet1.addCell({ type: "string", formula: "=A2" });
    sheet1.nextRow();

    sheet1.addCellValue(2);
    sheet1.addCellValue(new Date("2019-04-02"));
    sheet1.addCellValue("テスト2");
    sheet1.addCell({ type: "string", formula: "=A2+A3" });
    sheet1.nextRow();

    sheet1.addCellValue(3);
    sheet1.addCellValue(new Date("2019-04-03"));
    sheet1.addCellValue("テスト3");
    sheet1.addCell({ type: "string", formula: "=A2+A3+A4" });
    sheet1.nextRow();

    sheet1.end();

    await workbook.end();
  });
});
