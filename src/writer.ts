import { Writable } from "stream";
import archiver from "archiver";

import { Style, Predicate } from "./style";

/**
 * xlsxファイルを作成するためのクラスです。
 */
export class WorkbookWriter {
  private sheets: string[] = [];
  public archive: archiver.Archiver;
  public currentSheet?: SheetWriter;

  /**
   *
   * @param stream xlsxファイルの書き込み先
   */
  public constructor(stream: Writable) {
    this.archive = archiver("zip");
    this.archive.pipe(stream);

    this.archive.append(xml.app(), { name: "docProps/app.xml" });
    this.archive.append(xml.core(), { name: "docProps/core.xml" });
    this.archive.append(xml.sharedStrings(), { name: "xl/sharedStrings.xml" });

    this.archive.append(xml.rels(), { name: "_rels/.rels" });
  }

  /**
   * データ登録が完了したことを通知します。
   */
  public async end(): Promise<void> {

    this.archive.append(xml.contentTypes(this.sheets), { name: "[Content_Types].xml" });
    this.archive.append(xml.workbookRel(this.sheets), { name: "xl/_rels/workbook.xml.rels" });
    this.archive.append(xml.workbook(this.sheets), { name: "xl/workbook.xml" });
    return this.archive.finalize();
  }

  /**
   * シートを作成します。
   *
   * @param sheetName シート名
   */
  public sheet(sheetName: string): SheetWriter {
    if (this.currentSheet != null) {
      throw new Error(`${this.currentSheet.sheetName}が編集中です。`);
    }

    this.sheets.push(sheetName);
    this.currentSheet = new SheetWriter(this, sheetName, this.sheets.length);
    return this.currentSheet;
  }

  /**
   * スタイルを登録します。
   *
   * @param name スタイル名
   * @param style スタイル
   */
  public addStyle(name: string, style: Style): void {}
}

/**
 * シートを作成するためのインターフェースを提供するクラスです。
 */
export class SheetWriter {
  private book: WorkbookWriter;

  private rows: Array<Array<Cell>> = [];

  public sheetName: string;

  private index: number;

  public constructor(book: WorkbookWriter, sheetName: string, index: number) {
    this.book = book;
    this.sheetName = sheetName;
    this.index = index;
  }

  /**
   * このシートでスタイルを適用する条件を定義します。
   * predicateがtrueを返すときにnameで定義されたスタイルが適用されます。
   *
   * @param name スタイル名
   * @param predicate スタイルの適用条件
   */
  public addStyleRule(name: string, predicate: Predicate): void {}

  public skipRows(rows: number) {
    for (let i = 0; i < rows; i++) {
      this.rows.push([]);
    }
  }

  /**
   * 指定した値をもつ行を作成します。
   *
   * @param values セルの値
   */
  public writeRow(values: Row): void;

  /**
   * 指定した値を持つ行を作成します。
   * 先頭offset列まで空欄にしてからvalueを適用します。
   *
   * @param offset valuesを書き出すまでに追加する列の数
   * @param values セルの値
   */
  public writeRow(offset: number, values: Row): void;

  public writeRow(offset: unknown, values?: unknown): void {
    if (arguments.length === 1) {
      values = offset as Row;
      offset = 0;
    }
    const _offset = offset as number;
    const _values = values as Row;

    const row: Cell[] = [];

    for (let i = 0; i < _offset; i++) {
      row.push({ value: null });
    }
    _values.forEach(v => {
      let c: Cell;
      if (v == null) {
        c = { value: null };
      } else if (typeof v === "number" || typeof v === "string") {
        c = { value: v };
      } else if (v instanceof Date) {
        c = { value: v.toISOString() };
      } else {
        c = { value: null };
      }
      row.push(c);
    });
    this.rows.push(row);
  }

  /*
   * 指定したセルに、指定したスタイルを適用した場合の結果を返します。
   * 実際にはセルは書き込みません。
   * デバッグ用です。
   *
   * @param pos セルの位置(A1形式)
   * @param cell セルの値
   */
  // public describe(pos: string, cell: CellDef): Description {
  //  return {};
  // }

  /**
   * シートへのデータ登録が完了したことを通知します。
   */
  public end(): void {
    this.book.archive.append(xml.sheet(this.rows), { name: `xl/worksheets/sheet${this.index}.xml` });
    this.book.currentSheet = undefined;
  }
}

export type Row = Array<StyledCell | unknown>;

/** セルの値が表す型 */
export type CellType = "blank" | "boolean" | "formula" | "numeric" | "string";

export interface Cell {
  value: unknown;
}

/**
 * セルの定義。valueとformulaはどちらか一方のみ。
 */
export interface StyledCell {
  /** セルに設定する値 */
  value?: unknown;

  /** セルに設定する数式 */
  formula?: string;

  /** セルが表す値の型 */
  type?: CellType;

  /** セルに適用されるスタイルの名称 */
  styles?: string[];
}

const PACKAGE = require("../package.json");

export const col2alpha = ((cache: string[]) => (col: number): string => {
  if (cache[col] == null) {
    if (col < 26) {
      cache[col] = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[col];
    } else {
      cache[col] = `${col2alpha(Math.floor((col + 1) / 26) - 1)}${col2alpha(col % 26)}`;
    }
  }
  return cache[col];
})([]);

const ref = (r: number, c: number) => `${col2alpha(c)}${r + 1}`;

const xml = {
  escape(s: string) {
    if (s == null) {
      return "";
    }

    return s.replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
  },
  stringify <T>(array: Array<T>, mapper: (v: T, n:number) => string) {
    return array.map(mapper).map(s => s.trim()).join();
  },
  trim(s: string){
    return s.trim().replace(/\n */g, " ");
  },
  rels() {
    return this.trim(`
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            Target="xl/workbook.xml"/>
        <Relationship Id="rId2"
            Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
            Target="docProps/core.xml"/>
        <Relationship Id="rId3"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
            Target="docProps/app.xml"/>
      </Relationships>
    `);
  },
  workbookRel(names: string[]) {
    return this.trim(`
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      ${this.stringify(names, (_, index) => `
        <Relationship Id="rSheet${index + 1}" Target="worksheets/sheet${index + 1}.xml"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" />
      `)}
        <Relationship Id="rId2" Target="sharedStrings.xml"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" />
        <Relationship Id="rId3" Target="styles.xml"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" />
      </Relationships>
    `);
  },
  contentTypes(names: string[]) {
    return this.trim(`
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
        <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
        <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
        <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
      ${this.stringify(names, (_, index) => `
        <Override PartName="/xl/worksheets/sheet${index + 1}.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
      `)}
      </Types>
    `);
  },
  workbook(names: string[]) {
    return this.trim(`
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <workbook
          xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9303"/>
        <workbookPr defaultThemeVersion="124226"/>
        <bookViews>
          <workbookView xWindow="480" yWindow="60" windowWidth="18195" windowHeight="8505"/>
        </bookViews>
        <sheets>
        ${this.stringify(names, (name, index) => `
          <sheet name="${name}" sheetId="${index + 1}" r:id="rSheet${index + 1}" />
        `)}
        </sheets>
        <calcPr calcId="145621" />
      </workbook>
    `);
  },
  sheet(rows: Array<Array<Cell>>) {
    return this.trim(`
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <worksheet
          xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
          mc:Ignorable="x14ac"
          xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
        <sheetViews>
          <sheetView workbookViewId="0"/>
        </sheetViews>
        <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
        <sheetData>
        ${this.stringify(rows, (row, rowIndex) => `
          <row r="${rowIndex + 1}">
          ${this.stringify(row, (cell, cellIndex) => xml.cell(rowIndex, cellIndex, cell))}
          </row>
        `)}
        </sheetData>
      </worksheet>
    `);
  },
  cell(rowIndex: number, cellIndex: number, cell: Cell) {
    if (typeof cell.value === "number") {
      // 数値
      return `<c r="${ref(rowIndex, cellIndex)}" t="n"><v>${cell.value}</v></c>`;
    }
    // 文字列扱い
    return `<c r="${ref(rowIndex, cellIndex)}" t="inlineStr"><is><t>${this.escape((cell as any).value)}</t></is></c>`;
  },
  app() {
    return this.trim(`
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Properties
          xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
          xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
        <Application>${PACKAGE.name}</Application>
        <DocSecurity>0</DocSecurity>
        <ScaleCrop>false</ScaleCrop>
        <Company></Company>
        <LinksUpToDate>false</LinksUpToDate>
        <SharedDoc>false</SharedDoc>
        <HyperlinksChanged>false</HyperlinksChanged>
        <AppVersion>${PACKAGE.version}</AppVersion>
      </Properties>
    `);
  },
  core() {
    const today = new Date().toISOString();

    return this.trim(`
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <cp:coreProperties
          xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
          xmlns:dc="http://purl.org/dc/elements/1.1/"
          xmlns:dcterms="http://purl.org/dc/terms/"
          xmlns:dcmitype="http://purl.org/dc/dcmitype/"
          xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
        <dc:creator>${PACKAGE.name}</dc:creator>
        <cp:lastModifiedBy>${PACKAGE.name}</cp:lastModifiedBy>
        <dcterms:created xsi:type="dcterms:W3CDTF">${today}</dcterms:created>
        <dcterms:modified xsi:type="dcterms:W3CDTF">${today}</dcterms:modified>
      </cp:coreProperties>
    `);
  },
  sharedStrings() {
    return this.trim(`
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>
    `);
  },
};
