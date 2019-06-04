import { Writable } from "stream";
import archiver from "archiver";

import { StyledCell } from "./cell";
import { Style, Predicate } from "./style";
import {
  AnyValueConverter,
  DateValueConverter,
  NullConverter,
  NumberValueConverter,
  ValueConverter,
  ValueType,
} from "./value";
import { xml } from "./xml";

/**
 * xlsxファイルを作成するためのクラスです。
 */
export class WorkbookWriter {
  public static DEFAULT_CONVERTERS: ValueConverter<any>[] = [
    new NumberValueConverter(),
    new DateValueConverter(),
    new NullConverter(),
    new AnyValueConverter(),
  ];

  public converters: ValueConverter<any>[];
  private sheets: string[] = [];
  public archive: archiver.Archiver;
  public currentSheet?: SheetWriter;

  /**
   *
   * @param stream xlsxファイルの書き込み先
   */
  public constructor(stream: Writable) {
    this.converters = [...WorkbookWriter.DEFAULT_CONVERTERS];
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

  private rows: Array<Row>;
  private currentRow: Row = [];

  public sheetName: string;

  private index: number;

  public constructor(book: WorkbookWriter, sheetName: string, index: number) {
    this.book = book;
    this.sheetName = sheetName;
    this.index = index;
    this.rows = [];
    this.nextRow(); // 1行目
  }

  /**
   * このシートでスタイルを適用する条件を定義します。
   * predicateがtrueを返すときにnameで定義されたスタイルが適用されます。
   *
   * @param name スタイル名
   * @param predicate スタイルの適用条件
   */
  public addStyleRule(name: string, predicate: Predicate): void {}

  public addCell(cell: CellDef): void {
    const c: Partial<StyledCell> = {
      type: cell.type,
    };
    c.styles = [];

    if (cell.formula != null) {
      c.formula = cell.formula;
    } else if (cell.value != null) {
      const converter = this.findConverterFor(cell.value);
      c.value = converter.convert(cell.value);
      c.type = converter.type;
    }
    if (cell.styles != null) {
      c.styles.push(... cell.styles);
    }

    this.currentRow.push(c as StyledCell);
  }

  public addCellValue(value: unknown): void {
    const converter = this.findConverterFor(value);
    this.currentRow.push({
      value: converter.convert(value),
      type: converter.type,
      styles: [],
    });
  }

  private findConverterFor(value: unknown): ValueConverter<unknown> {
    for (const converter of this.book.converters) {
      if (converter.canApply(value)) {
        return converter;
      }
    }
    throw new Error("converter not found for " + value);
  }

  public skipCells(cells: number) {
    for (let i = 0; i < cells; i++) {
      this.currentRow.push({ type: "blank", styles: [] });
    }
  }

  public nextRow(): void {
    this.currentRow = [];
    this.rows.push(this.currentRow);
  }

  public skipRows(rows: number) {
    for (let i = 0; i < rows; i++) {
      this.nextRow();
    }
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

export interface CellDef {
  value?: unknown;
  formula?: string;
  type: ValueType;
  styles?: string[];
}

export type Row = Array<StyledCell>;
