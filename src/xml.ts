import { ref, StyledCell } from "./cell";

const PACKAGE = require("../package.json");

const types: { [key: string]: string } = {
  string: "s",
  number: "n",
  date: "d",
  boolean: "b",
};

export const escape = (s?: string | null) => {
  if (s == null) {
    return "";
  }

  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&t;")
    .replace(/"/g, "&quot;");
};
export const stringify = <T>(
  array: Array<T>,
  mapper: (v: T, n: number) => string
) => {
  return array
    .map(mapper)
    .map(s => s.trim())
    .join("");
};
export const trim = (s: string) => {
  return s.trim().replace(/\n */g, " ");
};

export const xml = {
  rels() {
    return trim(`
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
    return trim(`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        ${stringify(
          names,
          (_, index) => `
          <Relationship Id="rSheet${index + 1}" Target="worksheets/sheet${index + 1}.xml"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" />
        `
        )}
          <Relationship Id="rId2" Target="sharedStrings.xml"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" />
          <Relationship Id="rId3" Target="styles.xml"
              Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" />
        </Relationships>
      `);
  },
  contentTypes(names: string[]) {
    return trim(`
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
          <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
          <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
          <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
        ${stringify(
          names,
          (_, index) => `
          <Override PartName="/xl/worksheets/sheet${index + 1}.xml"
              ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
        `)}
        </Types>
      `);
  },
  workbook(names: string[]) {
    return trim(`
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
          ${stringify(
            names,
            (name, index) => `
            <sheet name="${name}" sheetId="${index + 1}" r:id="rSheet${index + 1}" />
          `
          )}
          </sheets>
          <calcPr calcId="145621" />
        </workbook>
      `);
  },
  sheet(rows: Array<Array<StyledCell>>) {
    return trim(`
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
          ${stringify(
            rows,
            (row, rowIndex) => `
            <row r="${rowIndex + 1}">
            ${stringify(row, (cell, cellIndex) =>
              xml.cell(rowIndex, cellIndex, cell)
            )}
            </row>
          `)}
          </sheetData>
        </worksheet>
      `);
  },
  cell(rowIndex: number, cellIndex: number, cell: StyledCell): string {
    const r = ref(rowIndex, cellIndex);
    if (cell.formula != null) {
      // 数式
      return `
        <c r="${r}" t="${types[cell.type]}">
          <f>${escape(cell.formula)}</f>
        </c>`;
    } else if (cell.type === "number") {
      // 数値
      return `<c r="${r}" t="n">
        <v>${cell.value}</v>
      </c>`;
    } else if (cell.type === "date") {
      // 日付
      return `<c r="${r}" t="d"><v>${cell.value}</v></c>`;
    } else if (cell.type === "string") {
      // 文字列扱い
      return `<c r="${r}" t="inlineStr">
        <is><t>${escape(cell.value)}</t></is>
      </c>`;
    } else {
      return `<c r="${r}"></c>`;
    }
  },
  app() {
    return trim(`
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

    return trim(`
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
    return trim(`
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>
    `);
  }
};
