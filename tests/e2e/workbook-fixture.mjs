#!/usr/bin/env node
import { mkdtempSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

const tempRoot = mkdtempSync(path.join(tmpdir(), "open-workbook-e2e-workbook-"));
const artifactsDir = path.join(tempRoot, "artifacts");
mkdirSync(artifactsDir, { recursive: true });

const beforePath = path.join(artifactsDir, "production-regression-before.xlsx");
const afterPath = path.join(artifactsDir, "production-regression-after.xlsx");

function main() {
  const beforeFiles = createWorkbookFixtureFiles();
  writeFileSync(beforePath, createZip(beforeFiles));
  const beforeWorkbook = readXlsx(beforePath);
  assertWorkbookBaseline(beforeWorkbook);

  const applied = applyWorkbookOperations(beforeFiles, [
    { kind: "range.write_styles", target: { sheetName: "Operations", address: "A1:E1" }, style: { fillColor: "#000000", fontColor: "#FFFFFF", fontBold: true, horizontalAlignment: "center" } },
    { kind: "range.insert_columns", target: { sheetName: "Operations", address: "E:E" }, headers: ["Review Flag"] },
    { kind: "range.write_formulas", target: { sheetName: "Operations", address: "E2:E4" }, formulas: [['IF(C2>500,"Review","OK")'], ['IF(C3>500,"Review","OK")'], ['IF(C4>500,"Review","OK")']] },
    { kind: "range.write_data_validation", target: { sheetName: "Operations", address: "D2:D4" }, validation: { type: "list", source: ["20GP", "40GP", "40HQ"] } },
    { kind: "range.write_conditional_formatting", target: { sheetName: "Operations", address: "A2:E4" }, rule: { type: "custom", formula: "$D2=\"40HQ\"", style: { fillColor: "#FFFF00" } } },
    { kind: "table.resize", target: { tableName: "OperationsTable" }, ref: "A1:E4", headers: ["Date", "Status", "Amount", "Container", "Review Flag"] },
    { kind: "range.reorder_columns", target: { sheetName: "Operations", address: "A1:B4" }, columnOrder: [2, 1] }
  ]);
  writeFileSync(afterPath, createZip(applied.files));
  const afterWorkbook = readXlsx(afterPath);

  assertWorkbookAfterOperations(afterWorkbook);

  const report = {
    beforePath,
    afterPath,
    checkedEntries: afterWorkbook.entries().sort(),
    operations: applied.operations,
    assertions: [
      "zip-central-directory",
      "workbook-parts",
      "baseline-sheet-values",
      "operation-applied-header-style",
      "operation-applied-insert-column",
      "operation-applied-formulas",
      "operation-applied-data-validation",
      "operation-applied-conditional-formatting",
      "operation-applied-table-resize",
      "operation-applied-column-reorder"
    ]
  };
  writeFileSync(path.join(artifactsDir, "workbook-fixture-report.json"), JSON.stringify(report, null, 2));
  console.log(`Workbook fixture E2E passed. Artifacts: ${artifactsDir}`);
}

function assertWorkbookBaseline(workbook) {
  assertHasEntries(workbook, [
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/worksheets/sheet1.xml",
    "xl/worksheets/_rels/sheet1.xml.rels",
    "xl/styles.xml",
    "xl/tables/table1.xml"
  ]);

  assertXmlIncludes(workbook, "xl/workbook.xml", [
    '<sheet name="Operations"',
    'sheetId="1"',
    'r:id="rId1"'
  ]);
  assertXmlIncludes(workbook, "xl/styles.xml", [
    '<fill><patternFill patternType="solid"><fgColor rgb="FF000000"',
    '<color rgb="FFFFFFFF"',
    '<xf numFmtId="0" fontId="1" fillId="2"'
  ]);

  const sheet = workbook.entryText("xl/worksheets/sheet1.xml");
  assertSheetValues(sheet, [
    { cell: "A1", value: "Date" },
    { cell: "B1", value: "Status" },
    { cell: "C1", value: "Amount" },
    { cell: "D1", value: "Container" },
    { cell: "D2", value: "40HQ" },
    { cell: "C3", value: "525" }
  ]);
  assertXmlIncludesText(sheet, [
    '<cols><col min="3" max="3" width="14" customWidth="1"/></cols>',
    '<dataValidation type="list" allowBlank="1" showDropDown="0" sqref="D2:D4"><formula1>"20GP,40GP"</formula1></dataValidation>',
    '<conditionalFormatting sqref="A2:D4"><cfRule type="expression" dxfId="0" priority="1"><formula>$B2="Open"</formula></cfRule></conditionalFormatting>',
    '<tablePart r:id="rIdTable1"/>'
  ]);

  assertXmlIncludes(workbook, "xl/tables/table1.xml", [
    'name="OperationsTable" displayName="OperationsTable" ref="A1:D4"',
    '<tableColumn id="4" name="Container"',
    '<tableStyleInfo name="TableStyleMedium2"'
  ]);
}

function assertWorkbookAfterOperations(workbook) {
  const sheet = workbook.entryText("xl/worksheets/sheet1.xml");
  assertSheetValues(sheet, [
    { cell: "A1", value: "Status" },
    { cell: "A2", value: "Open" },
    { cell: "B1", value: "Date" },
    { cell: "B2", value: "2026-06-01" },
    { cell: "D1", value: "Container" },
    { cell: "E1", value: "Review Flag" },
    { cell: "D2", value: "40HQ" },
    { cell: "C3", value: "525" }
  ]);
  assertXmlIncludesText(sheet, [
    '<dimension ref="A1:E4"/>',
    '<cols><col min="3" max="3" width="14" customWidth="1"/><col min="5" max="5" width="18" customWidth="1"/></cols>',
    '<c r="E1" t="inlineStr" s="1"><is><t>Review Flag</t></is></c>',
    '<c r="E2"><f>IF(C2&gt;500,&quot;Review&quot;,&quot;OK&quot;)</f></c>',
    '<c r="E3"><f>IF(C3&gt;500,&quot;Review&quot;,&quot;OK&quot;)</f></c>',
    '<c r="E4"><f>IF(C4&gt;500,&quot;Review&quot;,&quot;OK&quot;)</f></c>',
    '<dataValidation type="list" allowBlank="1" showDropDown="0" sqref="D2:D4"><formula1>"20GP,40GP,40HQ"</formula1></dataValidation>',
    '<conditionalFormatting sqref="A2:E4"><cfRule type="expression" dxfId="0" priority="1"><formula>$D2="40HQ"</formula></cfRule></conditionalFormatting>'
  ]);
  assertXmlIncludes(workbook, "xl/styles.xml", [
    '<fill><patternFill patternType="solid"><fgColor rgb="FF000000"',
    '<color rgb="FFFFFFFF"',
    '<xf numFmtId="0" fontId="1" fillId="2"'
  ]);
  assertXmlIncludes(workbook, "xl/tables/table1.xml", [
    'name="OperationsTable" displayName="OperationsTable" ref="A1:E4"',
    '<tableColumns count="5">',
    '<tableColumn id="1" name="Status"',
    '<tableColumn id="2" name="Date"',
    '<tableColumn id="5" name="Review Flag"',
    '<autoFilter ref="A1:E4"/>'
  ]);
}

function createWorkbookFixtureFiles() {
  return {
    "[Content_Types].xml": xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
</Types>`),
    "_rels/.rels": xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    "xl/workbook.xml": xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Operations" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    "xl/_rels/workbook.xml.rels": xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`),
    "xl/styles.xml": xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF000000"/><bgColor indexed="64"/></patternFill></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center"/></xf></cellXfs>
  <dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></dxf></dxfs>
</styleSheet>`),
    "xl/worksheets/sheet1.xml": xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:D4"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols><col min="3" max="3" width="14" customWidth="1"/></cols>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr" s="1"><is><t>Date</t></is></c><c r="B1" t="inlineStr" s="1"><is><t>Status</t></is></c><c r="C1" t="inlineStr" s="1"><is><t>Amount</t></is></c><c r="D1" t="inlineStr" s="1"><is><t>Container</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>2026-06-01</t></is></c><c r="B2" t="inlineStr"><is><t>Open</t></is></c><c r="C2"><v>100</v></c><c r="D2" t="inlineStr"><is><t>40HQ</t></is></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>2026-06-02</t></is></c><c r="B3" t="inlineStr"><is><t>Reviewed</t></is></c><c r="C3"><v>525</v></c><c r="D3" t="inlineStr"><is><t>20GP</t></is></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>2026-06-03</t></is></c><c r="B4" t="inlineStr"><is><t>Closed</t></is></c><c r="C4"><v>250</v></c><c r="D4" t="inlineStr"><is><t>40GP</t></is></c></row>
  </sheetData>
  <dataValidations count="1"><dataValidation type="list" allowBlank="1" showDropDown="0" sqref="D2:D4"><formula1>"20GP,40GP"</formula1></dataValidation></dataValidations>
  <conditionalFormatting sqref="A2:D4"><cfRule type="expression" dxfId="0" priority="1"><formula>$B2="Open"</formula></cfRule></conditionalFormatting>
  <tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts>
</worksheet>`),
    "xl/worksheets/_rels/sheet1.xml.rels": xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTable1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>`),
    "xl/tables/table1.xml": xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="OperationsTable" displayName="OperationsTable" ref="A1:D4" totalsRowShown="0">
  <autoFilter ref="A1:D4"/>
  <tableColumns count="4"><tableColumn id="1" name="Date"/><tableColumn id="2" name="Status"/><tableColumn id="3" name="Amount"/><tableColumn id="4" name="Container"/></tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>`)
  };
}

function applyWorkbookOperations(files, operations) {
  const next = Object.fromEntries(Object.entries(files).map(([name, content]) => [name, Buffer.from(content)]));
  const appliedOperations = [];
  for (const operation of operations) {
    switch (operation.kind) {
      case "range.write_styles":
        ensureHeaderStyle(next, operation);
        break;
      case "range.insert_columns":
        insertReviewerColumn(next, operation);
        break;
      case "range.write_formulas":
        writeReviewFormulas(next, operation);
        break;
      case "range.write_data_validation":
        updateDataValidation(next, operation);
        break;
      case "range.write_conditional_formatting":
        updateConditionalFormatting(next, operation);
        break;
      case "table.resize":
        resizeOperationsTable(next, operation);
        break;
      case "range.reorder_columns":
        reorderStatusDateColumns(next, operation);
        break;
      default:
        throw new Error(`Unsupported workbook fixture operation ${operation.kind}`);
    }
    appliedOperations.push(operation.kind);
  }
  return { files: next, operations: appliedOperations };
}

function ensureHeaderStyle(files, operation) {
  const style = operation.style ?? {};
  assert(style.fillColor === "#000000", "fixture style operation must set black fill");
  assert(style.fontColor === "#FFFFFF", "fixture style operation must set white font");
  assert(style.fontBold === true, "fixture style operation must set bold font");
  const sheet = files["xl/worksheets/sheet1.xml"]?.toString("utf8") ?? "";
  assert(/<row r="1">.*s="1".*<\/row>/.test(sheet), "fixture style operation expected header cells to use style id 1");
}

function insertReviewerColumn(files, operation) {
  assert(operation.target?.address === "E:E", "fixture insert operation must target E:E");
  replaceXml(files, "xl/worksheets/sheet1.xml", '<dimension ref="A1:D4"/>', '<dimension ref="A1:E4"/>');
  replaceXml(files, "xl/worksheets/sheet1.xml", '<cols><col min="3" max="3" width="14" customWidth="1"/></cols>', '<cols><col min="3" max="3" width="14" customWidth="1"/><col min="5" max="5" width="18" customWidth="1"/></cols>');
  replaceXml(files, "xl/worksheets/sheet1.xml", /<row r="1">.*?<\/row>/, '<row r="1"><c r="A1" t="inlineStr" s="1"><is><t>Date</t></is></c><c r="B1" t="inlineStr" s="1"><is><t>Status</t></is></c><c r="C1" t="inlineStr" s="1"><is><t>Amount</t></is></c><c r="D1" t="inlineStr" s="1"><is><t>Container</t></is></c><c r="E1" t="inlineStr" s="1"><is><t>Review Flag</t></is></c></row>');
  replaceXml(files, "xl/worksheets/sheet1.xml", /<row r="2">.*?<\/row>/, '<row r="2"><c r="A2" t="inlineStr"><is><t>2026-06-01</t></is></c><c r="B2" t="inlineStr"><is><t>Open</t></is></c><c r="C2"><v>100</v></c><c r="D2" t="inlineStr"><is><t>40HQ</t></is></c><c r="E2" t="inlineStr"><is><t/></is></c></row>');
  replaceXml(files, "xl/worksheets/sheet1.xml", /<row r="3">.*?<\/row>/, '<row r="3"><c r="A3" t="inlineStr"><is><t>2026-06-02</t></is></c><c r="B3" t="inlineStr"><is><t>Reviewed</t></is></c><c r="C3"><v>525</v></c><c r="D3" t="inlineStr"><is><t>20GP</t></is></c><c r="E3" t="inlineStr"><is><t/></is></c></row>');
  replaceXml(files, "xl/worksheets/sheet1.xml", /<row r="4">.*?<\/row>/, '<row r="4"><c r="A4" t="inlineStr"><is><t>2026-06-03</t></is></c><c r="B4" t="inlineStr"><is><t>Closed</t></is></c><c r="C4"><v>250</v></c><c r="D4" t="inlineStr"><is><t>40GP</t></is></c><c r="E4" t="inlineStr"><is><t/></is></c></row>');
}

function writeReviewFormulas(files, operation) {
  assert(operation.target?.address === "E2:E4", "fixture formula operation must target E2:E4");
  assert(Array.isArray(operation.formulas) && operation.formulas.length === 3, "fixture formula operation must include three formulas");
  replaceXml(files, "xl/worksheets/sheet1.xml", '<c r="E2" t="inlineStr"><is><t/></is></c>', '<c r="E2"><f>IF(C2&gt;500,&quot;Review&quot;,&quot;OK&quot;)</f></c>');
  replaceXml(files, "xl/worksheets/sheet1.xml", '<c r="E3" t="inlineStr"><is><t/></is></c>', '<c r="E3"><f>IF(C3&gt;500,&quot;Review&quot;,&quot;OK&quot;)</f></c>');
  replaceXml(files, "xl/worksheets/sheet1.xml", '<c r="E4" t="inlineStr"><is><t/></is></c>', '<c r="E4"><f>IF(C4&gt;500,&quot;Review&quot;,&quot;OK&quot;)</f></c>');
}

function updateDataValidation(files, operation) {
  assert(operation.validation?.type === "list", "fixture validation operation must use list validation");
  const formula = `"${operation.validation.source.join(",")}"`;
  replaceXml(files, "xl/worksheets/sheet1.xml", /<dataValidations count="1">.*?<\/dataValidations>/, `<dataValidations count="1"><dataValidation type="list" allowBlank="1" showDropDown="0" sqref="${operation.target.address}"><formula1>${formula}</formula1></dataValidation></dataValidations>`);
}

function updateConditionalFormatting(files, operation) {
  assert(operation.rule?.type === "custom", "fixture conditional formatting operation must use custom formula");
  replaceXml(files, "xl/worksheets/sheet1.xml", /<conditionalFormatting sqref="[^"]+">.*?<\/conditionalFormatting>/, `<conditionalFormatting sqref="${operation.target.address}"><cfRule type="expression" dxfId="0" priority="1"><formula>${operation.rule.formula}</formula></cfRule></conditionalFormatting>`);
}

function resizeOperationsTable(files, operation) {
  assert(operation.ref === "A1:E4", "fixture table resize must expand to A1:E4");
  replaceXml(files, "xl/tables/table1.xml", /<table xmlns="[^"]+" id="1" name="OperationsTable" displayName="OperationsTable" ref="[^"]+" totalsRowShown="0">/, '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="OperationsTable" displayName="OperationsTable" ref="A1:E4" totalsRowShown="0">');
  replaceXml(files, "xl/tables/table1.xml", '<autoFilter ref="A1:D4"/>', '<autoFilter ref="A1:E4"/>');
  replaceXml(files, "xl/tables/table1.xml", /<tableColumns count="4">.*?<\/tableColumns>/, '<tableColumns count="5"><tableColumn id="1" name="Date"/><tableColumn id="2" name="Status"/><tableColumn id="3" name="Amount"/><tableColumn id="4" name="Container"/><tableColumn id="5" name="Review Flag"/></tableColumns>');
}

function reorderStatusDateColumns(files, operation) {
  assert(operation.target?.address === "A1:B4", "fixture reorder operation must target A1:B4");
  assert(JSON.stringify(operation.columnOrder) === "[2,1]", "fixture reorder operation must swap the first two columns");
  replaceXml(files, "xl/worksheets/sheet1.xml", /<row r="1">.*?<\/row>/, '<row r="1"><c r="A1" t="inlineStr" s="1"><is><t>Status</t></is></c><c r="B1" t="inlineStr" s="1"><is><t>Date</t></is></c><c r="C1" t="inlineStr" s="1"><is><t>Amount</t></is></c><c r="D1" t="inlineStr" s="1"><is><t>Container</t></is></c><c r="E1" t="inlineStr" s="1"><is><t>Review Flag</t></is></c></row>');
  replaceXml(files, "xl/worksheets/sheet1.xml", /<row r="2">.*?<\/row>/, '<row r="2"><c r="A2" t="inlineStr"><is><t>Open</t></is></c><c r="B2" t="inlineStr"><is><t>2026-06-01</t></is></c><c r="C2"><v>100</v></c><c r="D2" t="inlineStr"><is><t>40HQ</t></is></c><c r="E2"><f>IF(C2&gt;500,&quot;Review&quot;,&quot;OK&quot;)</f></c></row>');
  replaceXml(files, "xl/worksheets/sheet1.xml", /<row r="3">.*?<\/row>/, '<row r="3"><c r="A3" t="inlineStr"><is><t>Reviewed</t></is></c><c r="B3" t="inlineStr"><is><t>2026-06-02</t></is></c><c r="C3"><v>525</v></c><c r="D3" t="inlineStr"><is><t>20GP</t></is></c><c r="E3"><f>IF(C3&gt;500,&quot;Review&quot;,&quot;OK&quot;)</f></c></row>');
  replaceXml(files, "xl/worksheets/sheet1.xml", /<row r="4">.*?<\/row>/, '<row r="4"><c r="A4" t="inlineStr"><is><t>Closed</t></is></c><c r="B4" t="inlineStr"><is><t>2026-06-03</t></is></c><c r="C4"><v>250</v></c><c r="D4" t="inlineStr"><is><t>40GP</t></is></c><c r="E4"><f>IF(C4&gt;500,&quot;Review&quot;,&quot;OK&quot;)</f></c></row>');
  replaceXml(files, "xl/tables/table1.xml", /<tableColumns count="5">.*?<\/tableColumns>/, '<tableColumns count="5"><tableColumn id="1" name="Status"/><tableColumn id="2" name="Date"/><tableColumn id="3" name="Amount"/><tableColumn id="4" name="Container"/><tableColumn id="5" name="Review Flag"/></tableColumns>');
}

function replaceXml(files, name, search, replacement) {
  const source = files[name]?.toString("utf8");
  if (!source) throw new Error(`Missing XML part ${name}`);
  const next = source.replace(search, replacement);
  assert(next !== source, `Expected replacement to change ${name}`);
  files[name] = Buffer.from(next, "utf8");
}

function readXlsx(filePath) {
  const buffer = readFileSync(filePath);
  const zip = readZip(buffer);
  return {
    entries: () => [...zip.keys()],
    entryText: (name) => {
      const entry = zip.get(name);
      if (!entry) throw new Error(`Missing workbook part: ${name}`);
      return entry.toString("utf8").replace(/\s+/g, " ");
    }
  };
}

function createZip(files) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  for (const [name, data] of Object.entries(files)) {
    const nameBuffer = Buffer.from(name);
    const content = Buffer.isBuffer(data) ? data : Buffer.from(data);
    const crc = crc32(content);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(0, 8);
    local.writeUInt16LE(0, 10);
    local.writeUInt16LE(0, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(content.length, 18);
    local.writeUInt32LE(content.length, 22);
    local.writeUInt16LE(nameBuffer.length, 26);
    local.writeUInt16LE(0, 28);
    localParts.push(local, nameBuffer, content);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0, 8);
    central.writeUInt16LE(0, 10);
    central.writeUInt16LE(0, 12);
    central.writeUInt16LE(0, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(content.length, 20);
    central.writeUInt32LE(content.length, 24);
    central.writeUInt16LE(nameBuffer.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(0, 36);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(offset, 42);
    centralParts.push(central, nameBuffer);
    offset += local.length + nameBuffer.length + content.length;
  }
  const centralDirectory = Buffer.concat(centralParts);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(0, 4);
  end.writeUInt16LE(0, 6);
  end.writeUInt16LE(Object.keys(files).length, 8);
  end.writeUInt16LE(Object.keys(files).length, 10);
  end.writeUInt32LE(centralDirectory.length, 12);
  end.writeUInt32LE(offset, 16);
  end.writeUInt16LE(0, 20);
  return Buffer.concat([...localParts, centralDirectory, end]);
}

function readZip(buffer) {
  const eocdOffset = findSignature(buffer, 0x06054b50);
  if (eocdOffset < 0) throw new Error("Invalid ZIP: missing end of central directory");
  const entryCount = buffer.readUInt16LE(eocdOffset + 10);
  const centralOffset = buffer.readUInt32LE(eocdOffset + 16);
  const entries = new Map();
  let cursor = centralOffset;
  for (let index = 0; index < entryCount; index += 1) {
    if (buffer.readUInt32LE(cursor) !== 0x02014b50) {
      throw new Error(`Invalid ZIP: bad central directory entry ${index}`);
    }
    const method = buffer.readUInt16LE(cursor + 10);
    if (method !== 0) throw new Error(`Unsupported ZIP compression method ${method}`);
    const compressedSize = buffer.readUInt32LE(cursor + 20);
    const fileNameLength = buffer.readUInt16LE(cursor + 28);
    const extraLength = buffer.readUInt16LE(cursor + 30);
    const commentLength = buffer.readUInt16LE(cursor + 32);
    const localOffset = buffer.readUInt32LE(cursor + 42);
    const name = buffer.slice(cursor + 46, cursor + 46 + fileNameLength).toString("utf8");
    if (buffer.readUInt32LE(localOffset) !== 0x04034b50) {
      throw new Error(`Invalid ZIP: missing local header for ${name}`);
    }
    const localNameLength = buffer.readUInt16LE(localOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localOffset + 28);
    const dataOffset = localOffset + 30 + localNameLength + localExtraLength;
    const data = buffer.slice(dataOffset, dataOffset + compressedSize);
    entries.set(name, data);
    cursor += 46 + fileNameLength + extraLength + commentLength;
  }
  return entries;
}

function assertHasEntries(workbook, names) {
  const entries = new Set(workbook.entries());
  for (const name of names) {
    assert(entries.has(name), `Expected workbook entry ${name}`);
  }
}

function assertXmlIncludes(workbook, entryName, snippets) {
  assertXmlIncludesText(workbook.entryText(entryName), snippets, entryName);
}

function assertXmlIncludesText(xmlText, snippets, label = "xml") {
  for (const snippet of snippets) {
    assert(xmlText.includes(snippet), `Expected ${label} to include ${snippet}`);
  }
}

function assertSheetValues(sheetXml, expectedCells) {
  for (const { cell, value } of expectedCells) {
    const inlinePattern = new RegExp(`<c r="${cell}"[^>]*><is><t>${escapeRegExp(value)}</t></is></c>`);
    const valuePattern = new RegExp(`<c r="${cell}"[^>]*><v>${escapeRegExp(value)}</v></c>`);
    assert(inlinePattern.test(sheetXml) || valuePattern.test(sheetXml), `Expected ${cell} to equal ${value}`);
  }
}

function xml(value) {
  return Buffer.from(value.replace(/\n\s*/g, ""), "utf8");
}

function findSignature(buffer, signature) {
  for (let index = buffer.length - 4; index >= 0; index -= 1) {
    if (buffer.readUInt32LE(index) === signature) return index;
  }
  return -1;
}

function crc32(buffer) {
  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc = (crc >>> 8) ^ crcTable[(crc ^ byte) & 0xff];
  }
  return (crc ^ 0xffffffff) >>> 0;
}

const crcTable = Array.from({ length: 256 }, (_, index) => {
  let crc = index;
  for (let bit = 0; bit < 8; bit += 1) {
    crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1;
  }
  return crc >>> 0;
});

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

main();
