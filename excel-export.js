const COLUMNS = ["A", "B", "C", "D", "E", "F", "G", "H"];

function xmlEscape(value) {
  return String(value ?? "").replace(/[&<>"']/g, (character) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&apos;",
  }[character]));
}

function cellXml(reference, value, style, numeric = false) {
  const styleAttribute = style ? ` s="${style}"` : "";
  if (value === "" || value === null || value === undefined) {
    return `<c r="${reference}"${styleAttribute}/>`;
  }
  if (numeric && Number.isFinite(Number(value))) {
    return `<c r="${reference}"${styleAttribute} t="n"><v>${Number(value)}</v></c>`;
  }
  return `<c r="${reference}"${styleAttribute} t="inlineStr"><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
}

function rowXml(rowNumber, cells) {
  return `<row r="${rowNumber}">${cells.map((cell, index) =>
    cellXml(`${COLUMNS[index]}${rowNumber}`, cell.value, cell.style, cell.numeric)
  ).join("")}</row>`;
}

function makeWorksheet(sheet, computedRows) {
  const rows = [];
  const merges = ["A1:H1", "A5:C5", "D5:F5", "G5:G6", "H5:H6"];
  const addRow = (cells) => {
    const rowNumber = rows.length + 1;
    rows.push(rowXml(rowNumber, cells));
  };
  const label = (value) => ({ value, style: 4 });
  const value = (text) => ({ value: text });
  const header = (text) => ({ value: text, style: 3 });
  const group = (text) => ({ value: text, style: 2 });
  const blank = () => ({ value: "", style: 5 });

  addRow([{ value: "FIELD LEVEL BOOK", style: 1 }]);
  addRow([
    label("Date"), value(sheet.date), label("Level Type"), value(sheet.levelType),
    label("Surveyor"), value(sheet.surveyor), label("Booker"), value(sheet.booker),
  ]);
  addRow([
    label("From Point"), value(sheet.fromPoint), label("To Point"), value(sheet.toPoint),
    label("Starting RL"), value(sheet.startRL), blank(), blank(),
  ]);
  addRow([]);
  addRow([group("Staff Readings"), blank(), blank(), group("Elevation"), blank(), blank(), header("Distance"), header("Remarks")]);
  addRow([
    header("Back Sight"), header("Inter. Sight"), header("Fore Sight"),
    header("Rise"), header("Fall"), header("Reduced Level"), blank(), blank(),
  ]);

  computedRows.forEach((row) => {
    const numericCell = (raw) => ({ value: raw, style: 6, numeric: true });
    const computedCell = (number) => ({ value: number === null ? "" : number.toFixed(3), style: 6, numeric: number !== null });
    addRow([
      numericCell(row.bs), numericCell(row.is), numericCell(row.fs),
      computedCell(row.rise), computedCell(row.fall), computedCell(row.rl),
      numericCell(row.distance), value(row.remarks),
    ]);
  });

  const lastDataRow = rows.length;
  const checkTitleRow = lastDataRow + 2;
  merges.push(`A${checkTitleRow}:H${checkTitleRow}`);
  rows.push(rowXml(checkTitleRow, [{ value: "ARITHMETIC CHECK", style: 1 }]));

  const numericValues = (property) => computedRows.reduce((total, row) => {
    const number = parseFloat(row[property]);
    return total + (Number.isFinite(number) ? number : 0);
  }, 0);
  const sumBs = numericValues("bs");
  const sumFs = numericValues("fs");
  const sumRise = numericValues("rise");
  const sumFall = numericValues("fall");
  const firstRL = computedRows.length ? computedRows[0].rl : null;
  const lastRL = computedRows.length ? computedRows[computedRows.length - 1].rl : null;
  const checks = [
    ["Σ Back Sight", sumBs],
    ["Σ Fore Sight", sumFs],
    ["ΣBS − ΣFS", sumBs - sumFs],
    ["Σ Rise − Σ Fall", sumRise - sumFall],
    ["Last RL − First RL", firstRL !== null && lastRL !== null ? lastRL - firstRL : null],
  ];
  checks.forEach(([checkLabel, number], index) => {
    const rowNumber = checkTitleRow + index + 1;
    rows.push(rowXml(rowNumber, [label(checkLabel), { value: number === null ? "—" : number.toFixed(3), style: 5, numeric: number !== null }]));
  });

  const finalRow = checkTitleRow + checks.length;
  const worksheetRows = rows.join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:H${finalRow}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>${COLUMNS.map((column, index) => `<col min="${index + 1}" max="${index + 1}" width="${index % 2 === 0 ? 18 : 20}" customWidth="1"/>`).join("")}</cols>
  <sheetData>${worksheetRows}</sheetData>
  <mergeCells count="${merges.length}">${merges.map((range) => `<mergeCell ref="${range}"/>`).join("")}</mergeCells>
</worksheet>`;
}

function workbookFiles(sheet, computedRows) {
  const types = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;
  const rootRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
  const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Field Book" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;
  const workbookRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="16"/><color rgb="FF1E2A3A"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1E2A3A"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFC3CFD9"/></left><right style="thin"><color rgb="FFC3CFD9"/></right><top style="thin"><color rgb="FFC3CFD9"/></top><bottom style="thin"><color rgb="FFC3CFD9"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="7">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
  return [
    ["[Content_Types].xml", types],
    ["_rels/.rels", rootRelationships],
    ["xl/workbook.xml", workbook],
    ["xl/_rels/workbook.xml.rels", workbookRelationships],
    ["xl/worksheets/sheet1.xml", makeWorksheet(sheet, computedRows)],
    ["xl/styles.xml", styles],
  ];
}

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let index = 0; index < table.length; index++) {
    let value = index;
    for (let bit = 0; bit < 8; bit++) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    table[index] = value >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function write16(view, offset, value) {
  view.setUint16(offset, value, true);
}

function write32(view, offset, value) {
  view.setUint32(offset, value >>> 0, true);
}

function zipStore(files) {
  const encoder = new TextEncoder();
  const localParts = [];
  const centralParts = [];
  let localOffset = 0;

  for (const [fileName, content] of files) {
    const nameBytes = encoder.encode(fileName);
    const data = encoder.encode(content);
    const checksum = crc32(data);
    const localHeader = new Uint8Array(30 + nameBytes.length);
    const localView = new DataView(localHeader.buffer);
    write32(localView, 0, 0x04034b50);
    write16(localView, 4, 20);
    write16(localView, 6, 0x0800);
    write16(localView, 8, 0);
    write32(localView, 14, checksum);
    write32(localView, 18, data.length);
    write32(localView, 22, data.length);
    write16(localView, 26, nameBytes.length);
    localHeader.set(nameBytes, 30);
    localParts.push(localHeader, data);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    const centralView = new DataView(centralHeader.buffer);
    write32(centralView, 0, 0x02014b50);
    write16(centralView, 4, 20);
    write16(centralView, 6, 20);
    write16(centralView, 8, 0x0800);
    write16(centralView, 10, 0);
    write32(centralView, 16, checksum);
    write32(centralView, 20, data.length);
    write32(centralView, 24, data.length);
    write16(centralView, 28, nameBytes.length);
    write32(centralView, 42, localOffset);
    centralHeader.set(nameBytes, 46);
    centralParts.push(centralHeader);
    localOffset += localHeader.length + data.length;
  }

  const centralSize = centralParts.reduce((total, part) => total + part.length, 0);
  const end = new Uint8Array(22);
  const endView = new DataView(end.buffer);
  write32(endView, 0, 0x06054b50);
  write16(endView, 8, files.length);
  write16(endView, 10, files.length);
  write32(endView, 12, centralSize);
  write32(endView, 16, localOffset);

  const archive = new Uint8Array(localOffset + centralSize + end.length);
  let offset = 0;
  for (const part of [...localParts, ...centralParts, end]) {
    archive.set(part, offset);
    offset += part.length;
  }
  return archive;
}

export function createFieldBookWorkbook(sheet, computedRows) {
  return zipStore(workbookFiles(sheet, computedRows));
}
