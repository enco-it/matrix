/* global JSZip, MatrixAgent */
(function (global) {
  let templateBuf = null;

  function letterToIndex(letter) {
    let n = 0;
    for (let i = 0; i < letter.length; i++) n = n * 26 + (letter.charCodeAt(i) - 64);
    return n;
  }

  function retargetFormula(formula, row, last) {
    if (global.MatrixAgent && MatrixAgent.retargetFormula) {
      return MatrixAgent.retargetFormula(formula, row, last);
    }
    let s = String(formula);
    s = s.replace(/\$([A-Za-z]{1,3})\$8:\$\1\$\d+/g, (_, col) => `$${col}$8:$${col}$${last}`);
    s = s.replace(/\$8:\$\d+/g, `$8:$${last}`);
    s = s.replace(/(\$?[A-Za-z]{1,3})(\$?)8(?!\d)/g, (m, col, abs) => (abs ? m : col + row));
    return s;
  }

  async function loadTemplate() {
    if (templateBuf) return templateBuf;
    const res = await fetch("template.xlsm");
    if (!res.ok) throw new Error("Не загружен шаблон ГП-10 (template.xlsm)");
    templateBuf = await res.arrayBuffer();
    return templateBuf;
  }

  function xmlEscape(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function cellCol(xml) {
    const m = xml.match(/\br="([A-Z]+)\d+"/);
    return m ? m[1] : "";
  }

  function styleAttr(xml) {
    const m = xml.match(/\bs="(\d+)"/);
    return m ? ` s="${m[1]}"` : "";
  }

  function writeCell(src, col, row, value) {
    const st = styleAttr(src);
    const ref = `r="${col}${row}"`;
    if (value === null || value === undefined || value === "") {
      return `<c ${ref}${st}/>`;
    }
    if (typeof value === "number" && Number.isFinite(value)) {
      return `<c ${ref}${st}><v>${value}</v></c>`;
    }
    return `<c ${ref}${st} t="inlineStr"><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
  }

  function cloneDataRow(row8xml, rowNum, last, values) {
    const open = row8xml.match(/^<row([^>]*)>/);
    const closeAt = row8xml.lastIndexOf("</row>");
    const attrs = open[1].replace(/\br="8"/, `r="${rowNum}"`);
    const inner = row8xml.slice(open[0].length, closeAt);
    const parts = inner.match(/<c\b[^>]*\/>|<c\b[^>]*>[\s\S]*?<\/c>/g) || [];
    const byCol = {};
    parts.forEach((c) => {
      const col = cellCol(c);
      if (col) byCol[col] = c;
    });
    Object.keys(values).forEach((col) => {
      if (!byCol[col]) byCol[col] = `<c r="${col}8"/>`;
    });
    const cols = Object.keys(byCol).sort((a, b) => letterToIndex(a) - letterToIndex(b));
    const cells = cols.map((col) => {
      const src = byCol[col];
      if (Object.prototype.hasOwnProperty.call(values, col)) {
        return writeCell(src, col, rowNum, values[col]);
      }
      let c = src.replace(/<f t="shared"[^>]*>([\s\S]*?)<\/f>/g, "<f>$1</f>");
      c = c.replace(/\br="([A-Z]+)8"/, `r="$1${rowNum}"`);
      if (/<f[ >]/.test(c)) {
        c = c.replace(/<f>([\s\S]*?)<\/f>/g, (_, body) => `<f>${retargetFormula(body, rowNum, last)}</f>`);
        c = c.replace(/<v>[\s\S]*?<\/v>/g, "");
        return c;
      }
      return writeCell(src, col, rowNum, null);
    });
    return `<row${attrs}>${cells.join("")}</row>`;
  }

  function rowValues(row) {
    const out = {};
    Object.keys(row).forEach((k) => {
      if (k === "notes" || row[k] === undefined) return;
      if (!/^[A-Z]+$/.test(k)) return;
      out[k] = row[k];
    });
    return out;
  }

  async function stripVba(zip) {
    zip.remove("xl/vbaProject.bin");
    let ct = await zip.file("[Content_Types].xml").async("string");
    ct = ct.replace(/<Override PartName="\/xl\/vbaProject.bin"[^/]*\/>/, "");
    ct = ct.replace(
      "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    );
    zip.file("[Content_Types].xml", ct);
    let rels = await zip.file("xl/_rels/workbook.xml.rels").async("string");
    rels = rels.replace(/<Relationship [^>]*vbaProject[^/]*\/>/, "");
    zip.file("xl/_rels/workbook.xml.rels", rels);
  }

  async function exportWorkbook(rows, meta, fmt) {
    if (typeof JSZip === "undefined") throw new Error("JSZip не загружен");
    const buf = await loadTemplate();
    const zip = await JSZip.loadAsync(buf);
    let sheet = await zip.file("xl/worksheets/sheet1.xml").async("string");
    const start = sheet.indexOf('<row r="8"');
    const end = sheet.indexOf("</row>", start);
    if (start < 0 || end < 0) throw new Error("В шаблоне нет строки 8");
    const row8 = sheet.slice(start, end + 6);
    const last = 7 + rows.length;
    const dataXml = rows.map((row, i) => cloneDataRow(row8, 8 + i, last, rowValues(row))).join("");
    sheet = sheet.slice(0, start) + dataXml + sheet.slice(end + 6);
    sheet = sheet.replace(/ref="A1:MK\d+"/, `ref="A1:MK${last}"`);
    sheet = sheet.replace(/5202/g, String(last)).replace(/5201/g, String(last));
    zip.file("xl/worksheets/sheet1.xml", sheet);
    if (zip.file("xl/workbook.xml")) {
      let wb = await zip.file("xl/workbook.xml").async("string");
      wb = wb.replace(/\$MK\$\d+/g, `$MK$${last}`);
      zip.file("xl/workbook.xml", wb);
    }
    const asXlsm = fmt !== "xlsx";
    if (!asXlsm) await stripVba(zip);
    const mime = asXlsm
      ? "application/vnd.ms-excel.sheet.macroEnabled.12"
      : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    return zip.generateAsync({ type: "blob", mimeType: mime, compression: "DEFLATE" });
  }

  global.MatrixAgent = global.MatrixAgent || {};
  global.MatrixAgent.exportWorkbook = exportWorkbook;
  global.MatrixAgent.loadTemplate = loadTemplate;
})(window);
