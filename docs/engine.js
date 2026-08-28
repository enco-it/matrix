/* global XLSX, pdfjsLib */
(function (global) {
  const PAGE_FLOORS_K2 = {
    1: [1], 2: [2], 3: [3, 6, 9, 12], 4: [4, 7, 10], 5: [5, 8, 11],
    6: [13, 15], 7: [14, 16], 8: [17, 18, 19, 20, 21, 22, 23, 24],
    9: [25], 10: [26], 11: [27], 12: [28], 13: [28], 14: [29, 30],
    15: [31], 16: [32], 17: [],
  };

  const HEADERS7 = [
    "", "ГП", "Секция", "Этаж", "Номер квартиры", "Номер на этаже", "Кол-во квартир на этаже",
    "Этажность", "Кол-во комнат", "Жилая S квартиры",
    "S без учета балконов/лоджий/террас (без зимнего сада)",
    "S без учета балконов/лоджий/террас (с зимним садом)",
    "Общая S балконов/лоджий/террас (без кф)",
    "S с учетом балконов/лоджий/террас (без кф)",
    "S балконов/лоджий/террас (с кф)",
    "Продаваемая S с учетом балконов/лоджий/террас (с кф)",
  ];

  const STEPS = [
    { id: "accept", label: "Приём исходного файла" },
    { id: "format", label: "Определение формата" },
    { id: "extract", label: "Извлечение штампов квартир" },
    { id: "compose", label: "Сборка строк матрицы" },
    { id: "book", label: "Запись листа и формул" },
    { id: "metrics", label: "Подсчёт метрик" },
  ];

  function nfloat(s) {
    return parseFloat(String(s).replace(",", "."));
  }

  function soften(text) {
    const out = [];
    let acc = "";
    String(text || "").replace(/\r/g, "").split("\n").forEach((line) => {
      const s = line.trim();
      if (!s) {
        if (acc) { out.push(acc); acc = ""; }
        return;
      }
      if (s.length === 1) acc += s;
      else {
        if (acc) { out.push(acc); acc = ""; }
        out.push(s);
      }
    });
    if (acc) out.push(acc);
    return out.join(" ");
  }

  function stampRecord(rooms, size, a, b, c, d, num, pageNo) {
    const areas = [nfloat(a), nfloat(b), nfloat(c), nfloat(d)].sort((x, y) => x - y);
    return {
      rooms: parseInt(rooms, 10),
      size: String(size).toUpperCase().replace("М", "M"),
      live: areas[0],
      s_wo: areas[1],
      s_raw: areas[2],
      s_kf: areas[3],
      num: parseInt(num, 10),
      page: pageNo,
    };
  }

  function parseStampsFromText(text, pageNo, fileName) {
    const out = [];
    const spaceRe = /([1-6])\s*([SMLМм])\s+(\d+[.,]\d+)\s+(\d+[.,]\d+)\s+(\d+[.,]\d+)\s+(\d+[.,]\d+)\s*(?:[№Nn#]|N[oо]\.?)\s*(\d{1,4})/gi;
    const glueRe = /([1-6])([SMLМм])(\d+[.,]\d+)(\d+[.,]\d+)(\d+[.,]\d+)(\d+[.,]\d+)(?:[№Nn#])(\d{1,4})/gi;
    const revRe = /(?:[№Nn#]|N[oо]\.?)\s*(\d{1,4})\s*([1-6])\s*([SMLМм])\s+(\d+[.,]\d+)\s+(\d+[.,]\d+)\s+(\d+[.,]\d+)\s+(\d+[.,]\d+)/gi;
    const blobs = [
      String(text || ""),
      String(text || "").replace(/\s+/g, " "),
      String(text || "").replace(/\n/g, ""),
      soften(text),
    ];
    blobs.forEach((blob) => {
      let m;
      spaceRe.lastIndex = 0;
      while ((m = spaceRe.exec(blob))) out.push(stampRecord(m[1], m[2], m[3], m[4], m[5], m[6], m[7], pageNo));
      glueRe.lastIndex = 0;
      while ((m = glueRe.exec(blob.replace(/\s+/g, "")))) out.push(stampRecord(m[1], m[2], m[3], m[4], m[5], m[6], m[7], pageNo));
      revRe.lastIndex = 0;
      while ((m = revRe.exec(blob))) out.push(stampRecord(m[2], m[3], m[4], m[5], m[6], m[7], m[1], pageNo));
    });
    return attachFloorSection(out, text, pageNo, fileName);
  }

  function mergeItems(items) {
    const pts = (items || [])
      .filter((it) => it && it.str)
      .map((it) => {
        const tr = it.transform || [1, 0, 0, 1, 0, 0];
        const h = Math.abs(it.height || tr[3] || 8);
        return { t: it.str, x: tr[4], y: tr[5], w: it.width || h * 0.5 * it.str.length, h };
      })
      .sort((a, b) => b.y - a.y || a.x - b.x);
    const toks = [];
    pts.forEach((p) => {
      const last = toks[toks.length - 1];
      const same = last && Math.abs(last.y - p.y) <= Math.max(2.2, last.h * 0.45);
      const gap = last ? p.x - (last.x + last.w) : 99;
      if (same && gap < Math.max(1.8, last.h * 0.4) && gap > -2) {
        last.t += p.t;
        last.w = p.x + p.w - last.x;
        return;
      }
      toks.push({ t: p.t, x: p.x, y: p.y, w: p.w, h: p.h });
    });
    return toks;
  }

  function parseStampsFromItems(items, pageNo, fileName) {
    const toks = mergeItems(items);
    const out = [];
    parseStampsFromText(toks.map((t) => t.t).join(" "), pageNo, fileName).forEach((s) => out.push(s));
    parseStampsFromText(toks.map((t) => t.t).join("\n"), pageNo, fileName).forEach((s) => out.push(s));
    toks.forEach((tok) => {
      parseStampsFromText(tok.t, pageNo, fileName).forEach((s) => out.push(s));
    });
    const typeRe = /^[1-6][SMLМм]$/i;
    const areaRe = /^\d+[.,]\d+$/;
    const types = toks.filter((t) => typeRe.test(t.t.trim()));
    const areas = toks.filter((t) => areaRe.test(t.t) && nfloat(t.t) >= 3 && nfloat(t.t) <= 450);
    const nums = [];
    toks.forEach((t) => {
      const m = t.t.replace(/\s+/g, "").match(/^[№Nn#]?(\d{1,4})$/);
      if (!m) return;
      if (m[1].length >= 4 && t.t.indexOf("№") < 0) return;
      nums.push({ tok: t, num: parseInt(m[1], 10) });
    });
    types.forEach((tp) => {
      const col = areas
        .filter((a) => Math.abs(a.x - tp.x) < 36 && Math.abs(a.y - tp.y) < 80)
        .sort((a, b) => Math.abs(a.y - tp.y) - Math.abs(b.y - tp.y));
      if (col.length < 4) return;
      const four = col.slice(0, 4);
      const cy = four.reduce((s, a) => s + a.y, 0) / 4;
      const ncand = nums
        .filter((n) => Math.abs(n.tok.x - tp.x) < 42 && Math.abs(n.tok.y - cy) < 90)
        .sort((a, b) => Math.hypot(a.tok.x - tp.x, a.tok.y - cy) - Math.hypot(b.tok.x - tp.x, b.tok.y - cy));
      if (!ncand.length) return;
      const tm = tp.t.trim().match(/^([1-6])([SMLМм])$/i);
      out.push(stampRecord(tm[1], tm[2], four[0].t, four[1].t, four[2].t, four[3].t, ncand[0].num, pageNo));
    });
    const pageText = toks.map((t) => t.t).join("\n");
    return attachFloorSection(out, pageText, pageNo, fileName);
  }

  function attachFloorSection(stamps, text, pageNo, fileName) {
    const isK2 = /[кkКK]-?2|айвазов/i.test(fileName || "");
    const floors = isK2 ? (PAGE_FLOORS_K2[pageNo] || []) : guessFloors(text || "");
    stamps.forEach((s) => {
      s.floor = floors.length === 1 ? floors[0] : floors[0] || null;
      s.floor_candidates = floors;
      s.floor_uncertain = floors.length > 1;
      s.section = guessSection(text || "");
    });
    return stamps;
  }

  function guessFloors(text) {
    let m = text.match(/План\s+(\d+)(?:\s*[-–]\s*(\d+))?\s*этаж/i);
    if (!m) m = text.match(/(\d{1,2})\s*[-–]\s*(\d{1,2})\s*этаж/i);
    if (!m) {
      const one = text.match(/(?:этаж|эт\.?)\s*(\d{1,2})\b/i) || text.match(/\b(\d{1,2})\s*этаж/i);
      return one ? [parseInt(one[1], 10)] : [];
    }
    const a = parseInt(m[1], 10);
    const b = m[2] ? parseInt(m[2], 10) : a;
    const r = [];
    for (let i = a; i <= b; i++) r.push(i);
    return r;
  }

  function guessSection(text) {
    const m = text.match(/\b(\d{1,2}\.[1-9])\b/);
    return m ? parseFloat(m[1]) : 1;
  }

  function guessGp(fileName) {
    const m = String(fileName || "").match(/(\d{3,4})/);
    return m ? parseInt(m[1], 10) : 1506;
  }

  function uniqueByNum(stamps) {
    const map = new Map();
    for (const s of stamps) {
      if (!map.has(s.num)) map.set(s.num, s);
    }
    return [...map.values()].sort((a, b) => (a.floor || 0) - (b.floor || 0) || a.num - b.num);
  }

  function composeRow(apt, seq, onFloor, floorCount, gp) {
    const notes = [];
    const nRooms = Math.max(1, apt.rooms || 1);
    const studio = nRooms === 1 && apt.size === "S";
    const live = apt.live;
    const wo = apt.s_wo;
    const remain = Math.round((wo - live) * 100) / 100;
    const kitchen = remain > 0 ? Math.round(remain * 50) / 100 : null;
    const hall = remain > 0 ? Math.round((remain - (kitchen || 0)) * 100) / 100 : null;
    if (apt.floor_uncertain) notes.push("этаж по первому листу типовой серии");
    const liveParts = [];
    if (nRooms === 1) liveParts.push(live);
    else {
      const part = Math.round((live / nRooms) * 100) / 100;
      for (let i = 0; i < nRooms - 1; i++) liveParts.push(part);
      liveParts.push(Math.round((live - part * (nRooms - 1)) * 100) / 100);
    }
    const row = {
      A: seq,
      B: gp,
      C: apt.section,
      D: apt.floor,
      E: apt.num,
      F: onFloor,
      G: floorCount,
      AF: studio ? "студия" : "Кухня-гостиная",
      AG: kitchen && kitchen >= 4 ? kitchen : undefined,
      T: hall && hall > 0.2 ? hall : undefined,
      BI: remain >= 6 ? 4 : undefined,
      notes,
    };
    if (studio) {
      row.AJ = "Студия";
      row.AK = "Студия";
      row.AL = liveParts[0];
    } else {
      const p = ["AJ", "AO", "AT", "AY", "BD"];
      const t = ["AK", "AP", "AU", "AZ", "BE"];
      const s = ["AL", "AQ", "AV", "BA", "BF"];
      liveParts.slice(0, 5).forEach((area, i) => {
        row[p[i]] = "Жилая комната";
        row[t[i]] = "Комната стандарт";
        row[s[i]] = area;
      });
    }
    const summer = Math.round((apt.s_raw - apt.s_wo) * 100) / 100;
    const summerKf = Math.round((apt.s_kf - apt.s_wo) * 100) / 100;
    if (summer > 0.2) {
      row.CE = summer;
      row.CF = summerKf > 0.05 ? summerKf : summer;
      row.CG = "стандартное";
      row.CH = "комната";
    }
    return row;
  }

  function metrics(apts, rows, fileName, sourceKind) {
    const floors = {};
    const kitchens = {};
    let review = 0;
    let studio = 0;
    for (const a of apts) {
      const f = a.floor == null ? "—" : String(a.floor);
      floors[f] = (floors[f] || 0) + 1;
    }
    for (const r of rows) {
      kitchens[r.AF] = (kitchens[r.AF] || 0) + 1;
      if (r.AF === "студия") studio += 1;
      review += (r.notes || []).length;
    }
    return {
      fileName,
      sourceKind,
      apartments: apts.length,
      pagesHint: apts.reduce((m, a) => Math.max(m, a.page || 0), 0),
      studio,
      kitchenLiving: kitchens["Кухня-гостиная"] || 0,
      reviewFlags: review,
      floors,
      kitchens,
      minFloor: Math.min(...apts.map((a) => a.floor).filter((x) => x != null)),
      maxFloor: Math.max(...apts.map((a) => a.floor).filter((x) => x != null)),
    };
  }

  function colLetter(n) {
    let s = "";
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function formulas(row, last) {
    return {
      H: `=_xlfn.MAXIFS($D$8:$D$${last},$C$8:$C$${last},C${row},$B$8:$B$${last},B${row})`,
      I: `=IF(AND(AF${row}="студия",AK${row}<>"",AL${row}>0),"ст",IF(AL${row}>0,COUNTA(AL${row},AQ${row},AV${row},BA${row},BF${row}),""))`,
      J: `=SUM(AL${row},AQ${row},AV${row},BA${row},BF${row})`,
      K: `=SUM(Q${row}:BM${row})`,
      L: `=K${row}+CB${row}`,
      M: `=SUM(BP${row},BT${row},BX${row},CE${row},CI${row},CM${row},CP${row},CS${row})`,
      N: `=L${row}+M${row}`,
      O: `=SUM(BQ${row},BU${row},BY${row},CF${row},CJ${row},CN${row},CQ${row},CT${row})`,
      P: `=L${row}+O${row}`,
    };
  }

  function letterToIndex(letter) {
    let n = 0;
    for (let i = 0; i < letter.length; i++) n = n * 26 + (letter.charCodeAt(i) - 64);
    return n;
  }

  function buildWorkbook(rows, meta) {
    const last = 7 + rows.length;
    const aoa = [];
    aoa[0] = ["Матрица квартирографии · автозаполнение агентом"];
    aoa[1] = ["Источник", meta.fileName, "Формат", meta.sourceKind];
    aoa[6] = HEADERS7.slice();
    while (aoa[6].length < 90) aoa[6].push("");
    aoa[6][19] = "Прихожая";
    aoa[6][31] = "Тип кухни";
    aoa[6][32] = "S кухни";
    aoa[6][35] = "Тип помещения 1";
    aoa[6][36] = "Тип комнаты 1";
    aoa[6][37] = "Комната 1";
    aoa[6][60] = "С/у1";
    aoa[6][82] = "Лоджия";
    aoa[6][83] = "Лоджия с кф";

    const keys = ["A","B","C","D","E","F","G","T","AF","AG","AJ","AK","AL","AO","AP","AQ","AT","AU","AV","BI","CE","CF","CG","CH"];
    rows.forEach((row, i) => {
      const ridx = 8 + i;
      const line = new Array(90);
      const f = formulas(ridx, last);
      Object.entries(f).forEach(([k, v]) => { line[letterToIndex(k) - 1] = v; });
      keys.forEach((k) => {
        if (row[k] !== undefined) line[letterToIndex(k) - 1] = row[k];
      });
      aoa[ridx - 1] = line;
    });

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 6 }, { wch: 8 }, { wch: 10 }, { wch: 8 }, { wch: 14 }, { wch: 12 }, { wch: 14 }];
    const review = [["Лот", "Этаж", "Секция", "Причина"]];
    rows.forEach((row) => {
      (row.notes || []).forEach((n) => review.push([row.E, row.D, row.C, n]));
    });
    review.push(["", "", "", "Обработка в браузере. RVT не читается без IFC/PDF."]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Квартирография (жилое)");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(review), "Требует проверки");
    return wb;
  }

  function fromDemoJson(items, fileName) {
    const apts = items.map((it) => ({
      num: it.num,
      floor: it.floor,
      section: it.sec,
      rooms: it.studio ? 1 : 2,
      size: it.studio ? "S" : "M",
      live: it.r1 || 20,
      s_wo: (it.r1 || 20) + (it.sk || 10) + (it.hall || 8),
      s_raw: (it.r1 || 20) + (it.sk || 10) + (it.hall || 8),
      s_kf: (it.r1 || 20) + (it.sk || 10) + (it.hall || 8),
      page: 0,
      kitchenPreset: it.kitchen,
    }));
    const byFloor = {};
    apts.forEach((a) => { byFloor[a.floor] = (byFloor[a.floor] || 0) + 1; });
    const on = {};
    const seen = {};
    const rows = apts.map((a, i) => {
      seen[a.floor] = (seen[a.floor] || 0) + 1;
      const row = composeRow(a, i + 1, seen[a.floor], byFloor[a.floor], 1506);
      if (items[i].kitchen) {
        row.AF = items[i].kitchen;
        if (items[i].kitchen === "студия") {
          row.AJ = "Студия";
          row.AK = "Студия";
        }
      }
      return row;
    });
    return { apts, rows, meta: metrics(apts, rows, fileName, "demo") };
  }

  async function extractPdf(file, onPage) {
    const buf = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({
      data: new Uint8Array(buf),
      cMapUrl: "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/cmaps/",
      cMapPacked: true,
      standardFontDataUrl: "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/standard_fonts/",
    }).promise;
    const stamps = [];
    let chars = 0;
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const tc = await page.getTextContent();
      const items = tc.items || [];
      chars += items.reduce((n, it) => n + (it.str ? it.str.length : 0), 0);
      const text = items.map((it) => it.str || "").join("\n");
      stamps.push(...parseStampsFromText(text, i, file.name));
      stamps.push(...parseStampsFromItems(items, i, file.name));
      if (onPage) onPage(i, pdf.numPages);
    }
    return { stamps: uniqueByNum(stamps), pages: pdf.numPages, chars };
  }

  function composeAll(apts, gp) {
    const byFloor = {};
    apts.forEach((a) => { byFloor[a.floor] = (byFloor[a.floor] || 0) + 1; });
    const seen = {};
    return apts.map((a, i) => {
      seen[a.floor] = (seen[a.floor] || 0) + 1;
      return composeRow(a, i + 1, seen[a.floor], byFloor[a.floor] || 1, gp);
    });
  }

  function detectKind(file) {
    const n = (file.name || "").toLowerCase();
    if (n.endsWith(".pdf")) return "pdf";
    if (n.endsWith(".rvt")) return "rvt";
    if (n.endsWith(".dwg")) return "dwg";
    if (n.endsWith(".xlsx") || n.endsWith(".xlsm")) return "xlsx";
    return "unknown";
  }

  global.MatrixAgent = {
    STEPS,
    parseStampsFromText,
    parseStampsFromItems,
    guessGp,
    uniqueByNum,
    composeAll,
    metrics,
    buildWorkbook,
    fromDemoJson,
    extractPdf,
    detectKind,
    colLetter,
  };
})(window);
