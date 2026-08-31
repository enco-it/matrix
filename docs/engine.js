/* global XLSX, pdfjsLib */
(function (global) {
  const PAGE_FLOORS_K2 = {
    1: [1], 2: [2], 3: [3, 6, 9, 12], 4: [4, 7, 10], 5: [5, 8, 11],
    6: [13, 15], 7: [14, 16], 8: [17, 18, 19, 20, 21, 22, 23, 24],
    9: [25], 10: [26], 11: [27], 12: [28], 13: [28], 14: [29, 30],
    15: [31], 16: [32], 17: [],
  };

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
    const dottedRe = /(\d+[.,]\d+)\s+(\d+[.,]\d+)\s+(\d+[.,]\d+)\s+(\d{2}\.\d\.\d{1,2})\s+(\d+[.,]\d+)/g;
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
      const glued = blob.replace(/\s+/g, "");
      if (glued.length < 800) {
        while ((m = glueRe.exec(glued))) out.push(stampRecord(m[1], m[2], m[3], m[4], m[5], m[6], m[7], pageNo));
      }
      revRe.lastIndex = 0;
      while ((m = revRe.exec(blob))) out.push(stampRecord(m[2], m[3], m[4], m[5], m[6], m[7], m[1], pageNo));
      dottedRe.lastIndex = 0;
      while ((m = dottedRe.exec(blob))) {
        const rec = fromDottedId(m[4], [nfloat(m[1]), nfloat(m[2]), nfloat(m[3]), nfloat(m[5])], pageNo);
        if (rec) out.push(rec);
      }
    });
    return out;
  }

  function fromDottedId(id, areas, pageNo) {
    const m = String(id).match(/^(\d{2})\.(\d)\.(\d{1,2})$/);
    if (!m) return null;
    const sorted = areas.filter((x) => Number.isFinite(x)).sort((a, b) => a - b);
    if (sorted.length < 3 || sorted[0] < 3 || sorted[sorted.length - 1] > 200) return null;
    const live = sorted[0];
    const rec = stampRecord(
      live < 14 ? 1 : 2,
      live < 14 ? "S" : "M",
      sorted[0],
      sorted[1],
      sorted[2],
      sorted[3] != null ? sorted[3] : sorted[2],
      m[3],
      pageNo
    );
    rec.gp = parseInt(m[1][0], 10);
    rec.section = parseFloat(`${m[1][0]}.${m[2]}`);
    rec.code = id;
    return rec;
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

  function toksToText(toks) {
    let y = null;
    const lines = [];
    let cur = [];
    (toks || []).forEach((t) => {
      if (y == null || Math.abs(t.y - y) > Math.max(4, t.h * 0.65)) {
        if (cur.length) lines.push(cur.join(" "));
        cur = [t.t];
        y = t.y;
      } else cur.push(t.t);
    });
    if (cur.length) lines.push(cur.join(" "));
    return lines.join("\n");
  }

  function dist(a, b) {
    return Math.hypot((a.x || 0) - (b.x || 0), (a.y || 0) - (b.y || 0));
  }

  function guessFloorsGeo(toks) {
    const plans = (toks || []).filter((t) => /^План$/i.test(String(t.t).trim()));
    const floorWords = (toks || []).filter((t) => /^этаж/i.test(String(t.t).trim()) || /^подвал/i.test(String(t.t).trim()));
    const ranges = (toks || []).filter((t) => /^\d{1,2}[-–]\d{1,2}$/.test(String(t.t).trim()));
    const ones = (toks || []).filter((t) => {
      const s = String(t.t).trim();
      return /^\d{1,2}$/.test(s) && +s >= 1 && +s <= 40;
    });
    const titles = [];
    plans.forEach((p) => {
      const fw = floorWords.filter((f) => dist(f, p) < 160);
      if (!fw.length) return;
      const rng = ranges.filter((r) => dist(r, p) < 160).sort((a, b) => dist(a, p) - dist(b, p));
      const one = ones.filter((o) => dist(o, p) < 160).sort((a, b) => dist(a, p) - dist(b, p));
      if (rng.length) titles.push(rng[0].t.replace("–", "-"));
      else if (/подвал/i.test(fw[0].t)) titles.push("подвал");
      else if (one.length) titles.push(one[0].t);
    });
    const uniq = [...new Set(titles.map((x) => String(x).replace(/\s+/g, "").toLowerCase()))];
    if (uniq.length >= 3) return [];
    if (!uniq.length) return null;
    const t = uniq[0];
    const m = String(t).match(/^(\d{1,2})[-–](\d{1,2})$/);
    if (m) {
      const a = parseInt(m[1], 10);
      const b = parseInt(m[2], 10);
      if (a >= 1 && a <= 40 && b >= a && b <= 40) {
        const r = [];
        for (let i = a; i <= b; i++) r.push(i);
        return r;
      }
    }
    if (t === "подвал") return [];
    const n = parseInt(t, 10);
    return n >= 1 && n <= 40 ? [n] : null;
  }

  function mergeNumeric(toks) {
    const out = [];
    (toks || []).forEach((p) => {
      const last = out[out.length - 1];
      const numish = /^[\d.,]+$/.test(p.t);
      const lastNum = last && /^[\d.,]+$/.test(last.t);
      const same = last && Math.abs(last.y - p.y) <= Math.max(3, last.h * 0.55);
      const gap = last ? p.x - (last.x + last.w) : 99;
      if (same && lastNum && numish && last.t.length <= 3 && p.t.length <= 3 && gap < 8 && gap > -3) {
        last.t += p.t;
        last.w = p.x + p.w - last.x;
        return;
      }
      out.push({ t: p.t, x: p.x, y: p.y, w: p.w, h: p.h });
    });
    return out;
  }

  function itemsToToks(items) {
    return (items || [])
      .filter((it) => it && it.str)
      .map((it) => {
        const tr = it.transform || [1, 0, 0, 1, 0, 0];
        const h = Math.abs(it.height || tr[3] || 8);
        return { t: it.str, x: tr[4], y: tr[5], w: it.width || h * 0.5 * it.str.length, h };
      });
  }

  function collectDotted(toks, pageNo, out) {
    const dottedId = /^\d{2}\.\d\.\d{1,2}$/;
    const areaTok = /^\d{1,3}[.,]\d{2}$/;
    const ids = (toks || []).filter((t) => dottedId.test(String(t.t).trim()));
    const areaToks = (toks || []).filter((t) => areaTok.test(t.t) && nfloat(t.t) >= 3 && nfloat(t.t) <= 200);
    const seenId = {};
    ids.forEach((tp) => {
      const id = tp.t.trim();
      if (seenId[id]) return;
      const col = areaToks
        .filter((a) => Math.abs(a.x - tp.x) < 30 && Math.abs(a.y - tp.y) < 60)
        .sort((a, b) => a.y - b.y);
      if (col.length < 3) return;
      const rec = fromDottedId(id, col.slice(0, 4).map((a) => nfloat(a.t)), pageNo);
      if (!rec) return;
      seenId[id] = true;
      out.push(rec);
    });
  }

  function parseStampsFromItems(items, pageNo, fileName) {
    const rawToks = itemsToToks(items);
    const merged = mergeItems(items);
    const mergedNum = mergeNumeric(merged);
    const out = [];
    const pageText = toksToText(rawToks) + "\n" + toksToText(merged);
    collectDotted(rawToks, pageNo, out);
    collectDotted(mergedNum, pageNo, out);
    parseStampsFromText(toksToText(rawToks), pageNo, fileName).forEach((s) => out.push(s));
    parseStampsFromText(toksToText(merged), pageNo, fileName).forEach((s) => out.push(s));
    const typeRe = /^[1-6][SMLМм]$/i;
    const areaRe = /^\d+[.,]\d+$/;
    const types = merged.filter((t) => typeRe.test(t.t.trim()));
    const areas = merged.filter((t) => areaRe.test(t.t) && nfloat(t.t) >= 3 && nfloat(t.t) <= 450);
    const nums = [];
    merged.forEach((t) => {
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
    return attachFloorSection(out, pageText, pageNo, fileName, rawToks);
  }

  function attachFloorSection(stamps, text, pageNo, fileName, toks) {
    const isK2 = /[кkКK]-?2|айвазов/i.test(fileName || "");
    const dottedCodes = new Set((stamps || []).filter((s) => s.code).map((s) => s.code));
    if (!isK2 && dottedCodes.size >= 12) {
      stamps = (stamps || []).filter((s) => !s.code);
    }
    let floors = isK2 ? (PAGE_FLOORS_K2[pageNo] || []) : guessFloors(text || "");
    if (!isK2) {
      const geo = guessFloorsGeo(toks || []);
      if (geo && geo.length === 0) floors = [];
      else if (geo && geo.length) floors = geo;
    }
    const expand = !isK2 && floors.length > 1;
    const out = [];
    stamps.forEach((s) => {
      const section = s.section != null ? s.section : guessSection(text || "");
      if (!isK2 && !floors.length) {
        if (s.code) return;
        s.floor = null;
        s.floor_candidates = [];
        s.floor_uncertain = true;
        s.section = section;
        out.push(s);
        return;
      }
      if (expand) {
        floors.forEach((fl) => {
          out.push(Object.assign({}, s, {
            floor: fl,
            floor_candidates: floors,
            floor_uncertain: false,
            section,
          }));
        });
        return;
      }
      s.floor = floors.length === 1 ? floors[0] : floors[0] || null;
      s.floor_candidates = floors;
      s.floor_uncertain = floors.length > 1;
      s.section = section;
      out.push(s);
    });
    return out;
  }

  function guessFloors(text) {
    const titles = [];
    const titleRe = /План\s+(подвала|\d+(?:\s*[-–]\s*\d+)?\s*этаж[аей]*)/gi;
    let tm;
    while ((tm = titleRe.exec(text || ""))) titles.push(tm[1].replace(/\s+/g, " ").toLowerCase());
    if ([...new Set(titles)].length >= 3) return [];
    let m = (text || "").match(/План\s+(\d+)\s*[-–]\s*(\d+)\s*этаж/i);
    if (m) {
      const a = parseInt(m[1], 10);
      const b = parseInt(m[2], 10);
      if (a >= 1 && a <= 40 && b >= a && b <= 40) {
        const r = [];
        for (let i = a; i <= b; i++) r.push(i);
        return r;
      }
    }
    m = (text || "").match(/План\s+(\d+)\s*этажа/i);
    if (m) {
      const n = parseInt(m[1], 10);
      return n >= 1 && n <= 40 ? [n] : [];
    }
    m = (text || "").match(/План\s+(\d+)(?:\s*[-–]\s*(\d+))?\s*этаж/i);
    if (!m) m = (text || "").match(/(\d{1,2})\s*[-–]\s*(\d{1,2})\s*этаж/i);
    if (!m) {
      const one = (text || "").match(/(?:этаж|эт\.?)\s*(\d{1,2})\b/i) || (text || "").match(/\b(\d{1,2})\s*этаж/i);
      return one ? [parseInt(one[1], 10)] : [];
    }
    const a = parseInt(m[1], 10);
    const b = m[2] ? parseInt(m[2], 10) : a;
    const r = [];
    for (let i = a; i <= b; i++) r.push(i);
    return r;
  }

  function guessSection(text) {
    const titled = (text || "").match(/Секци[яи]\s*(\d{1,2}\.\d)/i);
    if (titled) return parseFloat(titled[1]);
    const m = (text || "").match(/\b(\d(?:\.\d)?)\s*секц/i);
    return m ? parseFloat(m[1]) : 1;
  }

  function guessGp(fileName, stamps) {
    if (stamps && stamps.length) {
      const hit = stamps.find((s) => s.gp != null);
      if (hit) return hit.gp;
    }
    const named = String(fileName || "").match(/(?:ГП|GP)[-_]?(\d{1,2})\b/i);
    if (named) return parseInt(named[1], 10);
    const dash = String(fileName || "").match(/(\d{4,5})-(\d{1,2})/);
    if (dash) return parseInt(dash[2], 10);
    const m = String(fileName || "").match(/(\d{3,5})/);
    return m ? parseInt(m[1], 10) : 1506;
  }

  function uniqueByNum(stamps) {
    return uniqueByApt(stamps);
  }

  function uniqueByApt(stamps) {
    const map = new Map();
    for (const s of stamps) {
      const key = `${s.floor == null ? "" : s.floor}|${s.section}|${s.num}`;
      if (!map.has(key)) map.set(key, s);
    }
    return [...map.values()].sort(
      (a, b) => (a.floor || 0) - (b.floor || 0) || (a.section || 0) - (b.section || 0) || a.num - b.num
    );
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

  function letterToIndex(letter) {
    let n = 0;
    for (let i = 0; i < letter.length; i++) n = n * 26 + (letter.charCodeAt(i) - 64);
    return n;
  }

  function retargetFormula(formula, row, last) {
    let s = String(formula);
    s = s.replace(/\$([A-Za-z]{1,3})\$8:\$\1\$\d+/g, (_, col) => `$${col}$8:$${col}$${last}`);
    s = s.replace(/\$8:\$\d+/g, `$8:$${last}`);
    s = s.replace(/(\$?[A-Za-z]{1,3})(\$?)8(?!\d)/g, (m, col, abs) => (abs ? m : col + row));
    return s;
  }

  let skeleton = null;
  async function loadSkeleton() {
    if (skeleton) return skeleton;
    const res = await fetch("template-skeleton.json");
    if (!res.ok) throw new Error("Не загружен шаблон матрицы ГП-10");
    skeleton = await res.json();
    return skeleton;
  }

  function buildWorkbook(rows, meta, skel) {
    if (!skel) throw new Error("Нет шаблона матрицы");
    const last = 7 + rows.length;
    const maxCol = skel.maxCol || 349;
    const aoa = [];
    (skel.headers || []).forEach((row) => {
      const line = (row || []).slice();
      while (line.length < maxCol) line.push(null);
      aoa.push(line);
    });
    while (aoa.length < 7) aoa.push(new Array(maxCol).fill(null));
    if (!aoa[0]) aoa[0] = new Array(maxCol).fill(null);
    aoa[0][0] = `Матрица квартирографии · ${meta.fileName || ""} · ${rows.length} кв.`;

    const formulas = skel.formulas || [];
    rows.forEach((row, i) => {
      const ridx = 8 + i;
      const line = new Array(maxCol).fill(null);
      formulas.forEach((f, c) => {
        if (f) line[c] = retargetFormula(f, ridx, last);
      });
      Object.keys(row).forEach((k) => {
        if (k === "notes" || row[k] === undefined) return;
        const idx = letterToIndex(k);
        if (idx > 0) line[idx - 1] = row[k];
      });
      aoa[ridx - 1] = line;
    });

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!ref"] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: Math.max(6, last - 1), c: maxCol - 1 },
    });
    if (skel.merges && skel.merges.length && XLSX.utils.decode_range) {
      ws["!merges"] = skel.merges.map((ref) => XLSX.utils.decode_range(ref));
    }
    if (skel.colWidths) {
      ws["!cols"] = skel.colWidths.map((w) => (w ? { wch: w } : {}));
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, skel.living || "Квартирография (жилое)");
    Object.keys(skel.sheets || {}).forEach((name) => {
      const grid = skel.sheets[name] || [];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(grid), name.slice(0, 31));
    });

    const review = [["Лот", "Этаж", "Секция", "Причина"]];
    rows.forEach((row) => {
      (row.notes || []).forEach((n) => review.push([row.E, row.D, row.C, n]));
    });
    review.push(["", "", "", "Лист и формулы — как в шаблоне ГП-10. VBA макросов в выгрузке нет."]);
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
      stamps.push(...parseStampsFromItems(items, i, file.name));
      if (onPage) onPage(i, pdf.numPages);
    }
    return { stamps: uniqueByApt(stamps), pages: pdf.numPages, chars };
  }

  function composeAll(apts, gp) {
    const byFloor = {};
    apts.forEach((a) => { byFloor[a.floor] = (byFloor[a.floor] || 0) + 1; });
    const seen = {};
    return apts.map((a, i) => {
      seen[a.floor] = (seen[a.floor] || 0) + 1;
      return composeRow(a, i + 1, seen[a.floor], byFloor[a.floor] || 1, a.gp || gp);
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
    uniqueByApt,
    composeAll,
    metrics,
    loadSkeleton,
    retargetFormula,
    buildWorkbook,
    fromDemoJson,
    extractPdf,
    detectKind,
    colLetter,
  };
})(window);
