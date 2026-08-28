/* global MatrixAgent, pdfjsLib, XLSX */
(function () {
  const fileEl = document.getElementById("file");
  const drop = document.getElementById("drop");
  const fileLabel = document.getElementById("fileLabel");
  const btnRun = document.getElementById("btnRun");
  const btnDemo = document.getElementById("btnDemo");
  const btnDl = document.getElementById("btnDl");
  const stepsEl = document.getElementById("steps");
  const bar = document.getElementById("bar");
  const status = document.getElementById("status");
  const metricsBlock = document.getElementById("metricsBlock");
  const statsEl = document.getElementById("stats");
  const fmtHint = document.getElementById("fmtHint");

  let file = null;
  let fmt = "xlsx";
  let lastBook = null;
  let lastName = "matrix";
  let demoItems = null;
  let useDemo = false;

  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

  function renderSteps(state) {
    stepsEl.innerHTML = MatrixAgent.STEPS.map((s) => {
      const st = state[s.id] || "idle";
      const mark = st === "ok" ? "ok" : st === "run" ? "run" : st === "err" ? "err" : "";
      const right = st === "ok" ? "готово" : st === "run" ? "…" : st === "err" ? "ошибка" : "";
      return `<li><i class="dot ${mark}"></i><span>${s.label}</span><span class="hint" style="margin:0">${right}</span></li>`;
    }).join("");
  }

  const stepState = {};
  MatrixAgent.STEPS.forEach((s) => { stepState[s.id] = "idle"; });
  renderSteps(stepState);

  function setFmt(next) {
    fmt = next;
    document.querySelectorAll(".seg-btn").forEach((b) => {
      b.classList.toggle("on", b.getAttribute("data-fmt") === fmt);
    });
    fmtHint.textContent =
      fmt === "xlsm"
        ? "xlsm — контейнер Excel с макросами. VBA образца не копируется (в браузере её нет); лист и формулы записываются."
        : "xlsx — Excel без макросов. Лист и формулы как в образце ГП-10.";
  }

  document.getElementById("fmtXlsx").onclick = () => setFmt("xlsx");
  document.getElementById("fmtXlsm").onclick = () => setFmt("xlsm");

  function setFile(f, demo) {
    file = f;
    useDemo = !!demo;
    fileLabel.textContent = f ? f.name : "Выберите или перетащите файл";
    btnRun.disabled = !f && !useDemo;
    status.textContent = useDemo ? "Демо К-2 загружено — можно запускать" : f ? "Файл принят" : "Ожидание файла";
  }

  fileEl.addEventListener("change", () => {
    if (fileEl.files[0]) setFile(fileEl.files[0], false);
  });
  ["dragenter", "dragover"].forEach((ev) => {
    drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add("drag"); });
  });
  ["dragleave", "drop"].forEach((ev) => {
    drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove("drag"); });
  });
  drop.addEventListener("drop", (e) => {
    const f = e.dataTransfer.files[0];
    if (f) setFile(f, false);
  });

  function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

  async function mark(id, st) {
    stepState[id] = st;
    const done = MatrixAgent.STEPS.filter((s) => stepState[s.id] === "ok").length;
    bar.style.width = `${Math.round((done / MatrixAgent.STEPS.length) * 100)}%`;
    renderSteps(stepState);
    await sleep(st === "run" ? 180 : 40);
  }

  function bars(el, obj) {
    const max = Math.max(1, ...Object.values(obj));
    el.innerHTML = Object.entries(obj)
      .sort((a, b) => (Number(a[0]) || 0) - (Number(b[0]) || 0))
      .map(([k, v]) => {
        const w = Math.max(4, Math.round((v / max) * 100));
        return `<div class="hbar"><b>${k}</b><i><em style="width:${w}%"></em></i><span>${v}</span></div>`;
      })
      .join("");
  }

  function showMetrics(meta, rows) {
    metricsBlock.hidden = false;
    statsEl.innerHTML = [
      [meta.apartments, "квартир"],
      [meta.studio, "студий"],
      [meta.kitchenLiving, "кухня-гостиная"],
      [meta.reviewFlags, "пометок к проверке"],
      [`${meta.minFloor}–${meta.maxFloor}`, "этажи"],
      [meta.sourceKind, "источник"],
    ]
      .map(([v, l]) => `<div class="stat"><b>${v}</b><span>${l}</span></div>`)
      .join("");
    bars(document.getElementById("floorBars"), meta.floors);
    bars(document.getElementById("kitBars"), meta.kitchens);
    const prev = rows.slice(0, 12);
    document.getElementById("preview").innerHTML =
      "<thead><tr><th>№</th><th>Секция</th><th>Этаж</th><th>Лот</th><th>Кухня</th><th>Жилая S штампа</th></tr></thead>" +
      "<tbody>" +
      prev
        .map((r, i) => {
          const aLive = r.AL || "";
          return `<tr><td>${r.A}</td><td>${r.C}</td><td>${r.D}</td><td>${r.E}</td><td>${r.AF}</td><td>${aLive}</td></tr>`;
        })
        .join("") +
      (rows.length > 12 ? `<tr><td colspan="6">… ещё ${rows.length - 12} строк</td></tr>` : "") +
      "</tbody>";
    btnDl.hidden = false;
    metricsBlock.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  async function run() {
    btnRun.disabled = true;
    btnDl.hidden = true;
    lastBook = null;
    MatrixAgent.STEPS.forEach((s) => { stepState[s.id] = "idle"; });
    renderSteps(stepState);
    bar.style.width = "0";
    try {
      await mark("accept", "run");
      if (!file && !useDemo) throw new Error("Файл не выбран");
      await mark("accept", "ok");

      await mark("format", "run");
      let kind = useDemo ? "demo" : MatrixAgent.detectKind(file);
      if (kind === "rvt") {
        status.textContent = "RVT в браузере не разбирается. Нужен PDF планировок.";
        await mark("format", "err");
        throw new Error("Загрузите PDF планировок этого корпуса (экспорт из Revit).");
      }
      if (kind === "dwg") {
        await mark("format", "err");
        throw new Error("DWG в браузере не читается. Используйте PDF.");
      }
      await mark("format", "ok");

      await mark("extract", "run");
      let apts;
      let rows;
      let meta;
      if (useDemo) {
        if (!demoItems) {
          const res = await fetch("demo-k2.json");
          demoItems = await res.json();
        }
        const packed = MatrixAgent.fromDemoJson(demoItems, "К-2_планировки.pdf (демо прогон)");
        apts = packed.apts;
        rows = packed.rows;
        meta = packed.meta;
        status.textContent = `Демо: ${apts.length} квартир корпуса К-2`;
      } else {
        status.textContent = "Чтение PDF…";
        const extracted = await MatrixAgent.extractPdf(file, (i, n) => {
          status.textContent = `Лист ${i} из ${n}`;
        });
        apts = extracted.stamps;
        if (!apts.length) throw new Error("Штампы квартир на PDF не найдены. Нужен векторный поэтажный план.");
        rows = MatrixAgent.composeAll(apts, 1506);
        meta = MatrixAgent.metrics(apts, rows, file.name, "pdf");
      }
      await mark("extract", "ok");

      await mark("compose", "run");
      await mark("compose", "ok");

      await mark("book", "run");
      lastBook = MatrixAgent.buildWorkbook(rows, meta);
      lastName = `Матрица квартирографии_${(file && file.name ? file.name.replace(/\.[^.]+$/, "") : "К-2")}`;
      await mark("book", "ok");

      await mark("metrics", "run");
      showMetrics(meta, rows);
      status.textContent = `Готово: ${meta.apartments} квартир. Можно скачать ${fmt}.`;
      await mark("metrics", "ok");
      bar.style.width = "100%";
    } catch (err) {
      status.textContent = err.message || String(err);
      const running = MatrixAgent.STEPS.find((s) => stepState[s.id] === "run");
      if (running) await mark(running.id, "err");
    } finally {
      btnRun.disabled = false;
    }
  }

  btnRun.onclick = run;
  btnDemo.onclick = () => {
    setFile({ name: "К-2_планировки.pdf (демо)" }, true);
  };
  btnDl.onclick = () => {
    if (!lastBook) return;
    const bookType = fmt === "xlsm" ? "xlsm" : "xlsx";
    XLSX.writeFile(lastBook, `${lastName}.${bookType}`, { bookType });
  };

  fetch("demo-k2.json")
    .then((r) => r.json())
    .then((j) => { demoItems = j; })
    .catch(() => {});
})();
