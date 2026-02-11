(() => {
  const $ = (id) => document.getElementById(id);
  const toast = (title, msg) => {
    $("toastTitle").textContent = title;
    $("toastMsg").textContent = msg;
    $("toast").classList.add("show");
    clearTimeout(toast._t);
    toast._t = setTimeout(() => $("toast").classList.remove("show"), 3800);
  };

  const pad2 = (n) => String(n).padStart(2, "0");
  const toDateKey = (d) => `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`;
  const toMonthKey = (d) => `${d.getFullYear()}-${pad2(d.getMonth()+1)}`;

  const formatES = (dateKey) => {
    const [y,m,d] = dateKey.split("-");
    return `${d}/${m}/${y}`;
  };

  const monthLabelES = (monthKey) => {
    const [y,m] = monthKey.split("-").map(Number);
    const dt = new Date(y, m-1, 1);
    return dt.toLocaleDateString("es-ES", { month:"long", year:"numeric" });
  };

  function formatDateTimeES(dt){
    if (!(dt instanceof Date) || Number.isNaN(dt.getTime())) return "—";
    const dd = pad2(dt.getDate());
    const mm = pad2(dt.getMonth()+1);
    const yy = dt.getFullYear();
    const hh = pad2(dt.getHours());
    const mi = pad2(dt.getMinutes());
    return `${dd}/${mm}/${yy} ${hh}:${mi}`;
  }

  function escapeHtml(str){
    return String(str).replace(/[&<>"']/g, s => ({
      "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#039;"
    }[s]));
  }

  // -----------------------------
  // Theme Toggle
  // -----------------------------
  const THEME_KEY = "urg_dashboard_theme";
  function applyTheme(theme){
    const t = (theme === "dark") ? "dark" : "light";
    document.documentElement.setAttribute("data-theme", t);
    localStorage.setItem(THEME_KEY, t);
    const sw = $("themeSwitch");
    if (sw) sw.checked = (t === "dark");
    try { refreshDoctorSection(); } catch {}
  }
  (function initTheme(){
    const saved = localStorage.getItem(THEME_KEY);
    applyTheme(saved === "dark" ? "dark" : "light");
  })();
  $("themeSwitch").addEventListener("change", (e)=> applyTheme(e.target.checked ? "dark" : "light"));

  // -----------------------------
  // Parse fechas
  // -----------------------------
  function parseDateAny(v){
    if (v == null) return null;
    const s = String(v).trim();
    if (!s) return null;

    const d0 = new Date(s);
    if (!Number.isNaN(d0.getTime())) return d0;

    const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (m){
      const dd = Number(m[1]), mm = Number(m[2]), yy = Number(m[3]) < 100 ? 2000 + Number(m[3]) : Number(m[3]);
      const HH = Number(m[4]||0), MM = Number(m[5]||0), SS = Number(m[6]||0);
      const d = new Date(yy, mm-1, dd, HH, MM, SS);
      if (!Number.isNaN(d.getTime())) return d;
    }
    return null;
  }

  function parseExcelDate(v){
    if (v == null) return null;
    if (v instanceof Date && !Number.isNaN(v.getTime())) return v;

    if (typeof v === "number" && window.XLSX?.SSF?.parse_date_code){
      const o = XLSX.SSF.parse_date_code(v);
      if (o && o.y && o.m && o.d){
        return new Date(o.y, (o.m||1)-1, o.d||1, o.H||0, o.M||0, o.S||0);
      }
    }

    if (typeof v === "string"){
      const d = new Date(v);
      if (!Number.isNaN(d.getTime())) return d;
      return parseDateAny(v);
    }
    return null;
  }

  const normDoctor = (x) => (x==null ? "" : String(x)).trim();
  const normPatient = (x) => (x==null ? "" : String(x)).trim();

  // -----------------------------
  // Parse HoraIn -> hour (0..23) o null
  // -----------------------------
  function parseHourAny(v){
    if (v == null) return null;

    if (v instanceof Date && !Number.isNaN(v.getTime())) {
      const h = v.getHours();
      return (h>=0 && h<=23) ? h : null;
    }

    if (typeof v === "number" && Number.isFinite(v)){
      if (v >= 0 && v <= 23) return Math.floor(v);
      if (v >= 0 && v < 1) return Math.floor(v * 24);
      if (v >= 0 && v < 24.5) return Math.floor(v);
      return null;
    }

    const s = String(v).trim();
    if (!s) return null;

    if (/^\d{1,2}$/.test(s)){
      const h = Number(s);
      return (h>=0 && h<=23) ? h : null;
    }

    const m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (m){
      const h = Number(m[1]);
      return (h>=0 && h<=23) ? h : null;
    }

    const d = new Date(s);
    if (!Number.isNaN(d.getTime())){
      const h = d.getHours();
      return (h>=0 && h<=23) ? h : null;
    }

    return null;
  }

  // -----------------------------
  // Estado + índices
  // -----------------------------
  let idx = null;
  let selectedDateKey = "";
  let selectedDoctor = "";
  let selectedMonthKey = "";
  let calMonthKey = "";

  // filas base completas (sin filtro)
  let allRows = [];
  let hasHoraIn = false;

  // 3 modos v5.x
  let hourMode = "8to23"; // "8to23" | "0to8" | "turno"

  function buildIndex(rows){
    const byDateDoctor = new Map(); // dateKey -> Map(doctor -> Set(patient))
    const doctorDays = new Map();   // doctor -> Map(dateKey -> count)
    const dayTotal = new Map();     // dateKey -> total pacientes únicos urgencias (union Atendido)

    for (const r of rows){
      if (!r.dateKey || !r.doctor || !r.patient) continue;
      if (!byDateDoctor.has(r.dateKey)) byDateDoctor.set(r.dateKey, new Map());
      const m = byDateDoctor.get(r.dateKey);
      if (!m.has(r.doctor)) m.set(r.doctor, new Set());
      m.get(r.doctor).add(r.patient);
    }

    for (const [dk, m] of byDateDoctor.entries()){
      const union = new Set();
      for (const setp of m.values()){
        for (const p of setp) union.add(p);
      }
      dayTotal.set(dk, union.size);
    }

    for (const [dk, m] of byDateDoctor.entries()){
      for (const [doc, setp] of m.entries()){
        if (!doctorDays.has(doc)) doctorDays.set(doc, new Map());
        doctorDays.get(doc).set(dk, setp.size);
      }
    }

    const byMonthDoctor = new Map();
    for (const [doc, dmap] of doctorDays.entries()){
      for (const [dk, cnt] of dmap.entries()){
        const mk = dk.slice(0,7);
        if (!byMonthDoctor.has(mk)) byMonthDoctor.set(mk, new Map());
        const mm = byMonthDoctor.get(mk);
        mm.set(doc, (mm.get(doc)||0) + cnt);
      }
    }

    const dayMax = new Map();
    const dayMin = new Map();
    for (const [dk, m] of byDateDoctor.entries()){
      let mx = 0, mn = Infinity;
      for (const setp of m.values()){
        const c = setp.size;
        mx = Math.max(mx, c);
        mn = Math.min(mn, c);
      }
      dayMax.set(dk, mx);
      dayMin.set(dk, Number.isFinite(mn) ? mn : 0);
    }

    return { byDateDoctor, doctorDays, byMonthDoctor, dayMax, dayMin, dayTotal };
  }

  const getAllDates = () => Array.from(idx.byDateDoctor.keys()).sort();
  function getAllDoctors(){
    const s = new Set();
    for (const m of idx.byDateDoctor.values()) for (const doc of m.keys()) s.add(doc);
    return Array.from(s).sort((a,b)=>a.localeCompare(b,"es"));
  }

  function doctorsForDay(dk){
    const m = idx.byDateDoctor.get(dk);
    if (!m) return [];
    const arr = Array.from(m.entries()).map(([doctor,setp]) => ({doctor, count:setp.size}));
    arr.sort((a,b)=> (b.count-a.count) || a.doctor.localeCompare(b.doctor,"es"));
    return arr;
  }

  function monthDoctorTotals(mk){
    const m = idx.byMonthDoctor.get(mk);
    if (!m) return [];
    const arr = Array.from(m.entries()).map(([doctor,total]) => ({doctor,total}));
    arr.sort((a,b)=> (b.total-a.total) || a.doctor.localeCompare(b.doctor,"es"));
    return arr;
  }

  // -----------------------------
  // CSV parser simple
  // -----------------------------
  function parseCSV(text){
    const rows = [];
    let i=0, field="", row=[], inQuotes=false;
    while (i < text.length){
      const c = text[i];
      if (inQuotes){
        if (c === '"'){
          if (text[i+1] === '"'){ field+='"'; i+=2; continue; }
          inQuotes=false; i++; continue;
        } else { field += c; i++; continue; }
      } else {
        if (c === '"'){ inQuotes=true; i++; continue; }
        if (c === ','){ row.push(field); field=""; i++; continue; }
        if (c === '\r'){ i++; continue; }
        if (c === '\n'){ row.push(field); field=""; rows.push(row); row=[]; i++; continue; }
        field += c; i++; continue;
      }
    }
    if (field.length || row.length){ row.push(field); rows.push(row); }
    return rows;
  }

  function fillSelect(sel, values, labelFn, selected){
    sel.innerHTML = "";
    for (const v of values){
      const opt = document.createElement("option");
      opt.value = v;
      opt.textContent = labelFn(v);
      sel.appendChild(opt);
    }
    sel.value = selected || (values[0] || "");
  }

  // -----------------------------
  // v5.x: modos hora + Turno
  // -----------------------------
  function addDaysKey(dateKey, deltaDays){
    const [y,m,d] = dateKey.split("-").map(Number);
    const dt = new Date(y, m-1, d);
    dt.setDate(dt.getDate() + deltaDays);
    return toDateKey(dt);
  }

  function hourModeLabel(){
    if (!hasHoraIn) return "HoraIn no disponible";
    if (hourMode === "8to23") return "HoraIn 8–23";
    if (hourMode === "0to8") return "HoraIn 0–8";
    return "Turno (08:00-08:00)";
  }

  function turnoWindowText(dateKey){
    const startKey = addDaysKey(dateKey, -1);
    return `Turno: 08:00 ${formatES(startKey)} → 08:00 ${formatES(dateKey)}`;
  }

  function updateHourButton(){
    const b = $("btnHourMode");
    if (!b) return;

    if (!hasHoraIn){
      b.textContent = "🕒 8–23";
      b.title = "Filtro horario (requiere columna HoraIn)";
      return;
    }

    if (hourMode === "8to23"){
      b.textContent = "🕒 8–23";
      b.title = "Modo: Día (HoraIn 8–23)";
    } else if (hourMode === "0to8"){
      b.textContent = "🕒 0–8";
      b.title = "Modo: Noche (HoraIn 0–8)";
    } else {
      b.textContent = "🕒 Turno";
      b.title = "Modo: Turno de Facultativo (08:00 día-1 → 08:00 día)";
    }
  }

  // Base -> Effective para índices (dashboard/calendario)
  function deriveEffectiveRows(){
    if (!allRows || !allRows.length) return [];

    if (!hasHoraIn){
      return allRows.map(r => ({
        dateKey: r.baseDateKey,
        monthKey: r.baseMonthKey,
        doctor: r.doctor,
        patient: r.patient
      }));
    }

    if (hourMode === "8to23"){
      return allRows
        .filter(r => r.hour == null || (r.hour >= 8 && r.hour <= 23))
        .map(r => ({
          dateKey: r.baseDateKey,
          monthKey: r.baseMonthKey,
          doctor: r.doctor,
          patient: r.patient
        }));
    }

    if (hourMode === "0to8"){
      return allRows
        .filter(r => r.hour == null || (r.hour >= 0 && r.hour < 8))
        .map(r => ({
          dateKey: r.baseDateKey,
          monthKey: r.baseMonthKey,
          doctor: r.doctor,
          patient: r.patient
        }));
    }

    // turno
    return allRows.map(r => ({
      dateKey: r.turnoDateKey,
      monthKey: r.turnoDateKey.slice(0,7),
      doctor: r.doctor,
      patient: r.patient
    }));
  }

  // Base -> Effective para “Pacientes Atendidos” (necesita dt para ordenar)
  function deriveEffectiveEventsForPatients(){
    if (!allRows || !allRows.length) return [];

    // dt real del evento: preferimos HoraIn como hora (si existe) sobre FECINGRESO (si trae hora ya)
    const buildEventDT = (r) => {
      if (r.dtBase instanceof Date && !Number.isNaN(r.dtBase.getTime())) {
        // si dtBase ya trae hora real y NO hay HoraIn interpretable, usamos dtBase
        if (r.hour == null) return r.dtBase;
      }
      // si hay hour, ponemos esa hora en el día base (00:00) con minutos 00
      if (r.baseDateKey && r.hour != null){
        const [y,m,d] = r.baseDateKey.split("-").map(Number);
        return new Date(y, m-1, d, r.hour, 0, 0);
      }
      // fallback
      return r.dtBase instanceof Date ? r.dtBase : null;
    };

    if (!hasHoraIn){
      return allRows.map(r => ({
        effDateKey: r.baseDateKey,
        doctor: r.doctor,
        patient: r.patient,
        dt: buildEventDT(r)
      }));
    }

    if (hourMode === "8to23"){
      return allRows
        .filter(r => r.hour == null || (r.hour >= 8 && r.hour <= 23))
        .map(r => ({
          effDateKey: r.baseDateKey,
          doctor: r.doctor,
          patient: r.patient,
          dt: buildEventDT(r)
        }));
    }

    if (hourMode === "0to8"){
      return allRows
        .filter(r => r.hour == null || (r.hour >= 0 && r.hour < 8))
        .map(r => ({
          effDateKey: r.baseDateKey,
          doctor: r.doctor,
          patient: r.patient,
          dt: buildEventDT(r)
        }));
    }

    // turno: agrupación por turnoDateKey, pero dt es el momento real
    return allRows.map(r => ({
      effDateKey: r.turnoDateKey,
      doctor: r.doctor,
      patient: r.patient,
      dt: buildEventDT(r)
    }));
  }

  function rebuildFromAllRows(){
    if (!allRows || !allRows.length) return;

    const effective = deriveEffectiveRows();
    idx = buildIndex(effective);

    const dates = getAllDates();
    const doctors = getAllDoctors();

    if (selectedDateKey && !dates.includes(selectedDateKey)) {
      selectedDateKey = dates[dates.length-1] || "";
    } else if (!selectedDateKey) {
      selectedDateKey = dates[dates.length-1] || "";
    }

    selectedMonthKey = selectedDateKey ? selectedDateKey.slice(0,7) : "";

    if (selectedDoctor && !doctors.includes(selectedDoctor)) {
      selectedDoctor = doctors[0] || "";
    } else if (!selectedDoctor) {
      selectedDoctor = doctors[0] || "";
    }

    $("selDate").disabled = !dates.length;
    $("selDoctor").disabled = !doctors.length;

    fillSelect($("selDate"), dates, formatES, selectedDateKey);
    fillSelect($("selDoctor"), doctors, (d)=>d, selectedDoctor);

    $("emptyState").style.display = "none";
    $("viewDash").style.display = "block";

    updateHourButton();
    refreshAll();
  }

  function loadNormalized(norm, filename){
    allRows = norm;
    $("pillFile").textContent = filename || "archivo";
    $("footLeft").textContent = "Última carga: " + new Date().toLocaleString("es-ES");

    selectedDateKey = "";
    selectedDoctor = "";
    selectedMonthKey = "";
    calMonthKey = "";

    updateHourButton();
    rebuildFromAllRows();
  }

  function loadFromCSVText(text, filename){
    const table = parseCSV(text);
    if (!table.length) throw new Error("CSV vacío");

    const header = table[0].map(h => (h||"").trim());
    const col = (name) => header.indexOf(name);

    const iF = col("FECINGRESO");
    const iM = col("NOMBMED_RES");
    const iA = col("Atendido");
    const iH = col("HoraIn");

    const missing = [];
    if (iF < 0) missing.push("FECINGRESO");
    if (iM < 0) missing.push("NOMBMED_RES");
    if (iA < 0) missing.push("Atendido");
    if (missing.length) throw new Error("Faltan columnas: " + missing.join(", "));

    hasHoraIn = (iH >= 0);

    const norm = [];
    for (let r=1; r<table.length; r++){
      const line = table[r];
      if (!line || !line.length) continue;

      const d = parseDateAny(line[iF]);
      if (!d) continue;

      const doctor = normDoctor(line[iM]);
      const patient = normPatient(line[iA]);
      if (!doctor || !patient) continue;

      const baseDateKey = toDateKey(d);
      const baseMonthKey = toMonthKey(d);

      const hour = hasHoraIn ? parseHourAny(line[iH]) : null;

      let turnoDateKey = baseDateKey;
      if (hour != null){
        turnoDateKey = (hour >= 8) ? addDaysKey(baseDateKey, 1) : baseDateKey;
      }

      norm.push({
        baseDateKey,
        baseMonthKey,
        turnoDateKey,
        doctor,
        patient,
        hour,
        dtBase: d
      });
    }
    if (!norm.length) toast("Sin datos útiles", "No se detectaron filas válidas. Revisa columnas y formato de fecha.");
    loadNormalized(norm, filename);
  }

  async function loadFromXLSXFile(file){
    if (!window.XLSX) throw new Error("No se cargó la librería XLSX (CDN).");
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type:"array", cellDates:true });
    const sheetName = wb.SheetNames[0];
    if (!sheetName) throw new Error("XLSX sin hojas.");
    const ws = wb.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
    if (!json.length) throw new Error("XLSX vacío (sin filas).");

    const required = ["FECINGRESO","NOMBMED_RES","Atendido"];
    const cols = Object.keys(json[0] || {});
    const missing = required.filter(r => !cols.includes(r));
    if (missing.length) throw new Error("Faltan columnas: " + missing.join(", "));

    hasHoraIn = cols.includes("HoraIn");

    const norm = [];
    for (const row of json){
      const d = parseExcelDate(row["FECINGRESO"]);
      if (!d) continue;

      const doctor = normDoctor(row["NOMBMED_RES"]);
      const patient = normPatient(row["Atendido"]);
      if (!doctor || !patient) continue;

      const baseDateKey = toDateKey(d);
      const baseMonthKey = toMonthKey(d);

      const hour = hasHoraIn ? parseHourAny(row["HoraIn"]) : null;

      let turnoDateKey = baseDateKey;
      if (hour != null){
        turnoDateKey = (hour >= 8) ? addDaysKey(baseDateKey, 1) : baseDateKey;
      }

      norm.push({
        baseDateKey,
        baseMonthKey,
        turnoDateKey,
        doctor,
        patient,
        hour,
        dtBase: d
      });
    }
    if (!norm.length) toast("Sin datos útiles", "No se detectaron filas válidas. Revisa columnas y formato de FECINGRESO.");
    loadNormalized(norm, file.name);
  }

  // -----------------------------
  // Render helpers
  // -----------------------------
  function renderTop3(arr){
    const box = $("top3");
    box.innerHTML = "";
    if (!arr.length){
      box.innerHTML = `<div class="empty">Sin registros para la fecha seleccionada.</div>`;
      return;
    }
    arr.forEach((x,i)=>{
      const div = document.createElement("div");
      div.style.display = "flex";
      div.style.justifyContent = "space-between";
      div.style.alignItems = "center";
      div.innerHTML = `
        <div class="docName" title="${escapeHtml(x.doctor)}">
          <span class="tag">#${i+1}</span>${escapeHtml(x.doctor)}
        </div>
        <div class="bigNum">${x.count}</div>
      `;
      box.appendChild(div);
    });
  }

  function renderRanking(container, arr, dateKey, clickableToDoctor){
    container.innerHTML = "";
    if (!arr.length){
      container.innerHTML = `<div class="empty">No hay registros para la fecha seleccionada.</div>`;
      return;
    }
    arr.forEach((x,i)=>{
      const item = document.createElement(clickableToDoctor ? "button" : "div");
      item.className = "rankItem";
      if (clickableToDoctor){
        item.style.cursor = "pointer";
        item.title = "Click para ver detalle del facultativo";
        item.addEventListener("click", () => {
          selectedDoctor = x.doctor;
          $("selDoctor").value = selectedDoctor;
          setActiveTab("dash");
          $("viewDash").style.display = "block";
          $("viewCal").style.display = "none";
          refreshAll();
        });
      }
      item.innerHTML = `
        <div class="left">
          <div class="docName" title="${escapeHtml(x.doctor)}">
            <span class="tag">#${i+1}</span>${escapeHtml(x.doctor)}
          </div>
          <div class="muted2">${formatES(dateKey)}</div>
        </div>
        <div class="bigNum">${x.count}</div>
      `;
      container.appendChild(item);
    });
  }

  // ✅ v5.1: Pacientes atendidos (detalle)
  // ✅ v5.1: Pacientes atendidos (detalle) — alineado (fecha/hora a la izquierda, Atendido a la derecha)
function renderPatientsAttended(){
  const list = $("patientsList");
  const sub = $("patientsSubtitle");
  if (!list || !sub) return;

  if (!idx || !selectedDoctor || !selectedDateKey){
    sub.textContent = "Selecciona un facultativo para ver el detalle del día";
    list.innerHTML = `<div class="empty">Sin selección.</div>`;
    return;
  }

  const events = deriveEffectiveEventsForPatients()
    .filter(e => e.effDateKey === selectedDateKey && e.doctor === selectedDoctor)
    .sort((a,b) => {
      const ta = a.dt instanceof Date ? a.dt.getTime() : 0;
      const tb = b.dt instanceof Date ? b.dt.getTime() : 0;
      return ta - tb;
    });

  const modeTxt = hourModeLabel();
  const extra = (hasHoraIn && hourMode === "turno") ? ` · ${turnoWindowText(selectedDateKey)}` : "";
  sub.textContent = `${selectedDoctor} · ${formatES(selectedDateKey)} · ${modeTxt}${extra}`;

  if (!events.length){
    list.innerHTML = `<div class="empty">No hay atenciones para este facultativo en el día seleccionado con el filtro actual.</div>`;
    return;
  }

  list.innerHTML = "";
  events.forEach((e, i) => {
    const item = document.createElement("div");
    item.className = "rankItem";

    const dtTxt = formatDateTimeES(e.dt);
    const atendidoKey = escapeHtml(e.patient);

    // Alineado: izquierda # + fecha/hora, derecha Atendido (código)
    item.innerHTML = `
      <div class="left">
        <div class="docName" style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
          <span class="tag">#${i+1}</span>
          <span class="muted2" style="margin:0">${escapeHtml(dtTxt)}</span>
        </div>
      </div>
      <div class="bigNum" style="font-weight:900">${atendidoKey}</div>
    `;

    list.appendChild(item);
  });
}


  // -----------------------------
  // Chart (sin cambios respecto a tu versión)
  // -----------------------------
  function roundRect(ctx, x, y, w, h, r){
    if (h <= 0) return;
    const rr = Math.min(r, w/2, h/2);
    ctx.beginPath();
    ctx.moveTo(x+rr, y);
    ctx.arcTo(x+w, y, x+w, y+h, rr);
    ctx.arcTo(x+w, y+h, x, y+h, rr);
    ctx.arcTo(x, y+h, x, y, rr);
    ctx.arcTo(x, y, x+w, y, rr);
    ctx.closePath();
    ctx.fill();
  }

  function drawBarChart(canvas, data){
    if (!canvas) return;
    const ctx = canvas.getContext("2d");

    const cs = getComputedStyle(document.documentElement);
    const GRID = cs.getPropertyValue("--chart-grid").trim();
    const INK  = cs.getPropertyValue("--chart-ink").trim();
    const BAR  = cs.getPropertyValue("--chart-bar").trim();
    const BAR_A = parseFloat(cs.getPropertyValue("--chart-bar-alpha")) || 0.55;
    const GLOW_A = parseFloat(cs.getPropertyValue("--chart-glow-alpha")) || 0.10;

    const cssW = canvas.clientWidth;
    const cssH = canvas.clientHeight;
    const dpr = window.devicePixelRatio || 1;
    canvas.width = Math.floor(cssW * dpr);
    canvas.height = Math.floor(cssH * dpr);
    ctx.setTransform(dpr,0,0,dpr,0,0);
    ctx.clearRect(0,0,cssW,cssH);

    const padL = 38, padR = 14, padT = 14, padB = 28;
    const w = cssW - padL - padR;
    const h = cssH - padT - padB;

    ctx.globalAlpha = 0.22;
    ctx.strokeStyle = GRID;
    ctx.lineWidth = 1;
    const lines = 4;
    for (let i=0;i<=lines;i++){
      const y = padT + (h/lines)*i;
      ctx.beginPath();
      ctx.moveTo(padL, y);
      ctx.lineTo(padL+w, y);
      ctx.stroke();
    }
    ctx.globalAlpha = 1;

    const maxV = Math.max(0, ...data.map(d=>d.value));
    const n = Math.max(1, data.length);
    const gap = 6;
    const barW = Math.max(6, (w - gap*(n-1)) / n);

    ctx.fillStyle = INK;
    ctx.font = "12px ui-sans-serif, system-ui";
    for (let i=0;i<=lines;i++){
      const v = Math.round(maxV - (maxV/lines)*i);
      const y = padT + (h/lines)*i + 4;
      ctx.fillText(String(v), 6, y);
    }

    data.forEach((d, i)=>{
      const x = padL + i*(barW+gap);
      const bh = maxV ? (d.value/maxV)*h : 0;
      const y = padT + (h - bh);

      ctx.globalAlpha = GLOW_A;
      ctx.fillStyle = BAR;
      roundRect(ctx, x, y-2, barW, bh+2, 10);

      ctx.globalAlpha = BAR_A;
      ctx.fillStyle = BAR;
      roundRect(ctx, x, y, barW, bh, 10);

      const theme = document.documentElement.getAttribute("data-theme");
      const labelFill = (theme === "dark")
        ? "rgba(0,0,0,.88)"
        : "rgba(255,255,255,.92)";

      const valueStr = String(d.value);
      const fontSize = Math.max(12, Math.min(15, Math.floor(barW * 0.35)));
      ctx.font = `900 ${fontSize}px ui-sans-serif, system-ui`;
      ctx.fillStyle = labelFill;
      ctx.globalAlpha = 1;

      const textW = ctx.measureText(valueStr).width;

      if (bh >= fontSize + 10) {
        const tx = x + (barW - textW)/2;
        const ty = y + bh/2 + fontSize/3;
        ctx.fillText(valueStr, tx, ty);
      } else {
        ctx.fillStyle = INK;
        const tx = x + (barW - textW)/2;
        const ty = y - 6;
        ctx.fillText(valueStr, tx, ty);
      }

      ctx.fillStyle = INK;
      ctx.font = "11px ui-sans-serif, system-ui";
      const lab = String(d.day);
      const tw = ctx.measureText(lab).width;
      ctx.fillText(lab, x + (barW - tw)/2, padT + h + 18);
    });
  }

  function refreshDoctorSection(){
    if (!idx || !selectedDoctor) return;

    const dayMap = idx.doctorDays.get(selectedDoctor) || new Map();
    const all = Array.from(dayMap.entries())
      .map(([dateKey,count])=>({dateKey,count}))
      .sort((a,b)=>a.dateKey.localeCompare(b.dateKey));

    const inMonth = all.filter(x => x.dateKey.startsWith(selectedMonthKey));
    const sumAll = all.reduce((acc,x)=>acc+x.count,0);
    const sumMonth = inMonth.reduce((acc,x)=>acc+x.count,0);

    const maxAll = all.length ? Math.max(...all.map(x=>x.count)) : 0;
    const minAll = all.length ? Math.min(...all.map(x=>x.count)) : 0;
    const maxMonth = inMonth.length ? Math.max(...inMonth.map(x=>x.count)) : 0;
    const minMonth = inMonth.length ? Math.min(...inMonth.map(x=>x.count)) : 0;

    $("docSubtitle").textContent = `Análisis de ${selectedDoctor}`;
    $("docMonthLabel").textContent = monthLabelES(selectedMonthKey);

    $("docDaysMonth").textContent = inMonth.length;
    $("docTotalMonth").innerHTML = `Total pacientes (mes): <b class="docBig">${sumMonth}</b>`;
    $("docMaxMonth").textContent = maxMonth;
    $("docMinMonth").textContent = minMonth;

    $("docDaysAll").textContent = all.length;
    $("docTotalAll").textContent = sumAll;
    $("docMaxAll").textContent = maxAll;
    $("docMinAll").textContent = minAll;

    $("docChartHint").textContent = `${monthLabelES(selectedMonthKey)} · ${inMonth.length} días con actividad`;
    $("docChartTotal").textContent = sumMonth;

    const chartData = inMonth.map(x => ({
      day: Number(x.dateKey.slice(-2)),
      value: x.count,
      dateKey: x.dateKey
    }));
    drawBarChart($("chart"), chartData);
  }

  // -----------------------------
  // Calendario
  // -----------------------------
  function renderCalendar(){
    if (!idx || !calMonthKey) return;

    $("calMonthLabel").textContent = monthLabelES(calMonthKey);

    $("calDow").innerHTML = "";
    ["L","M","X","J","V","S","D"].forEach(w=>{
      const d = document.createElement("div");
      d.className = "dow";
      d.textContent = w;
      $("calDow").appendChild(d);
    });

    const [y,m] = calMonthKey.split("-").map(Number);
    const first = new Date(y, m-1, 1);
    const firstDow = (first.getDay() + 6) % 7;
    const daysInMonth = new Date(y, m, 0).getDate();

    const totalMap = {};
    for (const [dk,v] of idx.dayTotal.entries()) totalMap[dk] = v;
    const monthKeys = Object.keys(totalMap).filter(dk => dk.startsWith(calMonthKey));
    const monthMaxTotal = Math.max(1, ...monthKeys.map(k=>totalMap[k]||0));

    const cells = [];
    for (let i=0;i<firstDow;i++) cells.push(null);
    for (let d=1; d<=daysInMonth; d++){
      const dk = `${y}-${pad2(m)}-${pad2(d)}`;
      cells.push(dk);
    }
    while (cells.length % 7 !== 0) cells.push(null);

    const grid = $("calGrid");
    grid.innerHTML = "";
    cells.forEach(dk=>{
      const div = document.createElement("div");
      if (!dk){
        div.style.height = "78px";
        div.style.borderRadius = "20px";
        grid.appendChild(div);
        return;
      }

      const maxMed = (idx.dayMax.get(dk) || 0);
      const total = (totalMap[dk] || 0);
      const intensity = Math.min(1, total / monthMaxTotal);

      div.className = "dayCell" + (dk === selectedDateKey ? " sel" : "");
      div.title = `${formatES(dk)} — Total urgencias: ${total} · Max por médico: ${maxMed}`;

      div.style.setProperty("--a1", (0.08 + 0.42*intensity).toFixed(3));
      div.style.setProperty("--a2", (0.06 + 0.30*intensity).toFixed(3));

      div.innerHTML = `
        <div class="top">
          <div class="d">${Number(dk.slice(-2))}</div>
          <div class="kpiMax">MAX <b>${maxMed}</b></div>
        </div>
        <div class="mid">
          <div class="totalBig">${total}</div>
        </div>
        <div class="totalLab">Total urgencias</div>
      `;

      div.addEventListener("click", ()=>{
        selectedDateKey = dk;
        $("selDate").value = selectedDateKey;
        selectedMonthKey = selectedDateKey.slice(0,7);
        refreshAll();
        setActiveTab("cal");
        $("viewDash").style.display = "none";
        $("viewCal").style.display = "block";
      });

      grid.appendChild(div);
    });
  }

  function renderCalRanking(){
    const prefix =
      (hasHoraIn && hourMode === "turno")
        ? "Ranking (Turno) — "
        : "Ranking (Día) — ";

    $("calRankTitle").textContent = selectedDateKey ? `${prefix}${formatES(selectedDateKey)}` : "—";
    const dayArr = doctorsForDay(selectedDateKey);
    renderRanking($("calRankList"), dayArr.slice(0,10), selectedDateKey, true);
  }

  // -----------------------------
  // Refresh principal
  // -----------------------------
  function refreshAll(){
    if (!idx || !selectedDateKey) return;

    selectedMonthKey = selectedDateKey.slice(0,7);
    if (!calMonthKey) calMonthKey = selectedMonthKey;

    $("pillDate").textContent = formatES(selectedDateKey);
    $("pillMonth").textContent = monthLabelES(selectedMonthKey);

    const modeTxt = hourModeLabel();
    const extra = (hasHoraIn && hourMode === "turno") ? ` · ${turnoWindowText(selectedDateKey)}` : "";

    $("dashSubtitle").textContent =
      `Resumen de actividad para ${formatES(selectedDateKey)} (Atendido distinto por médico y día · ${modeTxt}${extra})`;

    const dayArr = doctorsForDay(selectedDateKey);
    const dayActive = dayArr.length;
    const dayMax = idx.dayMax.get(selectedDateKey) || 0;
    const dayMin = idx.dayMin.get(selectedDateKey) || 0;

    $("kpiActive").textContent = dayActive;
    $("kpiDayMax").textContent = dayMax;
    $("kpiDayMin").textContent = dayMin;

    $("cardTotalDocs").textContent = dayActive;
    $("cardDayMax").textContent = dayMax;
    $("cardDayMin").textContent = dayMin;

    renderTop3(dayArr.slice(0,3));
    renderRanking($("rankTop10"), dayArr, selectedDateKey, true);

    const monthTotals = monthDoctorTotals(selectedMonthKey);
    const monthBest = monthTotals[0] || null;
    const monthWorst = monthTotals.length ? monthTotals[monthTotals.length-1] : null;

    $("cardMonthMax").textContent = monthBest ? monthBest.total : 0;
    $("cardMonthMin").textContent = monthWorst ? monthWorst.total : 0;
    $("hintMonthBest").textContent = monthBest ? `Mejor del mes: ${monthBest.doctor}` : "—";
    $("hintMonthWorst").textContent = monthWorst ? `Menor del mes: ${monthWorst.doctor}` : "—";

    $("monthBestName").textContent = monthBest ? monthBest.doctor : "—";
    $("monthBestTotal").textContent = "Total mes: " + (monthBest ? monthBest.total : 0);
    $("monthWorstName").textContent = monthWorst ? monthWorst.doctor : "—";
    $("monthWorstTotal").textContent = "Total mes: " + (monthWorst ? monthWorst.total : 0);

    const monthDates = getAllDates().filter(dk => dk.startsWith(selectedMonthKey));
    let mx = 0, mn = Infinity;
    for (const dk of monthDates){
      mx = Math.max(mx, idx.dayMax.get(dk)||0);
      mn = Math.min(mn, idx.dayMin.get(dk)||0);
    }
    $("monthDailyMax").textContent = mx;
    $("monthDailyMin").textContent = Number.isFinite(mn) ? mn : 0;

    refreshDoctorSection();
    renderPatientsAttended(); // ✅ v5.1
    renderCalendar();
    renderCalRanking();
    updateHourButton();
  }

  // Tabs
  function setActiveTab(which){
    if (which === "dash"){
      $("tabDash").classList.add("active");
      $("tabCal").classList.remove("active");
    } else {
      $("tabCal").classList.add("active");
      $("tabDash").classList.remove("active");
    }
  }

  // -----------------------------
  // Eventos UI
  // -----------------------------
  $("btnUpload").addEventListener("click", ()=> $("file").click());

  $("file").addEventListener("change", async (ev)=>{
    const file = ev.target.files && ev.target.files[0];
    if (!file) return;

    try{
      const name = (file.name || "").toLowerCase();
      if (name.endsWith(".csv")){
        const text = await file.text();
        loadFromCSVText(text, file.name);
        toast("CSV cargado", "Dashboard actualizado.");
      } else if (name.endsWith(".xlsx") || name.endsWith(".xls")){
        await loadFromXLSXFile(file);
        toast("XLSX cargado", "Dashboard actualizado.");
      } else {
        throw new Error("Formato no soportado. Usa .csv, .xlsx o .xls");
      }
    } catch (e){
      toast("Error al cargar", String(e.message || e));
    } finally {
      ev.target.value = "";
    }
  });

  // Botón modo horario: 8–23 -> 0–8 -> Turno -> 8–23
  $("btnHourMode").addEventListener("click", ()=>{
    if (!idx && (!allRows || !allRows.length)){
      return toast("Sin datos", "Primero carga un CSV/XLSX.");
    }
    if (!hasHoraIn){
      return toast("HoraIn no disponible", "Tu fichero no tiene la columna HoraIn, no se puede filtrar por horas ni por turnos.");
    }

    if (hourMode === "8to23") hourMode = "0to8";
    else if (hourMode === "0to8") hourMode = "turno";
    else hourMode = "8to23";

    rebuildFromAllRows();

    const msg =
      hourMode === "8to23"
        ? "Mostrando HoraIn 8–23."
        : hourMode === "0to8"
          ? "Mostrando HoraIn 0–8 (0–7)."
          : "Mostrando Turno: 08:00 día-1 → 08:00 día.";

    toast("Filtro horario", msg);
  });

  $("btnReset").addEventListener("click", ()=>{
    idx = null;
    selectedDateKey = "";
    selectedDoctor = "";
    selectedMonthKey = "";
    calMonthKey = "";

    allRows = [];
    hasHoraIn = false;
    hourMode = "8to23";
    updateHourButton();

    $("pillDate").textContent = "—";
    $("pillMonth").textContent = "—";
    $("pillFile").textContent = "—";
    $("selDate").innerHTML = `<option value="">(Carga un archivo)</option>`;
    $("selDoctor").innerHTML = `<option value="">(Sin datos)</option>`;
    $("selDate").disabled = true;
    $("selDoctor").disabled = true;
    $("emptyState").style.display = "block";
    $("viewDash").style.display = "none";
    $("viewCal").style.display = "none";
    $("footLeft").textContent = "Sin archivo cargado";
    setActiveTab("dash");
    toast("Reset", "Se limpió el dashboard.");
  });

  $("selDate").addEventListener("change", (e)=>{
    selectedDateKey = e.target.value;
    selectedMonthKey = selectedDateKey ? selectedDateKey.slice(0,7) : "";
    calMonthKey = selectedMonthKey;
    refreshAll();
  });

  $("selDoctor").addEventListener("change", (e)=>{
    selectedDoctor = e.target.value;
    refreshDoctorSection();
    renderPatientsAttended(); // ✅ v5.1
  });

  $("tabDash").addEventListener("click", ()=>{
    if (!idx) return toast("Sin datos", "Primero carga un CSV/XLSX.");
    setActiveTab("dash");
    $("viewDash").style.display = "block";
    $("viewCal").style.display = "none";
  });

  $("tabCal").addEventListener("click", ()=>{
    if (!idx) return toast("Sin datos", "Primero carga un CSV/XLSX.");
    setActiveTab("cal");
    $("viewDash").style.display = "none";
    $("viewCal").style.display = "block";
    renderCalendar();
    renderCalRanking();
  });

  $("btnGoDash").addEventListener("click", ()=>{
    setActiveTab("dash");
    $("viewDash").style.display = "block";
    $("viewCal").style.display = "none";
  });

  $("btnPrevMonth").addEventListener("click", ()=>{
    if (!calMonthKey) return;
    const [y,m] = calMonthKey.split("-").map(Number);
    const d = new Date(y, m-1, 1);
    d.setMonth(d.getMonth()-1);
    calMonthKey = toMonthKey(d);
    renderCalendar();
  });

  $("btnNextMonth").addEventListener("click", ()=>{
    if (!calMonthKey) return;
    const [y,m] = calMonthKey.split("-").map(Number);
    const d = new Date(y, m-1, 1);
    d.setMonth(d.getMonth()+1);
    calMonthKey = toMonthKey(d);
    renderCalendar();
  });

  window.addEventListener("resize", ()=> {
    if (!idx || !selectedDoctor) return;
    refreshDoctorSection();
  });

  updateHourButton();
})();
