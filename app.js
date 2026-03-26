/* global html2canvas */

const DEFAULT_PEOPLE = [
  "Georgi Valeriev",
  "Magüi Cerdá",
  "Antonella Sipan",
  "Iñigo Puyol",
  "Luz Romero",
  "Patricia López",
  "Jorge Romera",
  "Irene Peñalosa",
  "Maria Jose Rubio",
  "Alessandra Solis",
  "Adrian Garces",
  "Ignacio Rivas",
  "Alonso García",
  "Rodrigo Fernandez",
  "Lara Carrasco",
];

const DAYS = [
  { key: "JUE", label: "Jueves" },
  { key: "VIE", label: "Viernes" },
  { key: "SAB", label: "Sábado" },
  { key: "DOM", label: "Domingo" },
  { key: "LUN", label: "Lunes" },
  { key: "MAR", label: "Martes" },
  { key: "MIE", label: "Miércoles" },
];

const FRANJAS = [
  { key: "MANANA", label: "Mañana", hours: "8–14" },
  { key: "TARDE", label: "Tarde", hours: "14–22" },
  { key: "NOCHE", label: "Noche", hours: "22–8" },
];

const TIPOS = [
  { key: "FIJO", label: "Fijo" },
  { key: "BACKUP", label: "Back-up" },
];

function qs(id) {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Falta elemento #${id}`);
  return el;
}

function toISODate(d) {
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function parseISODate(s) {
  if (!s) return null;
  const [y, m, d] = s.split("-").map((x) => Number(x));
  const dt = new Date(y, m - 1, d);
  if (Number.isNaN(dt.getTime())) return null;
  return dt;
}

function addDays(date, n) {
  const d = new Date(date);
  d.setDate(d.getDate() + n);
  return d;
}

function startOfDay(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

function formatHumanDate(date) {
  return date.toLocaleDateString("es-ES", {
    weekday: "long",
    year: "numeric",
    month: "long",
    day: "2-digit",
  });
}

function clampStr(s) {
  return String(s ?? "").trim();
}

function normalizeName(name) {
  return clampStr(name).replace(/\s+/g, " ");
}

function todayISO() {
  return toISODate(new Date());
}

function computeWeekStartThursday(fromDate = new Date()) {
  // JS: 0=domingo..6=sábado. Queremos jueves=4.
  const d = startOfDay(fromDate);
  const dow = d.getDay();
  const target = 4;
  const diff = (dow - target + 7) % 7;
  return addDays(d, -diff);
}

function buildWeek(weekStart) {
  const base = startOfDay(weekStart);
  return DAYS.map((day, idx) => ({
    ...day,
    date: addDays(base, idx),
    iso: toISODate(addDays(base, idx)),
  }));
}

function slotId(iso, franja, tipo) {
  return `${iso}__${franja}__${tipo}`;
}

function splitSlotId(id) {
  const [iso, franja, tipo] = String(id).split("__");
  return { iso, franja, tipo };
}

function createEmptySchedule(weekStartISO) {
  const weekStart = parseISODate(weekStartISO);
  const week = buildWeek(weekStart);
  const slots = {};
  for (const d of week) {
    for (const f of FRANJAS) {
      for (const t of TIPOS) {
        const id = slotId(d.iso, f.key, t.key);
        slots[id] = {
          id,
          weekStart: weekStartISO,
          fecha: d.iso,
          franja: f.key,
          tipo: t.key,
          modo: "NORMAL", // NORMAL o TODOS
          asignadoA: "",
          nota: "",
        };
      }
    }
  }
  return { weekStart: weekStartISO, slots };
}

function localConfigGet() {
  try {
    const raw = localStorage.getItem("turnosPazy.config");
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function localConfigSet(next) {
  localStorage.setItem("turnosPazy.config", JSON.stringify(next));
}

function setStatus(text, kind = "muted") {
  const el = qs("status");
  el.className = `status ${kind === "muted" ? "muted" : kind}`;
  el.textContent = text;
}

function buildOption(label, value) {
  const opt = document.createElement("option");
  opt.value = value;
  opt.textContent = label;
  return opt;
}

function unique(arr) {
  return Array.from(new Set(arr));
}

function sortNames(names) {
  return [...names].sort((a, b) => a.localeCompare(b, "es", { sensitivity: "base" }));
}

function makePeopleModel(rawPeople, vacationsByISO) {
  const all = sortNames(
    unique((rawPeople?.length ? rawPeople : DEFAULT_PEOPLE).map(normalizeName)).filter(Boolean),
  );
  const byIso = vacationsByISO || {};
  return { all, vacationsByISO: byIso };
}

function availablePeopleForDate(peopleModel, iso) {
  const blocked = new Set((peopleModel.vacationsByISO?.[iso] || []).map(normalizeName));
  const excluded = getExcludedSet();
  return peopleModel.all.filter((n) => !blocked.has(normalizeName(n)) && !excluded.has(normalizeName(n)));
}

function scheduleToRows(schedule) {
  return Object.values(schedule.slots);
}

function computeSummary(schedule) {
  const counters = new Map();
  for (const slot of scheduleToRows(schedule)) {
    if (slot.modo === "TODOS") continue;
    const name = normalizeName(slot.asignadoA);
    if (!name) continue;
    const cur = counters.get(name) || { fijo: 0, backup: 0 };
    if (slot.tipo === "FIJO") cur.fijo += 1;
    if (slot.tipo === "BACKUP") cur.backup += 1;
    counters.set(name, cur);
  }
  const rows = Array.from(counters.entries()).map(([name, c]) => ({
    name,
    fijo: c.fijo,
    backup: c.backup,
    total: c.fijo + c.backup,
  }));
  rows.sort((a, b) => b.total - a.total || a.name.localeCompare(b.name, "es"));
  return rows;
}

function setSlot(schedule, id, patch) {
  if (!schedule.slots[id]) return;
  schedule.slots[id] = { ...schedule.slots[id], ...patch };
}

function setTodosForDayFranja(schedule, iso, franjaKey, on) {
  for (const t of TIPOS) {
    const id = slotId(iso, franjaKey, t.key);
    if (!schedule.slots[id]) continue;
    schedule.slots[id] = {
      ...schedule.slots[id],
      modo: on ? "TODOS" : "NORMAL",
      asignadoA: on ? "" : schedule.slots[id].asignadoA,
    };
  }
}

function isTodosDayFranja(schedule, iso, franjaKey) {
  const idF = slotId(iso, franjaKey, "FIJO");
  return schedule.slots[idF]?.modo === "TODOS";
}

function mulberry32(seed) {
  let t = seed >>> 0;
  return function rand() {
    t += 0x6d2b79f5;
    let x = t;
    x = Math.imul(x ^ (x >>> 15), x | 1);
    x ^= x + Math.imul(x ^ (x >>> 7), x | 61);
    return ((x ^ (x >>> 14)) >>> 0) / 4294967296;
  };
}

function hashStringToSeed(str) {
  const s = String(str || "");
  let h = 2166136261;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return h >>> 0;
}

function isWeekendISO(iso) {
  const d = parseISODate(iso);
  if (!d) return false;
  const dow = d.getDay(); // 0 dom, 6 sab
  return dow === 0 || dow === 6;
}

function initStats(people) {
  const stats = new Map();
  for (const p of people) {
    stats.set(p, {
      total: 0,
      fijo: 0,
      backup: 0,
      byFranja: { MANANA: 0, TARDE: 0, NOCHE: 0 },
      weekend: 0,
      weekday: 0,
      lastAssignedIso: null,
      lastAssignedDayIndex: null,
    });
  }
  return stats;
}

function recordAssignment(stats, name, iso, dayIndex, franjaKey, tipoKey) {
  const s = stats.get(name);
  if (!s) return;
  s.total += 1;
  if (tipoKey === "FIJO") s.fijo += 1;
  if (tipoKey === "BACKUP") s.backup += 1;
  if (s.byFranja[franjaKey] != null) s.byFranja[franjaKey] += 1;
  if (isWeekendISO(iso)) s.weekend += 1;
  else s.weekday += 1;
  s.lastAssignedIso = iso;
  s.lastAssignedDayIndex = dayIndex;
}

function scoreCandidate({ stats, name, iso, dayIndex, franjaKey, tipoKey, dayAssignedSet, rng }) {
  const s = stats.get(name);
  if (!s) return Number.POSITIVE_INFINITY;

  const weekend = isWeekendISO(iso);

  // Balance base: total, tipo, franja, finde/entre semana
  let score = 0;
  score += s.total * 10;
  score += (tipoKey === "FIJO" ? s.fijo : s.backup) * 12;
  score += (s.byFranja[franjaKey] || 0) * 8;
  score += (weekend ? s.weekend : s.weekday) * 7;

  // Penalizaciones: repetir el mismo día, días consecutivos, etc.
  if (dayAssignedSet.has(name)) score += 45; // no queremos saturar el mismo día

  if (s.lastAssignedDayIndex != null) {
    const delta = Math.abs(dayIndex - s.lastAssignedDayIndex);
    if (delta === 0) score += 60;
    if (delta === 1) score += 18; // evitar seguidos
  }

  // Ruido pequeño determinista para desempate
  score += rng() * 0.5;
  return score;
}

function pickBestCandidate(args) {
  const { candidates } = args;
  let best = null;
  let bestScore = Number.POSITIVE_INFINITY;
  for (const name of candidates) {
    const s = scoreCandidate({ ...args, name });
    if (s < bestScore) {
      best = name;
      bestScore = s;
    }
  }
  return best;
}

function generateEquitableAssignments(schedule, peopleModel) {
  const weekStart = parseISODate(schedule.weekStart);
  const week = buildWeek(weekStart);
  const rng = mulberry32(hashStringToSeed(schedule.weekStart));
  const stats = initStats(peopleModel.all);

  // Reset solo casillas en modo NORMAL (no tocamos TODOS)
  for (const slot of scheduleToRows(schedule)) {
    if (slot.modo !== "TODOS") slot.asignadoA = "";
  }

  // Primero asignamos FIJO, luego BACKUP para minimizar choques en el mismo día/franja.
  for (let dayIndex = 0; dayIndex < week.length; dayIndex++) {
    const day = week[dayIndex];
    const dayAssigned = new Set();

    for (const franja of FRANJAS) {
      if (isTodosDayFranja(schedule, day.iso, franja.key)) continue;

      // FIJO
      {
        const candidates = availablePeopleForDate(peopleModel, day.iso);
        const chosen = pickBestCandidate({
          stats,
          candidates,
          iso: day.iso,
          dayIndex,
          franjaKey: franja.key,
          tipoKey: "FIJO",
          dayAssignedSet: dayAssigned,
          rng,
        });
        if (chosen) {
          const id = slotId(day.iso, franja.key, "FIJO");
          setSlot(schedule, id, { asignadoA: chosen });
          dayAssigned.add(chosen);
          recordAssignment(stats, chosen, day.iso, dayIndex, franja.key, "FIJO");
        }
      }

      // BACKUP (intentamos que no sea la misma persona que FIJO)
      {
        const fijoId = slotId(day.iso, franja.key, "FIJO");
        const fijoName = normalizeName(schedule.slots[fijoId]?.asignadoA);
        let candidates = availablePeopleForDate(peopleModel, day.iso);
        candidates = candidates.filter((n) => normalizeName(n) !== fijoName);
        if (!candidates.length) candidates = availablePeopleForDate(peopleModel, day.iso); // fallback

        const chosen = pickBestCandidate({
          stats,
          candidates,
          iso: day.iso,
          dayIndex,
          franjaKey: franja.key,
          tipoKey: "BACKUP",
          dayAssignedSet: dayAssigned,
          rng,
        });
        if (chosen) {
          const id = slotId(day.iso, franja.key, "BACKUP");
          setSlot(schedule, id, { asignadoA: chosen });
          dayAssigned.add(chosen);
          recordAssignment(stats, chosen, day.iso, dayIndex, franja.key, "BACKUP");
        }
      }
    }
  }

  // Señales de salud: comprobar huecos
  const missing = scheduleToRows(schedule).filter((s) => s.modo !== "TODOS" && !normalizeName(s.asignadoA)).length;
  return { missing };
}

function apiBase() {
  const url = clampStr(qs("webAppUrl").value);
  if (!url) return null;
  return url.replace(/\/+$/, "");
}

async function apiCall(action, payload) {
  const base = apiBase();
  if (!base) throw new Error("Falta Web App URL.");
  // Apps Script + GitHub Pages: evitamos preflight usando un Content-Type "simple".
  const res = await fetch(`${base}?action=${encodeURIComponent(action)}`, {
    method: "POST",
    headers: { "content-type": "text/plain;charset=UTF-8" },
    body: JSON.stringify(payload || {}),
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`Error ${res.status}: ${txt || res.statusText}`);
  }
  const data = await res.json();
  if (!data?.ok) throw new Error(data?.error || "Error desconocido del servidor");
  return data;
}

function getExcludedSet() {
  const cfg = localConfigGet();
  const arr = Array.isArray(cfg.excludedPeople) ? cfg.excludedPeople : [];
  return new Set(arr.map(normalizeName).filter(Boolean));
}

function setExcludedPeople(nextNames) {
  const cfg = localConfigGet();
  cfg.excludedPeople = sortNames(unique(nextNames.map(normalizeName).filter(Boolean)));
  localConfigSet(cfg);
}

function renderPeopleList(peopleModel) {
  const wrap = qs("peopleList");
  wrap.innerHTML = "";
  const excluded = getExcludedSet();
  for (const p of peopleModel.all) {
    const chip = document.createElement("div");
    chip.className = `personChip${excluded.has(normalizeName(p)) ? " off" : ""}`;
    chip.textContent = p;
    wrap.appendChild(chip);
  }
}

function renderExcludeControl(peopleModel) {
  const el = document.getElementById("excludePeople");
  if (!el) return;

  const excluded = getExcludedSet();
  el.innerHTML = "";
  for (const name of peopleModel.all) {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    opt.selected = excluded.has(normalizeName(name));
    el.appendChild(opt);
  }

  el.onchange = () => {
    const selected = Array.from(el.selectedOptions).map((o) => o.value);
    setExcludedPeople(selected);
    renderAll(window.__schedule, window.__peopleModel);
  };
}

function renderVacationsThisWeek(vacationEvents) {
  const el = document.getElementById("vacationsThisWeek");
  if (!el) return;
  const ev = Array.isArray(vacationEvents) ? vacationEvents : [];
  if (!ev.length) {
    el.textContent = "—";
    return;
  }
  el.innerHTML = ev
    .map((x) => `<div><strong>${escapeHtml(x.iso)}</strong>: ${escapeHtml((x.names || []).join(", "))}</div>`)
    .join("");
}

function renderImagePreview(lastImage) {
  const el = document.getElementById("imagePreview");
  if (!el) return;
  if (!lastImage?.previewUrl) {
    el.textContent = "—";
    return;
  }
  el.innerHTML = `
    <img src="${escapeHtml(lastImage.previewUrl)}" alt="Imagen de turnos" />
    <div class="imageActions">
      <a class="linkBtn" href="${escapeHtml(lastImage.fileUrl || lastImage.previewUrl)}" target="_blank" rel="noreferrer">Abrir en Drive</a>
      <span class="muted">${escapeHtml(lastImage.fileName || "")}</span>
    </div>
  `;
}

function renderMiniCalendar(weekStartISO) {
  const el = document.getElementById("miniCalendar");
  if (!el) return;
  const ws = parseISODate(weekStartISO);
  if (!ws) return;
  const week = buildWeek(ws);
  const inWeek = new Set(week.map((d) => d.iso));

  const year = ws.getFullYear();
  const month = ws.getMonth();
  const first = new Date(year, month, 1);
  const last = new Date(year, month + 1, 0);
  const startDow = (first.getDay() + 6) % 7; // lunes=0

  const cells = [];
  for (let i = 0; i < startDow; i++) cells.push({ empty: true });
  for (let d = 1; d <= last.getDate(); d++) {
    const dt = new Date(year, month, d);
    const iso = toISODate(dt);
    cells.push({ day: d, iso, inWeek: inWeek.has(iso) });
  }
  while (cells.length % 7 !== 0) cells.push({ empty: true });

  const heads = ["L", "M", "X", "J", "V", "S", "D"];
  el.innerHTML =
    heads.map((h) => `<div class="calHead">${h}</div>`).join("") +
    cells
      .map((c) => {
        if (c.empty) return `<div class="calCell off"></div>`;
        const isThu = week[0].iso === c.iso;
        return `<div class="calCell${c.inWeek ? " inWeek" : ""}${isThu ? " isThu" : ""}">${c.day}</div>`;
      })
      .join("");
}

function renderScheduleTable(schedule, peopleModel) {
  const weekStart = parseISODate(schedule.weekStart);
  const week = buildWeek(weekStart);
  const tbody = qs("scheduleBody");
  tbody.innerHTML = "";

  for (const day of week) {
    for (const tipo of TIPOS) {
      const tr = document.createElement("tr");

      if (tipo.key === "FIJO") {
        const tdDay = document.createElement("td");
        tdDay.className = "dayCell";
        tdDay.rowSpan = 2;
        tdDay.innerHTML = `${day.label} <span class="daySub">${day.iso}</span>`;
        tr.appendChild(tdDay);
      }

      const tdTipo = document.createElement("td");
      const pill = document.createElement("span");
      pill.className = `typePill ${tipo.key === "FIJO" ? "fijo" : "backup"}`;
      pill.textContent = tipo.label;
      tdTipo.appendChild(pill);
      tr.appendChild(tdTipo);

      for (const franja of FRANJAS) {
        const td = document.createElement("td");

        const isTodos = isTodosDayFranja(schedule, day.iso, franja.key);
        if (isTodos) {
          const d = document.createElement("div");
          d.className = "slotTodos";
          d.textContent = "TODOS";
          td.appendChild(d);
        } else {
          const id = slotId(day.iso, franja.key, tipo.key);
          const sel = document.createElement("select");
          sel.className = "slotSelect";
          sel.dataset.slotId = id;

          const available = availablePeopleForDate(peopleModel, day.iso);
          sel.appendChild(buildOption("—", ""));
          for (const name of available) sel.appendChild(buildOption(name, name));
          sel.value = schedule.slots[id]?.asignadoA || "";

          sel.addEventListener("change", () => {
            setSlot(schedule, id, { asignadoA: sel.value });
            renderSummary(schedule);
            refreshChangeSlotOptions(schedule);
            if (typeof window.__autoSaveDebounced === "function") window.__autoSaveDebounced("edición");
          });

          td.appendChild(sel);
        }

        tr.appendChild(td);
      }

      tbody.appendChild(tr);
    }
  }
}

function renderSummary(schedule) {
  const rows = computeSummary(schedule);
  const body = qs("summaryBody");
  body.innerHTML = "";
  for (const r of rows) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(r.name)}</td>
      <td class="num">${r.fijo}</td>
      <td class="num">${r.backup}</td>
      <td class="num">${r.total}</td>
    `;
    body.appendChild(tr);
  }
}

function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));
}

function refreshChangeSlotOptions(schedule) {
  const slotSel = qs("changeSlot");
  const prev = slotSel.value;
  slotSel.innerHTML = "";

  const weekStart = parseISODate(schedule.weekStart);
  const week = buildWeek(weekStart);
  const allSlots = [];
  for (const day of week) {
    for (const franja of FRANJAS) {
      const isTodos = isTodosDayFranja(schedule, day.iso, franja.key);
      if (isTodos) continue;
      for (const tipo of TIPOS) {
        const id = slotId(day.iso, franja.key, tipo.key);
        allSlots.push({ id, day, franja, tipo });
      }
    }
  }

  for (const s of allSlots) {
    const label = `${s.day.label} ${s.day.iso} · ${s.franja.label} · ${s.tipo.label}`;
    slotSel.appendChild(buildOption(label, s.id));
  }

  slotSel.value = allSlots.some((x) => x.id === prev) ? prev : (allSlots[0]?.id || "");
  refreshChangePeopleOptions(schedule);
}

function refreshChangePeopleOptions(schedule) {
  const slotSel = qs("changeSlot");
  const fromSel = qs("changeFrom");
  const toSel = qs("changeTo");
  const { iso } = splitSlotId(slotSel.value || "");

  const names = sortNames(unique([...DEFAULT_PEOPLE, ...window.__peopleModel.all]).map(normalizeName).filter(Boolean));
  const available = iso ? availablePeopleForDate(window.__peopleModel, iso) : names;

  const current = slotSel.value ? (schedule.slots[slotSel.value]?.asignadoA || "") : "";

  fromSel.innerHTML = "";
  toSel.innerHTML = "";
  fromSel.appendChild(buildOption("—", ""));
  toSel.appendChild(buildOption("—", ""));
  for (const n of available) {
    fromSel.appendChild(buildOption(n, n));
    toSel.appendChild(buildOption(n, n));
  }

  fromSel.value = current || "";
  if (current) {
    const other = available.find((x) => x !== current) || "";
    toSel.value = other;
  }
}

function renderChangesHistory(changes) {
  const body = qs("changesBody");
  body.innerHTML = "";
  for (const c of changes) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(c.timestamp)}</td>
      <td>${escapeHtml(c.slotLabel)}</td>
      <td>${escapeHtml(c.antes || "—")}</td>
      <td>${escapeHtml(c.despues || "—")}</td>
      <td>${escapeHtml(c.motivo || "")}</td>
    `;
    body.appendChild(tr);
  }
}

function updateMeta(schedule) {
  const weekStart = parseISODate(schedule.weekStart);
  const week = buildWeek(weekStart);
  qs("metaWeek").textContent = `${week[0].iso} → ${week[6].iso}`;
  qs("metaGeneratedAt").textContent = new Date().toLocaleString("es-ES");
  qs("subtitle").textContent = `${formatHumanDate(week[0].date)} · hasta · ${formatHumanDate(week[6].date)}`;
}

function ensureWeekStartIsThursday(iso) {
  const d = parseISODate(iso);
  if (!d) return iso;
  const thu = computeWeekStartThursday(d);
  return toISODate(thu);
}

async function loadBootstrap(schedule) {
  const sheetId = clampStr(qs("sheetId").value);
  if (!sheetId) {
    setStatus("Falta Google Sheet ID. Puedes usar modo local por ahora.", "warn");
    return { people: DEFAULT_PEOPLE, vacationsByISO: {}, vacationEvents: [], changes: [], scheduleFromServer: null };
  }

  try {
    const data = await apiCall("getBootstrap", {
      sheetId,
      weekStart: schedule.weekStart,
    });

    return {
      people: (data.people || []).map(normalizeName).filter(Boolean),
      vacationsByISO: data.vacationsByISO || {},
      vacationEvents: data.vacationEvents || [],
      changes: data.changes || [],
      scheduleFromServer: data.schedule || null,
      config: data.config || null,
    };
  } catch (e) {
    setStatus(`No se pudo cargar de Sheets: ${e.message}. Usando modo local.`, "warn");
    return { people: DEFAULT_PEOPLE, vacationsByISO: {}, vacationEvents: [], changes: [], scheduleFromServer: null };
  }
}

function scheduleFromServerPayload(weekStartISO, payload) {
  const schedule = createEmptySchedule(weekStartISO);
  if (!payload?.rows?.length) return schedule;
  for (const r of payload.rows) {
    const id = slotId(r.fecha, r.franja, r.tipo);
    if (!schedule.slots[id]) continue;
    schedule.slots[id] = {
      ...schedule.slots[id],
      modo: r.modo || (r.tipo === "TODOS" ? "TODOS" : "NORMAL"),
      asignadoA: r.asignadoA || "",
      nota: r.nota || "",
    };
  }
  return schedule;
}

function wireTodosToggles(schedule) {
  // UX: doble click sobre cabecera de columna (fila FIJO) para marcar TODOS para ese día+franja.
  // Implementado como doble click sobre la celda del select (fila FIJO) cuando está en NORMAL.
  const tbody = qs("scheduleBody");
  tbody.addEventListener("dblclick", (ev) => {
    const sel = ev.target?.closest?.("select.slotSelect");
    if (!sel) return;
    const id = sel.dataset.slotId;
    if (!id) return;
    const { iso, franja } = splitSlotId(id);
    const franjaKey = franja;
    const on = !isTodosDayFranja(schedule, iso, franjaKey);
    setTodosForDayFranja(schedule, iso, franjaKey, on);
    renderAll(schedule, window.__peopleModel);
    setStatus(on ? "Marcado como TODOS (doble click para revertir)." : "TODOS desactivado.", "ok");
  });
}

function renderAll(schedule, peopleModel) {
  window.__schedule = schedule;
  window.__peopleModel = peopleModel;
  updateMeta(schedule);
  renderPeopleList(peopleModel);
  renderExcludeControl(peopleModel);
  renderMiniCalendar(schedule.weekStart);
  renderScheduleTable(schedule, peopleModel);
  renderSummary(schedule);
  refreshChangeSlotOptions(schedule);
}

function buildSlotLabel(id) {
  const { iso, franja, tipo } = splitSlotId(id);
  const d = parseISODate(iso);
  const dayLabel = d
    ? d.toLocaleDateString("es-ES", { weekday: "long" })
    : iso;
  const f = FRANJAS.find((x) => x.key === franja)?.label || franja;
  const t = TIPOS.find((x) => x.key === tipo)?.label || tipo;
  return `${capitalize(dayLabel)} ${iso} · ${f} · ${t}`;
}

function capitalize(s) {
  const str = String(s || "");
  return str ? str[0].toUpperCase() + str.slice(1) : str;
}

async function saveToSheets(schedule, changes) {
  const sheetId = clampStr(qs("sheetId").value);
  if (!sheetId) throw new Error("Falta Google Sheet ID.");
  await apiCall("saveWeek", {
    sheetId,
    weekStart: schedule.weekStart,
    rows: scheduleToRows(schedule),
    changes,
  });
}

async function saveImageToDrive(schedule) {
  const sheetId = clampStr(qs("sheetId").value);
  if (!sheetId) throw new Error("Falta Google Sheet ID.");

  const capture = qs("scheduleCapture");
  const canvas = await html2canvas(capture, {
    backgroundColor: null,
    scale: Math.min(2, window.devicePixelRatio || 1),
    useCORS: true,
  });
  const dataUrl = canvas.toDataURL("image/png");
  const filename = `Turnos_${schedule.weekStart}_${todayISO()}.png`;

  const data = await apiCall("uploadPng", {
    sheetId,
    weekStart: schedule.weekStart,
    filename,
    dataUrl,
  });
  return data;
}

function init() {
  const cfg = localConfigGet();
  const weekStart = ensureWeekStartIsThursday(cfg.weekStart || toISODate(computeWeekStartThursday(new Date())));
  qs("weekStart").value = weekStart;
  // Si se abre desde un Web App de Apps Script, podemos autodetectar la URL base.
  const currentBase = `${window.location.origin}${window.location.pathname}`.replace(/\/+$/, "");
  qs("webAppUrl").value = cfg.webAppUrl || (window.location.pathname.includes("/exec") ? currentBase : "");
  qs("sheetId").value = cfg.sheetId || "";

  let schedule = createEmptySchedule(weekStart);
  let changes = [];
  let lastImage = null;

  const autoSave = async (reason) => {
    const hasIntegration = Boolean(clampStr(qs("webAppUrl").value) && clampStr(qs("sheetId").value));
    if (!hasIntegration) return;
    try {
      setStatus(`Guardando… (${reason})`, "muted");
      await saveToSheets(schedule, changes);
      setStatus("Guardado.", "ok");
    } catch (e) {
      setStatus(`No se pudo guardar: ${e.message}`, "bad");
    }
  };
  const autoSaveDebounced = (() => {
    let t = null;
    return (reason) => {
      if (t) clearTimeout(t);
      t = setTimeout(() => autoSave(reason), 900);
    };
  })();
  window.__autoSaveDebounced = autoSaveDebounced;

  const updateIntegrationButtons = () => {
    const hasIntegration = Boolean(clampStr(qs("webAppUrl").value) && clampStr(qs("sheetId").value));
    qs("btnLoad").disabled = !hasIntegration;
    qs("btnSave").disabled = !hasIntegration;
    qs("btnSaveImage").disabled = !hasIntegration;
    qs("localHint").innerHTML = hasIntegration
      ? "Integración activa: puedes cargar/guardar en Sheets y subir imagen a Drive."
      : "Empieza en local: elige jueves y pulsa <strong>Generar</strong>. La conexión con Google la puedes activar después.";
  };

  const persist = () => {
    localConfigSet({
      weekStart: qs("weekStart").value,
      webAppUrl: qs("webAppUrl").value,
      sheetId: qs("sheetId").value,
    });
    updateIntegrationButtons();
  };

  const hardReload = async () => {
    persist();
    const iso = ensureWeekStartIsThursday(qs("weekStart").value);
    qs("weekStart").value = iso;
    schedule = createEmptySchedule(iso);
    changes = [];

    setStatus("Cargando datos…", "muted");
    const boot = await loadBootstrap(schedule);
    const peopleModel = makePeopleModel(boot.people, boot.vacationsByISO);

    if (boot.scheduleFromServer) {
      schedule = scheduleFromServerPayload(iso, boot.scheduleFromServer);
    }
    changes = (boot.changes || []).map((c) => ({
      ...c,
      slotLabel: c.slotLabel || buildSlotLabel(c.slotId || ""),
    }));

    renderAll(schedule, peopleModel);
    renderVacationsThisWeek(boot.vacationEvents);
    renderImagePreview(lastImage);
    wireTodosToggles(schedule);
    renderChangesHistory(changes);
    setStatus("Cargado.", "ok");
  };

  qs("btnLoad").addEventListener("click", () => hardReload());
  qs("weekStart").addEventListener("change", () => hardReload());
  qs("webAppUrl").addEventListener("change", persist);
  qs("sheetId").addEventListener("change", persist);

  qs("changeSlot").addEventListener("change", () => refreshChangePeopleOptions(schedule));

  const applyChange = async () => {
    const slotIdVal = qs("changeSlot").value;
    if (!slotIdVal) return;
    const mode = qs("changeMode").value;
    const from = normalizeName(qs("changeFrom").value);
    const to = normalizeName(qs("changeTo").value);
    const motivo = clampStr(qs("changeReason").value);

    const before = schedule.slots[slotIdVal]?.asignadoA || "";
    if (mode === "REASSIGN") {
      if (!to) {
        setStatus("Selecciona a quién reasignar.", "warn");
        return;
      }
      setSlot(schedule, slotIdVal, { asignadoA: to });
      changes.unshift({
        timestamp: new Date().toLocaleString("es-ES"),
        slotId: slotIdVal,
        slotLabel: buildSlotLabel(slotIdVal),
        antes: before || "—",
        despues: to,
        motivo,
      });
    } else {
      if (!from || !to) {
        setStatus("Selecciona 'De' y 'A' para intercambiar.", "warn");
        return;
      }
      // SWAP: slot seleccionada intercambia con otra casilla donde esté "to" dentro de la semana y mismo tipo/franja
      const { franja, tipo } = splitSlotId(slotIdVal);
      const weekSlots = Object.values(schedule.slots).filter(
        (s) => s.weekStart === schedule.weekStart && s.franja === franja && s.tipo === tipo && s.modo !== "TODOS",
      );
      const other = weekSlots.find((s) => normalizeName(s.asignadoA) === to);
      if (!other) {
        setStatus(`No encuentro otra casilla con ${to} en la misma franja/tipo para intercambiar.`, "warn");
        return;
      }
      const aId = slotIdVal;
      const bId = other.id;
      const aName = schedule.slots[aId]?.asignadoA || "";
      const bName = schedule.slots[bId]?.asignadoA || "";
      setSlot(schedule, aId, { asignadoA: bName });
      setSlot(schedule, bId, { asignadoA: aName });
      changes.unshift({
        timestamp: new Date().toLocaleString("es-ES"),
        slotId: aId,
        slotLabel: `${buildSlotLabel(aId)} ↔ ${buildSlotLabel(bId)}`,
        antes: `${aName || "—"} / ${bName || "—"}`,
        despues: `${bName || "—"} / ${aName || "—"}`,
        motivo,
      });
    }

    renderAll(schedule, window.__peopleModel);
    renderChangesHistory(changes);
    setStatus("Cambio aplicado.", "ok");
    autoSaveDebounced("cambio");
  };

  // Botón legacy (si existiera en versiones anteriores)
  const maybeApply = document.getElementById("btnApplyChange");
  if (maybeApply) maybeApply.addEventListener("click", applyChange);
  qs("btnApplyChange2").addEventListener("click", applyChange);

  qs("btnSave").addEventListener("click", async () => {
    try {
      setStatus("Guardando en Sheets…", "muted");
      await saveToSheets(schedule, changes);
      setStatus("Guardado en Sheets.", "ok");
    } catch (e) {
      setStatus(`No se pudo guardar: ${e.message}`, "bad");
    }
  });

  qs("btnSaveImage").addEventListener("click", async () => {
    if (qs("btnSaveImage").disabled) {
      setStatus("Configura Web App URL + Sheet ID para guardar imagen en Drive.", "warn");
      return;
    }
    try {
      setStatus("Generando PNG y subiendo a Drive…", "muted");
      const data = await saveImageToDrive(schedule);
      lastImage = data;
      renderImagePreview(lastImage);
      setStatus(`Imagen subida a Drive. Archivo: ${data.fileName}`, "ok");
      if (data.fileUrl) {
        // Mostrar enlace de forma sencilla sin romper el layout
        const msg = `Imagen subida a Drive: ${data.fileName}`;
        qs("status").textContent = msg;
        qs("status").title = data.fileUrl;
      }
    } catch (e) {
      setStatus(`No se pudo guardar imagen: ${e.message}`, "bad");
    }
  });

  qs("btnGenerate").addEventListener("click", () => {
    try {
      const res = generateEquitableAssignments(schedule, window.__peopleModel);
      renderAll(schedule, window.__peopleModel);
      renderChangesHistory(changes);
      if (res.missing) {
        setStatus(`Turnos generados con ${res.missing} huecos (por falta de disponibles).`, "warn");
      } else {
        setStatus("Turnos generados.", "ok");
      }
      autoSaveDebounced("auto");
    } catch (e) {
      setStatus(`No se pudo generar: ${e.message}`, "bad");
    }
  });

  updateIntegrationButtons();
  hardReload().catch((e) => setStatus(`Error: ${e.message}`, "bad"));
}

document.addEventListener("DOMContentLoaded", init);

