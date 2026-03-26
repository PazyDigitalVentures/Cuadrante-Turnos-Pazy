/* Copia del JS del frontend (se mantiene igual que app.js) */
/* global html2canvas */

// Pegado del contenido de turnos-pazy/app.js
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
          modo: "NORMAL",
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

function availablePeopleForDate(peopleModel, iso) {
  const blocked = new Set((peopleModel.vacationsByISO?.[iso] || []).map(normalizeName));
  const excluded = getExcludedSet();
  return peopleModel.all.filter((n) => !blocked.has(n) && !excluded.has(normalizeName(n)));
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
  const dow = d.getDay();
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
      lastAssignedDayIndex: null,
    });
  }
  return stats;
}

function recordAssignment(stats, name, dayIndex, franjaKey, tipoKey, iso) {
  const s = stats.get(name);
  if (!s) return;
  s.total += 1;
  if (tipoKey === "FIJO") s.fijo += 1;
  if (tipoKey === "BACKUP") s.backup += 1;
  if (s.byFranja[franjaKey] != null) s.byFranja[franjaKey] += 1;
  if (isWeekendISO(iso)) s.weekend += 1;
  else s.weekday += 1;
  s.lastAssignedDayIndex = dayIndex;
}

function scoreCandidate({ stats, name, iso, dayIndex, franjaKey, tipoKey, dayAssignedSet, rng }) {
  const s = stats.get(name);
  if (!s) return Number.POSITIVE_INFINITY;
  const weekend = isWeekendISO(iso);
  let score = 0;
  score += s.total * 10;
  score += (tipoKey === "FIJO" ? s.fijo : s.backup) * 12;
  score += (s.byFranja[franjaKey] || 0) * 8;
  score += (weekend ? s.weekend : s.weekday) * 7;
  if (dayAssignedSet.has(name)) score += 45;
  if (s.lastAssignedDayIndex != null) {
    const delta = Math.abs(dayIndex - s.lastAssignedDayIndex);
    if (delta === 0) score += 60;
    if (delta === 1) score += 18;
  }
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

  for (const slot of scheduleToRows(schedule)) {
    if (slot.modo !== "TODOS") slot.asignadoA = "";
  }

  for (let dayIndex = 0; dayIndex < week.length; dayIndex++) {
    const day = week[dayIndex];
    const dayAssigned = new Set();
    for (const franja of FRANJAS) {
      if (isTodosDayFranja(schedule, day.iso, franja.key)) continue;
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
          recordAssignment(stats, chosen, dayIndex, franja.key, "FIJO", day.iso);
        }
      }
      {
        const fijoId = slotId(day.iso, franja.key, "FIJO");
        const fijoName = normalizeName(schedule.slots[fijoId]?.asignadoA);
        let candidates = availablePeopleForDate(peopleModel, day.iso).filter((n) => normalizeName(n) !== fijoName);
        if (!candidates.length) candidates = availablePeopleForDate(peopleModel, day.iso);
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
          recordAssignment(stats, chosen, dayIndex, franja.key, "BACKUP", day.iso);
        }
      }
    }
  }

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
  const res = await fetch(`${base}?action=${encodeURIComponent(action)}`, {
    method: "POST",
    headers: { "content-type": "application/json" },
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
    setStatus("Falta Google Sheet ID.", "warn");
    return { people: DEFAULT_PEOPLE, vacationsByISO: {}, vacationEvents: [], changes: [], scheduleFromServer: null };
  }
  try {
    const data = await apiCall("getBootstrap", { sheetId, weekStart: schedule.weekStart });
    return {
      people: (data.people || []).map(normalizeName).filter(Boolean),
      vacationsByISO: data.vacationsByISO || {},
      vacationEvents: data.vacationEvents || [],
      changes: data.changes || [],
      scheduleFromServer: data.schedule || null,
      config: data.config || null,
    };
  } catch (e) {
    setStatus(`No se pudo cargar: ${e.message}`, "warn");
    return { people: DEFAULT_PEOPLE, vacationsByISO: {}, vacationEvents: [], changes: [], scheduleFromServer: null };
  }
}

function scheduleFromServerPayload(weekStartISO, payload) {
  const schedule = createEmptySchedule(weekStartISO);
  if (!payload?.rows?.length) return schedule;
  for (const r of payload.rows) {
    const id = slotId(r.fecha, r.franja, r.tipo);
    if (!schedule.slots[id]) continue;
    schedule.slots[id] = { ...schedule.slots[id], modo: r.modo || "NORMAL", asignadoA: r.asignadoA || "", nota: r.nota || "" };
  }
  return schedule;
}

function wireTodosToggles(schedule) {
  const tbody = qs("scheduleBody");
  tbody.addEventListener("dblclick", (ev) => {
    const sel = ev.target?.closest?.("select.slotSelect");
    if (!sel) return;
    const id = sel.dataset.slotId;
    if (!id) return;
    const { iso, franja } = splitSlotId(id);
    const on = !isTodosDayFranja(schedule, iso, franja);
    setTodosForDayFranja(schedule, iso, franja, on);
    renderAll(schedule, window.__peopleModel);
  });
}

function renderAll(schedule, peopleModel) {
  window.__schedule = schedule;
  window.__peopleModel = peopleModel;
  updateMeta(schedule);
  renderPeopleList(peopleModel);
  renderExcludeControl(peopleModel);
  renderScheduleTable(schedule, peopleModel);
  renderSummary(schedule);
  refreshChangeSlotOptions(schedule);
}

function buildSlotLabel(id) {
  const { iso, franja, tipo } = splitSlotId(id);
  const d = parseISODate(iso);
  const dayLabel = d ? d.toLocaleDateString("es-ES", { weekday: "long" }) : iso;
  const f = FRANJAS.find((x) => x.key === franja)?.label || franja;
  const t = TIPOS.find((x) => x.key === tipo)?.label || tipo;
  return `${dayLabel} ${iso} · ${f} · ${t}`;
}

async function saveToSheets(schedule, changes) {
  const sheetId = clampStr(qs("sheetId").value);
  if (!sheetId) throw new Error("Falta Google Sheet ID.");
  await apiCall("saveWeek", { sheetId, weekStart: schedule.weekStart, rows: scheduleToRows(schedule), changes });
}

async function saveImageToDrive(schedule) {
  const sheetId = clampStr(qs("sheetId").value);
  if (!sheetId) throw new Error("Falta Google Sheet ID.");
  const capture = qs("scheduleCapture");
  const canvas = await html2canvas(capture, { backgroundColor: null, scale: Math.min(2, window.devicePixelRatio || 1), useCORS: true });
  const dataUrl = canvas.toDataURL("image/png");
  const filename = `Turnos_${schedule.weekStart}_${todayISO()}.png`;
  return await apiCall("uploadPng", { sheetId, weekStart: schedule.weekStart, filename, dataUrl });
}

function init() {
  const cfg = localConfigGet();
  const weekStart = ensureWeekStartIsThursday(cfg.weekStart || toISODate(computeWeekStartThursday(new Date())));
  qs("weekStart").value = weekStart;
  const currentBase = `${window.location.origin}${window.location.pathname}`.replace(/\/+$/, "");
  qs("webAppUrl").value = cfg.webAppUrl || (window.location.pathname.includes("/exec") ? currentBase : "");
  qs("sheetId").value = cfg.sheetId || "";

  let schedule = createEmptySchedule(weekStart);
  let changes = [];

  const updateIntegrationButtons = () => {
    const hasIntegration = Boolean(clampStr(qs("webAppUrl").value) && clampStr(qs("sheetId").value));
    qs("btnLoad").disabled = !hasIntegration;
    qs("btnSave").disabled = !hasIntegration;
    qs("btnSaveImage").disabled = !hasIntegration;
  };

  const persist = () => {
    localConfigSet({ weekStart: qs("weekStart").value, webAppUrl: qs("webAppUrl").value, sheetId: qs("sheetId").value });
    updateIntegrationButtons();
  };

  const hardReload = async () => {
    persist();
    const iso = ensureWeekStartIsThursday(qs("weekStart").value);
    qs("weekStart").value = iso;
    schedule = createEmptySchedule(iso);
    changes = [];
    setStatus("Cargando…", "muted");
    const boot = await loadBootstrap(schedule);
    const peopleModel = makePeopleModel(boot.people, boot.vacationsByISO);
    if (boot.scheduleFromServer) schedule = scheduleFromServerPayload(iso, boot.scheduleFromServer);
    changes = (boot.changes || []).map((c) => ({ ...c, slotLabel: c.slotLabel || buildSlotLabel(c.slotId || "") }));
    renderAll(schedule, peopleModel);
    renderVacationsThisWeek(boot.vacationEvents);
    wireTodosToggles(schedule);
    renderChangesHistory(changes);
    setStatus("Listo.", "ok");
  };

  qs("btnLoad").addEventListener("click", () => hardReload());
  qs("weekStart").addEventListener("change", () => hardReload());
  qs("webAppUrl").addEventListener("change", persist);
  qs("sheetId").addEventListener("change", persist);
  qs("changeSlot").addEventListener("change", () => refreshChangePeopleOptions(schedule));

  qs("btnApplyChange2").addEventListener("click", async () => {
    const slotIdVal = qs("changeSlot").value;
    if (!slotIdVal) return;
    const mode = qs("changeMode").value;
    const from = normalizeName(qs("changeFrom").value);
    const to = normalizeName(qs("changeTo").value);
    const motivo = clampStr(qs("changeReason").value);
    const before = schedule.slots[slotIdVal]?.asignadoA || "";
    if (mode === "REASSIGN") {
      if (!to) return;
      setSlot(schedule, slotIdVal, { asignadoA: to });
      changes.unshift({ timestamp: new Date().toLocaleString("es-ES"), slotId: slotIdVal, slotLabel: buildSlotLabel(slotIdVal), antes: before || "—", despues: to, motivo });
    }
    renderAll(schedule, window.__peopleModel);
    renderChangesHistory(changes);
  });

  qs("btnSave").addEventListener("click", async () => {
    try { setStatus("Guardando…", "muted"); await saveToSheets(schedule, changes); setStatus("Guardado.", "ok"); }
    catch (e) { setStatus(`Error: ${e.message}`, "bad"); }
  });

  qs("btnSaveImage").addEventListener("click", async () => {
    try { setStatus("Subiendo imagen…", "muted"); const data = await saveImageToDrive(schedule); setStatus(`Imagen subida: ${data.fileName}`, "ok"); }
    catch (e) { setStatus(`Error: ${e.message}`, "bad"); }
  });

  qs("btnGenerate").addEventListener("click", () => {
    const res = generateEquitableAssignments(schedule, window.__peopleModel);
    renderAll(schedule, window.__peopleModel);
    renderChangesHistory(changes);
    setStatus(res.missing ? `Generado con ${res.missing} huecos.` : "Turnos generados.", res.missing ? "warn" : "ok");
  });

  updateIntegrationButtons();
  hardReload();
}

document.addEventListener("DOMContentLoaded", init);

