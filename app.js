/* global html2canvas */

const STORAGE_KEY = "turnosPazy.localState.v2";
const DEFAULT_PEOPLE = [
  "Georgi Valeriev",
  "Magui Cerda",
  "Antonella Sipan",
  "Inigo Puyol",
  "Luz Romero",
  "Patricia Lopez",
  "Jorge Romera",
  "Irene Penalosa",
  "Maria Jose Rubio",
  "Alessandra Solis",
  "Adrian Garces",
  "Ignacio Rivas",
  "Alonso Garcia",
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

const TIPOS = [{ key: "FIJO", label: "Fijo" }, { key: "BACKUP", label: "Back-up" }];

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

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function saveState(next) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
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
  const excluded = getExcludedSet(window.__state || {});
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

function getExcludedSet(state) {
  const arr = Array.isArray(state.excludedPeople) ? state.excludedPeople : [];
  return new Set(arr.map(normalizeName).filter(Boolean));
}

function setExcludedPeople(state, nextNames) {
  state.excludedPeople = sortNames(unique(nextNames.map(normalizeName).filter(Boolean)));
}

function renderPeopleList(peopleModel, state) {
  const wrap = qs("peopleList");
  wrap.innerHTML = "";
  const excluded = getExcludedSet(state);
  for (const p of peopleModel.all) {
    const chip = document.createElement("div");
    chip.className = `personChip${excluded.has(normalizeName(p)) ? " off" : ""}`;
    chip.textContent = p;
    wrap.appendChild(chip);
  }
}

function renderExcludeControl(peopleModel, state, onChange) {
  const el = document.getElementById("excludePeople");
  if (!el) return;

  const excluded = getExcludedSet(state);
  el.innerHTML = "";
  for (const name of peopleModel.all) {
    const row = document.createElement("label");
    row.className = "checkItem";
    const chk = document.createElement("input");
    chk.type = "checkbox";
    chk.value = name;
    chk.checked = excluded.has(normalizeName(name));
    chk.addEventListener("change", () => {
      const selected = Array.from(el.querySelectorAll("input[type='checkbox']:checked")).map((x) => x.value);
      setExcludedPeople(state, selected);
      if (chk.checked) {
        const vacSel = document.getElementById("vacPerson");
        if (vacSel) vacSel.value = name;
      }
      onChange();
    });
    row.appendChild(chk);
    const txt = document.createElement("span");
    txt.textContent = name;
    row.appendChild(txt);
    el.appendChild(row);
  }
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
      <span class="linkBtn">${escapeHtml(lastImage.githubPath || "")}</span>
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

  const offset = Number(window.__state?.calendarMonthOffset || 0);
  const target = new Date(ws.getFullYear(), ws.getMonth() + offset, 1);
  const year = target.getFullYear();
  const month = target.getMonth();
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
  const monthLabel = document.getElementById("monthLabel");
  if (monthLabel) monthLabel.textContent = target.toLocaleDateString("es-ES", { month: "long", year: "numeric" });
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

function renderScheduleTable(schedule, peopleModel, onSelectChange) {
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

        const id = slotId(day.iso, franja.key, tipo.key);
        const sel = document.createElement("select");
        sel.className = "slotSelect";
        sel.dataset.slotId = id;
        const available = availablePeopleForDate(peopleModel, day.iso);
        sel.appendChild(buildOption("—", ""));
        sel.appendChild(buildOption("TODOS", "__TODOS__"));
        for (const name of available) sel.appendChild(buildOption(name, name));
        const cur = schedule.slots[id];
        sel.value = cur?.modo === "TODOS" ? "__TODOS__" : (cur?.asignadoA || "");

        sel.addEventListener("change", () => {
          if (sel.value === "__TODOS__") setSlot(schedule, id, { modo: "TODOS", asignadoA: "" });
          else setSlot(schedule, id, { modo: "NORMAL", asignadoA: sel.value });
          renderSummary(schedule);
          onSelectChange();
        });

        td.appendChild(sel);

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
  if (!document.getElementById("changeSlot")) return;
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
  if (!document.getElementById("changeSlot")) return;
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
  if (!document.getElementById("changesBody")) return;
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

function buildVacationsByISO(vacationRanges) {
  const byISO = {};
  for (const row of vacationRanges || []) {
    const person = normalizeName(row.person);
    const from = parseISODate(row.from);
    const to = parseISODate(row.to);
    if (!person || !from || !to) continue;
    let cur = startOfDay(from);
    const end = startOfDay(to);
    while (cur <= end) {
      const iso = toISODate(cur);
      if (!byISO[iso]) byISO[iso] = [];
      byISO[iso].push(person);
      cur = addDays(cur, 1);
    }
  }
  return byISO;
}

function buildVacationEventsForWeek(vacationsByISO, weekStartISO) {
  const week = buildWeek(parseISODate(weekStartISO));
  return week
    .map((d) => ({ iso: d.iso, names: vacationsByISO[d.iso] || [] }))
    .filter((x) => x.names.length);
}

function normalizeState(input) {
  const nowWeek = ensureWeekStartIsThursday(toISODate(computeWeekStartThursday(new Date())));
  const next = input || {};
  return {
    weekStart: ensureWeekStartIsThursday(next.weekStart || nowWeek),
    people: sortNames(unique((next.people || DEFAULT_PEOPLE).map(normalizeName).filter(Boolean))),
    excludedPeople: sortNames(unique((next.excludedPeople || []).map(normalizeName).filter(Boolean))),
    vacationRanges: Array.isArray(next.vacationRanges) ? next.vacationRanges : [],
    changesByWeek: next.changesByWeek || {},
    schedulesByWeek: next.schedulesByWeek || {},
    imagesMeta: Array.isArray(next.imagesMeta) ? next.imagesMeta : [],
    generatedAtByWeek: next.generatedAtByWeek || {},
  };
}

function buildPeopleTextarea(people) {
  return sortNames(people).join("\n");
}

function parsePeopleTextarea(raw) {
  return sortNames(unique(String(raw || "").split("\n").map(normalizeName).filter(Boolean)));
}

function renderVacationManagerList(state) {
  const el = qs("vacationsManagerList");
  const rows = [...state.vacationRanges].sort((a, b) => `${a.person}${a.from}`.localeCompare(`${b.person}${b.from}`));
  if (!rows.length) {
    el.textContent = "—";
    return;
  }
  el.innerHTML = rows.map((r) => `<div>${escapeHtml(r.person)}: ${escapeHtml(r.from)} → ${escapeHtml(r.to)}</div>`).join("");
}

function renderVacationPersonOptions(state) {
  const sel = qs("vacPerson");
  sel.innerHTML = "";
  for (const p of state.people) sel.appendChild(buildOption(p, p));
}

function wireTodosToggles(schedule) {
  void schedule;
}

function renderAll(schedule, peopleModel, state, onExcludeChange, onSelectChange) {
  window.__schedule = schedule;
  window.__peopleModel = peopleModel;
  window.__state = state;
  updateMeta(schedule);
  renderPeopleList(peopleModel, state);
  renderExcludeControl(peopleModel, state, onExcludeChange);
  renderMiniCalendar(schedule.weekStart);
  renderScheduleTable(schedule, peopleModel, onSelectChange);
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

async function saveImageLocal(schedule) {
  const capture = qs("scheduleCapture");
  const canvas = await html2canvas(capture, {
    backgroundColor: null,
    scale: Math.min(2, window.devicePixelRatio || 1),
    useCORS: true,
  });
  const dataUrl = canvas.toDataURL("image/png");
  const yyyy = new Date().getFullYear();
  const mm = String(new Date().getMonth() + 1).padStart(2, "0");
  const filename = `Turnos_${schedule.weekStart}_${Date.now()}.png`;
  const a = document.createElement("a");
  a.href = dataUrl;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  return {
    fileName: filename,
    previewUrl: dataUrl,
    githubPath: `imagenes/${yyyy}/${mm}/${filename}`,
    createdAt: new Date().toLocaleString("es-ES"),
  };
}

function init() {
  const state = normalizeState(loadState());
  const weekStart = ensureWeekStartIsThursday(state.weekStart || toISODate(computeWeekStartThursday(new Date())));
  qs("weekStart").value = weekStart;

  let schedule = state.schedulesByWeek[weekStart] || createEmptySchedule(weekStart);
  let changes = state.changesByWeek[weekStart] || [];
  let lastImage = state.imagesMeta[0] || null;
  let peopleModel = makePeopleModel(state.people, buildVacationsByISO(state.vacationRanges));

  const persist = (reason) => {
    state.weekStart = schedule.weekStart;
    state.people = peopleModel.all;
    state.schedulesByWeek[schedule.weekStart] = schedule;
    state.changesByWeek[schedule.weekStart] = changes;
    saveState(state);
    setStatus(`Guardado local (${reason})`, "ok");
  };

  const persistDebounced = (() => {
    let t = null;
    return (reason) => {
      if (t) clearTimeout(t);
      t = setTimeout(() => persist(reason), 800);
    };
  })();

  const rerender = () => {
    renderAll(schedule, peopleModel, state, () => {
      rerender();
      persistDebounced("exclusion");
    }, () => persistDebounced("edicion"));
    renderVacationsThisWeek(buildVacationEventsForWeek(peopleModel.vacationsByISO, schedule.weekStart));
    renderVacationPersonOptions(state);
    renderChangesHistory(changes);
  };
  window.__rerender = rerender;

  const hardReload = async () => {
    const iso = ensureWeekStartIsThursday(qs("weekStart").value);
    qs("weekStart").value = iso;
    state.weekStart = iso;
    peopleModel = makePeopleModel(state.people, buildVacationsByISO(state.vacationRanges));
    schedule = state.schedulesByWeek[iso] || createEmptySchedule(iso);
    changes = state.changesByWeek[iso] || [];
    rerender();
    wireTodosToggles(schedule);
    persistDebounced("cambio-semana");
  };

  qs("weekStart").addEventListener("change", () => hardReload());
  qs("btnMonthPrev").addEventListener("click", () => {
    state.calendarMonthOffset = Number(state.calendarMonthOffset || 0) - 1;
    rerender();
    persistDebounced("mes");
  });
  qs("btnMonthNext").addEventListener("click", () => {
    state.calendarMonthOffset = Number(state.calendarMonthOffset || 0) + 1;
    rerender();
    persistDebounced("mes");
  });

  qs("btnSaveImage").addEventListener("click", async () => {
    try {
      setStatus("Generando PNG…", "muted");
      const data = await saveImageLocal(schedule);
      lastImage = data;
      state.imagesMeta.unshift(data);
      state.imagesMeta = state.imagesMeta.slice(0, 50);
      persist("imagen");
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
      persistDebounced("generar");
    } catch (e) {
      setStatus(`No se pudo generar: ${e.message}`, "bad");
    }
  });

  qs("btnAddVacation").addEventListener("click", () => {
    const person = normalizeName(qs("vacPerson").value);
    const from = clampStr(qs("vacFrom").value);
    const to = clampStr(qs("vacTo").value) || from;
    if (!person || !from || !to) {
      setStatus("Selecciona comercial y rango de fechas.", "warn");
      return;
    }
    state.vacationRanges.push({ person, from, to });
    peopleModel = makePeopleModel(state.people, buildVacationsByISO(state.vacationRanges));
    rerender();
    persist("vacaciones");
  });

  qs("btnRemoveVacation").addEventListener("click", () => {
    const person = normalizeName(qs("vacPerson").value);
    const from = clampStr(qs("vacFrom").value);
    const to = clampStr(qs("vacTo").value) || from;
    state.vacationRanges = state.vacationRanges.filter((x) => !(x.person === person && x.from === from && x.to === to));
    peopleModel = makePeopleModel(state.people, buildVacationsByISO(state.vacationRanges));
    rerender();
    persist("vacaciones");
  });

  hardReload().catch((e) => setStatus(`Error: ${e.message}`, "bad"));
}

document.addEventListener("DOMContentLoaded", init);

