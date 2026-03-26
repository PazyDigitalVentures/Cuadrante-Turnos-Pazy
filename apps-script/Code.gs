const SHEETS_DEF = {
  Comerciales: ["id", "nombre", "activo"],
  Vacaciones: ["nombre", "desde", "hasta", "motivo"],
  Turnos: ["weekStart", "fecha", "franja", "tipo", "modo", "asignadoA", "nota"],
  Cambios: ["timestamp", "weekStart", "slotId", "antes", "despues", "motivo", "autor"],
  Config: ["clave", "valor"],
};

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "";
  const asset = (e && e.parameter && e.parameter.asset) || "";

  // Servir assets (CSS/JS) como ficheros
  if (asset) {
    if (asset === "js") {
      return ContentService.createTextOutput(include("AppJsRaw")).setMimeType(
        ContentService.MimeType.JAVASCRIPT,
      );
    }
    // CSS como asset es poco fiable por el mimetype; lo dejamos inline en Index.html
    return ContentService.createTextOutput("Not found").setMimeType(ContentService.MimeType.TEXT);
  }

  // Si no hay acción, servimos el frontend.
  if (!action) {
    const html = HtmlService.createTemplateFromFile("Index")
      .evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    html.setTitle("Turnos Pazy");
    html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    return html;
  }
  // Si hay action, respondemos como API (JSON)
  return doPost(e);
}

function doPost(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || "";
    const payload = parseJsonBody_(e);
    let out;

    if (action === "getBootstrap") out = getBootstrap_(payload);
    else if (action === "saveWeek") out = saveWeek_(payload);
    else if (action === "uploadPng") out = uploadPng_(payload);
    else throw new Error("Accion no soportada: " + action);

    return jsonOut_({ ok: true, ...out });
  } catch (err) {
    return jsonOut_({
      ok: false,
      error: err && err.message ? err.message : String(err),
    });
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function parseJsonBody_(e) {
  try {
    const txt = (e && e.postData && e.postData.contents) || "{}";
    return JSON.parse(txt);
  } catch (_err) {
    throw new Error("JSON invalido en body");
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON,
  );
}

function requireStr_(v, name) {
  const s = String(v || "").trim();
  if (!s) throw new Error("Falta campo: " + name);
  return s;
}

function openAndEnsure_(sheetId) {
  const ss = SpreadsheetApp.openById(sheetId);
  ensureSheets_(ss);
  return ss;
}

function ensureSheets_(ss) {
  Object.keys(SHEETS_DEF).forEach(function (sheetName) {
    const headers = SHEETS_DEF[sheetName];
    let sh = ss.getSheetByName(sheetName);
    if (!sh) sh = ss.insertSheet(sheetName);

    if (sh.getLastRow() < 1) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      styleHeader_(sh, headers.length);
      return;
    }

    const current = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    const needsWrite = headers.some(function (h, i) {
      return String(current[i] || "").trim() !== h;
    });
    if (needsWrite) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      styleHeader_(sh, headers.length);
    }
  });
}

function styleHeader_(sh, width) {
  const rng = sh.getRange(1, 1, 1, width);
  rng.setBackground("#7A0A3F");
  rng.setFontColor("#FFFFFF");
  rng.setFontWeight("bold");
  rng.setHorizontalAlignment("center");
  sh.setFrozenRows(1);
}

function getDataRows_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= 1) return [];
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return values.map(function (row) {
    const obj = {};
    headers.forEach(function (h, i) {
      obj[String(h)] = row[i];
    });
    return obj;
  });
}

function formatIsoDate_(v) {
  if (!v) return "";
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s);
  if (isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function getConfigMap_(ss) {
  const sh = ss.getSheetByName("Config");
  const rows = getDataRows_(sh);
  const map = {};
  rows.forEach(function (r) {
    const k = String(r.clave || "").trim();
    if (!k) return;
    map[k] = String(r.valor == null ? "" : r.valor).trim();
  });
  return map;
}

function getBootstrap_(payload) {
  const sheetId = requireStr_(payload.sheetId, "sheetId");
  const weekStart = requireStr_(payload.weekStart, "weekStart");
  const ss = openAndEnsure_(sheetId);

  const people = readPeople_(ss);
  const vacFromCal = readVacationsFromCalendarVacaciones_(ss, weekStart, people);
  const vacationsByISO = Object.keys(vacFromCal.byISO || {}).length
    ? vacFromCal.byISO
    : readVacationsMap_(ss, weekStart);
  const schedule = readWeekSchedule_(ss, weekStart);
  const changes = readWeekChanges_(ss, weekStart);
  const config = getConfigMap_(ss);

  return {
    people: people,
    vacationsByISO: vacationsByISO,
    vacationEvents: vacFromCal.events || [],
    schedule: schedule,
    changes: changes,
    config: config,
  };
}

function readPeople_(ss) {
  const rows = getDataRows_(ss.getSheetByName("Comerciales"));
  return rows
    .filter(function (r) {
      const name = String(r.nombre || "").trim();
      const activeRaw = String(r.activo == null ? "TRUE" : r.activo).toLowerCase();
      const active = !(activeRaw === "false" || activeRaw === "0" || activeRaw === "no");
      return name && active;
    })
    .map(function (r) {
      return String(r.nombre).trim();
    });
}

function readVacationsMap_(ss, weekStart) {
  const rows = getDataRows_(ss.getSheetByName("Vacaciones"));
  const map = {};
  const weekDates = weekDates_(weekStart); // 7 iso
  rows.forEach(function (r) {
    const name = String(r.nombre || "").trim();
    if (!name) return;
    const fromIso = formatIsoDate_(r.desde);
    const toIso = formatIsoDate_(r.hasta);
    if (!fromIso || !toIso) return;
    weekDates.forEach(function (iso) {
      if (iso >= fromIso && iso <= toIso) {
        if (!map[iso]) map[iso] = [];
        map[iso].push(name);
      }
    });
  });
  return map;
}

function normalizeName_(s) {
  return String(s || "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function monthIndexFromSpanish_(name) {
  const n = normalizeName_(name);
  const map = {
    enero: 0,
    febrero: 1,
    marzo: 2,
    abril: 3,
    mayo: 4,
    junio: 5,
    julio: 6,
    agosto: 7,
    septiembre: 8,
    setiembre: 8,
    octubre: 9,
    noviembre: 10,
    diciembre: 11,
  };
  return map.hasOwnProperty(n) ? map[n] : null;
}

function readVacationsFromCalendarVacaciones_(ss, weekStartIso, people) {
  const sh = ss.getSheetByName("Calendario vacaciones");
  if (!sh) return { byISO: {}, events: [] };

  const weekDatesIso = weekDates_(weekStartIso);
  const weekDates = weekDatesIso.map(function (iso) {
    return new Date(iso + "T00:00:00");
  });

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 15 || lastCol < 4) return { byISO: {}, events: [] };

  // Fila 12: meses (a veces solo al inicio de bloque); rellenamos hacia la derecha
  const monthsRow = sh.getRange(12, 1, 1, lastCol).getValues()[0];
  const daysRow = sh.getRange(14, 1, 1, lastCol).getValues()[0];

  const monthsFilled = [];
  let currentMonthName = "";
  for (let c = 0; c < lastCol; c++) {
    const v = String(monthsRow[c] || "").trim();
    if (v) currentMonthName = v;
    monthsFilled[c] = currentMonthName;
  }

  // Para cada fecha de la semana, buscar su columna por (mes, día)
  const dateToCol = {};
  weekDates.forEach(function (d) {
    const targetMonth = d.getMonth();
    const targetDay = d.getDate();
    let foundCol = null;

    for (let c = 0; c < lastCol; c++) {
      const mi = monthIndexFromSpanish_(monthsFilled[c]);
      if (mi == null || mi !== targetMonth) continue;
      const dayVal = Number(daysRow[c]);
      if (!dayVal || dayVal !== targetDay) continue;
      foundCol = c + 1; // 1-indexed
      break;
    }

    dateToCol[Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd")] = foundCol;
  });

  // Mapa nombre->fila buscando en columna C (solo comerciales)
  const peopleSet = {};
  (people || []).forEach(function (p) {
    peopleSet[normalizeName_(p)] = p;
  });

  const nameCol = sh.getRange(1, 3, lastRow, 1).getValues(); // columna C
  const rowByName = {};
  for (let r = 0; r < nameCol.length; r++) {
    const cell = nameCol[r][0];
    const key = normalizeName_(cell);
    if (peopleSet[key]) rowByName[peopleSet[key]] = r + 1;
  }

  const byISO = {};
  weekDatesIso.forEach(function (iso) {
    const col = dateToCol[iso];
    if (!col) return;
    const names = [];

    Object.keys(rowByName).forEach(function (name) {
      const row = rowByName[name];
      const v = sh.getRange(row, col).getValue();
      const sv = String(v || "").trim().toLowerCase();
      if (sv === "v") names.push(name);
    });

    if (names.length) byISO[iso] = names;
  });

  const events = Object.keys(byISO)
    .sort()
    .map(function (iso) {
      return { iso: iso, names: byISO[iso] };
    });

  return { byISO: byISO, events: events };
}

function weekDates_(weekStartIso) {
  const d0 = new Date(weekStartIso + "T00:00:00");
  const out = [];
  for (let i = 0; i < 7; i++) {
    const d = new Date(d0.getTime());
    d.setDate(d0.getDate() + i);
    out.push(Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd"));
  }
  return out;
}

function readWeekSchedule_(ss, weekStart) {
  const rows = getDataRows_(ss.getSheetByName("Turnos"));
  const out = rows
    .filter(function (r) {
      return String(r.weekStart || "").trim() === weekStart;
    })
    .map(function (r) {
      return {
        weekStart: String(r.weekStart || "").trim(),
        fecha: formatIsoDate_(r.fecha),
        franja: String(r.franja || "").trim(),
        tipo: String(r.tipo || "").trim(),
        modo: String(r.modo || "NORMAL").trim(),
        asignadoA: String(r.asignadoA || "").trim(),
        nota: String(r.nota || "").trim(),
      };
    });
  return { rows: out };
}

function readWeekChanges_(ss, weekStart) {
  const rows = getDataRows_(ss.getSheetByName("Cambios"));
  return rows
    .filter(function (r) {
      return String(r.weekStart || "").trim() === weekStart;
    })
    .slice(-200)
    .reverse()
    .map(function (r) {
      return {
        timestamp: String(r.timestamp || ""),
        weekStart: String(r.weekStart || ""),
        slotId: String(r.slotId || ""),
        antes: String(r.antes || ""),
        despues: String(r.despues || ""),
        motivo: String(r.motivo || ""),
        autor: String(r.autor || ""),
      };
    });
}

function saveWeek_(payload) {
  const sheetId = requireStr_(payload.sheetId, "sheetId");
  const weekStart = requireStr_(payload.weekStart, "weekStart");
  const rows = Array.isArray(payload.rows) ? payload.rows : [];
  const changes = Array.isArray(payload.changes) ? payload.changes : [];

  const ss = openAndEnsure_(sheetId);
  writeWeekSchedule_(ss, weekStart, rows);
  writeChanges_(ss, weekStart, changes);
  return { savedRows: rows.length, savedChanges: changes.length };
}

function writeWeekSchedule_(ss, weekStart, rows) {
  const sh = ss.getSheetByName("Turnos");
  const all = getDataRows_(sh);
  const keep = all.filter(function (r) {
    return String(r.weekStart || "").trim() !== weekStart;
  });
  const merged = keep.concat(
    rows.map(function (r) {
      return {
        weekStart: weekStart,
        fecha: formatIsoDate_(r.fecha),
        franja: String(r.franja || "").trim(),
        tipo: String(r.tipo || "").trim(),
        modo: String(r.modo || "NORMAL").trim(),
        asignadoA: String(r.asignadoA || "").trim(),
        nota: String(r.nota || "").trim(),
      };
    }),
  );

  sh.clearContents();
  sh.getRange(1, 1, 1, SHEETS_DEF.Turnos.length).setValues([SHEETS_DEF.Turnos]);
  styleHeader_(sh, SHEETS_DEF.Turnos.length);
  if (merged.length) {
    const values = merged.map(function (r) {
      return SHEETS_DEF.Turnos.map(function (h) {
        return r[h] == null ? "" : r[h];
      });
    });
    sh.getRange(2, 1, values.length, SHEETS_DEF.Turnos.length).setValues(values);
  }
}

function writeChanges_(ss, weekStart, changes) {
  const sh = ss.getSheetByName("Cambios");
  const all = getDataRows_(sh);
  const keep = all.filter(function (r) {
    return String(r.weekStart || "").trim() !== weekStart;
  });
  const merged = keep.concat(
    changes.map(function (c) {
      return {
        timestamp: String(c.timestamp || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")),
        weekStart: weekStart,
        slotId: String(c.slotId || ""),
        antes: String(c.antes || ""),
        despues: String(c.despues || ""),
        motivo: String(c.motivo || ""),
        autor: String(c.autor || ""),
      };
    }),
  );

  sh.clearContents();
  sh.getRange(1, 1, 1, SHEETS_DEF.Cambios.length).setValues([SHEETS_DEF.Cambios]);
  styleHeader_(sh, SHEETS_DEF.Cambios.length);
  if (merged.length) {
    const values = merged.map(function (r) {
      return SHEETS_DEF.Cambios.map(function (h) {
        return r[h] == null ? "" : r[h];
      });
    });
    sh.getRange(2, 1, values.length, SHEETS_DEF.Cambios.length).setValues(values);
  }
}

function uploadPng_(payload) {
  const sheetId = requireStr_(payload.sheetId, "sheetId");
  const weekStart = requireStr_(payload.weekStart, "weekStart");
  const filename = requireStr_(payload.filename, "filename");
  const dataUrl = requireStr_(payload.dataUrl, "dataUrl");
  const ss = openAndEnsure_(sheetId);
  const cfg = getConfigMap_(ss);

  const base64 = dataUrl.replace(/^data:image\/png;base64,/, "");
  if (!base64 || base64 === dataUrl) throw new Error("PNG invalido en dataUrl");
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, "image/png", filename);

  const folderId = String(cfg.carpetaDriveId || "").trim();
  const folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  const file = folder.createFile(blob);

  appendImageLog_(ss, weekStart, filename, file.getId(), file.getUrl());
  return {
    fileId: file.getId(),
    fileName: file.getName(),
    fileUrl: file.getUrl(),
  };
}

function appendImageLog_(ss, weekStart, filename, fileId, fileUrl) {
  const shName = "Imagenes";
  let sh = ss.getSheetByName(shName);
  const headers = ["timestamp", "weekStart", "filename", "fileId", "fileUrl"];
  if (!sh) {
    sh = ss.insertSheet(shName);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader_(sh, headers.length);
  }
  sh.appendRow([
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"),
    weekStart,
    filename,
    fileId,
    fileUrl,
  ]);
}

