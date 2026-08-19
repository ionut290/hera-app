"use strict";

const DEFAULT_REGION = "www.appsheet.com";
const ROME_TIMEZONE = "Europe/Rome";

function normalizeText(value) {
  return String(value ?? "").trim();
}

function getPath(object, path) {
  if (!path) return undefined;
  return String(path)
    .split(".")
    .filter(Boolean)
    .reduce((current, key) => (current == null ? undefined : current[key]), object);
}

function toDate(value) {
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value;
  if (value && typeof value.toDate === "function") {
    try {
      return toDate(value.toDate());
    } catch (_) {
      return null;
    }
  }
  if (value && Number.isFinite(Number(value.seconds))) {
    return new Date(Number(value.seconds) * 1000);
  }
  if (typeof value === "string" || typeof value === "number") {
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }
  return null;
}

function formatRomeDate(value) {
  const date = toDate(value);
  if (!date) return "";
  const parts = new Intl.DateTimeFormat("it-IT", {
    timeZone: ROME_TIMEZONE,
    day: "2-digit",
    month: "2-digit",
    year: "numeric"
  }).formatToParts(date);
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${values.day}/${values.month}/${values.year}`;
}

function formatRomeTime(value) {
  const date = toDate(value);
  if (!date) return "";
  return new Intl.DateTimeFormat("it-IT", {
    timeZone: ROME_TIMEZONE,
    hour: "2-digit",
    minute: "2-digit",
    hourCycle: "h23"
  }).format(date);
}

function normalizeStatus(value) {
  const status = normalizeText(value).toLocaleUpperCase("it-IT").replace(/_/g, " ");
  if (["DONE", "COMPLETATO", "COMPLETATA"].includes(status)) return "FATTO";
  return status;
}

function firstPresent(values) {
  for (const value of values) {
    if (value === false || value === 0) return value;
    if (value !== undefined && value !== null && String(value).trim() !== "") return value;
  }
  return undefined;
}

function getLat(source) {
  return firstPresent([
    getPath(source, "latitudine"),
    getPath(source, "gpsY"),
    getPath(source, "lat"),
    getPath(source, "latitude")
  ]);
}

function getLon(source) {
  return firstPresent([
    getPath(source, "longitudine"),
    getPath(source, "gpsX"),
    getPath(source, "lon"),
    getPath(source, "lng"),
    getPath(source, "longitude")
  ]);
}

function resolveSpecialValue(source, token, context) {
  switch (token) {
    case "$id":
      return context.id;
    case "$parentId":
      return context.parentId;
    case "$path":
      return context.path;
    case "$deleted":
      return Boolean(context.deleted);
    case "$doneStatus":
      return source?.done === true ? "FATTO" : "DA FARE";
    case "$doneDate":
      return formatRomeDate(source?.doneAt);
    case "$doneTime":
      return formatRomeTime(source?.doneAt);
    case "$operator":
      return firstPresent([source?.operatoreNome, source?.operatore, source?.doneBy]);
    case "$lat":
      return getLat(source);
    case "$lon":
      return getLon(source);
    case "$coordinates": {
      const lat = getLat(source);
      const lon = getLon(source);
      return lat !== undefined && lon !== undefined ? `${lat},${lon}` : undefined;
    }
    case "$commessaName":
      return firstPresent([
        context?.commessa?.nome,
        context?.commessa?.nomeCommessa,
        context?.commessa?.commessa,
        context?.commessa?.denominazione
      ]);
    default:
      return undefined;
  }
}

function resolveSourceValue(source, sourcePath, context) {
  if (Array.isArray(sourcePath)) {
    return firstPresent(sourcePath.map((entry) => resolveSourceValue(source, entry, context)));
  }
  if (typeof sourcePath !== "string") return sourcePath;
  if (sourcePath.startsWith("$")) return resolveSpecialValue(source, sourcePath, context);
  return getPath(source, sourcePath);
}

function normalizeValue(value) {
  if (value == null) return value;
  if (value instanceof Date) return value.toISOString();
  if (typeof value?.toDate === "function") {
    try {
      return value.toDate().toISOString();
    } catch (_) {
      return null;
    }
  }
  if (Array.isArray(value)) return value.map(normalizeValue);
  if (typeof value === "object") {
    if (Number.isFinite(value.latitude) && Number.isFinite(value.longitude)) {
      return `${value.latitude},${value.longitude}`;
    }
    const normalized = {};
    for (const [key, nestedValue] of Object.entries(value)) normalized[key] = normalizeValue(nestedValue);
    return normalized;
  }
  return value;
}

function buildMappedRow(source, tableConfig, context = {}) {
  const fieldMap = tableConfig?.fields && typeof tableConfig.fields === "object" ? tableConfig.fields : {};
  const row = {};
  for (const [appSheetColumn, sourcePath] of Object.entries(fieldMap)) {
    let value = resolveSourceValue(source, sourcePath, context);
    if (appSheetColumn === "STATO" && value !== undefined) value = normalizeStatus(value);
    if (value !== undefined && value !== "") row[appSheetColumn] = normalizeValue(value);
  }
  return row;
}

function normalizeMatch(value) {
  return normalizeText(value)
    .toLocaleUpperCase("it-IT")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

function getCommessaName(commessa) {
  return firstPresent([
    commessa?.nome,
    commessa?.nomeCommessa,
    commessa?.commessa,
    commessa?.denominazione,
    commessa?.titolo
  ]);
}

function selectImpiantiTable(config, commessa) {
  const name = normalizeMatch(getCommessaName(commessa));
  const routes = Array.isArray(config?.tables?.impiantiRoutes) ? config.tables.impiantiRoutes : [];
  for (const route of routes) {
    const matches = Array.isArray(route?.matchNames) ? route.matchNames : [];
    if (matches.some((candidate) => normalizeMatch(candidate) === name)) return route;
  }
  return config?.tables?.impianti || null;
}

function validateTableConfig(tableConfig) {
  if (!tableConfig || typeof tableConfig !== "object") return { valid: false, reason: "table-config-missing" };
  if (tableConfig.enabled === false) return { valid: false, reason: "table-disabled" };
  if (!normalizeText(tableConfig.tableName)) return { valid: false, reason: "table-name-missing" };
  if (!normalizeText(tableConfig.keyColumn)) return { valid: false, reason: "key-column-missing" };
  if (!tableConfig.fields || typeof tableConfig.fields !== "object") return { valid: false, reason: "field-map-missing" };
  if (!Object.prototype.hasOwnProperty.call(tableConfig.fields, tableConfig.keyColumn)) {
    return { valid: false, reason: "key-column-not-mapped" };
  }
  return { valid: true, reason: "" };
}

function validateConfig(config) {
  if (!config || config.enabled !== true) return { valid: false, reason: "disabled" };
  if (!normalizeText(config.appId)) return { valid: false, reason: "app-id-missing" };
  const commesse = validateTableConfig(config.tables?.commesse);
  if (!commesse.valid) return { valid: false, reason: `commesse:${commesse.reason}` };
  const impianti = validateTableConfig(config.tables?.impianti);
  if (!impianti.valid) return { valid: false, reason: `impianti:${impianti.reason}` };
  return { valid: true, reason: "" };
}

function shouldSyncWrite(beforeExists, afterExists, tableConfig) {
  if (!afterExists) return tableConfig?.syncDeletes === true ? "Delete" : "Ignore";
  return beforeExists ? "Edit" : "Add";
}

function buildActionBody(action, rows, options = {}) {
  return {
    Action: action,
    Properties: {
      Locale: options.locale || "it-IT",
      Timezone: options.timezone || ROME_TIMEZONE,
      ...(options.userId ? { UserId: options.userId } : {})
    },
    Rows: rows
  };
}

function buildActionUrl(config, tableName) {
  const region = normalizeText(config?.region) || DEFAULT_REGION;
  const appId = normalizeText(config?.appId);
  return `https://${region}/api/v2/apps/${encodeURIComponent(appId)}/tables/${encodeURIComponent(tableName)}/Action`;
}

function mapInboundRow(row, inboundFields) {
  if (!row || typeof row !== "object") return {};
  if (!inboundFields || typeof inboundFields !== "object") return {};
  const patch = {};
  for (const [firestoreField, appSheetColumn] of Object.entries(inboundFields)) {
    if (!appSheetColumn) continue;
    if (Object.prototype.hasOwnProperty.call(row, appSheetColumn)) {
      patch[firestoreField] = row[appSheetColumn];
    }
  }
  return patch;
}

function hasOwnEntries(object) {
  return Boolean(object && typeof object === "object" && Object.keys(object).length);
}

module.exports = {
  DEFAULT_REGION,
  ROME_TIMEZONE,
  normalizeText,
  getPath,
  toDate,
  formatRomeDate,
  formatRomeTime,
  normalizeStatus,
  firstPresent,
  resolveSourceValue,
  normalizeValue,
  buildMappedRow,
  normalizeMatch,
  getCommessaName,
  selectImpiantiTable,
  validateTableConfig,
  validateConfig,
  shouldSyncWrite,
  buildActionBody,
  buildActionUrl,
  mapInboundRow,
  hasOwnEntries
};
