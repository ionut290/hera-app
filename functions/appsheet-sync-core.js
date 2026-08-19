"use strict";

const DEFAULT_REGION = "www.appsheet.com";

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

function resolveSourceValue(source, sourcePath, context) {
  switch (sourcePath) {
    case "$id":
      return context.id;
    case "$parentId":
      return context.parentId;
    case "$path":
      return context.path;
    case "$deleted":
      return Boolean(context.deleted);
    default:
      return getPath(source, sourcePath);
  }
}

function buildMappedRow(source, tableConfig, context = {}) {
  const fieldMap = tableConfig?.fields && typeof tableConfig.fields === "object" ? tableConfig.fields : {};
  const row = {};
  for (const [appSheetColumn, sourcePath] of Object.entries(fieldMap)) {
    const value = resolveSourceValue(source, sourcePath, context);
    if (value !== undefined) row[appSheetColumn] = normalizeValue(value);
  }
  return row;
}

function validateTableConfig(tableConfig) {
  if (!tableConfig || typeof tableConfig !== "object") return { valid: false, reason: "table-config-missing" };
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
      Timezone: options.timezone || "Europe/Rome",
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
  normalizeText,
  getPath,
  normalizeValue,
  buildMappedRow,
  validateTableConfig,
  validateConfig,
  shouldSyncWrite,
  buildActionBody,
  buildActionUrl,
  mapInboundRow,
  hasOwnEntries
};
