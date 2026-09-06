"use strict";

const crypto = require("crypto");

const CODE_PREFIX = "REC-";
const MIN_CODE_LENGTH = 16;
const MAX_SECRET_LENGTH = 128;
const MIN_PASSWORD_LENGTH = 10;

function clean(value) {
  return String(value || "").trim();
}

function isRecoveryCodeCandidate(value) {
  const code = clean(value);
  return code.startsWith(CODE_PREFIX)
    && code.length >= MIN_CODE_LENGTH
    && code.length <= MAX_SECRET_LENGTH
    && !/\s/.test(code);
}

function isValidNewPassword(value) {
  const password = String(value || "");
  return password.length >= MIN_PASSWORD_LENGTH && password.length <= MAX_SECRET_LENGTH;
}

function createSalt() {
  return crypto.randomBytes(24).toString("base64");
}

function hashRecoveryCode(value, salt) {
  return crypto.scryptSync(clean(value), String(salt || ""), 64).toString("base64");
}

function secureEqual(left, right) {
  try {
    const a = Buffer.from(String(left || ""), "base64");
    const b = Buffer.from(String(right || ""), "base64");
    return a.length === b.length && crypto.timingSafeEqual(a, b);
  } catch (_) {
    return false;
  }
}

function opaqueId(value) {
  return crypto.createHash("sha256").update(String(value || "")).digest("hex");
}

module.exports = {
  CODE_PREFIX,
  MIN_CODE_LENGTH,
  MIN_PASSWORD_LENGTH,
  MAX_SECRET_LENGTH,
  isRecoveryCodeCandidate,
  isValidNewPassword,
  createSalt,
  hashRecoveryCode,
  secureEqual,
  opaqueId
};
