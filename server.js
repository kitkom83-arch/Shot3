const dotenv = require("dotenv");
dotenv.config({ override: true });

const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");
const crypto = require("crypto");
const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const { parse: parseCsvSync } = require("csv-parse/sync");
const { parse: parseCsvStream } = require("csv-parse");
const { stringify } = require("csv-stringify/sync");
const { TelegramClient, Api } = require("telegram");
const { StringSession } = require("telegram/sessions");
const {
  APP_USER,
  APP_PASSWORD,
  APP_SECRET,
  AUTH_COOKIE,
  PUBLIC_DIR,
  parseCookies,
  signToken,
  verifyToken,
  setAuthCookie,
  clearAuthCookie,
  requireAuth,
} = require("./auth-addon");


const app = express();

const HOST = process.env.HOST || "0.0.0.0";
const PORT = toPositiveNumber(process.env.PORT, 3000);
const MAX_UPLOAD_MB = Math.max(200, toPositiveNumber(process.env.MAX_UPLOAD_MB, 200));
const TRUST_PROXY = String(process.env.TRUST_PROXY || "").trim() === "1";
const FALLBACK_BATCH_SIZE = 1;
const FALLBACK_MAX_PER_RUN = 20;
const FLOOD_STALE_THRESHOLD_SEC = toPositiveNumber(process.env.FLOOD_STALE_THRESHOLD_SEC, 3600);
const FLOOD_MAX_BACKOFF_SEC = toPositiveNumber(process.env.FLOOD_MAX_BACKOFF_SEC, 7200);

const APP_ROOT = __dirname;
const DATA_DIR = path.resolve(process.env.DATA_DIR || path.join(APP_ROOT, "data"));
const RAW_DIR = path.join(DATA_DIR, "input_raw");
const CLEAN_DIR = path.join(DATA_DIR, "input_clean");
const OUTPUT_DIR = path.join(DATA_DIR, "output");
const SESSION_DIR = path.join(DATA_DIR, "session");
const LOG_DIR = path.join(DATA_DIR, "logs");
const RECOVERY_DIR = path.join(DATA_DIR, "recovery");

const FILES = {
  accounts: path.join(DATA_DIR, "accounts.json"),
  appState: path.join(DATA_DIR, "app_state.json"),
  settings: path.join(DATA_DIR, "settings.json"),
  intakeLatest: path.join(CLEAN_DIR, "intake_latest.json"),
  cleanReadyCsv: path.join(CLEAN_DIR, "clean_ready.csv"),
  invalidRowsCsv: path.join(CLEAN_DIR, "invalid_rows.csv"),
  duplicatePhonesCsv: path.join(CLEAN_DIR, "duplicate_phones.csv"),
  cleanDebugCsv: path.join(CLEAN_DIR, "clean_debug.csv"),
  cleanRejectsCsv: path.join(CLEAN_DIR, "clean_rejects.csv"),
  intakeSummaryJson: path.join(CLEAN_DIR, "summary.json"),
  currentJob: path.join(OUTPUT_DIR, "current_job.json"),
  jobState: path.join(OUTPUT_DIR, "job_state.json"),
  latestJson: path.join(OUTPUT_DIR, "telegram_matches.json"),
  latestCsv: path.join(OUTPUT_DIR, "telegram_matches.csv"),
  allJson: path.join(OUTPUT_DIR, "telegram_matches_all.json"),
  allCsv: path.join(OUTPUT_DIR, "telegram_matches_all.csv"),
  retryCsv: path.join(OUTPUT_DIR, "retry_rows.csv"),
  autoState: path.join(OUTPUT_DIR, "auto_state.json"),
  runLog: path.join(LOG_DIR, "run_log.json"),
  remainingCsv: path.join(OUTPUT_DIR, "remaining_only.csv"),
};


const CRITICAL_BACKUP_FILES = new Set([
  path.resolve(FILES.currentJob),
  path.resolve(FILES.jobState),
  path.resolve(FILES.autoState),
  path.resolve(FILES.settings),
  path.resolve(FILES.appState),
  path.resolve(FILES.accounts),
  path.resolve(FILES.runLog),
]);

const recoveryNotices = [];

function lastGoodPath(filePath) {
  return `${filePath}.lastgood`;
}

function shouldKeepLastGood(filePath) {
  return CRITICAL_BACKUP_FILES.has(path.resolve(filePath));
}

function pushRecoveryNotice(level, code, title, message, extra = {}) {
  recoveryNotices.push({ at: nowIso(), level, code, title, message, ...extra });
  if (recoveryNotices.length > 80) recoveryNotices.splice(0, recoveryNotices.length - 80);
}

const DEFAULT_SETTINGS = {
  autoRun: false,
  batchSize: FALLBACK_BATCH_SIZE,
  maxContactsPerRun: FALLBACK_MAX_PER_RUN,
  delayBetweenRunsSec: 120,
  retryPauseSec: 600,
  retryRatioThreshold: 0.2,
  waitFloodAutomatically: true,
};

if (TRUST_PROXY) app.set("trust proxy", true);

let isProcessing = false;
const activeClients = new Map();
const pendingAuthClients = new Map();
let autoTimer = null;

function toPositiveNumber(value, fallback) {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

function nowIso() {
  return new Date().toISOString();
}

function computeRemainingSec(nextRunAt) {
  if (!nextRunAt) return 0;
  const ts = new Date(nextRunAt).getTime();
  if (!Number.isFinite(ts)) return 0;
  return Math.max(0, Math.ceil((ts - Date.now()) / 1000));
}

function displayQueueStatus(status) {
  const value = String(status || '').toLowerCase();
  if (value === 'waiting_retry_cooldown') return 'WAITING_RETRY';
  if (value === 'waiting_flood') return 'WAITING_FLOOD';
  if (value === 'paused_flood_stale') return 'PAUSED_FLOOD_STALE';
  if (value === 'completed') return 'COMPLETED';
  if (value === 'paused') return 'PAUSED';
  if (value === 'ready') return 'READY';
  if (value === 'running') return 'RUNNING';
  return value ? value.toUpperCase() : 'IDLE';
}

function displayAutoStatus(status) {
  const value = String(status || '').toLowerCase();
  if (value === 'off') return 'OFF';
  if (value === 'running') return 'RUNNING';
  if (value === 'waiting_flood') return 'WAITING_FLOOD';
  if (value === 'waiting_retry_cooldown' || value === 'waiting_retry') return 'WAITING_RETRY';
  if (value === 'paused_flood_stale') return 'PAUSED_FLOOD_STALE';
  if (value === 'paused') return 'PAUSED';
  return value ? value.toUpperCase() : 'OFF';
}

function shouldRequireFloodReset(job) {
  return String(job?.status || "").toLowerCase() === "paused_flood_stale";
}

function logInfo(message) {
  console.log(`[INFO] ${message}`);
}

function onlyDigits(value) {
  return String(value || "").replace(/\D+/g, "");
}

function sanitizeFileName(name) {
  return String(name || "upload").replace(/[^a-zA-Z0-9._-]+/g, "_");
}

function makeId(prefix = "id") {
  return `${prefix}_${Date.now().toString(36)}_${crypto.randomBytes(4).toString("hex")}`;
}

function chunkArray(items, size) {
  const chunks = [];
  for (let i = 0; i < items.length; i += size) chunks.push(items.slice(i, i + size));
  return chunks;
}

function encryptionKey() {
  return crypto.createHash("sha256").update(APP_SECRET).digest();
}

function encryptText(value) {
  const plain = String(value || "");
  if (!plain) return "";
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", encryptionKey(), iv);
  const encrypted = Buffer.concat([cipher.update(plain, "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();
  return `${iv.toString("base64")}:${tag.toString("base64")}:${encrypted.toString("base64")}`;
}

function decryptText(value) {
  const encoded = String(value || "");
  if (!encoded) return "";
  const parts = encoded.split(":");
  if (parts.length !== 3) return "";
  const [ivB64, tagB64, dataB64] = parts;
  const decipher = crypto.createDecipheriv("aes-256-gcm", encryptionKey(), Buffer.from(ivB64, "base64"));
  decipher.setAuthTag(Buffer.from(tagB64, "base64"));
  const plain = Buffer.concat([
    decipher.update(Buffer.from(dataB64, "base64")),
    decipher.final(),
  ]);
  return plain.toString("utf8");
}

function maskSecret(value, start = 4, end = 2) {
  const text = String(value || "");
  if (!text) return "";
  if (text.length <= start + end) return "*".repeat(text.length);
  return `${text.slice(0, start)}${"*".repeat(text.length - start - end)}${text.slice(-end)}`;
}

function stripPhoneFormatting(value) {
  return String(value || "").replace(/[\s.\-()]/g, "").trim();
}

function normalizeThaiPhoneDetailed(value) {
  const raw = String(value || "").trim();
  if (!raw) return { ok: false, normalized: "", reason: "empty_phone", sanitized: "", digits: "" };

  const sanitized = stripPhoneFormatting(raw);
  if (!sanitized) return { ok: false, normalized: "", reason: "empty_phone", sanitized: "", digits: "" };

  let digits = onlyDigits(sanitized);
  if (!digits) return { ok: false, normalized: "", reason: "no_digits", sanitized, digits: "" };

  if (digits.startsWith("0066")) digits = digits.slice(2);

  let normalized = "";
  if (digits.length === 10 && digits.startsWith("0")) {
    normalized = `+66${digits.slice(1)}`;
  } else if (digits.length === 11 && digits.startsWith("66")) {
    normalized = `+${digits}`;
  } else if (digits.length === 12 && digits.startsWith("660")) {
    normalized = `+66${digits.slice(3)}`;
  } else {
    return { ok: false, normalized: "", reason: "invalid_length_or_prefix", sanitized, digits };
  }

  if (!/^\+66[689]\d{8}$/.test(normalized)) {
    return { ok: false, normalized: "", reason: "unsupported_thai_mobile_pattern", sanitized, digits };
  }

  return { ok: true, normalized, reason: "", sanitized, digits };
}

function normalizeThaiPhoneStrict(value) {
  return normalizeThaiPhoneDetailed(value).normalized;
}

function scorePhoneCandidate(value) {
  const raw = String(value || "").trim();
  if (!raw) return 0;
  const cleaned = stripPhoneFormatting(raw);
  if (!cleaned) return 0;
  const digits = onlyDigits(cleaned);
  if (!digits) return 0;

  if (/^\+?66\d{9}$/.test(cleaned) || /^0\d{9}$/.test(cleaned)) return 6;
  if ((digits.length === 10 && digits.startsWith("0")) || (digits.length === 11 && digits.startsWith("66"))) return 5;
  if (digits.length >= 9 && digits.length <= 12) return 2;
  return 0;
}

function inferPhoneColumnBySamples(headers, rows, sampleSize = 20) {
  if (!Array.isArray(headers) || !headers.length || !Array.isArray(rows) || !rows.length) {
    return { column: "", source: "sample_scan", confidence: 0, score: 0 };
  }

  let bestColumn = "";
  let bestScore = -1;
  const sampleRows = rows.slice(0, sampleSize);
  for (const header of headers) {
    let score = 0;
    for (const row of sampleRows) {
      score += scorePhoneCandidate(row?.[header]);
    }
    if (score > bestScore) {
      bestScore = score;
      bestColumn = header;
    }
  }

  const maxScore = sampleRows.length * 6;
  const confidence = maxScore > 0 ? Number((bestScore / maxScore).toFixed(3)) : 0;
  if (!bestColumn || bestScore <= 0) return { column: "", source: "sample_scan", confidence: 0, score: 0 };
  return { column: bestColumn, source: "sample_scan", confidence, score: bestScore };
}

function validateBootConfig() {
  const issues = [];
  const password = String(APP_PASSWORD || "").trim();
  const secret = String(APP_SECRET || "").trim();

  if (!password) {
    issues.push("APP_PASSWORD ยังไม่ได้ตั้งค่า");
  } else if (password.length < 8 || /changeme|default|password/i.test(password)) {
    issues.push("APP_PASSWORD ไม่ปลอดภัยพอ (ห้ามใช้ค่า default)");
  }

  if (!secret) {
    issues.push("APP_SECRET ยังไม่ได้ตั้งค่า");
  } else if (secret.length < 32 || /changeme|default|secret/i.test(secret)) {
    issues.push("APP_SECRET ไม่ปลอดภัยพอ (ต้องยาว >= 32 และไม่ใช่ค่า default)");
  }

  if (issues.length) {
    throw new Error(`CONFIG_ERROR: ${issues.join(" | ")}`);
  }
}

function computeFloodBackoffSec(baseSeconds, consecutiveCount) {
  const safeBase = Math.max(1, toPositiveNumber(baseSeconds, 60));
  const safeCount = Math.max(1, Math.floor(Number(consecutiveCount || 1)));
  const computed = safeBase * Math.pow(2, safeCount - 1);
  return Math.min(FLOOD_MAX_BACKOFF_SEC, Math.max(1, Math.round(computed)));
}

function floodStaleAgeSec(job) {
  if (!job) return 0;
  const marker = job.lastFloodAt || job.updatedAt || job.nextRunAt || "";
  const ts = marker ? new Date(marker).getTime() : NaN;
  if (!Number.isFinite(ts)) return 0;
  return Math.max(0, Math.floor((Date.now() - ts) / 1000));
}

function toSafeStringId(value) {
  if (value === null || value === undefined) return "";
  try {
    return String(value);
  } catch {
    return "";
  }
}

function parseFloodSeconds(message) {
  const text = String(message || "");
  const match = text.match(/wait of (\d+) seconds/i) || text.match(/FLOOD_WAIT[_ ]?(\d+)/i);
  return match ? Number(match[1]) : 0;
}

function normalizeErrorMessage(error) {
  return String(error?.message || error?.errorMessage || error || "");
}

function isAuthSessionError(error) {
  const upper = normalizeErrorMessage(error).toUpperCase();
  return (
    upper.includes("UNSUPPORTED STATE OR UNABLE TO AUTHENTICATE DATA")
    || upper.includes("AUTH_KEY_UNREGISTERED")
    || upper.includes("SESSION_REVOKED")
    || upper.includes("SESSION_EXPIRED")
    || upper.includes("AUTH_KEY")
    || upper.includes("PASSWORD_HASH_INVALID")
    || upper.includes("API_ID_INVALID")
    || upper.includes("API_HASH_INVALID")
    || upper.includes("NEEDS_RELOGIN")
    || upper.includes("AUTH_INVALID")
    || upper.includes("SESSION_INVALID")
    || upper.includes("DECRYPT")
  );
}

function makeAccountAuthError(message, accountStatus = "needs_relogin") {
  const err = new Error(message || "session ใช้งานไม่ได้ กรุณาล็อกอินใหม่");
  err.isAccountAuthError = true;
  err.accountStatus = accountStatus;
  return err;
}

function extractTelegramError(error) {
  const raw = normalizeErrorMessage(error) || "ไม่ทราบสาเหตุ";
  const upper = String(raw).toUpperCase();

  if (upper.includes("SESSION_PASSWORD_NEEDED")) return "บัญชีนี้เปิด 2FA อยู่ กรุณาใส่รหัสผ่าน 2FA";
  if (upper.includes("PHONE_CODE_INVALID")) return "รหัส OTP ไม่ถูกต้อง";
  if (upper.includes("PHONE_CODE_EXPIRED")) return "รหัส OTP หมดอายุแล้ว กรุณาส่ง OTP ใหม่";
  if (upper.includes("PHONE_NUMBER_INVALID")) return "เบอร์โทรไม่ถูกต้อง";
  if (upper.includes("PHONE_NUMBER_UNOCCUPIED")) return "เบอร์นี้ยังไม่มีบัญชี Telegram";
  if (upper.includes("PHONE_NUMBER_BANNED")) return "เบอร์นี้ถูกแบนจาก Telegram";
  if (upper.includes("PHONE_NUMBER_FLOOD")) return "เบอร์นี้ขอรหัสบ่อยเกินไป กรุณารอสักพัก";
  if (upper.includes("API_ID_INVALID")) return "API_ID หรือ API_HASH ไม่ถูกต้อง";
  if (upper.includes("AUTH_KEY_UNREGISTERED")) return "session ใช้งานไม่ได้แล้ว กรุณาล็อกอินใหม่";
  if (upper.includes("UNSUPPORTED STATE OR UNABLE TO AUTHENTICATE DATA")) return "session ถอดรหัสไม่ได้หรือไม่ถูกต้อง กรุณารีล็อกอินบัญชีนี้ใหม่";
  if (upper.includes("PASSWORD_HASH_INVALID")) return "รหัส 2FA ไม่ถูกต้อง";
  if (upper.includes("FLOOD_WAIT") || upper.includes("A WAIT OF")) return raw;

  return raw;
}

function sessionFilePath(accountId) {
  const safeId = String(accountId || "").replace(/[^a-zA-Z0-9_-]/g, "_");
  return path.join(SESSION_DIR, `${safeId}.session`);
}

async function readSessionFromFile(accountId) {
  const filePath = sessionFilePath(accountId);
  if (!fs.existsSync(filePath)) return "";
  try {
    return String(await fsp.readFile(filePath, "utf8")).trim();
  } catch {
    return "";
  }
}

async function writeSessionToFile(accountId, sessionString) {
  const filePath = sessionFilePath(accountId);
  const text = String(sessionString || "").trim();
  if (!text) {
    await unlinkIfExists(filePath);
    return;
  }
  await writeTextAtomic(filePath, text);
}

async function loadAccountSessionString(account) {
  const encoded = String(account?.sessionEnc || "").trim();
  if (encoded) {
    try {
      const decrypted = decryptText(encoded).trim();
      if (decrypted) return { sessionString: decrypted, source: "accounts_json" };
    } catch (error) {
      const fromFile = await readSessionFromFile(account?.id);
      if (fromFile) return { sessionString: fromFile, source: "session_file" };
      throw makeAccountAuthError("session เดิมของบัญชีนี้ใช้งานไม่ได้ กรุณาล็อกอินใหม่", "auth_invalid");
    }
  }
  const fromFile = await readSessionFromFile(account?.id);
  if (fromFile) return { sessionString: fromFile, source: "session_file" };
  return { sessionString: "", source: "" };
}

async function ensureDirectories() {
  await Promise.all([
    fsp.mkdir(DATA_DIR, { recursive: true }),
    fsp.mkdir(RAW_DIR, { recursive: true }),
    fsp.mkdir(CLEAN_DIR, { recursive: true }),
    fsp.mkdir(OUTPUT_DIR, { recursive: true }),
    fsp.mkdir(SESSION_DIR, { recursive: true }),
    fsp.mkdir(LOG_DIR, { recursive: true }),
    fsp.mkdir(RECOVERY_DIR, { recursive: true }),
  ]);
}


async function writeTextAtomic(filePath, text) {
  await fsp.mkdir(path.dirname(filePath), { recursive: true });
  const tmpFile = path.join(
    path.dirname(filePath),
    `.${path.basename(filePath)}.${process.pid}.${Date.now()}.${crypto.randomBytes(4).toString("hex")}.tmp`
  );
  await fsp.writeFile(tmpFile, text, "utf8");
  await fsp.rename(tmpFile, filePath);
}

async function writeJson(filePath, data) {
  const text = JSON.stringify(data, null, 2);
  await writeTextAtomic(filePath, text);
  if (shouldKeepLastGood(filePath)) {
    await writeTextAtomic(lastGoodPath(filePath), text).catch(() => {});
  }
}

async function stashBrokenJson(filePath, raw = "") {
  try {
    await fsp.mkdir(RECOVERY_DIR, { recursive: true });
    const badPath = path.join(
      RECOVERY_DIR,
      `${path.basename(filePath, path.extname(filePath))}.bad.${Date.now()}.json`
    );
    if (typeof raw === "string" && raw.length) {
      await fsp.writeFile(badPath, raw, "utf8");
    } else if (fs.existsSync(filePath)) {
      await fsp.copyFile(filePath, badPath);
    }
    return badPath;
  } catch {
    return "";
  }
}

async function readJson(filePath, fallback = null) {
  if (!fs.existsSync(filePath)) return fallback;

  let raw = "";
  try {
    raw = await fsp.readFile(filePath, "utf8");
    return JSON.parse(raw);
  } catch (error) {
    const badPath = await stashBrokenJson(filePath, raw);
    const backupPath = lastGoodPath(filePath);

    if (fs.existsSync(backupPath)) {
      try {
        const backupRaw = await fsp.readFile(backupPath, "utf8");
        const parsed = JSON.parse(backupRaw);
        await writeTextAtomic(filePath, backupRaw).catch(() => {});
        pushRecoveryNotice(
          "warning",
          "JSON_RECOVERED",
          "ระบบกู้ไฟล์ JSON อัตโนมัติ",
          `${path.basename(filePath)} เสีย ระบบกู้กลับจาก backup ล่าสุดแล้ว`,
          { file: path.basename(filePath), badFile: badPath ? path.basename(badPath) : "" }
        );
        return parsed;
      } catch {
        // ignore backup parse failure and fall through to fallback
      }
    }

    if (fallback !== null && fallback !== undefined) {
      try {
        await writeJson(filePath, fallback);
        pushRecoveryNotice(
          "warning",
          "JSON_RECOVERED_FALLBACK",
          "ระบบกู้ไฟล์ JSON อัตโนมัติ",
          `${path.basename(filePath)} เสีย ระบบกู้กลับด้วยค่า fallback แล้ว`,
          { file: path.basename(filePath), badFile: badPath ? path.basename(badPath) : "" }
        );
        return fallback;
      } catch {}
    }

    pushRecoveryNotice(
      "danger",
      "JSON_RECOVER_FAILED",
      "พบไฟล์ JSON เสีย",
      `${path.basename(filePath)} อ่านไม่ได้ ระบบจะใช้ค่า fallback และพักงานไว้ถ้าจำเป็น`,
      { file: path.basename(filePath), badFile: badPath ? path.basename(badPath) : "" }
    );
    return fallback;
  }
}


async function seedLastGoodBackup(filePath) {
  if (!shouldKeepLastGood(filePath)) return;
  if (!fs.existsSync(filePath) || fs.existsSync(lastGoodPath(filePath))) return;
  try {
    const raw = await fsp.readFile(filePath, 'utf8');
    JSON.parse(raw);
    await writeTextAtomic(lastGoodPath(filePath), raw);
  } catch {
    // ignore broken source during startup; auto-recover will handle later
  }
}

async function unlinkIfExists(filePath) {
  if (fs.existsSync(filePath)) await fsp.unlink(filePath);
}

async function appendRunLog(entry) {
  const current = await readJson(FILES.runLog, []);
  current.push({ at: nowIso(), ...entry });
  await writeJson(FILES.runLog, current.slice(-1000));
}

async function loadSettings() {
  const saved = await readJson(FILES.settings, DEFAULT_SETTINGS);
  return { ...DEFAULT_SETTINGS, ...(saved || {}) };
}

async function saveSettings(partial) {
  const current = await loadSettings();
  const next = {
    ...current,
    ...partial,
    batchSize: toPositiveNumber(partial.batchSize ?? current.batchSize, current.batchSize),
    maxContactsPerRun: toPositiveNumber(partial.maxContactsPerRun ?? current.maxContactsPerRun, current.maxContactsPerRun),
    delayBetweenRunsSec: toPositiveNumber(partial.delayBetweenRunsSec ?? current.delayBetweenRunsSec, current.delayBetweenRunsSec),
    retryPauseSec: toPositiveNumber(partial.retryPauseSec ?? current.retryPauseSec, current.retryPauseSec),
    retryRatioThreshold: Math.max(0, Math.min(1, Number(partial.retryRatioThreshold ?? current.retryRatioThreshold))),
    autoRun: Boolean(partial.autoRun ?? current.autoRun),
    waitFloodAutomatically: Boolean(partial.waitFloodAutomatically ?? current.waitFloodAutomatically),
  };
  await writeJson(FILES.settings, next);
  return next;
}

async function loadAccountsState() {
  const accounts = await readJson(FILES.accounts, []);
  const appState = await readJson(FILES.appState, { selectedAccountId: "" });
  return { accounts, appState };
}

async function saveAccountsState(accounts, appState) {
  await writeJson(FILES.accounts, accounts);
  await writeJson(FILES.appState, appState);
}

async function getAccounts() {
  const { accounts, appState } = await loadAccountsState();
  const selectedId = appState?.selectedAccountId || "";
  return accounts.map((account) => ({
    id: account.id,
    label: account.label,
    apiId: account.apiId,
    apiHashMasked: (() => {
      try {
        return maskSecret(decryptText(account.apiHashEnc), 4, 2);
      } catch {
        return "(decrypt_failed)";
      }
    })(),
    phone: account.phone,
    phoneMasked: maskSecret(account.phone, 4, 2),
    status: account.status || "new",
    hasSession: Boolean(account.sessionEnc) || fs.existsSync(sessionFilePath(account.id)),
    selected: account.id === selectedId,
    pendingCode: Boolean(account.pendingPhoneCodeHash),
    lastUsedAt: account.lastUsedAt || "",
    createdAt: account.createdAt || "",
    updatedAt: account.updatedAt || "",
    lastError: account.lastError || "",
    me: account.me || null,
  }));
}

async function getAccountById(accountId) {
  const { accounts, appState } = await loadAccountsState();
  const account = accounts.find((item) => item.id === accountId);
  return { account, accounts, appState };
}

async function updateAccountById(accountId, updater) {
  const { accounts, appState } = await loadAccountsState();
  const index = accounts.findIndex((item) => item.id === accountId);
  if (index === -1) throw new Error("ไม่พบบัญชีที่เลือก");
  const next = await updater({ ...accounts[index] }, accounts, appState);
  if (next) accounts[index] = next;
  await saveAccountsState(accounts, appState);
  return { account: accounts[index], accounts, appState };
}

async function createClient(apiId, apiHash, sessionString = "") {
  const client = new TelegramClient(new StringSession(sessionString), Number(apiId), String(apiHash), {
    connectionRetries: 5,
  });
  await client.connect();
  return client;
}

async function disconnectClientMap(map, accountId) {
  const client = map.get(accountId);
  map.delete(accountId);
  if (client) {
    try { await client.disconnect(); } catch {}
  }
}

async function markAccountForRelogin(accountId, message = "session ใช้งานไม่ได้ กรุณาล็อกอินใหม่", status = "needs_relogin") {
  try {
    await updateAccountById(accountId, (current) => ({
      ...current,
      status,
      sessionEnc: "",
      pendingPhoneCodeHash: "",
      pendingCodeType: "",
      pendingTimeout: 0,
      awaitingPassword: false,
      lastError: message,
      me: null,
      updatedAt: nowIso(),
    }));
  } catch {}
}

async function resetAccountSession(accountId, message = "ล้าง session แล้ว กรุณาล็อกอินใหม่") {
  await disconnectClientMap(activeClients, accountId);
  await disconnectClientMap(pendingAuthClients, accountId);
  await unlinkIfExists(sessionFilePath(accountId));
  await markAccountForRelogin(accountId, message, "needs_relogin");
}

async function ensureAuthorizedAccountClient(account) {
  const accountId = String(account?.id || "");
  if (!accountId) throw makeAccountAuthError("ไม่พบบัญชีที่เลือก", "auth_invalid");
  if (!String(account?.apiId || "").trim()) throw makeAccountAuthError("บัญชีนี้ยังไม่ได้ตั้งค่า API_ID กรุณาตรวจสอบบัญชี", "auth_invalid");

  let apiHash = "";
  try {
    apiHash = decryptText(account.apiHashEnc || "").trim();
  } catch {
    await markAccountForRelogin(accountId, "API_HASH ของบัญชีนี้ถอดรหัสไม่ได้ กรุณารีล็อกอินบัญชีใหม่", "auth_invalid");
    throw makeAccountAuthError("API_HASH ของบัญชีนี้ถอดรหัสไม่ได้ กรุณารีล็อกอินบัญชีใหม่", "auth_invalid");
  }
  if (!apiHash) {
    await markAccountForRelogin(accountId, "API_HASH ของบัญชีนี้ไม่ครบ กรุณารีล็อกอินบัญชีใหม่", "auth_invalid");
    throw makeAccountAuthError("API_HASH ของบัญชีนี้ไม่ครบ กรุณารีล็อกอินบัญชีใหม่", "auth_invalid");
  }

  const { sessionString } = await loadAccountSessionString(account);
  if (!sessionString) {
    await markAccountForRelogin(accountId, "บัญชีนี้ยังไม่มี session ที่ใช้งานได้ กรุณาล็อกอินใหม่", "needs_relogin");
    throw makeAccountAuthError("บัญชีนี้ยังไม่มี session ที่ใช้งานได้ กรุณาล็อกอินใหม่", "needs_relogin");
  }

  const cached = activeClients.get(account.id);
  if (cached) {
    try {
      if (await cached.checkAuthorization()) return cached;
    } catch {
      await disconnectClientMap(activeClients, account.id);
    }
  }

  let client = null;
  let authorized = false;
  try {
    client = await createClient(account.apiId, apiHash, sessionString);
    authorized = await client.checkAuthorization();
  } catch (error) {
    if (client) {
      try { await client.disconnect(); } catch {}
    }
    const friendly = "session ใช้งานไม่ได้ กรุณาล็อกอินใหม่";
    if (isAuthSessionError(error)) {
      await markAccountForRelogin(accountId, friendly, "needs_relogin");
      throw makeAccountAuthError(friendly, "needs_relogin");
    }
    throw error;
  }
  if (!authorized) {
    try { await client.disconnect(); } catch {}
    await markAccountForRelogin(accountId, "session ของบัญชีนี้หมดอายุแล้ว กรุณาล็อกอินใหม่", "needs_relogin");
    throw makeAccountAuthError("session ของบัญชีนี้หมดอายุแล้ว กรุณาล็อกอินใหม่", "needs_relogin");
  }

  await updateAccountById(accountId, (current) => ({
    ...current,
    status: "connected",
    lastError: "",
    updatedAt: nowIso(),
  })).catch(() => {});
  activeClients.set(account.id, client);
  return client;
}

function detectColumn(headers, type) {
  const phoneKeys = ["phone", "telephone", "tel", "mobile", "โทร", "เบอร์", "เบอร์โทร", "โทรศัพท์"];
  const nameKeys = ["name", "ชื่อ", "fullname", "full_name", "customer", "contact"];
  const list = type === "phone" ? phoneKeys : nameKeys;

  const normalized = headers.map((header) => ({
    original: header,
    safe: String(header || "").trim().toLowerCase(),
  }));

  for (const item of normalized) {
    if (list.some((k) => item.safe.includes(k))) return item.original;
  }
  return "";
}

function detectPhoneColumn(headers, rows) {
  const direct = detectColumn(headers, "phone");
  if (direct) return { column: direct, source: "header_match", confidence: 1, score: 0 };
  if (headers.length === 1) return { column: headers[0], source: "single_column", confidence: 0.5, score: 0 };
  return inferPhoneColumnBySamples(headers, rows, 20);
}

function parseTxtRows(rawText) {
  const lines = String(rawText || "").split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  return lines.map((line, idx) => ({
    __rowNumber: idx + 1,
    phone: line,
    name: "",
  }));
}

async function parseCsvRows(filePath) {
  const rows = [];
  await new Promise((resolve, reject) => {
    let rowNumber = 0;
    const parser = parseCsvStream({
      columns: true,
      skip_empty_lines: true,
      bom: true,
      trim: true,
      relax_column_count: true,
    });

    parser.on("readable", () => {
      let row;
      while ((row = parser.read()) !== null) {
        rowNumber += 1;
        rows.push({ __rowNumber: rowNumber, ...row });
      }
    });
    parser.on("error", reject);
    parser.on("end", resolve);

    fs.createReadStream(filePath).on("error", reject).pipe(parser);
  });
  return rows;
}

function parseWorkbookRows(filePath) {
  const workbook = XLSX.readFile(filePath, { raw: false, dense: true });
  const firstSheet = workbook.SheetNames[0];
  if (!firstSheet) throw new Error("ไฟล์ Excel ไม่มี sheet ที่อ่านได้");
  const sheet = workbook.Sheets[firstSheet];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  return rows.map((row, idx) => ({ __rowNumber: idx + 1, ...row }));
}

function normalizeInputRows(rows) {
  if (!Array.isArray(rows) || !rows.length) {
    return { cleanRows: [], invalidRows: [], duplicateRows: [], rejectRows: [], debugRows: [], summary: null };
  }

  const headers = Object.keys(rows[0] || {}).filter((key) => key !== "__rowNumber");
  const phoneDetection = detectPhoneColumn(headers, rows);
  const phoneColumn = phoneDetection.column;
  let nameColumn = detectColumn(headers, "name");

  const cleanRows = [];
  const invalidRows = [];
  const duplicateRows = [];
  const rejectRows = [];
  const debugRows = [];
  const seen = new Set();

  rows.forEach((row, index) => {
    const rowNumber = Number(row.__rowNumber || index + 1);
    const rawPhone = phoneColumn ? String(row[phoneColumn] || "").trim() : "";
    const normalizedResult = normalizeThaiPhoneDetailed(rawPhone);
    const normalizedPhone = normalizedResult.normalized;
    const nameRaw = nameColumn ? String(row[nameColumn] || "").trim() : "";
    const name = nameRaw || `C${String(cleanRows.length + 1).padStart(6, "0")}`;
    let status = "READY";
    let reason = "";

    if (!normalizedPhone) {
      status = "REJECT";
      reason = normalizedResult.reason || "invalid_phone";
      invalidRows.push({ rowNumber, name: nameRaw, rawPhone, reason });
      rejectRows.push({
        rowNumber,
        name: nameRaw,
        rawPhone,
        sanitizedPhone: normalizedResult.sanitized || "",
        digits: normalizedResult.digits || "",
        reason,
      });
      debugRows.push({
        rowNumber,
        detectedPhoneColumn: phoneColumn || "",
        rawPhone,
        sanitizedPhone: normalizedResult.sanitized || "",
        digits: normalizedResult.digits || "",
        normalizedPhone: "",
        status,
        reason,
      });
      return;
    }

    if (seen.has(normalizedPhone)) {
      status = "DUPLICATE";
      reason = "duplicate_phone";
      duplicateRows.push({ rowNumber, name, rawPhone, normalizedPhone, reason });
      rejectRows.push({
        rowNumber,
        name,
        rawPhone,
        sanitizedPhone: normalizedResult.sanitized || "",
        digits: normalizedResult.digits || "",
        normalizedPhone,
        reason,
      });
      debugRows.push({
        rowNumber,
        detectedPhoneColumn: phoneColumn || "",
        rawPhone,
        sanitizedPhone: normalizedResult.sanitized || "",
        digits: normalizedResult.digits || "",
        normalizedPhone,
        status,
        reason,
      });
      return;
    }

    seen.add(normalizedPhone);
    cleanRows.push({
      rowNumber,
      sourceIndex: cleanRows.length + 1,
      name,
      rawPhone,
      normalizedPhone,
      phone: normalizedPhone,
    });
    debugRows.push({
      rowNumber,
      detectedPhoneColumn: phoneColumn || "",
      rawPhone,
      sanitizedPhone: normalizedResult.sanitized || "",
      digits: normalizedResult.digits || "",
      normalizedPhone,
      status,
      reason,
    });
  });

  const summary = {
    originalRows: rows.length,
    totalRows: rows.length,
    readyRows: cleanRows.length,
    invalidRows: invalidRows.length,
    duplicateRows: duplicateRows.length,
    phoneColumn: phoneColumn || "",
    detectedPhoneColumn: phoneColumn || "",
    phoneColumnDetectionSource: phoneDetection.source || "unknown",
    phoneColumnConfidence: Number(phoneDetection.confidence || 0),
    nameColumn,
    generatedAt: nowIso(),
  };

  return { cleanRows, invalidRows, duplicateRows, rejectRows, debugRows, summary };
}

function cleanRowsToCsv(rows) {
  return stringify(rows.map((row) => ({
    name: row.name,
    phone: row.phone,
    rawPhone: row.rawPhone,
    rowNumber: row.rowNumber,
  })), {
    header: true,
    columns: ["name", "phone", "rawPhone", "rowNumber"],
  });
}

function simpleRowsToCsv(rows, columns) {
  return stringify(rows.map((row) => {
    const out = {};
    columns.forEach((c) => out[c] = row[c] ?? "");
    return out;
  }), { header: true, columns });
}

function dynamicRowsToCsv(rows, preferredColumns = []) {
  const list = Array.isArray(rows) ? rows : [];
  const dynamicColumns = new Set(preferredColumns);
  for (const row of list) {
    for (const key of Object.keys(row || {})) dynamicColumns.add(key);
  }
  const columns = Array.from(dynamicColumns);
  return stringify(
    list.map((row) => {
      const out = {};
      for (const c of columns) out[c] = row?.[c] ?? "";
      return out;
    }),
    { header: true, columns }
  );
}

async function saveIntakeOutputs(filename, result) {
  const meta = {
    originalFilename: filename,
    ...result.summary,
    files: {
      cleanReadyCsv: "clean_ready.csv",
      invalidRowsCsv: "invalid_rows.csv",
      duplicatePhonesCsv: "duplicate_phones.csv",
      cleanDebugCsv: "clean_debug.csv",
      cleanRejectsCsv: "clean_rejects.csv",
      summaryJson: "summary.json",
    },
    preview: result.cleanRows.slice(0, 30),
  };
  await fsp.writeFile(FILES.cleanReadyCsv, cleanRowsToCsv(result.cleanRows), "utf8");
  await fsp.writeFile(FILES.invalidRowsCsv, simpleRowsToCsv(result.invalidRows, ["rowNumber", "name", "rawPhone", "reason"]), "utf8");
  await fsp.writeFile(FILES.duplicatePhonesCsv, simpleRowsToCsv(result.duplicateRows, ["rowNumber", "name", "rawPhone", "normalizedPhone", "reason"]), "utf8");
  await fsp.writeFile(
    FILES.cleanRejectsCsv,
    dynamicRowsToCsv(result.rejectRows, ["rowNumber", "name", "rawPhone", "sanitizedPhone", "digits", "normalizedPhone", "reason"]),
    "utf8"
  );
  await fsp.writeFile(
    FILES.cleanDebugCsv,
    dynamicRowsToCsv(result.debugRows, ["rowNumber", "detectedPhoneColumn", "rawPhone", "sanitizedPhone", "digits", "normalizedPhone", "status", "reason"]),
    "utf8"
  );
  await writeJson(FILES.intakeSummaryJson, meta);
  await writeJson(FILES.intakeLatest, meta);
  await appendRunLog({
    type: "clean_summary",
    totalRows: meta.totalRows || 0,
    readyRows: meta.readyRows || 0,
    invalidRows: meta.invalidRows || 0,
    duplicateRows: meta.duplicateRows || 0,
    detectedPhoneColumn: meta.detectedPhoneColumn || "",
  });
  logInfo(
    `Cleaner summary total=${meta.totalRows || 0} ready=${meta.readyRows || 0} invalid=${meta.invalidRows || 0} duplicate=${meta.duplicateRows || 0} phoneColumn=${meta.detectedPhoneColumn || "-"}`
  );
  return meta;
}

async function loadCleanRows() {
  if (!fs.existsSync(FILES.cleanReadyCsv)) return [];
  const raw = await fsp.readFile(FILES.cleanReadyCsv, "utf8");
  const rows = parseCsvSync(raw, { columns: true, skip_empty_lines: true, bom: true, trim: true });
  return rows.map((row, idx) => ({
    sourceIndex: idx + 1,
    name: String(row.name || "").trim(),
    rawPhone: String(row.rawPhone || row.phone || "").trim(),
    normalizedPhone: String(row.phone || "").trim(),
  }));
}

function buildContacts(rows) {
  return rows.map((row, index) => {
    const clientId = BigInt(Date.now()) * 1000n + BigInt(index + 1);
    return {
      ...row,
      clientId,
      contact: new Api.InputPhoneContact({
        clientId,
        phone: row.normalizedPhone,
        firstName: row.name || `Contact ${row.sourceIndex}`,
        lastName: "",
      }),
    };
  });
}

async function importBatch(client, batchRows) {
  const contacts = buildContacts(batchRows);
  const result = await client.invoke(
    new Api.contacts.ImportContacts({
      contacts: contacts.map((item) => item.contact),
    })
  );

  const usersById = new Map();
  for (const user of result.users || []) usersById.set(toSafeStringId(user.id), user);

  const matchedByClientId = new Map();
  for (const imported of result.imported || []) {
    const clientId = toSafeStringId(imported.clientId);
    const userId = toSafeStringId(imported.userId);
    const user = usersById.get(userId);
    matchedByClientId.set(clientId, {
      userId,
      username: user?.username || "",
      firstName: user?.firstName || "",
      lastName: user?.lastName || "",
      phone: user?.phone ? `+${onlyDigits(user.phone)}` : "",
    });
  }

  const retrySet = new Set((result.retryContacts || []).map((value) => toSafeStringId(value)));

  const rows = contacts.map((item) => {
    const clientId = toSafeStringId(item.clientId);
    const matched = matchedByClientId.get(clientId);
    const isRetry = retrySet.has(clientId);

    return {
      sourceIndex: item.sourceIndex,
      name: item.name,
      rawPhone: item.rawPhone,
      normalizedPhone: item.normalizedPhone,
      status: matched ? "YES" : isRetry ? "RETRY" : "NO",
      hasTelegram: Boolean(matched),
      retry: isRetry,
      telegramUserId: matched?.userId || "",
      telegramUsername: matched?.username || "",
      telegramFirstName: matched?.firstName || "",
      telegramLastName: matched?.lastName || "",
      telegramPhone: matched?.phone || "",
      clientId,
    };
  });

  return {
    rows,
    importedCount: Array.isArray(result.imported) ? result.imported.length : 0,
    retryContacts: Array.from(retrySet),
  };
}

function rowsToCsv(rows) {
  return stringify(
    (rows || []).map((row) => ({
      sourceIndex: row.sourceIndex,
      name: row.name,
      rawPhone: row.rawPhone,
      normalizedPhone: row.normalizedPhone,
      status: row.status,
      telegramUserId: row.telegramUserId,
      telegramUsername: row.telegramUsername,
      telegramFirstName: row.telegramFirstName,
      telegramLastName: row.telegramLastName,
      telegramPhone: row.telegramPhone,
      clientId: row.clientId,
    })),
    {
      header: true,
      columns: [
        "sourceIndex",
        "name",
        "rawPhone",
        "normalizedPhone",
        "status",
        "telegramUserId",
        "telegramUsername",
        "telegramFirstName",
        "telegramLastName",
        "telegramPhone",
        "clientId",
      ],
    }
  );
}

async function saveLatestOutputs(payload) {
  await writeJson(FILES.latestJson, payload);
  await fsp.writeFile(FILES.latestCsv, rowsToCsv(payload.rows || []), "utf8");
}

async function saveAllOutputs(rows) {
  await writeJson(FILES.allJson, rows || []);
  await fsp.writeFile(FILES.allCsv, rowsToCsv(rows || []), "utf8");
  const retryRows = (rows || []).filter((row) => row.status === "RETRY");
  await fsp.writeFile(FILES.retryCsv, rowsToCsv(retryRows), "utf8");
}

async function readLatestMeta() {
  const latest = await readJson(FILES.latestJson, null);
  return latest?.meta || null;
}

async function getProcessedSourceIndexSet() {
  const allRows = await readJson(FILES.allJson, []);
  return new Set((Array.isArray(allRows) ? allRows : []).map((row) => String(row?.sourceIndex || '')).filter(Boolean));
}

async function writeRemainingCsv(job) {
  if (!job || !Array.isArray(job.rows)) {
    await unlinkIfExists(FILES.remainingCsv);
    return [];
  }
  const doneSet = await getProcessedSourceIndexSet();
  const remainingRows = job.rows.filter((row) => !doneSet.has(String(row?.sourceIndex || '')));
  await fsp.writeFile(FILES.remainingCsv, cleanRowsToCsv(remainingRows), 'utf8');
  return remainingRows;
}

function buildAlerts(job, latestMeta, settings, runLog = [], notices = []) {
  const alerts = [];
  const remainingSec = computeRemainingSec(job?.nextRunAt);
  const latestNotice = Array.isArray(notices) && notices.length ? notices[notices.length - 1] : null;
  if (latestNotice) {
    alerts.push({
      level: latestNotice.level || 'warning',
      code: latestNotice.code || 'RECOVERY',
      title: latestNotice.title || 'ระบบกู้คืนอัตโนมัติ',
      message: latestNotice.message || ''
    });
  }
  if (!job) {
    alerts.push({ level: 'info', code: 'NO_JOB', title: 'ยังไม่มีคิวงาน', message: 'ล้างไฟล์และสร้างคิวก่อน ระบบถึงจะเริ่มวิ่งได้' });
    return alerts;
  }

  if (displayAutoStatus(job?.autoStatus) === 'WAITING_FLOOD') {
    alerts.push({
      level: 'warning',
      code: 'WAITING_FLOOD',
      title: 'กำลังรอ FloodWait',
      message: `รออีก ${remainingSec} วินาที${settings?.autoRun ? ' และระบบจะทำต่อเอง' : ' แล้วค่อยกดทำต่อ'}`
    });
  }

  if (displayQueueStatus(job?.status) === 'WAITING_RETRY') {
    alerts.push({
      level: 'warning',
      code: 'WAITING_RETRY',
      title: 'กำลังพักก่อนลอง RETRY ใหม่',
      message: `รออีก ${remainingSec} วินาที เพื่อให้ระบบวนกลับไปลองรายการ RETRY`
    });
  }

  if (displayQueueStatus(job?.status) === 'PAUSED') {
    alerts.push({ level: 'warning', code: 'PAUSED', title: 'คิวถูกพักอยู่', message: 'กดทำต่อออโต้หรือทำต่อรอบถัดไปเมื่อพร้อม' });
  }

  if (displayQueueStatus(job?.status) === 'PAUSED_FLOOD_STALE') {
    alerts.push({
      level: 'warning',
      code: 'PAUSED_FLOOD_STALE',
      title: 'คิวโดนพักจาก FloodWait ค้างนาน',
      message: 'ตรวจสอบบัญชีและกดรีเซ็ต flood state ก่อนเริ่มใหม่'
    });
  }

  if (displayQueueStatus(job?.status) === 'COMPLETED') {
    alerts.push({ level: 'success', code: 'COMPLETED', title: 'คิวนี้ทำครบแล้ว', message: 'ดาวน์โหลดผลรวม หรือสร้างคิวใหม่จากไฟล์ที่เหลือได้เลย' });
  }

  if (Number(job?.lastRetryRatio || 0) > Number(settings?.retryRatioThreshold || 0.2)) {
    alerts.push({ level: 'warning', code: 'RETRY_RATIO_HIGH', title: 'RETRY สูงกว่าค่าที่ตั้งไว้', message: `retry ratio ล่าสุด ${Number(job?.lastRetryRatio || 0).toFixed(2)} สูงกว่าค่า ${Number(settings?.retryRatioThreshold || 0.2).toFixed(2)}` });
  }

  const latestError = (Array.isArray(runLog) ? [...runLog].reverse() : []).find((item) => ['auto_error', 'account_error'].includes(item?.type));
  if (latestError?.message) {
    alerts.push({ level: 'danger', code: 'LAST_ERROR', title: 'มีปัญหาล่าสุด', message: latestError.message });
  }

  if (!alerts.length && latestMeta) {
    alerts.push({ level: 'info', code: 'RUNNING_OK', title: 'ระบบกำลังทำงาน', message: `ล่าสุดทำถึง ${latestMeta.processedRows || 0}/${latestMeta.totalRows || 0} และสถานะ ${displayAutoStatus(job?.autoStatus)}` });
  }

  return alerts.slice(0, 5);
}

async function buildDashboardPayload() {
  const [settings, job, latestMeta, runLog] = await Promise.all([
    loadSettings(),
    getCurrentJob(),
    readLatestMeta(),
    readJson(FILES.runLog, []),
  ]);

  const remainingSec = computeRemainingSec(job?.nextRunAt);
  const progressPct = job?.totalRows ? Math.min(100, Math.round((Number(job.processedRows || 0) / Number(job.totalRows || 1)) * 100)) : 0;
  const runEvents = (Array.isArray(runLog) ? [...runLog] : []).slice(-15).reverse().map((entry) => ({
    at: entry.at || '',
    type: entry.type || 'event',
    text: entry.message || entry.code || [entry.type, entry.batchNumber ? `รอบ ${entry.batchNumber}` : '', Number.isFinite(entry.seconds) ? `${entry.seconds} วิ` : ''].filter(Boolean).join(' • '),
  }));
  const noticeEvents = [...recoveryNotices].slice(-10).reverse().map((entry) => ({
    at: entry.at || '',
    type: entry.code || 'recovery',
    text: entry.message || ''
  }));
  const events = [...noticeEvents, ...runEvents].slice(0, 15);

  return {
    now: nowIso(),
    settings,
    job: job ? {
      ...job,
      displayStatus: displayQueueStatus(job.status),
      displayAutoStatus: displayAutoStatus(job.autoStatus),
      remainingSec,
      progressPct,
    } : null,
    latestMeta,
    cards: {
      done: Number(job?.processedRows || 0),
      remaining: Number(job?.remainingRows || 0),
      yes: Number(job?.matchedCountTotal || 0),
      no: Number(job?.unmatchedCountTotal || 0),
      retry: Number(job?.retryCountTotal || 0),
      floodWaitCount: Number(job?.floodWaitCount || 0),
      progressPct,
      account: job?.lockedAccountLabel || '-',
      queueStatus: displayQueueStatus(job?.status),
    },
    alerts: buildAlerts(job, latestMeta, settings, runLog, recoveryNotices),
    events,
    recovery: [...recoveryNotices].slice(-10).reverse(),
  };
}

async function buildAutoState(job) {
  const settings = await loadSettings();
  const remainingSec = computeRemainingSec(job?.nextRunAt);
  return {
    autoRun: Boolean(settings.autoRun),
    status: displayAutoStatus(job?.autoStatus || (settings.autoRun ? "RUNNING" : "OFF")),
    queueStatus: displayQueueStatus(job?.status || "idle"),
    nextRunAt: job?.nextRunAt || "",
    waitUntil: job?.nextRunAt || "",
    remainingSec,
    nextAction: settings.autoRun ? "auto_resume" : "manual_resume",
    lockedAccountLabel: job?.lockedAccountLabel || "",
    lockedAccountPhone: job?.lockedAccountPhone || "",
    file: job?.sourceFile || "",
    floodWaitSec: displayQueueStatus(job?.status) === "WAITING_FLOOD" ? remainingSec : (job?.lastFloodWaitSec || 0),
    floodWaitCount: Number(job?.floodWaitCount || 0),
    retryRatio: job?.lastRetryRatio || 0,
    processedRows: job?.processedRows || 0,
    remainingRows: job?.remainingRows || 0,
    updatedAt: nowIso(),
  };
}


async function getCurrentJob() {
  const [jobState, currentJob] = await Promise.all([
    readJson(FILES.jobState, null),
    readJson(FILES.currentJob, null),
  ]);

  const job = jobState || currentJob;
  if (!job) return null;

  const normalized = {
    ...job,
    status: job.status || "paused",
    autoStatus: job.autoStatus || (job.status === "completed" ? "OFF" : "PAUSED"),
    floodWaitCount: Number(job.floodWaitCount || 0),
    consecutiveFloodCount: Number(job.consecutiveFloodCount || 0),
    rows: Array.isArray(job.rows) ? job.rows : (Array.isArray(currentJob?.rows) ? currentJob.rows : job.rows),
  };

  if (jobState && !currentJob) {
    await writeJson(FILES.currentJob, normalized).catch(() => {});
  } else if (currentJob && !jobState) {
    await writeJson(FILES.jobState, normalized).catch(() => {});
  }

  return normalized;
}

async function saveCurrentJob(job) {
  await writeJson(FILES.currentJob, job);
  await writeJson(FILES.jobState, job);
  await writeJson(FILES.autoState, await buildAutoState(job));
  await writeRemainingCsv(job).catch(() => {});
  return job;
}

async function getSelectedAccount(requireConnected = false) {
  const { accounts, appState } = await loadAccountsState();
  const selectedId = appState?.selectedAccountId || "";
  if (!selectedId) throw new Error("ยังไม่ได้เลือกบัญชี Telegram");
  const account = accounts.find((item) => item.id === selectedId);
  if (!account) throw new Error("ไม่พบบัญชีที่เลือก");
  if (requireConnected) {
    const hasSession = Boolean(account.sessionEnc) || fs.existsSync(sessionFilePath(account.id));
    if (!hasSession) throw new Error("บัญชีที่เลือกยังไม่ได้เชื่อม Telegram");
  }
  return account;
}

async function getLockedAccountForJob(job) {
  const { account } = await getAccountById(job.lockedAccountId);
  if (!account) throw new Error("ไม่พบบัญชีที่ล็อกกับคิวนี้");
  const hasSession = Boolean(account.sessionEnc) || fs.existsSync(sessionFilePath(account.id));
  if (!hasSession) throw new Error("บัญชีที่ล็อกกับคิวนี้ยังไม่ได้เชื่อม Telegram");
  return account;
}

async function createJobFromLatestClean() {
  const cleanRows = await loadCleanRows();
  if (!cleanRows.length) throw new Error("ยังไม่มี clean_ready.csv กรุณาล้างไฟล์ก่อน");
  const selectedAccount = await getSelectedAccount(true);
  const settings = await loadSettings();

  const job = {
    id: makeId("job"),
    sourceFile: "clean_ready.csv",
    totalRows: cleanRows.length,
    nextIndex: 0,
    processedRows: 0,
    remainingRows: cleanRows.length,
    lockedAccountId: selectedAccount.id,
    lockedAccountLabel: selectedAccount.label,
    lockedAccountPhone: selectedAccount.phone,
    perRunLimit: settings.maxContactsPerRun,
    chunkSize: settings.batchSize,
    delayBetweenRunsSec: settings.delayBetweenRunsSec,
    retryPauseSec: settings.retryPauseSec,
    retryRatioThreshold: settings.retryRatioThreshold,
    waitFloodAutomatically: settings.waitFloodAutomatically,
    matchedCountTotal: 0,
    unmatchedCountTotal: 0,
    retryCountTotal: 0,
    lastBatchNumber: 0,
    lastBatchStart: 0,
    lastBatchEnd: 0,
    lastBatchSize: 0,
    lastGeneratedAt: "",
    lastFloodWaitSec: 0,
    lastFloodWaitBaseSec: 0,
    floodWaitCount: 0,
    consecutiveFloodCount: 0,
    lastFloodAt: "",
    lastRetryRatio: 0,
    nextRunAt: "",
    autoStatus: settings.autoRun ? "RUNNING" : "OFF",
    status: "ready",
    rows: cleanRows,
    createdAt: nowIso(),
    updatedAt: nowIso(),
  };

  await saveAllOutputs([]);
  await saveCurrentJob(job);
  await unlinkIfExists(FILES.latestJson);
  await unlinkIfExists(FILES.latestCsv);
  await appendRunLog({ type: "job_created", jobId: job.id, lockedAccount: job.lockedAccountLabel });

  return job;
}

async function processNextBatch() {
  const job = await getCurrentJob();
  if (!job) throw new Error("ยังไม่มีคิวงาน กรุณาล้างไฟล์และสร้างคิวก่อน");
  if (!Array.isArray(job.rows) || !job.rows.length) throw new Error("ข้อมูลคิวหาย กรุณาสร้างคิวใหม่");
  if (job.status === "completed") throw new Error("คิวนี้ทำครบแล้ว");
  if (job.status === "paused") throw new Error("คิวนี้ถูกพักอยู่");
  if (job.status === "paused_flood_stale") throw new Error("คิวนี้ถูกพักเพราะ FloodWait ค้างนาน กรุณารีเซ็ต flood state ก่อน");
  if (job.status === "waiting_flood" && job.nextRunAt && Date.now() < new Date(job.nextRunAt).getTime()) {
    throw new Error(`ยังต้องรอถึง ${job.nextRunAt}`);
  }

  const selected = await getSelectedAccount(false);
  if (!selected?.id) throw new Error("ยังไม่ได้เลือกบัญชี Telegram");
  if (selected.id !== job.lockedAccountId) {
    throw new Error("บัญชีที่เลือกอยู่ไม่ตรงกับบัญชีที่ล็อกกับคิวนี้ กรุณากด 'ใช้บัญชีนี้' ให้ตรงก่อนรัน");
  }

  const account = await getLockedAccountForJob(job);
  const client = await ensureAuthorizedAccountClient(account);

  const start = Number(job.nextIndex || 0);
  const end = Math.min(start + Number(job.perRunLimit || 100), job.rows.length);
  const sourceBatch = job.rows.slice(start, end);
  if (!sourceBatch.length) throw new Error("ไม่พบข้อมูลในรอบถัดไป");

  const chunks = chunkArray(sourceBatch, Number(job.chunkSize || 1));
  const currentBatchRows = [];
  let importedCount = 0;
  let batchRetryCount = 0;

  for (const chunk of chunks) {
    const result = await importBatch(client, chunk);
    currentBatchRows.push(...result.rows);
    importedCount += result.importedCount;
    batchRetryCount += result.rows.filter((row) => row.status === "RETRY").length;
  }

  const matchedCount = currentBatchRows.filter((row) => row.status === "YES").length;
  const retryCount = currentBatchRows.filter((row) => row.status === "RETRY").length;
  const unmatchedCount = currentBatchRows.filter((row) => row.status === "NO").length;

  const allRows = await readJson(FILES.allJson, []);
  allRows.push(...currentBatchRows);
  await saveAllOutputs(allRows);

  const now = nowIso();
  const nextIndex = end;
  const batchNumber = Math.floor(start / Number(job.perRunLimit || 100)) + 1;
  const retryRatio = currentBatchRows.length ? retryCount / currentBatchRows.length : 0;

  let nextStatus = nextIndex >= job.rows.length ? "completed" : "ready";
  let nextAutoStatus = job.autoStatus || "OFF";
  let nextRunAt = "";

  if (nextStatus !== "completed" && retryRatio > Number(job.retryRatioThreshold || 0.2)) {
    nextStatus = "waiting_retry_cooldown";
    nextAutoStatus = "WAITING_RETRY";
    nextRunAt = new Date(Date.now() + Number(job.retryPauseSec || 300) * 1000).toISOString();
  } else if (nextStatus !== "completed" && nextAutoStatus === "RUNNING") {
    nextRunAt = new Date(Date.now() + Number(job.delayBetweenRunsSec || 60) * 1000).toISOString();
  }

  const nextJob = {
    ...job,
    nextIndex,
    processedRows: nextIndex,
    remainingRows: Math.max(0, Number(job.totalRows || 0) - nextIndex),
    matchedCountTotal: Number(job.matchedCountTotal || 0) + matchedCount,
    unmatchedCountTotal: Number(job.unmatchedCountTotal || 0) + unmatchedCount,
    retryCountTotal: Number(job.retryCountTotal || 0) + retryCount,
    lastBatchNumber: batchNumber,
    lastBatchStart: start + 1,
    lastBatchEnd: end,
    lastBatchSize: currentBatchRows.length,
    lastGeneratedAt: now,
    lastFloodWaitSec: 0,
    lastFloodWaitBaseSec: 0,
    lastRetryRatio: retryRatio,
    consecutiveFloodCount: 0,
    nextRunAt,
    status: nextStatus,
    autoStatus: nextStatus === "completed" ? "OFF" : nextAutoStatus,
    updatedAt: now,
  };

  const latestPayload = {
    meta: {
      sourceFile: nextJob.sourceFile,
      batchNumber,
      batchStart: start + 1,
      batchEnd: end,
      batchSize: currentBatchRows.length,
      totalRows: nextJob.totalRows,
      processedRows: nextJob.processedRows,
      remainingRows: nextJob.remainingRows,
      matchedCount,
      unmatchedCount,
      retryCount,
      matchedCountTotal: nextJob.matchedCountTotal,
      unmatchedCountTotal: nextJob.unmatchedCountTotal,
      retryCountTotal: nextJob.retryCountTotal,
      importedCount,
      generatedAt: now,
      done: nextJob.status === "completed",
      perRunLimit: nextJob.perRunLimit,
      chunkSize: nextJob.chunkSize,
      accountId: account.id,
      accountLabel: account.label,
      accountPhone: account.phone,
      retryRatio,
      autoStatus: nextJob.autoStatus,
      nextRunAt: nextJob.nextRunAt,
      floodWaitCount: Number(nextJob.floodWaitCount || 0),
    },
    rows: currentBatchRows,
  };

  await saveLatestOutputs(latestPayload);
  await saveCurrentJob(nextJob);
  await appendRunLog({
    type: "batch_done",
    jobId: job.id,
    batchNumber,
    matchedCount,
    unmatchedCount,
    retryCount,
    retryRatio,
    processedRows: nextJob.processedRows,
    floodWaitCount: Number(nextJob.floodWaitCount || 0),
    nextStatus,
  });
  await appendRunLog({
    type: "run_summary",
    processed: nextJob.processedRows,
    yes: nextJob.matchedCountTotal,
    no: nextJob.unmatchedCountTotal,
    retry: nextJob.retryCountTotal,
    floodWaitCount: Number(nextJob.floodWaitCount || 0),
  });
  logInfo(
    `Run summary processed=${nextJob.processedRows} yes=${nextJob.matchedCountTotal} no=${nextJob.unmatchedCountTotal} retry=${nextJob.retryCountTotal} floodWaitCount=${Number(nextJob.floodWaitCount || 0)}`
  );

  return { latestBatch: latestPayload, job: nextJob };
}

async function markFloodStateStaleIfNeeded(job) {
  if (!job || job.status !== "waiting_flood") return job;
  const ageSec = floodStaleAgeSec(job);
  if (ageSec < FLOOD_STALE_THRESHOLD_SEC) return job;
  const nextJob = {
    ...job,
    status: "paused_flood_stale",
    autoStatus: "PAUSED_FLOOD_STALE",
    nextRunAt: "",
    updatedAt: nowIso(),
  };
  await saveCurrentJob(nextJob);
  await appendRunLog({
    type: "flood_stale_paused",
    jobId: job.id,
    accountLabel: job.lockedAccountLabel || "",
    accountPhone: job.lockedAccountPhone || "",
    staleAgeSec: ageSec,
    thresholdSec: FLOOD_STALE_THRESHOLD_SEC,
  });
  logInfo(`Flood stale paused job=${job.id} ageSec=${ageSec} thresholdSec=${FLOOD_STALE_THRESHOLD_SEC}`);
  return nextJob;
}

async function handleFloodWait(job, error) {
  const settings = await loadSettings();
  const baseFloodSeconds = parseFloodSeconds(error?.message || error?.errorMessage || error) || Number(job.delayBetweenRunsSec || settings.delayBetweenRunsSec || 120);
  const consecutiveFloodCount = Number(job.consecutiveFloodCount || 0) + 1;
  const appliedFloodSec = computeFloodBackoffSec(baseFloodSeconds, consecutiveFloodCount);
  const nextRunAt = new Date(Date.now() + Math.max(1, appliedFloodSec) * 1000).toISOString();
  const nextJob = {
    ...job,
    status: settings.waitFloodAutomatically ? "waiting_flood" : "paused",
    autoStatus: settings.waitFloodAutomatically ? "WAITING_FLOOD" : "PAUSED",
    lastFloodWaitSec: appliedFloodSec,
    lastFloodWaitBaseSec: baseFloodSeconds,
    lastFloodAt: nowIso(),
    consecutiveFloodCount,
    floodWaitCount: Number(job.floodWaitCount || 0) + 1,
    nextRunAt,
    updatedAt: nowIso(),
  };
  await saveCurrentJob(nextJob);
  await appendRunLog({
    type: "flood_wait",
    jobId: job.id,
    seconds: appliedFloodSec,
    baseSeconds: baseFloodSeconds,
    consecutiveFloodCount,
    accountLabel: job.lockedAccountLabel || "",
    accountPhone: job.lockedAccountPhone || "",
    retryAt: nextRunAt,
    nextRunAt,
  });
  logInfo(
    `FloodWait account=${job.lockedAccountLabel || "-"} base=${baseFloodSeconds}s applied=${appliedFloodSec}s consecutive=${consecutiveFloodCount} retryAt=${nextRunAt}`
  );
  return nextJob;
}

const storage = multer.diskStorage({
  destination: (_req, _file, callback) => callback(null, RAW_DIR),
  filename: (_req, file, callback) => callback(null, `${Date.now()}-${sanitizeFileName(file.originalname)}`),
});

const upload = multer({
  storage,
  limits: { fileSize: MAX_UPLOAD_MB * 1024 * 1024 },
});

function isSupportedIntakeExt(ext) {
  return [".csv", ".txt", ".xlsx", ".xls", ".xlsm", ".xlsb", ".ods"].includes(ext);
}

async function parseInputRowsFromFile(filePath, originalName = "") {
  const ext = path.extname(originalName || filePath || "").toLowerCase();
  if (!isSupportedIntakeExt(ext)) {
    throw new Error("รองรับเฉพาะ csv, txt, xlsx, xls, xlsm, xlsb, ods");
  }
  if (ext === ".txt") {
    const rawText = await fsp.readFile(filePath, "utf8");
    return parseTxtRows(rawText);
  }
  if (ext === ".csv") return parseCsvRows(filePath);
  return parseWorkbookRows(filePath); // Excel/ODS: read first sheet only
}

async function processIntakeFile(filePath, originalName = "") {
  const rows = await parseInputRowsFromFile(filePath, originalName);
  const result = normalizeInputRows(rows);
  const meta = await saveIntakeOutputs(originalName || path.basename(filePath), result);
  return { meta, result };
}

async function splitCsvFile(filePath, rowsPerFile = 50000) {
  const safeRowsPerFile = Math.max(1000, Math.floor(Number(rowsPerFile || 50000)));
  const outDir = path.join(RAW_DIR, `split_${Date.now()}`);
  await fsp.mkdir(outDir, { recursive: true });

  const parser = parseCsvStream({
    columns: true,
    skip_empty_lines: true,
    bom: true,
    trim: true,
    relax_column_count: true,
  });
  const stream = fs.createReadStream(filePath).pipe(parser);
  const outFiles = [];
  let batch = [];
  let totalRows = 0;
  let part = 1;

  async function flushBatch() {
    if (!batch.length) return;
    const columns = Object.keys(batch[0] || {});
    const csv = stringify(batch, { header: true, columns });
    const outPath = path.join(outDir, `${path.basename(filePath, path.extname(filePath))}.part${String(part).padStart(3, "0")}.csv`);
    await fsp.writeFile(outPath, csv, "utf8");
    outFiles.push(outPath);
    part += 1;
    batch = [];
  }

  for await (const row of stream) {
    batch.push(row);
    totalRows += 1;
    if (batch.length >= safeRowsPerFile) await flushBatch();
  }
  await flushBatch();

  return { outDir, files: outFiles, totalRows, rowsPerFile: safeRowsPerFile };
}

app.use(express.json({ limit: "2mb" }));
app.use(express.urlencoded({ extended: true }));

app.post("/auth/login", (req, res) => {
  const { username, password } = req.body || {};
  if (username !== APP_USER || password !== APP_PASSWORD) {
    return res.status(401).json({ ok: false, message: "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง" });
  }
  const token = signToken({ u: username, exp: Date.now() + 1000 * 60 * 60 * 12 });
  setAuthCookie(res, token);
  return res.json({ ok: true, username });
});

app.post("/auth/logout", (_req, res) => {
  clearAuthCookie(res);
  return res.json({ ok: true });
});

app.get("/auth/me", (req, res) => {
  const cookies = parseCookies(req);
  const payload = verifyToken(cookies[AUTH_COOKIE]);
  if (!payload) return res.json({ loggedIn: false });
  return res.json({ loggedIn: true, username: payload.u });
});

app.use("/download", requireAuth);
app.use("/api", (req, res, next) => {
  if (req.path === "/health") return next();
  return requireAuth(req, res, next);
});

app.get("/login", (_req, res) => res.sendFile(path.join(PUBLIC_DIR, "login.html")));
app.get("/", (req, res) => {
  const cookies = parseCookies(req);
  const payload = verifyToken(cookies[AUTH_COOKIE]);
  if (!payload) return res.redirect("/login");
  return res.sendFile(path.join(PUBLIC_DIR, "index.html"));
});
app.use(express.static(PUBLIC_DIR, { index: false }));

app.get("/api/health", async (_req, res) => {
  const job = await getCurrentJob();
  const accounts = await getAccounts();
  const settings = await loadSettings();
  res.json({
    ok: true,
    host: HOST,
    port: PORT,
    dataDir: DATA_DIR,
    accounts,
    job,
    settings,
    isProcessing,
  });
});

app.get("/api/settings", async (_req, res) => {
  res.json(await loadSettings());
});

app.post("/api/settings", async (req, res) => {
  try {
    const settings = await saveSettings(req.body || {});
    const job = await getCurrentJob();
    if (job && job.status !== "completed") {
      const nextJob = {
        ...job,
        perRunLimit: settings.maxContactsPerRun,
        chunkSize: settings.batchSize,
        delayBetweenRunsSec: settings.delayBetweenRunsSec,
        retryPauseSec: settings.retryPauseSec,
        retryRatioThreshold: settings.retryRatioThreshold,
        waitFloodAutomatically: settings.waitFloodAutomatically,
        autoStatus: settings.autoRun ? (job.autoStatus === "OFF" ? "RUNNING" : job.autoStatus) : "OFF",
        updatedAt: nowIso(),
      };
      await saveCurrentJob(nextJob);
    }
    res.json({ ok: true, settings });
  } catch (error) {
    res.status(400).json({ error: error.message || "บันทึกค่าไม่สำเร็จ" });
  }
});

app.get("/api/auto-status", async (_req, res) => {
  const settings = await loadSettings();
  const job = await markFloodStateStaleIfNeeded(await getCurrentJob());
  res.json({
    autoRun: settings.autoRun,
    status: job?.autoStatus || "OFF",
    queueStatus: displayQueueStatus(job?.status || "idle"),
    nextRunAt: job?.nextRunAt || "",
    lockedAccountLabel: job?.lockedAccountLabel || "",
    lockedAccountPhone: job?.lockedAccountPhone || "",
    file: job?.sourceFile || "",
    floodWaitSec: job?.lastFloodWaitSec || 0,
    floodWaitCount: Number(job?.floodWaitCount || 0),
    retryRatio: job?.lastRetryRatio || 0,
    processedRows: job?.processedRows || 0,
    remainingRows: job?.remainingRows || 0,
    requiresFloodReset: shouldRequireFloodReset(job),
  });
});

app.get("/api/dashboard", async (_req, res) => {
  await markFloodStateStaleIfNeeded(await getCurrentJob());
  res.json(await buildDashboardPayload());
});

app.get("/api/accounts", async (_req, res) => {
  try {
    const accounts = await getAccounts();
    res.json({ accounts });
  } catch (error) {
    res.status(500).json({ error: error.message || "โหลดบัญชีไม่สำเร็จ" });
  }
});

app.post("/api/accounts", async (req, res) => {
  try {
    const label = String(req.body.label || "").trim();
    const apiId = String(req.body.apiId || "").trim();
    const apiHash = String(req.body.apiHash || "").trim();
    const phone = normalizeThaiPhoneStrict(req.body.phone || "");

    if (!label) throw new Error("กรุณาใส่ชื่อบัญชี");
    if (!apiId) throw new Error("กรุณาใส่ API_ID");
    if (!apiHash) throw new Error("กรุณาใส่ API_HASH");
    if (!phone) throw new Error("กรุณาใส่ PHONE ให้ถูกต้อง");

    const { accounts, appState } = await loadAccountsState();
    const duplicate = accounts.find((item) => item.label === label || item.phone === phone);
    if (duplicate) throw new Error("มีบัญชีชื่อนี้หรือเบอร์นี้อยู่แล้ว");

    const account = {
      id: makeId("acc"),
      label,
      apiId,
      apiHashEnc: encryptText(apiHash),
      phone,
      sessionEnc: "",
      status: "new",
      pendingPhoneCodeHash: "",
      pendingCodeType: "",
      pendingTimeout: 0,
      awaitingPassword: false,
      lastError: "",
      lastUsedAt: "",
      me: null,
      createdAt: nowIso(),
      updatedAt: nowIso(),
    };

    accounts.push(account);
    if (!appState.selectedAccountId) appState.selectedAccountId = account.id;
    await saveAccountsState(accounts, appState);

    res.json({ ok: true, message: "เพิ่มบัญชีแล้ว", accounts: await getAccounts() });
  } catch (error) {
    res.status(400).json({ error: error.message || "เพิ่มบัญชีไม่สำเร็จ" });
  }
});

app.post("/api/accounts/:id/send-code", async (req, res) => {
  const accountId = String(req.params.id || "");
  try {
    const { account } = await getAccountById(accountId);
    if (!account) throw new Error("ไม่พบบัญชี");

    await disconnectClientMap(activeClients, accountId);
    await disconnectClientMap(pendingAuthClients, accountId);

    let apiHash = "";
    try {
      apiHash = decryptText(account.apiHashEnc || "").trim();
    } catch {
      throw makeAccountAuthError("API_HASH ของบัญชีนี้ถอดรหัสไม่ได้ กรุณาแก้บัญชีแล้วล็อกอินใหม่", "auth_invalid");
    }
    if (!apiHash) throw makeAccountAuthError("API_HASH ของบัญชีนี้ไม่ครบ กรุณาแก้บัญชีแล้วล็อกอินใหม่", "auth_invalid");
    const client = await createClient(account.apiId, apiHash, "");
    const sent = await client.invoke(new Api.auth.SendCode({
      phoneNumber: account.phone,
      apiId: Number(account.apiId),
      apiHash,
      settings: new Api.CodeSettings({}),
    }));

    pendingAuthClients.set(accountId, client);
    await updateAccountById(accountId, (current) => ({
      ...current,
      status: "code_sent",
      pendingPhoneCodeHash: sent.phoneCodeHash || "",
      pendingCodeType: sent.type?.className || sent.type?.constructor?.name || "sentCodeType",
      pendingTimeout: Number(sent.timeout || 0),
      awaitingPassword: false,
      lastError: "",
      updatedAt: nowIso(),
    }));

    res.json({
      ok: true,
      message: "ส่ง OTP แล้ว กรุณาไปกรอกรหัสในหน้าเว็บ",
      accounts: await getAccounts(),
    });
  } catch (error) {
    await disconnectClientMap(pendingAuthClients, accountId);
    const nextStatus = isAuthSessionError(error) ? "auth_invalid" : "error";
    try {
      await updateAccountById(accountId, (current) => ({
        ...current,
        status: nextStatus,
        lastError: extractTelegramError(error),
        updatedAt: nowIso(),
      }));
    } catch {}
    res.status(400).json({ error: extractTelegramError(error) });
  }
});

app.post("/api/accounts/:id/verify-code", async (req, res) => {
  const accountId = String(req.params.id || "");
  try {
    const code = String(req.body.code || "").trim();
    if (!code) throw new Error("กรุณาใส่รหัส OTP");
    const { account } = await getAccountById(accountId);
    if (!account) throw new Error("ไม่พบบัญชี");

    const client = pendingAuthClients.get(accountId);
    if (!client) throw new Error("รอบ OTP นี้หมดแล้ว กรุณากดส่ง OTP ใหม่");
    if (!account.pendingPhoneCodeHash) throw new Error("ไม่พบ phoneCodeHash กรุณากดส่ง OTP ใหม่");

    try {
      await client.invoke(new Api.auth.SignIn({
        phoneNumber: account.phone,
        phoneCodeHash: account.pendingPhoneCodeHash,
        phoneCode: code,
      }));
    } catch (error) {
      const raw = String(error?.message || "");
      if (raw.toUpperCase().includes("SESSION_PASSWORD_NEEDED")) {
        await updateAccountById(accountId, (current) => ({
          ...current,
          status: "awaiting_password",
          awaitingPassword: true,
          lastError: "",
          updatedAt: nowIso(),
        }));

        return res.json({
          ok: true,
          needPassword: true,
          message: "บัญชีนี้เปิด 2FA อยู่ กรุณาใส่รหัสผ่าน 2FA",
          accounts: await getAccounts(),
        });
      }
      throw error;
    }

    const me = await client.getMe();
    const sessionString = client.session.save();
    await writeSessionToFile(accountId, sessionString);
    activeClients.set(accountId, client);
    pendingAuthClients.delete(accountId);

    await updateAccountById(accountId, (current) => ({
      ...current,
      sessionEnc: encryptText(sessionString),
      status: "connected",
      pendingPhoneCodeHash: "",
      pendingCodeType: "",
      pendingTimeout: 0,
      awaitingPassword: false,
      lastError: "",
      me: me ? {
        id: toSafeStringId(me.id),
        username: me.username || "",
        firstName: me.firstName || "",
        lastName: me.lastName || "",
        phone: me.phone ? `+${onlyDigits(me.phone)}` : "",
      } : null,
      updatedAt: nowIso(),
    }));

    res.json({ ok: true, message: "เชื่อม Telegram สำเร็จแล้ว", accounts: await getAccounts() });
  } catch (error) {
    const nextStatus = isAuthSessionError(error) ? "needs_relogin" : (await (async () => {
      try {
        const { account } = await getAccountById(accountId);
        return account?.awaitingPassword ? "awaiting_password" : "code_sent";
      } catch {
        return "code_sent";
      }
    })());
    try {
      await updateAccountById(accountId, (current) => ({
        ...current,
        status: nextStatus,
        lastError: extractTelegramError(error),
        updatedAt: nowIso(),
      }));
    } catch {}
    res.status(400).json({ error: extractTelegramError(error) });
  }
});

app.post("/api/accounts/:id/verify-password", async (req, res) => {
  const accountId = String(req.params.id || "");
  try {
    const password = String(req.body.password || "").trim();
    if (!password) throw new Error("กรุณาใส่รหัส 2FA");
    const { account } = await getAccountById(accountId);
    if (!account) throw new Error("ไม่พบบัญชี");

    const client = pendingAuthClients.get(accountId);
    if (!client) throw new Error("รอบล็อกอินนี้หมดแล้ว กรุณากดส่ง OTP ใหม่");

    let apiHash = "";
    try {
      apiHash = decryptText(account.apiHashEnc || "").trim();
    } catch {
      throw makeAccountAuthError("API_HASH ของบัญชีนี้ถอดรหัสไม่ได้ กรุณาแก้บัญชีแล้วล็อกอินใหม่", "auth_invalid");
    }
    if (!apiHash) throw makeAccountAuthError("API_HASH ของบัญชีนี้ไม่ครบ กรุณาแก้บัญชีแล้วล็อกอินใหม่", "auth_invalid");
    await client.signInWithPassword(
      { apiId: Number(account.apiId), apiHash },
      {
        password: async () => password,
        onError: (err) => { throw err; },
      }
    );

    const me = await client.getMe();
    const sessionString = client.session.save();
    await writeSessionToFile(accountId, sessionString);
    activeClients.set(accountId, client);
    pendingAuthClients.delete(accountId);

    await updateAccountById(accountId, (current) => ({
      ...current,
      sessionEnc: encryptText(sessionString),
      status: "connected",
      pendingPhoneCodeHash: "",
      pendingCodeType: "",
      pendingTimeout: 0,
      awaitingPassword: false,
      lastError: "",
      me: me ? {
        id: toSafeStringId(me.id),
        username: me.username || "",
        firstName: me.firstName || "",
        lastName: me.lastName || "",
        phone: me.phone ? `+${onlyDigits(me.phone)}` : "",
      } : null,
      updatedAt: nowIso(),
    }));

    res.json({ ok: true, message: "ยืนยัน 2FA สำเร็จแล้ว", accounts: await getAccounts() });
  } catch (error) {
    const nextStatus = isAuthSessionError(error) ? "needs_relogin" : "awaiting_password";
    try {
      await updateAccountById(accountId, (current) => ({
        ...current,
        status: nextStatus,
        awaitingPassword: nextStatus === "awaiting_password",
        lastError: extractTelegramError(error),
        updatedAt: nowIso(),
      }));
    } catch {}
    res.status(400).json({ error: extractTelegramError(error) });
  }
});

app.post("/api/accounts/:id/select", async (req, res) => {
  const accountId = String(req.params.id || "");
  try {
    const { account, accounts, appState } = await getAccountById(accountId);
    if (!account) throw new Error("ไม่พบบัญชี");
    const hasSession = Boolean(account.sessionEnc) || fs.existsSync(sessionFilePath(account.id));
    if (!hasSession) throw new Error("บัญชีนี้ยังไม่ได้เชื่อม Telegram");
    appState.selectedAccountId = accountId;
    await saveAccountsState(accounts, appState);
    res.json({ ok: true, message: `เลือกบัญชี ${account.label} แล้ว`, accounts: await getAccounts() });
  } catch (error) {
    res.status(400).json({ error: error.message || "เลือกบัญชีไม่สำเร็จ" });
  }
});

app.post("/api/accounts/:id/reset-session", async (req, res) => {
  const accountId = String(req.params.id || "");
  try {
    const { account } = await getAccountById(accountId);
    if (!account) throw new Error("ไม่พบบัญชี");
    await resetAccountSession(accountId, "ล้าง session แล้ว กรุณาล็อกอินใหม่");
    res.json({ ok: true, message: `ล้าง session ของบัญชี ${account.label} แล้ว กรุณาส่ง OTP ใหม่`, accounts: await getAccounts() });
  } catch (error) {
    res.status(400).json({ error: error.message || "ล้าง session ไม่สำเร็จ" });
  }
});

app.delete("/api/accounts/:id", async (req, res) => {
  const accountId = String(req.params.id || "");
  try {
    const { accounts, appState } = await loadAccountsState();
    const index = accounts.findIndex((item) => item.id === accountId);
    if (index === -1) throw new Error("ไม่พบบัญชี");
    const [removed] = accounts.splice(index, 1);
    if (appState.selectedAccountId === accountId) {
      const next = accounts.find((item) => item.sessionEnc) || accounts[0] || null;
      appState.selectedAccountId = next?.id || "";
    }
    await saveAccountsState(accounts, appState);
    await disconnectClientMap(activeClients, accountId);
    await disconnectClientMap(pendingAuthClients, accountId);
    await unlinkIfExists(sessionFilePath(accountId));

    res.json({ ok: true, message: `ลบบัญชี ${removed.label} แล้ว`, accounts: await getAccounts() });
  } catch (error) {
    res.status(400).json({ error: error.message || "ลบบัญชีไม่สำเร็จ" });
  }
});

app.post("/api/intake/upload", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) throw new Error("กรุณาอัปโหลดไฟล์ก่อน");
    const { meta } = await processIntakeFile(req.file.path, req.file.originalname);
    res.json({ ok: true, message: "ล้างไฟล์เรียบร้อย", summary: meta, preview: meta.preview || [] });
  } catch (error) {
    res.status(400).json({ error: error.message || "ล้างไฟล์ไม่สำเร็จ" });
  }
});

app.post("/api/intake/import-from-path", async (req, res) => {
  try {
    const inputPath = String(req.body?.filePath || "").trim();
    if (!inputPath) throw new Error("กรุณาระบุ filePath");
    const resolved = path.resolve(inputPath);
    if (!fs.existsSync(resolved)) throw new Error("ไม่พบไฟล์ตาม path ที่ระบุ");
    const ext = path.extname(resolved).toLowerCase();
    if (!isSupportedIntakeExt(ext)) throw new Error("รองรับเฉพาะ csv, txt, xlsx, xls, xlsm, xlsb, ods");
    const { meta } = await processIntakeFile(resolved, path.basename(resolved));
    res.json({ ok: true, message: "นำเข้าไฟล์จาก path สำเร็จ", source: resolved, summary: meta, preview: meta.preview || [] });
  } catch (error) {
    res.status(400).json({ error: error.message || "นำเข้าไฟล์จาก path ไม่สำเร็จ" });
  }
});

app.post("/api/intake/split-csv", async (req, res) => {
  try {
    const inputPath = String(req.body?.filePath || "").trim();
    const rowsPerFile = toPositiveNumber(req.body?.rowsPerFile, 50000);
    if (!inputPath) throw new Error("กรุณาระบุ filePath");
    const resolved = path.resolve(inputPath);
    if (!fs.existsSync(resolved)) throw new Error("ไม่พบไฟล์ตาม path ที่ระบุ");
    if (path.extname(resolved).toLowerCase() !== ".csv") throw new Error("split อัตโนมัติรองรับเฉพาะไฟล์ csv");
    const result = await splitCsvFile(resolved, rowsPerFile);
    res.json({
      ok: true,
      message: "แยกไฟล์ csv สำเร็จ",
      source: resolved,
      outDir: result.outDir,
      totalRows: result.totalRows,
      rowsPerFile: result.rowsPerFile,
      files: result.files,
    });
  } catch (error) {
    res.status(400).json({ error: error.message || "แยกไฟล์ไม่สำเร็จ" });
  }
});

app.get("/api/intake/latest", async (_req, res) => {
  const latest = await readJson(FILES.intakeLatest, null);
  if (!latest) return res.status(404).json({ error: "ยังไม่มีผลล้างไฟล์ล่าสุด" });
  res.json(latest);
});

app.post("/api/jobs/create-from-clean", async (_req, res) => {
  try {
    const job = await createJobFromLatestClean();
    const settings = await loadSettings();
    res.json({ ok: true, message: "สร้างคิวจาก clean_ready.csv แล้ว", job, settings });
  } catch (error) {
    res.status(400).json({ error: error.message || "สร้างคิวไม่สำเร็จ" });
  }
});

app.get("/api/job-status", async (_req, res) => {
  const job = await markFloodStateStaleIfNeeded(await getCurrentJob());
  if (!job) return res.status(404).json({ error: "ยังไม่มีคิวงาน" });
  res.json(job);
});

app.get("/api/latest", async (_req, res) => {
  const latest = await readJson(FILES.latestJson, null);
  if (!latest) return res.status(404).json({ error: "ยังไม่มีผลล่าสุด" });
  res.json(latest);
});

app.get('/api/remaining', async (_req, res) => {
  const job = await getCurrentJob();
  if (!job || !Array.isArray(job.rows)) return res.status(404).json({ error: 'ยังไม่มีคิวงาน' });
  const doneSet = await getProcessedSourceIndexSet();
  const remainingRows = job.rows.filter((row) => !doneSet.has(String(row?.sourceIndex || '')));
  res.json({
    ok: true,
    file: 'remaining_only.csv',
    count: remainingRows.length,
    totalRows: job.totalRows || 0,
    processedRows: job.processedRows || 0,
    rows: remainingRows.slice(0, 30),
  });
});

app.post("/api/run-next", async (_req, res) => {
  if (isProcessing) return res.status(409).json({ error: "ระบบกำลังประมวลผลอยู่" });
  isProcessing = true;
  try {
    const staleChecked = await markFloodStateStaleIfNeeded(await getCurrentJob());
    if (staleChecked?.status === "paused_flood_stale") {
      return res.status(409).json({ error: "คิวถูกพักเพราะ FloodWait ค้างนาน กรุณารีเซ็ต flood state ก่อน" });
    }
    const result = await processNextBatch();
    res.json({ ok: true, message: "ทำรอบถัดไปเรียบร้อย", ...result });
  } catch (error) {
    const job = await getCurrentJob();
    if (job && (error?.isAccountAuthError || isAuthSessionError(error))) {
      if (job.autoStatus !== "OFF" || job.status !== "paused") {
        await saveCurrentJob({
          ...job,
          autoStatus: "OFF",
          status: "paused",
          updatedAt: nowIso(),
        });
      }
      await appendRunLog({ type: "account_auth_invalid", jobId: job.id, message: extractTelegramError(error) });
      return res.status(409).json({ error: extractTelegramError(error) || "session ใช้งานไม่ได้ กรุณาล็อกอินใหม่", code: "ACCOUNT_AUTH_INVALID" });
    }
    if (job && parseFloodSeconds(error?.message)) {
      const nextJob = await handleFloodWait(job, error);
      return res.status(429).json({ error: extractTelegramError(error), job: nextJob });
    }
    const friendly = extractTelegramError(error) || "ทำรอบถัดไปไม่สำเร็จ";
    const statusCode = /ยังไม่ได้เลือกบัญชี|ไม่พบบัญชี|ไม่ตรงกับบัญชี/.test(friendly) ? 409 : 400;
    res.status(statusCode).json({ error: friendly });
  } finally {
    isProcessing = false;
  }
});

app.post("/api/job/start-auto", async (_req, res) => {
  const job = await markFloodStateStaleIfNeeded(await getCurrentJob());
  if (!job) return res.status(404).json({ error: "ยังไม่มีคิวงาน" });
  if (job.status === "paused_flood_stale") {
    return res.status(409).json({ error: "คิวถูกพักเพราะ FloodWait ค้างนาน กรุณารีเซ็ต flood state ก่อน" });
  }
  const settings = await saveSettings({ autoRun: true });
  const now = Date.now();
  const waitFloodTs = job.nextRunAt ? new Date(job.nextRunAt).getTime() : 0;
  const waitRetryTs = job.nextRunAt ? new Date(job.nextRunAt).getTime() : 0;
  const keepFloodWaiting = job.status === "waiting_flood" && waitFloodTs && now < waitFloodTs;
  const keepRetryWaiting = job.status === "waiting_retry_cooldown" && waitRetryTs && now < waitRetryTs;
  const nextStatus = keepFloodWaiting
    ? "waiting_flood"
    : keepRetryWaiting
      ? "waiting_retry_cooldown"
      : (["paused", "ready", "waiting_retry_cooldown", "waiting_flood"].includes(job.status) ? "ready" : job.status);
  const nextAutoStatus = keepFloodWaiting ? "WAITING_FLOOD" : (keepRetryWaiting ? "WAITING_RETRY" : "RUNNING");

  const nextJob = {
    ...job,
    autoStatus: nextAutoStatus,
    status: nextStatus,
    nextRunAt: keepFloodWaiting || keepRetryWaiting ? job.nextRunAt : (job.nextRunAt || new Date().toISOString()),
    updatedAt: nowIso(),
  };
  await saveCurrentJob(nextJob);
  res.json({ ok: true, settings, job: nextJob });
});

app.post("/api/job/pause-auto", async (_req, res) => {
  const job = await getCurrentJob();
  if (!job) return res.status(404).json({ error: "ยังไม่มีคิวงาน" });
  const settings = await saveSettings({ autoRun: false });
  const nextJob = { ...job, autoStatus: "PAUSED", status: "paused", updatedAt: nowIso() };
  await saveCurrentJob(nextJob);
  res.json({ ok: true, settings, job: nextJob });
});

app.post("/api/job/resume-auto", async (_req, res) => {
  const job = await markFloodStateStaleIfNeeded(await getCurrentJob());
  if (!job) return res.status(404).json({ error: "ยังไม่มีคิวงาน" });
  if (job.status === "paused_flood_stale") {
    return res.status(409).json({ error: "คิวถูกพักเพราะ FloodWait ค้างนาน กรุณารีเซ็ต flood state ก่อน" });
  }
  const settings = await saveSettings({ autoRun: true });
  const now = Date.now();
  const waitFloodTs = job.nextRunAt ? new Date(job.nextRunAt).getTime() : 0;
  const waitRetryTs = job.nextRunAt ? new Date(job.nextRunAt).getTime() : 0;
  const keepFloodWaiting = job.status === "waiting_flood" && waitFloodTs && now < waitFloodTs;
  const keepRetryWaiting = job.status === "waiting_retry_cooldown" && waitRetryTs && now < waitRetryTs;
  const nextJob = {
    ...job,
    autoStatus: keepFloodWaiting ? "WAITING_FLOOD" : (keepRetryWaiting ? "WAITING_RETRY" : "RUNNING"),
    status: keepFloodWaiting ? "waiting_flood" : (keepRetryWaiting ? "waiting_retry_cooldown" : "ready"),
    nextRunAt: keepFloodWaiting || keepRetryWaiting ? job.nextRunAt : new Date().toISOString(),
    updatedAt: nowIso(),
  };
  await saveCurrentJob(nextJob);
  res.json({ ok: true, settings, job: nextJob });
});

async function resetFloodStateAndClearStaleJob() {
  const job = await getCurrentJob();
  if (!job) throw new Error("ยังไม่มีคิวงาน");
  if (!["waiting_flood", "paused_flood_stale", "paused"].includes(job.status)) {
    throw new Error("สถานะคิวนี้ไม่ใช่ flood state ที่รีเซ็ตได้");
  }

  const settings = await saveSettings({
    autoRun: false,
    maxContactsPerRun: 20,
    batchSize: 1,
  });
  const nextJob = {
    ...job,
    status: "ready",
    autoStatus: "OFF",
    perRunLimit: Number(settings.maxContactsPerRun || 20),
    chunkSize: Number(settings.batchSize || 1),
    nextRunAt: "",
    lastFloodWaitSec: 0,
    lastFloodWaitBaseSec: 0,
    consecutiveFloodCount: 0,
    updatedAt: nowIso(),
  };
  await saveCurrentJob(nextJob);
  await appendRunLog({
    type: "flood_reset",
    jobId: job.id,
    accountLabel: job.lockedAccountLabel || "",
    accountPhone: job.lockedAccountPhone || "",
  });
  return { job: nextJob, settings };
}

app.post("/api/job/reset-flood-state", async (_req, res) => {
  try {
    const result = await resetFloodStateAndClearStaleJob();
    res.json({ ok: true, ...result });
  } catch (error) {
    if (error.message === "ยังไม่มีคิวงาน") return res.status(404).json({ error: error.message });
    if (error.message === "สถานะคิวนี้ไม่ใช่ flood state ที่รีเซ็ตได้") return res.status(400).json({ error: error.message });
    res.status(500).json({ error: error.message || "รีเซ็ต flood state ไม่สำเร็จ" });
  }
});

app.post("/api/job/reset-flood-stale", async (_req, res) => {
  try {
    const result = await resetFloodStateAndClearStaleJob();
    res.json({ ok: true, ...result });
  } catch (error) {
    if (error.message === "ยังไม่มีคิวงาน") return res.status(404).json({ error: error.message });
    if (error.message === "สถานะคิวนี้ไม่ใช่ flood state ที่รีเซ็ตได้") return res.status(400).json({ error: error.message });
    res.status(500).json({ error: error.message || "รีเซ็ต flood state ไม่สำเร็จ" });
  }
});

app.post("/api/reset-job", async (_req, res) => {
  if (isProcessing) return res.status(409).json({ error: "กำลังประมวลผลอยู่ ล้างคิวตอนนี้ไม่ได้" });
  try {
    for (const filePath of [
      FILES.currentJob,
      FILES.jobState,
      FILES.latestJson,
      FILES.latestCsv,
      FILES.allJson,
      FILES.allCsv,
      FILES.retryCsv,
      FILES.autoState,
      FILES.remainingCsv,
    ]) await unlinkIfExists(filePath);
    res.json({ ok: true, message: "ล้างคิวและผลลัพธ์เรียบร้อย" });
  } catch (error) {
    res.status(500).json({ error: error.message || "ล้างคิวไม่สำเร็จ" });
  }
});

app.get("/download/:name", async (req, res) => {
  const allowed = new Map([
    ["telegram_matches.csv", FILES.latestCsv],
    ["telegram_matches.json", FILES.latestJson],
    ["telegram_matches_all.csv", FILES.allCsv],
    ["telegram_matches_all.json", FILES.allJson],
    ["retry_rows.csv", FILES.retryCsv],
    ["job_state.json", FILES.currentJob],
    ["clean_ready.csv", FILES.cleanReadyCsv],
    ["invalid_rows.csv", FILES.invalidRowsCsv],
    ["duplicate_phones.csv", FILES.duplicatePhonesCsv],
    ["clean_debug.csv", FILES.cleanDebugCsv],
    ["clean_rejects.csv", FILES.cleanRejectsCsv],
    ["summary.json", FILES.intakeSummaryJson],
    ["run_log.json", FILES.runLog],
    ["remaining_only.csv", FILES.remainingCsv],
  ]);

  const requested = String(req.params.name || "");
  if (!allowed.has(requested)) return res.status(404).send("ไม่พบไฟล์");
  const filePath = allowed.get(requested);
  if (!fs.existsSync(filePath)) return res.status(404).send("ยังไม่มีไฟล์นี้");
  res.download(filePath);
});

app.use((error, _req, res, _next) => {
  if (error instanceof multer.MulterError && error.code === "LIMIT_FILE_SIZE") {
    return res.status(413).json({
      error: `ไฟล์ใหญ่เกิน limit ปัจจุบัน ${MAX_UPLOAD_MB} MB`,
      code: "FILE_TOO_LARGE",
      currentLimitMb: MAX_UPLOAD_MB,
      alternatives: [
        "ใช้ /api/intake/import-from-path สำหรับนำเข้าไฟล์จาก disk path",
        "ใช้ /api/intake/split-csv เพื่อแยกไฟล์ csv เป็นหลายไฟล์ก่อนนำเข้า",
      ],
    });
  }
  const message = error?.message || "คำขอไม่ถูกต้อง";
  res.status(400).json({ error: message });
});

async function autoTick() {
  if (isProcessing) return;
  const settings = await loadSettings();
  const job = await markFloodStateStaleIfNeeded(await getCurrentJob());
  if (!job || !settings.autoRun) return;
  if (job.status === "completed" || job.status === "paused" || job.status === "paused_flood_stale") return;

  const now = Date.now();
  const nextRunTime = job.nextRunAt ? new Date(job.nextRunAt).getTime() : 0;
  if (job.status === "waiting_flood" || job.status === "waiting_retry_cooldown" || job.status === "ready") {
    if (nextRunTime && now < nextRunTime) return;
  }

  isProcessing = true;
  try {
    const nextJobBefore = await getCurrentJob();
    if (!nextJobBefore) return;
    const normalizedJob = {
      ...nextJobBefore,
      status: nextJobBefore.status === "waiting_flood" || nextJobBefore.status === "waiting_retry_cooldown" ? "ready" : nextJobBefore.status,
      autoStatus: "RUNNING",
      updatedAt: nowIso(),
    };
    await saveCurrentJob(normalizedJob);
    await processNextBatch();
  } catch (error) {
    const current = await getCurrentJob();
    if (current && (error?.isAccountAuthError || isAuthSessionError(error))) {
      await saveCurrentJob({
        ...current,
        autoStatus: "OFF",
        status: "paused",
        updatedAt: nowIso(),
      });
      pushRecoveryNotice('warning', 'AUTO_AUTH_INVALID', 'ต้องล็อกอินบัญชีใหม่', extractTelegramError(error) || "session ใช้งานไม่ได้ กรุณาล็อกอินใหม่");
      await appendRunLog({ type: "auto_auth_invalid", jobId: current.id, message: extractTelegramError(error) });
    } else if (current && parseFloodSeconds(error?.message)) {
      await handleFloodWait(current, error);
    } else {
      if (current) {
        await saveCurrentJob({
          ...current,
          autoStatus: "PAUSED",
          status: "paused",
          updatedAt: nowIso(),
        });
      }
      pushRecoveryNotice('danger', 'AUTO_ERROR', 'รอบออโต้มีปัญหา', extractTelegramError(error));
      await appendRunLog({ type: "auto_error", message: extractTelegramError(error) });
    }
  } finally {
    isProcessing = false;
  }
}

(async () => {
  validateBootConfig();
  await ensureDirectories();

  for (const filePath of CRITICAL_BACKUP_FILES) {
    await seedLastGoodBackup(filePath);
  }

  if (!fs.existsSync(FILES.accounts)) await writeJson(FILES.accounts, []);
  if (!fs.existsSync(FILES.appState)) await writeJson(FILES.appState, { selectedAccountId: "" });
  if (!fs.existsSync(FILES.settings)) await writeJson(FILES.settings, DEFAULT_SETTINGS);
  if (!fs.existsSync(FILES.runLog)) await writeJson(FILES.runLog, []);
  if (!fs.existsSync(FILES.autoState)) await writeJson(FILES.autoState, await buildAutoState(null));

  autoTimer = setInterval(() => {
    autoTick().catch((err) => console.error("[AUTO]", err));
  }, 5000);

  await markFloodStateStaleIfNeeded(await getCurrentJob());

  app.listen(PORT, HOST, () => {
    logInfo(`Telegram All-in-One running on http://${HOST}:${PORT}`);
    logInfo(`DATA_DIR=${DATA_DIR}`);
  });
})();
