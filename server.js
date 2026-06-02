const dotenv = require("dotenv");
dotenv.config({ override: true });

const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");
const crypto = require("crypto");
const { AsyncLocalStorage } = require("async_hooks");
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
const CHUNK_FILE_SIZE = Math.max(1000, toPositiveNumber(process.env.CHUNK_FILE_SIZE, 1000));
const HIGH_RETRY_MANUAL_PAUSE_RATIO = 0.8;
const STATE_RENAME_RETRY_DELAYS_MS = [300, 800, 1500];
const SIMULATE_STATE_RENAME_EPERM = String(process.env.SIMULATE_STATE_RENAME_EPERM || "").trim();
const SIMULATE_STATE_RENAME_EPERM_COUNT = Math.max(0, Math.floor(Number(process.env.SIMULATE_STATE_RENAME_EPERM_COUNT || 0)));

const APP_ROOT = __dirname;
const DATA_DIR = path.resolve(process.env.DATA_DIR || path.join(APP_ROOT, "data"));
const SYSTEM_IDS = ["A", "B", "C", "D", "E"];
const DEFAULT_SYSTEM_ID = "A";
const SYSTEMS_DIR = path.join(DATA_DIR, "systems");
const LEGACY_DIRS = {
  raw: path.join(DATA_DIR, "input_raw"),
  clean: path.join(DATA_DIR, "input_clean"),
  output: path.join(DATA_DIR, "output"),
  uploads: path.join(DATA_DIR, "uploads"),
  chunks: path.join(DATA_DIR, "chunks"),
  jobs: path.join(DATA_DIR, "jobs"),
  session: path.join(DATA_DIR, "session"),
  logs: path.join(DATA_DIR, "logs"),
};
const RECOVERY_DIR = path.join(DATA_DIR, "recovery");

const systemContext = new AsyncLocalStorage();
const systemPathCache = new Map();
const CRITICAL_BACKUP_FILE_NAMES = new Set([
  "current_job.json",
  "job_state.json",
  "auto_state.json",
  "settings.json",
  "app_state.json",
  "accounts.json",
  "run_log.json",
]);
const recoveryNoticesBySystem = new Map();

function normalizeSystemId(value) {
  const systemId = String(value || DEFAULT_SYSTEM_ID).trim().toUpperCase();
  return SYSTEM_IDS.includes(systemId) ? systemId : "";
}

function currentSystemId() {
  return normalizeSystemId(systemContext.getStore()?.systemId) || DEFAULT_SYSTEM_ID;
}

function withSystem(systemId, fn) {
  const normalized = normalizeSystemId(systemId);
  if (!normalized) throw new Error("systemId ไม่ถูกต้อง ต้องเป็น A, B, C, D หรือ E");
  return systemContext.run({ systemId: normalized }, fn);
}

function getSystemPaths(systemId = currentSystemId()) {
  const id = normalizeSystemId(systemId);
  if (!id) throw new Error("systemId ไม่ถูกต้อง ต้องเป็น A, B, C, D หรือ E");
  if (systemPathCache.has(id)) return systemPathCache.get(id);

  const root = path.join(SYSTEMS_DIR, id);
  const inputDir = path.join(root, "input");
  const cleanDir = path.join(inputDir, "clean");
  const rawDir = path.join(inputDir, "raw");
  const uploadsDir = path.join(inputDir, "uploads");
  const outputDir = path.join(root, "output");
  const jobsDir = path.join(root, "jobs");
  const chunksDir = path.join(jobsDir, "chunks");
  const sessionDir = path.join(root, "session");
  const logDir = path.join(root, "logs");
  const files = {
    accounts: path.join(root, "accounts.json"),
    appState: path.join(root, "app_state.json"),
    settings: path.join(root, "settings.json"),
    intakeLatest: path.join(cleanDir, "intake_latest.json"),
    cleanReadyCsv: path.join(cleanDir, "clean_ready.csv"),
    invalidRowsCsv: path.join(cleanDir, "invalid_rows.csv"),
    duplicatePhonesCsv: path.join(cleanDir, "duplicate_phones.csv"),
    cleanDebugCsv: path.join(cleanDir, "clean_debug.csv"),
    cleanRejectsCsv: path.join(cleanDir, "clean_rejects.csv"),
    intakeSummaryJson: path.join(cleanDir, "summary.json"),
    currentJob: path.join(outputDir, "current_job.json"),
    jobState: path.join(outputDir, "job_state.json"),
    latestJson: path.join(outputDir, "telegram_matches.json"),
    latestCsv: path.join(outputDir, "telegram_matches.csv"),
    allJson: path.join(outputDir, "telegram_matches_all.json"),
    allCsv: path.join(outputDir, "telegram_matches_all.csv"),
    processedCsv: path.join(outputDir, "processed_only.csv"),
    processedXlsx: path.join(outputDir, "processed_only.xlsx"),
    retryCsv: path.join(outputDir, "retry_rows.csv"),
    yesOnlyCsv: path.join(outputDir, "telegram_yes_only.csv"),
    yesOnlyJson: path.join(outputDir, "telegram_yes_only.json"),
    noOnlyCsv: path.join(outputDir, "telegram_no_only.csv"),
    noOnlyJson: path.join(outputDir, "telegram_no_only.json"),
    retryOnlyCsv: path.join(outputDir, "telegram_retry_only.csv"),
    retryOnlyJson: path.join(outputDir, "telegram_retry_only.json"),
    invalidOnlyCsv: path.join(outputDir, "telegram_invalid_only.csv"),
    invalidOnlyJson: path.join(outputDir, "telegram_invalid_only.json"),
    marketingAllowedCsv: path.join(outputDir, "marketing_allowed_only.csv"),
    marketingAllowedJson: path.join(outputDir, "marketing_allowed_only.json"),
    winbackCsv: path.join(outputDir, "winback_only.csv"),
    winbackJson: path.join(outputDir, "winback_only.json"),
    retryLaterCsv: path.join(outputDir, "retry_later_only.csv"),
    retryLaterJson: path.join(outputDir, "retry_later_only.json"),
    autoState: path.join(outputDir, "auto_state.json"),
    runLog: path.join(logDir, "run_log.json"),
    importDiagnosticLog: path.join(logDir, "import_diagnostic_log.jsonl"),
    remainingCsv: path.join(outputDir, "remaining_only.csv"),
    remainingXlsx: path.join(outputDir, "remaining_only.xlsx"),
  };
  const paths = { id, root, inputDir, cleanDir, rawDir, uploadsDir, outputDir, jobsDir, chunksDir, sessionDir, logDir, files };
  systemPathCache.set(id, paths);
  return paths;
}

const FILES = new Proxy({}, {
  get(_target, prop) {
    return getSystemPaths().files[prop];
  },
});

function lastGoodPath(filePath) {
  return `${filePath}.lastgood`;
}

function shouldKeepLastGood(filePath) {
  return CRITICAL_BACKUP_FILE_NAMES.has(path.basename(filePath));
}

function getRecoveryNotices(systemId = currentSystemId()) {
  const id = normalizeSystemId(systemId) || DEFAULT_SYSTEM_ID;
  if (!recoveryNoticesBySystem.has(id)) recoveryNoticesBySystem.set(id, []);
  return recoveryNoticesBySystem.get(id);
}

function pushRecoveryNotice(level, code, title, message, extra = {}) {
  const recoveryNotices = getRecoveryNotices();
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

const processingBySystem = new Map();
const activeClients = new Map();
const pendingAuthClients = new Map();
const autoTimers = new Map();
const simulatedRenameFailures = new Map();
let contactClientIdCounter = 0n;

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
  if (value === 'waiting_retry_manual') return 'WAITING_RETRY_MANUAL';
  if (value === 'paused_too_many_retry') return 'PAUSED_TOO_MANY_RETRY';
  if (value === 'waiting_retry_cooldown' || value === 'paused_retry_cooldown') return 'WAITING_RETRY';
  if (value === 'waiting_flood') return 'WAITING_FLOOD';
  if (value === 'auth_required') return 'AUTH_REQUIRED';
  if (value === 'paused_flood_stale') return 'PAUSED_FLOOD_STALE';
  if (value === 'completed') return 'COMPLETED';
  if (value === 'paused') return 'PAUSED';
  if (value === 'ready') return 'READY';
  if (value === 'running') return 'RUNNING';
  if (value === 'auth_required') return 'AUTH_REQUIRED';
  return value ? value.toUpperCase() : 'IDLE';
}

function displayAutoStatus(status) {
  const value = String(status || '').toLowerCase();
  if (value === 'off') return 'OFF';
  if (value === 'running') return 'RUNNING';
  if (value === 'auth_required') return 'AUTH_REQUIRED';
  if (value === 'waiting_flood') return 'WAITING_FLOOD';
  if (value === 'waiting_retry_cooldown' || value === 'paused_retry_cooldown' || value === 'waiting_retry') return 'WAITING_RETRY';
  if (value === 'paused_flood_stale') return 'PAUSED_FLOOD_STALE';
  if (value === 'paused') return 'PAUSED';
  return value ? value.toUpperCase() : 'OFF';
}

function isHighRetryManualPause(job) {
  return Boolean(
    job &&
    String(job.status || "").toLowerCase() !== "completed" &&
    Number(job.lastRetryRatio || 0) >= HIGH_RETRY_MANUAL_PAUSE_RATIO
  );
}

function shouldRequireFloodReset(job) {
  return String(job?.status || "").toLowerCase() === "paused_flood_stale";
}

function logInfo(message) {
  console.log(`[INFO] ${message}`);
}

function safePathForResponse(filePath) {
  const resolved = path.resolve(String(filePath || ""));
  const relative = path.relative(APP_ROOT, resolved);
  if (!relative.startsWith("..") && !path.isAbsolute(relative)) return relative.replace(/\\/g, "/");
  return path.basename(resolved);
}

function onlyDigits(value) {
  return String(value || "").replace(/\D+/g, "");
}

function sanitizeFileName(name) {
  return String(name || "upload").replace(/[^a-zA-Z0-9._-]+/g, "_");
}

function clientMapKey(accountId, systemId = currentSystemId()) {
  return `${normalizeSystemId(systemId) || DEFAULT_SYSTEM_ID}:${String(accountId || "")}`;
}

function isSystemProcessing(systemId = currentSystemId()) {
  return Boolean(processingBySystem.get(normalizeSystemId(systemId) || DEFAULT_SYSTEM_ID));
}

function setSystemProcessing(value, systemId = currentSystemId()) {
  processingBySystem.set(normalizeSystemId(systemId) || DEFAULT_SYSTEM_ID, Boolean(value));
}

function decodeUploadFileName(name) {
  const raw = String(name || "").trim();
  if (!raw) return "upload";
  try {
    const decoded = Buffer.from(raw, "latin1").toString("utf8").trim();
    if (!decoded) return raw;
    const replacementRatio = (decoded.match(/\uFFFD/g) || []).length / Math.max(1, decoded.length);
    return replacementRatio > 0.2 ? raw : decoded;
  } catch {
    return raw;
  }
}

function toAsciiDownloadName(name, fallback = "download") {
  const safe = sanitizeFileName(String(name || fallback));
  return safe || `${fallback}.dat`;
}

function encodeRFC5987ValueChars(str) {
  return encodeURIComponent(String(str || "download"))
    .replace(/['()]/g, escape)
    .replace(/\*/g, "%2A");
}

function sendFileDownload(res, filePath, displayName) {
  const asciiName = toAsciiDownloadName(displayName, path.basename(filePath));
  const utf8Name = encodeRFC5987ValueChars(displayName || path.basename(filePath));
  res.setHeader("Content-Disposition", `attachment; filename="${asciiName}"; filename*=UTF-8''${utf8Name}`);
  res.sendFile(path.resolve(filePath));
}

function fileStem(name) {
  const base = String(name || "").trim();
  if (!base) return "";
  const ext = path.extname(base);
  return ext ? base.slice(0, -ext.length) : base;
}

function normalizeJobLabel(inputLabel, fallbackFileName = "") {
  const raw = String(inputLabel || "").trim();
  if (raw) return raw.slice(0, 60);
  const fallback = fileStem(fallbackFileName) || String(fallbackFileName || "").trim() || "งานไม่ระบุชื่อ";
  return fallback.slice(0, 60);
}

function normalizeJobNote(note) {
  return String(note || "").trim().slice(0, 200);
}

function parseSourceLabelBonusConfig() {
  try {
    const raw = String(process.env.SOURCE_LABEL_BONUS_JSON || "").trim();
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") return {};
    const out = {};
    for (const [key, value] of Object.entries(parsed)) {
      const safeKey = String(key || "").trim().toLowerCase();
      const safeValue = Number(value);
      if (!safeKey || !Number.isFinite(safeValue)) continue;
      out[safeKey] = Math.round(safeValue);
    }
    return out;
  } catch {
    return {};
  }
}

const SOURCE_LABEL_BONUS_CONFIG = parseSourceLabelBonusConfig();

function normalizeSourceLabel(value, fallback = "unknown") {
  const text = String(value || "").trim();
  return text || String(fallback || "unknown").trim() || "unknown";
}

function normalizeConsentStatus(value) {
  const text = String(value || "").trim().toLowerCase();
  if (["yes", "y", "true", "1", "consented", "optin", "opt_in"].includes(text)) return "yes";
  if (["no", "n", "false", "0", "optout", "opt_out", "denied"].includes(text)) return "no";
  return "unknown";
}

function normalizeCustomerStatus(value) {
  const text = String(value || "").trim().toLowerCase();
  if (["buyer", "customer", "paid", "purchase"].includes(text)) return "buyer";
  if (["old", "existing", "returning"].includes(text)) return "old";
  return "new";
}

function computeSourceLabelBonus(sourceLabel, bonusOverrides = SOURCE_LABEL_BONUS_CONFIG) {
  const key = String(sourceLabel || "").trim().toLowerCase();
  if (!key) return 0;
  const value = Number(bonusOverrides?.[key] || 0);
  return Number.isFinite(value) ? Math.round(value) : 0;
}

function computeLeadScore(row, bonusOverrides = SOURCE_LABEL_BONUS_CONFIG) {
  let score = 0;
  const status = String(row?.status || "").toUpperCase();
  if (status === "YES") score += 20;
  else if (status === "NO") score -= 20;
  else if (status === "RETRY") score -= 10;

  const consentStatus = normalizeConsentStatus(row?.consentStatus);
  if (consentStatus === "yes") score += 50;

  const customerStatus = normalizeCustomerStatus(row?.customerStatus);
  if (customerStatus === "old") score += 30;
  else if (customerStatus === "buyer") score += 40;

  score += computeSourceLabelBonus(row?.sourceLabel, bonusOverrides);
  return score;
}

function computeNextAction(row) {
  const status = String(row?.status || "").toUpperCase();
  const consentStatus = normalizeConsentStatus(row?.consentStatus);
  const customerStatus = normalizeCustomerStatus(row?.customerStatus);
  const score = Number(row?.leadScore || 0);

  if (status === "INVALID") return "FIX_PHONE";
  if (status === "RETRY") return "RETRY_LATER";
  if (consentStatus !== "yes") return "HOLD_FOR_CONSENT";
  if (score <= 0) return "LOW_QUALITY";
  if ((customerStatus === "old" || customerStatus === "buyer") && status !== "RETRY") return "WINBACK";
  if (status === "YES" && score >= 70) return "MARKETING_ALLOWED";
  if (status === "NO") return "RETARGETING_ADS";
  return "RETARGETING_ADS";
}

function withSegmentation(row, fallbackSourceLabel = "unknown", bonusOverrides = SOURCE_LABEL_BONUS_CONFIG) {
  const sourceLabel = normalizeSourceLabel(row?.sourceLabel, fallbackSourceLabel);
  const consentStatus = normalizeConsentStatus(row?.consentStatus);
  const customerStatus = normalizeCustomerStatus(row?.customerStatus);
  const leadScore = computeLeadScore({ ...row, sourceLabel, consentStatus, customerStatus }, bonusOverrides);
  const nextAction = computeNextAction({ ...row, sourceLabel, consentStatus, customerStatus, leadScore });
  return {
    ...row,
    sourceLabel,
    consentStatus,
    customerStatus,
    leadScore,
    nextAction,
  };
}

function buildAliasPool(group = "A", size = 8) {
  const g = String(group || "A").trim().toUpperCase().replace(/[^A-Z0-9]/g, "") || "A";
  const out = [];
  for (let i = 1; i <= Math.max(1, Number(size || 8)); i += 1) out.push(`${g}${i}`);
  return out;
}

function parseAliasGroupFromNote(jobNote = "") {
  const text = String(jobNote || "");
  const match = text.match(/(?:aliasSet|aliasGroup|pool)\s*[:=]\s*([A-Za-z0-9]+)/i);
  return match ? String(match[1] || "").trim().toUpperCase() : "";
}

function deriveAliasGroup(jobLabel = "", jobNote = "") {
  const fromNote = parseAliasGroupFromNote(jobNote);
  if (fromNote) return fromNote;
  const label = String(jobLabel || "").trim().toUpperCase();
  if (label.startsWith("B")) return "B";
  if (label.startsWith("C")) return "C";
  return "A";
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

function maskPhone(value) {
  const text = String(value || "");
  const digits = onlyDigits(text);
  if (!digits) return "";
  return maskSecret(digits, 3, 2);
}

function sanitizeDiagnosticErrorMessage(error) {
  const text = normalizeErrorMessage(error).replace(/\s+/g, " ").trim();
  if (!text) return "";
  return text
    .replace(/\b\d{8,15}\b/g, (value) => maskPhone(value) || "***")
    .slice(0, 500);
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

function nextContactClientId() {
  contactClientIdCounter = (contactClientIdCounter + 1n) % 999999n;
  return BigInt(Date.now()) * 1000000n + contactClientIdCounter;
}

function validateCheckerPhone(row) {
  const raw = row?.normalizedPhone || row?.phone || row?.rawPhone || "";
  const normalized = normalizeThaiPhoneStrict(raw);
  if (!normalized) return { ok: false, normalizedPhone: "", reason: "invalid_phone" };
  return { ok: true, normalizedPhone: normalized, reason: "" };
}

function makeCheckerResultRow(row, status, extra = {}) {
  return withSegmentation({
    sourceIndex: row.sourceIndex,
    name: row.name,
    rawPhone: row.rawPhone,
    normalizedPhone: extra.normalizedPhone ?? row.normalizedPhone ?? row.phone ?? "",
    status,
    hasTelegram: status === "YES",
    retry: status === "RETRY",
    telegramUserId: extra.telegramUserId || "",
    telegramUsername: extra.telegramUsername || "",
    telegramFirstName: extra.telegramFirstName || "",
    telegramLastName: extra.telegramLastName || "",
    telegramPhone: extra.telegramPhone || "",
    clientId: extra.clientId || "",
    sourceLabel: row.sourceLabel || "unknown",
    consentStatus: row.consentStatus || "unknown",
    customerStatus: row.customerStatus || "new",
    reason: extra.reason || "",
  }, row.sourceLabel || "unknown");
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

function makeAuthRequiredError(message = "บัญชี Telegram ยังไม่พร้อม กรุณารีล็อกอิน", accountStatus = "needs_relogin") {
  const err = makeAccountAuthError(message, accountStatus);
  err.type = "AUTH_REQUIRED";
  err.code = "AUTH_REQUIRED";
  return err;
}

function authRequiredMessage(error) {
  const raw = normalizeErrorMessage(error);
  if (/UNSUPPORTED STATE OR UNABLE TO AUTHENTICATE DATA/i.test(raw)) {
    return "Session ใช้ไม่ได้ ต้องล็อกอินบัญชี Telegram ใหม่";
  }
  return "บัญชี Telegram ยังไม่พร้อม กรุณารีล็อกอิน";
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
  return path.join(getSystemPaths().sessionDir, `${safeId}.session`);
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
  const dirs = [DATA_DIR, SYSTEMS_DIR, RECOVERY_DIR];
  for (const systemId of SYSTEM_IDS) {
    const p = getSystemPaths(systemId);
    dirs.push(p.root, p.inputDir, p.rawDir, p.uploadsDir, p.cleanDir, p.outputDir, p.jobsDir, p.chunksDir, p.sessionDir, p.logDir);
  }
  await Promise.all(dirs.map((dir) => fsp.mkdir(dir, { recursive: true })));
}

function makeTempFilePath(filePath) {
  return path.join(
    path.dirname(filePath),
    `.${path.basename(filePath)}.${process.pid}.${Date.now()}.${crypto.randomBytes(8).toString("hex")}.tmp`
  );
}

function isRetriableReplaceError(error) {
  return ["EPERM", "EACCES", "EBUSY"].includes(String(error?.code || ""));
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function backupJsonPath(filePath) {
  const ext = path.extname(filePath);
  if (ext.toLowerCase() === ".json") return filePath.slice(0, -ext.length) + ".backup.json";
  return `${filePath}.backup.json`;
}

function maybeSimulateStateRenameError(filePath) {
  if (!SIMULATE_STATE_RENAME_EPERM) return;
  const target = path.basename(SIMULATE_STATE_RENAME_EPERM);
  if (target && path.basename(filePath) !== target) return;
  const key = path.resolve(filePath);
  const used = simulatedRenameFailures.get(key) || 0;
  if (used >= SIMULATE_STATE_RENAME_EPERM_COUNT) return;
  simulatedRenameFailures.set(key, used + 1);
  const error = new Error(`simulated EPERM rename for ${path.basename(filePath)}`);
  error.code = "EPERM";
  throw error;
}

async function replaceFileWithRetry(tmpFile, filePath) {
  let lastError = null;

  for (let attempt = 0; attempt <= STATE_RENAME_RETRY_DELAYS_MS.length; attempt += 1) {
    try {
      maybeSimulateStateRenameError(filePath);
      await fsp.rename(tmpFile, filePath);
      return;
    } catch (error) {
      lastError = error;
      if (!isRetriableReplaceError(error) || attempt === STATE_RENAME_RETRY_DELAYS_MS.length) break;
      await sleep(STATE_RENAME_RETRY_DELAYS_MS[attempt]);
    }
  }

  if (lastError) {
    lastError.message = `${lastError.message} (rename failed after retries ${STATE_RENAME_RETRY_DELAYS_MS.join(",")}ms)`;
  }
  throw lastError || new Error(`rename failed for ${filePath}`);
}


async function writeTextAtomic(filePath, text) {
  await fsp.mkdir(path.dirname(filePath), { recursive: true });
  const tmpFile = makeTempFilePath(filePath);
  await fsp.writeFile(tmpFile, text, "utf8");
  try {
    await replaceFileWithRetry(tmpFile, filePath);
  } catch (error) {
    await unlinkIfExists(tmpFile).catch(() => {});
    throw error;
  }
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
  try {
    await writeJson(FILES.runLog, current.slice(-1000));
  } catch (error) {
    warnStateWriteFailure(FILES.runLog, error);
  }
}

async function appendImportDiagnosticLog(entry) {
  const record = { at: nowIso(), ...entry };
  try {
    await fsp.mkdir(path.dirname(FILES.importDiagnosticLog), { recursive: true });
    await fsp.appendFile(FILES.importDiagnosticLog, `${JSON.stringify(record)}\n`, "utf8");
  } catch (error) {
    warnStateWriteFailure(FILES.importDiagnosticLog, error);
  }
  return record;
}

async function readImportDiagnosticLog(limit = 50) {
  if (!fs.existsSync(FILES.importDiagnosticLog)) return [];
  const raw = await fsp.readFile(FILES.importDiagnosticLog, "utf8").catch(() => "");
  const rows = raw
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => {
      try {
        return JSON.parse(line);
      } catch {
        return null;
      }
    })
    .filter(Boolean);
  if (limit === "all") return rows;
  const count = Math.max(1, Math.min(500, Math.floor(Number(limit || 50))));
  return rows.slice(-count);
}

function buildImportDiagnosticSummary(records = []) {
  const list = Array.isArray(records) ? records.filter(Boolean) : [];
  const latest = list[list.length - 1] || null;
  const ratioRows = list.filter((entry) => Number.isFinite(Number(entry.retryRatio))).slice(-20);
  const avgRetryRatio = ratioRows.length
    ? ratioRows.reduce((sum, entry) => sum + Number(entry.retryRatio || 0), 0) / ratioRows.length
    : 0;
  const maxRetryRatio = ratioRows.length
    ? Math.max(...ratioRows.map((entry) => Number(entry.retryRatio || 0)))
    : 0;
  return {
    lastAt: latest?.at || "",
    lastJobId: latest?.jobId || "",
    lastAccountLabel: latest?.accountLabel || "",
    totalCalls: list.length,
    avgRetryRatio,
    maxRetryRatio,
    lastImportedCount: Number(latest?.importedCount || 0),
    lastUsersCount: Number(latest?.usersCount || 0),
    lastRetryContactsCount: Number(latest?.retryContactsCount || 0),
    lastContactsInCall: Number(latest?.contactsInCall || 0),
    lastPerRunLimit: Number(latest?.perRunLimit || 0),
    lastChunkSize: Number(latest?.chunkSize || 0),
  };
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
  const key = clientMapKey(accountId);
  const client = map.get(key);
  map.delete(key);
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
  if (!accountId) throw makeAuthRequiredError("บัญชี Telegram ยังไม่พร้อม กรุณารีล็อกอิน", "auth_invalid");
  if (!String(account?.apiId || "").trim()) throw makeAuthRequiredError("บัญชี Telegram ยังไม่พร้อม กรุณารีล็อกอิน", "auth_invalid");

  let apiHash = "";
  try {
    apiHash = decryptText(account.apiHashEnc || "").trim();
  } catch {
    await markAccountForRelogin(accountId, "API_HASH ของบัญชีนี้ถอดรหัสไม่ได้ กรุณารีล็อกอินบัญชีใหม่", "auth_invalid");
    throw makeAuthRequiredError("บัญชี Telegram ยังไม่พร้อม กรุณารีล็อกอิน", "auth_invalid");
  }
  if (!apiHash) {
    await markAccountForRelogin(accountId, "API_HASH ของบัญชีนี้ไม่ครบ กรุณารีล็อกอินบัญชีใหม่", "auth_invalid");
    throw makeAuthRequiredError("บัญชี Telegram ยังไม่พร้อม กรุณารีล็อกอิน", "auth_invalid");
  }

  const { sessionString } = await loadAccountSessionString(account);
  if (!sessionString) {
    await markAccountForRelogin(accountId, "บัญชีนี้ยังไม่มี session ที่ใช้งานได้ กรุณาล็อกอินใหม่", "needs_relogin");
    throw makeAuthRequiredError("บัญชี Telegram ยังไม่พร้อม กรุณารีล็อกอิน", "needs_relogin");
  }

  const cached = activeClients.get(clientMapKey(account.id));
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
      const message = authRequiredMessage(error);
      await markAccountForRelogin(accountId, message, /UNSUPPORTED STATE OR UNABLE TO AUTHENTICATE DATA/i.test(normalizeErrorMessage(error)) ? "auth_invalid" : "needs_relogin");
      throw makeAuthRequiredError(message, /UNSUPPORTED STATE OR UNABLE TO AUTHENTICATE DATA/i.test(normalizeErrorMessage(error)) ? "auth_invalid" : "needs_relogin");
    }
    throw error;
  }
  if (!authorized) {
    try { await client.disconnect(); } catch {}
    await markAccountForRelogin(accountId, "session ของบัญชีนี้หมดอายุแล้ว กรุณาล็อกอินใหม่", "needs_relogin");
    throw makeAuthRequiredError("บัญชี Telegram ยังไม่พร้อม กรุณารีล็อกอิน", "needs_relogin");
  }

  await updateAccountById(accountId, (current) => ({
    ...current,
    status: "connected",
    lastError: "",
    updatedAt: nowIso(),
  })).catch(() => {});
  activeClients.set(clientMapKey(account.id), client);
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
    const primaryPhone = phoneColumn ? String(row[phoneColumn] || "").trim() : "";
    const fallbackRawPhone = String(row.rawPhone || row.raw_phone || row.rawphone || "").trim();
    const rawPhone = primaryPhone || fallbackRawPhone;
    const sourceLabelRaw = String(row.sourceLabel || row.source_label || row.source || row.list || row.campaign || "").trim();
    const consentRaw = String(row.consentStatus || row.consent_status || row.consent || "").trim();
    const customerRaw = String(row.customerStatus || row.customer_status || row.customer || row.customer_type || "").trim();
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
      sourceLabel: normalizeSourceLabel(sourceLabelRaw, "unknown"),
      consentStatus: normalizeConsentStatus(consentRaw),
      customerStatus: normalizeCustomerStatus(customerRaw),
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
    sourceLabel: row.sourceLabel || "",
    consentStatus: row.consentStatus || "",
    customerStatus: row.customerStatus || "",
  })), {
    header: true,
    columns: ["name", "phone", "rawPhone", "rowNumber", "sourceLabel", "consentStatus", "customerStatus"],
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

async function saveIntakeOutputs(filename, result, options = {}) {
  const jobLabel = normalizeJobLabel(options.jobLabel, filename);
  const jobNote = normalizeJobNote(options.jobNote);
  const cleanReadyPath = FILES.cleanReadyCsv;
  const meta = {
    fileName: filename,
    originalFilename: filename,
    masterFile: String(options.masterFile || ""),
    jobLabel,
    jobNote,
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
  await fsp.writeFile(cleanReadyPath, cleanRowsToCsv(result.cleanRows), "utf8");
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
    fileName: meta.fileName || "",
    jobLabel: meta.jobLabel || "",
    totalRows: meta.totalRows || 0,
    readyRows: meta.readyRows || 0,
    invalidRows: meta.invalidRows || 0,
    duplicateRows: meta.duplicateRows || 0,
    detectedPhoneColumn: meta.detectedPhoneColumn || "",
  });
  logInfo(
    `Cleaner summary label=${meta.jobLabel || "-"} total=${meta.totalRows || 0} ready=${meta.readyRows || 0} invalid=${meta.invalidRows || 0} duplicate=${meta.duplicateRows || 0} phoneColumn=${meta.detectedPhoneColumn || "-"}`
  );
  logInfo(`UPLOAD CLEAN_READY PATH ${safePathForResponse(cleanReadyPath)}`);
  logInfo(`UPLOAD CLEAN_READY EXISTS ${fs.existsSync(cleanReadyPath)}`);
  return meta;
}

async function buildCleanReadyState(systemId = currentSystemId(), summary = null) {
  const id = normalizeSystemId(systemId) || DEFAULT_SYSTEM_ID;
  const cleanReadyPath = getSystemPaths(id).files.cleanReadyCsv;
  const fallback = summary || await readJson(getSystemPaths(id).files.intakeLatest, null) || {};
  return {
    systemId: id,
    cleanReadyExists: fs.existsSync(cleanReadyPath),
    cleanReadyPath: safePathForResponse(cleanReadyPath),
    readyRows: Number(fallback.readyRows || 0),
    invalidRows: Number(fallback.invalidRows || 0),
    duplicateRows: Number(fallback.duplicateRows || 0),
  };
}

async function loadCleanRows(systemId = currentSystemId()) {
  const cleanReadyPath = getSystemPaths(systemId).files.cleanReadyCsv;
  if (!fs.existsSync(cleanReadyPath)) return [];
  const raw = await fsp.readFile(cleanReadyPath, "utf8");
  const rows = parseCsvSync(raw, { columns: true, skip_empty_lines: true, bom: true, trim: true });
  return rows.map((row, idx) => ({
    sourceIndex: idx + 1,
    name: String(row.name || "").trim(),
    rawPhone: String(row.rawPhone || row.phone || "").trim(),
    normalizedPhone: String(row.phone || "").trim(),
    sourceLabel: normalizeSourceLabel(row.sourceLabel || row.source || "", "unknown"),
    consentStatus: normalizeConsentStatus(row.consentStatus || row.consent || ""),
    customerStatus: normalizeCustomerStatus(row.customerStatus || row.customer || ""),
  }));
}

function makeContactFirstName(jobLabel, alias, row) {
  const label = normalizeJobLabel(jobLabel, row?.name || row?.rawPhone || row?.normalizedPhone || "");
  const safeName = String(row?.name || "").trim() || String(row?.rawPhone || row?.normalizedPhone || "").trim() || `Contact ${row?.sourceIndex || ""}`;
  const text = `${alias} [${label}] ${safeName}`.trim();
  return text.slice(0, 64);
}

function buildContacts(rows, jobLabel = "", aliasPool = []) {
  const pool = Array.isArray(aliasPool) && aliasPool.length ? aliasPool : buildAliasPool("A", 8);
  return rows.map((row, index) => {
    const clientId = nextContactClientId();
    const sourceIndex = Math.max(1, Number(row?.sourceIndex || index + 1));
    const alias = pool[(sourceIndex - 1) % pool.length];
    const normalizedPhone = row.normalizedPhone || row.phone || "";
    return {
      ...row,
      normalizedPhone,
      clientId,
      contactAlias: alias,
      contact: new Api.InputPhoneContact({
        clientId,
        phone: normalizedPhone,
        firstName: makeContactFirstName(jobLabel, alias, row),
        lastName: "",
      }),
    };
  });
}

async function importBatch(client, batchRows, jobLabel = "", aliasPool = [], diagnostic = {}) {
  const sourceIndexes = (batchRows || [])
    .map((row) => Number(row?.sourceIndex || 0))
    .filter((value) => Number.isFinite(value) && value > 0);
  const sourceIndexStart = sourceIndexes.length ? Math.min(...sourceIndexes) : 0;
  const sourceIndexEnd = sourceIndexes.length ? Math.max(...sourceIndexes) : 0;
  const validRows = [];
  const invalidRows = [];
  for (const row of batchRows || []) {
    const phoneCheck = validateCheckerPhone(row);
    if (!phoneCheck.ok) {
      invalidRows.push(makeCheckerResultRow(row, "INVALID", {
        normalizedPhone: row?.normalizedPhone || row?.phone || "",
        reason: phoneCheck.reason,
      }));
    } else {
      validRows.push({ ...row, normalizedPhone: phoneCheck.normalizedPhone, phone: phoneCheck.normalizedPhone });
    }
  }

  if (!validRows.length) {
    logInfo(`ImportContacts result imported=0 retryContacts=0 users=0 invalid=${invalidRows.length}`);
    return {
      rows: invalidRows,
      importedCount: 0,
      retryContacts: [],
      usersCount: 0,
      invalidCount: invalidRows.length,
    };
  }

  const contacts = buildContacts(validRows, jobLabel, aliasPool);
  let result;
  const startedAt = Date.now();
  try {
    result = await client.invoke(
      new Api.contacts.ImportContacts({
        contacts: contacts.map((item) => item.contact),
      })
    );
  } catch (error) {
    await appendImportDiagnosticLog({
      systemId: currentSystemId(),
      accountId: diagnostic.accountId || "",
      accountLabel: diagnostic.accountLabel || "",
      accountPhoneMasked: maskPhone(diagnostic.accountPhone || ""),
      jobId: diagnostic.jobId || "",
      jobLabel: diagnostic.jobLabel || jobLabel || "",
      batchNumber: diagnostic.batchNumber || 0,
      batchStart: diagnostic.batchStart || 0,
      batchEnd: diagnostic.batchEnd || 0,
      sourceIndexStart,
      sourceIndexEnd,
      perRunLimit: Number(diagnostic.perRunLimit || 0),
      chunkSize: Number(diagnostic.chunkSize || 0),
      contactsInCall: contacts.length,
      importedCount: 0,
      usersCount: 0,
      retryContactsCount: 0,
      yesCount: 0,
      noCount: 0,
      retryCount: 0,
      retryRatio: null,
      floodWaitSec: parseFloodSeconds(error?.message || error?.errorMessage || error),
      errorCode: String(error?.code || error?.errorCode || error?.name || ""),
      errorMessage: sanitizeDiagnosticErrorMessage(error),
      durationMs: Date.now() - startedAt,
    });
    throw error;
  }

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

  logInfo(
    `ImportContacts result imported=${Array.isArray(result.imported) ? result.imported.length : 0} retryContacts=${Array.isArray(result.retryContacts) ? result.retryContacts.length : 0} users=${Array.isArray(result.users) ? result.users.length : 0} invalid=${invalidRows.length}`
  );

  const rows = contacts.map((item) => {
    const clientId = toSafeStringId(item.clientId);
    const matched = matchedByClientId.get(clientId);
    const isRetry = retrySet.has(clientId);
    return makeCheckerResultRow(item, matched ? "YES" : isRetry ? "RETRY" : "NO", {
      clientId,
      telegramUserId: matched?.userId || "",
      telegramUsername: matched?.username || "",
      telegramFirstName: matched?.firstName || "",
      telegramLastName: matched?.lastName || "",
      telegramPhone: matched?.phone || "",
    });
  });
  const yesCount = rows.filter((row) => row.status === "YES").length;
  const noCount = rows.filter((row) => row.status === "NO").length;
  const retryCount = rows.filter((row) => row.status === "RETRY").length;
  const retryRatio = rows.length ? retryCount / rows.length : 0;
  const diagnosticRecord = await appendImportDiagnosticLog({
    systemId: currentSystemId(),
    accountId: diagnostic.accountId || "",
    accountLabel: diagnostic.accountLabel || "",
    accountPhoneMasked: maskPhone(diagnostic.accountPhone || ""),
    jobId: diagnostic.jobId || "",
    jobLabel: diagnostic.jobLabel || jobLabel || "",
    batchNumber: diagnostic.batchNumber || 0,
    batchStart: diagnostic.batchStart || 0,
    batchEnd: diagnostic.batchEnd || 0,
    sourceIndexStart,
    sourceIndexEnd,
    perRunLimit: Number(diagnostic.perRunLimit || 0),
    chunkSize: Number(diagnostic.chunkSize || 0),
    contactsInCall: contacts.length,
    importedCount: Array.isArray(result.imported) ? result.imported.length : 0,
    usersCount: Array.isArray(result.users) ? result.users.length : 0,
    retryContactsCount: retrySet.size,
    yesCount,
    noCount,
    retryCount,
    retryRatio,
    floodWaitSec: 0,
    errorCode: "",
    errorMessage: "",
    durationMs: Date.now() - startedAt,
  });

  return {
    rows: [...rows, ...invalidRows],
    importedCount: Array.isArray(result.imported) ? result.imported.length : 0,
    retryContacts: Array.from(retrySet),
    usersCount: Array.isArray(result.users) ? result.users.length : 0,
    invalidCount: invalidRows.length,
    diagnostic: diagnosticRecord,
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
      sourceLabel: row.sourceLabel || "",
      consentStatus: row.consentStatus || "unknown",
      customerStatus: row.customerStatus || "new",
      leadScore: Number(row.leadScore || 0),
      nextAction: row.nextAction || "",
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
        "sourceLabel",
        "consentStatus",
        "customerStatus",
        "leadScore",
        "nextAction",
      ],
    }
  );
}

function writeRowsToXlsx(filePath, rows, columns) {
  const workbook = XLSX.utils.book_new();
  const normalizedRows = (Array.isArray(rows) ? rows : []).map((row) => {
    const out = {};
    for (const c of columns) out[c] = row?.[c] ?? "";
    return out;
  });
  const sheet = XLSX.utils.json_to_sheet(normalizedRows, { header: columns });
  XLSX.utils.book_append_sheet(workbook, sheet, "data");
  XLSX.writeFile(workbook, filePath);
}

const PROCESSED_COLUMNS = [
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
  "sourceLabel",
  "consentStatus",
  "customerStatus",
  "leadScore",
  "nextAction",
];

const REMAINING_COLUMNS = ["name", "phone", "rawPhone", "rowNumber", "sourceLabel", "consentStatus", "customerStatus"];

async function saveLatestOutputs(payload) {
  await writeJson(FILES.latestJson, payload);
  await fsp.writeFile(FILES.latestCsv, rowsToCsv(payload.rows || []), "utf8");
}

function byStatus(rows, status) {
  const wanted = String(status || "").toUpperCase();
  return (Array.isArray(rows) ? rows : []).filter((row) => String(row?.status || "").toUpperCase() === wanted);
}

function byNextAction(rows, action) {
  const wanted = String(action || "").toUpperCase();
  return (Array.isArray(rows) ? rows : []).filter((row) => String(row?.nextAction || "").toUpperCase() === wanted);
}

async function saveSegmentExports(rows) {
  const allRows = Array.isArray(rows) ? rows : [];
  const yesRows = byStatus(allRows, "YES");
  const noRows = byStatus(allRows, "NO");
  const retryRows = byStatus(allRows, "RETRY");
  const invalidRows = byStatus(allRows, "INVALID");
  const marketingAllowedRows = byNextAction(allRows, "MARKETING_ALLOWED");
  const winbackRows = byNextAction(allRows, "WINBACK");
  const retryLaterRows = byNextAction(allRows, "RETRY_LATER");

  await Promise.all([
    fsp.writeFile(FILES.yesOnlyCsv, rowsToCsv(yesRows), "utf8"),
    writeJson(FILES.yesOnlyJson, yesRows),
    fsp.writeFile(FILES.noOnlyCsv, rowsToCsv(noRows), "utf8"),
    writeJson(FILES.noOnlyJson, noRows),
    fsp.writeFile(FILES.retryOnlyCsv, rowsToCsv(retryRows), "utf8"),
    writeJson(FILES.retryOnlyJson, retryRows),
    fsp.writeFile(FILES.invalidOnlyCsv, rowsToCsv(invalidRows), "utf8"),
    writeJson(FILES.invalidOnlyJson, invalidRows),
    fsp.writeFile(FILES.marketingAllowedCsv, rowsToCsv(marketingAllowedRows), "utf8"),
    writeJson(FILES.marketingAllowedJson, marketingAllowedRows),
    fsp.writeFile(FILES.winbackCsv, rowsToCsv(winbackRows), "utf8"),
    writeJson(FILES.winbackJson, winbackRows),
    fsp.writeFile(FILES.retryLaterCsv, rowsToCsv(retryLaterRows), "utf8"),
    writeJson(FILES.retryLaterJson, retryLaterRows),
  ]);
}

function buildSegmentationSummary(rows) {
  const list = Array.isArray(rows) ? rows : [];
  const statusCount = { YES: 0, NO: 0, RETRY: 0, INVALID: 0 };
  const consentCount = { yes: 0, no: 0, unknown: 0 };
  const sourceMap = new Map();

  for (const row of list) {
    const status = String(row?.status || "").toUpperCase();
    if (status === "YES" || status === "NO" || status === "RETRY" || status === "INVALID") statusCount[status] += 1;
    const consent = normalizeConsentStatus(row?.consentStatus);
    consentCount[consent] += 1;

    const sourceLabel = normalizeSourceLabel(row?.sourceLabel, "unknown");
    const item = sourceMap.get(sourceLabel) || { sourceLabel, total: 0, yes: 0, no: 0, retry: 0, invalid: 0 };
    item.total += 1;
    if (status === "YES") item.yes += 1;
    if (status === "NO") item.no += 1;
    if (status === "RETRY") item.retry += 1;
    if (status === "INVALID") item.invalid += 1;
    sourceMap.set(sourceLabel, item);
  }

  const sources = Array.from(sourceMap.values()).map((item) => ({
    ...item,
    yesRate: item.total ? Number((item.yes / item.total).toFixed(4)) : 0,
  }));
  const topSourceByYesRate = sources
    .filter((item) => item.total > 0)
    .sort((a, b) => (b.yesRate - a.yesRate) || (b.total - a.total))
    .slice(0, 5);
  const lowQualitySources = sources
    .filter((item) => item.total >= 3 && (item.yesRate <= 0.1 || item.retry / item.total >= 0.5))
    .sort((a, b) => (a.yesRate - b.yesRate) || (b.retry - a.retry))
    .slice(0, 5);

  return {
    totalRows: list.length,
    telegramStatus: statusCount,
    consentStatus: consentCount,
    topSourceByYesRate,
    lowQualitySources,
  };
}

async function saveAllOutputs(rows) {
  await writeJson(FILES.allJson, rows || []);
  const allRows = rows || [];
  await fsp.writeFile(FILES.allCsv, rowsToCsv(allRows), "utf8");
  await fsp.writeFile(FILES.processedCsv, rowsToCsv(allRows), "utf8");
  writeRowsToXlsx(FILES.processedXlsx, allRows, PROCESSED_COLUMNS);
  const retryRows = (rows || []).filter((row) => row.status === "RETRY");
  await fsp.writeFile(FILES.retryCsv, rowsToCsv(retryRows), "utf8");
  await saveSegmentExports(allRows);
}

async function syncDerivedOutputFilesFromState() {
  const allRows = await readJson(FILES.allJson, []);
  if (Array.isArray(allRows)) {
    await fsp.writeFile(FILES.processedCsv, rowsToCsv(allRows), "utf8");
    writeRowsToXlsx(FILES.processedXlsx, allRows, PROCESSED_COLUMNS);
    const retryRows = allRows.filter((row) => row.status === "RETRY");
    await fsp.writeFile(FILES.retryCsv, rowsToCsv(retryRows), "utf8");
    await saveSegmentExports(allRows);
  }
  const current = await getCurrentJob();
  await writeRemainingCsv(current).catch(() => {});
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
    await unlinkIfExists(FILES.remainingXlsx);
    return [];
  }
  const doneSet = await getProcessedSourceIndexSet();
  const remainingRows = job.rows.filter((row) => !doneSet.has(String(row?.sourceIndex || '')));
  await fsp.writeFile(FILES.remainingCsv, cleanRowsToCsv(remainingRows), 'utf8');
  writeRowsToXlsx(FILES.remainingXlsx, remainingRows.map((row) => ({
    name: row.name,
    phone: row.phone || row.normalizedPhone || "",
    rawPhone: row.rawPhone || "",
    rowNumber: row.rowNumber || row.sourceIndex || "",
    sourceLabel: row.sourceLabel || "",
    consentStatus: row.consentStatus || "unknown",
    customerStatus: row.customerStatus || "new",
  })), REMAINING_COLUMNS);
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
      message: `รอบนี้พบรายการรอลองใหม่จำนวนมาก ระบบได้แยกไฟล์ที่ทำเสร็จแล้วและไฟล์ที่เหลือไว้ให้แล้ว (รออีก ${remainingSec} วินาที)`
    });
  }

  if (displayQueueStatus(job?.status) === 'PAUSED_TOO_MANY_RETRY') {
    alerts.push({
      level: 'warning',
      code: 'PAUSED_TOO_MANY_RETRY',
      title: 'Telegram กำลังให้พักบัญชี',
      message: `retry ratio ล่าสุด ${Number(job?.lastRetryRatio || 0).toFixed(2)} สูงเกิน ${HIGH_RETRY_MANUAL_PAUSE_RATIO.toFixed(2)} ระบบปิด Auto แล้ว ต้องกดทำรอบถัดไปเองเท่านั้น`
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

  if (displayQueueStatus(job?.status) === 'AUTH_REQUIRED') {
    alerts.push({
      level: 'warning',
      code: 'AUTH_REQUIRED',
      title: 'ต้องล็อกอินบัญชี Telegram ใหม่',
      message: job?.authRequiredMessage || 'บัญชี Telegram ยังไม่พร้อม กรุณารีล็อกอิน'
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
  const recoveryNotices = getRecoveryNotices();
  const [settings, job, latestBatch, runLog, accountsState] = await Promise.all([
    loadSettings(),
    getCurrentJob(),
    readJson(FILES.latestJson, null),
    readJson(FILES.runLog, []),
    loadAccountsState(),
  ]);
  const latestMeta = latestBatch?.meta || null;
  const allRows = await readJson(FILES.allJson, []);
  const importLog = await readImportDiagnosticLog(50);
  const segmentation = buildSegmentationSummary(allRows);
  const appState = accountsState?.appState || {};

  const remainingSec = computeRemainingSec(job?.nextRunAt);
  const progressPct = job?.totalRows ? Math.min(100, Math.round((Number(job.processedRows || 0) / Number(job.totalRows || 1)) * 100)) : 0;
  const runEvents = (Array.isArray(runLog) ? [...runLog] : []).slice(-15).reverse().map((entry) => ({
    at: entry.at || '',
    type: entry.type || 'event',
    text: entry.message || entry.code || [
      entry.jobLabel ? `งาน ${entry.jobLabel}` : '',
      entry.type,
      entry.batchNumber ? `รอบ ${entry.batchNumber}` : '',
      Number.isFinite(entry.seconds) ? `${entry.seconds} วิ` : '',
    ].filter(Boolean).join(' • '),
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
    latestBatch,
    cards: {
      done: Number(job?.processedRows || 0),
      remaining: Number(job?.remainingRows || 0),
      yes: Number(job?.matchedCountTotal || 0),
      no: Number(job?.unmatchedCountTotal || 0),
      retry: Number(job?.retryCountTotal || 0),
      invalid: Number(job?.invalidCountTotal || 0),
      floodWaitCount: Number(job?.floodWaitCount || 0),
      progressPct,
      account: job?.lockedAccountLabel || '-',
      jobLabel: job?.jobLabel || '-',
      queueStatus: displayQueueStatus(job?.status),
      totalChunks: Array.isArray(job?.chunks) ? job.chunks.length : 0,
      currentChunkIndex: Number(job?.currentChunkIndex || 0),
      currentRowIndex: Number(job?.currentRowIndex || 0),
      chunkFileSize: Number(job?.chunkFileSize || CHUNK_FILE_SIZE),
    },
    alerts: buildAlerts(job, latestMeta, settings, runLog, recoveryNotices),
    events,
    recovery: [...recoveryNotices].slice(-10).reverse(),
    importDiagnostics: {
      summary: buildImportDiagnosticSummary(importLog),
      recent: importLog,
    },
    lastJob: buildJobSnapshot(job) || appState.lastJobMeta || null,
    downloads: {
      processedCsv: fs.existsSync(FILES.processedCsv),
      processedXlsx: fs.existsSync(FILES.processedXlsx),
      remainingCsv: fs.existsSync(FILES.remainingCsv),
      remainingXlsx: fs.existsSync(FILES.remainingXlsx),
      runLog: fs.existsSync(FILES.runLog),
      allResultsCsv: fs.existsSync(FILES.allCsv),
      allResultsJson: fs.existsSync(FILES.allJson),
      yesOnlyCsv: fs.existsSync(FILES.yesOnlyCsv),
      yesOnlyJson: fs.existsSync(FILES.yesOnlyJson),
      noOnlyCsv: fs.existsSync(FILES.noOnlyCsv),
      noOnlyJson: fs.existsSync(FILES.noOnlyJson),
      retryOnlyCsv: fs.existsSync(FILES.retryOnlyCsv),
      retryOnlyJson: fs.existsSync(FILES.retryOnlyJson),
      invalidOnlyCsv: fs.existsSync(FILES.invalidOnlyCsv),
      invalidOnlyJson: fs.existsSync(FILES.invalidOnlyJson),
      marketingAllowedCsv: fs.existsSync(FILES.marketingAllowedCsv),
      marketingAllowedJson: fs.existsSync(FILES.marketingAllowedJson),
      winbackCsv: fs.existsSync(FILES.winbackCsv),
      winbackJson: fs.existsSync(FILES.winbackJson),
      retryLaterCsv: fs.existsSync(FILES.retryLaterCsv),
      retryLaterJson: fs.existsSync(FILES.retryLaterJson),
    },
    segmentation,
  };
}

async function buildAutoState(job) {
  const settings = await loadSettings();
  const remainingSec = computeRemainingSec(job?.nextRunAt);
  const retryManualPause = isHighRetryManualPause(job);
  return {
    autoRun: retryManualPause ? false : Boolean(settings.autoRun),
    status: retryManualPause ? "OFF" : displayAutoStatus(job?.autoStatus || (settings.autoRun ? "RUNNING" : "OFF")),
    queueStatus: displayQueueStatus(job?.status || "idle"),
    nextRunAt: job?.nextRunAt || "",
    waitUntil: job?.nextRunAt || "",
    remainingSec,
    nextAction: retryManualPause ? "manual_retry_required" : (settings.autoRun ? "auto_resume" : "manual_resume"),
    lockedAccountLabel: job?.lockedAccountLabel || "",
    lockedAccountPhone: job?.lockedAccountPhone || "",
    file: job?.sourceFile || "",
    fileName: job?.fileName || job?.sourceFile || "",
    jobLabel: job?.jobLabel || "",
    jobNote: job?.jobNote || "",
    floodWaitSec: displayQueueStatus(job?.status) === "WAITING_FLOOD" ? remainingSec : (job?.lastFloodWaitSec || 0),
    floodWaitCount: Number(job?.floodWaitCount || 0),
    retryRatio: job?.lastRetryRatio || 0,
    retryRatioHigh: retryManualPause,
    manualRetryRequired: Boolean(job?.manualRetryRequired || retryManualPause),
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
    fileName: job.fileName || job.sourceFile || "",
    jobLabel: normalizeJobLabel(job.jobLabel, job.fileName || job.sourceFile || ""),
    jobNote: normalizeJobNote(job.jobNote),
    selectedAccountId: job.selectedAccountId || job.lockedAccountId || "",
    status: job.status || "paused",
    autoStatus: job.autoStatus || (job.status === "completed" ? "OFF" : "PAUSED"),
    floodWaitCount: Number(job.floodWaitCount || 0),
    consecutiveFloodCount: Number(job.consecutiveFloodCount || 0),
    chunkFileSize: Number(job.chunkFileSize || CHUNK_FILE_SIZE),
    chunks: Array.isArray(job.chunks) ? job.chunks : [],
    currentChunkIndex: Number(job.currentChunkIndex || 0),
    currentRowIndex: Number(job.currentRowIndex || 0),
    rows: Array.isArray(job.rows) ? job.rows : (Array.isArray(currentJob?.rows) ? currentJob.rows : job.rows),
  };

  if (jobState && !currentJob) {
    await writeJson(FILES.currentJob, normalized).catch(() => {});
  } else if (currentJob && !jobState) {
    await writeJson(FILES.jobState, normalized).catch(() => {});
  }

  return normalized;
}

function buildJobSnapshot(job) {
  if (!job) return null;
  return {
    jobId: job.id || "",
    fileName: job.fileName || job.sourceFile || "",
    jobLabel: job.jobLabel || job.sourceFile || "",
    jobNote: job.jobNote || "",
    aliasGroup: job.aliasGroup || "A",
    createdAt: job.createdAt || "",
    selectedAccountId: job.selectedAccountId || job.lockedAccountId || "",
    selectedAccountLabel: job.lockedAccountLabel || "",
    total: Number(job.totalRows || 0),
    processed: Number(job.processedRows || 0),
    remaining: Number(job.remainingRows || 0),
    status: job.status || "",
    autoStatus: job.autoStatus || "",
    updatedAt: nowIso(),
  };
}

function warnStateWriteFailure(filePath, error, extra = {}) {
  const file = path.basename(filePath);
  const message = error?.message || "write state failed";
  const backupNote = extra.backupPath ? ` backup=${extra.backupPath}` : "";
  console.warn(`[WARN] state write failed file=${file}: ${message}${backupNote}`);
  pushRecoveryNotice(
    "warning",
    "STATE_WRITE_FAILED",
    "บันทึก state ไม่สำเร็จ",
    extra.backupPath
      ? `${file} เขียนทับไฟล์หลักไม่สำเร็จ ระบบบันทึกสำรองไว้ที่ ${path.basename(extra.backupPath)} และยังส่งผลลัพธ์รอบนี้กลับได้`
      : `${file} เขียนไม่สำเร็จ แต่ระบบยังส่งผลลัพธ์รอบนี้กลับได้`,
    { file, error: message, ...extra }
  );
}

async function writeStateBackupJson(filePath, data) {
  const backupPath = backupJsonPath(filePath);
  await fsp.mkdir(path.dirname(backupPath), { recursive: true });
  await fsp.writeFile(backupPath, JSON.stringify(data, null, 2), "utf8");
  return backupPath;
}

async function writeStateJsonBestEffort(filePath, data) {
  try {
    await writeJson(filePath, data);
    return null;
  } catch (error) {
    try {
      const backupPath = await writeStateBackupJson(filePath, data);
      warnStateWriteFailure(filePath, error, { backupPath });
      return {
        file: path.basename(filePath),
        message: `${path.basename(filePath)} เขียนไฟล์หลักไม่สำเร็จ จึงบันทึกสำรองไว้ที่ ${path.basename(backupPath)}`,
        error: error?.message || "write state failed",
        backup: path.basename(backupPath),
      };
    } catch (backupError) {
      warnStateWriteFailure(filePath, error, { backupError: backupError?.message || "backup write failed" });
      return {
        file: path.basename(filePath),
        message: `${path.basename(filePath)} เขียนไฟล์หลักและไฟล์สำรองไม่สำเร็จ แต่ API ยังไม่ล้ม`,
        error: error?.message || "write state failed",
        backupError: backupError?.message || "backup write failed",
      };
    }
  }
}

async function saveCurrentJob(job) {
  const warnings = [];
  for (const warning of await Promise.all([
    writeStateJsonBestEffort(FILES.currentJob, job),
    writeStateJsonBestEffort(FILES.jobState, job),
    writeJobStateMirror(job).catch((error) => {
      warnStateWriteFailure(path.join(getSystemPaths().jobsDir, String(job?.id || ""), "job_state.json"), error);
      return { file: "job_state.json", error: error?.message || "write state failed" };
    }),
    buildAutoState(job)
      .then((autoState) => writeStateJsonBestEffort(FILES.autoState, autoState))
      .catch((error) => {
        warnStateWriteFailure(FILES.autoState, error);
        return { file: path.basename(FILES.autoState), error: error?.message || "write state failed" };
      }),
  ])) {
    if (warning) warnings.push(warning);
  }
  await writeRemainingCsv(job).catch(() => {});
  try {
    const { accounts, appState } = await loadAccountsState();
    appState.lastJobMeta = buildJobSnapshot(job);
    const warning = await writeStateJsonBestEffort(FILES.appState, appState);
    if (warning) warnings.push(warning);
  } catch (error) {
    warnStateWriteFailure(FILES.appState, error);
  }
  if (warnings.length) job.stateWriteWarnings = warnings;
  else delete job.stateWriteWarnings;
  return job;
}

async function pauseHighRetryForManual(job, reason = "retry_ratio_high") {
  if (!isHighRetryManualPause(job)) return job;
  if (job.status === "paused_too_many_retry" && job.autoStatus === "OFF" && !job.nextRunAt) return job;

  await saveSettings({ autoRun: false }).catch((error) => warnStateWriteFailure(FILES.settings, error));
  const nextJob = {
    ...job,
    status: "paused_too_many_retry",
    autoStatus: "OFF",
    nextRunAt: "",
    manualRetryRequired: true,
    highRetryPauseReason: reason,
    updatedAt: nowIso(),
  };
  await saveCurrentJob(nextJob);
  await appendRunLog({
    type: "retry_manual_pause",
    jobId: nextJob.id,
    retryRatio: Number(nextJob.lastRetryRatio || 0),
    retryThreshold: HIGH_RETRY_MANUAL_PAUSE_RATIO,
    reason,
  });
  return nextJob;
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

async function inspectRunNextAuthState(job) {
  const { accounts, appState } = await loadAccountsState();
  const selectedAccountId = String(appState?.selectedAccountId || "");
  const account = accounts.find((item) => item.id === selectedAccountId) || null;
  const sessionPath = account?.id ? sessionFilePath(account.id) : "";
  const sessionPathExists = sessionPath ? fs.existsSync(sessionPath) : false;
  logInfo(
    `Run-next preflight selectedAccountId=${selectedAccountId || "-"} accountStatus=${account?.status || "missing"} sessionPathExists=${sessionPathExists} hasSessionEnc=${Boolean(account?.sessionEnc)} jobStatus=${job?.status || "-"} totalRows=${Number(job?.totalRows || 0)} processedRows=${Number(job?.processedRows || 0)}`
  );
  return { accounts, appState, selectedAccountId, account, sessionPath, sessionPathExists };
}

async function assertRunNextAuthReady(job) {
  const state = await inspectRunNextAuthState(job);
  const { selectedAccountId, account, sessionPathExists } = state;
  if (!selectedAccountId) throw makeAuthRequiredError("บัญชี Telegram ยังไม่พร้อม กรุณารีล็อกอิน", "needs_relogin");
  if (!account) throw makeAuthRequiredError("บัญชี Telegram ยังไม่พร้อม กรุณารีล็อกอิน", "auth_invalid");
  if (job?.lockedAccountId && selectedAccountId !== job.lockedAccountId) {
    throw new Error("บัญชีที่เลือกอยู่ไม่ตรงกับบัญชีที่ล็อกกับคิวนี้ กรุณากด 'ใช้บัญชีนี้' ให้ตรงก่อนรัน");
  }
  if (!String(account.apiId || "").trim()) {
    await markAccountForRelogin(account.id, "API_ID ของบัญชีนี้ไม่ครบ กรุณารีล็อกอินบัญชีใหม่", "auth_invalid");
    throw makeAuthRequiredError("บัญชี Telegram ยังไม่พร้อม กรุณารีล็อกอิน", "auth_invalid");
  }
  try {
    const apiHash = decryptText(account.apiHashEnc || "").trim();
    if (!apiHash) {
      await markAccountForRelogin(account.id, "API_HASH ของบัญชีนี้ไม่ครบ กรุณารีล็อกอินบัญชีใหม่", "auth_invalid");
      throw makeAuthRequiredError("บัญชี Telegram ยังไม่พร้อม กรุณารีล็อกอิน", "auth_invalid");
    }
  } catch (error) {
    if (error?.isAccountAuthError) throw error;
    await markAccountForRelogin(account.id, "API_HASH ของบัญชีนี้ถอดรหัสไม่ได้ กรุณารีล็อกอินบัญชีใหม่", "auth_invalid");
    throw makeAuthRequiredError("บัญชี Telegram ยังไม่พร้อม กรุณารีล็อกอิน", "auth_invalid");
  }

  try {
    const { sessionString, source } = await loadAccountSessionString(account);
    if (!sessionString) {
      await markAccountForRelogin(account.id, "บัญชีนี้ยังไม่มี session ที่ใช้งานได้ กรุณาล็อกอินใหม่", "needs_relogin");
      throw makeAuthRequiredError("บัญชี Telegram ยังไม่พร้อม กรุณารีล็อกอิน", "needs_relogin");
    }
    logInfo(`Run-next preflight sessionSource=${source || "unknown"} sessionPathExists=${sessionPathExists}`);
  } catch (error) {
    if (error?.isAccountAuthError || isAuthSessionError(error)) {
      const message = authRequiredMessage(error);
      await markAccountForRelogin(account.id, message, /UNSUPPORTED STATE OR UNABLE TO AUTHENTICATE DATA/i.test(normalizeErrorMessage(error)) ? "auth_invalid" : "needs_relogin");
      throw makeAuthRequiredError(message, /UNSUPPORTED STATE OR UNABLE TO AUTHENTICATE DATA/i.test(normalizeErrorMessage(error)) ? "auth_invalid" : "needs_relogin");
    }
    throw error;
  }

  return account;
}

async function findPhoneUsageInOtherSystem(phone, ownSystemId = currentSystemId()) {
  const normalizedPhone = String(phone || "").trim();
  const own = normalizeSystemId(ownSystemId) || DEFAULT_SYSTEM_ID;
  if (!normalizedPhone) return null;
  for (const systemId of SYSTEM_IDS) {
    if (systemId === own) continue;
    const usage = await withSystem(systemId, async () => {
      const job = await getCurrentJob();
      if (!job) return null;
      const status = String(job.status || "").toLowerCase();
      const autoStatus = String(job.autoStatus || "").toUpperCase();
      const activeStatus = !["completed"].includes(status) && !["OFF"].includes(autoStatus);
      const runnableStatus = ["ready", "running", "waiting_flood", "paused_retry_cooldown", "waiting_retry_cooldown"].includes(status);
      if ((activeStatus || runnableStatus) && String(job.lockedAccountPhone || "") === normalizedPhone) {
        return { systemId, jobId: job.id || "", jobLabel: job.jobLabel || "", status: job.status || "", autoStatus: job.autoStatus || "" };
      }
      return null;
    });
    if (usage) return usage;
  }
  return null;
}

async function assertAccountPhoneNotRunningElsewhere(account) {
  const usage = await findPhoneUsageInOtherSystem(account?.phone);
  if (!usage) return;
  throw new Error(`บัญชี Telegram เบอร์ ${account.phone} ถูกใช้อยู่ในระบบ ${usage.systemId} (${usage.status}/${usage.autoStatus}) ห้ามรันพร้อมกันหลายระบบ`);
}

class MissingCleanReadyError extends Error {
  constructor(details) {
    super("ยังไม่มี clean_ready.csv กรุณาอัปโหลดไฟล์ก่อน");
    this.name = "MissingCleanReadyError";
    this.details = details;
  }
}

async function getCleanReadyDebug(systemId = currentSystemId()) {
  const id = normalizeSystemId(systemId) || DEFAULT_SYSTEM_ID;
  const paths = getSystemPaths(id);
  const expectedCleanReadyPath = paths.files.cleanReadyCsv;
  const [latestIntakeSummary, summaryFallback] = await Promise.all([
    readJson(paths.files.intakeLatest, null),
    readJson(paths.files.intakeSummaryJson, null),
  ]);
  return {
    systemId: id,
    expectedCleanReadyPath: safePathForResponse(expectedCleanReadyPath),
    cleanReadyPath: safePathForResponse(expectedCleanReadyPath),
    existsCleanReady: fs.existsSync(expectedCleanReadyPath),
    latestIntakeSummary: latestIntakeSummary || summaryFallback || null,
  };
}

async function createJobFromLatestClean(systemId = currentSystemId()) {
  const id = normalizeSystemId(systemId) || DEFAULT_SYSTEM_ID;
  return withSystem(id, async () => {
  const cleanReadyDebug = await getCleanReadyDebug(id);
  logInfo(`CREATE JOB EXPECTED CLEAN_READY PATH ${cleanReadyDebug.cleanReadyPath}`);
  logInfo(`CREATE JOB CLEAN_READY EXISTS ${cleanReadyDebug.existsCleanReady}`);
  if (!cleanReadyDebug.existsCleanReady) {
    logInfo(`CREATE JOB RESULT missing_clean_ready system=${id} path=${cleanReadyDebug.cleanReadyPath}`);
    throw new MissingCleanReadyError(cleanReadyDebug);
  }
  const cleanRows = await loadCleanRows(id);
  if (!cleanRows.length) {
    logInfo(`CREATE JOB RESULT empty_clean_ready system=${id} path=${cleanReadyDebug.cleanReadyPath}`);
    throw new Error("clean_ready.csv ไม่มีข้อมูลพร้อมใช้");
  }
  const selectedAccount = await getSelectedAccount(true);
  await assertAccountPhoneNotRunningElsewhere(selectedAccount);
  const settings = await loadSettings();
  const intakeLatest = await readJson(FILES.intakeLatest, null);
  const fileName = intakeLatest?.fileName || intakeLatest?.originalFilename || "clean_ready.csv";
  const jobLabel = normalizeJobLabel(intakeLatest?.jobLabel, fileName);
  const jobNote = normalizeJobNote(intakeLatest?.jobNote);
  const aliasGroup = deriveAliasGroup(jobLabel, jobNote);
  const aliasPool = buildAliasPool(aliasGroup, 8);
  const chunkFileSize = CHUNK_FILE_SIZE;
  const chunks = buildChunkSpecs(cleanRows, chunkFileSize);
  const masterFile = String(intakeLatest?.masterFile || "");

  const job = {
    id: makeId("job"),
    sourceFile: "clean_ready.csv",
    fileName,
    jobLabel,
    jobNote,
    aliasGroup,
    aliasPool,
    masterFile,
    chunkFileSize,
    chunks,
    currentChunkIndex: chunks.length ? 1 : 0,
    currentRowIndex: 0,
    totalRows: cleanRows.length,
    nextIndex: 0,
    processedRows: 0,
    remainingRows: cleanRows.length,
    selectedAccountId: selectedAccount.id,
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
    invalidCountTotal: 0,
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
  await writeChunkFiles(job);
  await saveCurrentJob(job);
  await unlinkIfExists(FILES.latestJson);
  await unlinkIfExists(FILES.latestCsv);
  await appendRunLog({
    type: "job_created",
    jobId: job.id,
    jobLabel: job.jobLabel,
    fileName: job.fileName,
    lockedAccount: job.lockedAccountLabel,
  });

  logInfo(`CREATE JOB RESULT ok system=${id} jobId=${job.id} totalRows=${job.totalRows}`);
  return job;
  });
}

async function processNextBatch() {
  const job = await getCurrentJob();
  if (!job) throw new Error("ยังไม่มีคิวงาน กรุณาล้างไฟล์และสร้างคิวก่อน");
  if (!Array.isArray(job.rows) || !job.rows.length) throw new Error("ข้อมูลคิวหาย กรุณาสร้างคิวใหม่");
  if (job.status === "completed") throw new Error("คิวนี้ทำครบแล้ว");
  if (job.status === "paused") throw new Error("คิวนี้ถูกพักอยู่");
  if (job.status === "paused_retry_cooldown" && job.nextRunAt && Date.now() < new Date(job.nextRunAt).getTime()) {
    throw new Error(`คิวพักรอลองใหม่ถึง ${job.nextRunAt}`);
  }
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
  await assertAccountPhoneNotRunningElsewhere(account);
  const client = await ensureAuthorizedAccountClient(account);

  const chunkStates = Array.isArray(job.chunks) && job.chunks.length
    ? job.chunks.map((item) => ({ ...item }))
    : buildChunkSpecs(job.rows || [], Number(job.chunkFileSize || CHUNK_FILE_SIZE));
  let activeChunk = chunkStates.find((item) => item.status !== "DONE" && item.status !== "FAILED") || null;
  if (!activeChunk) throw new Error("ไม่พบ chunk ที่รันต่อได้");
  const activeChunkIndex = Math.max(0, Number(activeChunk.chunkIndex || 1) - 1);
  activeChunk.status = "RUNNING";
  const chunkStart = Number(activeChunk.start || 0);
  const chunkEndExclusive = Number(activeChunk.endExclusive || 0);
  const offsetInChunk = Math.max(0, Number(activeChunk.currentRowIndex || 0));
  const start = Math.min(chunkStart + offsetInChunk, chunkEndExclusive);
  const end = Math.min(start + Number(job.perRunLimit || 100), chunkEndExclusive);
  const sourceBatch = job.rows.slice(start, end);
  if (!sourceBatch.length) throw new Error("ไม่พบข้อมูลในรอบถัดไป");

  const chunks = chunkArray(sourceBatch, Number(job.chunkSize || 1));
  const batchNumber = Math.floor(start / Number(job.perRunLimit || 100)) + 1;
  const currentBatchRows = [];
  const importDiagnostics = [];
  let importedCount = 0;
  let batchRetryCount = 0;
  const aliasPool = Array.isArray(job.aliasPool) && job.aliasPool.length
    ? job.aliasPool
    : buildAliasPool(deriveAliasGroup(job.jobLabel, job.jobNote), 8);

  let callOffset = 0;
  for (const chunk of chunks) {
    const batchStart = start + callOffset + 1;
    const batchEnd = start + callOffset + chunk.length;
    callOffset += chunk.length;
    const result = await importBatch(client, chunk, job.jobLabel || job.sourceFile || "งาน", aliasPool, {
      systemId: currentSystemId(),
      accountId: account.id,
      accountLabel: account.label,
      accountPhone: account.phone,
      jobId: job.id,
      jobLabel: job.jobLabel || job.sourceFile || "",
      batchNumber,
      batchStart,
      batchEnd,
      perRunLimit: Number(job.perRunLimit || 0),
      chunkSize: Number(job.chunkSize || 0),
    });
    currentBatchRows.push(...result.rows);
    importedCount += result.importedCount;
    batchRetryCount += result.rows.filter((row) => row.status === "RETRY").length;
    if (result.diagnostic) importDiagnostics.push(result.diagnostic);
  }

  const matchedCount = currentBatchRows.filter((row) => row.status === "YES").length;
  const retryCount = currentBatchRows.filter((row) => row.status === "RETRY").length;
  const unmatchedCount = currentBatchRows.filter((row) => row.status === "NO").length;
  const invalidCount = currentBatchRows.filter((row) => row.status === "INVALID").length;

  const allRows = await readJson(FILES.allJson, []);
  allRows.push(...currentBatchRows);
  await saveAllOutputs(allRows);

  const now = nowIso();
  const nextIndex = end;
  const processedInChunk = Math.max(0, end - chunkStart);
  const chunkDone = end >= chunkEndExclusive;
  const updatedChunk = {
    ...activeChunk,
    currentRowIndex: processedInChunk,
    status: chunkDone ? "DONE" : "PENDING",
  };
  chunkStates[activeChunkIndex] = updatedChunk;
  const nextChunk = chunkStates.find((item) => item.status !== "DONE" && item.status !== "FAILED") || null;
  const retryRatio = currentBatchRows.length ? retryCount / currentBatchRows.length : 0;

  let nextStatus = nextChunk ? "ready" : "completed";
  let nextAutoStatus = job.autoStatus || "OFF";
  let nextRunAt = "";

  const highRetryManualPause = nextStatus !== "completed" && retryRatio >= HIGH_RETRY_MANUAL_PAUSE_RATIO;
  if (highRetryManualPause) {
    await saveSettings({ autoRun: false }).catch((error) => warnStateWriteFailure(FILES.settings, error));
    nextStatus = "paused_too_many_retry";
    nextAutoStatus = "OFF";
    nextRunAt = "";
  } else if (nextStatus !== "completed" && retryRatio > Number(job.retryRatioThreshold || 0.2)) {
    nextStatus = "paused_retry_cooldown";
    nextAutoStatus = "WAITING_RETRY";
    nextRunAt = new Date(Date.now() + Number(job.retryPauseSec || 300) * 1000).toISOString();
  } else if (nextStatus !== "completed" && nextAutoStatus === "RUNNING") {
    nextRunAt = new Date(Date.now() + Number(job.delayBetweenRunsSec || 60) * 1000).toISOString();
  }

  const nextJob = {
    ...job,
    chunks: chunkStates,
    currentChunkIndex: nextChunk ? Number(nextChunk.chunkIndex || 0) : Number(updatedChunk.chunkIndex || 0),
    currentRowIndex: nextChunk ? Number(nextChunk.currentRowIndex || 0) : Number(updatedChunk.currentRowIndex || 0),
    nextIndex,
    processedRows: nextIndex,
    remainingRows: Math.max(0, Number(job.totalRows || 0) - nextIndex),
    matchedCountTotal: Number(job.matchedCountTotal || 0) + matchedCount,
    unmatchedCountTotal: Number(job.unmatchedCountTotal || 0) + unmatchedCount,
    retryCountTotal: Number(job.retryCountTotal || 0) + retryCount,
    invalidCountTotal: Number(job.invalidCountTotal || 0) + invalidCount,
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
    manualRetryRequired: highRetryManualPause,
    highRetryPauseReason: highRetryManualPause ? "retry_ratio_high" : "",
    updatedAt: now,
  };

  const latestPayload = {
    meta: {
      sourceFile: nextJob.sourceFile,
      fileName: nextJob.fileName || nextJob.sourceFile,
      jobLabel: nextJob.jobLabel || nextJob.sourceFile,
      jobNote: nextJob.jobNote || "",
      aliasGroup: nextJob.aliasGroup || "A",
      aliasPool: Array.isArray(nextJob.aliasPool) ? nextJob.aliasPool : buildAliasPool(nextJob.aliasGroup || "A", 8),
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
      invalidCount,
      matchedCountTotal: nextJob.matchedCountTotal,
      unmatchedCountTotal: nextJob.unmatchedCountTotal,
      retryCountTotal: nextJob.retryCountTotal,
      invalidCountTotal: nextJob.invalidCountTotal,
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
      manualRetryRequired: Boolean(nextJob.manualRetryRequired),
      highRetryPauseReason: nextJob.highRetryPauseReason || "",
      nextRunAt: nextJob.nextRunAt,
      floodWaitCount: Number(nextJob.floodWaitCount || 0),
      chunkFileSize: Number(nextJob.chunkFileSize || CHUNK_FILE_SIZE),
      totalChunks: Array.isArray(nextJob.chunks) ? nextJob.chunks.length : 0,
      currentChunkIndex: Number(nextJob.currentChunkIndex || 0),
      currentRowIndex: Number(nextJob.currentRowIndex || 0),
    },
    rows: currentBatchRows,
  };

  await saveLatestOutputs(latestPayload);
  await saveCurrentJob(nextJob);
  await appendRunLog({
    type: "batch_done",
    jobId: job.id,
    jobLabel: nextJob.jobLabel || "",
    fileName: nextJob.fileName || "",
    batchNumber,
    matchedCount,
    unmatchedCount,
    retryCount,
    invalidCount,
    retryRatio,
    processedRows: nextJob.processedRows,
    floodWaitCount: Number(nextJob.floodWaitCount || 0),
    nextStatus,
    manualRetryRequired: highRetryManualPause,
  });
  await appendRunLog({
    type: "run_summary",
    processed: nextJob.processedRows,
    yes: nextJob.matchedCountTotal,
    no: nextJob.unmatchedCountTotal,
    retry: nextJob.retryCountTotal,
    invalid: nextJob.invalidCountTotal,
    floodWaitCount: Number(nextJob.floodWaitCount || 0),
  });
  logInfo(
    `Run summary processed=${nextJob.processedRows} yes=${nextJob.matchedCountTotal} no=${nextJob.unmatchedCountTotal} retry=${nextJob.retryCountTotal} invalid=${nextJob.invalidCountTotal} floodWaitCount=${Number(nextJob.floodWaitCount || 0)}`
  );
  if (importDiagnostics.length) {
    const diagnosticSummary = buildImportDiagnosticSummary(importDiagnostics);
    logInfo(
      `Import diagnostic summary system=${currentSystemId()} jobId=${job.id} batch=${batchNumber} avgRetryRatio=${diagnosticSummary.avgRetryRatio.toFixed(2)} imported=${diagnosticSummary.lastImportedCount} users=${diagnosticSummary.lastUsersCount} retryContacts=${diagnosticSummary.lastRetryContactsCount} calls=${diagnosticSummary.totalCalls} contactsInLastCall=${diagnosticSummary.lastContactsInCall}`
    );
  }

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
  destination: (req, _file, callback) => {
    const systemId = normalizeSystemId(req.params?.systemId || req.systemId || currentSystemId()) || DEFAULT_SYSTEM_ID;
    const dir = getSystemPaths(systemId).rawDir;
    fs.mkdir(dir, { recursive: true }, (error) => callback(error, dir));
  },
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

async function processIntakeFile(filePath, originalName = "", options = {}) {
  const rows = await parseInputRowsFromFile(filePath, originalName);
  const result = normalizeInputRows(rows);
  const meta = await saveIntakeOutputs(originalName || path.basename(filePath), result, options);
  return { meta, result };
}

async function persistMasterUpload(filePath, originalName = "") {
  const ext = path.extname(originalName || filePath || "").toLowerCase() || ".csv";
  const base = sanitizeFileName(fileStem(originalName || path.basename(filePath)) || "master");
  const targetName = `${base}_${Date.now()}${ext}`;
  const targetPath = path.join(getSystemPaths().uploadsDir, targetName);
  await fsp.copyFile(filePath, targetPath);
  return targetPath;
}

function buildChunkSpecs(rows = [], chunkFileSize = CHUNK_FILE_SIZE) {
  const specs = [];
  const size = Math.max(1000, Math.floor(Number(chunkFileSize || CHUNK_FILE_SIZE)));
  const total = Array.isArray(rows) ? rows.length : 0;
  let chunkIndex = 0;
  for (let start = 0; start < total; start += size) {
    const endExclusive = Math.min(total, start + size);
    chunkIndex += 1;
    specs.push({
      chunkIndex,
      fileName: `chunk_${String(chunkIndex).padStart(3, "0")}.csv`,
      start,
      endExclusive,
      totalRows: endExclusive - start,
      currentRowIndex: 0,
      status: "PENDING",
    });
  }
  return specs;
}

async function writeChunkFiles(job) {
  const jobId = String(job?.id || "");
  if (!jobId) return [];
  const dir = path.join(getSystemPaths().chunksDir, jobId);
  await fsp.mkdir(dir, { recursive: true });
  const specs = Array.isArray(job?.chunks) ? job.chunks : [];
  const rows = Array.isArray(job?.rows) ? job.rows : [];
  for (const spec of specs) {
    const part = rows.slice(Number(spec.start || 0), Number(spec.endExclusive || 0));
    const filePath = path.join(dir, spec.fileName);
    await fsp.writeFile(filePath, cleanRowsToCsv(part), "utf8");
  }
  return specs;
}

async function writeJobStateMirror(job) {
  const jobId = String(job?.id || "");
  if (!jobId) return;
  const dir = path.join(getSystemPaths().jobsDir, jobId);
  await fsp.mkdir(dir, { recursive: true });
  const state = {
    jobId,
    masterFile: job.masterFile || "",
    totalRows: Number(job.totalRows || 0),
    chunkFileSize: Number(job.chunkFileSize || CHUNK_FILE_SIZE),
    currentChunkIndex: Number(job.currentChunkIndex || 0),
    currentRowIndex: Number(job.currentRowIndex || 0),
    processed: Number(job.processedRows || 0),
    yes: Number(job.matchedCountTotal || 0),
    no: Number(job.unmatchedCountTotal || 0),
    retry: Number(job.retryCountTotal || 0),
    invalid: Number(job.invalidCountTotal || 0),
    status: String(job.status || ""),
    retryRatio: Number(job.lastRetryRatio || 0),
    lastRunAt: job.lastGeneratedAt || "",
    nextRunAt: job.nextRunAt || "",
    selectedAccountId: job.selectedAccountId || job.lockedAccountId || "",
    chunks: Array.isArray(job.chunks) ? job.chunks : [],
    updatedAt: nowIso(),
  };
  return writeStateJsonBestEffort(path.join(dir, "job_state.json"), state);
}

async function markJobAuthRequired(job, message = "บัญชี Telegram ยังไม่พร้อม กรุณารีล็อกอิน") {
  if (!job) return null;
  await saveSettings({ autoRun: false }).catch((error) => warnStateWriteFailure(FILES.settings, error));
  const nextJob = {
    ...job,
    status: "AUTH_REQUIRED",
    autoStatus: "OFF",
    nextRunAt: "",
    authRequiredAt: nowIso(),
    authRequiredMessage: message,
    updatedAt: nowIso(),
  };
  await saveCurrentJob(nextJob);
  await appendRunLog({
    type: "auth_required",
    jobId: job.id,
    message,
    selectedAccountId: job.selectedAccountId || job.lockedAccountId || "",
  });
  return nextJob;
}

async function splitCsvFile(filePath, rowsPerFile = 50000) {
  const safeRowsPerFile = Math.max(1000, Math.floor(Number(rowsPerFile || 50000)));
  const outDir = path.join(getSystemPaths().rawDir, `split_${Date.now()}`);
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

async function runNextForCurrentSystem(res) {
  if (isSystemProcessing()) return res.status(409).json({ error: "ระบบนี้กำลังประมวลผลอยู่" });
  setSystemProcessing(true);
  try {
    const staleChecked = await pauseHighRetryForManual(await markFloodStateStaleIfNeeded(await getCurrentJob()), "manual_run");
    if (staleChecked?.status === "paused_flood_stale") {
      return res.status(409).json({ error: "คิวถูกพักเพราะ FloodWait ค้างนาน กรุณารีเซ็ต flood state ก่อน" });
    }
    await assertRunNextAuthReady(staleChecked);
    const result = await processNextBatch();
    return res.json({ ok: true, message: "ทำรอบถัดไปเรียบร้อย", ...result });
  } catch (error) {
    const job = await getCurrentJob();
    if (job && (error?.isAccountAuthError || isAuthSessionError(error))) {
      const message = error?.code === "AUTH_REQUIRED" || error?.type === "AUTH_REQUIRED"
        ? (error.message || authRequiredMessage(error))
        : authRequiredMessage(error);
      const nextJob = await markJobAuthRequired(job, message);
      await appendRunLog({ type: "account_auth_invalid", jobId: job.id, message });
      return res.status(409).json({
        ok: false,
        type: "AUTH_REQUIRED",
        code: "AUTH_REQUIRED",
        message,
        error: message,
        job: nextJob,
      });
    }
    if (job && parseFloodSeconds(error?.message)) {
      const nextJob = await handleFloodWait(job, error);
      return res.status(429).json({ error: extractTelegramError(error), job: nextJob });
    }
    const friendly = extractTelegramError(error) || "ทำรอบถัดไปไม่สำเร็จ";
    const statusCode = /ยังไม่ได้เลือกบัญชี|ไม่พบบัญชี|ไม่ตรงกับบัญชี|ห้ามรันพร้อมกัน/.test(friendly) ? 409 : 400;
    return res.status(statusCode).json({ error: friendly });
  } finally {
    setSystemProcessing(false);
  }
}

async function startAutoForCurrentSystem(res) {
  const job = await pauseHighRetryForManual(await markFloodStateStaleIfNeeded(await getCurrentJob()), "start_auto");
  if (!job) return res.status(404).json({ error: "ยังไม่มีคิวงาน" });
  if (job.status === "paused_flood_stale") {
    return res.status(409).json({ error: "คิวถูกพักเพราะ FloodWait ค้างนาน กรุณารีเซ็ต flood state ก่อน" });
  }
  if (isHighRetryManualPause(job)) {
    return res.status(409).json({
      error: "retry ratio ล่าสุดสูงมาก Telegram กำลังให้พักบัญชี ปิด Auto แล้ว กรุณากดทำรอบถัดไปเองเมื่อพร้อม",
      code: "PAUSED_TOO_MANY_RETRY",
      job,
    });
  }
  const account = await getLockedAccountForJob(job);
  await assertAccountPhoneNotRunningElsewhere(account);
  const settings = await saveSettings({ autoRun: true });
  const now = Date.now();
  const waitFloodTs = job.nextRunAt ? new Date(job.nextRunAt).getTime() : 0;
  const waitRetryTs = job.nextRunAt ? new Date(job.nextRunAt).getTime() : 0;
  const keepFloodWaiting = job.status === "waiting_flood" && waitFloodTs && now < waitFloodTs;
  const keepRetryWaiting = (job.status === "waiting_retry_cooldown" || job.status === "paused_retry_cooldown") && waitRetryTs && now < waitRetryTs;
  const nextStatus = keepFloodWaiting
    ? "waiting_flood"
    : keepRetryWaiting
      ? "paused_retry_cooldown"
      : (["paused", "ready", "waiting_retry_cooldown", "paused_retry_cooldown", "waiting_flood"].includes(job.status) ? "ready" : job.status);
  const nextAutoStatus = keepFloodWaiting ? "WAITING_FLOOD" : (keepRetryWaiting ? "WAITING_RETRY" : "RUNNING");
  const nextJob = {
    ...job,
    autoStatus: nextAutoStatus,
    status: nextStatus,
    nextRunAt: keepFloodWaiting || keepRetryWaiting ? job.nextRunAt : (job.nextRunAt || new Date().toISOString()),
    updatedAt: nowIso(),
  };
  await saveCurrentJob(nextJob);
  return res.json({ ok: true, settings, job: nextJob });
}

async function pauseAutoForCurrentSystem(res) {
  const job = await getCurrentJob();
  if (!job) return res.status(404).json({ error: "ยังไม่มีคิวงาน" });
  const settings = await saveSettings({ autoRun: false });
  const nextJob = { ...job, autoStatus: "PAUSED", status: "paused", updatedAt: nowIso() };
  await saveCurrentJob(nextJob);
  return res.json({ ok: true, settings, job: nextJob });
}

async function resumeAutoForCurrentSystem(res) {
  const job = await pauseHighRetryForManual(await markFloodStateStaleIfNeeded(await getCurrentJob()), "resume_auto");
  if (!job) return res.status(404).json({ error: "ยังไม่มีคิวงาน" });
  if (job.status === "paused_flood_stale") {
    return res.status(409).json({ error: "คิวถูกพักเพราะ FloodWait ค้างนาน กรุณารีเซ็ต flood state ก่อน" });
  }
  if (isHighRetryManualPause(job)) {
    return res.status(409).json({
      error: "retry ratio ล่าสุดสูงมาก Telegram กำลังให้พักบัญชี ปิด Auto แล้ว กรุณากดทำรอบถัดไปเองเมื่อพร้อม",
      code: "PAUSED_TOO_MANY_RETRY",
      job,
    });
  }
  const account = await getLockedAccountForJob(job);
  await assertAccountPhoneNotRunningElsewhere(account);
  const settings = await saveSettings({ autoRun: true });
  const now = Date.now();
  const waitFloodTs = job.nextRunAt ? new Date(job.nextRunAt).getTime() : 0;
  const waitRetryTs = job.nextRunAt ? new Date(job.nextRunAt).getTime() : 0;
  const keepFloodWaiting = job.status === "waiting_flood" && waitFloodTs && now < waitFloodTs;
  const keepRetryWaiting = (job.status === "waiting_retry_cooldown" || job.status === "paused_retry_cooldown") && waitRetryTs && now < waitRetryTs;
  const nextJob = {
    ...job,
    autoStatus: keepFloodWaiting ? "WAITING_FLOOD" : (keepRetryWaiting ? "WAITING_RETRY" : "RUNNING"),
    status: keepFloodWaiting ? "waiting_flood" : (keepRetryWaiting ? "paused_retry_cooldown" : "ready"),
    nextRunAt: keepFloodWaiting || keepRetryWaiting ? job.nextRunAt : new Date().toISOString(),
    updatedAt: nowIso(),
  };
  await saveCurrentJob(nextJob);
  return res.json({ ok: true, settings, job: nextJob });
}

async function resumeManualForCurrentSystem(res) {
  const job = await getCurrentJob();
  if (!job) return res.status(404).json({ error: "ยังไม่มีคิวงาน" });
  if (job.status === "paused_flood_stale") {
    return res.status(409).json({ error: "คิวถูกพักเพราะ FloodWait ค้างนาน ห้ามปลดพักแบบ manual" });
  }
  if (job.status === "paused_too_many_retry") {
    return res.status(409).json({ error: "คิวถูกพักเพราะ RETRY สูง ห้ามปลดพักแบบ manual" });
  }
  const waitTs = job.nextRunAt ? new Date(job.nextRunAt).getTime() : 0;
  const waitingStatus = ["waiting_flood", "waiting_retry_cooldown", "paused_retry_cooldown"].includes(job.status);
  if (waitingStatus && waitTs && Date.now() < waitTs) {
    return res.status(409).json({ error: `คิวยังต้องรอถึง ${job.nextRunAt}` });
  }
  if (job.status !== "paused") return res.status(400).json({ error: "ปลดพักแบบ manual ได้เฉพาะคิวที่ paused เท่านั้น" });

  const settings = await saveSettings({ autoRun: false });
  const nextJob = {
    ...job,
    status: "ready",
    autoStatus: "OFF",
    nextRunAt: "",
    updatedAt: nowIso(),
  };
  await saveCurrentJob(nextJob);
  await appendRunLog({
    type: "manual_resume",
    jobId: job.id,
    jobLabel: job.jobLabel || "",
    fileName: job.fileName || job.sourceFile || "",
    processedRows: Number(job.processedRows || 0),
    remainingRows: Number(job.remainingRows || 0),
  });
  return res.json({
    ok: true,
    message: "ปลดพักคิวสำหรับตรวจมือแล้ว",
    job: nextJob,
    settings,
  });
}

async function resetJobForCurrentSystem(res) {
  if (isSystemProcessing()) return res.status(409).json({ error: "ระบบนี้กำลังประมวลผลอยู่ ล้างคิวตอนนี้ไม่ได้" });
  try {
    for (const filePath of [
      FILES.currentJob,
      FILES.jobState,
      FILES.latestJson,
      FILES.latestCsv,
      FILES.allJson,
      FILES.allCsv,
      FILES.processedCsv,
      FILES.processedXlsx,
      FILES.retryCsv,
      FILES.yesOnlyCsv,
      FILES.yesOnlyJson,
      FILES.noOnlyCsv,
      FILES.noOnlyJson,
      FILES.retryOnlyCsv,
      FILES.retryOnlyJson,
      FILES.invalidOnlyCsv,
      FILES.invalidOnlyJson,
      FILES.marketingAllowedCsv,
      FILES.marketingAllowedJson,
      FILES.winbackCsv,
      FILES.winbackJson,
      FILES.retryLaterCsv,
      FILES.retryLaterJson,
      FILES.autoState,
      FILES.remainingCsv,
      FILES.remainingXlsx,
      FILES.cleanReadyCsv,
      FILES.invalidRowsCsv,
      FILES.duplicatePhonesCsv,
      FILES.cleanDebugCsv,
      FILES.cleanRejectsCsv,
      FILES.intakeSummaryJson,
      FILES.intakeLatest,
    ]) await unlinkIfExists(filePath);
    return res.json({ ok: true, message: "ล้างคิวและผลลัพธ์ของระบบนี้เรียบร้อย" });
  } catch (error) {
    return res.status(500).json({ error: error.message || "ล้างคิวไม่สำเร็จ" });
  }
}

function getDownloadTarget(name) {
  const allowed = new Map([
    ["telegram_matches.csv", { path: FILES.latestCsv, displayName: "telegram_matches.csv" }],
    ["telegram_matches.json", { path: FILES.latestJson, displayName: "telegram_matches.json" }],
    ["all.csv", { path: FILES.allCsv, displayName: "all.csv" }],
    ["all.json", { path: FILES.allJson, displayName: "all.json" }],
    ["telegram_matches_all.csv", { path: FILES.allCsv, displayName: "telegram_matches_all.csv" }],
    ["telegram_matches_all.json", { path: FILES.allJson, displayName: "telegram_matches_all.json" }],
    ["retry_rows.csv", { path: FILES.retryCsv, displayName: "retry_rows.csv" }],
    ["job_state.json", { path: FILES.currentJob, displayName: "job_state.json" }],
    ["clean_ready.csv", { path: FILES.cleanReadyCsv, displayName: "clean_ready.csv" }],
    ["invalid_rows.csv", { path: FILES.invalidRowsCsv, displayName: "invalid_rows.csv" }],
    ["duplicate_phones.csv", { path: FILES.duplicatePhonesCsv, displayName: "duplicate_phones.csv" }],
    ["clean_debug.csv", { path: FILES.cleanDebugCsv, displayName: "clean_debug.csv" }],
    ["clean_rejects.csv", { path: FILES.cleanRejectsCsv, displayName: "clean_rejects.csv" }],
    ["summary.json", { path: FILES.intakeSummaryJson, displayName: "summary.json" }],
    ["run_log.json", { path: FILES.runLog, displayName: "run_log.json" }],
    ["run-log", { path: FILES.runLog, displayName: "run_log.json" }],
    ["processed.csv", { path: FILES.processedCsv, displayName: "processed_only.csv" }],
    ["processed.xlsx", { path: FILES.processedXlsx, displayName: "processed_only.xlsx" }],
    ["yes_only.csv", { path: FILES.yesOnlyCsv, displayName: "telegram_yes_only.csv" }],
    ["yes.csv", { path: FILES.yesOnlyCsv, displayName: "yes.csv" }],
    ["yes_only.json", { path: FILES.yesOnlyJson, displayName: "telegram_yes_only.json" }],
    ["no_only.csv", { path: FILES.noOnlyCsv, displayName: "telegram_no_only.csv" }],
    ["no.csv", { path: FILES.noOnlyCsv, displayName: "no.csv" }],
    ["no_only.json", { path: FILES.noOnlyJson, displayName: "telegram_no_only.json" }],
    ["retry_only.csv", { path: FILES.retryOnlyCsv, displayName: "telegram_retry_only.csv" }],
    ["retry.csv", { path: FILES.retryOnlyCsv, displayName: "retry.csv" }],
    ["retry_only.json", { path: FILES.retryOnlyJson, displayName: "telegram_retry_only.json" }],
    ["invalid.csv", { path: FILES.invalidOnlyCsv, displayName: "invalid.csv" }],
    ["invalid.json", { path: FILES.invalidOnlyJson, displayName: "invalid.json" }],
    ["marketing_allowed.csv", { path: FILES.marketingAllowedCsv, displayName: "marketing_allowed_only.csv" }],
    ["marketing_allowed.json", { path: FILES.marketingAllowedJson, displayName: "marketing_allowed_only.json" }],
    ["winback.csv", { path: FILES.winbackCsv, displayName: "winback_only.csv" }],
    ["winback.json", { path: FILES.winbackJson, displayName: "winback_only.json" }],
    ["retry_later.csv", { path: FILES.retryLaterCsv, displayName: "retry_later_only.csv" }],
    ["retry_later.json", { path: FILES.retryLaterJson, displayName: "retry_later_only.json" }],
    ["remaining.csv", { path: FILES.remainingCsv, displayName: "remaining_only.csv" }],
    ["remaining.xlsx", { path: FILES.remainingXlsx, displayName: "remaining_only.xlsx" }],
    ["remaining_only.csv", { path: FILES.remainingCsv, displayName: "remaining_only.csv" }],
    ["remaining_only.xlsx", { path: FILES.remainingXlsx, displayName: "remaining_only.xlsx" }],
  ]);
  return allowed.get(String(name || ""));
}

app.use(express.json({ limit: "2mb" }));
app.use(express.urlencoded({ extended: true }));

// Public healthcheck for Railway/infra probes. Must never require auth.
app.get("/api/health", (_req, res) => {
  res.status(200).json({ ok: true });
});

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

function systemScope(req, res, next) {
  const systemId = normalizeSystemId(req.params?.systemId || req.body?.systemId || req.query?.systemId);
  if (!systemId) return res.status(400).json({ error: "systemId ไม่ถูกต้อง ต้องเป็น A, B, C, D หรือ E" });
  req.systemId = systemId;
  return systemContext.run({ systemId }, next);
}

app.use("/api/systems/:systemId", systemScope);

app.get("/login", (_req, res) => res.sendFile(path.join(PUBLIC_DIR, "login.html")));
app.use(express.static(path.join(__dirname, "public")));
app.get("/", (_req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});
app.get("/index.html", (_req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

async function buildSystemSummary(systemId) {
  return withSystem(systemId, async () => {
    const [job, accountsState, settings] = await Promise.all([
      getCurrentJob(),
      loadAccountsState(),
      loadSettings(),
    ]);
    const selectedId = accountsState?.appState?.selectedAccountId || "";
    const selected = (accountsState?.accounts || []).find((item) => item.id === selectedId);
    return {
      id: systemId,
      label: `ระบบ ${systemId}`,
      status: displayQueueStatus(job?.status || "idle"),
      autoStatus: displayAutoStatus(job?.autoStatus || (settings.autoRun ? "RUNNING" : "OFF")),
      autoRun: Boolean(settings.autoRun),
      account: selected?.label || job?.lockedAccountLabel || "",
      accountPhone: selected?.phone || job?.lockedAccountPhone || "",
      processed: Number(job?.processedRows || 0),
      remaining: Number(job?.remainingRows || 0),
      yes: Number(job?.matchedCountTotal || 0),
      no: Number(job?.unmatchedCountTotal || 0),
      retry: Number(job?.retryCountTotal || 0),
      invalid: Number(job?.invalidCountTotal || 0),
      jobId: job?.id || "",
      jobLabel: job?.jobLabel || "",
      processing: isSystemProcessing(systemId),
    };
  });
}

app.get("/api/systems", async (_req, res) => {
  try {
    const systems = await Promise.all(SYSTEM_IDS.map((systemId) => buildSystemSummary(systemId)));
    res.json({ systems });
  } catch (error) {
    res.status(500).json({ error: error.message || "โหลดระบบไม่สำเร็จ" });
  }
});

app.get("/api/settings", async (_req, res) => {
  res.json(await loadSettings());
});

app.post("/api/settings", async (req, res) => {
  try {
    const settings = await saveSettings(req.body || {});
    const job = await pauseHighRetryForManual(await getCurrentJob(), "settings_update");
    const effectiveSettings = isHighRetryManualPause(job) ? await saveSettings({ autoRun: false }) : settings;
    let savedJob = job;
    if (job && job.status !== "completed") {
      const nextJob = {
        ...job,
        perRunLimit: effectiveSettings.maxContactsPerRun,
        chunkSize: effectiveSettings.batchSize,
        delayBetweenRunsSec: effectiveSettings.delayBetweenRunsSec,
        retryPauseSec: effectiveSettings.retryPauseSec,
        retryRatioThreshold: effectiveSettings.retryRatioThreshold,
        waitFloodAutomatically: effectiveSettings.waitFloodAutomatically,
        autoStatus: isHighRetryManualPause(job) ? "OFF" : (effectiveSettings.autoRun ? (job.autoStatus === "OFF" ? "RUNNING" : job.autoStatus) : "OFF"),
        status: isHighRetryManualPause(job) ? "paused_too_many_retry" : job.status,
        nextRunAt: isHighRetryManualPause(job) ? "" : job.nextRunAt,
        updatedAt: nowIso(),
      };
      savedJob = await saveCurrentJob(nextJob);
    }
    res.json({ ok: true, settings: effectiveSettings, job: savedJob });
  } catch (error) {
    res.status(400).json({ error: error.message || "บันทึกค่าไม่สำเร็จ" });
  }
});

app.get("/api/auto-status", async (_req, res) => {
  const job = await pauseHighRetryForManual(await markFloodStateStaleIfNeeded(await getCurrentJob()), "auto_status");
  const settings = await loadSettings();
  res.json({
    autoRun: settings.autoRun,
    status: job?.autoStatus || "OFF",
    queueStatus: displayQueueStatus(job?.status || "idle"),
    nextRunAt: job?.nextRunAt || "",
    lockedAccountLabel: job?.lockedAccountLabel || "",
    lockedAccountPhone: job?.lockedAccountPhone || "",
    file: job?.sourceFile || "",
    fileName: job?.fileName || job?.sourceFile || "",
    jobLabel: job?.jobLabel || "",
    jobNote: job?.jobNote || "",
    floodWaitSec: job?.lastFloodWaitSec || 0,
    floodWaitCount: Number(job?.floodWaitCount || 0),
    retryRatio: job?.lastRetryRatio || 0,
    retryRatioHigh: isHighRetryManualPause(job),
    manualRetryRequired: Boolean(job?.manualRetryRequired || isHighRetryManualPause(job)),
    processedRows: job?.processedRows || 0,
    remainingRows: job?.remainingRows || 0,
    requiresFloodReset: shouldRequireFloodReset(job),
  });
});

app.get("/api/dashboard", async (_req, res) => {
  await pauseHighRetryForManual(await markFloodStateStaleIfNeeded(await getCurrentJob()), "dashboard");
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

    pendingAuthClients.set(clientMapKey(accountId), client);
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

    const client = pendingAuthClients.get(clientMapKey(accountId));
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
    activeClients.set(clientMapKey(accountId), client);
    pendingAuthClients.delete(clientMapKey(accountId));

    await updateAccountById(accountId, (current) => ({
      ...current,
      sessionEnc: encryptText(sessionString),
      status: "ready",
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

    const client = pendingAuthClients.get(clientMapKey(accountId));
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
    activeClients.set(clientMapKey(accountId), client);
    pendingAuthClients.delete(clientMapKey(accountId));

    await updateAccountById(accountId, (current) => ({
      ...current,
      sessionEnc: encryptText(sessionString),
      status: "ready",
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
    const displayName = decodeUploadFileName(req.file.originalname);
    const masterFile = await persistMasterUpload(req.file.path, displayName);
    const { meta } = await processIntakeFile(req.file.path, displayName, {
      jobLabel: req.body?.jobLabel,
      jobNote: req.body?.jobNote,
      masterFile,
    });
    const cleanReadyState = await buildCleanReadyState(currentSystemId(), meta);
    res.json({
      ok: true,
      message: "ล้างไฟล์เรียบร้อย",
      ...cleanReadyState,
      summary: { ...meta, ...cleanReadyState },
      preview: meta.preview || [],
    });
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
    const masterFile = await persistMasterUpload(resolved, path.basename(resolved));
    const { meta } = await processIntakeFile(resolved, path.basename(resolved), {
      jobLabel: req.body?.jobLabel,
      jobNote: req.body?.jobNote,
      masterFile,
    });
    const cleanReadyState = await buildCleanReadyState(currentSystemId(), meta);
    res.json({
      ok: true,
      message: "นำเข้าไฟล์จาก path สำเร็จ",
      source: resolved,
      ...cleanReadyState,
      summary: { ...meta, ...cleanReadyState },
      preview: meta.preview || [],
    });
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
  res.json({ ...latest, ...(await buildCleanReadyState(currentSystemId(), latest)) });
});

app.post("/api/jobs/create-from-clean", async (_req, res) => {
  try {
    const job = await createJobFromLatestClean();
    const settings = await loadSettings();
    res.json({ ok: true, message: "สร้างคิวจาก clean_ready.csv แล้ว", job, settings });
  } catch (error) {
    res.status(400).json({
      error: error.message || "สร้างคิวไม่สำเร็จ",
      ...(error?.details || {}),
    });
  }
});

app.get("/api/job-status", async (_req, res) => {
  const job = await pauseHighRetryForManual(await markFloodStateStaleIfNeeded(await getCurrentJob()), "job_status");
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
  if (isSystemProcessing()) return res.status(409).json({ error: "ระบบนี้กำลังประมวลผลอยู่" });
  setSystemProcessing(true);
  try {
    const staleChecked = await pauseHighRetryForManual(await markFloodStateStaleIfNeeded(await getCurrentJob()), "manual_run");
    if (staleChecked?.status === "paused_flood_stale") {
      return res.status(409).json({ error: "คิวถูกพักเพราะ FloodWait ค้างนาน กรุณารีเซ็ต flood state ก่อน" });
    }
    await assertRunNextAuthReady(staleChecked);
    const result = await processNextBatch();
    res.json({ ok: true, message: "ทำรอบถัดไปเรียบร้อย", ...result });
  } catch (error) {
    const job = await getCurrentJob();
    if (job && (error?.isAccountAuthError || isAuthSessionError(error))) {
      const message = error?.code === "AUTH_REQUIRED" || error?.type === "AUTH_REQUIRED"
        ? (error.message || authRequiredMessage(error))
        : authRequiredMessage(error);
      const nextJob = await markJobAuthRequired(job, message);
      await appendRunLog({ type: "account_auth_invalid", jobId: job.id, message });
      return res.status(409).json({
        ok: false,
        type: "AUTH_REQUIRED",
        code: "AUTH_REQUIRED",
        message,
        error: message,
        job: nextJob,
      });
    }
    if (job && parseFloodSeconds(error?.message)) {
      const nextJob = await handleFloodWait(job, error);
      return res.status(429).json({ error: extractTelegramError(error), job: nextJob });
    }
    const friendly = extractTelegramError(error) || "ทำรอบถัดไปไม่สำเร็จ";
    const statusCode = /ยังไม่ได้เลือกบัญชี|ไม่พบบัญชี|ไม่ตรงกับบัญชี/.test(friendly) ? 409 : 400;
    res.status(statusCode).json({ error: friendly });
  } finally {
    setSystemProcessing(false);
  }
});

app.post("/api/job/start-auto", async (_req, res) => {
  const job = await pauseHighRetryForManual(await markFloodStateStaleIfNeeded(await getCurrentJob()), "start_auto");
  if (!job) return res.status(404).json({ error: "ยังไม่มีคิวงาน" });
  if (job.status === "paused_flood_stale") {
    return res.status(409).json({ error: "คิวถูกพักเพราะ FloodWait ค้างนาน กรุณารีเซ็ต flood state ก่อน" });
  }
  if (isHighRetryManualPause(job)) {
    return res.status(409).json({
      error: "retry ratio ล่าสุดสูงมาก Telegram กำลังให้พักบัญชี ปิด Auto แล้ว กรุณากดทำรอบถัดไปเองเมื่อพร้อม",
      code: "PAUSED_TOO_MANY_RETRY",
      job,
    });
  }
  const settings = await saveSettings({ autoRun: true });
  const now = Date.now();
  const waitFloodTs = job.nextRunAt ? new Date(job.nextRunAt).getTime() : 0;
  const waitRetryTs = job.nextRunAt ? new Date(job.nextRunAt).getTime() : 0;
  const keepFloodWaiting = job.status === "waiting_flood" && waitFloodTs && now < waitFloodTs;
  const keepRetryWaiting = (job.status === "waiting_retry_cooldown" || job.status === "paused_retry_cooldown") && waitRetryTs && now < waitRetryTs;
  const nextStatus = keepFloodWaiting
    ? "waiting_flood"
    : keepRetryWaiting
      ? "paused_retry_cooldown"
      : (["paused", "ready", "waiting_retry_cooldown", "paused_retry_cooldown", "waiting_flood"].includes(job.status) ? "ready" : job.status);
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
  const job = await pauseHighRetryForManual(await markFloodStateStaleIfNeeded(await getCurrentJob()), "resume_auto");
  if (!job) return res.status(404).json({ error: "ยังไม่มีคิวงาน" });
  if (job.status === "paused_flood_stale") {
    return res.status(409).json({ error: "คิวถูกพักเพราะ FloodWait ค้างนาน กรุณารีเซ็ต flood state ก่อน" });
  }
  if (isHighRetryManualPause(job)) {
    return res.status(409).json({
      error: "retry ratio ล่าสุดสูงมาก Telegram กำลังให้พักบัญชี ปิด Auto แล้ว กรุณากดทำรอบถัดไปเองเมื่อพร้อม",
      code: "PAUSED_TOO_MANY_RETRY",
      job,
    });
  }
  const settings = await saveSettings({ autoRun: true });
  const now = Date.now();
  const waitFloodTs = job.nextRunAt ? new Date(job.nextRunAt).getTime() : 0;
  const waitRetryTs = job.nextRunAt ? new Date(job.nextRunAt).getTime() : 0;
  const keepFloodWaiting = job.status === "waiting_flood" && waitFloodTs && now < waitFloodTs;
  const keepRetryWaiting = (job.status === "waiting_retry_cooldown" || job.status === "paused_retry_cooldown") && waitRetryTs && now < waitRetryTs;
  const nextJob = {
    ...job,
    autoStatus: keepFloodWaiting ? "WAITING_FLOOD" : (keepRetryWaiting ? "WAITING_RETRY" : "RUNNING"),
    status: keepFloodWaiting ? "waiting_flood" : (keepRetryWaiting ? "paused_retry_cooldown" : "ready"),
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
  if (isSystemProcessing()) return res.status(409).json({ error: "ระบบนี้กำลังประมวลผลอยู่ ล้างคิวตอนนี้ไม่ได้" });
  try {
    for (const filePath of [
      FILES.currentJob,
      FILES.jobState,
      FILES.latestJson,
      FILES.latestCsv,
      FILES.allJson,
      FILES.allCsv,
      FILES.processedCsv,
      FILES.processedXlsx,
      FILES.retryCsv,
      FILES.yesOnlyCsv,
      FILES.yesOnlyJson,
      FILES.noOnlyCsv,
      FILES.noOnlyJson,
      FILES.retryOnlyCsv,
      FILES.retryOnlyJson,
      FILES.invalidOnlyCsv,
      FILES.invalidOnlyJson,
      FILES.marketingAllowedCsv,
      FILES.marketingAllowedJson,
      FILES.winbackCsv,
      FILES.winbackJson,
      FILES.retryLaterCsv,
      FILES.retryLaterJson,
      FILES.autoState,
      FILES.remainingCsv,
      FILES.remainingXlsx,
      FILES.cleanReadyCsv,
      FILES.invalidRowsCsv,
      FILES.duplicatePhonesCsv,
      FILES.cleanDebugCsv,
      FILES.cleanRejectsCsv,
      FILES.intakeSummaryJson,
      FILES.intakeLatest,
    ]) await unlinkIfExists(filePath);
    res.json({ ok: true, message: "ล้างคิวและผลลัพธ์เรียบร้อย" });
  } catch (error) {
    res.status(500).json({ error: error.message || "ล้างคิวไม่สำเร็จ" });
  }
});

app.get("/api/systems/:systemId/status", async (_req, res) => {
  await pauseHighRetryForManual(await markFloodStateStaleIfNeeded(await getCurrentJob()), "system_status");
  res.json(await buildDashboardPayload());
});

app.get("/api/systems/:systemId/dashboard", async (_req, res) => {
  await pauseHighRetryForManual(await markFloodStateStaleIfNeeded(await getCurrentJob()), "dashboard");
  res.json(await buildDashboardPayload());
});

app.get("/api/systems/:systemId/settings", async (_req, res) => {
  res.json(await loadSettings());
});

app.post("/api/systems/:systemId/settings", async (req, res) => {
  try {
    const settings = await saveSettings(req.body || {});
    const job = await pauseHighRetryForManual(await getCurrentJob(), "settings_update");
    const effectiveSettings = isHighRetryManualPause(job) ? await saveSettings({ autoRun: false }) : settings;
    let savedJob = job;
    if (job && job.status !== "completed") {
      const nextJob = {
        ...job,
        perRunLimit: effectiveSettings.maxContactsPerRun,
        chunkSize: effectiveSettings.batchSize,
        delayBetweenRunsSec: effectiveSettings.delayBetweenRunsSec,
        retryPauseSec: effectiveSettings.retryPauseSec,
        retryRatioThreshold: effectiveSettings.retryRatioThreshold,
        waitFloodAutomatically: effectiveSettings.waitFloodAutomatically,
        autoStatus: isHighRetryManualPause(job) ? "OFF" : (effectiveSettings.autoRun ? (job.autoStatus === "OFF" ? "RUNNING" : job.autoStatus) : "OFF"),
        status: isHighRetryManualPause(job) ? "paused_too_many_retry" : job.status,
        nextRunAt: isHighRetryManualPause(job) ? "" : job.nextRunAt,
        updatedAt: nowIso(),
      };
      savedJob = await saveCurrentJob(nextJob);
    }
    res.json({ ok: true, settings: effectiveSettings, job: savedJob });
  } catch (error) {
    res.status(400).json({ error: error.message || "บันทึกค่าไม่สำเร็จ" });
  }
});

app.get("/api/systems/:systemId/auto-status", async (_req, res) => {
  const job = await pauseHighRetryForManual(await markFloodStateStaleIfNeeded(await getCurrentJob()), "auto_status");
  const settings = await loadSettings();
  res.json({
    autoRun: settings.autoRun,
    status: job?.autoStatus || "OFF",
    queueStatus: displayQueueStatus(job?.status || "idle"),
    nextRunAt: job?.nextRunAt || "",
    lockedAccountLabel: job?.lockedAccountLabel || "",
    lockedAccountPhone: job?.lockedAccountPhone || "",
    file: job?.sourceFile || "",
    fileName: job?.fileName || job?.sourceFile || "",
    jobLabel: job?.jobLabel || "",
    jobNote: job?.jobNote || "",
    floodWaitSec: job?.lastFloodWaitSec || 0,
    floodWaitCount: Number(job?.floodWaitCount || 0),
    retryRatio: job?.lastRetryRatio || 0,
    retryRatioHigh: isHighRetryManualPause(job),
    manualRetryRequired: Boolean(job?.manualRetryRequired || isHighRetryManualPause(job)),
    processedRows: job?.processedRows || 0,
    remainingRows: job?.remainingRows || 0,
    requiresFloodReset: shouldRequireFloodReset(job),
  });
});

app.get("/api/systems/:systemId/accounts", async (_req, res) => {
  try {
    res.json({ accounts: await getAccounts() });
  } catch (error) {
    res.status(500).json({ error: error.message || "โหลดบัญชีไม่สำเร็จ" });
  }
});

app.post("/api/systems/:systemId/accounts", async (req, res) => {
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
    if (duplicate) throw new Error("มีบัญชีชื่อนี้หรือเบอร์นี้อยู่แล้วในระบบนี้");
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

app.post("/api/systems/:systemId/accounts/:id/send-code", async (req, res) => {
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
    pendingAuthClients.set(clientMapKey(accountId), client);
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
    res.json({ ok: true, message: "ส่ง OTP แล้ว กรุณาไปกรอกรหัสในหน้าเว็บ", accounts: await getAccounts() });
  } catch (error) {
    await disconnectClientMap(pendingAuthClients, accountId);
    const nextStatus = isAuthSessionError(error) ? "auth_invalid" : "error";
    try {
      await updateAccountById(accountId, (current) => ({ ...current, status: nextStatus, lastError: extractTelegramError(error), updatedAt: nowIso() }));
    } catch {}
    res.status(400).json({ error: extractTelegramError(error) });
  }
});

app.post("/api/systems/:systemId/accounts/:id/verify-code", async (req, res) => {
  const accountId = String(req.params.id || "");
  try {
    const code = String(req.body.code || "").trim();
    if (!code) throw new Error("กรุณาใส่รหัส OTP");
    const { account } = await getAccountById(accountId);
    if (!account) throw new Error("ไม่พบบัญชี");
    const client = pendingAuthClients.get(clientMapKey(accountId));
    if (!client) throw new Error("รอบ OTP นี้หมดแล้ว กรุณากดส่ง OTP ใหม่");
    if (!account.pendingPhoneCodeHash) throw new Error("ไม่พบ phoneCodeHash กรุณากดส่ง OTP ใหม่");
    try {
      await client.invoke(new Api.auth.SignIn({ phoneNumber: account.phone, phoneCodeHash: account.pendingPhoneCodeHash, phoneCode: code }));
    } catch (error) {
      if (String(error?.message || "").toUpperCase().includes("SESSION_PASSWORD_NEEDED")) {
        await updateAccountById(accountId, (current) => ({ ...current, status: "awaiting_password", awaitingPassword: true, lastError: "", updatedAt: nowIso() }));
        return res.json({ ok: true, needPassword: true, message: "บัญชีนี้เปิด 2FA อยู่ กรุณาใส่รหัสผ่าน 2FA", accounts: await getAccounts() });
      }
      throw error;
    }
    const me = await client.getMe();
    const sessionString = client.session.save();
    await writeSessionToFile(accountId, sessionString);
    activeClients.set(clientMapKey(accountId), client);
    pendingAuthClients.delete(clientMapKey(accountId));
    await updateAccountById(accountId, (current) => ({
      ...current,
      sessionEnc: encryptText(sessionString),
      status: "ready",
      pendingPhoneCodeHash: "",
      pendingCodeType: "",
      pendingTimeout: 0,
      awaitingPassword: false,
      lastError: "",
      me: me ? { id: toSafeStringId(me.id), username: me.username || "", firstName: me.firstName || "", lastName: me.lastName || "", phone: me.phone ? `+${onlyDigits(me.phone)}` : "" } : null,
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
      await updateAccountById(accountId, (current) => ({ ...current, status: nextStatus, lastError: extractTelegramError(error), updatedAt: nowIso() }));
    } catch {}
    res.status(400).json({ error: extractTelegramError(error) });
  }
});

app.post("/api/systems/:systemId/accounts/:id/verify-password", async (req, res) => {
  const accountId = String(req.params.id || "");
  try {
    const password = String(req.body.password || "").trim();
    if (!password) throw new Error("กรุณาใส่รหัส 2FA");
    const { account } = await getAccountById(accountId);
    if (!account) throw new Error("ไม่พบบัญชี");
    const client = pendingAuthClients.get(clientMapKey(accountId));
    if (!client) throw new Error("รอบล็อกอินนี้หมดแล้ว กรุณากดส่ง OTP ใหม่");
    let apiHash = "";
    try {
      apiHash = decryptText(account.apiHashEnc || "").trim();
    } catch {
      throw makeAccountAuthError("API_HASH ของบัญชีนี้ถอดรหัสไม่ได้ กรุณาแก้บัญชีแล้วล็อกอินใหม่", "auth_invalid");
    }
    if (!apiHash) throw makeAccountAuthError("API_HASH ของบัญชีนี้ไม่ครบ กรุณาแก้บัญชีแล้วล็อกอินใหม่", "auth_invalid");
    await client.signInWithPassword({ apiId: Number(account.apiId), apiHash }, { password: async () => password, onError: (err) => { throw err; } });
    const me = await client.getMe();
    const sessionString = client.session.save();
    await writeSessionToFile(accountId, sessionString);
    activeClients.set(clientMapKey(accountId), client);
    pendingAuthClients.delete(clientMapKey(accountId));
    await updateAccountById(accountId, (current) => ({
      ...current,
      sessionEnc: encryptText(sessionString),
      status: "ready",
      pendingPhoneCodeHash: "",
      pendingCodeType: "",
      pendingTimeout: 0,
      awaitingPassword: false,
      lastError: "",
      me: me ? { id: toSafeStringId(me.id), username: me.username || "", firstName: me.firstName || "", lastName: me.lastName || "", phone: me.phone ? `+${onlyDigits(me.phone)}` : "" } : null,
      updatedAt: nowIso(),
    }));
    res.json({ ok: true, message: "ยืนยัน 2FA สำเร็จแล้ว", accounts: await getAccounts() });
  } catch (error) {
    const nextStatus = isAuthSessionError(error) ? "needs_relogin" : "awaiting_password";
    try {
      await updateAccountById(accountId, (current) => ({ ...current, status: nextStatus, awaitingPassword: nextStatus === "awaiting_password", lastError: extractTelegramError(error), updatedAt: nowIso() }));
    } catch {}
    res.status(400).json({ error: extractTelegramError(error) });
  }
});

app.post("/api/systems/:systemId/accounts/:id/select", async (req, res) => {
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

app.post("/api/systems/:systemId/accounts/:id/reset-session", async (req, res) => {
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

app.delete("/api/systems/:systemId/accounts/:id", async (req, res) => {
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

app.post("/api/systems/:systemId/upload", upload.single("file"), async (req, res) => {
  try {
    const systemId = normalizeSystemId(req.params.systemId);
    if (!systemId) throw new Error("systemId ไม่ถูกต้อง ต้องเป็น A, B, C, D หรือ E");
    if (!req.file) throw new Error("กรุณาอัปโหลดไฟล์ก่อน");
    logInfo(`SYSTEM UPLOAD REQUEST systemId=${systemId}`);
    const response = await withSystem(systemId, async () => {
      const displayName = decodeUploadFileName(req.file.originalname);
      const masterFile = await persistMasterUpload(req.file.path, displayName);
      const { meta } = await processIntakeFile(req.file.path, displayName, {
        jobLabel: req.body?.jobLabel,
        jobNote: req.body?.jobNote,
        masterFile,
      });
      const cleanReadyState = await buildCleanReadyState(systemId, meta);
      return {
        ok: true,
        message: "ล้างไฟล์เรียบร้อย",
        ...cleanReadyState,
        summary: { ...meta, ...cleanReadyState },
        preview: meta.preview || [],
      };
    });
    logInfo(`SYSTEM UPLOAD RESPONSE systemId=${response.systemId} cleanReadyPath=${response.cleanReadyPath}`);
    res.json(response);
  } catch (error) {
    res.status(400).json({ error: error.message || "ล้างไฟล์ไม่สำเร็จ" });
  }
});

app.post("/api/systems/:systemId/intake/upload", upload.single("file"), async (req, res) => {
  try {
    const systemId = normalizeSystemId(req.params.systemId);
    if (!systemId) throw new Error("systemId ไม่ถูกต้อง ต้องเป็น A, B, C, D หรือ E");
    if (!req.file) throw new Error("กรุณาอัปโหลดไฟล์ก่อน");
    logInfo(`SYSTEM INTAKE UPLOAD REQUEST systemId=${systemId}`);
    const response = await withSystem(systemId, async () => {
      const displayName = decodeUploadFileName(req.file.originalname);
      const masterFile = await persistMasterUpload(req.file.path, displayName);
      const { meta } = await processIntakeFile(req.file.path, displayName, {
        jobLabel: req.body?.jobLabel,
        jobNote: req.body?.jobNote,
        masterFile,
      });
      const cleanReadyState = await buildCleanReadyState(systemId, meta);
      return {
        ok: true,
        message: "ล้างไฟล์เรียบร้อย",
        ...cleanReadyState,
        summary: { ...meta, ...cleanReadyState },
        preview: meta.preview || [],
      };
    });
    logInfo(`SYSTEM INTAKE UPLOAD RESPONSE systemId=${response.systemId} cleanReadyPath=${response.cleanReadyPath}`);
    res.json(response);
  } catch (error) {
    res.status(400).json({ error: error.message || "ล้างไฟล์ไม่สำเร็จ" });
  }
});

app.get("/api/systems/:systemId/intake/latest", async (req, res) => {
  const systemId = normalizeSystemId(req.params.systemId);
  if (!systemId) return res.status(400).json({ error: "systemId ไม่ถูกต้อง ต้องเป็น A, B, C, D หรือ E" });
  await withSystem(systemId, async () => {
    const latest = await readJson(FILES.intakeLatest, null);
    if (!latest) return res.status(404).json({ error: "ยังไม่มีผลล้างไฟล์ล่าสุด" });
    return res.json({ ...latest, ...(await buildCleanReadyState(systemId, latest)) });
  });
});

app.post("/api/systems/:systemId/job/create", async (req, res) => {
  try {
    const systemId = normalizeSystemId(req.params.systemId);
    if (!systemId) throw new Error("systemId ไม่ถูกต้อง ต้องเป็น A, B, C, D หรือ E");
    const { job, settings } = await withSystem(systemId, async () => ({
      job: await createJobFromLatestClean(systemId),
      settings: await loadSettings(),
    }));
    res.json({ ok: true, systemId, message: "สร้างคิวจาก clean_ready.csv แล้ว", job, settings });
  } catch (error) {
    res.status(400).json({
      error: error.message || "สร้างคิวไม่สำเร็จ",
      ...(error?.details || {}),
    });
  }
});

app.get("/api/systems/:systemId/diagnostics/import-log", async (req, res) => {
  const systemId = normalizeSystemId(req.params.systemId);
  if (!systemId) return res.status(400).json({ error: "systemId ไม่ถูกต้อง ต้องเป็น A, B, C, D หรือ E" });
  await withSystem(systemId, async () => {
    const rows = await readImportDiagnosticLog(req.query?.limit || 50);
    res.json({
      ok: true,
      systemId,
      count: rows.length,
      rows,
    });
  });
});

app.get("/api/systems/:systemId/diagnostics/import-summary", async (req, res) => {
  const systemId = normalizeSystemId(req.params.systemId);
  if (!systemId) return res.status(400).json({ error: "systemId ไม่ถูกต้อง ต้องเป็น A, B, C, D หรือ E" });
  await withSystem(systemId, async () => {
    const rows = await readImportDiagnosticLog("all");
    res.json({
      ok: true,
      systemId,
      ...buildImportDiagnosticSummary(rows),
    });
  });
});

app.post("/api/systems/:systemId/job/run-next", async (_req, res) => {
  await runNextForCurrentSystem(res);
});

app.post("/api/systems/:systemId/job/start-auto", async (_req, res) => {
  try {
    await startAutoForCurrentSystem(res);
  } catch (error) {
    res.status(409).json({ error: error.message || "เริ่ม Auto ไม่สำเร็จ" });
  }
});

app.post("/api/systems/:systemId/job/pause-auto", async (_req, res) => {
  await pauseAutoForCurrentSystem(res);
});

app.post("/api/systems/:systemId/job/resume-auto", async (_req, res) => {
  try {
    await resumeAutoForCurrentSystem(res);
  } catch (error) {
    res.status(409).json({ error: error.message || "ทำต่อ Auto ไม่สำเร็จ" });
  }
});

app.post("/api/systems/:systemId/job/resume-manual", async (_req, res) => {
  try {
    await resumeManualForCurrentSystem(res);
  } catch (error) {
    res.status(409).json({ error: error.message || "ปลดพักคิวสำหรับตรวจมือไม่สำเร็จ" });
  }
});

app.post("/api/systems/:systemId/job/reset-flood-state", async (_req, res) => {
  try {
    const result = await resetFloodStateAndClearStaleJob();
    res.json({ ok: true, ...result });
  } catch (error) {
    if (error.message === "ยังไม่มีคิวงาน") return res.status(404).json({ error: error.message });
    if (error.message === "สถานะคิวนี้ไม่ใช่ flood state ที่รีเซ็ตได้") return res.status(400).json({ error: error.message });
    res.status(500).json({ error: error.message || "รีเซ็ต flood state ไม่สำเร็จ" });
  }
});

app.post("/api/systems/:systemId/job/reset-flood-stale", async (_req, res) => {
  try {
    const result = await resetFloodStateAndClearStaleJob();
    res.json({ ok: true, ...result });
  } catch (error) {
    if (error.message === "ยังไม่มีคิวงาน") return res.status(404).json({ error: error.message });
    if (error.message === "สถานะคิวนี้ไม่ใช่ flood state ที่รีเซ็ตได้") return res.status(400).json({ error: error.message });
    res.status(500).json({ error: error.message || "รีเซ็ต flood state ไม่สำเร็จ" });
  }
});

app.post("/api/systems/:systemId/job/reset", async (_req, res) => {
  await resetJobForCurrentSystem(res);
});

app.get("/api/systems/:systemId/job/status", async (_req, res) => {
  const job = await pauseHighRetryForManual(await markFloodStateStaleIfNeeded(await getCurrentJob()), "job_status");
  if (!job) return res.status(404).json({ error: "ยังไม่มีคิวงาน" });
  res.json(job);
});

app.get("/api/systems/:systemId/latest", async (_req, res) => {
  const latest = await readJson(FILES.latestJson, null);
  if (!latest) return res.status(404).json({ error: "ยังไม่มีผลล่าสุด" });
  res.json(latest);
});

app.get("/api/systems/:systemId/remaining", async (_req, res) => {
  const job = await getCurrentJob();
  if (!job || !Array.isArray(job.rows)) return res.status(404).json({ error: "ยังไม่มีคิวงาน" });
  const doneSet = await getProcessedSourceIndexSet();
  const remainingRows = job.rows.filter((row) => !doneSet.has(String(row?.sourceIndex || "")));
  res.json({
    ok: true,
    file: "remaining_only.csv",
    count: remainingRows.length,
    totalRows: job.totalRows || 0,
    processedRows: job.processedRows || 0,
    rows: remainingRows.slice(0, 30),
  });
});

app.get("/api/systems/:systemId/download/:name", async (req, res) => {
  const target = getDownloadTarget(req.params.name);
  if (!target) return res.status(404).send("ไม่พบไฟล์");
  if (!fs.existsSync(target.path)) return res.status(404).send("ยังไม่มีไฟล์นี้");
  sendFileDownload(res, target.path, target.displayName);
});

app.get("/download/:name", async (req, res) => {
  const allowed = new Map([
    ["telegram_matches.csv", { path: FILES.latestCsv, displayName: "telegram_matches.csv" }],
    ["telegram_matches.json", { path: FILES.latestJson, displayName: "telegram_matches.json" }],
    ["all.csv", { path: FILES.allCsv, displayName: "all.csv" }],
    ["all.json", { path: FILES.allJson, displayName: "all.json" }],
    ["telegram_matches_all.csv", { path: FILES.allCsv, displayName: "telegram_matches_all.csv" }],
    ["telegram_matches_all.json", { path: FILES.allJson, displayName: "telegram_matches_all.json" }],
    ["retry_rows.csv", { path: FILES.retryCsv, displayName: "retry_rows.csv" }],
    ["job_state.json", { path: FILES.currentJob, displayName: "job_state.json" }],
    ["clean_ready.csv", { path: FILES.cleanReadyCsv, displayName: "clean_ready.csv" }],
    ["invalid_rows.csv", { path: FILES.invalidRowsCsv, displayName: "invalid_rows.csv" }],
    ["duplicate_phones.csv", { path: FILES.duplicatePhonesCsv, displayName: "duplicate_phones.csv" }],
    ["clean_debug.csv", { path: FILES.cleanDebugCsv, displayName: "clean_debug.csv" }],
    ["clean_rejects.csv", { path: FILES.cleanRejectsCsv, displayName: "clean_rejects.csv" }],
    ["summary.json", { path: FILES.intakeSummaryJson, displayName: "summary.json" }],
    ["run_log.json", { path: FILES.runLog, displayName: "run_log.json" }],
    ["run-log", { path: FILES.runLog, displayName: "run_log.json" }],
    ["processed.csv", { path: FILES.processedCsv, displayName: "processed_only.csv" }],
    ["processed.xlsx", { path: FILES.processedXlsx, displayName: "processed_only.xlsx" }],
    ["yes_only.csv", { path: FILES.yesOnlyCsv, displayName: "telegram_yes_only.csv" }],
    ["yes.csv", { path: FILES.yesOnlyCsv, displayName: "yes.csv" }],
    ["yes_only.json", { path: FILES.yesOnlyJson, displayName: "telegram_yes_only.json" }],
    ["no_only.csv", { path: FILES.noOnlyCsv, displayName: "telegram_no_only.csv" }],
    ["no.csv", { path: FILES.noOnlyCsv, displayName: "no.csv" }],
    ["no_only.json", { path: FILES.noOnlyJson, displayName: "telegram_no_only.json" }],
    ["retry_only.csv", { path: FILES.retryOnlyCsv, displayName: "telegram_retry_only.csv" }],
    ["retry.csv", { path: FILES.retryOnlyCsv, displayName: "retry.csv" }],
    ["retry_only.json", { path: FILES.retryOnlyJson, displayName: "telegram_retry_only.json" }],
    ["invalid.csv", { path: FILES.invalidOnlyCsv, displayName: "invalid.csv" }],
    ["invalid.json", { path: FILES.invalidOnlyJson, displayName: "invalid.json" }],
    ["marketing_allowed.csv", { path: FILES.marketingAllowedCsv, displayName: "marketing_allowed_only.csv" }],
    ["marketing_allowed.json", { path: FILES.marketingAllowedJson, displayName: "marketing_allowed_only.json" }],
    ["winback.csv", { path: FILES.winbackCsv, displayName: "winback_only.csv" }],
    ["winback.json", { path: FILES.winbackJson, displayName: "winback_only.json" }],
    ["retry_later.csv", { path: FILES.retryLaterCsv, displayName: "retry_later_only.csv" }],
    ["retry_later.json", { path: FILES.retryLaterJson, displayName: "retry_later_only.json" }],
    ["remaining.csv", { path: FILES.remainingCsv, displayName: "remaining_only.csv" }],
    ["remaining.xlsx", { path: FILES.remainingXlsx, displayName: "remaining_only.xlsx" }],
    ["remaining_only.csv", { path: FILES.remainingCsv, displayName: "remaining_only.csv" }],
    ["remaining_only.xlsx", { path: FILES.remainingXlsx, displayName: "remaining_only.xlsx" }],
  ]);

  const requested = String(req.params.name || "");
  if (!allowed.has(requested)) return res.status(404).send("ไม่พบไฟล์");
  const target = allowed.get(requested);
  const filePath = target.path;
  if (!fs.existsSync(filePath)) return res.status(404).send("ยังไม่มีไฟล์นี้");
  sendFileDownload(res, filePath, target.displayName);
});

app.use((req, res) => {
  res.status(404).json({ error: "Route not found" });
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

async function autoTick(systemId = DEFAULT_SYSTEM_ID) {
  return withSystem(systemId, async () => {
  if (isSystemProcessing()) return;
  const settings = await loadSettings();
  const job = await pauseHighRetryForManual(await markFloodStateStaleIfNeeded(await getCurrentJob()), "auto_loop");
  if (!job || !settings.autoRun) return;
  if (job.status === "completed" || job.status === "paused" || job.status === "paused_flood_stale" || job.status === "paused_too_many_retry" || displayQueueStatus(job.status) === "AUTH_REQUIRED") return;

  const now = Date.now();
  const nextRunTime = job.nextRunAt ? new Date(job.nextRunAt).getTime() : 0;
  if (job.status === "waiting_flood" || job.status === "waiting_retry_cooldown" || job.status === "paused_retry_cooldown" || job.status === "ready") {
    if (nextRunTime && now < nextRunTime) return;
  }

  setSystemProcessing(true);
  try {
    const nextJobBefore = await getCurrentJob();
    if (!nextJobBefore) return;
    await assertRunNextAuthReady(nextJobBefore);
    const normalizedJob = {
      ...nextJobBefore,
      status: nextJobBefore.status === "waiting_flood" || nextJobBefore.status === "waiting_retry_cooldown" || nextJobBefore.status === "paused_retry_cooldown" ? "ready" : nextJobBefore.status,
      autoStatus: "RUNNING",
      updatedAt: nowIso(),
    };
    await saveCurrentJob(normalizedJob);
    await processNextBatch();
  } catch (error) {
    const current = await getCurrentJob();
    if (current && (error?.isAccountAuthError || isAuthSessionError(error))) {
      const message = error?.code === "AUTH_REQUIRED" || error?.type === "AUTH_REQUIRED"
        ? (error.message || authRequiredMessage(error))
        : authRequiredMessage(error);
      await markJobAuthRequired(current, message);
      pushRecoveryNotice('warning', 'AUTO_AUTH_INVALID', 'ต้องล็อกอินบัญชีใหม่', message);
      await appendRunLog({ type: "auto_auth_invalid", jobId: current.id, message });
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
    setSystemProcessing(false);
  }
  });
}

async function copyFileIfMissing(source, target) {
  if (!fs.existsSync(source) || fs.existsSync(target)) return;
  await fsp.mkdir(path.dirname(target), { recursive: true });
  await fsp.copyFile(source, target);
}

async function copyDirIfMissing(source, target) {
  if (!fs.existsSync(source) || fs.existsSync(target)) return;
  await fsp.mkdir(path.dirname(target), { recursive: true });
  await fsp.cp(source, target, { recursive: true, force: false, errorOnExist: false });
}

async function migrateLegacyDataToSystemA() {
  const p = getSystemPaths(DEFAULT_SYSTEM_ID);
  await copyFileIfMissing(path.join(DATA_DIR, "accounts.json"), p.files.accounts);
  await copyFileIfMissing(path.join(DATA_DIR, "app_state.json"), p.files.appState);
  await copyFileIfMissing(path.join(DATA_DIR, "settings.json"), p.files.settings);
  await copyDirIfMissing(LEGACY_DIRS.output, p.outputDir);
  await copyDirIfMissing(LEGACY_DIRS.session, p.sessionDir);
  await copyDirIfMissing(LEGACY_DIRS.logs, p.logDir);
  await copyDirIfMissing(LEGACY_DIRS.jobs, p.jobsDir);
  await copyDirIfMissing(LEGACY_DIRS.chunks, p.chunksDir);
  await copyDirIfMissing(LEGACY_DIRS.clean, p.cleanDir);
  await copyDirIfMissing(LEGACY_DIRS.raw, p.rawDir);
  await copyDirIfMissing(LEGACY_DIRS.uploads, p.uploadsDir);
}

async function initializeSystemState(systemId) {
  await withSystem(systemId, async () => {
    for (const filePath of Object.values(getSystemPaths(systemId).files)) {
      await seedLastGoodBackup(filePath);
    }
    if (!fs.existsSync(FILES.accounts)) await writeJson(FILES.accounts, []);
    if (!fs.existsSync(FILES.appState)) await writeJson(FILES.appState, { selectedAccountId: "" });
    if (!fs.existsSync(FILES.settings)) await writeJson(FILES.settings, DEFAULT_SETTINGS);
    if (!fs.existsSync(FILES.runLog)) await writeJson(FILES.runLog, []);
    if (!fs.existsSync(FILES.autoState)) await writeJson(FILES.autoState, await buildAutoState(null));
    await markFloodStateStaleIfNeeded(await getCurrentJob());
    await syncDerivedOutputFilesFromState().catch(() => {});
  });
}

(async () => {
  validateBootConfig();
  await ensureDirectories();
  await migrateLegacyDataToSystemA();
  for (const systemId of SYSTEM_IDS) {
    await initializeSystemState(systemId);
    const timer = setInterval(() => {
      autoTick(systemId).catch((err) => console.error(`[AUTO:${systemId}]`, err));
    }, 5000);
    autoTimers.set(systemId, timer);
  }

  app.listen(PORT, HOST, () => {
    logInfo(`Telegram All-in-One running on http://${HOST}:${PORT}`);
    logInfo(`DATA_DIR=${DATA_DIR}`);
    logInfo(`SYSTEMS_DIR=${SYSTEMS_DIR}`);
  });
})();
