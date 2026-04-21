const crypto = require("crypto");
const path = require("path");

const APP_USER = String(process.env.APP_USER || "admin").trim();
const APP_PASSWORD = String(process.env.APP_PASSWORD || "").trim();
const APP_SECRET = String(process.env.APP_SECRET || "").trim();
const AUTH_COOKIE = "tg_checker_auth";
const AUTH_TTL_MS = 1000 * 60 * 60 * 12; // 12h
const PUBLIC_DIR = path.join(__dirname, "public");

function parseCookies(req) {
  const header = req.headers.cookie || "";
  const out = {};
  header.split(";").forEach((part) => {
    const idx = part.indexOf("=");
    if (idx > -1) {
      const key = part.slice(0, idx).trim();
      const val = decodeURIComponent(part.slice(idx + 1).trim());
      out[key] = val;
    }
  });
  return out;
}

function signToken(payload) {
  const body = Buffer.from(JSON.stringify(payload)).toString("base64url");
  const sig = crypto.createHmac("sha256", APP_SECRET).update(body).digest("base64url");
  return `${body}.${sig}`;
}

function verifyToken(token) {
  try {
    if (!token || !token.includes(".")) return null;
    const [body, sig] = token.split(".");
    const expected = crypto.createHmac("sha256", APP_SECRET).update(body).digest("base64url");
    if (sig !== expected) return null;
    const payload = JSON.parse(Buffer.from(body, "base64url").toString("utf8"));
    if (!payload.exp || Date.now() > payload.exp) return null;
    return payload;
  } catch {
    return null;
  }
}

function setAuthCookie(res, token) {
  const isProd = process.env.NODE_ENV === "production";
  const cookie = [
    `${AUTH_COOKIE}=${encodeURIComponent(token)}`,
    "Path=/",
    "HttpOnly",
    "SameSite=Lax",
    `Max-Age=${Math.floor(AUTH_TTL_MS / 1000)}`,
  ];
  if (isProd) cookie.push("Secure");
  res.setHeader("Set-Cookie", cookie.join("; "));
}

function clearAuthCookie(res) {
  res.setHeader("Set-Cookie", `${AUTH_COOKIE}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0`);
}

function requireAuth(req, res, next) {
  const cookies = parseCookies(req);
  const payload = verifyToken(cookies[AUTH_COOKIE]);
  if (!payload) {
    return res.status(401).json({ ok: false, error: "unauthorized" });
  }
  req.auth = payload;
  next();
}

module.exports = {
  APP_USER,
  APP_PASSWORD,
  APP_SECRET,
  AUTH_COOKIE,
  AUTH_TTL_MS,
  PUBLIC_DIR,
  parseCookies,
  signToken,
  verifyToken,
  setAuthCookie,
  clearAuthCookie,
  requireAuth,
};
