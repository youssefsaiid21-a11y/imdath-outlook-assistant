// =============================================================================
// Imdath (Sphinx) - Express Server
// =============================================================================
// This is the backend for an Outlook add-in that helps users draft emails using
// AI (Claude), pull CRM context from HubSpot, and send/reply via Microsoft
// Graph API. It also supports auto-tracking emails to HubSpot deals via
// Graph webhook subscriptions.
//
// Key integrations:
//   - Microsoft Graph API: email read/send/reply, contact search, thread history
//   - HubSpot CRM: contacts, deals, pipelines, email logging
//   - Anthropic Claude: AI email generation, refinement, summarization
//   - Groq Whisper: voice-to-text transcription for voice note input
// =============================================================================

"use strict";

// Load environment variables from .env in the project root
require("dotenv").config({ path: require("path").join(__dirname, "../.env") });

const express  = require("express");
const cors     = require("cors");
const helmet   = require("helmet");
const path     = require("path");
const https    = require("https");
const fs       = require("fs");
const os       = require("os");
const crypto   = require("crypto");

// Internal modules - each wraps a specific external API
const graph      = require("./graph");      // Microsoft Graph API helpers
const hubspot    = require("./hubspot");    // HubSpot CRM API helpers
const claude     = require("./claude");     // Anthropic Claude AI helpers
const tokenStore = require("./tokenStore"); // Persists user refresh tokens and session data

const app  = express();
app.set('trust proxy', 1); // Trust first proxy (nginx) so req.ip reflects the real client IP
const PORT = process.env.PORT || 3000;

// ── Constants (env-var backed, with sensible defaults) ────────
const ALLOWED_TONES     = (process.env.ALLOWED_TONES || "Professional,Friendly,Formal,Concise,Persuasive").split(",").map(t => t.trim()).filter(Boolean);
const CONTEXT_CACHE_TTL = parseInt(process.env.CONTEXT_CACHE_TTL_MS || String(5 * 60 * 1000), 10); // How long context lookups are cached (default 5 min)
const RATE_LIMIT        = parseInt(process.env.RATE_LIMIT        || "20",    10);  // Max requests per window per client
const RATE_WINDOW       = parseInt(process.env.RATE_WINDOW_MS    || "60000", 10);  // Rate limit window in ms (default 1 min)
const BODY_LIMIT        = process.env.BODY_LIMIT || "2mb";

const MAX_CACHE_ENTRIES = 5000; // Hard cap to prevent unbounded memory growth
const contextCache  = new Map(); // Caches /api/context results: cacheKey -> { data, expiresAt }
const rateLimitMap  = new Map(); // Tracks request counts per client: clientId -> { count, resetAt }

// ── CORS Origins ─────────────────────────────────────────────
// Only Outlook web clients and our own server are allowed to make requests.
// This prevents random websites from calling our API using a user's session.
const allowedOrigins = [
  "https://outlook.office365.com",
  "https://outlook.office.com",
  "https://outlook.live.com",
  "https://outlook.cloud.microsoft.com",
];
// Optionally allow our own server URL (e.g. for testing or direct access)
if (process.env.SERVER_URL) {
  try {
    const u = new URL(process.env.SERVER_URL);
    if (u.protocol === "https:" || u.protocol === "http:") allowedOrigins.push(process.env.SERVER_URL);
    else console.warn("[cors] SERVER_URL ignored — invalid protocol:", process.env.SERVER_URL);
  } catch {
    console.warn("[cors] SERVER_URL ignored — malformed URL:", process.env.SERVER_URL);
  }
}

// ── Graph Webhook Rate Limiting ───────────────────────────────
// The webhook endpoint is mounted BEFORE the CORS middleware because
// Microsoft Graph sends notifications without an Origin header - CORS
// would reject them otherwise.
const WEBHOOK_RATE_LIMIT  = 120;  // max requests per minute per IP
const WEBHOOK_RATE_WINDOW = 60000;
const webhookRateLimitMap = new Map(); // ip -> { count, resetAt }

// Per-IP rate limiter specifically for the webhook endpoint.
// Separate from the main rate limiter because webhooks come from Microsoft's
// servers, not end users, and need higher throughput.
function webhookRateLimit(req, res, next) {
  const ip = req.ip || req.socket?.remoteAddress || "unknown";
  const now = Date.now();
  const entry = webhookRateLimitMap.get(ip);

  if (!entry || now > entry.resetAt) {
    // Prevent map from growing unbounded - purge expired entries when it gets large
    if (webhookRateLimitMap.size > 10000) {
      for (const [k, v] of webhookRateLimitMap) {
        if (now > v.resetAt) webhookRateLimitMap.delete(k);
      }
    }
    webhookRateLimitMap.set(ip, { count: 1, resetAt: now + WEBHOOK_RATE_WINDOW });
    return next();
  }
  if (entry.count >= WEBHOOK_RATE_LIMIT) {
    return res.status(429).send("Too many requests");
  }
  entry.count++;
  next();
}

// Shared secret set when creating the Graph subscription. Microsoft sends it
// back in every notification as "clientState" so we can verify authenticity.
const GRAPH_WEBHOOK_SECRET = process.env.GRAPH_WEBHOOK_SECRET || "";
if (!GRAPH_WEBHOOK_SECRET) {
  console.warn("[webhook] GRAPH_WEBHOOK_SECRET is not set — webhook notifications will be rejected. Auto-tracking is disabled until a secret is configured.");
}

// ── Webhook processing queue ─────────────────────────────────
// Notifications are queued and processed one at a time (serially).
// This prevents flooding the event loop and avoids hitting HubSpot's
// rate limits when many emails arrive at once.
const _webhookQueue = [];
let _webhookDraining = false;

// Processes queued notifications sequentially until the queue is empty
async function _drainWebhookQueue() {
  _webhookDraining = true;
  while (_webhookQueue.length > 0) {
    const notification = _webhookQueue.shift();
    try {
      const autoTracker = require("./autoTracker");
      await autoTracker.processNotification(notification);
    } catch (err) {
      console.error("[webhook] processNotification error:", err.message);
    }
  }
  _webhookDraining = false;
}

// POST /api/webhooks/graph
// Receives email change notifications from Microsoft Graph.
// Must respond within 3 seconds (Microsoft requirement), so actual
// processing happens asynchronously via the queue above.
// No CORS middleware - Microsoft doesn't send Origin headers.
app.post("/api/webhooks/graph", webhookRateLimit, express.json({ limit: "64kb" }), (req, res) => {
  // Microsoft sends a validation request when first creating a subscription -
  // we must echo back the token to prove we own this endpoint
  if (req.query.validationToken) {
    console.log("[webhook] Validation handshake received");
    res.set("Content-Type", "text/plain");
    return res.status(200).send(req.query.validationToken);
  }

  // Reject all notifications if no webhook secret is configured
  if (!GRAPH_WEBHOOK_SECRET) {
    console.warn("[webhook] Rejecting notifications — GRAPH_WEBHOOK_SECRET not configured");
    return res.status(202).send();
  }

  // Change notifications
  const notifications = req.body?.value;
  if (!Array.isArray(notifications) || notifications.length === 0) {
    return res.status(202).send();
  }

  // Only process notifications whose clientState matches our secret -
  // this confirms they genuinely came from our subscription
  const valid = notifications.filter(n => n.clientState === GRAPH_WEBHOOK_SECRET);
  if (valid.length === 0) {
    console.warn("[webhook] All notifications had invalid clientState — ignoring");
    return res.status(202).send();
  }

  console.log(`[webhook] Received ${valid.length} notification(s)`);

  // Enqueue and start draining if not already running
  for (const notification of valid) {
    _webhookQueue.push(notification);
  }
  if (!_webhookDraining) _drainWebhookQueue();

  // Respond immediately - processing continues in the background
  res.status(202).send();
});

// ── CORS Middleware ───────────────────────────────────────────
// Applied AFTER the webhook route so Microsoft notifications aren't blocked.
// Blocks requests from any origin not in the allowedOrigins list.
app.use(cors({
  origin: (origin, callback) => {
    if (!origin || allowedOrigins.includes(origin)) return callback(null, true);
    console.warn(`[cors] BLOCKED origin: ${origin}`);
    callback(new Error("Not allowed by CORS"));
  },
  credentials: true,
}));

// ── Security Headers (Helmet) ────────────────────────────────
// Helmet for security headers, but CSP DISABLED for Office Add-in compatibility.
// Office.js requires dynamic script loading, telemetry iframes, and host communication
// across multiple Microsoft domains that are difficult to whitelist exhaustively.
// The add-in runs inside Outlook's sandboxed iframe — Outlook enforces its own CSP.
app.use(helmet({
  contentSecurityPolicy: false,     // Disabled — Office.js needs too many dynamic origins
  crossOriginEmbedderPolicy: false, // Office.js webview needs cross-origin resources
  crossOriginOpenerPolicy: false,
  crossOriginResourcePolicy: false, // Outlook webview loads resources cross-origin
  xFrameOptions: false,            // Outlook embeds add-in in iframe; frame-ancestors handled by Outlook
  // Keep the useful non-breaking headers:
  strictTransportSecurity: process.env.NODE_ENV === "production"
    ? { maxAge: 31536000, includeSubDomains: true }
    : false,
}));

// Parse JSON request bodies (with a size limit to prevent abuse)
app.use(express.json({ limit: BODY_LIMIT }));

// Serve the Outlook add-in's static files (HTML, JS, CSS) from public/.
// IMPORTANT: Only public/ is served - never the project root, which contains
// .env, server source, and other sensitive files.
// Cache-busting headers ensure Outlook always loads the latest version.
app.use(express.static(path.join(__dirname, "../public"), {
  setHeaders: (res, filePath) => {
    if (/\.(html|js|css)$/i.test(filePath)) {
      res.setHeader("Cache-Control", "no-store, must-revalidate");
      res.setHeader("Pragma", "no-cache");
      res.setHeader("Expires", "0");
    }
  },
}));

// ── Rate Limiting ─────────────────────────────────────────────
// Identifies clients by hashed bearer token (preferred) or IP address.
// Hashing ensures raw tokens are never stored as map keys.
// Respects X-Forwarded-For via trust proxy so nginx doesn't collapse all users to one IP.
function getClientId(req) {
  const auth = req.headers.authorization;
  if (auth && auth.startsWith("Bearer ")) {
    return crypto.createHash("sha256").update(auth.slice(7)).digest("hex").substring(0, 32);
  }
  // Trust X-Forwarded-For only if app.get('trust proxy') is set
  const ip = req.ip || req.socket?.remoteAddress || "unknown";
  return ip;
}

// Middleware: rejects requests when a client exceeds RATE_LIMIT within RATE_WINDOW
function rateLimit(req, res, next) {
  const clientId = getClientId(req);
  const now = Date.now();
  const entry = rateLimitMap.get(clientId);

  if (!entry || now > entry.resetAt) {
    // Cap map size — purge all expired entries before adding a new one
    if (rateLimitMap.size > 10000) {
      for (const [k, v] of rateLimitMap) {
        if (now > v.resetAt) rateLimitMap.delete(k);
      }
    }
    rateLimitMap.set(clientId, { count: 1, resetAt: now + RATE_WINDOW });
    return next();
  }
  if (entry.count >= RATE_LIMIT) {
    return res.status(429).json({ error: "Too many requests — wait a minute and try again." });
  }
  entry.count++;
  next();
}

// Purge expired rate limit entries every 5 minutes to prevent unbounded growth
const rateLimitCleanupInterval = setInterval(() => {
  const now = Date.now();
  for (const [key, val] of rateLimitMap) {
    if (now > val.resetAt) rateLimitMap.delete(key);
  }
  for (const [key, val] of webhookRateLimitMap) {
    if (now > val.resetAt) webhookRateLimitMap.delete(key);
  }
}, 5 * 60 * 1000);

// Purge expired context cache entries every 10 minutes to prevent unbounded memory growth
const contextCacheCleanupInterval = setInterval(() => {
  const now = Date.now();
  for (const [key, val] of contextCache) {
    if (now > val.expiresAt) contextCache.delete(key);
  }
}, 10 * 60 * 1000);

// ── Auth Helpers ──────────────────────────────────────────────
// Pulls the bearer token from the Authorization header, or returns null
function extractToken(req) {
  const auth = req.headers.authorization;
  if (!auth || !auth.startsWith("Bearer ")) return null;
  return auth.slice(7);
}

// ── JWT Validation ───────────────────────────────────────────
// Validate Microsoft Azure AD tokens using JWKS (JSON Web Key Set).
// The JWKS endpoint and expected claims are derived from env vars so
// nothing is hardcoded to a specific tenant.
const AZURE_TENANT_ID = process.env.AZURE_TENANT_ID;
const AZURE_CLIENT_ID = process.env.AZURE_CLIENT_ID;
const JWT_VALIDATION_ENABLED = !!(AZURE_TENANT_ID && AZURE_CLIENT_ID);

if (!JWT_VALIDATION_ENABLED) {
  if (process.env.NODE_ENV === "production") {
    console.error("[FATAL] AZURE_TENANT_ID or AZURE_CLIENT_ID not set in production — refusing to start without JWT validation.");
    process.exit(1);
  }
  console.warn("[auth] AZURE_TENANT_ID or AZURE_CLIENT_ID not set — JWT validation DISABLED. Set both in .env for production.");
}

// Cache validated tokens for 2 minutes to avoid calling Graph /me on every request
const _tokenValidationCache = new Map(); // tokenHash -> { oid, expiresAt }
const TOKEN_VALIDATION_TTL = 2 * 60 * 1000;

// Middleware: requires a valid Microsoft Graph access token in the Authorization header.
// Validates the token by calling Graph's /me endpoint (since Graph tokens are opaque
// and can't be verified locally by third-party apps). Sets req.token and req.tokenPayload.
async function requireToken(req, res, next) {
  const token = extractToken(req);
  if (!token) return res.status(401).json({ error: "Missing access token" });

  // The token is a Microsoft Graph access token. Graph tokens are opaque /
  // special-format JWTs that Microsoft says should NOT be validated locally
  // by third-party apps. Instead, we verify the token by calling Graph's
  // lightweight /me endpoint. If Graph accepts it, the token is valid and
  // belongs to a real user in our tenant. Results are cached briefly.
  if (JWT_VALIDATION_ENABLED) {
    const tokenHash = crypto.createHash("sha256").update(token).digest("hex");
    const cached = _tokenValidationCache.get(tokenHash);
    if (cached && Date.now() < cached.expiresAt) {
      req.tokenPayload = { oid: cached.oid };
    } else {
      try {
        const controller = new AbortController();
        const timer = setTimeout(() => controller.abort(), 8000);
        const meResp = await fetch("https://graph.microsoft.com/v1.0/me?$select=id", {
          headers: { Authorization: `Bearer ${token}` },
          signal: controller.signal,
        });
        clearTimeout(timer);
        if (!meResp.ok) {
          console.warn("[auth] Graph /me validation failed:", meResp.status);
          return res.status(401).json({ error: "Invalid or expired access token" });
        }
        const meData = await meResp.json();
        req.tokenPayload = { oid: meData.id };
        _tokenValidationCache.set(tokenHash, { oid: meData.id, expiresAt: Date.now() + TOKEN_VALIDATION_TTL });
        // Prune stale entries
        if (_tokenValidationCache.size > 500) {
          const now = Date.now();
          for (const [k, v] of _tokenValidationCache) { if (now > v.expiresAt) _tokenValidationCache.delete(k); }
        }
      } catch (err) {
        console.warn("[auth] Graph /me validation error:", err.message);
        return res.status(401).json({ error: "Invalid or expired access token" });
      }
    }
  }

  req.token = token;
  next();
}

// ── Validation Helpers ────────────────────────────────────────

// Cache key: SHA-256 hash of token + email so raw tokens never appear in logs or memory keys
function makeCacheKey(token, email) {
  return crypto.createHash("sha256").update(`${token}:${email}`).digest("hex");
}

// Basic email format check - not exhaustive but catches most typos
function isValidEmail(email) {
  return typeof email === "string" && /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email);
}

// Validates Graph API identifiers (messageId, conversationId).
// Only allows alphanumeric and safe URL characters to prevent injection.
function isValidGraphId(id, maxLen = 500) {
  return typeof id === "string" && id.length > 0 && id.length <= maxLen && /^[A-Za-z0-9\-_/+=.]+$/.test(id);
}

// ── Error Handler ─────────────────────────────────────────────
// Central error handler for all routes. Hides internal details on 5xx errors
// but passes through 4xx errors with their actual message for client debugging.
function handleError(res, err, context) {
  console.error(`[${context}]`, err.message || err);
  const status = err.status || 500;
  // Only mask true server errors (5xx) — surface 4xx errors with their actual message
  if (status >= 500) {
    return res.status(status).json({ error: "An error occurred. Please try again." });
  }
  res.status(status).json({ error: err.message || "Unexpected error" });
}

// =============================================================================
// AUTH ROUTES
// These handle the OAuth2 authorization code flow with Azure AD.
// The server acts as a "confidential client" - it owns the refresh token
// and the client (Outlook add-in) only ever holds short-lived access tokens.
// =============================================================================

// POST /api/auth/exchange
// Exchanges an OAuth2 authorization code for tokens (part of the sign-in flow).
// Body: { code, verifier } - the auth code and PKCE code verifier.
// Returns: { accessToken, expiresIn, account } - the access token for Graph API calls.
// The refresh token is stored server-side and never sent to the client.
app.post("/api/auth/exchange", async (req, res) => {
  const { code, verifier } = req.body;
  if (!code || !verifier) return res.status(400).json({ error: "Missing code or verifier" });
  try {
    const params = new URLSearchParams({
      client_id: process.env.AZURE_CLIENT_ID,
      grant_type: "authorization_code",
      code,
      redirect_uri: process.env.SERVER_URL + "/auth-complete.html",
      code_verifier: verifier,
      scope: "openid profile User.Read Mail.Read Mail.ReadWrite Mail.Send MailboxSettings.Read offline_access"
    });
    // Confidential client: include client_secret for Web platform
    if (process.env.AZURE_CLIENT_SECRET) {
      params.set("client_secret", process.env.AZURE_CLIENT_SECRET);
    }
    const response = await fetch(
      `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/oauth2/v2.0/token`,
      { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: params.toString() }
    );
    if (!response.ok) {
      const errText = await response.text().catch(() => "Unknown error");
      console.error("[auth/exchange] Microsoft token endpoint returned", response.status, errText.substring(0, 200));
      return res.status(502).json({ error: "Authorization failed. Please try again." });
    }
    let data;
    try { data = await response.json(); }
    catch { return res.status(502).json({ error: "Invalid response from Microsoft token endpoint" }); }
    if (data.error) {
      console.error("[auth/exchange] Microsoft error:", data.error, "-", data.error_description);
      return res.status(400).json({ error: "Authorization failed. Please try signing in again." });
    }
    if (!data.id_token) return res.status(502).json({ error: "Missing id_token in Microsoft response" });
    const idParts = data.id_token.split(".");
    if (idParts.length !== 3) return res.status(502).json({ error: "Malformed id_token from Microsoft" });
    const idPayload = JSON.parse(Buffer.from(idParts[1], "base64url").toString());

    // Server exclusively owns the refresh token — never sent to client.
    // This prevents token rotation conflicts between client and server.
    if (data.refresh_token && idPayload.oid) {
      tokenStore.saveUser(idPayload.oid, {
        refreshToken: data.refresh_token,
        email:        idPayload.preferred_username,
        name:         idPayload.name,
      });
    }

    res.json({
      accessToken:  data.access_token,
      expiresIn:    data.expires_in    || 3600,
      account: { name: idPayload.name || idPayload.preferred_username, username: idPayload.preferred_username, oid: idPayload.oid }
    });
  } catch (err) {
    handleError(res, err, "auth/exchange");
  }
});

// POST /api/auth/handoff
// One-time call after sign-in: client hands off its refresh token to the server.
// From this point, the server exclusively owns the refresh token - the client
// deletes its copy and only holds short-lived access tokens.
app.post("/api/auth/handoff", requireToken, rateLimit, async (req, res) => {
  const userId = req.tokenPayload?.oid;
  const rt = req.body?.refreshToken;
  if (!userId) return res.status(400).json({ error: "Could not identify user" });
  if (!rt) return res.status(400).json({ error: "Missing refreshToken" });

  let email = "", name = "";
  try {
    const parts = req.token.split(".");
    if (parts.length === 3) {
      const payload = JSON.parse(Buffer.from(parts[1], "base64url").toString());
      email = payload.preferred_username || payload.upn || "";
      name  = payload.name || "";
    }
  } catch (e) { /* ignore */ }

  tokenStore.saveUser(userId, { refreshToken: rt, email, name });
  console.log(`[auth/handoff] Stored refresh token for user ${userId} (${email})`);
  res.json({ ok: true });
});

// POST /api/auth/refresh
// Called when the client's access token expires. The server uses its stored
// refresh token to get a new access token from Azure AD.
// NOTE: The expired access token in the Authorization header is used ONLY to
// extract the user's oid (not validated via Graph /me). After refreshing,
// we verify the new token's identity to prevent oid spoofing attacks.
app.post("/api/auth/refresh", rateLimit, async (req, res) => {
  // Extract oid from the (possibly expired) access token — don't validate via Graph /me
  const token = extractToken(req);
  if (!token) return res.status(401).json({ error: "Missing access token" });

  let oid;
  try {
    const parts = token.split(".");
    if (parts.length === 3) {
      const payload = JSON.parse(Buffer.from(parts[1], "base64url").toString());
      oid = payload.oid;
    }
  } catch (e) { /* JWT decode failed */ }

  if (!oid) return res.status(401).json({ error: "Could not identify user from token" });

  const user = tokenStore.getUser(oid);
  if (!user?.refreshToken) {
    return res.status(401).json({ error: "No stored session — please sign in again" });
  }

  try {
    const tokens = await graph.refreshAccessToken(user.refreshToken);

    // Verify the new access token actually belongs to the claimed user.
    // Without this check, an attacker could forge an expired JWT with any oid
    // and hijack another user's refresh token.
    let verifiedOid;
    try {
      const meResp = await fetch("https://graph.microsoft.com/v1.0/me?$select=id", {
        headers: { Authorization: `Bearer ${tokens.accessToken}` },
      });
      if (!meResp.ok) throw new Error(`Graph /me returned ${meResp.status}`);
      const meData = await meResp.json();
      verifiedOid = meData.id;
    } catch (verifyErr) {
      console.error("[auth/refresh] Post-refresh identity verification failed:", verifyErr.message);
      return res.status(401).json({ error: "Identity verification failed — please sign in again" });
    }

    if (verifiedOid !== oid) {
      console.error(`[auth/refresh] OID MISMATCH: token claimed ${oid}, Graph returned ${verifiedOid}`);
      return res.status(401).json({ error: "Identity mismatch — please sign in again" });
    }

    // Save rotated refresh token + cache new access token
    tokenStore.saveUser(oid, {
      refreshToken:         tokens.refreshToken,
      accessToken:          tokens.accessToken,
      accessTokenExpiresAt: new Date(Date.now() + 55 * 60 * 1000).toISOString(),
    });
    res.json({
      accessToken: tokens.accessToken,
      expiresIn:   tokens.expiresIn || 3600,
    });
  } catch (err) {
    console.error("[auth/refresh] Token refresh failed:", err.message);
    // Refresh token is dead — user must re-sign-in
    res.status(401).json({ error: "Session expired — please sign in again" });
  }
});

// POST /api/auth/logout
// Clears the server-side session: refresh token, cached access token,
// and any active Graph webhook subscription for this user.
app.post("/api/auth/logout", requireToken, async (req, res) => {
  const oid = req.tokenPayload?.oid;
  if (!oid) return res.status(400).json({ error: "Could not identify user" });

  tokenStore.deleteUser(oid);
  subscriptions.deleteSubscription(oid, req.token).catch(() => {});
  console.log(`[auth/logout] Cleared session for user ${oid}`);
  res.json({ ok: true });
});

// =============================================================================
// CONTEXT ROUTE
// The main data-gathering endpoint. When a user opens an email in Outlook,
// the add-in calls this to fetch everything needed to generate a smart reply:
// CRM data from HubSpot, email history from Graph, and the user's writing style.
// =============================================================================

// POST /api/context
// Body: { email: string } - the email address of the person the user is emailing.
// Returns: { recipient, deals, crmNotes, emailHistory, senderStyleProfile, ... }
// Results are cached per user+email for CONTEXT_CACHE_TTL to reduce API calls.
app.post("/api/context", requireToken, rateLimit, async (req, res) => {
  const { email } = req.body;
  console.log(`[context] origin=${req.headers.origin || "none"} email=${JSON.stringify(email)} valid=${isValidEmail(email)}`);
  if (!email || !isValidEmail(email)) {
    return res.status(400).json({ error: "Valid email is required" });
  }

  const cacheKey = makeCacheKey(req.token, email);
  const cached = contextCache.get(cacheKey);
  if (cached && Date.now() < cached.expiresAt) {
    return res.json(cached.data);
  }

  try {
    // Fire all four lookups in parallel - each can fail independently without
    // blocking the others (that's why we use allSettled, not Promise.all)
    const [styleResult, historyResult, hubspotResult, portalIdResult] = await Promise.allSettled([
      graph.getSentMailStyle(req.token),
      graph.getEmailHistory(req.token, email),
      hubspot.getContactByEmail(email),
      hubspot.getPortalId(),
    ]);

    if (styleResult.status   === "rejected") console.warn("[context] style fetch failed:", styleResult.reason?.message);
    if (historyResult.status === "rejected") console.warn("[context] history fetch failed:", historyResult.reason?.message);
    if (hubspotResult.status === "rejected") console.warn("[context] hubspot fetch failed:", hubspotResult.reason?.message);
    if (portalIdResult.status === "rejected") console.warn("[context] portalId fetch failed:", portalIdResult.reason?.message);

    const style    = styleResult.status    === "fulfilled" ? styleResult.value    : null;
    const history  = historyResult.status  === "fulfilled" ? historyResult.value  : [];
    const crm      = hubspotResult.status  === "fulfilled" ? hubspotResult.value  : null;
    const portalId = portalIdResult.status === "fulfilled" ? portalIdResult.value : null;

    const recipient = crm?.contact || { name: "", email, company: "", jobTitle: "", lifecycleStage: "", hubspotOwner: "" };

    const hubspotConfigured = !!process.env.HUBSPOT_TOKEN;
    const context = {
      recipient,
      deals:             crm?.deals   || [],
      crmNotes:          crm?.notes   || [],
      emailHistory:      history,
      senderStyleProfile: style,
      hubspotSearched:   hubspotConfigured, // tells UI whether HubSpot was actually queried
      ...(portalId ? { hubspotPortalId: portalId } : {}),
    };

    // Enforce cache size limit to prevent unbounded memory growth
    if (contextCache.size >= MAX_CACHE_ENTRIES) {
      const now = Date.now();
      for (const [k, v] of contextCache) { if (now > v.expiresAt) contextCache.delete(k); }
    }
    if (contextCache.size < MAX_CACHE_ENTRIES) {
      contextCache.set(cacheKey, { data: context, expiresAt: Date.now() + CONTEXT_CACHE_TTL });
    }
    res.json(context);
  } catch (err) {
    handleError(res, err, "POST /api/context");
  }
});

// =============================================================================
// AI EMAIL GENERATION ROUTES
// =============================================================================

// POST /api/generate
// Generates a new email draft using Claude AI.
// Body: { context, tone, senderName, customNote, isReply?, conversationId? }
// If isReply is true, fetches the email thread server-side for full context.
// Returns: { subject, body } - the AI-generated email content.
app.post("/api/generate", requireToken, rateLimit, async (req, res) => {
  const { context, tone, senderName, customNote, isReply, conversationId } = req.body;
  if (!customNote) return res.status(400).json({ error: "customNote is required" });

  const safeTone = ALLOWED_TONES.includes(tone) ? tone : "Professional";

  try {
    // If reply mode, fetch thread context server-side for security
    let threadContext = [];
    if (isReply && conversationId) {
      if (!isValidGraphId(conversationId)) return res.status(400).json({ error: "Invalid conversationId format" });
      threadContext = await graph.getEmailThread(req.token, conversationId);
    }
    const result = await claude.generateEmail({ context, tone: safeTone, senderName, customNote, isReply: !!isReply, threadContext });
    res.json(result);
  } catch (err) {
    handleError(res, err, "POST /api/generate");
  }
});

// POST /api/refine
// Iteratively refines an existing draft based on user feedback (chat-style).
// Body: { context, currentSubject, currentBody, userMessage, conversationId?, chatHistory? }
// Returns: { subject, body } - the refined email content.
app.post("/api/refine", requireToken, rateLimit, async (req, res) => {
  const { context, currentSubject, currentBody, userMessage, conversationId, chatHistory } = req.body;
  if (!userMessage) return res.status(400).json({ error: "userMessage is required" });

  try {
    // Fetch thread context if replying to a conversation (same as /api/generate)
    let threadContext = [];
    if (conversationId) {
      if (!isValidGraphId(conversationId)) return res.status(400).json({ error: "Invalid conversationId format" });
      threadContext = await graph.getEmailThread(req.token, conversationId);
    }
    const result = await claude.refineEmail({ context, currentSubject, currentBody, userMessage, threadContext, chatHistory: chatHistory || [] });
    res.json(result);
  } catch (err) {
    handleError(res, err, "POST /api/refine");
  }
});

// POST /api/transcribe
// Accepts raw audio bytes and transcribes them using Groq's Whisper API.
// Content-Type must be application/octet-stream (raw binary, not multipart).
// Returns: { text } - the transcribed text.
// Used for voice-to-text input in the add-in.
app.post(
  "/api/transcribe",
  requireToken,
  rateLimit,
  express.raw({ type: "application/octet-stream", limit: "25mb" }),
  async (req, res) => {
    if (!process.env.GROQ_API_KEY) {
      return res.status(500).json({ error: "GROQ_API_KEY not configured on server" });
    }
    if (!req.body || !req.body.length) {
      return res.status(400).json({ error: "Audio body required" });
    }
    try {
      const form = new FormData();
      form.append("file", new Blob([req.body], { type: "audio/webm" }), "audio.webm");
      form.append("model", process.env.GROQ_STT_MODEL || "whisper-large-v3-turbo");
      form.append("response_format", "json");
      form.append("language", "en");
      form.append("temperature", "0");

      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), 30000);
      try {
        const r = await fetch("https://api.groq.com/openai/v1/audio/transcriptions", {
          method: "POST",
          headers: { Authorization: `Bearer ${process.env.GROQ_API_KEY}` },
          body: form,
          signal: controller.signal,
        });
        if (!r.ok) {
          const errText = await r.text().catch(() => "");
          console.error("[transcribe] Groq error", r.status, errText.substring(0, 300));
          return res.status(502).json({ error: `Transcription failed (${r.status})` });
        }
        const data = await r.json();
        res.json({ text: (data.text || "").trim() });
      } finally {
        clearTimeout(timer);
      }
    } catch (err) {
      handleError(res, err, "POST /api/transcribe");
    }
  }
);

// POST /api/contextualize-voice
// Takes raw transcribed speech and cleans it up using Claude - removes filler
// words, fixes grammar, and structures it as a proper email instruction.
// Body: { rawText }
// Returns: { text } - the cleaned-up text.
app.post("/api/contextualize-voice", requireToken, rateLimit, async (req, res) => {
  const { rawText } = req.body;
  if (!rawText) return res.status(400).json({ error: "rawText is required" });

  try {
    const text = await claude.contextualizeVoice(rawText);
    res.json({ text });
  } catch (err) {
    handleError(res, err, "POST /api/contextualize-voice");
  }
});

// =============================================================================
// HUBSPOT CRM ROUTES
// =============================================================================

// POST /api/hubspot/log
// Logs an email interaction as a CRM note on a HubSpot contact.
// Uses Claude to summarize the email body before saving.
// Body: { contactEmail, subject, body }
// Returns: { success, noteId }
app.post("/api/hubspot/log", requireToken, rateLimit, async (req, res) => {
  const { contactEmail, subject, body } = req.body;
  if (!contactEmail) return res.status(400).json({ error: "contactEmail is required" });

  try {
    const summary = body ? await claude.summarizeForCRM(body) : subject;
    const note = await hubspot.createNote(contactEmail, subject || "(no subject)", summary);
    res.json({ success: true, noteId: note.id });
  } catch (err) {
    handleError(res, err, "POST /api/hubspot/log");
  }
});

// =============================================================================
// GRAPH EMAIL ROUTES
// These proxy email operations through the server so the client never calls
// Graph directly for mutations (send, reply, draft).
// =============================================================================

// POST /api/graph/send
// Sends a new email via Microsoft Graph on behalf of the user.
// Body: { toEmail, subject, htmlBody }
// Returns: { success }
app.post("/api/graph/send", requireToken, rateLimit, async (req, res) => {
  const { toEmail, subject, htmlBody } = req.body;
  if (!toEmail || !isValidEmail(toEmail)) {
    return res.status(400).json({ error: "Valid toEmail is required" });
  }
  if (subject && subject.length > 500) return res.status(400).json({ error: "Subject too long (max 500 chars)" });
  if (htmlBody && htmlBody.length > 100000) return res.status(400).json({ error: "Email body too long (max 100KB)" });
  try {
    await graph.sendEmail(req.token, { toEmail, subject, htmlBody });
    res.json({ success: true });
  } catch (err) {
    handleError(res, err, "POST /api/graph/send");
  }
});

// POST /api/graph/reply
// Sends a reply (or reply-all) to a specific message via Graph API.
// Body: { messageId, htmlBody, replyAll }
// Returns: { success }
app.post("/api/graph/reply", requireToken, rateLimit, async (req, res) => {
  const { messageId, htmlBody, replyAll } = req.body;
  if (!isValidGraphId(messageId, 2000)) {
    return res.status(400).json({ error: "messageId must be a valid Graph API identifier" });
  }
  if (!htmlBody)  return res.status(400).json({ error: "htmlBody is required" });
  try {
    await graph.replyToMessage(req.token, messageId, htmlBody, !!replyAll);
    res.json({ success: true });
  } catch (err) {
    handleError(res, err, "POST /api/graph/reply");
  }
});

// POST /api/graph/create-reply-draft
// Creates a reply-all draft (not sent) so the user can review it in Outlook.
// Body: { messageId, htmlBody }
// Returns: { draftId, webLink } - webLink lets the client open the draft in Outlook
app.post("/api/graph/create-reply-draft", requireToken, rateLimit, async (req, res) => {
  const { messageId, htmlBody } = req.body;
  if (!isValidGraphId(messageId, 2000)) {
    return res.status(400).json({ error: "messageId must be a valid Graph API identifier" });
  }
  if (!htmlBody) return res.status(400).json({ error: "htmlBody is required" });
  try {
    const result = await graph.createReplyAllDraft(req.token, messageId, htmlBody);
    res.json(result);
  } catch (err) {
    handleError(res, err, "POST /api/graph/create-reply-draft");
  }
});

// POST /api/graph/create-draft
// Creates a brand new draft message (not a reply - no existing thread).
// Body: { to (array of emails), subject, htmlBody }
// Returns: { draftId, webLink }
app.post("/api/graph/create-draft", requireToken, rateLimit, async (req, res) => {
  const { to, subject, htmlBody } = req.body;
  if (!htmlBody) return res.status(400).json({ error: "htmlBody is required" });
  try {
    const result = await graph.createNewDraft(req.token, { to: to || [], subject, htmlBody });
    res.json(result);
  } catch (err) {
    handleError(res, err, "POST /api/graph/create-draft");
  }
});

// GET /api/graph/signature
// Returns the user's email signature extracted from their most recent sent email.
// Graph API doesn't have a dedicated signature endpoint, so we parse sent mail.
app.get("/api/graph/signature", requireToken, rateLimit, async (req, res) => {
  try {
    const signature = await graph.getSignatureFromSentMail(req.token);
    res.json({ signature });
  } catch (err) {
    handleError(res, err, "GET /api/graph/signature");
  }
});

// GET /api/hubspot/contact?email=...
// Looks up a contact in HubSpot by email address.
// Returns contact details, associated deals, and CRM notes.
app.get("/api/hubspot/contact", requireToken, rateLimit, async (req, res) => {
  const { email } = req.query;
  if (!email) return res.status(400).json({ error: "email query param is required" });

  try {
    const result = await hubspot.getContactByEmail(email);
    if (!result) return res.status(404).json({ error: "Contact not found" });
    res.json(result);
  } catch (err) {
    handleError(res, err, "GET /api/hubspot/contact");
  }
});

// GET /api/graph/history?email=...
// Returns recent email history between the current user and the given email address.
app.get("/api/graph/history", requireToken, rateLimit, async (req, res) => {
  const { email } = req.query;
  if (!email) return res.status(400).json({ error: "email query param is required" });

  try {
    const history = await graph.getEmailHistory(req.token, email);
    res.json(history);
  } catch (err) {
    handleError(res, err, "GET /api/graph/history");
  }
});

// GET /api/graph/style
// Analyzes the user's recent sent emails to build a writing style profile.
// Claude uses this to mimic the user's natural tone in generated emails.
app.get("/api/graph/style", requireToken, rateLimit, async (req, res) => {
  try {
    const style = await graph.getSentMailStyle(req.token);
    res.json(style);
  } catch (err) {
    handleError(res, err, "GET /api/graph/style");
  }
});

// POST /api/feedback
// Records user feedback (thumbs up/down) on generated emails.
// Currently just logs - can be persisted to a database later for model improvement.
app.post("/api/feedback", requireToken, rateLimit, async (req, res) => {
  const { positive, tone, hasEdit } = req.body;
  // Log for now — can be persisted to DB later
  console.log(`[feedback] positive=${positive} tone=${tone} hasEdit=${hasEdit}`);
  res.json({ ok: true });
});

// GET /api/people?q=:query
// Contact autocomplete - searches the user's Graph contacts/directory.
// Requires at least 2 characters. Returns matching people with emails.
app.get("/api/people", requireToken, rateLimit, async (req, res) => {
  const q = (req.query.q || "").trim();
  if (!q || q.length < 2) return res.json([]);
  try {
    const results = await graph.searchPeople(req.token, q);
    res.json(results);
  } catch (err) {
    handleError(res, err, "GET /api/people");
  }
});

// GET /api/analytics
// Legacy endpoint kept for backwards compatibility.
// Returns email activity analytics from Graph.
app.get("/api/analytics", requireToken, rateLimit, async (req, res) => {
  try {
    const data = await graph.getEmailAnalytics(req.token);
    res.json(data);
  } catch (err) {
    handleError(res, err, "GET /api/analytics");
  }
});

// =============================================================================
// FOLLOW-UP ROUTES
// These power the "Follow-ups" tab, which shows deals that need attention.
// =============================================================================

// GET /api/followups
// Returns open HubSpot deals the current user is involved in, sorted by
// staleness (days since last activity). Deals are matched by:
//   1. HubSpot owner ID (user owns the deal)
//   2. Contact involvement (user's email appears as a contact on the deal)
// Each deal is enriched with associated contact emails for the follow-up flow.
app.get("/api/followups", requireToken, rateLimit, async (req, res) => {
  try {
    // 1. Extract user email from JWT
    let userEmail = null;
    try {
      const parts = req.token.split(".");
      if (parts.length === 3) {
        const payload = JSON.parse(Buffer.from(parts[1], "base64url").toString());
        userEmail = (payload.preferred_username || payload.upn || "").toLowerCase();
      }
    } catch (e) { /* ignore */ }

    // 2. Resolve HubSpot owner ID (gracefully returns null if scope unavailable)
    const ownerId = userEmail ? await hubspot.getOwnerIdByEmail(userEmail) : null;

    // 3. Fetch all open deals (globally cached, 2-min TTL)
    const allDeals = await hubspot.getFollowupDeals();

    // 4. Fast-filter: deals where user is the owner (no API calls needed)
    const ownerMatched = ownerId ? allDeals.filter(d => d.ownerId === ownerId) : [];
    const ownerIds = new Set(ownerMatched.map(d => d.id));

    // Enrich owner-matched deals with contacts (batched to avoid HubSpot rate limits)
    const ENRICH_BATCH = 5;
    for (let i = 0; i < ownerMatched.length; i += ENRICH_BATCH) {
      const batch = ownerMatched.slice(i, i + ENRICH_BATCH);
      const results = await Promise.allSettled(batch.map(d => hubspot.getDealContacts(d.id)));
      for (let j = 0; j < batch.length; j++) {
        if (results[j].status === "fulfilled") {
          batch[j].contacts = results[j].value;
          if (results[j].value.length > 0 && results[j].value[0].company) {
            batch[j].company = results[j].value[0].company;
          }
        }
      }
    }

    // Also check non-owned deals - the user might be a participant (e.g. CC'd)
    let userDeals = [...ownerMatched];
    if (userEmail) {
      const nonOwned = allDeals.filter(d => !ownerIds.has(d.id)).slice(0, 30);
      for (let i = 0; i < nonOwned.length; i += ENRICH_BATCH) {
        const batch = nonOwned.slice(i, i + ENRICH_BATCH);
        const results = await Promise.allSettled(batch.map(d => hubspot.getDealContacts(d.id)));
        for (let j = 0; j < batch.length; j++) {
          if (results[j].status === "fulfilled") {
            batch[j].contacts = results[j].value;
            if (results[j].value.length > 0 && results[j].value[0].company) {
              batch[j].company = results[j].value[0].company;
            }
            if (results[j].value.some(c => c.email === userEmail)) {
              userDeals.push(batch[j]);
            }
          }
        }
      }
    }

    // 5. Sort by staleness
    userDeals.sort((a, b) => b.daysSince - a.daysSince);

    res.json({ deals: userDeals });
  } catch (err) {
    handleError(res, err, "GET /api/followups");
  }
});

// POST /api/followups/draft
// Generates an AI follow-up email for a specific deal.
// Body: { dealId, dealName, dealStage, daysSince, contactEmail, contactName, company }
//
// Logic flow:
//   1. Check if there's a manually linked email thread for this deal (via dedupStore)
//   2. If linked thread exists, fetch it and resolve the external contact from it
//   3. If no linked thread, resolve best contact from HubSpot deal associations
//   4. Generate follow-up email with Claude using deal context + email history
//   5. Return the draft + replyMeta (if replying in a thread) or resolvedContact (if new compose)
app.post("/api/followups/draft", requireToken, rateLimit, async (req, res) => {
  const { dealId, dealName, dealStage, daysSince, contactEmail, contactName, company } = req.body;
  if (!dealName) return res.status(400).json({ error: "dealName is required" });

  try {
    let lastEmail = null;
    let thread = [];
    let senderStyleProfile = null;
    let resolvedContactName = contactName;
    let resolvedContactEmail = contactEmail;

    // Decode user email from JWT to exclude from contact resolution
    let userEmail = null;
    try {
      const parts = req.token.split(".");
      if (parts.length === 3) {
        const payload = JSON.parse(Buffer.from(parts[1], "base64url").toString());
        userEmail = (payload.preferred_username || payload.upn || "").toLowerCase();
      }
    } catch (_) {}

    // Only reply in a thread if the user explicitly linked it to this deal
    // (via "Link Email" button). We don't guess threads from contact search
    // because the wrong thread could be selected.
    const dedupStore = require("./dedupStore");
    const linkedConvId = dealId ? dedupStore.getConversationIdForDeal(dealId) : null;

    if (linkedConvId) {
      try {
        const [latestResult, threadResult] = await Promise.allSettled([
          graph.getLatestMessageInThread(req.token, linkedConvId),
          graph.getEmailThreadForLogging(req.token, linkedConvId),
        ]);
        if (latestResult.status === "fulfilled" && latestResult.value) {
          lastEmail = latestResult.value;
        }
        if (threadResult.status === "fulfilled") {
          thread = threadResult.value || [];
        }

        // Walk the thread backwards to find the most recent external
        // participant (someone who isn't the current user)
        if (thread.length > 0 && userEmail) {
          for (let i = thread.length - 1; i >= 0; i--) {
            const msg = thread[i];
            const fromEmail = (msg.fromEmail || "").toLowerCase();
            if (fromEmail && fromEmail !== userEmail) {
              resolvedContactEmail = fromEmail;
              resolvedContactName = msg.fromName || fromEmail;
              break;
            }
            const toList = (msg.toEmails || "").split(";").map(e => e.trim().toLowerCase()).filter(Boolean);
            const external = toList.find(e => e !== userEmail);
            if (external) {
              resolvedContactEmail = external;
              resolvedContactName = null;
              break;
            }
          }
        }
      } catch (err) {
        console.warn("[followups/draft] Failed to fetch linked thread:", err.message);
      }
    }

    // Fallback: resolve best contact from HubSpot deal associations.
    // Picks the contact with the most recent email exchange.
    // This does NOT set replyMeta - without a linked thread, the client opens a new compose window.
    if (dealId && !resolvedContactName) {
      try {
        const dealContacts = await hubspot.getDealContacts(dealId);
        if (dealContacts.length > 0) {
          // Find the contact with the most recent email for best context
          const emailResults = await Promise.allSettled(
            dealContacts.slice(0, 5).map(c => graph.getLastEmailWithContact(req.token, c.email))
          );

          let bestContact = null;
          let bestDate = null;

          for (let i = 0; i < emailResults.length; i++) {
            if (emailResults[i].status !== "fulfilled" || !emailResults[i].value) continue;
            const receivedAt = emailResults[i].value.receivedAt ? new Date(emailResults[i].value.receivedAt).getTime() : 0;
            if (!bestDate || receivedAt > bestDate) {
              bestDate = receivedAt;
              bestContact = dealContacts[i];
            }
          }

          if (bestContact) {
            resolvedContactEmail = bestContact.email;
            resolvedContactName = bestContact.name;
          } else {
            resolvedContactEmail = dealContacts[0].email;
            resolvedContactName = dealContacts[0].name;
          }

          // Fetch email history with the resolved contact for Claude context
          if (thread.length === 0 && resolvedContactEmail) {
            try {
              thread = await graph.getEmailHistory(req.token, resolvedContactEmail) || [];
            } catch (_) {}
          }
        }
      } catch (err) {
        console.warn("[followups/draft] Deal contact resolution failed:", err.message);
      }
    }

    // Fetch style profile
    try {
      senderStyleProfile = await graph.getSentMailStyle(req.token);
    } catch (_) {}

    const result = await claude.generateFollowUp({
      dealName,
      dealStage,
      daysSince: daysSince || 0,
      contactName: resolvedContactName || contactName,
      company,
      thread,
      senderStyleProfile,
    });

    // replyMeta tells the client whether to reply in a thread or open a new compose.
    // Only set when we have an explicitly linked thread - otherwise null.
    let replyMeta = null;
    if (lastEmail && linkedConvId) {
      replyMeta = {
        messageId:      lastEmail.messageId,
        conversationId: lastEmail.conversationId,
        subject:        lastEmail.subject,
      };
    }

    // Include resolved contact so the client can set the To field in compose mode
    res.json({ ...result, replyMeta, resolvedContact: { email: resolvedContactEmail, name: resolvedContactName } });
  } catch (err) {
    handleError(res, err, "POST /api/followups/draft");
  }
});

// GET /api/graph/thread?conversationId=...
// Returns all messages in an email conversation thread.
// Used by the client to display thread context in reply mode.
app.get("/api/graph/thread", requireToken, rateLimit, async (req, res) => {
  const { conversationId } = req.query;
  if (!conversationId || !isValidGraphId(conversationId)) return res.status(400).json({ error: "Valid conversationId is required" });
  try {
    const thread = await graph.getEmailThread(req.token, conversationId);
    res.json(thread);
  } catch (err) {
    handleError(res, err, "GET /api/graph/thread");
  }
});

// GET /api/hubspot/pipelines
// Returns all deal pipelines and their stages from HubSpot.
// Used to populate the pipeline/stage dropdowns when creating deals.
app.get("/api/hubspot/pipelines", requireToken, rateLimit, async (req, res) => {
  try {
    const pipelines = await hubspot.getPipelines();
    res.json(pipelines);
  } catch (err) {
    handleError(res, err, "GET /api/hubspot/pipelines");
  }
});

// GET /api/hubspot/deal-properties
// Returns all writable deal properties from HubSpot (cached 5 min).
// Used to dynamically build the deal creation form.
app.get("/api/hubspot/deal-properties", requireToken, rateLimit, async (req, res) => {
  try {
    const props = await hubspot.getDealProperties();
    res.json(props);
  } catch (err) {
    handleError(res, err, "GET /api/hubspot/deal-properties");
  }
});

// POST /api/hubspot/deal
// Creates a new deal in HubSpot and associates it with the given contact.
// Body: { contactEmail, dealname, ...otherDealProperties }
// contactEmail and dealname are required; all other fields pass through as deal properties.
app.post("/api/hubspot/deal", requireToken, rateLimit, async (req, res) => {
  const { contactEmail, ...dealProps } = req.body;
  if (!contactEmail)      return res.status(400).json({ error: "contactEmail is required" });
  if (!dealProps.dealname) return res.status(400).json({ error: "dealname is required" });
  try {
    const result = await hubspot.createDeal(contactEmail, dealProps);
    res.json({ success: true, ...result });
  } catch (err) {
    handleError(res, err, "POST /api/hubspot/deal");
  }
});

// POST /api/suggest-deal-name
// Uses Claude to suggest a deal name based on contact info and existing deal names.
// Body: { contactName, company, jobTitle }
// Returns: { name }
app.post("/api/suggest-deal-name", requireToken, rateLimit, async (req, res) => {
  const { contactName, company, jobTitle } = req.body;
  try {
    const existingDealNames = await hubspot.getRecentDealNames();
    const name = await claude.suggestDealName({ contactName, company, jobTitle, existingDealNames });
    res.json({ name });
  } catch (err) {
    handleError(res, err, "POST /api/suggest-deal-name");
  }
});

// GET /api/hubspot/deals/search?q=...
// Searches HubSpot deals by name. Used for the "Link to Deal" search UI.
app.get("/api/hubspot/deals/search", requireToken, rateLimit, async (req, res) => {
  const q = (req.query.q || "").trim();
  try {
    const deals = await hubspot.searchDeals(q);
    res.json(deals);
  } catch (err) {
    handleError(res, err, "GET /api/hubspot/deals/search");
  }
});

// POST /api/hubspot/deal/link
// Associates one or more contacts with an existing HubSpot deal.
// Body: { contactEmails: string[], dealId } or { contactEmail: string, dealId }
// Supports both array (new) and single string (backwards compat) formats.
// Skips contacts already on the deal. Returns: { success, linked, alreadyLinked, errors }
app.post("/api/hubspot/deal/link", requireToken, rateLimit, async (req, res) => {
  const { contactEmail, contactEmails, dealId } = req.body;

  // Normalize: accept array (new) or single string (backwards compat)
  let emails = [];
  if (Array.isArray(contactEmails) && contactEmails.length > 0) {
    emails = contactEmails.filter(e => typeof e === "string" && e.includes("@"));
  } else if (contactEmail) {
    emails = [contactEmail];
  }

  if (emails.length === 0) return res.status(400).json({ error: "contactEmail or contactEmails is required" });
  if (!dealId)             return res.status(400).json({ error: "dealId is required" });

  // Cap to prevent abuse
  emails = emails.slice(0, 20);
  console.log(`[deal/link] Linking ${emails.length} contact(s) to deal ${dealId}: ${emails.join(", ")}`);

  // Pre-fetch existing contacts on the deal to detect already-linked
  let existingEmails = new Set();
  try {
    const existing = await hubspot.getDealContacts(dealId);
    existingEmails = new Set(existing.map(c => c.email.toLowerCase()));
  } catch (e) {
    console.warn("[deal/link] Could not fetch existing contacts:", e.message);
  }

  let linked = 0;
  let alreadyLinked = 0;
  const errors = [];

  for (const email of emails) {
    if (existingEmails.has(email.toLowerCase())) {
      alreadyLinked++;
      continue;
    }
    try {
      await hubspot.linkContactToDeal(email, dealId);
      linked++;
    } catch (err) {
      console.error(`[deal/link] FAILED: ${email} → deal ${dealId}:`, err.message);
      errors.push({ email, error: err.message });
    }
  }

  console.log(`[deal/link] Result: ${linked} linked, ${alreadyLinked} already on deal, ${errors.length} failed → deal ${dealId}`);
  res.json({ success: true, linked, alreadyLinked, errors });
});

// POST /api/hubspot/deal/preview-link-email
// Dry run - shows which emails would be linked without writing to HubSpot.
// Body: { conversationId, subject }
// Returns: { messages } - list of emails in the thread with basic metadata.
app.post("/api/hubspot/deal/preview-link-email", requireToken, rateLimit, async (req, res) => {
  const { conversationId, subject } = req.body;
  if (!conversationId) return res.status(400).json({ error: "conversationId is required" });
  try {
    const messages = await graph.getEmailThreadForLogging(req.token, conversationId);

    res.json({
      messages: messages.map(m => ({
        subject:   m.subject,
        fromEmail: m.fromEmail,
        fromName:  m.fromName,
        sentAt:    m.sentAt,
      })),
    });
  } catch (err) {
    handleError(res, err, "POST /api/hubspot/deal/preview-link-email");
  }
});

// POST /api/hubspot/deal/link-email
// The core "Link Email" action. Fetches an email thread via Graph API, then
// logs each message as a separate email engagement on the HubSpot deal.
// Body: { dealId, conversationId, contactEmail, subject, force? }
//
// This is the most complex route because it:
//   1. Validates participant/domain/subject overlap to prevent linking the wrong thread
//   2. Deduplicates against emails already on the deal
//   3. Responds immediately (to avoid nginx timeouts) then processes in the background
//   4. Auto-links thread participants as contacts on the deal (matching domain only)
//   5. Saves a conversation-to-deal mapping for future follow-up replies
//
// Pass force=true to skip validation checks (participant/subject mismatch).
app.post("/api/hubspot/deal/link-email", requireToken, rateLimit, async (req, res) => {
  const { dealId, conversationId, contactEmail, subject } = req.body;
  if (!dealId) return res.status(400).json({ error: "dealId is required" });
  if (!conversationId || !isValidGraphId(conversationId)) return res.status(400).json({ error: "Valid conversationId is required" });
  try {
    // Decode user email from the Graph API JWT (second segment = payload)
    let userEmail = null;
    try {
      const parts = req.token.split(".");
      if (parts.length === 3) {
        const payload = JSON.parse(Buffer.from(parts[1], "base64url").toString());
        userEmail = payload.preferred_username || payload.upn || null;
      }
    } catch (jwtErr) {
      console.warn("[deal/link-email] Could not decode JWT for user email:", jwtErr.message);
    }

    // Fire all independent lookups in parallel to cut latency
    const [messagesResult, contactResult, existingEmailsResult, dealContactsResult] = await Promise.allSettled([
      graph.getEmailThreadForLogging(req.token, conversationId),
      contactEmail ? hubspot.getContactByEmail(contactEmail) : Promise.resolve(null),
      hubspot.getDealEmails(dealId),
      hubspot.getDealContacts(dealId),
    ]);

    const messages = messagesResult.status === "fulfilled" ? messagesResult.value : [];
    if (messages.length === 0) {
      return res.status(404).json({ error: "No messages found in this thread" });
    }

    // ── Validation (skipped if force=true) ──────────────────────
    // Three checks to prevent linking the wrong thread to a deal:
    //   1. External participant overlap - are the right people on this thread?
    //   2. Domain overlap - is this thread with the right company?
    //   3. Subject consistency - does this thread match existing emails on the deal?
    if (!req.body.force) {
      try {
        const dealContacts = dealContactsResult.status === "fulfilled" ? dealContactsResult.value : [];
        const internalDomain = userEmail ? userEmail.split("@")[1]?.toLowerCase() : null;
        const isExternal = (email) => {
          if (!email || !internalDomain) return true;
          return !email.toLowerCase().endsWith("@" + internalDomain);
        };

        // Collect external participants from the thread
        const threadParticipants = new Set();
        const threadDomains = new Set();
        for (const msg of messages) {
          for (const addr of [msg.fromEmail, ...(msg.toEmails || "").split(";"), ...(msg.ccEmails || "").split(";")]) {
            const trimmed = (addr || "").trim().toLowerCase();
            if (trimmed && isExternal(trimmed)) {
              threadParticipants.add(trimmed);
              const domain = trimmed.split("@")[1];
              if (domain) threadDomains.add(domain);
            }
          }
        }

        // Check 1: External participant overlap
        const externalDealEmails = dealContacts.map(c => c.email).filter(isExternal);
        const externalDealDomains = new Set(externalDealEmails.map(e => e.split("@")[1]).filter(Boolean));
        const hasParticipantOverlap = [...threadParticipants].some(e => externalDealEmails.includes(e));

        if (!hasParticipantOverlap && externalDealEmails.length > 0) {
          // Check 2: Domain overlap (weaker signal — same company, different person)
          const hasDomainOverlap = [...threadDomains].some(d => externalDealDomains.has(d));
          if (!hasDomainOverlap) {
            console.warn(`[deal/link-email] No participant or domain overlap: thread domains=[${[...threadDomains].join(",")}], deal domains=[${[...externalDealDomains].join(",")}]`);
            return res.status(422).json({
              error: "No email participants match this deal's contacts or companies. Please verify you selected the correct deal.",
              code: "NO_CONTACT_OVERLAP",
              dealContacts: dealContacts.filter(c => isExternal(c.email)).map(c => c.name || c.email),
            });
          }
        }

        // Check 3: Subject consistency — if the deal already has emails,
        // verify the new thread's subject is related to at least one of them.
        // This catches the case where the right contact has multiple unrelated threads.
        const existingEmails = existingEmailsResult.status === "fulfilled" ? existingEmailsResult.value : [];
        if (existingEmails.length > 0 && messages.length > 0) {
          const { normalizeSubject } = require("./graph");
          const existingSubjects = new Set(existingEmails.map(e => normalizeSubject(e.subject)).filter(Boolean));
          const newThreadSubject = normalizeSubject(messages[0].subject);

          if (newThreadSubject && existingSubjects.size > 0 && !existingSubjects.has(newThreadSubject)) {
            // New thread subject doesn't match ANY existing email on the deal
            const existingSubjectList = [...new Set(existingEmails.map(e => e.subject).filter(Boolean))].slice(0, 3);
            console.warn(`[deal/link-email] Subject mismatch: new thread="${messages[0].subject}", existing=[${existingSubjectList.join("; ")}]`);
            return res.status(422).json({
              error: `This thread ("${messages[0].subject}") doesn't match any existing emails on this deal. Please verify you selected the correct deal and thread.`,
              code: "SUBJECT_MISMATCH",
              existingSubjects: existingSubjectList,
            });
          }
        }
      } catch (e) {
        console.warn("[deal/link-email] Validation check failed, proceeding anyway:", e.message);
      }
    }

    // Use already-fetched contact result
    let preResolvedContactId = null;
    if (contactResult.status === "fulfilled" && contactResult.value) {
      preResolvedContactId = contactResult.value?.contact?.hubspotId || null;
    } else if (contactResult.status === "rejected") {
      console.warn("[deal/link-email] Contact lookup failed, will skip contact association:", contactResult.reason?.message);
    }

    // Use already-fetched existing emails for dedup
    let existingEmailSubjects = new Set();
    if (existingEmailsResult.status === "fulfilled") {
      for (const e of existingEmailsResult.value) {
        existingEmailSubjects.add((e.subject || "") + "|" + (e.timestamp || ""));
      }
    } else {
      console.warn("[deal/link-email] Could not check existing emails, proceeding without dedup:", existingEmailsResult.reason?.message);
    }

    // Count how many emails will be linked (after dedup)
    const toLink = messages.filter(msg => {
      const key = (msg.subject || "") + "|" + (msg.sentAt || "");
      return !existingEmailSubjects.has(key);
    });
    const skipped = messages.length - toLink.length;
    if (skipped > 0) console.log(`[deal/link-email] Will skip ${skipped} already-linked email(s)`);

    // Save conversation→deal mapping immediately (before background processing)
    if (conversationId) {
      const dedupStore = require("./dedupStore");
      const threadSubject = messages.length > 0 ? messages[0].subject : (subject || "");
      dedupStore.setConversationDeal(conversationId, dealId, threadSubject);
      console.log(`[deal/link-email] Saved conversation mapping: ${conversationId} → deal ${dealId} (subject: "${threadSubject}")`);
    }

    // Respond immediately — HubSpot linking continues in the background.
    // This prevents nginx 502 timeouts on large threads (each email = 3 rate-limited API calls).
    res.json({ success: true, linked: toLink.length, total: messages.length });

    // ── Background processing (after response sent) ──────────────
    // Log each message as a separate HubSpot email engagement
    let linked = 0;
    for (const msg of toLink) {
      try {
        await hubspot.linkEmailToDeal(
          { subject: msg.subject, bodyHtml: msg.bodyHtml, bodyText: msg.bodyText, fromEmail: msg.fromEmail, toEmail: msg.toEmails, ccEmail: msg.ccEmails, sentAt: msg.sentAt },
          dealId,
          contactEmail || null,
          { contactId: preResolvedContactId, userEmail }
        );
        linked++;
      } catch (err) {
        console.warn(`[deal/link-email] Failed to log message "${msg.subject}":`, err.message);
      }
    }

    // Auto-associate external participants from the thread as contacts on the deal.
    // Filters out: internal domain emails, noreply addresses, and already-linked contacts.
    // Only links contacts whose domain matches an existing deal contact (safety net).
    const allParticipantEmails = new Set();
    for (const msg of messages) {
      if (msg.fromEmail) allParticipantEmails.add(msg.fromEmail.toLowerCase());
      for (const addr of (msg.toEmails || "").split(";")) {
        if (addr.trim()) allParticipantEmails.add(addr.trim().toLowerCase());
      }
      for (const addr of (msg.ccEmails || "").split(";")) {
        if (addr.trim()) allParticipantEmails.add(addr.trim().toLowerCase());
      }
    }
    if (userEmail) allParticipantEmails.delete(userEmail.toLowerCase());

    // Filter out automated/system email addresses that aren't real people
    const NON_PERSON_PREFIXES = ["noreply", "no-reply", "no_reply", "donotreply", "do-not-reply", "notifications", "notification", "mailer-daemon", "postmaster", "calendar", "bounce", "auto", "system", "admin@", "info@", "support@", "help@", "sales@"];
    for (const email of [...allParticipantEmails]) {
      const local = email.split("@")[0];
      if (NON_PERSON_PREFIXES.some(p => local === p || local.startsWith(p + "+"))) {
        allParticipantEmails.delete(email);
      }
    }
    if (userEmail) {
      const userDomain = userEmail.split("@")[1];
      if (userDomain) {
        for (const email of [...allParticipantEmails]) {
          if (email.endsWith("@" + userDomain)) allParticipantEmails.delete(email);
        }
      }
    }

    if (allParticipantEmails.size > 0) {
      let existingContactEmails = new Set();
      const existingContactDomains = new Set();
      if (dealContactsResult.status === "fulfilled") {
        for (const c of dealContactsResult.value) {
          existingContactEmails.add(c.email);
          const domain = c.email.split("@")[1]?.toLowerCase();
          if (domain) existingContactDomains.add(domain);
        }
      }

      // Only auto-link contacts whose domain matches an existing deal contact.
      // This prevents unrelated participants (e.g. CC'd from a different company)
      // from being added to the deal and triggering HubSpot auto-associations.
      const contactsToLink = [];
      const skippedDomains = [];
      for (const email of allParticipantEmails) {
        if (existingContactEmails.has(email)) continue; // already on deal
        const domain = email.split("@")[1]?.toLowerCase();
        if (existingContactDomains.size > 0 && domain && !existingContactDomains.has(domain)) {
          skippedDomains.push(email);
          continue;
        }
        contactsToLink.push(email);
      }
      if (skippedDomains.length > 0) {
        console.log(`[deal/link-email] Skipped ${skippedDomains.length} participant(s) from non-matching domains: ${skippedDomains.join(", ")}`);
      }

      let contactsLinked = 0;
      for (const email of contactsToLink) {
        try {
          await hubspot.linkContactToDeal(email, dealId);
          contactsLinked++;
        } catch (err) {
          console.warn(`[deal/link-email] Contact link failed for ${email}:`, err.message);
        }
      }
      if (contactsLinked > 0) console.log(`[deal/link-email] Linked ${contactsLinked} participant contact(s) to deal ${dealId}`);
    }

    console.log(`[deal/link-email] COMPLETED: ${linked}/${toLink.length} emails → deal ${dealId}`);
  } catch (err) {
    console.error(`[deal/link-email] FAILED:`, err.message);
    handleError(res, err, "POST /api/hubspot/deal/link-email");
  }
});

// PATCH /api/hubspot/deal/:id/stage
// Moves a deal to a different pipeline stage (e.g. "Proposal" to "Closed Won").
// Body: { stageId }
// Also invalidates any cached /api/context entries that reference this deal.
app.patch("/api/hubspot/deal/:id/stage", requireToken, rateLimit, async (req, res) => {
  const { id } = req.params;
  const { stageId } = req.body || {};
  if (!id || !/^\d+$/.test(id))    return res.status(400).json({ error: "Invalid deal id" });
  if (!stageId || typeof stageId !== "string") return res.status(400).json({ error: "stageId is required" });
  if (stageId.length > 64) return res.status(400).json({ error: "stageId too long" });
  try {
    const result = await hubspot.moveDealStage(id, stageId);
    // Invalidate only cached /api/context entries that contain this deal
    for (const [key, entry] of contextCache) {
      const deals = entry?.data?.deals;
      if (Array.isArray(deals) && deals.some((d) => d && d.id === id)) {
        contextCache.delete(key);
      }
    }
    res.json({ success: true, ...result });
  } catch (err) {
    handleError(res, err, "PATCH /api/hubspot/deal/:id/stage");
  }
});

// POST /api/hubspot/contact
// Creates a new contact in HubSpot. Optionally associates with a company.
// Body: { email, firstname, lastname, company, jobtitle, companyId }
app.post("/api/hubspot/contact", requireToken, rateLimit, async (req, res) => {
  const { email, firstname, lastname, company, jobtitle, companyId } = req.body;
  if (!email) return res.status(400).json({ error: "email is required" });
  try {
    const result = await hubspot.createContact({ email, firstname, lastname, company, jobtitle, companyId });
    res.json({ success: true, ...result });
  } catch (err) {
    handleError(res, err, "POST /api/hubspot/contact");
  }
});

// GET /api/hubspot/companies/search?q=...
// Searches HubSpot companies by name. Used for company autocomplete.
app.get("/api/hubspot/companies/search", requireToken, rateLimit, async (req, res) => {
  const q = (req.query.q || "").trim();
  try {
    const companies = await hubspot.searchCompanies(q);
    res.json(companies);
  } catch (err) {
    handleError(res, err, "GET /api/hubspot/companies/search");
  }
});

// POST /api/hubspot/company
// Creates a new company in HubSpot.
// Body: { name, domain } - at least one is required.
app.post("/api/hubspot/company", requireToken, rateLimit, async (req, res) => {
  const { name, domain } = req.body;
  if (!domain && !name) return res.status(400).json({ error: "name or domain is required" });
  try {
    const result = await hubspot.createCompany({ name, domain });
    res.json({ success: true, ...result });
  } catch (err) {
    handleError(res, err, "POST /api/hubspot/company");
  }
});

// POST /api/suggest-response
// Reads an email thread and uses Claude to suggest what the user should reply.
// Body: { conversationId, recipientName?, company? }
// Returns: { suggestion } - a plain-text suggestion (not a full draft).
app.post("/api/suggest-response", requireToken, rateLimit, async (req, res) => {
  const { conversationId, recipientName, company } = req.body;
  if (!conversationId) return res.status(400).json({ error: "conversationId is required" });
  try {
    const thread = await graph.getEmailThread(req.token, conversationId);
    const suggestion = await claude.suggestResponse({ thread, recipientName, company });
    res.json({ suggestion });
  } catch (err) {
    handleError(res, err, "POST /api/suggest-response");
  }
});

// =============================================================================
// AUTO-TRACK ROUTES
// Auto-tracking creates a Microsoft Graph webhook subscription on the user's
// mailbox. When new emails arrive, Graph notifies our /api/webhooks/graph
// endpoint, which then automatically links matching emails to HubSpot deals.
// =============================================================================
const subscriptions = require("./subscriptions"); // Manages Graph webhook subscription lifecycle

// POST /api/auto-track/enable
// Creates a Graph webhook subscription for the user's mailbox.
// The server stores the refresh token so it can renew the subscription
// and process webhooks even when the user isn't actively using the add-in.
app.post("/api/auto-track/enable", requireToken, rateLimit, async (req, res) => {
  try {
    const userId = req.tokenPayload?.oid;
    if (!userId) return res.status(400).json({ error: "Could not determine user identity" });

    // Server already has the refresh token from auth/exchange.
    // No need for client to send it — server is the exclusive RT owner.
    const user = tokenStore.getUser(userId);
    if (!user?.refreshToken) {
      return res.status(400).json({ error: "Please sign out and sign in again to enable auto-tracking" });
    }

    // Cache the live access token for immediate webhook processing (~1h).
    // After it expires, autoTracker falls back to the server-side refresh token.
    const accessToken = req.token;
    tokenStore.saveUser(userId, {
      accessToken,
      accessTokenExpiresAt: new Date(Date.now() + 55 * 60 * 1000).toISOString(),
    });

    // Delete existing subscription if any (clean slate)
    const existing = subscriptions.getSubscription(userId);
    if (existing) {
      await subscriptions.deleteSubscription(userId, accessToken).catch(() => {});
    }

    await subscriptions.createSubscription(userId, accessToken);
    res.json({ enabled: true });
  } catch (err) {
    handleError(res, err, "POST /api/auto-track/enable");
  }
});

// POST /api/auto-track/disable
// Deletes the Graph webhook subscription and clears the local record.
// Uses the refresh token to get a fresh access token for the Graph API call.
app.post("/api/auto-track/disable", requireToken, rateLimit, async (req, res) => {
  try {
    const userId = req.tokenPayload?.oid;
    if (!userId) return res.status(400).json({ error: "Could not determine user identity" });

    const user = tokenStore.getUser(userId);
    if (user?.refreshToken) {
      try {
        const tokens = await graph.refreshAccessToken(user.refreshToken);
        await subscriptions.deleteSubscription(userId, tokens.accessToken);
      } catch (err) {
        console.warn("[auto-track/disable] Could not delete subscription via API:", err.message);
      }
    }
    // Also remove locally even if API call failed
    const entry = subscriptions.getSubscription(userId);
    if (entry) {
      await subscriptions.deleteSubscription(userId, "").catch(() => {});
    }
    res.json({ enabled: false });
  } catch (err) {
    handleError(res, err, "POST /api/auto-track/disable");
  }
});

// GET /api/auto-track/status
// Checks whether auto-tracking is currently active for the user.
// Returns: { enabled: boolean }
app.get("/api/auto-track/status", requireToken, rateLimit, async (req, res) => {
  try {
    const userId = req.tokenPayload?.oid;
    if (!userId) return res.json({ enabled: false });

    const sub = subscriptions.getSubscription(userId);
    const enabled = !!(sub && new Date(sub.expiresAt).getTime() > Date.now());
    res.json({ enabled });
  } catch (err) {
    handleError(res, err, "GET /api/auto-track/status");
  }
});

// Graph subscriptions expire after a max of 3 days. This timer renews any
// subscriptions expiring within 12 hours, running every 6 hours.
const _renewalInterval = setInterval(() => {
  subscriptions.renewAllDue().catch(err => console.error("[subscriptions] renewAllDue error:", err.message));
}, 6 * 60 * 60 * 1000);
// Also renew 30s after startup
setTimeout(() => {
  subscriptions.renewAllDue().catch(err => console.error("[subscriptions] startup renewAllDue error:", err.message));
}, 30000);

// ── Health Check ──────────────────────────────────────────────
// Simple uptime probe used by monitoring tools and nginx health checks.
// No auth required - just confirms the server is running.
app.get("/health", (req, res) => {
  res.json({ status: "ok", timestamp: new Date().toISOString() });
});

// ── Server Start ──────────────────────────────────────────────
// Production: plain HTTP behind nginx (nginx handles SSL termination).
// Development: HTTPS with local dev certificates (required by Office Add-ins).
let server;
if (process.env.NODE_ENV === "production") {
  const http = require("http");
  server = http.createServer(app).listen(PORT, () => {
    console.log(`Imdath server running on port ${PORT}`);
  });
} else {
  try {
    const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
    const sslOptions = {
      key:  fs.readFileSync(path.join(certDir, "localhost.key")),
      cert: fs.readFileSync(path.join(certDir, "localhost.crt")),
      ca:   fs.readFileSync(path.join(certDir, "ca.crt")),
    };
    server = https.createServer(sslOptions, app).listen(PORT, () => {
      console.log(`Imdath server running on https://localhost:${PORT}`);
    });
  } catch (err) {
    console.error("[server] SSL certificates not found. Run: npm run install-certs");
    console.error("[server]", err.message);
    process.exit(1);
  }
}

// Set socket timeouts to mitigate slowloris attacks (where an attacker holds
// connections open indefinitely by sending data very slowly).
// headersTimeout must be greater than keepAliveTimeout per Node.js docs.
if (server) {
  server.setTimeout(30000);       // Max time for a single request
  server.keepAliveTimeout = 65000; // How long to keep idle connections open
  server.headersTimeout = 66000;   // Max time to receive complete request headers
}

// ── Graceful Shutdown ─────────────────────────────────────────
// On SIGTERM/SIGINT: stop accepting new connections, clear all timers,
// wait for in-flight requests to finish, then exit. Forces exit after 10s
// if connections don't close (prevents zombie processes).
function gracefulShutdown(signal) {
  console.log(`[server] ${signal} received — shutting down gracefully`);
  clearInterval(rateLimitCleanupInterval);
  clearInterval(contextCacheCleanupInterval);
  clearInterval(_renewalInterval);
  if (server) {
    server.close(() => {
      console.log("[server] All connections closed — exiting");
      process.exit(0);
    });
    // Force exit if connections don't close within 10 seconds
    setTimeout(() => {
      console.warn("[server] Forceful shutdown after timeout");
      process.exit(1);
    }, 10000).unref();
  } else {
    process.exit(0);
  }
}

process.on("SIGTERM", () => gracefulShutdown("SIGTERM"));
process.on("SIGINT",  () => gracefulShutdown("SIGINT"));

// ── Process Error Handlers ────────────────────────────────────
// Unhandled promise rejections: log but keep running (non-fatal).
// Uncaught exceptions: log and crash - the process is in an unknown state
// and PM2 will restart it automatically.
process.on("unhandledRejection", (reason, promise) => {
  console.error("[server] Unhandled promise rejection:", reason);
});

process.on("uncaughtException", (err) => {
  console.error("[server] Uncaught exception — shutting down:", err);
  process.exit(1);
});
