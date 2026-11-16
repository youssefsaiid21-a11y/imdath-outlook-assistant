/**
 * tokenStore.js - Secure OAuth token storage for authenticated users.
 *
 * Splits token storage into two layers for security:
 * - Persisted (disk): refresh tokens, email, name - survives restarts
 * - Volatile (memory): access tokens - lost on restart (by design)
 *
 * Access tokens are short-lived (~1h) and can always be regenerated from
 * the refresh token, so writing them to disk is an unnecessary security risk.
 * The persisted store uses atomic writes (tmp + rename) to prevent corruption.
 */
"use strict";

const fs   = require("fs");
const path = require("path");

const DATA_DIR   = process.env.DATA_DIR || path.join(__dirname, "../.data");
const TOKEN_FILE = path.join(DATA_DIR, "tokens.json");

// Persisted store: userId -> { refreshToken, email, name, updatedAt }
// Written to disk - NEVER contains access tokens.
let _persisted = new Map();

// Memory-only store: userId -> { accessToken, accessTokenExpiresAt }
// Never written to disk - lost on restart (by design).
const _volatile = new Map();

// ── Initialization ───────────────────────────────────────────
function ensureDataDir() {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true, mode: 0o700 });
  }
}

// Loads persisted tokens from disk on startup. Strips any access token
// fields that may have been written by older versions of this module.
function load() {
  try {
    ensureDataDir();
    if (fs.existsSync(TOKEN_FILE)) {
      const raw = fs.readFileSync(TOKEN_FILE, "utf8");
      const parsed = JSON.parse(raw);
      // Strip any access token fields that may exist from older versions
      for (const [key, val] of Object.entries(parsed)) {
        delete val.accessToken;
        delete val.accessTokenExpiresAt;
      }
      _persisted = new Map(Object.entries(parsed));
      console.log(`[tokenStore] Loaded ${_persisted.size} user(s)`);
    }
  } catch (err) {
    console.error("[tokenStore] Failed to load:", err.message);
  }
}

function persist() {
  try {
    ensureDataDir();
    // Build a clean object that never includes access token fields
    const obj = {};
    for (const [key, val] of _persisted) {
      obj[key] = {
        refreshToken: val.refreshToken || "",
        email:        val.email || "",
        name:         val.name || "",
        updatedAt:    val.updatedAt || "",
      };
    }
    const tmpFile = TOKEN_FILE + ".tmp";
    fs.writeFileSync(tmpFile, JSON.stringify(obj, null, 2), { mode: 0o600 });
    fs.renameSync(tmpFile, TOKEN_FILE);
  } catch (err) {
    console.error("[tokenStore] Failed to persist:", err.message);
  }
}

// ── Public API ───────────────────────────────────────────────

// Saves or updates a user's tokens. Refresh token goes to disk (persisted),
// access token stays in memory only (volatile). Merges with existing data
// so callers can update individual fields without losing the rest.
function saveUser(userId, { refreshToken, email, name, accessToken, accessTokenExpiresAt }) {
  if (!userId) return;

  // Update persisted fields
  const existing = _persisted.get(userId) || {};
  _persisted.set(userId, {
    refreshToken: refreshToken || existing.refreshToken || "",
    email:        email || existing.email || "",
    name:         name  || existing.name  || "",
    updatedAt:    new Date().toISOString(),
  });
  persist();

  // Update volatile (memory-only) fields
  if (accessToken || accessTokenExpiresAt) {
    const existingVol = _volatile.get(userId) || {};
    _volatile.set(userId, {
      accessToken:          accessToken || existingVol.accessToken || "",
      accessTokenExpiresAt: accessTokenExpiresAt || existingVol.accessTokenExpiresAt || "",
    });
  }
}

// Returns a merged view of persisted + volatile data for a user.
// Callers see a single object with both refresh and access token fields.
function getUser(userId) {
  const persisted = _persisted.get(userId);
  if (!persisted) return null;
  const vol = _volatile.get(userId) || {};
  // Merge both stores — callers see the same shape as before
  return {
    ...persisted,
    accessToken:          vol.accessToken || "",
    accessTokenExpiresAt: vol.accessTokenExpiresAt || "",
  };
}

// Returns all users with merged persisted + volatile data.
function getAllUsers() {
  return Array.from(_persisted.entries()).map(([id, data]) => {
    const vol = _volatile.get(id) || {};
    return { id, ...data, accessToken: vol.accessToken || "", accessTokenExpiresAt: vol.accessTokenExpiresAt || "" };
  });
}

// Removes a user from both stores and persists the change.
function deleteUser(userId) {
  const deleted = _persisted.delete(userId);
  _volatile.delete(userId);
  if (deleted) persist();
  return deleted;
}

// Load on require
load();

module.exports = { saveUser, getUser, getAllUsers, deleteUser };
