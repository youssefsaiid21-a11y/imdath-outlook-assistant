/**
 * subscriptions.js - Microsoft Graph webhook subscription management.
 *
 * Graph API sends push notifications when new emails arrive, but subscriptions
 * expire after a maximum of 3 days. This module handles:
 * - Creating new webhook subscriptions for users who enable auto-tracking
 * - Renewing subscriptions before they expire (checked periodically)
 * - Deleting subscriptions when auto-tracking is disabled
 * - Mapping subscription IDs back to user IDs (for webhook processing)
 *
 * Subscriptions are persisted to disk so they survive server restarts.
 */
"use strict";

const fs         = require("fs");
const path       = require("path");
const fetch      = require("node-fetch");
const tokenStore = require("./tokenStore");
const graph      = require("./graph");

const GRAPH_BASE   = "https://graph.microsoft.com/v1.0";
const DATA_DIR     = process.env.DATA_DIR || path.join(__dirname, "../.data");
const SUBS_FILE    = path.join(DATA_DIR, "subscriptions.json");
const TIMEOUT_MS   = 15000;
const MAX_EXPIRY_DAYS = 3; // Graph API maximum for mail subscription lifetime

// In-memory: userId -> { subscriptionId, expiresAt (ISO string) }
let _store = new Map();

// ── Persistence ──────────────────────────────────────────────
function ensureDataDir() {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true, mode: 0o700 });
  }
}

function load() {
  try {
    ensureDataDir();
    if (fs.existsSync(SUBS_FILE)) {
      const parsed = JSON.parse(fs.readFileSync(SUBS_FILE, "utf8"));
      _store = new Map(Object.entries(parsed));
      console.log(`[subscriptions] Loaded ${_store.size} subscription(s)`);
    }
  } catch (err) {
    console.error("[subscriptions] Failed to load:", err.message);
  }
}

function persist() {
  try {
    ensureDataDir();
    const tmpFile = SUBS_FILE + ".tmp";
    fs.writeFileSync(tmpFile, JSON.stringify(Object.fromEntries(_store), null, 2), { mode: 0o600 });
    fs.renameSync(tmpFile, SUBS_FILE);
  } catch (err) {
    console.error("[subscriptions] Failed to persist:", err.message);
  }
}

// ── Graph API Helpers ────────────────────────────────────────

// Simplified Graph API fetch for subscription CRUD operations.
// Separate from graph.js because subscriptions use different endpoints and error handling.
async function graphFetch(url, token, options = {}) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);
  try {
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      signal: controller.signal,
      ...options,
    });
    const text = await res.text();
    if (!res.ok) throw new Error(`Graph ${options.method || "GET"} ${res.status}: ${text.substring(0, 300)}`);
    return text ? JSON.parse(text) : {};
  } finally {
    clearTimeout(timer);
  }
}

// ── Public API ───────────────────────────────────────────────

// Creates a Graph webhook subscription that notifies us when new emails arrive.
// The clientState secret is used to validate incoming webhook notifications.
async function createSubscription(userId, accessToken) {
  const webhookSecret = process.env.GRAPH_WEBHOOK_SECRET;
  const serverUrl     = process.env.SERVER_URL;
  if (!webhookSecret) throw new Error("GRAPH_WEBHOOK_SECRET not configured");
  if (!serverUrl)     throw new Error("SERVER_URL not configured");

  const expiresAt = new Date(Date.now() + MAX_EXPIRY_DAYS * 24 * 60 * 60 * 1000 - 60000);

  const data = await graphFetch(`${GRAPH_BASE}/subscriptions`, accessToken, {
    method: "POST",
    body: JSON.stringify({
      changeType:         "created",
      resource:           "/me/messages",
      notificationUrl:    serverUrl + "/api/webhooks/graph",
      expirationDateTime: expiresAt.toISOString(),
      clientState:        webhookSecret,
    }),
  });

  _store.set(userId, {
    subscriptionId: data.id,
    expiresAt:      data.expirationDateTime || expiresAt.toISOString(),
  });
  persist();

  console.log(`[subscriptions] Created subscription ${data.id} for user ${userId}, expires ${data.expirationDateTime}`);
  return data;
}

// Extends the expiration of an existing subscription by another 3 days.
async function renewSubscription(userId, accessToken) {
  const entry = _store.get(userId);
  if (!entry) throw new Error(`No subscription for user ${userId}`);

  const expiresAt = new Date(Date.now() + MAX_EXPIRY_DAYS * 24 * 60 * 60 * 1000 - 60000);

  const data = await graphFetch(`${GRAPH_BASE}/subscriptions/${entry.subscriptionId}`, accessToken, {
    method: "PATCH",
    body: JSON.stringify({ expirationDateTime: expiresAt.toISOString() }),
  });

  entry.expiresAt = data.expirationDateTime || expiresAt.toISOString();
  persist();

  console.log(`[subscriptions] Renewed subscription ${entry.subscriptionId} for user ${userId}, expires ${entry.expiresAt}`);
  return data;
}

// Deletes a subscription from Graph API and removes it from local storage.
// Silently ignores 404 (already deleted on Graph's side).
async function deleteSubscription(userId, accessToken) {
  const entry = _store.get(userId);
  if (!entry) return;

  try {
    await graphFetch(`${GRAPH_BASE}/subscriptions/${entry.subscriptionId}`, accessToken, {
      method: "DELETE",
    });
  } catch (err) {
    // 404 = already deleted, that's fine
    if (!err.message.includes("404")) {
      console.warn(`[subscriptions] Delete failed for user ${userId}:`, err.message);
    }
  }

  _store.delete(userId);
  persist();
  console.log(`[subscriptions] Deleted subscription for user ${userId}`);
}

// Returns the subscription entry for a user, or null if none exists.
function getSubscription(userId) {
  return _store.get(userId) || null;
}

// Reverse lookup: finds which user owns a given subscription ID.
// Called by autoTracker when a webhook notification arrives.
function getUserBySubscriptionId(subscriptionId) {
  for (const [userId, entry] of _store) {
    if (entry.subscriptionId === subscriptionId) return userId;
  }
  return null;
}

// Called periodically (e.g. every few hours) to renew subscriptions expiring
// within 12 hours. Refreshes the user's OAuth token and extends the subscription.
// Removes stale subscriptions if the user's refresh token is invalid.
async function renewAllDue() {
  const now = Date.now();
  const renewThreshold = 12 * 60 * 60 * 1000; // 12 hours

  for (const [userId, entry] of _store) {
    const expiresAt = new Date(entry.expiresAt).getTime();
    if (expiresAt - now > renewThreshold) continue; // not due yet

    console.log(`[subscriptions] Renewing subscription for user ${userId} (expires ${entry.expiresAt})`);
    try {
      const user = tokenStore.getUser(userId);
      if (!user?.refreshToken) {
        console.warn(`[subscriptions] No refresh token for user ${userId} — removing subscription`);
        _store.delete(userId);
        persist();
        continue;
      }

      const tokens = await graph.refreshAccessToken(user.refreshToken);
      // Update stored refresh token in case it was rotated
      tokenStore.saveUser(userId, { refreshToken: tokens.refreshToken, email: user.email, name: user.name });

      await renewSubscription(userId, tokens.accessToken);
    } catch (err) {
      console.error(`[subscriptions] Renewal failed for user ${userId}:`, err.message);
      // If token is invalid, remove the subscription
      if (err.message.includes("invalid_grant") || err.message.includes("AADSTS")) {
        console.warn(`[subscriptions] Removing stale subscription for user ${userId}`);
        _store.delete(userId);
        persist();
      }
    }
  }
}

// Load on require
load();

module.exports = { createSubscription, renewSubscription, deleteSubscription, getSubscription, getUserBySubscriptionId, renewAllDue };
