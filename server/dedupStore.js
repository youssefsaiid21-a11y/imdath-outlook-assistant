/**
 * dedupStore.js - Deduplication and conversation-to-deal mapping store.
 *
 * Manages three persistent data stores:
 * 1. Per-user dedup: tracks which emails have been processed (prevents re-linking)
 * 2. Global deal links: tracks which email+deal pairs exist (prevents cross-user duplicates)
 * 3. Conversation-deal map: maps Graph conversationIds to HubSpot deal IDs
 *    (set when user manually links an email thread to a deal via Sphinx UI)
 *
 * All stores are backed by JSON files in the .data directory. Writes use
 * atomic rename (write to .tmp, then rename) to prevent corruption on crash.
 * Entries are pruned by age (30 days) and size caps to prevent unbounded growth.
 */
"use strict";

const fs   = require("fs");
const path = require("path");

const DATA_DIR         = process.env.DATA_DIR || path.join(__dirname, "../.data");
const DEDUP_FILE       = path.join(DATA_DIR, "linked-emails.json");
const DEAL_LINK_FILE   = path.join(DATA_DIR, "deal-links.json");
const CONV_MAP_FILE    = path.join(DATA_DIR, "conversation-deals.json");
const MAX_PER_USER     = 10000;   // Max tracked emails per user
const MAX_GLOBAL       = 50000;   // Max global deal-link entries
const MAX_CONV_MAP     = 10000;   // Max conversation-deal mappings
const MAX_AGE_MS       = 30 * 24 * 60 * 60 * 1000; // 30 days

// Per-user: userId -> Map<internetMessageId, timestamp>
let _store = new Map();
// Global: "msgId:dealId" -> timestamp (prevents cross-user duplicates)
let _dealLinks = new Map();
// Conversation -> deal mapping: conversationId -> { dealId, timestamp }
let _convMap = new Map();

function ensureDataDir() {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true, mode: 0o700 });
  }
}

function load() {
  try {
    ensureDataDir();
    if (fs.existsSync(DEDUP_FILE)) {
      const parsed = JSON.parse(fs.readFileSync(DEDUP_FILE, "utf8"));
      for (const [userId, entries] of Object.entries(parsed)) {
        _store.set(userId, new Map(Object.entries(entries)));
      }
      console.log(`[dedupStore] Loaded per-user dedup data for ${_store.size} user(s)`);
    }
    if (fs.existsSync(DEAL_LINK_FILE)) {
      const parsed = JSON.parse(fs.readFileSync(DEAL_LINK_FILE, "utf8"));
      _dealLinks = new Map(Object.entries(parsed));
      console.log(`[dedupStore] Loaded ${_dealLinks.size} global deal-link(s)`);
    }
    if (fs.existsSync(CONV_MAP_FILE)) {
      const parsed = JSON.parse(fs.readFileSync(CONV_MAP_FILE, "utf8"));
      // Normalize all keys on load so old entries (pre-normalization) are findable
      for (const [key, val] of Object.entries(parsed)) {
        _convMap.set(normalizeConvId(key), val);
      }
      console.log(`[dedupStore] Loaded ${_convMap.size} conversation-deal mapping(s)`);
    }
  } catch (err) {
    console.error("[dedupStore] Failed to load:", err.message);
  }
}

// Saves per-user dedup store to disk using atomic write (tmp + rename).
function persist() {
  try {
    ensureDataDir();
    const obj = {};
    for (const [userId, entries] of _store) {
      obj[userId] = Object.fromEntries(entries);
    }
    const tmpFile = DEDUP_FILE + ".tmp";
    fs.writeFileSync(tmpFile, JSON.stringify(obj, null, 2), { mode: 0o600 });
    fs.renameSync(tmpFile, DEDUP_FILE);
  } catch (err) {
    console.error("[dedupStore] Failed to persist per-user store:", err.message);
  }
}

function persistDealLinks() {
  try {
    ensureDataDir();
    const tmpFile = DEAL_LINK_FILE + ".tmp";
    fs.writeFileSync(tmpFile, JSON.stringify(Object.fromEntries(_dealLinks), null, 2), { mode: 0o600 });
    fs.renameSync(tmpFile, DEAL_LINK_FILE);
  } catch (err) {
    console.error("[dedupStore] Failed to persist deal-links:", err.message);
  }
}

// Removes expired entries (older than 30 days) and enforces size cap per user.
function prune(userId) {
  const entries = _store.get(userId);
  if (!entries) return;
  const now = Date.now();
  for (const [msgId, ts] of entries) {
    if (now - ts > MAX_AGE_MS) entries.delete(msgId);
  }
  // Cap size — remove oldest entries
  if (entries.size > MAX_PER_USER) {
    const sorted = [...entries.entries()].sort((a, b) => a[1] - b[1]);
    const toRemove = sorted.slice(0, entries.size - MAX_PER_USER);
    for (const [msgId] of toRemove) entries.delete(msgId);
  }
}

// Checks if a message was already processed for a given user.
function hasLinked(userId, messageId) {
  const entries = _store.get(userId);
  return entries ? entries.has(messageId) : false;
}

// Records that a message was processed. Triggers prune + persist.
function markLinked(userId, messageId) {
  if (!_store.has(userId)) _store.set(userId, new Map());
  const entries = _store.get(userId);
  entries.set(messageId, Date.now());
  prune(userId);
  persist();
}

// ── Global deal-link dedup (cross-user) ──────────────────────

// Checks if a specific email has already been linked to a specific deal.
// Uses a composite key "messageId:dealId" to track unique pairings.
function hasDealLink(messageId, dealId) {
  return _dealLinks.has(messageId + ":" + dealId);
}

function markDealLink(messageId, dealId) {
  _dealLinks.set(messageId + ":" + dealId, Date.now());
  // Prune old entries
  if (_dealLinks.size > MAX_GLOBAL) {
    const now = Date.now();
    for (const [key, ts] of _dealLinks) {
      if (now - ts > MAX_AGE_MS) _dealLinks.delete(key);
    }
    // If still over limit, remove oldest
    if (_dealLinks.size > MAX_GLOBAL) {
      const sorted = [..._dealLinks.entries()].sort((a, b) => a[1] - b[1]);
      const toRemove = sorted.slice(0, _dealLinks.size - MAX_GLOBAL);
      for (const [key] of toRemove) _dealLinks.delete(key);
    }
  }
  persistDealLinks();
}

// ── Conversation → Deal mapping (for threads linked manually) ──
function persistConvMap() {
  try {
    ensureDataDir();
    const tmpFile = CONV_MAP_FILE + ".tmp";
    fs.writeFileSync(tmpFile, JSON.stringify(Object.fromEntries(_convMap), null, 2), { mode: 0o600 });
    fs.renameSync(tmpFile, CONV_MAP_FILE);
  } catch (err) {
    console.error("[dedupStore] Failed to persist conversation map:", err.message);
  }
}

// Normalize base64 vs base64url encoding differences in conversationIds.
// Graph API inconsistently uses + vs -, / vs _, and trailing = padding
// depending on which endpoint returns the ID. This ensures lookups match.
function normalizeConvId(id) {
  if (!id) return "";
  return id.replace(/\-/g, "/").replace(/_/g, "+").replace(/=+$/g, "");
}

// Records that a conversation thread is linked to a deal. Called when the user
// manually links an email to a deal in the UI. Conversation mappings never expire
// by age (only by size cap) because deals can go quiet for months between emails.
function setConversationDeal(conversationId, dealId, subject) {
  if (!conversationId || !dealId) return;
  _convMap.set(normalizeConvId(conversationId), { dealId, subject: subject || "", timestamp: Date.now() });
  // Only prune by size — conversation-deal mappings represent explicit user intent
  // and must not expire by age (deals can go quiet for months between emails).
  if (_convMap.size > MAX_CONV_MAP) {
    const sorted = [..._convMap.entries()].sort((a, b) => a[1].timestamp - b[1].timestamp);
    const toRemove = sorted.slice(0, _convMap.size - MAX_CONV_MAP);
    for (const [key] of toRemove) _convMap.delete(key);
  }
  persistConvMap();
}

// Returns the deal ID linked to a conversation, or null if not mapped.
function getConversationDeal(conversationId) {
  if (!conversationId) return null;
  const key = normalizeConvId(conversationId);
  const entry = _convMap.get(key);
  if (!entry) return null;
  return entry.dealId;
}

// Returns the full mapping entry (dealId, subject, timestamp) or null.
function getConversationDealEntry(conversationId) {
  if (!conversationId) return null;
  const key = normalizeConvId(conversationId);
  return _convMap.get(key) || null;
}

// Reverse lookup: finds the conversationId linked to a given deal.
// Used by follow-up feature to find the email thread for a deal.
function getConversationIdForDeal(dealId) {
  if (!dealId) return null;
  const target = String(dealId);
  for (const [convId, entry] of _convMap) {
    if (String(entry.dealId) === target) return convId;
  }
  return null;
}

// Load on require
load();

module.exports = { hasLinked, markLinked, hasDealLink, markDealLink, setConversationDeal, getConversationDeal, getConversationDealEntry, getConversationIdForDeal };
