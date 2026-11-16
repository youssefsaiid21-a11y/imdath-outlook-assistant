/**
 * graph.js - Microsoft Graph API client for Outlook email operations.
 *
 * Handles all interactions with the Microsoft Graph API including:
 * - Reading and searching emails
 * - Sending emails and creating drafts (new or reply)
 * - Fetching email threads for CRM logging
 * - Email analytics (unread count, priority scoring, response time)
 * - OAuth token refresh for server-side operations
 * - Extracting email signatures from sent mail
 *
 * All Graph API calls use abort-controller timeouts to prevent hanging requests.
 */
"use strict";

const fetch = require("node-fetch");
const { extractStyleProfile } = require("./styleProfile");

const GRAPH_BASE       = "https://graph.microsoft.com/v1.0";
const TIMEOUT_MS       = parseInt(process.env.GRAPH_TIMEOUT_MS    || "15000", 10);
const SENT_MAIL_LIMIT  = parseInt(process.env.SENT_MAIL_LIMIT     || "25",    10);
const HISTORY_LIMIT    = parseInt(process.env.EMAIL_HISTORY_LIMIT || "10",    10);

/** Escape a value for use inside OData $search strings.
 *  Whitelist approach: strip anything that isn't alphanumeric, space, or common email chars. */
function escapeODataSearch(str) {
  if (!str) return "";
  return str.replace(/[^a-zA-Z0-9._%+@\- ]/g, "");
}

/** Strip reply/forward prefixes and normalise a subject for comparison.
 *  Handles Re:, Fwd:, FW:, AW: (German), SV: (Scandinavian), etc. */
function normalizeSubject(subj) {
  let s = (typeof subj === "string" ? subj : "").trim();
  // Strip gateway tags like [Ext], [External], [Secure], etc.
  s = s.replace(/\[(?:ext|external|extern|secure|encrypted|caution|spam)\]\s*/gi, "");
  // Strip calendar date/time suffix (everything after @ with date-like content)
  s = s.replace(/\s*@\s+.*$/i, "");
  // Loop to strip all prefixes — they can be nested in any order:
  // Re: Accepted: RE: FW: [Ext] Re: Subject
  const prefixRe = /^(re|fw|fwd|aw|sv|ref|rif|odp|ynt|accepted|declined|tentative|tentatively accepted|updated invitation|updated|cancelled invitation|cancelled|canceled)\s*:\s*/i;
  let prev;
  do { prev = s; s = s.replace(prefixRe, ""); } while (s !== prev);
  return s.trim().toLowerCase();
}

/** Escape a value for use inside OData $filter eq 'value' strings.
 *  Conversation IDs contain base64 chars (=, /, +) that must be preserved.
 *  The only dangerous char in OData single-quoted strings is the single quote itself. */
function escapeODataFilter(str) {
  if (!str) return "";
  return str.replace(/'/g, "''");
}

// ── Helpers ───────────────────────────────────────────────────

// POST wrapper with timeout - used for sending emails, creating drafts, etc.
async function graphPost(url, token, body) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);
  try {
    const res = await fetch(url, {
      method:  "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body:    JSON.stringify(body),
      signal:  controller.signal,
    });
    // 202 No Content is success for sendMail
    if (res.status === 202 || res.status === 200) return { ok: true };
    const text = await res.text();
    const err = new Error(`Graph API ${res.status}: ${text}`);
    err.status = res.status;
    throw err;
  } catch (err) {
    if (err.name === "AbortError") throw new Error(`Graph API request timed out after ${TIMEOUT_MS}ms`);
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

// GET wrapper with timeout - used for fetching emails, threads, analytics data
async function graphGet(url, token) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);
  try {
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
      signal: controller.signal,
    });
    if (!res.ok) {
      const body = await res.text();
      const err = new Error(`Graph API ${res.status}: ${body}`);
      err.status = res.status;
      throw err;
    }
    try {
      return await res.json();
    } catch (parseErr) {
      throw new Error(`Graph API response not valid JSON`);
    }
  } catch (err) {
    if (err.name === "AbortError") throw new Error(`Graph API request timed out after ${TIMEOUT_MS}ms`);
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

/** Sanitize HTML — remove scripts, styles, event handlers, and dangerous attributes */
function sanitizeEmailHtml(html) {
  if (!html) return "";
  return html
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "")
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
    .replace(/<iframe[^>]*>[\s\S]*?<\/iframe>/gi, "")
    .replace(/<object[^>]*>[\s\S]*?<\/object>/gi, "")
    .replace(/<embed[^>]*>/gi, "")
    .replace(/\son\w+\s*=\s*"[^"]*"/gi, "")   // inline handlers (double-quoted)
    .replace(/\son\w+\s*=\s*'[^']*'/gi, "")   // inline handlers (single-quoted)
    .replace(/\son\w+\s*=[^\s>]*/gi, "")       // inline handlers (unquoted)
    .replace(/javascript\s*:/gi, "")            // javascript: protocol
    .trim();
}

/** Strip HTML tags and decode named + numeric entities */
function stripHtml(html) {
  if (!html) return "";
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    // Named entities
    .replace(/&nbsp;/gi,   " ")
    .replace(/&amp;/gi,    "&")
    .replace(/&lt;/gi,     "<")
    .replace(/&gt;/gi,     ">")
    .replace(/&quot;/gi,   '"')
    .replace(/&apos;/gi,   "'")
    .replace(/&#39;/gi,    "'")
    .replace(/&mdash;/gi,  "—")
    .replace(/&ndash;/gi,  "–")
    .replace(/&lsquo;/gi,  "\u2018")
    .replace(/&rsquo;/gi,  "\u2019")
    .replace(/&ldquo;/gi,  "\u201C")
    .replace(/&rdquo;/gi,  "\u201D")
    // Decimal numeric entities (e.g. &#8220;)
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(parseInt(n, 10)))
    // Hex numeric entities (e.g. &#x201C;)
    .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/\r\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

// ── Sent Mail Style ────────────────────────────────────────────

// Fetches the user's recent sent emails and builds a writing style profile.
// Claude uses this profile to mimic the sender's voice when ghostwriting.
async function getSentMailStyle(token) {
  const url =
    `${GRAPH_BASE}/me/mailFolders/SentItems/messages` +
    `?$top=${SENT_MAIL_LIMIT}&$select=subject,body,toRecipients&$orderby=sentDateTime desc`;

  const data = await graphGet(url, token);
  const messages = data.value || [];

  const emails = messages.map((m) => ({
    subject:    m.subject || "",
    bodyText:   stripHtml(m.body?.content || ""),
    recipients: (m.toRecipients || []).map((r) => r.emailAddress?.address).filter(Boolean),
  }));

  const profile = extractStyleProfile(emails);

  // Filter out likely auto-replies and empty messages (keep short genuine emails)
  profile.samples = emails
    .filter((e) => e.bodyText.length > 20)
    .slice(0, SENT_MAIL_LIMIT)
    .map((e) => e.bodyText.substring(0, 800));

  return profile;
}

// ── Email History ──────────────────────────────────────────────

// Fetches recent emails exchanged with a specific recipient.
// Used to give Claude context about prior conversations when drafting.
async function getEmailHistory(token, recipientEmail) {
  // Graph $search expects: $search="\"email\""  (KQL quoted phrase)
  // We encode the inner value but NOT the outer quotes which are part of OData syntax.
  const searchTerm = encodeURIComponent('"' + escapeODataSearch(recipientEmail) + '"');
  const searchUrl =
    `${GRAPH_BASE}/me/messages` +
    `?$search=${searchTerm}` +
    `&$top=${HISTORY_LIMIT}&$select=subject,receivedDateTime,bodyPreview,from&$orderby=receivedDateTime desc`;

  let data;
  try {
    data = await graphGet(searchUrl, token);
  } catch (originalErr) {
    // $search + $orderby can't be combined on some tenants — retry without $orderby
    console.warn("[graph] getEmailHistory: $search+$orderby unsupported, retrying:", originalErr.message);
    const fallbackUrl =
      `${GRAPH_BASE}/me/messages` +
      `?$search=${searchTerm}` +
      `&$top=${HISTORY_LIMIT}&$select=subject,receivedDateTime,bodyPreview,from`;
    try {
      data = await graphGet(fallbackUrl, token);
    } catch (fallbackErr) {
      console.error("[graph] getEmailHistory fallback failed:", fallbackErr.message);
      return [];
    }
  }

  return (data.value || []).map((m) => ({
    subject: m.subject || "(no subject)",
    date:    m.receivedDateTime ? m.receivedDateTime.substring(0, 10) : "",
    snippet: (m.bodyPreview || "").substring(0, 120),
    from:    m.from?.emailAddress?.address || "",
  }));
}

// ── Send Email ─────────────────────────────────────────────────

// Sends a new email (not a reply) via Graph API and saves to Sent Items.
async function sendEmail(token, { toEmail, subject, htmlBody }) {
  return graphPost(`${GRAPH_BASE}/me/sendMail`, token, {
    message: {
      subject: subject || "(no subject)",
      body: { contentType: "HTML", content: htmlBody || "" },
      toRecipients: [{ emailAddress: { address: toEmail } }],
    },
    saveToSentItems: true,
  });
}

// ── Email Analytics ────────────────────────────────────────────

// ── People Search ──────────────────────────────────────────────

// Searches the user's contacts/people graph for matching names or emails.
// Returns up to 8 results with name, email, title, and company.
async function searchPeople(token, query) {
  const q = encodeURIComponent('"' + escapeODataSearch(query.trim()) + '"');
  const url =
    `${GRAPH_BASE}/me/people?$search="${q}"&$top=8` +
    `&$select=displayName,emailAddresses,jobTitle,companyName`;
  const data = await graphGet(url, token);
  return (data.value || [])
    .filter(p => p.emailAddresses && p.emailAddresses.length > 0)
    .map(p => ({
      name:    p.displayName || "",
      email:   p.emailAddresses[0].address,
      title:   p.jobTitle      || "",
      company: p.companyName   || "",
    }));
}

// Computes email analytics: unread count, priority-scored urgent messages,
// and average response time. Fetches unread inbox, sent items, and full inbox
// in parallel to minimize latency.
async function getEmailAnalytics(token) {
  const [unreadResult, sentResult, inboxResult] = await Promise.allSettled([
    // Unread messages — for count + urgency + priority (includes webLink)
    graphGet(
      `${GRAPH_BASE}/me/mailFolders/Inbox/messages` +
      `?$filter=isRead eq false&$top=50` +
      `&$select=subject,from,receivedDateTime,importance,conversationId,webLink,id` +
      `&$orderby=receivedDateTime desc`,
      token
    ),
    // Recent sent items — for response time
    graphGet(
      `${GRAPH_BASE}/me/mailFolders/SentItems/messages` +
      `?$top=25&$select=subject,sentDateTime,conversationId&$orderby=sentDateTime desc`,
      token
    ),
    // Recent inbox (all, read+unread, larger window) — for response time matching
    graphGet(
      `${GRAPH_BASE}/me/mailFolders/Inbox/messages` +
      `?$top=100&$select=receivedDateTime,conversationId&$orderby=receivedDateTime desc`,
      token
    ),
  ]);

  const unreadMessages = unreadResult.status === "fulfilled" ? (unreadResult.value?.value || []) : [];
  const sentMessages   = sentResult.status   === "fulfilled" ? (sentResult.value?.value   || []) : [];
  const inboxMessages  = inboxResult.status  === "fulfilled" ? (inboxResult.value?.value  || []) : [];

  const now = Date.now();

  // Scores an unread email by urgency: high importance + age in hours.
  // Older unread emails get higher scores to surface neglected messages.
  function priorityScore(m) {
    const receivedTime = new Date(m.receivedDateTime).getTime();
    if (isNaN(receivedTime)) return 0; // skip invalid dates
    const ageH = (now - receivedTime) / 3600000;
    let score = 0;
    if (m.importance === "high") score += 60;
    if      (ageH > 72) score += 50;
    else if (ageH > 48) score += 35;
    else if (ageH > 24) score += 20;
    else if (ageH > 12) score += 10;
    else if (ageH > 4)  score += 5;
    return score;
  }

  // All unread with scores, sorted by priority descending
  const scoredUnread = unreadMessages.map(m => ({
    subject:    m.subject || "(no subject)",
    from:       m.from?.emailAddress?.address || "",
    fromName:   m.from?.emailAddress?.name    || "",
    receivedAt: m.receivedDateTime,
    importance: m.importance,
    ageHours:   Math.floor((now - new Date(m.receivedDateTime).getTime()) / 3600000),
    webLink:    m.webLink || null,
    score:      priorityScore(m),
  })).sort((a, b) => b.score - a.score);

  // Top 5 priority + general urgency list (10 items)
  const topPriority   = scoredUnread.slice(0, 5);
  const urgentMessages = scoredUnread
    .filter(m => m.importance === "high" || m.ageHours > 24)
    .slice(0, 10);

  // Avg response time: for each sent reply, find the most recent received message
  // in the same conversation that arrived BEFORE the reply was sent.
  // This measures how fast the user typically replies to incoming emails.
  let avgResponseTimeHours = null;
  if (sentMessages.length > 0 && inboxMessages.length > 0) {
    // Build a lookup: conversationId -> timestamp of the latest inbound message
    const convMap = new Map();
    for (const m of inboxMessages) {
      if (!m.conversationId || !m.receivedDateTime) continue;
      const t = new Date(m.receivedDateTime).getTime();
      const existing = convMap.get(m.conversationId);
      if (!existing || t > existing) convMap.set(m.conversationId, t);
    }

    const deltas = [];
    for (const sent of sentMessages) {
      if (!sent.conversationId || !sent.sentDateTime) continue;
      // Only count messages that are actual replies (have Re:/Sv:/etc prefix)
      const sub = (sent.subject || "").toLowerCase().trimStart();
      if (!/^(re|sv|aw|ref|rif|odp|ynt):\s*/i.test(sub)) continue;
      const receivedTime = convMap.get(sent.conversationId);
      if (!receivedTime) continue;
      const sentTime = new Date(sent.sentDateTime).getTime();
      // Only count if reply came after the received message (sanity check)
      if (sentTime <= receivedTime) continue;
      const delta = sentTime - receivedTime;
      // Cap at 7 days; exclude immediate bot-replies (< 30s)
      if (delta >= 30000 && delta < 7 * 24 * 3600 * 1000) deltas.push(delta);
    }

    // Require at least 2 data points for a meaningful average
    if (deltas.length >= 2) {
      avgResponseTimeHours = (deltas.reduce((a, b) => a + b, 0) / deltas.length) / 3600000;
    }
  }

  return {
    unreadCount:         unreadMessages.length,
    urgentCount:         urgentMessages.length,
    urgentMessages,
    topPriority,
    avgResponseTimeHours,
  };
}

// ── Get Email Thread ──────────────────────────────────────────
// Fetches messages in the same conversation as the given messageId,
// used for Reply Mode to provide thread context to Claude.
async function getEmailThread(token, conversationId) {
  // Escape for OData filter to prevent injection
  const safeConvId = escapeODataFilter(conversationId);
  const filter = `conversationId eq '${safeConvId}'`;
  const url =
    `${GRAPH_BASE}/me/messages` +
    `?$filter=${encodeURIComponent(filter)}` +
    `&$top=10&$select=subject,body,from,toRecipients,receivedDateTime` +
    `&$orderby=receivedDateTime desc`;

  let data;
  try {
    data = await graphGet(url, token);
  } catch (err) {
    // Some tenants don't support $filter + $orderby — retry without $orderby
    try {
      const fallbackUrl =
        `${GRAPH_BASE}/me/messages` +
        `?$filter=${encodeURIComponent(filter)}` +
        `&$top=10&$select=subject,body,from,toRecipients,receivedDateTime`;
      data = await graphGet(fallbackUrl, token);
    } catch (err2) {
      console.warn("[graph] getEmailThread failed (with fallback):", err2.message);
      return [];
    }
  }

  return (data.value || []).map((m) => ({
    subject: m.subject || "(no subject)",
    from:    m.from?.emailAddress?.name || m.from?.emailAddress?.address || "",
    date:    m.receivedDateTime ? m.receivedDateTime.substring(0, 10) : "",
    body:    stripHtml(m.body?.content || "").substring(0, 1200),
  }));
}

// ── Reply to Message ──────────────────────────────────────────
// Sends a reply (or reply-all) to an existing message via Graph API.
// Uses the /reply or /replyAll action which preserves the quoted thread
// and automatically addresses the correct recipients.
async function replyToMessage(token, messageId, htmlBody, replyAll = false) {
  const action = replyAll ? "replyAll" : "reply";
  const url = `${GRAPH_BASE}/me/messages/${encodeURIComponent(messageId)}/${action}`;
  return graphPost(url, token, {
    message: {
      body: { contentType: "HTML", content: htmlBody || "" },
    },
  });
}

// ── Create Reply-All Draft ───────────────────────────────────
// Creates a draft reply-all on a message and sets the body.
// Returns { draftId, webLink } so the client can open it in Outlook.
async function createReplyAllDraft(token, messageId, htmlBody) {
  // Step 1: Create the draft (POST /me/messages/{id}/createReplyAll)
  const createUrl = `${GRAPH_BASE}/me/messages/${encodeURIComponent(messageId)}/createReplyAll`;
  const controller1 = new AbortController();
  const timer1 = setTimeout(() => controller1.abort(), TIMEOUT_MS);
  let draft;
  try {
    const res = await fetch(createUrl, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({}),
      signal: controller1.signal,
    });
    if (!res.ok) {
      const text = await res.text();
      const err = new Error(`Graph API ${res.status}: ${text}`);
      err.status = res.status;
      throw err;
    }
    draft = await res.json();
  } finally {
    clearTimeout(timer1);
  }

  // Step 2: Update the draft body (PATCH /me/messages/{draftId})
  const patchUrl = `${GRAPH_BASE}/me/messages/${encodeURIComponent(draft.id)}`;
  const controller2 = new AbortController();
  const timer2 = setTimeout(() => controller2.abort(), TIMEOUT_MS);
  try {
    const res = await fetch(patchUrl, {
      method: "PATCH",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        body: { contentType: "HTML", content: htmlBody || "" },
      }),
      signal: controller2.signal,
    });
    if (!res.ok) {
      const text = await res.text();
      const err = new Error(`Graph API PATCH ${res.status}: ${text}`);
      err.status = res.status;
      throw err;
    }
    const updated = await res.json();
    return {
      draftId: updated.id,
      webLink: updated.webLink || draft.webLink || null,
    };
  } finally {
    clearTimeout(timer2);
  }
}

// ── Create New Draft (no thread) ─────────────────────────────
// Creates a brand-new draft message via Graph API.
// Returns { draftId, webLink } so the client can open it in Outlook.
async function createNewDraft(token, { to, subject, htmlBody }) {
  const url = `${GRAPH_BASE}/me/messages`;
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);
  try {
    const res = await fetch(url, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        subject: subject || "",
        body: { contentType: "HTML", content: htmlBody || "" },
        toRecipients: (to || []).map(email => ({
          emailAddress: { address: email },
        })),
      }),
      signal: controller.signal,
    });
    if (!res.ok) {
      const text = await res.text();
      const err = new Error(`Graph API ${res.status}: ${text}`);
      err.status = res.status;
      throw err;
    }
    const draft = await res.json();
    return { draftId: draft.id, webLink: draft.webLink || null };
  } finally {
    clearTimeout(timer);
  }
}

// ── Get Full Email Thread for HubSpot Logging ────────────────
// Returns individual messages with full details needed for creating
// separate email engagements in HubSpot.
async function getEmailThreadForLogging(token, conversationId) {
  const safeConvId = escapeODataFilter(conversationId);
  const filter = `conversationId eq '${safeConvId}'`;
  const url =
    `${GRAPH_BASE}/me/messages` +
    `?$filter=${encodeURIComponent(filter)}` +
    `&$top=25&$select=id,internetMessageId,subject,body,from,toRecipients,ccRecipients,receivedDateTime` +
    `&$orderby=receivedDateTime asc`;

  let data;
  try {
    data = await graphGet(url, token);
  } catch (err) {
    // Some tenants don't support $filter + $orderby — retry without $orderby
    try {
      const fallbackUrl =
        `${GRAPH_BASE}/me/messages` +
        `?$filter=${encodeURIComponent(filter)}` +
        `&$top=25&$select=id,internetMessageId,subject,body,from,toRecipients,ccRecipients,receivedDateTime`;
      data = await graphGet(fallbackUrl, token);
    } catch (err2) {
      throw new Error(`Failed to fetch email thread for logging: ${err2.message}`);
    }
  }

  // Deduplicate: Graph API can return the same message from multiple folders
  // (Inbox + Sent Items). Use internetMessageId as the canonical key since it
  // is the same across all folder copies; fall back to Graph id.
  const seen = new Set();
  const unique = (data.value || []).filter((m) => {
    const key = m.internetMessageId || m.id;
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });

  // Subject filter: conversationId can group unrelated threads with similar
  // subjects. Keep only messages whose normalised subject matches the anchor
  // (first message in the thread, which is the oldest after asc ordering).
  let filtered = unique;
  if (unique.length > 1) {
    const anchor = normalizeSubject(unique[0].subject);
    filtered = unique.filter(m => normalizeSubject(m.subject) === anchor);
    if (filtered.length < unique.length) {
      console.warn(`[graph] getEmailThreadForLogging: subject filter removed ${unique.length - filtered.length} unrelated message(s)`);
    }
  }

  return filtered.map((m) => ({
    subject:   m.subject || "(no subject)",
    fromEmail: m.from?.emailAddress?.address || "",
    fromName:  m.from?.emailAddress?.name || "",
    toEmails:  (m.toRecipients || []).map(r => r.emailAddress?.address).filter(Boolean).join(";"),
    ccEmails:  (m.ccRecipients || []).map(r => r.emailAddress?.address).filter(Boolean).join(";"),
    sentAt:    m.receivedDateTime || new Date().toISOString(),
    bodyHtml:  sanitizeEmailHtml(m.body?.content || "").substring(0, 65536),
    bodyText:  stripHtml(m.body?.content || "").substring(0, 10000),
  }));
}

// ── Signature Extraction from Sent Mail ──────────────────────
// Fetches the most recent sent email and extracts the signature HTML.
// Used as a fallback when the client has no cached signature (e.g. Graph reply
// from read mode without ever opening a compose window).

// Tries multiple signature detection strategies in priority order:
// 1. OWA signature div (id="Signature")
// 2. Prefixed variant (id="x_Signature") - added by some Outlook versions
// 3. Class-based (class contains "signature")
// 4. HR divider - classic Outlook desktop style
// 5. Unix "-- " delimiter
function extractSignatureFromHtml(html) {
  if (!html) return "";
  // Only search the tail portion to avoid false positives from quoted content
  const tail = html.length > 2000 ? html.slice(-2000) : html;
  const offset = html.length - tail.length;

  // OWA / New Outlook: <div id="Signature">
  const owaMatch = tail.match(/<div[^>]+id=["']Signature["'][^>]*>[\s\S]*/i);
  if (owaMatch) return owaMatch[0];
  // Prefixed variant: <div id="x_Signature">
  const prefixedMatch = tail.match(/<div[^>]+id=["']x_Signature["'][^>]*>[\s\S]*/i);
  if (prefixedMatch) return prefixedMatch[0];
  // Class-based (e.g. class="OutlookSignature")
  const classMatch = tail.match(/<div[^>]+class=["'][^"']*signature[^"']*["'][^>]*>[\s\S]*/i);
  if (classMatch) return classMatch[0];
  // Classic Outlook: <hr> divider
  const hrIdx = tail.search(/<hr[\s>]/i);
  if (hrIdx !== -1) return tail.slice(hrIdx);
  // Unix-style "-- " delimiter
  const unixIdx = tail.indexOf("-- <br");
  if (unixIdx !== -1) return tail.slice(unixIdx);
  return "";
}

// Fetches the user's most recent sent email and extracts the signature HTML from it.
async function getSignatureFromSentMail(token) {
  const url =
    `${GRAPH_BASE}/me/mailFolders/SentItems/messages` +
    `?$top=1&$select=body&$orderby=sentDateTime desc`;
  const data = await graphGet(url, token);
  const messages = (data && data.value) || [];
  if (messages.length === 0) return "";
  const bodyHtml = (messages[0].body && messages[0].body.content) || "";
  return extractSignatureFromHtml(bodyHtml);
}

// ── Token Refresh (server-side) ─────────────────────────────

// Exchanges an OAuth refresh token for a new access token via Azure AD.
// Used by autoTracker and subscription renewal when the cached token expires.
async function refreshAccessToken(refreshToken) {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  if (!tenantId || !clientId) throw new Error("AZURE_TENANT_ID and AZURE_CLIENT_ID required for token refresh");

  const params = new URLSearchParams({
    client_id:     clientId,
    grant_type:    "refresh_token",
    refresh_token: refreshToken,
    scope:         "openid profile User.Read Mail.Read Mail.ReadWrite Mail.Send MailboxSettings.Read offline_access",
  });
  // Confidential client: include client_secret for Web platform tokens
  if (process.env.AZURE_CLIENT_SECRET) {
    params.set("client_secret", process.env.AZURE_CLIENT_SECRET);
  }

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);
  try {
    const res = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      method:  "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body:    params.toString(),
      signal:  controller.signal,
    });
    const data = await res.json();
    if (data.error || !data.access_token) {
      throw new Error(`Token refresh failed: ${data.error || "no access_token"} — ${data.error_description || ""}`);
    }
    return {
      accessToken:  data.access_token,
      refreshToken: data.refresh_token || refreshToken,
      expiresIn:    data.expires_in || 3600,
    };
  } catch (err) {
    if (err.name === "AbortError") throw new Error("Token refresh timed out");
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

// ── Fetch Single Message by ID ──────────────────────────────

// Fetches a single email by its Graph message ID. Returns null for drafts.
// Used by autoTracker when processing webhook notifications about new emails.
async function getMessageById(token, messageId) {
  const select = "id,internetMessageId,conversationId,subject,body,from,toRecipients,ccRecipients,receivedDateTime,isDraft";
  const data = await graphGet(`${GRAPH_BASE}/me/messages/${encodeURIComponent(messageId)}?$select=${select}`, token);

  if (data.isDraft) return null; // skip drafts

  const fromAddr = data.from?.emailAddress?.address || "";
  const toAddrs  = (data.toRecipients || []).map(r => r.emailAddress?.address).filter(Boolean);
  const ccAddrs  = (data.ccRecipients || []).map(r => r.emailAddress?.address).filter(Boolean);
  const rawHtml  = data.body?.contentType === "html" ? data.body.content : "";
  const rawText  = data.body?.contentType === "text" ? data.body.content : "";

  return {
    id:               data.id,
    internetMessageId: data.internetMessageId || data.id,
    conversationId:   data.conversationId || null,
    subject:          data.subject || "(no subject)",
    fromEmail:        fromAddr,
    fromName:         data.from?.emailAddress?.name || fromAddr,
    toEmails:         toAddrs.join(";"),
    ccEmails:         ccAddrs.join(";"),
    allParticipants:  [fromAddr, ...toAddrs, ...ccAddrs].filter(Boolean),
    sentAt:           data.receivedDateTime || new Date().toISOString(),
    bodyHtml:         sanitizeEmailHtml((rawHtml || "").substring(0, 65536)),
    bodyText:         stripHtml(rawHtml || rawText || "").substring(0, 10000),
  };
}

// ── Find Last Email Thread with a Contact ────────────────────
// Searches for the most recent email involving the contact, returns
// the message ID (for reply), conversationId, and participant breakdown
// (last sender → To, everyone else → CC, excluding the user).
async function getLastEmailWithContact(token, contactEmail) {
  const searchTerm = encodeURIComponent('"' + escapeODataSearch(contactEmail) + '"');
  const url =
    `${GRAPH_BASE}/me/messages` +
    `?$search=${searchTerm}` +
    `&$top=1&$select=id,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime` +
    `&$orderby=receivedDateTime desc`;

  let data;
  try {
    data = await graphGet(url, token);
  } catch (err) {
    // Fallback without $orderby
    const fallbackUrl =
      `${GRAPH_BASE}/me/messages` +
      `?$search=${searchTerm}` +
      `&$top=1&$select=id,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime`;
    try {
      data = await graphGet(fallbackUrl, token);
    } catch (err2) {
      console.warn("[graph] getLastEmailWithContact failed:", err2.message);
      return null;
    }
  }

  const messages = data?.value || [];
  if (messages.length === 0) return null;

  const msg = messages[0];
  const fromAddr = msg.from?.emailAddress?.address || "";
  const fromName = msg.from?.emailAddress?.name || fromAddr;
  const toAddrs  = (msg.toRecipients || []).map(r => ({ email: r.emailAddress?.address || "", name: r.emailAddress?.name || "" })).filter(r => r.email);
  const ccAddrs  = (msg.ccRecipients || []).map(r => ({ email: r.emailAddress?.address || "", name: r.emailAddress?.name || "" })).filter(r => r.email);

  return {
    messageId:      msg.id,
    conversationId: msg.conversationId || null,
    subject:        msg.subject || "(no subject)",
    receivedAt:     msg.receivedDateTime || null,
    from:           { email: fromAddr, name: fromName },
    to:             toAddrs,
    cc:             ccAddrs,
  };
}

// Fetches the most recent message in a conversation thread by conversationId.
// Used by follow-up feature to find the right message to reply to.
async function getLatestMessageInThread(token, conversationId) {
  const safeConvId = escapeODataFilter(conversationId);
  const filter = `conversationId eq '${safeConvId}'`;
  const url =
    `${GRAPH_BASE}/me/messages` +
    `?$filter=${encodeURIComponent(filter)}` +
    `&$top=1&$select=id,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime` +
    `&$orderby=receivedDateTime desc`;

  let data;
  try {
    data = await graphGet(url, token);
  } catch (err) {
    // Fallback without $orderby
    const fallbackUrl =
      `${GRAPH_BASE}/me/messages` +
      `?$filter=${encodeURIComponent(filter)}` +
      `&$top=1&$select=id,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime`;
    data = await graphGet(fallbackUrl, token);
  }

  const messages = data?.value || [];
  if (messages.length === 0) return null;

  const msg = messages[0];
  return {
    messageId:      msg.id,
    conversationId: msg.conversationId || conversationId,
    subject:        msg.subject || "(no subject)",
    receivedAt:     msg.receivedDateTime || null,
  };
}

module.exports = { getSentMailStyle, getEmailHistory, sendEmail, getEmailAnalytics, searchPeople, getEmailThread, getEmailThreadForLogging, replyToMessage, createReplyAllDraft, createNewDraft, getSignatureFromSentMail, normalizeSubject, refreshAccessToken, getMessageById, getLastEmailWithContact, getLatestMessageInThread };
