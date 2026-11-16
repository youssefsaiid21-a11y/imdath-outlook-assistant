"use strict";
/* Sphinx Utils - shared helpers with no cross-module dependencies.
   This file MUST be loaded first via <script> because all other modules
   depend on these functions. It has zero imports from other project files. */

// Prevent uncaught errors from crashing the Office host pane.
// Returning true tells the browser we handled the error, so it won't
// show a default error dialog inside the Outlook taskpane.
window.onerror = function (msg, src, line, col, err) {
  console.error("[Imdath] Uncaught error:", msg, src, line, err);
  return true;
};
window.onunhandledrejection = function (e) {
  console.error("[Imdath] Unhandled rejection:", e.reason);
  e.preventDefault();
};

// ── Helpers ───────────────────────────────────────────────────
var _toastTimer = null; // Active timer for auto-hiding toast notifications

// Clean up all timers when the taskpane unloads to avoid memory leaks.
// These timer vars are defined in other modules but accessible here because
// all files share the same global scope via <script> tags.
window.addEventListener("unload", () => {
  clearInterval(_recipientPollTimer);
  clearTimeout(_hubspotBtnTimer);
  clearTimeout(_toastTimer);
});

// Safe text setter - silently no-ops if the element doesn't exist
function setText(id, val) {
  const el = document.getElementById(id);
  if (el) el.textContent = val;
}

// Sanitize user/API strings before injecting into innerHTML to prevent XSS
function escapeHtml(str) {
  return str
    .replace(/&/g,  "&amp;")
    .replace(/</g,  "&lt;")
    .replace(/>/g,  "&gt;")
    .replace(/"/g,  "&quot;")
    .replace(/'/g,  "&#39;");
}

// Map HubSpot deal stage names to CSS badge colors for visual status
function getStageBadgeClass(stage) {
  if (!stage) return "badge-grey";
  const s = stage.toLowerCase();
  if (s.includes("closed won") || /\bwon\b/.test(s))                             return "badge-green";
  if (s.includes("proposal") || s.includes("progress") || s.includes("negotiat")) return "badge-orange";
  return "badge-grey";
}

// Display a brief notification toast at the bottom of the taskpane.
// type can be "success" or "error" for color styling; auto-hides after 3s.
function showToast(message, type) {
  const toast = document.getElementById("status-toast");
  if (!toast) return;
  toast.textContent = message;
  toast.classList.remove("show", "success", "error");
  // Force reflow so re-adding "show" triggers the CSS transition even for back-to-back toasts
  void toast.offsetWidth;
  toast.classList.add("show");
  if (type) toast.classList.add(type);
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => {
    toast.classList.remove("show", "success", "error");
  }, 3000);
}

// Convert a numeric "days since last contact" into a human-readable label.
// 999 is a sentinel value from the server meaning "no email ever sent".
function formatDaysAgo(days) {
  if (days >= 999) return "Never";
  if (days === 0)  return "Today";
  if (days === 1)  return "1d ago";
  return days + "d ago";
}

// ── Pane Size Toggle ──────────────────────────────────────────
// Lets users switch between full and compact taskpane layouts.
// Preference persists in localStorage across sessions.
function togglePaneSize() {
  const isCompact = document.body.classList.toggle("compact");
  localStorage.setItem("imdath_compact", isCompact ? "1" : "0");
  updateSizeIcon(isCompact);
}

// Called on startup to restore the user's last pane size preference
function restorePaneSize() {
  const compact = localStorage.getItem("imdath_compact") === "1";
  if (compact) document.body.classList.add("compact");
  updateSizeIcon(compact);
}

// Swap the shrink/expand icon to reflect the current pane state
function updateSizeIcon(isCompact) {
  const shrink = document.getElementById("size-icon-shrink");
  const expand = document.getElementById("size-icon-expand");
  if (shrink) shrink.style.display = isCompact ? "none"  : "block";
  if (expand) expand.style.display = isCompact ? "block" : "none";
}

"use strict";
/* Sphinx Auth - session management, token refresh, and the apiFetch wrapper.
   Handles sign-in via an Office dialog, silent token refresh, and automatic
   401-retry logic so every other module can call apiFetch() worry-free.
   Depends on: utils.js (showToast, setText). */

// ── State (auth) ──────────────────────────────────────────────
var API_BASE = "/api";            // All API calls are proxied through the same origin
var accessToken    = null;        // JWT access token for the current session
var currentAccount = null;        // { name, username, ... } of the signed-in user

// ── Session Management ────────────────────────────────────────
// Sessions are stored in localStorage so they survive taskpane close/reopen.
// Only the access token and account info are stored client-side; the refresh
// token is exclusively owned by the server for security.
var SESSION_KEY = "imdath_session";

// Persist session to localStorage. Expiry is set 60s before actual token
// expiry so we refresh proactively instead of hitting 401s.
function saveSession(data) {
  const session = {
    accessToken:  data.accessToken,
    // Refresh token is owned exclusively by the server — never stored client-side.
    expiresAt:    Date.now() + ((data.expiresIn || 3600) - 60) * 1000,
    account:      data.account,
  };
  // Use localStorage so session persists across add-in pane open/close.
  // When AT expires, silentSignIn calls /api/auth/refresh to get a new one.
  localStorage.setItem(SESSION_KEY, JSON.stringify(session));
}

// Remove all client-side session data (sign-out or expired)
function clearSession() {
  localStorage.removeItem(SESSION_KEY);
  sessionStorage.removeItem(SESSION_KEY); // clean up legacy
  try { localStorage.removeItem("imdath_signature"); } catch (e) {}
}

// Decode the payload segment of a JWT without verifying the signature
function decodeJwtPayload(token) {
  const parts = token.split(".");
  if (parts.length !== 3) throw new Error("Invalid JWT structure");
  return JSON.parse(atob(parts[1].replace(/-/g, "+").replace(/_/g, "/")));
}

// Ask the server to issue a new access token using the refresh token it holds.
// The client sends its (possibly expired) access token so the server can
// identify which user to refresh. Returns { accessToken, expiresIn } or null.
async function refreshSession(expiredToken) {
  try {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 15000);
    let resp;
    try {
      resp = await fetch(`${API_BASE}/auth/refresh`, {
        method:  "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${expiredToken}` },
        signal:  controller.signal,
      });
    } finally {
      clearTimeout(timer);
    }
    if (!resp.ok) return null;
    const data = await resp.json();
    if (!data.accessToken) return null;
    // Preserve existing account info — server doesn't return it on refresh
    return {
      accessToken: data.accessToken,
      expiresIn:   data.expiresIn || 3600,
    };
  } catch (e) {
    console.warn("[Imdath] refreshSession failed:", e.message);
    return null;
  }
}

// Try to restore a session from localStorage without showing UI.
// If the access token is still valid, use it directly.
// If expired, attempt a server-side refresh. Returns true if signed in.
async function silentSignIn() {
  try {
    const stored = localStorage.getItem(SESSION_KEY) || sessionStorage.getItem(SESSION_KEY);
    if (!stored) return false;
    let session;
    try {
      session = JSON.parse(stored);
    } catch {
      // Corrupted session data — clear it and force re-login
      clearSession();
      return false;
    }

    if (session.expiresAt > Date.now()) {
      accessToken    = session.accessToken;
      currentAccount = session.account;
      updateAuthUI(currentAccount.name || currentAccount.username);
      return true;
    }

    // Access token expired — ask server to refresh (server owns the refresh token)
    if (session.accessToken) {
      const refreshed = await refreshSession(session.accessToken);
      if (refreshed) {
        // Preserve account info from the original session
        saveSession({ ...refreshed, account: session.account });
        accessToken    = refreshed.accessToken;
        currentAccount = session.account;
        updateAuthUI(currentAccount.name || currentAccount.username);
        return true;
      }
    }

    clearSession();
    return false;
  } catch (e) {
    console.warn("[Imdath] silentSignIn failed:", e.message);
    return false;
  }
}

// Sign-in / sign-out toggle. When signed in, clears all state and notifies
// the server. When signed out, opens an Office dialog for Microsoft OAuth.
function handleAuthBtn() {
  if (accessToken) {
    // Notify server to revoke tokens/subscriptions (fire-and-forget)
    fetch(API_BASE + "/auth/logout", {
      method: "POST",
      headers: { Authorization: "Bearer " + accessToken }
    }).catch(function () {});
    accessToken      = null;
    currentAccount   = null;
    currentContext   = null;
    _lastLoadedEmail = null;
    _cachedSignature = null;
    _pipelinesCache  = null;
    _followupsLoaded = false;
    _hasGenerated    = false;
    _allRecipients   = [];
    _followupDeals   = [];
    _followupTarget  = null;
    _followupReplyMeta = null;
    _followupDraftGeneration = 0;
    _followupSending = false;
    try {
      var fuList = document.getElementById("followup-list");
      if (fuList) fuList.innerHTML = "";
      setText("stat-total-deals", "—");
      setText("stat-overdue", "—");
      setText("stat-stalest", "—");
    } catch (e) {}
    clearInterval(_recipientPollTimer);
    _recipientPollTimer = null;
    dismissReplyMode();
    clearSession();
    hideMainContent();
    updateAuthUI(null);
    return;
  }

  Office.context.ui.displayDialogAsync(
    window.location.origin + "/auth-dialog.html",
    { height: 60, width: 30, promptBeforeOpen: false },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showToast("Could not open sign-in window. Error: " + asyncResult.error.message, "error");
        return;
      }

      const dialog  = asyncResult.value;
      const AUTH_KEY = "imdath_auth_result";
      sessionStorage.removeItem(AUTH_KEY);

      // Guard: both messageParent callback and sessionStorage poll can fire,
      // so we use a flag to ensure only the first one processes the result.
      let authHandled = false;

      // Process the auth result from the dialog. Supports two flows:
      // 1. "success" - dialog already exchanged the code and returns tokens
      // 2. "code" - dialog returns an auth code for the client to exchange
      async function handleAuthResult(msg) {
        try {
          if (msg.status === "success") {
            accessToken    = msg.accessToken;
            currentAccount = msg.account;
            saveSession(msg); // RT is NOT stored — saveSession excludes it
            updateAuthUI(currentAccount.name || currentAccount.username);
            showMainContent();
            initRecipientDetection();
            // Hand off refresh token to server (one-time, with retry).
            // Server needs the RT for background tasks like webhook processing.
            // Client never stores the RT - only the server keeps it.
            if (msg.refreshToken && msg.accessToken) {
              var handoffOk = false;
              for (var _attempt = 0; _attempt < 2 && !handoffOk; _attempt++) {
                try {
                  var hRes = await fetch(API_BASE + "/auth/handoff", {
                    method:  "POST",
                    headers: { "Content-Type": "application/json", Authorization: "Bearer " + msg.accessToken },
                    body:    JSON.stringify({ refreshToken: msg.refreshToken })
                  });
                  if (hRes.ok) {
                    handoffOk = true;
                    console.log("[Imdath] Refresh token handed off to server");
                  }
                } catch (hErr) {
                  console.warn("[Imdath] RT handoff attempt " + (_attempt + 1) + " failed:", hErr.message);
                }
              }
              if (!handoffOk) {
                showToast("Session saved partially \u2014 sign out and back in if auto-tracking doesn\u2019t work.", "error");
              }
            }
          } else if (msg.status === "code") {
            const resp = await fetch(API_BASE + "/auth/exchange", {
              method:  "POST",
              headers: { "Content-Type": "application/json" },
              body:    JSON.stringify({ code: msg.code, verifier: msg.verifier })
            });
            const data = await resp.json();
            if (data.error) { showToast("Sign-in failed: " + data.error, "error"); return; }
            accessToken    = data.accessToken;
            currentAccount = data.account;
            saveSession(data);
            updateAuthUI(currentAccount.name || currentAccount.username);
            showMainContent();
            initRecipientDetection();
          } else {
            showToast("Sign-in failed: " + (msg.message || "Unknown error"), "error");
          }
        } catch (e) {
          showToast("Sign-in failed. Please try again.", "error");
        }
      }

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        if (authHandled) return;
        authHandled = true;
        clearInterval(pollTimer);
        dialog.close();
        try { handleAuthResult(JSON.parse(arg.message)); }
        catch (e) { console.error("[Imdath] Auth message parse failed:", e); }
      });

      // Fallback: poll sessionStorage every 500ms. Some Outlook hosts don't
      // support messageParent after a redirect, so the dialog writes to
      // sessionStorage and we pick it up here.
      const pollTimer = setInterval(function () {
        const stored = sessionStorage.getItem(AUTH_KEY);
        if (!stored) return;
        if (authHandled) { clearInterval(pollTimer); return; }
        authHandled = true;
        clearInterval(pollTimer);
        sessionStorage.removeItem(AUTH_KEY);
        try { dialog.close(); } catch (e) {}
        try { handleAuthResult(JSON.parse(stored)); }
        catch (e) { console.error("[Imdath] Auth poll parse failed:", e); }
      }, 500);

      dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
        clearInterval(pollTimer);
        if (authHandled) return;
        // One final check in case dialog closed right after storing result
        const stored = sessionStorage.getItem(AUTH_KEY);
        if (stored) {
          authHandled = true;
          sessionStorage.removeItem(AUTH_KEY);
          try { handleAuthResult(JSON.parse(stored)); }
          catch (e) { console.error("[Imdath] Auth dialog-event parse failed:", e); }
        }
      });
    }
  );
}

// Accessor for the current access token - used by all modules
function getToken() {
  return accessToken || null;
}

// ── API Fetch Wrapper (handles 401 globally) ───────────────────
// Drop-in replacement for fetch(). On a 401, it automatically attempts
// one token refresh and retries. If that also fails, it signs the user out
// and shows a "session expired" banner. All API calls should use this.
async function apiFetch(url, options = {}, _retried) {
  const res = await fetch(url, options);
  if (res.status === 401 && !_retried) {
    // Attempt a token refresh before giving up
    var oldToken = accessToken;
    var refreshed = oldToken ? await refreshSession(oldToken) : null;
    if (refreshed) {
      var stored = localStorage.getItem(SESSION_KEY);
      var acct = currentAccount;
      try { acct = JSON.parse(stored).account || acct; } catch (_e) {}
      saveSession({ ...refreshed, account: acct });
      accessToken = refreshed.accessToken;
      // Retry with new token
      var retryOpts = Object.assign({}, options, {
        headers: Object.assign({}, options.headers || {}, { Authorization: "Bearer " + refreshed.accessToken })
      });
      return apiFetch(url, retryOpts, true);
    }
    accessToken    = null;
    currentContext = null;
    clearSession();
    updateAuthUI(null);
    hideMainContent();
    showSessionExpiredBanner();
    throw new Error("Session expired — please sign in again.");
  }
  if (res.status === 401 && _retried) {
    accessToken    = null;
    currentContext = null;
    clearSession();
    updateAuthUI(null);
    hideMainContent();
    showSessionExpiredBanner();
    throw new Error("Session expired — please sign in again.");
  }
  return res;
}

// Extract a human-readable error message from a failed API response
async function parseApiError(res) {
  try { const d = await res.json(); return d.error || `Server error ${res.status}`; }
  catch (_) { return `Server error ${res.status}`; }
}

function showSessionExpiredBanner() {
  const p = document.getElementById("auth-banner-subtitle");
  if (p) {
    p.textContent = "Your session expired. Sign in to continue.";
    p.style.color = "var(--color-warning)";
  }
}

// ── Auth UI ───────────────────────────────────────────────────
// Toggle the sign-in button label and show/hide the main content area
function updateAuthUI(userName) {
  const btn = document.getElementById("auth-btn");
  if (userName) {
    btn.textContent = "Sign out";
    btn.title = userName;
  } else {
    btn.textContent = "Sign in";
  }
}

function showMainContent() {
  document.getElementById("auth-banner").style.display  = "none";
  document.getElementById("main-content").style.display = "block";
}

function hideMainContent() {
  document.getElementById("auth-banner").style.display  = "block";
  document.getElementById("main-content").style.display = "none";
}
"use strict";
/* Sphinx Recipient - detection, autocomplete, and multi-recipient chips.
   Handles two Outlook modes:
   - Compose mode: polls the To field every 3s for recipient changes
   - Read mode: reads the sender from the currently open email
   Depends on: utils.js, auth.js, context.js (loadContext, setContextLoading). */

// ── State (recipient) ─────────────────────────────────────────
var _recipientPollTimer = null;     // setInterval ID for compose-mode To-field polling
var _lastLoadedEmail    = null;     // Prevents re-loading context for the same email
var _recipientSearchTimer = null;   // Debounce timer for autocomplete search
var _dropdownActiveIndex  = -1;     // Keyboard-navigable index in autocomplete dropdown
var _allRecipients      = [];       // All To-field recipients (for multi-recipient chips)

// ── Recipient Detection ───────────────────────────────────────
// Entry point called after sign-in and on item change. Detects whether
// we're in compose mode (has item.to.getAsync) or read mode (has item.from)
// and sets up the appropriate recipient handling.
function initRecipientDetection() {
  if (!Office || !Office.context || !Office.context.mailbox) return;
  const item = Office.context.mailbox.item;
  if (!item) return;

  if (item.to && item.to.getAsync) {
    // Compose mode — poll for recipient changes
    hideReadModeBanner();
    clearInterval(_recipientPollTimer);
    checkRecipient();
    _recipientPollTimer = setInterval(checkRecipient, 3000);
    detectReplyMode();
    // If the user generated a draft in read mode and then clicked Reply
    // through Sphinx, Outlook spawns a fresh compose pane. Restore the
    // draft that was persisted just before the reply form opened.
    try { restorePendingReplyDraft(); } catch (e) { console.warn("[Imdath] restorePendingReplyDraft failed:", e); }
  } else if (item.from) {
    // Read mode — load context for the sender and show reply options
    const fromEmail = item.from.emailAddress;
    const fromName  = item.from.displayName;
    if (fromEmail && fromEmail !== _lastLoadedEmail) {
      _lastLoadedEmail = fromEmail;
      setRecipientAndLoad(fromEmail, fromName);
    }
    showReadModeBanner();
  }
}

// Poll the Outlook To field for changes. Only triggers a context reload
// when the primary (first) recipient email actually changes.
function checkRecipient() {
  if (!Office.context || !Office.context.mailbox) return;
  const item = Office.context.mailbox.item;
  if (!item || !item.to) return;
  item.to.getAsync((result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.warn("[Imdath] getAsync failed:", result.error?.message);
      return;
    }
    const recipients = result.value || [];
    if (recipients.length === 0) {
      if (_lastLoadedEmail !== null) {
        _lastLoadedEmail = null;
        _allRecipients = [];
        setContextPending("No recipient yet");
        renderRecipientChips();
      }
      return;
    }
    const recip = recipients[0];
    const email = recip.emailAddress;

    // Store all recipients and render chips
    _allRecipients = recipients.map(r => ({ email: r.emailAddress, name: r.displayName || r.emailAddress }));
    renderRecipientChips();

    if (email && email !== _lastLoadedEmail) {
      _lastLoadedEmail = email;
      setRecipientAndLoad(email, recip.displayName);
    }
  });
}

// Update the recipient input field and trigger a full context load from the server
function setRecipientAndLoad(email, name) {
  _userClosedContext = false; // new recipient — reset so context opens fresh
  const input = document.getElementById("recipient-input");
  if (input) input.value = name ? `${name} <${email}>` : email;
  // Hide deal panel from previous contact while loading
  const dealSection = document.getElementById("deal-section");
  if (dealSection) dealSection.style.display = "none";
  setContextLoading();
  loadContext(email);
}

// ── Recipient Autocomplete ────────────────────────────────────
// Lets users manually search for contacts by typing in the recipient field.
// Debounces input to avoid excessive API calls.

// Triggered on each keystroke in the recipient input field
function onRecipientInput(input) {
  const q = input.value.trim();
  clearTimeout(_recipientSearchTimer);
  if (q.length < 2) { hideRecipientDropdown(); return; }
  _recipientSearchTimer = setTimeout(() => searchPeople(q), 280);
}

// Keyboard navigation for the autocomplete dropdown (arrows, Enter, Escape)
function onRecipientKeydown(event) {
  const dd = document.getElementById("recipient-dropdown");
  const items = dd.querySelectorAll(".recipient-option");
  if (!items.length) return;

  if (event.key === "ArrowDown") {
    event.preventDefault();
    _dropdownActiveIndex = Math.min(_dropdownActiveIndex + 1, items.length - 1);
    highlightDropdownItem(items);
  } else if (event.key === "ArrowUp") {
    event.preventDefault();
    _dropdownActiveIndex = Math.max(_dropdownActiveIndex - 1, 0);
    highlightDropdownItem(items);
  } else if (event.key === "Enter" && _dropdownActiveIndex >= 0) {
    event.preventDefault();
    items[_dropdownActiveIndex].click();
  } else if (event.key === "Escape") {
    hideRecipientDropdown();
  }
}

function highlightDropdownItem(items) {
  items.forEach((el, i) => el.classList.toggle("active", i === _dropdownActiveIndex));
}

// Called on input blur. Delay hides the dropdown so mousedown on an option
// can fire before the dropdown disappears.
function scheduleHideDropdown() {
  setTimeout(hideRecipientDropdown, 200);
}

function hideRecipientDropdown() {
  const dd = document.getElementById("recipient-dropdown");
  if (dd) dd.style.display = "none";
  _dropdownActiveIndex = -1;
}

// Search the server's people/contacts index for autocomplete results
async function searchPeople(q) {
  const token = getToken();
  if (!token) return;
  try {
    const res = await apiFetch(`${API_BASE}/people?q=${encodeURIComponent(q)}`, {
      headers: { "Authorization": `Bearer ${token}` },
    });
    if (!res.ok) {
      console.warn("[Imdath] searchPeople failed:", res.status);
      return;
    }
    const people = await res.json();
    showRecipientDropdown(people);
  } catch (e) {
    console.warn("[Imdath] searchPeople failed:", e);
  }
}

// Render the autocomplete dropdown with search results
function showRecipientDropdown(people) {
  const dd = document.getElementById("recipient-dropdown");
  if (!dd) return;
  dd.innerHTML = "";
  _dropdownActiveIndex = -1;

  if (people.length === 0) { dd.style.display = "none"; return; }

  people.forEach(p => {
    const opt = document.createElement("div");
    opt.className = "recipient-option";
    opt.setAttribute("role", "option");

    const name = document.createElement("span");
    name.className   = "ro-name";
    name.textContent = p.name || p.email;

    const meta = document.createElement("span");
    meta.className   = "ro-meta";
    meta.textContent = p.name ? p.email : (p.title || p.company || "");

    opt.appendChild(name);
    opt.appendChild(meta);

    // mousedown prevents blur from firing before click
    opt.addEventListener("mousedown", (e) => {
      e.preventDefault();
      selectRecipient(p.email, p.name);
    });

    dd.appendChild(opt);
  });

  dd.style.display = "block";
}

// Called when user picks a contact from autocomplete. Updates both the
// add-in's input field and Outlook's actual To field via setAsync.
function selectRecipient(email, name) {
  const input = document.getElementById("recipient-input");
  if (input) input.value = name ? `${name} <${email}>` : email;
  hideRecipientDropdown();

  // Update Outlook To field if in compose mode
  if (Office?.context?.mailbox?.item?.to?.setAsync) {
    Office.context.mailbox.item.to.setAsync(
      [{ emailAddress: email, displayName: name || email }],
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.warn("[Imdath] setAsync failed:", result.error?.message);
        }
      }
    );
  }

  if (email !== _lastLoadedEmail) {
    _lastLoadedEmail = email;
    setContextLoading();
    loadContext(email);
  }
}

// ── Multi-Recipient Chips ─────────────────────────────────────
// When there are multiple To recipients, show the primary one as a chip
// plus a "+N" badge for the rest. Hidden when there's 0 or 1 recipient.
function renderRecipientChips() {
  const container = document.getElementById("multi-recipient-chips");
  if (!container) return;

  if (_allRecipients.length <= 1) {
    container.style.display = "none";
    container.innerHTML = "";
    return;
  }

  container.style.display = "flex";
  container.innerHTML = "";

  // Show first recipient as "primary" + count badge for rest
  const primary = _allRecipients[0];
  const chip = document.createElement("span");
  chip.className = "recipient-chip primary";
  chip.textContent = primary.name || primary.email;
  chip.title = primary.email;
  container.appendChild(chip);

  if (_allRecipients.length > 1) {
    const more = document.createElement("span");
    more.className = "recipient-chip more";
    more.textContent = `+${_allRecipients.length - 1}`;
    more.title = _allRecipients.slice(1).map(r => r.name || r.email).join(", ");
    container.appendChild(more);
  }
}
"use strict";
/* Sphinx Context - the CRM context card, deal stage management, and auto-track.
   Fetches contact/deal/email-history data from the server for the current
   recipient and renders it in the collapsible context panel.
   Depends on: utils.js, auth.js. */

// ── State (context) ───────────────────────────────────────────
var currentContext    = null;    // Full context object from /api/context (recipient, deals, notes, etc.)
var contextOpen       = false;   // Whether the context panel is expanded
var _userClosedContext = false;   // True if user explicitly closed the panel (don't auto-reopen)
var _pipelinesCache   = null;    // Cached HubSpot pipelines/stages to avoid repeated fetches
var _pipelinesPromise = null;    // In-flight pipelines request to dedupe concurrent calls
var _moveLinkRef      = null;    // Reference to the "Move" button that opened the stage panel
var _moveStageKeyHandler = null; // Stored keyboard handler for the move-stage select

// ── Context Loading ───────────────────────────────────────────
// Fetch CRM context (contact info, deals, email history, notes) from the
// server for the given email address and render it in the context card.
async function loadContext(email) {
  if (!email) return;
  const token = getToken();
  if (!token) return;

  try {
    const res = await apiFetch(`${API_BASE}/context`, {
      method:  "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
      body:    JSON.stringify({ email }),
    });

    if (!res.ok) throw new Error(await parseApiError(res));
    const ctx = await res.json();
    currentContext = ctx;
    renderContext(ctx);
  } catch (err) {
    setContextError(`Could not load context: ${err.message}`);
  }
}

// Open the HubSpot contact/create URL in the user's default browser
function openHubSpotLink() {
  const link = document.getElementById("hubspot-create-link");
  const url = link?.dataset?.url;
  if (!url) return;
  try {
    Office.context.ui.openBrowserWindow(url);
  } catch (e) {
    // Fallback for environments where openBrowserWindow isn't available
    window.open(url, "_blank");
  }
}

// Force-reload context from the server (clears pipeline cache too)
function refreshContext() {
  _pipelinesCache = null;
  const email = currentContext?.recipient?.email;
  if (!email) return;
  setContextLoading();
  loadContext(email);
  showToast("Refreshing context...");
}

// Render the full context card: contact info, deal stage badge, deal amount,
// last email, CRM notes, HubSpot connection status, and HubSpot links.
function renderContext(ctx) {
  const r       = ctx.recipient  || {};
  const deals   = ctx.deals      || [];
  const notes   = ctx.crmNotes   || [];
  const history = ctx.emailHistory || [];

  document.getElementById("context-dot").className = "context-dot loaded";

  const name = r.name || r.email || "Unknown";
  document.getElementById("context-title").textContent    = name;
  document.getElementById("context-subtitle").textContent = r.company || r.jobTitle || "";

  setText("ctx-name",    r.name    || "—");
  setText("ctx-company", r.company || "—");
  setText("ctx-title",   r.jobTitle || "—");

  // Reset CRM note row so it doesn't persist from previous contact
  document.getElementById("ctx-note-row").style.display = "none";

  // Deal stage badge — use DOM methods to avoid XSS from API data
  const stageEl = document.getElementById("ctx-stage");
  stageEl.textContent = "";
  if (deals.length > 0) {
    const d = deals[0];
    renderStageBadge(stageEl, d);
  } else {
    stageEl.textContent = "No deals";
  }
  // Multi-deal indicator: only the most recent deal is shown/movable.
  // Backend sorts by hs_lastmodifieddate DESC so deals[0] is the freshest.
  const moreRow = document.getElementById("ctx-more-deals-row");
  if (moreRow) moreRow.remove();
  if (deals.length > 1) {
    const dealValueRow = document.getElementById("ctx-deal")?.closest(".context-row");
    const row = document.createElement("div");
    row.id = "ctx-more-deals-row";
    row.className = "context-row";
    const label = document.createElement("span");
    label.className = "context-label";
    label.textContent = "";
    const val = document.createElement("span");
    val.className = "context-value";
    val.style.color    = "var(--color-text-muted)";
    val.style.fontSize = "11px";
    val.style.fontStyle = "italic";
    val.textContent = `+${deals.length - 1} more deal(s) — only the most recent is shown`;
    row.appendChild(label);
    row.appendChild(val);
    if (dealValueRow && dealValueRow.parentNode) {
      dealValueRow.parentNode.insertBefore(row, dealValueRow.nextSibling);
    }
  }
  // Ensure move panel is hidden whenever we re-render
  hideMoveStagePanel();

  // Deal info — treat 0 as valid amount (not falsy)
  if (deals.length > 0) {
    const d = deals[0];
    const amount    = d.amount != null ? `$${Number(d.amount).toLocaleString()}` : "";
    const closeDate = d.closeDate ? ` · closes ${d.closeDate}` : "";
    setText("ctx-deal", `${d.name || ""}${amount !== "" ? " · " + amount : ""}${closeDate}`);
  } else {
    setText("ctx-deal", "—");
  }

  // Last email
  if (history.length > 0) {
    const h = history[0];
    setText("ctx-last-email", `${h.date || ""} · ${h.subject || ""}`);
  } else {
    setText("ctx-last-email", "No previous emails");
  }

  // CRM note
  if (notes.length > 0) {
    document.getElementById("ctx-note").textContent        = `"${notes[0]}"`;
    document.getElementById("ctx-note-row").style.display  = "block";
  }

  // HubSpot connection status
  const hsRow    = document.getElementById("ctx-hs-status-row");
  const hsStatus = document.getElementById("ctx-hs-status");
  if (hsRow && hsStatus) {
    hsRow.style.display = "flex";
    if (r.hubspotId) {
      hsStatus.textContent  = "Connected ✓";
      hsStatus.className    = "context-value hs-connected";
    } else if (ctx.hubspotSearched) {
      // Server tried HubSpot but contact not found
      hsStatus.textContent  = "Not in HubSpot";
      hsStatus.className    = "context-value hs-missing";
    } else {
      // HubSpot not configured on server
      hsStatus.textContent  = "Not connected";
      hsStatus.className    = "context-value hs-missing";
    }
  }

  // HubSpot link — use portalId from context if available for correct URLs
  const wrap = document.getElementById("hubspot-link-wrap");
  const link = document.getElementById("hubspot-create-link");
  const portalId = ctx.hubspotPortalId || "";
  if (wrap) wrap.style.display = "inline";
  if (link) {
    if (r.hubspotId) {
      link.dataset.url = portalId
        ? `https://app.hubspot.com/contacts/${portalId}/contact/${r.hubspotId}`
        : `https://app.hubspot.com/contacts/${r.hubspotId}`;
      link.textContent = "View in HubSpot";
    } else {
      link.dataset.url = portalId
        ? `https://app.hubspot.com/contacts/${portalId}/contact/create?email=${encodeURIComponent(r.email || "")}`
        : `https://app.hubspot.com/contacts/create?email=${encodeURIComponent(r.email || "")}`;
      link.textContent = "Create in HubSpot";
    }
  }

  document.getElementById("context-skeleton").style.display = "none";
  document.getElementById("context-loaded").style.display   = "block";
  document.getElementById("context-error").style.display    = "none";

  // Hide deal panel by default — user opens it via "+ New Deal" link
  const dealSection = document.getElementById("deal-section");
  if (dealSection) {
    dealSection.style.display = "none";
  }

  // Auto-open on first load per recipient; respect explicit user close
  if (!_userClosedContext) openContextPanel();

  // Check auto-track status now that context loaded successfully (non-blocking)
  checkAutoTrackStatus();
}

// Show skeleton loading state in the context card while data is being fetched
function setContextLoading() {
  document.getElementById("context-dot").className = "context-dot loading";
  setText("context-title",    "Loading context...");
  setText("context-subtitle", "");
  document.getElementById("context-skeleton").style.display = "block";
  document.getElementById("context-loaded").style.display   = "none";
  document.getElementById("context-error").style.display    = "none";
  openContextPanel();
}

function setContextError(msg) {
  document.getElementById("context-dot").className = "context-dot error";
  setText("context-title",    "Context unavailable");
  setText("context-subtitle", "");
  document.getElementById("context-skeleton").style.display = "none";
  document.getElementById("context-loaded").style.display   = "none";
  document.getElementById("context-error").style.display    = "block";
  document.getElementById("context-error-msg").textContent  = msg;
  if (!_userClosedContext) openContextPanel();
}

// Show a neutral "waiting" state (e.g., "No recipient yet") - no skeleton, no error
function setContextPending(msg) {
  document.getElementById("context-dot").className = "context-dot";
  setText("context-title",    msg);
  setText("context-subtitle", "");
  document.getElementById("context-skeleton").style.display = "none";
  document.getElementById("context-loaded").style.display   = "none";
  document.getElementById("context-error").style.display    = "none";
}

// ── Context Card Toggle ───────────────────────────────────────
function toggleContext() {
  contextOpen ? closeContextPanel() : openContextPanel();
}

function openContextPanel() {
  contextOpen = true;
  document.getElementById("context-body").classList.add("open");
  document.getElementById("context-toggle").classList.add("open");
  document.getElementById("context-chevron").classList.add("open");
}

function closeContextPanel() {
  contextOpen        = false;
  _userClosedContext = true; // remember user's preference
  document.getElementById("context-body").classList.remove("open");
  document.getElementById("context-toggle").classList.remove("open");
  document.getElementById("context-chevron").classList.remove("open");
}


// ── Move Deal Stage ───────────────────────────────────────────
// Render a colored badge for the deal's current stage, with a "Move" button
// that opens a dropdown to change the stage without leaving the add-in.
function renderStageBadge(stageEl, deal) {
  if (!stageEl || !deal) return;
  stageEl.textContent = "";
  const span = document.createElement("span");
  span.className   = `badge ${getStageBadgeClass(deal.stage)}`;
  span.textContent = deal.stage || "Unknown";
  stageEl.appendChild(span);
  if (deal.id && deal.pipelineId) {
    const moveBtn = document.createElement("button");
    moveBtn.className   = "btn-link move-stage-link";
    moveBtn.textContent = "Move";
    moveBtn.title       = `Move "${deal.name || "deal"}" to another stage`;
    moveBtn.setAttribute("aria-haspopup", "listbox");
    moveBtn.setAttribute("aria-controls", "move-stage-row");
    moveBtn.setAttribute("aria-expanded", "false");
    moveBtn.onclick     = (e) => showMoveStagePanel(e.currentTarget);
    stageEl.appendChild(document.createTextNode(" "));
    stageEl.appendChild(moveBtn);
  }
}

// Open the move-stage panel: load pipeline stages into a <select> and let
// the user pick a new stage. Dedupes concurrent requests and validates data.
async function showMoveStagePanel(triggerEl) {
  const deal = currentContext?.deals?.[0];
  if (!deal || !deal.id || !deal.pipelineId) {
    showToast("Deal info missing — cannot move.", "error");
    return;
  }
  const row = document.getElementById("move-stage-row");
  const select = document.getElementById("move-stage-select");
  const confirmBtn = document.getElementById("move-stage-confirm");
  if (!row || !select) return;
  // FIX-3: dedupe — if already visible, don't re-run
  if (row.style.display === "flex") return;

  // FIX-13: token check BEFORE showing the row
  const token = getToken();
  if (!token) {
    showToast("Not signed in — please sign in and try again.", "error");
    return;
  }

  // Track which Move button opened the panel so we can restore focus on hide
  if (triggerEl) _moveLinkRef = triggerEl;
  if (_moveLinkRef) _moveLinkRef.setAttribute("aria-expanded", "true");

  row.style.display = "flex";
  select.innerHTML = '<option value="">Loading stages…</option>';
  select.disabled = true;
  if (confirmBtn) confirmBtn.disabled = true; // FIX-6

  try {
    if (!_pipelinesCache) {
      // FIX-3: dedupe in-flight requests across rapid clicks
      _pipelinesPromise ||= apiFetch(`${API_BASE}/hubspot/pipelines`, {
        headers: { "Authorization": `Bearer ${token}` },
      }).then(async r => {
        if (!r.ok) throw new Error(await parseApiError(r));
        return r.json();
      });
      let data;
      try {
        data = await _pipelinesPromise;
      } finally {
        _pipelinesPromise = null;
      }
      // FIX-5: validate before caching
      const valid = Array.isArray(data) && data.length > 0
        && data.every(p => Array.isArray(p.stages) && p.stages.length > 0);
      if (!valid) {
        throw new Error("Pipelines response invalid");
      }
      _pipelinesCache = data;
    }
    const pipeline = _pipelinesCache.find(p => p.id === deal.pipelineId);
    if (!pipeline || !pipeline.stages?.length) {
      throw new Error("Pipeline stages unavailable");
    }
    select.innerHTML = "";
    pipeline.stages.forEach(s => {
      const opt = document.createElement("option");
      opt.value = s.id;
      opt.textContent = s.label;
      if (s.id === deal.stageId) opt.selected = true;
      select.appendChild(opt);
    });
    select.disabled = false;
    if (confirmBtn) confirmBtn.disabled = false; // FIX-6

    // FIX-4: focus + keyboard support
    select.focus();
    if (_moveStageKeyHandler) {
      select.removeEventListener("keydown", _moveStageKeyHandler);
    }
    _moveStageKeyHandler = function (e) {
      if (e.key === "Enter") {
        e.preventDefault();
        confirmMoveStage();
      } else if (e.key === "Escape") {
        e.preventDefault();
        hideMoveStagePanel();
      }
    };
    select.addEventListener("keydown", _moveStageKeyHandler);
  } catch (err) {
    select.innerHTML = '<option value="">Failed to load</option>';
    if (confirmBtn) confirmBtn.disabled = true; // FIX-6: leave disabled on failure
    showToast(`Could not load stages: ${err.message}`, "error");
  }
}

function hideMoveStagePanel() {
  const row = document.getElementById("move-stage-row");
  if (row) row.style.display = "none";
  const select = document.getElementById("move-stage-select");
  if (select) {
    if (_moveStageKeyHandler) {
      select.removeEventListener("keydown", _moveStageKeyHandler);
    }
    select.innerHTML = ""; // FIX-15
  }
  // FIX-4: restore focus and ARIA on the trigger
  if (_moveLinkRef) {
    _moveLinkRef.setAttribute("aria-expanded", "false");
    try { _moveLinkRef.focus(); } catch (e) {}
  }
}

// Send the stage change to HubSpot via PATCH and update the UI optimistically
async function confirmMoveStage() {
  const deal = currentContext?.deals?.[0];
  const select = document.getElementById("move-stage-select");
  const btn = document.getElementById("move-stage-confirm");
  if (!deal?.id || !select) return;
  const newStageId = select.value;
  if (!newStageId) {
    showToast("Pick a stage first.", "error");
    return;
  }
  if (newStageId === deal.stageId) {
    showToast("Deal is already at that stage.");
    hideMoveStagePanel();
    return;
  }
  const token = getToken();
  if (!token) return;

  const origLabel = btn?.textContent;
  if (btn) { btn.disabled = true; btn.textContent = "Moving…"; }
  try {
    const res = await apiFetch(`${API_BASE}/hubspot/deal/${encodeURIComponent(deal.id)}/stage`, {
      method: "PATCH",
      headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({ stageId: newStageId }),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) {
      if (res.status === 429) {
        throw new Error("Too many requests — wait a minute and try again.");
      }
      throw new Error(data.error || `Server error ${res.status}`);
    }

    // Update local state so the UI is consistent without a full re-fetch
    deal.stageId = data.stageId;
    deal.stage   = data.stageLabel || deal.stage;

    // Re-render the stage badge inline via shared helper
    const stageEl = document.getElementById("ctx-stage");
    renderStageBadge(stageEl, deal);
    hideMoveStagePanel();
    showToast(`Moved "${deal.name}" to "${deal.stage}".`, "success");
  } catch (err) {
    showToast(`Move failed: ${err.message}`, "error");
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = origLabel || "Move"; }
  }
}

// ── Auto-Track ────────────────────────────────────────────────
// Auto-track uses server-side Graph webhooks to automatically log emails
// to HubSpot deals. The checkbox state is cached in localStorage for
// instant UI rendering, then confirmed with the server.
async function checkAutoTrackStatus() {
  const token = getToken();
  if (!token) return;

  try {
    const row = document.getElementById("auto-track-row");
    const cb  = document.getElementById("auto-track-checkbox");
    if (!row || !cb) return;

    // Show the row immediately with cached state
    const cached = localStorage.getItem("imdath_autotrack");
    if (cached === "true") cb.checked = true;
    row.style.display = "flex";

    const res = await apiFetch(`${API_BASE}/auto-track/status`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (res.ok) {
      const data = await res.json();
      cb.checked = data.enabled;
      localStorage.setItem("imdath_autotrack", String(data.enabled));
    }
  } catch (e) {
    console.warn("[Imdath] checkAutoTrackStatus failed:", e.message);
  }
}

// Enable or disable auto-tracking. Disables the checkbox during the API call
// to prevent double-toggling, and reverts on failure.
async function toggleAutoTrack(enabled) {
  const token = getToken();
  const cb = document.getElementById("auto-track-checkbox");
  if (!token) {
    if (cb) cb.checked = false;
    showToast("Please sign in first.", "error");
    return;
  }

  // Disable checkbox during the API call to prevent double-toggling
  if (cb) cb.disabled = true;

  try {
    const endpoint = enabled ? "/auto-track/enable" : "/auto-track/disable";
    // Send the refresh token to the server. The server will do the refresh
    // itself — we must NOT refresh client-side first because Microsoft rotates
    // tokens on each use, and a token can only be used once.
    // Server already has the refresh token from auth/exchange — no need to send it
    const res = await apiFetch(`${API_BASE}${endpoint}`, {
      method:  "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error || `Server error ${res.status}`);
    }
    const data = await res.json();
    if (cb) cb.checked = data.enabled;
    localStorage.setItem("imdath_autotrack", String(data.enabled));
    showToast(data.enabled ? "Auto-tracking enabled" : "Auto-tracking disabled", "success");
  } catch (e) {
    console.error("[Imdath] toggleAutoTrack failed:", e);
    if (cb) cb.checked = !enabled; // revert
    showToast("Could not toggle auto-tracking: " + e.message, "error");
  } finally {
    if (cb) cb.disabled = false;
  }
}

"use strict";
/* Sphinx Compose - AI email generation, refinement chat, reply mode,
   Outlook body insertion, voice input, HubSpot logging, and clipboard copy.
   This is the largest frontend module. It manages the compose tab's full lifecycle.
   Depends on: utils.js, auth.js, context.js (currentContext). */

// ── State (compose) ───────────────────────────────────────────
var selectedTone   = localStorage.getItem("imdath_tone") || "Professional"; // Persisted tone preference
var isGenerating   = false;          // Guards against double-clicking Generate
var isRefining     = false;          // Guards against double-clicking Apply in chat
var _hasGenerated  = false;          // True after first successful generation (enables Regenerate label)
var _cachedSignature = null;         // User's email signature (extracted or fetched), cached in memory
var _replyMode       = false;        // True when composing a reply (not a new email)
var _replyThreadSubject = "";        // Subject of the thread being replied to (without Re:/Fw: prefix)
var _replyConversationId = null;     // Outlook conversation ID for the reply thread
var _replyAll      = false;          // True for Reply All, false for Reply
var _replyMessageId = null;          // Graph API message ID for sending replies from read mode
var _hubspotBtnTimer = null;         // Timer to reset the "Log to HubSpot" button after success
var PENDING_DRAFT_KEY = "imdath_pending_reply_draft"; // localStorage key for cross-pane draft persistence
var PENDING_DRAFT_TTL_MS = 10 * 60 * 1000;           // Pending drafts expire after 10 minutes

// ── Tone Selection ────────────────────────────────────────────
// Tone pills (Professional, Friendly, etc.) control the AI's writing style.
// Selection is persisted in localStorage and restored on next open.
function selectTone(pill) {
  document.querySelectorAll(".tone-pill").forEach((p) => p.classList.remove("selected"));
  pill.classList.add("selected");
  selectedTone = pill.dataset.tone;
  localStorage.setItem("imdath_tone", selectedTone);
}

// Restore the previously selected tone pill from localStorage on init
function restoreTone() {
  const saved = localStorage.getItem("imdath_tone");
  if (!saved) return;
  const VALID_TONES = ["Professional","Friendly","Formal","Concise","Persuasive"];
  if (!VALID_TONES.includes(saved)) return;
  const pill = document.querySelector(`.tone-pill[data-tone="${saved}"]`);
  if (pill) {
    document.querySelectorAll(".tone-pill").forEach((p) => p.classList.remove("selected"));
    pill.classList.add("selected");
    selectedTone = saved;
  }
}

// ── Email Generation ──────────────────────────────────────────
// Sends context + user prompt + tone to the server's /api/generate endpoint.
// The server uses Claude to compose a full email (subject + body).
async function generateEmail() {
  if (isGenerating) return;

  // Warn before overwriting a refined draft
  if (_hasGenerated) {
    const bodyHasContent = document.getElementById("body-textarea").value.trim().length > 0;
    if (bodyHasContent && !confirm("This will overwrite your current draft. Continue?")) return;
  }

  const customNote = document.getElementById("custom-notes").value.trim();
  if (!customNote) {
    showToast("Tell me the goal — type a prompt or use the mic.", "error");
    return;
  }
  if (customNote.length > 2000) {
    showToast("Prompt is too long — keep it under 2000 characters.", "error");
    return;
  }
  if (!currentContext) {
    showToast("Add a recipient first — context needs to load before generating.", "error");
    return;
  }

  const token = getToken();
  if (!token) return;

  const senderName = Office.context?.mailbox?.userProfile?.displayName || "";

  // FIX-9: Clear stale refinement chat on new generation
  const chatHistoryEl = document.getElementById("chat-history");
  if (chatHistoryEl) chatHistoryEl.innerHTML = "";
  _chatHistory = [];

  setGenerateLoading(true);

  try {
    const res = await apiFetch(`${API_BASE}/generate`, {
      method:  "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
      body:    JSON.stringify({ context: currentContext, tone: selectedTone, senderName, customNote, isReply: _replyMode, conversationId: _replyConversationId }),
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error || `Server error ${res.status}`);
    }

    const data = await res.json();
    document.getElementById("subject-input").value   = data.subject || "";
    document.getElementById("body-textarea").value   = data.body    || "";
    window._lastGeneratedBody = data.body || "";

    // Show sections that are hidden until first generation
    document.getElementById("preview-section").classList.add("visible");
    document.getElementById("chat-section").classList.add("visible");

    // Show feedback row; reset buttons for new generation
    const feedbackRow = document.getElementById("feedback-row");
    if (feedbackRow) {
      feedbackRow.style.display = "flex";
      ["feedback-up", "feedback-down"].forEach(id => {
        const btn = document.getElementById(id);
        if (btn) { btn.disabled = false; btn.classList.remove("active"); }
      });
    }

    try { document.getElementById("preview-section").scrollIntoView({ behavior: "smooth", block: "nearest" }); } catch (e) {}
    _hasGenerated = true;
    updateGenerateBtn();
    showToast("Email generated.", "success");
  } catch (err) {
    showToast(`Generation failed: ${err.message}`, "error");
  } finally {
    setGenerateLoading(false);
  }
}

// Toggle the Generate button between loading spinner and normal state
function setGenerateLoading(loading) {
  isGenerating = loading;
  const btn = document.getElementById("generate-btn");
  btn.disabled    = loading;
  btn.textContent = "";
  if (loading) {
    const spinner = document.createElement("span");
    spinner.className = "spinner";
    btn.appendChild(spinner);
    btn.appendChild(document.createTextNode(" Generating..."));
  } else {
    updateGenerateBtn();
  }
}

// Update Generate button label based on state: "Regenerate", "Generate Reply", or "Generate Email"
function updateGenerateBtn() {
  const btn = document.getElementById("generate-btn");
  if (_hasGenerated) {
    btn.textContent = "Regenerate";
  } else {
    btn.textContent = _replyMode ? "Generate Reply" : "Generate Email";
  }
}

// ── Signature Extraction ──────────────────────────────────────
// Outlook auto-injects the user's signature into the compose body. When we
// call setAsync to replace the body with generated content, we need to
// preserve that signature. These helpers extract it from the HTML so we
// can re-append it after our content.

// Strip the outer <html><head>...<body>...</body></html> wrapper that
// Outlook's getAsync returns, leaving only the inner body content.
function unwrapOutlookHtml(html) {
  if (!html) return "";
  const bodyMatch = html.match(/<body[^>]*>([\s\S]*)<\/body>/i);
  return bodyMatch ? bodyMatch[1].trim() : html;
}

// Try multiple heuristics to find the email signature in an HTML body.
// Different Outlook versions and email clients use different markers,
// so we check each pattern in order of specificity.
function extractSignature(html) {
  if (!html) return "";
  // OWA / New Outlook desktop: wraps signature in <div id="Signature">
  const owaMatch = html.match(/<div[^>]+id=["']Signature["'][^>]*>[\s\S]*/i);
  if (owaMatch) return owaMatch[0];
  // Some Outlook versions prefix the id: <div id="x_Signature">
  const prefixedMatch = html.match(/<div[^>]+id=["']x_Signature["'][^>]*>[\s\S]*/i);
  if (prefixedMatch) return prefixedMatch[0];
  // Class-based signatures (e.g. class="gmail_signature" or class="OutlookSignature")
  const classMatch = html.match(/<div[^>]+class=["'][^"']*signature[^"']*["'][^>]*>[\s\S]*/i);
  if (classMatch) return classMatch[0];
  // Classic Outlook: signature follows an <hr> divider
  const hrIdx = html.search(/<hr[\s>]/i);
  if (hrIdx !== -1) return html.slice(hrIdx);
  // Unix-style "-- " signature delimiter
  const unixIdx = html.indexOf("-- <br");
  if (unixIdx !== -1) return html.slice(unixIdx);
  return "";
}

// ── Insert into Outlook ───────────────────────────────────────
// Main action button handler. In compose mode, inserts the generated email
// into the active Outlook compose window. In read-mode reply, sends via Graph API.
function insertIntoOutlook() {
  // If we entered reply mode from a read-mode email, there's no compose
  // window to insert into — send the reply via Graph API instead.
  if (_replyMessageId) {
    _sendReplyViaGraph(_replyAll);
    return;
  }

  const subject = document.getElementById("subject-input").value;
  const body    = document.getElementById("body-textarea").value;

  const item = Office.context.mailbox.item;
  if (!item) {
    showToast("No active compose window.", "error");
    return;
  }

  item.subject.setAsync(subject, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      showToast("Could not set subject.", "error");
    }
  });

  // _doSetBody reads the current body, strips any previous imdath block,
  // and rebuilds as [generated content] + [everything else incl. signature].
  _doSetBody(item, body);
}

// Core body-insertion logic. Wraps generated content in an identifiable div
// so re-inserts can replace only our content without touching the signature.
// In reply mode, uses prependAsync to avoid scroll-jump issues.
function _doSetBody(item, body) {
  const generatedHtml = escapeHtml(body).replace(/\n/g, "<br>");
  // Wrap in identifiable div so we can replace on re-insert without touching the signature
  const wrappedHtml = '<div id="imdath-generated">' + generatedHtml + '</div>';

  // Reply mode: use prependAsync to insert at the top without replacing the whole
  // body. setAsync replaces everything, which causes Outlook to re-render and
  // scroll to the bottom of the quoted thread. prependAsync avoids this by only
  // adding content above the existing body, keeping the cursor at the top.
  if (_replyMode && item.body.prependAsync) {
    // First strip any previous imdath block, then prepend the new one
    item.body.getAsync(Office.CoercionType.Html, (getResult) => {
      if (getResult.status === Office.AsyncResultStatus.Succeeded) {
        const currentHtml = getResult.value || "";
        const hadPrevious = /<div[^>]+id=["']imdath-generated["']/i.test(currentHtml);
        if (hadPrevious) {
          // Re-insert: strip old block then setAsync (unavoidable full replace)
          let stripped = currentHtml.replace(/<div[^>]+id=["']imdath-generated["'][^>]*>[\s\S]*?<\/div>/i, "");
          const bodyInner = unwrapOutlookHtml(stripped);
          const finalHtml = wrappedHtml + bodyInner;
          item.body.setAsync(finalHtml, { coercionType: Office.CoercionType.Html }, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              showToast("Drafted in Outlook. Click Send to send.", "success");
            } else {
              _doSetBodyFallback(item, body, wrappedHtml);
            }
          });
          return;
        }
      }
      // First insert or getAsync failed — prepend (no scroll jump)
      item.body.prependAsync(wrappedHtml, { coercionType: Office.CoercionType.Html }, (r) => {
        if (r.status === Office.AsyncResultStatus.Succeeded) {
          showToast("Drafted in Outlook. Click Send to send.", "success");
        } else {
          console.warn("[Imdath] prependAsync failed:", r.error);
          _doSetBodyFallback(item, body, wrappedHtml);
        }
      });
    });
    return;
  }

  // New compose: read current body → strip any previous imdath-generated block →
  // rebuild as [new content] + [everything else including signature] → write back.
  // This preserves the signature because we round-trip whatever getAsync returns.
  item.body.getAsync(Office.CoercionType.Html, (getResult) => {
    if (getResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.warn("[Imdath] _doSetBody getAsync failed, falling back to setAsync");
      _doSetBodyFallback(item, body, wrappedHtml);
      return;
    }

    let currentHtml = getResult.value || "";
    console.log("[Imdath] _doSetBody current body length:", currentHtml.length);

    // Strip previous imdath-generated block if present (re-insert case)
    currentHtml = currentHtml.replace(/<div[^>]+id=["']imdath-generated["'][^>]*>[\s\S]*?<\/div>/i, "");

    // Extract inner body content (strip <html><head>…<body> wrapper)
    const bodyInner = unwrapOutlookHtml(currentHtml);

    // Combine: generated content first, then everything that was already there (signature etc.)
    const finalHtml = wrappedHtml + bodyInner;

    item.body.setAsync(finalHtml, { coercionType: Office.CoercionType.Html }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        showToast("Drafted in Outlook. Click Send to send.", "success");
      } else {
        console.warn("[Imdath] setAsync HTML failed:", result.error);
        _doSetBodyFallback(item, body, wrappedHtml);
      }
    });
  });
}

// Last-resort fallback — tries prependAsync (preserves signature natively) then setAsync
function _doSetBodyFallback(item, body, wrappedHtml) {
  if (item.body.prependAsync) {
    item.body.prependAsync(wrappedHtml, { coercionType: Office.CoercionType.Html }, (r) => {
      showToast(r.status === Office.AsyncResultStatus.Succeeded ? "Drafted in Outlook." : "Could not set body.",
        r.status === Office.AsyncResultStatus.Succeeded ? "success" : "error");
    });
  } else {
    item.body.setAsync(body, { coercionType: Office.CoercionType.Text }, (r2) => {
      showToast(r2.status === Office.AsyncResultStatus.Succeeded ? "Drafted in Outlook." : "Could not set body.",
        r2.status === Office.AsyncResultStatus.Succeeded ? "success" : "error");
    });
  }
}

// ── Voice Input ───────────────────────────────────────────────
// Opens an Office dialog for speech-to-text input. The dialog handles
// microphone access and Web Speech API, then sends the transcript back
// via messageParent. Used by both the notes prompt and the refinement chat.
function openVoiceDialog(btnId, onText) {
  const micBtn = document.getElementById(btnId);
  micBtn.disabled = true;

  const dialogUrl = window.location.origin + "/voice-dialog.html";

  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 40, width: 30, promptBeforeOpen: false },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        micBtn.disabled = false;
        showToast("Could not open voice dialog.", "error");
        return;
      }

      const dialog = result.value;
      micBtn.classList.add("listening");

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
        dialog.close();
        micBtn.classList.remove("listening");
        try {
          const msg = JSON.parse(arg.message);
          if (msg.type === "voice" && msg.text) {
            try {
              await onText(msg.text, micBtn);
            } catch (e) {
              console.error("[Imdath] Voice onText callback failed:", e);
              micBtn.disabled = false;
            }
          } else {
            micBtn.disabled = false;
          }
        } catch (e) {
          console.error("[Imdath] Voice dialog message parse failed:", e);
          micBtn.disabled = false;
        }
      });

      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
        micBtn.classList.remove("listening");
        micBtn.disabled = false;
      });
    }
  );
}

// Voice for the notes/prompt textarea - appends raw transcript directly
function toggleVoice() {
  openVoiceDialog("mic-btn", async (text, micBtn) => {
    const textarea = document.getElementById("custom-notes");
    const existing = textarea.value.trim();
    textarea.value = existing ? existing + " " + text : text;
    showToast("Voice captured.", "success");
    micBtn.disabled = false;
  });
}

// Voice for the refinement chat - sends transcript through /api/contextualize-voice
// to clean it up before placing it in the chat input. Falls back to raw text on error.
function toggleVoiceForChat() {
  openVoiceDialog("chat-mic-btn", async (text, micBtn) => {
    const input = document.getElementById("chat-input");
    input.value    = "Processing...";
    input.disabled = true;

    try {
      const token      = getToken();
      const controller = new AbortController();
      const timer      = setTimeout(() => controller.abort(), 12000);
      let res;
      try {
        res = await apiFetch(`${API_BASE}/contextualize-voice`, {
          method:  "POST",
          headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
          body:    JSON.stringify({ rawText: text }),
          signal:  controller.signal,
        });
      } finally {
        clearTimeout(timer);
      }
      if (!res.ok) throw new Error(await parseApiError(res));
      const data  = await res.json();
      input.value = data.text || text;
      showToast("Voice processed.", "success");
    } catch (e) {
      // Fall back to raw text if API fails or times out
      input.value = text;
      showToast("Voice captured.", "success");
    } finally {
      input.disabled  = false;
      micBtn.disabled = false;
    }
  });
}

// ── Feedback ──────────────────────────────────────────────────
// Thumbs up/down on generated emails. Sent to the server for quality tracking.
// Non-critical - failures are silently logged.
async function submitFeedback(positive) {
  const token = getToken();
  if (!token) return;

  // Visual confirmation
  const upBtn   = document.getElementById("feedback-up");
  const downBtn = document.getElementById("feedback-down");
  if (upBtn)   upBtn.disabled   = true;
  if (downBtn) downBtn.disabled = true;
  if (positive && upBtn)   upBtn.classList.add("active");
  if (!positive && downBtn) downBtn.classList.add("active");
  showToast(positive ? "Thanks for the feedback!" : "Got it — we'll improve.", "success");

  try {
    await apiFetch(`${API_BASE}/feedback`, {
      method:  "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
      body:    JSON.stringify({
        positive,
        tone:    selectedTone,
        hasEdit: document.getElementById("body-textarea").value.trim() !==
                 (window._lastGeneratedBody || ""),
      }),
    });
  } catch (e) {
    // Non-critical — don't surface to user
    console.warn("[Imdath] feedback failed:", e);
  }
}

// ── Copy to Clipboard ─────────────────────────────────────────
async function copyToClipboard() {
  const subject = document.getElementById("subject-input").value;
  const body    = document.getElementById("body-textarea").value;
  const text    = subject ? `Subject: ${subject}\n\n${body}` : body;
  try {
    await navigator.clipboard.writeText(text);
    showToast("Copied to clipboard.", "success");
  } catch {
    showToast("Copy failed — select text manually.", "error");
  }
}

// ── HubSpot Logging ───────────────────────────────────────────
// Manually log the current generated email as an activity on the HubSpot contact.
// Separate from auto-track, which logs emails automatically via webhooks.
async function logToHubSpot() {
  if (!currentContext?.recipient?.email) {
    showToast("No contact to log to.", "error");
    return;
  }

  const subject = document.getElementById("subject-input").value;
  const body    = document.getElementById("body-textarea").value;
  if (!subject && !body) {
    showToast("No email to log.", "error");
    return;
  }

  const token = getToken();
  const btn   = document.getElementById("hubspot-log-btn");
  btn.disabled    = true;
  btn.textContent = "Logging...";

  try {
    const res = await apiFetch(`${API_BASE}/hubspot/log`, {
      method:  "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
      body:    JSON.stringify({ contactEmail: currentContext.recipient.email, subject, body }),
    });

    if (!res.ok) throw new Error(await parseApiError(res));
    showToast("Logged to HubSpot.", "success");
    btn.textContent = "Logged ✓";

    clearTimeout(_hubspotBtnTimer);
    _hubspotBtnTimer = setTimeout(() => {
      btn.textContent = "Log to HubSpot";
      btn.disabled    = false;
    }, 3000);
  } catch (err) {
    showToast(`HubSpot log failed: ${err.message}`, "error");
    btn.textContent = "Log to HubSpot";
    btn.disabled    = false;
  }
}

// ── Refinement Chat ───────────────────────────────────────────
// A conversational interface for iteratively refining the generated email.
// The user types instructions ("make it shorter", "add a CTA") and the
// server applies them while maintaining multi-turn context.
var _chatHistory = []; // Accumulates user instructions so the server has full conversation context

// Enter sends the chat message; Shift+Enter allows newlines
function handleChatKey(event) {
  if (event.key === "Enter" && !event.shiftKey) {
    event.preventDefault();
    sendChatMessage();
  }
}

// Send a refinement instruction to /api/refine and update the email preview
async function sendChatMessage() {
  if (isRefining) return;

  const input   = document.getElementById("chat-input");
  const message = input.value.trim();
  if (!message) return;

  input.value = "";
  appendChatBubble(message, "user");

  const currentSubject = document.getElementById("subject-input").value;
  const currentBody    = document.getElementById("body-textarea").value;

  if (!currentSubject && !currentBody) {
    appendChatBubble("Generate an email first, then I can refine it.", "assistant");
    return;
  }

  const token = getToken();
  if (!token) return;

  if (!currentContext) {
    appendChatBubble("No context loaded — add a recipient first.", "assistant");
    return;
  }

  setRefineLoading(true);
  const bubbleId = `typing-bubble-${Date.now()}`;
  appendChatBubble("Updating...", "assistant", bubbleId);

  try {
    const res = await apiFetch(`${API_BASE}/refine`, {
      method:  "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
      body:    JSON.stringify({ context: currentContext, currentSubject, currentBody, userMessage: message, conversationId: _replyConversationId || null, chatHistory: _chatHistory }),
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error || `Server error ${res.status}`);
    }

    const data = await res.json();
    const typing = document.getElementById(bubbleId);
    if (typing) typing.remove();

    document.getElementById("subject-input").value = data.subject || currentSubject;
    document.getElementById("body-textarea").value = data.body    || currentBody;
    _chatHistory.push(message);
    if (data.chatReply) {
      appendChatBubble(data.chatReply, "assistant");
    } else {
      appendChatBubble("Done. Check the preview above.", "assistant");
    }
  } catch (err) {
    const typing = document.getElementById(bubbleId);
    if (typing) {
      typing.textContent = `Error: ${err.message}`;
    } else {
      appendChatBubble(`Error: ${err.message}`, "assistant");
    }
  } finally {
    setRefineLoading(false);
  }
}

// Add a chat bubble to the refinement history (user or assistant)
function appendChatBubble(text, role, id) {
  const history = document.getElementById("chat-history");
  const bubble  = document.createElement("div");
  bubble.className = `chat-bubble ${role}`;
  bubble.textContent = text;
  if (id) bubble.id = id;
  history.appendChild(bubble);
  history.scrollTop = history.scrollHeight;
}

function setRefineLoading(loading) {
  isRefining = loading;
  const btn = document.getElementById("chat-send-btn");
  btn.disabled    = loading;
  btn.textContent = loading ? "..." : "Apply";
}

// ── Read Mode Banner ──────────────────────────────────────────
// When viewing (not composing) an email, we show a banner with Reply/Reply All
// buttons. The user can generate a reply and send it via Graph API without
// ever opening a compose window.
function showReadModeBanner() {
  const banner = document.getElementById("read-mode-banner");
  if (banner) banner.style.display = "block";
}

function hideReadModeBanner() {
  const banner = document.getElementById("read-mode-banner");
  if (banner) banner.style.display = "none";
}

// Convert the generated plain-text body to safe HTML for Graph API replies
function _buildReplyHtmlBody() {
  if (!_hasGenerated) return "";
  const body = (document.getElementById("body-textarea")?.value || "").trim();
  if (!body) return "";
  return escapeHtml(body).replace(/\n/g, "<br>");
}

// Send a reply via the server's Graph API endpoint (used in read mode where
// there's no compose window). Resolves the signature from three tiers:
// in-memory cache, localStorage, or a server fetch from sent mail.
async function _sendReplyViaGraph(replyAll) {
  if (!_hasGenerated) { showToast("Generate a reply first.", "error"); return; }

  const insertBtn = document.querySelector("#preview-section .btn-primary");
  const origText  = insertBtn ? insertBtn.textContent : "";
  let sent = false;
  try {
    if (insertBtn) { insertBtn.disabled = true; insertBtn.textContent = "Sending…"; }

    const messageId = _replyMessageId;
    if (!messageId) throw new Error("No message ID — cannot send reply");
    const generatedHtml = _buildReplyHtmlBody();
    if (!generatedHtml) throw new Error("No email body to send");
    const token = getToken();
    if (!token) throw new Error("Not signed in");

    // 3-tier signature resolution: in-memory → localStorage → server fetch
    let signature = _cachedSignature || "";
    if (!signature) {
      try { signature = localStorage.getItem("imdath_signature") || ""; } catch (e) { /* private browsing */ }
    }
    if (!signature) {
      try {
        const sigRes = await apiFetch(`${API_BASE}/graph/signature`, {
          headers: { Authorization: `Bearer ${token}` },
        });
        if (sigRes.ok) {
          const sigData = await sigRes.json();
          signature = sigData.signature || "";
          if (signature) {
            _cachedSignature = signature;
            try { localStorage.setItem("imdath_signature", signature); } catch (e) { /* private browsing */ }
          }
        }
      } catch (e) { console.warn("[Imdath] Could not fetch signature from sent mail:", e); }
    }

    const htmlBody = generatedHtml + signature;

    const res = await apiFetch(`${API_BASE}/graph/reply`, {
      method:  "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
      body:    JSON.stringify({ messageId, htmlBody, replyAll }),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error || `Server error ${res.status}`);
    }
    sent = true;
    showToast(replyAll ? "Reply All sent!" : "Reply sent!", "success");
  } catch (e) {
    console.error("[Imdath] _sendReplyViaGraph failed:", e);
    showToast("Could not send reply: " + e.message, "error");
  } finally {
    if (insertBtn) {
      if (sent) {
        insertBtn.disabled = true;
        insertBtn.textContent = "Sent \u2713";
      } else {
        insertBtn.disabled = false;
        insertBtn.textContent = origText;
      }
    }
  }
}

// Enter reply mode from a read-mode email without opening a new compose tab.
// Converts the Outlook EWS item ID to a Graph API REST ID, activates reply
// mode UI, and prepares the "Send Reply" button.
function _enterReplyFromReadMode(replyAll) {
  const item = Office.context?.mailbox?.item;
  if (!item) { showToast("No email selected.", "error"); return; }

  _replyMode = true;
  _replyAll  = replyAll;
  _replyConversationId = item.conversationId || null;

  // item.itemId is an EWS ID — must convert to REST format for Graph API.
  // On Outlook iOS it's already REST-compatible; everywhere else (OWA, desktop)
  // we need convertToRestId().
  const rawId = item.itemId || null;
  if (rawId && Office.context.mailbox.convertToRestId) {
    try {
      _replyMessageId = Office.context.mailbox.convertToRestId(
        rawId, Office.MailboxEnums.RestVersion.v2_0
      );
    } catch (e) {
      console.warn("[Imdath] convertToRestId failed, using raw ID:", e);
      _replyMessageId = rawId;
    }
  } else {
    _replyMessageId = rawId;
  }

  // In read mode, item.subject is a plain string (not an object with getAsync)
  const subject = (typeof item.subject === "string")
    ? item.subject
    : (item.subject?.toString?.() || "");
  _replyThreadSubject = subject.replace(/^(re|fw):\s*/i, "");

  hideReadModeBanner();
  showReplyBanner();

  // Change the "Insert into Outlook" button to "Send Reply" since there is
  // no compose window — we'll send via Graph API when the user clicks it.
  const insertBtn = document.querySelector("#preview-section .btn-primary");
  if (insertBtn) {
    insertBtn.textContent = replyAll ? "Send Reply All" : "Send Reply";
  }
  const caption = document.querySelector("#preview-section .action-caption");
  if (caption) caption.textContent = "Sends the reply from your account via Outlook.";
}

function triggerReplyAll() {
  _enterReplyFromReadMode(true);
}

function triggerReply() {
  _enterReplyFromReadMode(false);
}

// ── Cross-Pane Draft Persistence ──────────────────────────────
// Problem: clicking Reply/Reply All in Outlook spawns a brand-new compose
// pane with its own taskpane instance, losing any generated draft.
// Solution: save the draft to localStorage keyed by conversationId before
// the reply form opens, then restore it when the new pane initializes.

// Save the current draft state so it survives a pane switch
function savePendingReplyDraft() {
  try {
    const item = Office.context?.mailbox?.item;
    const convId = item && item.conversationId;
    if (!convId) return;
    const subject = document.getElementById("subject-input")?.value || "";
    const body    = document.getElementById("body-textarea")?.value || "";
    const notes   = document.getElementById("custom-notes")?.value || "";
    // Only persist if there's something worth restoring.
    if (!subject && !body && !notes) return;
    const payload = {
      conversationId: convId,
      subject,
      body,
      notes,
      tone: selectedTone,
      hasGenerated: !!_hasGenerated,
      savedAt: Date.now(),
    };
    localStorage.setItem(PENDING_DRAFT_KEY, JSON.stringify(payload));
  } catch (e) {
    console.warn("[Imdath] savePendingReplyDraft failed:", e);
  }
}

// Check localStorage for a saved draft matching the current conversation.
// Restores subject, body, notes, tone, and UI state. Returns true if restored.
function restorePendingReplyDraft() {
  try {
    const raw = localStorage.getItem(PENDING_DRAFT_KEY);
    if (!raw) return false;
    const draft = JSON.parse(raw);
    // Expire stale drafts
    if (!draft || !draft.savedAt || (Date.now() - draft.savedAt) > PENDING_DRAFT_TTL_MS) {
      localStorage.removeItem(PENDING_DRAFT_KEY);
      return false;
    }
    const item = Office.context?.mailbox?.item;
    const convId = item && item.conversationId;
    if (!convId || convId !== draft.conversationId) return false;

    const subjectEl = document.getElementById("subject-input");
    const bodyEl    = document.getElementById("body-textarea");
    const notesEl   = document.getElementById("custom-notes");
    if (subjectEl) subjectEl.value = draft.subject || "";
    if (bodyEl)    bodyEl.value    = draft.body    || "";
    if (notesEl)   notesEl.value   = draft.notes   || "";
    window._lastGeneratedBody = draft.body || "";

    if (draft.tone) {
      selectedTone = draft.tone;
      try { localStorage.setItem("imdath_tone", selectedTone); } catch (e) {}
      try { restoreTone && restoreTone(); } catch (e) {}
    }

    if (draft.hasGenerated) {
      _hasGenerated = true;
      const preview = document.getElementById("preview-section");
      const chat    = document.getElementById("chat-section");
      if (preview) preview.classList.add("visible");
      if (chat)    chat.classList.add("visible");
      const feedbackRow = document.getElementById("feedback-row");
      if (feedbackRow) feedbackRow.style.display = "flex";
      updateGenerateBtn();
    }

    localStorage.removeItem(PENDING_DRAFT_KEY);
    showToast("Restored your generated draft.", "success");
    return true;
  } catch (e) {
    console.warn("[Imdath] restorePendingReplyDraft failed:", e);
    try { localStorage.removeItem(PENDING_DRAFT_KEY); } catch (err) {}
    return false;
  }
}

// ── Reply Mode ────────────────────────────────────────────────
// In compose mode, detect whether we're in a reply (subject starts with Re:/Fw:)
// and activate reply mode UI accordingly.
function detectReplyMode() {
  if (!Office?.context?.mailbox?.item) return;
  const item = Office.context.mailbox.item;

  // Check if we're in a reply (item has conversationId and subject starts with Re:)
  if (item.conversationId && item.subject) {
    item.subject.getAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) return;
      const subject = result.value || "";
      if (/^(re|fw):/i.test(subject)) {
        _replyMode = true;
        _replyThreadSubject = subject.replace(/^(re|fw):\s*/i, "");
        _replyConversationId = item.conversationId;
        showReplyBanner();
      } else {
        dismissReplyMode();
      }
    });
  } else {
    dismissReplyMode();
  }
}

function showReplyBanner() {
  const banner = document.getElementById("reply-mode-banner");
  if (!banner) return;
  banner.style.display = "block";
  const subjectEl = document.getElementById("reply-thread-subject");
  if (subjectEl) subjectEl.textContent = _replyThreadSubject || "Thread";
  updateGenerateBtn();
}

// Exit reply mode and restore all UI elements to their default new-email state
function dismissReplyMode() {
  _replyMode = false;
  _replyConversationId = null;
  _replyThreadSubject = "";
  _replyAll = false;
  _replyMessageId = null;
  const banner = document.getElementById("reply-mode-banner");
  if (banner) banner.style.display = "none";
  // Restore the "Insert into Outlook" button if it was changed to "Send Reply"
  const insertBtn = document.querySelector("#preview-section .btn-primary");
  if (insertBtn) insertBtn.textContent = "Insert into Outlook";
  const caption = document.querySelector("#preview-section .action-caption");
  if (caption) caption.textContent = "Puts the draft in your compose window — you review before sending.";
  updateGenerateBtn();
}

// ── Suggest Response ──────────────────────────────────────────
// Reads the current email thread and asks the server to suggest what the
// user should say next. Populates the notes/prompt textarea so the user
// can review before generating.
async function suggestResponse() {
  const token = getToken();
  if (!token) return;

  const item = Office.context?.mailbox?.item;
  if (!item || !item.conversationId) {
    showToast("No email thread detected — open or reply to an email first.", "error");
    return;
  }

  const btn = document.getElementById("suggest-response-btn");
  const textarea = document.getElementById("custom-notes");
  const origText = btn.textContent;
  btn.disabled = true;
  btn.textContent = "...";

  try {
    const r = currentContext?.recipient || {};
    const res = await apiFetch(`${API_BASE}/suggest-response`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
      body: JSON.stringify({ conversationId: item.conversationId, recipientName: r.name, company: r.company }),
    });
    if (!res.ok) throw new Error(await parseApiError(res));
    const data = await res.json();
    if (data.suggestion) {
      textarea.value = data.suggestion;
      showToast("Response suggestion loaded.", "success");
    }
  } catch (err) {
    showToast(`Suggest failed: ${err.message}`, "error");
  } finally {
    btn.disabled = false;
    btn.textContent = origText;
  }
}

"use strict";
/* Sphinx CRM - HubSpot deal/contact/company creation and linking.
   Provides the UI panels for creating CRM records and linking email
   threads to existing deals. All CRM writes go through the server
   which proxies to the HubSpot API.
   Depends on: utils.js, auth.js, context.js (currentContext, refreshContext, _pipelinesCache). */

// ── State (crm) ───────────────────────────────────────────────
var _isCreatingDeal    = false;  // Guards against double-submit on deal creation
var _isCreatingContact = false;  // Guards against double-submit on contact creation
var _isCreatingCompany = false;  // Guards against double-submit on company creation
var _dealPropsCache    = null;   // Cached dynamic deal properties from HubSpot
var _linkDealSnapshot  = null;   // Frozen email state when Link panel opened (prevents stale links)
var _dealSearchTimer   = null;   // Debounce timer for deal search input
var _companySearchTimer = null;  // Debounce timer for company search input

// ── Create Deal ───────────────────────────────────────────────
// Show the deal creation form, loading pipeline/stage options and dynamic
// deal properties from HubSpot.
function showDealPanel() {
  const section = document.getElementById("deal-section");
  if (!section) return;
  section.style.display = "block";
  loadPipelines();
  loadDealProperties();
  section.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

// Fetch custom deal properties configured in HubSpot and render form fields
async function loadDealProperties() {
  const token = getToken();
  if (!token) return;
  const container = document.getElementById("deal-dynamic-fields");
  if (!container) return;

  if (_dealPropsCache) {
    renderDealProperties(_dealPropsCache, container);
    return;
  }

  container.innerHTML = '<div style="text-align:center; padding:12px; color:var(--color-text-muted); font-size:11px">Loading properties...</div>';

  try {
    const res = await apiFetch(`${API_BASE}/hubspot/deal-properties`, {
      headers: { "Authorization": `Bearer ${token}` },
    });
    if (!res.ok) throw new Error(await parseApiError(res));
    _dealPropsCache = await res.json();
    renderDealProperties(_dealPropsCache, container);
  } catch (err) {
    container.innerHTML = '<div style="text-align:center; padding:8px; color:var(--color-text-muted); font-size:11px">Could not load properties</div>';
    console.warn("[deal] loadDealProperties failed:", err.message);
  }
}

// Render dynamic deal property fields into the form. Small field types
// (number, date) are paired into 2-column rows for compact layout.
function renderDealProperties(props, container) {
  container.innerHTML = "";
  if (!props || props.length === 0) return;

  // Group into pairs for 2-column layout where field types are small
  const smallTypes = new Set(["number", "date", "datetime"]);
  let i = 0;
  while (i < props.length) {
    const p = props[i];
    const isSmall = smallTypes.has(p.type) && p.fieldType !== "textarea";
    const next = (i + 1 < props.length) ? props[i + 1] : null;
    const nextIsSmall = next && smallTypes.has(next.type) && next.fieldType !== "textarea";

    if (isSmall && nextIsSmall) {
      // Render as 2-column row
      const row = document.createElement("div");
      row.className = "deal-row-2col";
      row.appendChild(buildFieldGroup(p));
      row.appendChild(buildFieldGroup(next));
      container.appendChild(row);
      i += 2;
    } else {
      container.appendChild(buildFieldGroup(p));
      i += 1;
    }
  }

  // Set default close date to 30 days from now
  const closeDateEl = container.querySelector('[data-deal-prop="closedate"]');
  if (closeDateEl && !closeDateEl.value) {
    const d = new Date(Date.now() + 30 * 86400000);
    closeDateEl.value = d.toISOString().substring(0, 10);
  }
}

// Build a label + input element pair for a single deal property.
// Supports text, number, date, select (enumeration), boolean, and textarea.
function buildFieldGroup(prop) {
  const group = document.createElement("div");
  group.className = "field-group";

  const label = document.createElement("label");
  label.className = "field-label";
  label.textContent = prop.label;
  label.setAttribute("for", `deal-prop-${prop.name}`);
  group.appendChild(label);

  let input;
  if (prop.type === "enumeration" || (prop.options && prop.options.length > 0)) {
    input = document.createElement("select");
    const emptyOpt = document.createElement("option");
    emptyOpt.value = "";
    emptyOpt.textContent = `Select ${prop.label}...`;
    input.appendChild(emptyOpt);
    prop.options.forEach(o => {
      const opt = document.createElement("option");
      opt.value = o.value;
      opt.textContent = o.label;
      input.appendChild(opt);
    });
  } else if (prop.type === "bool") {
    input = document.createElement("select");
    const emptyOpt = document.createElement("option");
    emptyOpt.value = "";
    emptyOpt.textContent = `Select...`;
    input.appendChild(emptyOpt);
    ["true", "false"].forEach(v => {
      const opt = document.createElement("option");
      opt.value = v;
      opt.textContent = v === "true" ? "Yes" : "No";
      input.appendChild(opt);
    });
  } else if (prop.fieldType === "textarea") {
    input = document.createElement("textarea");
    input.rows = 2;
    input.placeholder = prop.description || "";
  } else if (prop.type === "number") {
    input = document.createElement("input");
    input.type = "number";
    input.step = "any";
    input.placeholder = "0";
  } else if (prop.type === "date" || prop.type === "datetime") {
    input = document.createElement("input");
    input.type = "date";
  } else {
    input = document.createElement("input");
    input.type = "text";
    input.placeholder = prop.description || "";
  }

  input.id = `deal-prop-${prop.name}`;
  input.setAttribute("data-deal-prop", prop.name);
  group.appendChild(input);
  return group;
}

// Fetch HubSpot pipelines and their stages (cached in _pipelinesCache)
async function loadPipelines() {
  const token = getToken();
  if (!token) return;

  const pipelineSelect = document.getElementById("deal-pipeline-select");
  const stageSelect    = document.getElementById("deal-stage-select");

  if (_pipelinesCache) {
    renderPipelineOptions(_pipelinesCache);
    return;
  }

  pipelineSelect.innerHTML = '<option value="">Loading...</option>';
  stageSelect.innerHTML    = '<option value="">Select pipeline first</option>';

  try {
    const res = await apiFetch(`${API_BASE}/hubspot/pipelines`, {
      headers: { "Authorization": `Bearer ${token}` },
    });
    if (!res.ok) throw new Error(await parseApiError(res));
    _pipelinesCache = await res.json();
    renderPipelineOptions(_pipelinesCache);
  } catch (err) {
    pipelineSelect.innerHTML = '<option value="">Failed to load</option>';
    showToast(`Could not load pipelines: ${err.message}`, "error");
  }
}

function renderPipelineOptions(pipelines) {
  const select = document.getElementById("deal-pipeline-select");
  select.innerHTML = "";

  if (!pipelines || pipelines.length === 0) {
    select.innerHTML = '<option value="">No pipelines found</option>';
    return;
  }

  pipelines.forEach((p, i) => {
    const opt = document.createElement("option");
    opt.value = p.id;
    opt.textContent = p.label;
    if (i === 0) opt.selected = true;
    select.appendChild(opt);
  });

  // Auto-populate stages for first pipeline
  onPipelineChange();
}

// When the pipeline dropdown changes, repopulate the stage dropdown with
// the stages belonging to the newly selected pipeline.
function onPipelineChange() {
  const pipelineId = document.getElementById("deal-pipeline-select").value;
  const stageSelect = document.getElementById("deal-stage-select");
  stageSelect.innerHTML = "";

  if (!pipelineId || !_pipelinesCache) {
    stageSelect.innerHTML = '<option value="">Select pipeline first</option>';
    return;
  }

  const pipeline = _pipelinesCache.find(p => p.id === pipelineId);
  if (!pipeline || !pipeline.stages || pipeline.stages.length === 0) {
    stageSelect.innerHTML = '<option value="">No stages</option>';
    return;
  }

  pipeline.stages.forEach((s, i) => {
    const opt = document.createElement("option");
    opt.value = s.id;
    opt.textContent = s.label;
    if (i === 0) opt.selected = true;
    stageSelect.appendChild(opt);
  });
}

// Ask the server to AI-suggest a deal name based on the contact's info
async function suggestDealName() {
  const token = getToken();
  if (!token || !currentContext?.recipient) return;

  const btn = document.getElementById("suggest-deal-btn");
  const input = document.getElementById("deal-name-input");
  btn.disabled = true;
  btn.textContent = "...";

  try {
    const r = currentContext.recipient;
    const res = await apiFetch(`${API_BASE}/suggest-deal-name`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
      body: JSON.stringify({ contactName: r.name, company: r.company, jobTitle: r.jobTitle }),
    });
    if (!res.ok) throw new Error(await parseApiError(res));
    const data = await res.json();
    if (data.name) input.value = data.name;
  } catch (err) {
    showToast(`Suggest failed: ${err.message}`, "error");
  } finally {
    btn.disabled = false;
    btn.textContent = "\u2605";
  }
}

async function createDeal() {
  if (_isCreatingDeal) return;
  const token = getToken();
  if (!token) return;

  const contactEmail = currentContext?.recipient?.email;
  if (!contactEmail) {
    showToast("No contact loaded — add a recipient first.", "error");
    return;
  }

  const dealname  = document.getElementById("deal-name-input").value.trim();
  if (!dealname) {
    showToast("Deal name is required.", "error");
    return;
  }

  const pipeline  = document.getElementById("deal-pipeline-select").value;
  const dealstage = document.getElementById("deal-stage-select").value;
  const notes     = document.getElementById("deal-notes").value.trim() || undefined;

  // Collect all dynamic property fields
  const body = { contactEmail, dealname, pipeline, dealstage, notes };
  document.querySelectorAll("[data-deal-prop]").forEach(el => {
    const val = el.value.trim();
    if (val) body[el.getAttribute("data-deal-prop")] = val;
  });

  _isCreatingDeal = true;
  const btn = document.getElementById("create-deal-btn");
  btn.disabled = true;
  btn.textContent = "";
  const spinner = document.createElement("span");
  spinner.className = "spinner";
  btn.appendChild(spinner);
  btn.appendChild(document.createTextNode(" Creating..."));

  try {
    const res = await apiFetch(`${API_BASE}/hubspot/deal`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
      body: JSON.stringify(body),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error || `Server error ${res.status}`);
    }
    const data = await res.json();
    showToast(`Deal "${data.dealname}" created in HubSpot.`, "success");
    btn.textContent = "Created \u2713";

    // FIX-6: Clear form and hide panel after successful creation
    clearDealForm();
    setTimeout(() => {
      const section = document.getElementById("deal-section");
      if (section) section.style.display = "none";
      btn.textContent = "Create Deal";
      btn.disabled = false;
    }, 1500);

    // Refresh context to show new deal
    refreshContext();
  } catch (err) {
    showToast(`Create deal failed: ${err.message}`, "error");
    btn.textContent = "Create Deal";
    btn.disabled = false;
  } finally {
    _isCreatingDeal = false;
  }
}


// ── Link to Existing Deal ─────────────────────────────────────
// ── Link to Existing Deal ─────────────────────────────────────
// (declared as var at top of this file)

function showLinkDealPanel() {
  const section = document.getElementById("link-deal-section");
  if (!section) return;

  // Freeze the current email's identity — all link operations will use this
  // snapshot instead of reading live state at click time.
  const item = Office.context?.mailbox?.item;
  if (!item || !item.conversationId) {
    showToast("No email thread open.", "error");
    return;
  }

  // In compose mode, item.subject is a Subject object (not a string) and
  // item.from is undefined. Read subject via getAsync when needed, and fall
  // back to currentContext for the contact email.
  const rawSubject = item.subject;
  const subjectIsString = (typeof rawSubject === "string");
  const contactEmail = item.from?.emailAddress || currentContext?.recipient?.email || "";

  const buildSnapshot = (subjectStr) => {
    _linkDealSnapshot = {
      conversationId: item.conversationId,
      subject:        subjectStr || "",
      contactEmail:   contactEmail,
      itemId:         item.itemId || null,
    };

    // Hide create deal panel if open
    const createSection = document.getElementById("deal-section");
    if (createSection) createSection.style.display = "none";

    section.style.display = "block";
    document.getElementById("deal-search-input").value = "";
    document.getElementById("deal-search-results").innerHTML = "";
    section.scrollIntoView({ behavior: "smooth", block: "nearest" });

    // Load recent deals immediately
    searchDealsForLink("");
  };

  if (subjectIsString) {
    buildSnapshot(rawSubject);
  } else if (rawSubject && typeof rawSubject.getAsync === "function") {
    rawSubject.getAsync((asyncResult) => {
      const subj = asyncResult.status === Office.AsyncResultStatus.Succeeded
        ? asyncResult.value
        : "";
      buildSnapshot(subj);
    });
  } else {
    buildSnapshot(rawSubject?.toString?.() || "");
  }
}

function hideLinkDealPanel() {
  _linkDealSnapshot = null;
  const section = document.getElementById("link-deal-section");
  if (section) section.style.display = "none";
}

function hideContactPanel() {
  const section = document.getElementById("contact-section");
  if (section) section.style.display = "none";
}

function hideCompanyPanel() {
  const section = document.getElementById("company-section");
  if (section) section.style.display = "none";
}

function onDealSearchInput(input) {
  clearTimeout(_dealSearchTimer);
  _dealSearchTimer = setTimeout(() => searchDealsForLink(input.value.trim()), 300);
}

async function searchDealsForLink(query) {
  const token = getToken();
  if (!token) return;

  const container = document.getElementById("deal-search-results");
  container.innerHTML = '<div style="padding:8px;font-size:11px;color:var(--color-text-muted)">Searching...</div>';

  try {
    const res = await apiFetch(`${API_BASE}/hubspot/deals/search?q=${encodeURIComponent(query)}`, {
      headers: { "Authorization": `Bearer ${token}` },
    });
    if (!res.ok) throw new Error(await parseApiError(res));
    const deals = await res.json();
    renderDealSearchResults(deals);
  } catch (err) {
    container.innerHTML = `<div style="padding:8px;font-size:11px;color:var(--color-error)">${escapeHtml(err.message)}</div>`;
  }
}

function renderDealSearchResults(deals) {
  const container = document.getElementById("deal-search-results");
  container.innerHTML = "";

  if (deals.length === 0) {
    container.innerHTML = '<div style="padding:8px;font-size:11px;color:var(--color-text-muted)">No deals found.</div>';
    return;
  }

  // Header showing result count so users know how many deals are visible
  const header = document.createElement("div");
  header.className = "deal-result-count";
  header.textContent = `${deals.length} deal${deals.length === 1 ? "" : "s"}`;
  container.appendChild(header);

  deals.forEach(d => {
    const item = document.createElement("div");
    item.className = "deal-result-item";

    const info = document.createElement("div");
    info.className = "deal-result-info";

    const name = document.createElement("div");
    name.className = "deal-result-name";
    name.textContent = d.name;

    const meta = document.createElement("div");
    meta.className = "deal-result-meta";
    const parts = [d.stage];
    if (d.amount != null) parts.push(`$${Number(d.amount).toLocaleString()}`);
    if (d.closeDate) parts.push(`closes ${d.closeDate}`);
    meta.textContent = parts.join(" · ");

    info.appendChild(name);
    info.appendChild(meta);

    const btnWrap = document.createElement("div");
    btnWrap.className = "deal-result-btns";

    const btn = document.createElement("button");
    btn.className = "deal-result-link-btn";
    btn.textContent = "Link Email";
    btn.addEventListener("click", function () { linkEmailToDealUI(d.id, d.name, btn); });

    const btn2 = document.createElement("button");
    btn2.className = "deal-result-link-btn deal-result-link-contact-btn";
    btn2.textContent = "Link Contacts";
    btn2.addEventListener("click", function () { linkContactToDeal(d.id, d.name, btn2); });

    btnWrap.appendChild(btn);
    btnWrap.appendChild(btn2);

    item.appendChild(info);
    item.appendChild(btnWrap);
    container.appendChild(item);
  });
}

// ── Link Email Thread to Deal ──────────────────────────────────
// Uses _linkDealSnapshot to prevent stale state from linking the wrong thread.
// Server-side subject validation catches cross-thread contamination.
async function linkEmailToDealUI(dealId, dealName, btn) {
  const token = getToken();
  if (!token) { showToast("Not signed in.", "error"); return; }

  // Use frozen snapshot — never read live item state for the link operation
  if (!_linkDealSnapshot) {
    showToast("Email changed — please reopen the link panel.", "error");
    return;
  }

  // Guard: verify user is still on the same email
  const item = Office.context?.mailbox?.item;
  if (item && item.itemId && _linkDealSnapshot.itemId && item.itemId !== _linkDealSnapshot.itemId) {
    showToast("You switched emails. Please reopen the link panel.", "error");
    _linkDealSnapshot = null;
    hideLinkDealPanel();
    return;
  }

  const { conversationId, subject, contactEmail } = _linkDealSnapshot;

  btn.disabled = true;
  btn.textContent = "...";

  try {
    const doLink = async (force) => {
      const payload = { dealId, conversationId, contactEmail, subject };
      if (force) payload.force = true;
      const res = await apiFetch(`${API_BASE}/hubspot/deal/link-email`, {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
        body: JSON.stringify(payload),
      });
      return res;
    };

    let res = await doLink(false);

    // Handle validation warnings (422) — show inline confirmation
    if (res.status === 422) {
      const errData = await res.json().catch(() => ({}));
      if (errData.code === "NO_CONTACT_OVERLAP" || errData.code === "SUBJECT_MISMATCH") {
        btn.textContent = "Link Email";
        btn.disabled = false;

        var warningMessage;
        var warningDetail = "";
        if (errData.code === "SUBJECT_MISMATCH") {
          warningMessage = "This looks like a different conversation. Are you sure you want to link it to \u201C" + dealName + "\u201D?";
          var existingSubj = (errData.existingSubjects || []).slice(0, 2);
          if (existingSubj.length > 0) {
            warningDetail = "Deal thread: \u201C" + existingSubj[0].replace(/^(re|fw|fwd):\s*/i, "") + "\u201D";
          }
        } else {
          warningMessage = "No email participants are contacts on \u201C" + dealName + "\u201D. Wrong deal?";
          var contacts = (errData.dealContacts || []).join(", ");
          if (contacts) warningDetail = "Deal contacts: " + contacts;
        }

        // Show warning card above search results
        var userConfirmed = await new Promise(function (resolve) {
          var prev = document.getElementById("link-overlap-warning");
          if (prev) prev.remove();

          var warn = document.createElement("div");
          warn.id = "link-overlap-warning";
          warn.style.cssText = "background:#fff4f0;border:1px solid #d83b01;border-radius:6px;padding:10px 12px;margin:8px 0;";
          var msg = document.createElement("div");
          msg.style.cssText = "font-size:12px;color:#d83b01;margin-bottom:8px;line-height:1.4;";
          msg.textContent = warningMessage;
          if (warningDetail) {
            var sub = document.createElement("div");
            sub.style.cssText = "font-size:11px;color:#666;margin-top:4px;";
            sub.textContent = warningDetail;
            msg.appendChild(sub);
          }
          warn.appendChild(msg);
          var btnRow = document.createElement("div");
          btnRow.style.cssText = "display:flex;gap:8px;";
          var forceBtn = document.createElement("button");
          forceBtn.textContent = "Link Anyway";
          forceBtn.style.cssText = "font-size:12px;padding:5px 14px;background:#d83b01;color:white;border:none;border-radius:4px;cursor:pointer;font-weight:600;";
          var cancelBtn = document.createElement("button");
          cancelBtn.textContent = "Cancel";
          cancelBtn.style.cssText = "font-size:12px;padding:5px 14px;background:#e8e8e8;color:#333;border:none;border-radius:4px;cursor:pointer;";
          forceBtn.addEventListener("click", function () { warn.remove(); resolve(true); });
          cancelBtn.addEventListener("click", function () { warn.remove(); resolve(false); });
          btnRow.appendChild(forceBtn);
          btnRow.appendChild(cancelBtn);
          warn.appendChild(btnRow);

          var resultsContainer = document.getElementById("deal-search-results");
          if (resultsContainer) {
            resultsContainer.insertBefore(warn, resultsContainer.firstChild);
            warn.scrollIntoView({ behavior: "smooth", block: "nearest" });
          }
        });

        if (!userConfirmed) return;
        btn.disabled = true;
        btn.textContent = "...";
        res = await doLink(true);
      }
    }

    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error || `Server error ${res.status}`);
    }
    const data = await res.json().catch(() => ({}));
    const count = data.linked || 1;
    btn.textContent = "Linked ✓";
    showToast(`${count} email${count === 1 ? "" : "s"} linked to "${dealName}".`, "success");
    refreshContext();
    setTimeout(function () { btn.textContent = "Link Email"; btn.disabled = false; }, 3000);
  } catch (err) {
    showToast(`Link failed: ${err.message}`, "error");
    btn.textContent = "Link Email";
    btn.disabled = false;
  }
}

// ── Gather All Email Participants ─────────────────────────────
function gatherAllParticipants() {
  const item = Office.context?.mailbox?.item;
  if (!item) return [];

  const emails = new Set();

  // Sender
  if (item.from?.emailAddress) {
    emails.add(item.from.emailAddress.toLowerCase());
  }

  // To recipients (item.to is an array of EmailAddressDetails in read mode)
  if (Array.isArray(item.to)) {
    item.to.forEach(function (r) {
      if (r.emailAddress) emails.add(r.emailAddress.toLowerCase());
    });
  }

  // CC recipients (item.cc is an array of EmailAddressDetails in read mode)
  if (Array.isArray(item.cc)) {
    item.cc.forEach(function (r) {
      if (r.emailAddress) emails.add(r.emailAddress.toLowerCase());
    });
  }

  // Exclude the signed-in user's own email
  var userEmail = Office.context?.mailbox?.userProfile?.emailAddress;
  if (userEmail) emails.delete(userEmail.toLowerCase());

  return Array.from(emails);
}

// ── Link Contacts to Deal ─────────────────────────────────────
async function linkContactToDeal(dealId, dealName, btn) {
  const token = getToken();
  if (!token) { showToast("Not signed in.", "error"); return; }

  var contactEmails = gatherAllParticipants();
  if (contactEmails.length === 0) {
    // Fallback: try context or sender
    var fallback = currentContext?.recipient?.email || Office.context?.mailbox?.item?.from?.emailAddress || "";
    if (!fallback) {
      showToast("No participants found — open an email first.", "error");
      return;
    }
    contactEmails = [fallback.toLowerCase()];
  }

  btn.disabled = true;
  btn.textContent = "...";

  try {
    const res = await apiFetch(`${API_BASE}/hubspot/deal/link`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
      body: JSON.stringify({ contactEmails, dealId }),
    });
    const data = await res.json().catch(function () { return {}; });
    if (!res.ok) {
      throw new Error(data.error || `Server error ${res.status}`);
    }
    const linked = data.linked || 0;
    const already = data.alreadyLinked || 0;
    const total = linked + already;

    btn.textContent = `Linked ${total} \u2713`;

    var parts = [];
    if (linked > 0) parts.push(linked + " new");
    if (already > 0) parts.push(already + " already on deal");
    showToast(parts.join(", ") + ` on "${dealName}".`, "success");

    refreshContext();
    setTimeout(function () { btn.textContent = "Link Contacts"; btn.disabled = false; }, 3000);
  } catch (err) {
    showToast(`Link failed: ${err.message}`, "error");
    btn.textContent = "Link Contacts";
    btn.disabled = false;
  }
}


// ── Create Contact ────────────────────────────────────────────
// ── Create Contact ────────────────────────────────────────────
// (declared as var at top of this file)

function showContactPanel() {
  clearContactForm();
  // Hide other panels
  const dealSection = document.getElementById("deal-section");
  if (dealSection) dealSection.style.display = "none";
  const linkSection = document.getElementById("link-deal-section");
  if (linkSection) linkSection.style.display = "none";
  const companySection = document.getElementById("company-section");
  if (companySection) companySection.style.display = "none";

  const section = document.getElementById("contact-section");
  if (!section) return;
  section.style.display = "block";

  // Pre-fill email from current recipient if not in HubSpot
  const email = currentContext?.recipient?.email || "";
  const emailInput = document.getElementById("contact-email");
  if (emailInput && !emailInput.value) emailInput.value = email;

  // Pre-fill name parts from context
  const name = currentContext?.recipient?.name || "";
  if (name) {
    const parts = name.split(" ");
    const fn = document.getElementById("contact-firstname");
    const ln = document.getElementById("contact-lastname");
    if (fn && !fn.value) fn.value = parts[0] || "";
    if (ln && !ln.value) ln.value = parts.slice(1).join(" ") || "";
  }

  section.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function onCompanySearchInput(input) {
  clearTimeout(_companySearchTimer);
  const q = input.value.trim();
  if (q.length < 2) {
    document.getElementById("company-search-results").innerHTML = "";
    return;
  }
  _companySearchTimer = setTimeout(() => searchCompaniesForContact(q), 300);
}

async function searchCompaniesForContact(query) {
  const token = getToken();
  if (!token) return;
  const container = document.getElementById("company-search-results");
  container.innerHTML = '<div style="padding:6px;font-size:11px;color:var(--color-text-muted)">Searching...</div>';

  try {
    const res = await apiFetch(`${API_BASE}/hubspot/companies/search?q=${encodeURIComponent(query)}`, {
      headers: { "Authorization": `Bearer ${token}` },
    });
    if (!res.ok) throw new Error(await parseApiError(res));
    const companies = await res.json();
    renderCompanyResults(companies);
  } catch (err) {
    container.innerHTML = `<div style="padding:6px;font-size:11px;color:var(--color-error)">${escapeHtml(err.message)}</div>`;
  }
}

function renderCompanyResults(companies) {
  const container = document.getElementById("company-search-results");
  container.innerHTML = "";
  if (companies.length === 0) {
    container.innerHTML = '<div style="padding:6px;font-size:11px;color:var(--color-text-muted)">No companies found.</div>';
    return;
  }
  companies.forEach(c => {
    const item = document.createElement("div");
    item.className = "deal-result-item";
    const info = document.createElement("div");
    info.className = "deal-result-info";
    const nameEl = document.createElement("div");
    nameEl.className = "deal-result-name";
    nameEl.textContent = c.name;
    const meta = document.createElement("div");
    meta.className = "deal-result-meta";
    meta.textContent = c.domain || "";
    info.appendChild(nameEl);
    info.appendChild(meta);
    const btn = document.createElement("button");
    btn.className = "deal-result-link-btn";
    btn.textContent = "Select";
    btn.onclick = () => selectCompany(c.id, c.name);
    item.appendChild(info);
    item.appendChild(btn);
    container.appendChild(item);
  });
}

function selectCompany(companyId, companyName) {
  document.getElementById("contact-company-id").value = companyId;
  document.getElementById("company-search-results").innerHTML = "";
  document.getElementById("contact-company-search").value = "";
  const badge = document.getElementById("selected-company");
  badge.textContent = companyName;
  badge.style.display = "inline-block";
}

async function createContact() {
  if (_isCreatingContact) return;
  const token = getToken();
  if (!token) return;

  const email = document.getElementById("contact-email").value.trim();
  if (!email || !email.includes("@")) {
    showToast("Valid email is required.", "error");
    return;
  }

  const firstname = document.getElementById("contact-firstname").value.trim();
  const lastname  = document.getElementById("contact-lastname").value.trim();
  const jobtitle  = document.getElementById("contact-jobtitle").value.trim();
  const companyId = document.getElementById("contact-company-id").value || undefined;
  const company   = document.getElementById("contact-company-search").value.trim() || undefined;

  _isCreatingContact = true;
  const btn = document.getElementById("create-contact-btn");
  btn.disabled = true;
  btn.textContent = "Creating...";

  try {
    const res = await apiFetch(`${API_BASE}/hubspot/contact`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
      body: JSON.stringify({ email, firstname, lastname, company, jobtitle, companyId }),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error || `Server error ${res.status}`);
    }
    showToast(`Contact "${firstname} ${lastname}" created in HubSpot.`, "success");
    btn.textContent = "Created \u2713";

    // FIX-6: Clear form and hide panel after successful creation
    clearContactForm();
    setTimeout(() => {
      const section = document.getElementById("contact-section");
      if (section) section.style.display = "none";
      btn.textContent = "Create Contact";
      btn.disabled = false;
    }, 1500);

    refreshContext();
  } catch (err) {
    showToast(`Create contact failed: ${err.message}`, "error");
    btn.textContent = "Create Contact";
    btn.disabled = false;
  } finally {
    _isCreatingContact = false;
  }
}


// ── Create Company ────────────────────────────────────────────
// ── Create Company ────────────────────────────────────────────
function showCreateCompanyPanel() {
  const section = document.getElementById("company-section");
  if (!section) return;
  section.style.display = "block";
  section.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

async function createCompany() {
  if (_isCreatingCompany) return;
  const token = getToken();
  if (!token) return;

  const name   = document.getElementById("company-name").value.trim();
  const domain = document.getElementById("company-domain").value.trim();
  if (!name && !domain) {
    showToast("Company name or domain is required.", "error");
    return;
  }

  _isCreatingCompany = true;
  const btn = document.getElementById("create-company-btn");
  btn.disabled = true;
  btn.textContent = "Creating...";

  try {
    const res = await apiFetch(`${API_BASE}/hubspot/company`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
      body: JSON.stringify({ name: name || domain, domain }),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error || `Server error ${res.status}`);
    }
    const data = await res.json();
    showToast(`Company "${data.name}" created in HubSpot.`, "success");
    btn.textContent = "Created \u2713";

    // Auto-select the new company in the contact form
    if (data.companyId) {
      selectCompany(data.companyId, data.name);
    }

    // FIX-6: Clear form and hide panel after successful creation
    clearCompanyForm();
    setTimeout(() => {
      const section = document.getElementById("company-section");
      if (section) section.style.display = "none";
      btn.textContent = "Create Company";
      btn.disabled = false;
    }, 1500);
  } catch (err) {
    showToast(`Create company failed: ${err.message}`, "error");
    btn.textContent = "Create Company";
    btn.disabled = false;
  } finally {
    _isCreatingCompany = false;
  }
}

// ── Form Reset Helpers ────────────────────────────────────────
function clearDealForm() {
  const nameInput = document.getElementById("deal-name-input");
  if (nameInput) nameInput.value = "";
  const notesInput = document.getElementById("deal-notes");
  if (notesInput) notesInput.value = "";
  document.querySelectorAll("[data-deal-prop]").forEach(el => { el.value = ""; });
  // Reset pipeline/stage selects to first option
  const pipelineSelect = document.getElementById("deal-pipeline-select");
  if (pipelineSelect && pipelineSelect.options.length > 0) pipelineSelect.selectedIndex = 0;
  onPipelineChange();
}

function clearContactForm() {
  ["contact-firstname", "contact-lastname", "contact-email", "contact-jobtitle", "contact-company-search", "contact-company-id"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = "";
  });
  const badge = document.getElementById("selected-company");
  if (badge) { badge.style.display = "none"; badge.textContent = ""; }
  const results = document.getElementById("company-search-results");
  if (results) results.innerHTML = "";
}

function clearCompanyForm() {
  ["company-name", "company-domain"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = "";
  });
}

"use strict";
/* Sphinx Follow-ups — follow-up tab, draft generation */

// ── State (follow-ups) ────────────────────────────────────────
var _followupsLoaded = false;
var _followupDeals   = [];
var _followupTarget  = null;
var _followupResolvedContact = null;
var _followupReplyMeta = null;
var _followupSending   = false;
var _followupDraftGeneration = 0;
var _followupRefining = false;
var STALE_DAYS = 7;

// ── Tab Switching ─────────────────────────────────────────────
function switchTab(tab) {
  document.getElementById("tab-compose").style.display    = tab === "compose"  ? "block" : "none";
  document.getElementById("tab-followups").style.display  = tab === "followups" ? "block" : "none";
  document.getElementById("tab-btn-compose").classList.toggle("active",    tab === "compose");
  document.getElementById("tab-btn-followups").classList.toggle("active",  tab === "followups");
  document.getElementById("tab-btn-compose").setAttribute("aria-selected",    tab === "compose");
  document.getElementById("tab-btn-followups").setAttribute("aria-selected",  tab === "followups");
  // FIX-10: Clear follow-up draft when switching away from follow-ups tab
  if (tab !== "followups") hideFollowupDraft();
  if (tab === "followups" && !_followupsLoaded) loadFollowups();
}

// ── Follow-ups ────────────────────────────────────────────────
async function loadFollowups() {
  const token = getToken();
  if (!token) return;

  const list = document.getElementById("followup-list");
  list.innerHTML = '<div class="analytics-skeleton"><div class="skeleton skeleton-line full"></div><div class="skeleton skeleton-line medium"></div><div class="skeleton skeleton-line full"></div><div class="skeleton skeleton-line medium"></div><div class="skeleton skeleton-line full"></div></div>';
  setText("stat-total-deals", "…");
  setText("stat-overdue",     "…");
  setText("stat-stalest",     "…");

  const btn = document.getElementById("followup-refresh-btn");
  btn.disabled = true;

  try {
    const res = await apiFetch(`${API_BASE}/followups`, {
      headers: { "Authorization": `Bearer ${token}` },
    });
    if (!res.ok) throw new Error(await parseApiError(res));
    const data = await res.json();
    _followupDeals = data.deals || [];

    const overdue = _followupDeals.filter(d => d.daysSince >= STALE_DAYS);
    const stalest = _followupDeals.length > 0 ? _followupDeals[0].daysSince : 0;

    setText("stat-total-deals", String(_followupDeals.length));
    setText("stat-overdue",     String(overdue.length));
    setText("stat-stalest",     stalest >= 999 ? "N/A" : stalest + "d");

    renderFollowupList(_followupDeals);
    _followupsLoaded = true;
  } catch (err) {
    list.innerHTML = `<div class="empty-state"><span class="icon">⚠️</span>${escapeHtml(err.message)}</div>`;
    setText("stat-total-deals", "—");
    setText("stat-overdue",     "—");
    setText("stat-stalest",     "—");
  } finally {
    btn.disabled = false;
  }
}

// formatDaysAgo is in utils.js (loaded before this file)

function renderFollowupList(deals) {
  const list = document.getElementById("followup-list");
  if (!list) return;

  if (deals.length === 0) {
    list.innerHTML = '<div class="empty-state"><span class="icon">✅</span>No deals found in HubSpot.</div>';
    return;
  }

  list.innerHTML = "";
  list.setAttribute("role", "list");
  deals.forEach(function (deal) {
    const isOverdue = deal.daysSince >= STALE_DAYS;
    const item = document.createElement("div");
    item.className = "followup-item" + (isOverdue ? " overdue" : "");
    item.setAttribute("role", "listitem");

    // Age badge
    const age = document.createElement("div");
    age.className = "followup-age" + (isOverdue ? " urgent" : "");
    age.textContent = formatDaysAgo(deal.daysSince);

    // Body
    const body = document.createElement("div");
    body.className = "followup-body";

    const name = document.createElement("div");
    name.className = "followup-deal-name";
    name.textContent = deal.name;

    const meta = document.createElement("div");
    meta.className = "followup-meta";
    meta.textContent = deal.stage + (deal.amount ? " · $" + Number(deal.amount).toLocaleString() : "");

    body.appendChild(name);
    body.appendChild(meta);

    // Follow up button
    const btn = document.createElement("button");
    btn.className = "followup-action-btn";
    btn.textContent = "Follow up";
    btn.title = "Draft a follow-up email for this deal";
    btn.setAttribute("aria-label", "Follow up on " + deal.name);
    btn.onclick = function () { startFollowupDraft(deal, btn); };

    item.appendChild(age);
    item.appendChild(body);
    item.appendChild(btn);
    list.appendChild(item);
  });
}

// (declared as var at top of this file)

async function startFollowupDraft(deal, clickedBtn) {
  _followupTarget = deal;
  _followupReplyMeta = null;
  _followupDraftGeneration++;
  const thisGeneration = _followupDraftGeneration;
  const section = document.getElementById("followup-draft-section");
  section.style.display = "block";
  try { section.scrollIntoView({ behavior: "smooth", block: "start" }); } catch (e) {}

  document.getElementById("followup-subject").value = "";
  document.getElementById("followup-body").value = "Generating follow-up...";
  document.getElementById("followup-insert-btn").disabled = true;
  document.getElementById("followup-insert-btn").style.display = "";
  document.getElementById("followup-reply-info").style.display = "none";

  // FIX-8: Disable the clicked "Follow up" button and show loading state
  if (clickedBtn) {
    clickedBtn.disabled = true;
    clickedBtn.textContent = "Drafting...";
  }

  const token = getToken();
  if (!token) {
    if (clickedBtn) { clickedBtn.disabled = false; clickedBtn.textContent = "Follow up"; }
    return;
  }

  try {
    // Use the first contact on the deal (enriched by GET /api/followups)
    const contact = (deal.contacts || [])[0] || {};
    const contactEmail = contact.email || null;
    const contactName  = contact.name || null;

    const res = await apiFetch(`${API_BASE}/followups/draft`, {
      method: "POST",
      headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        dealId: deal.id,
        dealName: deal.name,
        dealStage: deal.stage,
        daysSince: deal.daysSince,
        contactEmail: contactEmail,
        contactName: contactName,
        company: deal.company || "",
      }),
    });

    if (!res.ok) throw new Error(await parseApiError(res));
    const data = await res.json();

    // FIX-5: Discard stale response if user switched deals
    if (thisGeneration !== _followupDraftGeneration) return;

    document.getElementById("followup-subject").value = data.subject || "";
    document.getElementById("followup-body").value = data.body || "";
    document.getElementById("followup-insert-btn").disabled = false;

    // Store resolved contact for compose window
    _followupResolvedContact = data.resolvedContact || null;

    // Show reply routing info
    _followupReplyMeta = data.replyMeta || null;
    const infoEl = document.getElementById("followup-reply-info");
    if (_followupReplyMeta) {
      const threadSubject = _followupReplyMeta.subject || deal.name || "thread";
      infoEl.innerHTML =
        '<span class="followup-route-label">Replying in thread:</span> ' + escapeHtml(threadSubject);
      infoEl.style.display = "block";
      document.getElementById("followup-insert-btn").textContent = "Draft in Thread";
    } else {
      infoEl.innerHTML = '<em>No prior email thread found — will open a new compose window.</em>';
      infoEl.style.display = "block";
      document.getElementById("followup-insert-btn").textContent = "Open Compose";
    }
  } catch (err) {
    document.getElementById("followup-body").value = "Error: " + err.message;
    document.getElementById("followup-insert-btn").disabled = false;
  } finally {
    // FIX-8: Re-enable the "Follow up" button
    if (clickedBtn) { clickedBtn.disabled = false; clickedBtn.textContent = "Follow up"; }
  }
}

// (declared as var at top of this file)

async function refineFollowupDraft(instruction) {
  if (!instruction || !instruction.trim()) return;
  if (_followupRefining) return;
  _followupRefining = true;

  const bodyEl = document.getElementById("followup-body");
  const subjectEl = document.getElementById("followup-subject");
  const chips = document.querySelectorAll(".followup-chip");
  const customInput = document.getElementById("followup-custom-instruction");

  chips.forEach(function(c) { c.disabled = true; });
  const prevBody = bodyEl.value;
  bodyEl.value = "Refining...";

  try {
    const token = getToken();
    if (!token) throw new Error("Not signed in");

    const res = await apiFetch(`${API_BASE}/refine`, {
      method: "POST",
      headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        context: "Follow-up email for deal: " + (_followupTarget ? _followupTarget.name : ""),
        currentSubject: subjectEl.value,
        currentBody: prevBody,
        userMessage: instruction.trim(),
      }),
    });

    if (!res.ok) throw new Error(await parseApiError(res));
    const data = await res.json();

    bodyEl.value = data.body || prevBody;
    if (data.subject) subjectEl.value = data.subject;
    if (customInput) customInput.value = "";
  } catch (err) {
    bodyEl.value = prevBody;
    showToast("Refine error: " + err.message, "error");
  } finally {
    _followupRefining = false;
    chips.forEach(function(c) { c.disabled = false; });
  }
}

function hideFollowupDraft() {
  document.getElementById("followup-draft-section").style.display = "none";
  _followupTarget = null;
  _followupReplyMeta = null;
  _followupResolvedContact = null;
}

async function insertFollowupIntoOutlook() {
  if (_followupSending) return;
  _followupSending = true;

  const subject = document.getElementById("followup-subject").value;
  const body    = document.getElementById("followup-body").value;

  if (!body) {
    showToast("No draft to insert", "error");
    _followupSending = false;
    return;
  }

  const btn = document.getElementById("followup-insert-btn");
  btn.disabled = true;
  const origText = btn.textContent;
  btn.innerHTML = '<span class="spinner"></span> Creating draft...';

  try {
    const token = getToken();
    if (!token) throw new Error("Not signed in");

    const htmlBody = body.replace(/\n/g, "<br>");

    if (_followupReplyMeta && _followupReplyMeta.messageId) {
      // Create a reply-all draft in the thread, then open it
      const res = await apiFetch(`${API_BASE}/graph/create-reply-draft`, {
        method: "POST",
        headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({
          messageId: _followupReplyMeta.messageId,
          htmlBody: htmlBody,
        }),
      });

      if (!res.ok) {
        const errData = await res.json().catch(function() { return {}; });
        throw new Error(errData.error || "Failed to create draft");
      }

      const data = await res.json();

      showToast("Draft created — check your Drafts folder", "success");
      hideFollowupDraft();
    } else {
      // No prior thread — create a new draft via Graph API
      const contact = _followupResolvedContact || (_followupTarget?.contacts || [])[0] || {};
      const res = await apiFetch(`${API_BASE}/graph/create-draft`, {
        method: "POST",
        headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({
          to: contact.email ? [contact.email] : [],
          subject: subject,
          htmlBody: htmlBody,
        }),
      });

      if (!res.ok) {
        const errData = await res.json().catch(function() { return {}; });
        throw new Error(errData.error || "Failed to create draft");
      }

      showToast("Draft created — check your Drafts folder", "success");
      hideFollowupDraft();
    }
  } catch (err) {
    console.error("[followup] insertFollowupIntoOutlook error:", err);
    showToast("Error: " + err.message, "error");
  } finally {
    _followupSending = false;
    btn.disabled = false;
    btn.textContent = origText;
  }
}
"use strict";
/* Sphinx Init — Office.onReady entry point */

// ── State (init) ──────────────────────────────────────────────
var _itemChangedHandler = null;

// ── Office Initialization ─────────────────────────────────────
Office.onReady(function () {
  try {
    restorePaneSize();
    restoreTone();
    silentSignIn()
      .then((authed) => {
        if (authed) {
          showMainContent();
          try { initRecipientDetection(); } catch (e) { console.warn("[Imdath] initRecipientDetection failed:", e); }
        }
      })
      .catch((e) => { console.warn("[Imdath] silentSignIn failed:", e); });

    if (Office.context && Office.context.mailbox) {
      try {
        // Remove previous handler before adding to prevent stacking on reload
        if (_itemChangedHandler) {
          try { Office.context.mailbox.removeHandlerAsync(Office.EventType.ItemChanged, { handler: _itemChangedHandler }); } catch (e) {}
        }
        _itemChangedHandler = () => {
          _lastLoadedEmail   = null;
          currentContext     = null; // prevent stale recipient email from leaking into link operations
          _linkDealSnapshot  = null; // invalidate any pending link operation
          _cachedSignature   = null; // new compose window — signature may differ
          _hasGenerated      = false;
          _replyMessageId    = null;
          _replyAll          = false;
          _allRecipients     = [];
          _pipelinesCache    = null; // reload pipelines for new context
          dismissReplyMode();
          hideReadModeBanner();
          renderRecipientChips();
          // FIX-9: Clear chat history on recipient change
          const chatHistory = document.getElementById("chat-history");
          if (chatHistory) chatHistory.innerHTML = "";
          // Hide all CRM panels
          ["deal-section", "link-deal-section", "contact-section", "company-section", "move-stage-row"].forEach(id => {
            const el = document.getElementById(id);
            if (el) el.style.display = "none";
          });
          updateGenerateBtn();
          if (accessToken) {
            // Small delay to let Outlook update the item reference
            setTimeout(() => {
              try { initRecipientDetection(); } catch (e) { console.warn("[Imdath] initRecipientDetection on ItemChanged failed:", e); }
            }, 200);
          }
        };
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, _itemChangedHandler);
      } catch (e) { console.warn("[Imdath] addHandlerAsync failed:", e); }
    }
  } catch (e) {
    console.error("[Imdath] Office.onReady error:", e);
  }
});
