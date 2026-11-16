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
