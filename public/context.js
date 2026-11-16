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
