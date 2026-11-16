"use strict";
/* Sphinx Follow-ups - the follow-up tab that lists deals needing attention.
   Shows deals sorted by staleness (days since last email), lets users generate
   AI follow-up drafts, refine them, and create reply drafts via Graph API.
   Depends on: utils.js (formatDaysAgo, escapeHtml, showToast, setText),
   auth.js (getToken, apiFetch), compose.js (shared refine endpoint). */

// ── State (follow-ups) ────────────────────────────────────────
var _followupsLoaded = false;          // True after first successful load (prevents re-fetch on tab switch)
var _followupDeals   = [];             // Array of deal objects from /api/followups
var _followupTarget  = null;           // The deal currently being drafted for
var _followupResolvedContact = null;   // Server-resolved best contact for the follow-up
var _followupReplyMeta = null;         // Thread info (messageId, subject) if replying in an existing thread
var _followupSending   = false;        // Guards against double-click on insert/send
var _followupDraftGeneration = 0;      // Incremented per draft request - discards stale responses
var _followupRefining = false;         // Guards against double-click on refine chips
var STALE_DAYS = 7;                    // Deals with no email for 7+ days are marked "overdue"

// ── Tab Switching ─────────────────────────────────────────────
// Toggle between the "Compose" and "Follow-ups" tabs. Lazy-loads follow-up
// data on first switch. Cleans up draft UI when leaving the follow-ups tab.
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
// Fetch all deals from the server and render the list with analytics summary
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

// Render the deal list with age badges, stage/amount info, and action buttons
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

// Generate an AI follow-up draft for a specific deal. The server resolves
// the best contact and checks for an existing email thread to reply in.
// Uses a generation counter to discard stale responses if the user clicks
// another deal before the first draft finishes.
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

// Refine the follow-up draft using quick-action chips ("shorter", "more formal")
// or custom instructions. Reuses the same /api/refine endpoint as the compose chat.
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

// Create the follow-up as a draft in the user's mailbox via Graph API.
// Two paths: if replyMeta exists, create a reply draft in the existing thread;
// otherwise, create a brand-new draft addressed to the resolved contact.
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
