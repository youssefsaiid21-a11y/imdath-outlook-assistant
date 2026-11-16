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
