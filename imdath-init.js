"use strict";
/* Sphinx Init - Office.onReady entry point.
   This file MUST be loaded LAST because it calls functions from every other
   module (restorePaneSize, restoreTone, silentSignIn, initRecipientDetection,
   dismissReplyMode, etc.). It bootstraps the add-in when Office.js is ready
   and registers a handler for when the user switches between emails. */

// ── State (init) ──────────────────────────────────────────────
var _itemChangedHandler = null; // Reference to the ItemChanged handler (for cleanup on reload)

// ── Office Initialization ─────────────────────────────────────
// Office.onReady fires once the Office.js host (Outlook) is fully loaded.
// This is the add-in's main entry point.
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
        // When the user switches emails, reset ALL per-email state to prevent
        // data from one email leaking into another's context or actions.
        _itemChangedHandler = () => {
          _lastLoadedEmail   = null;
          currentContext     = null;
          _linkDealSnapshot  = null;
          _cachedSignature   = null;
          _hasGenerated      = false;
          _replyMessageId    = null;
          _replyAll          = false;
          _allRecipients     = [];
          _pipelinesCache    = null;
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
            // 200ms delay gives Outlook time to update the internal item reference
            // before we try to read the new email's recipients/sender
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
