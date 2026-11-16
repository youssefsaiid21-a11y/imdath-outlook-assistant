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
