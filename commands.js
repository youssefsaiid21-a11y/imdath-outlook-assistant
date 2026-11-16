/* ============================================================
   Sphinx — Event-based Command Handler
   Runs in a hidden background context (commands.html).
   Handles auto-launch events so the taskpane opens
   automatically when the user starts composing.
   ============================================================ */

"use strict";

// Auto-open the Sphinx taskpane when a new message compose window opens
function onNewMessageComposeHandler(event) {
  Office.addin.showAsTaskpane()
    .then(() => event.completed())
    .catch((err) => {
      console.warn("[Sphinx] Auto-launch failed:", String(err));
      event.completed();
    });
}

// Register the handler so Office can invoke it by name
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
