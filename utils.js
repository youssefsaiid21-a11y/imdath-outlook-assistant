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
