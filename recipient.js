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
