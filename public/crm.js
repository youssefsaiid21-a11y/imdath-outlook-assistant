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

// Submit the deal creation form to HubSpot. Collects all static and dynamic
// fields, creates the deal, then refreshes context to show it in the card.
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
// Links the current email thread's conversationId to a HubSpot deal.
// Uses a "snapshot" pattern: we freeze the email's identity (conversationId,
// subject, itemId) when the panel opens so the link operation can't
// accidentally target the wrong email if the user switches messages.

// Open the link-deal panel and freeze the current email's identity
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

// Search HubSpot deals by name for the link panel
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

// Render search results with "Link Email" and "Link Contacts" buttons per deal
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
