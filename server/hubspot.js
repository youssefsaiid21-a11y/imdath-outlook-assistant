/**
 * hubspot.js - HubSpot CRM API client.
 *
 * Handles all HubSpot operations including:
 * - Contact lookup, creation, and search
 * - Deal CRUD, stage changes, and deal search
 * - Linking emails and contacts to deals (associations)
 * - Notes, pipelines, companies, and owner lookups
 * - Follow-up deal tracking (stale deals needing attention)
 *
 * All API calls go through a rate-limit queue to stay within HubSpot's
 * rolling 10-second limit, with automatic retry on 429 responses.
 */
"use strict";

const fetch = require("node-fetch");

const HUBSPOT_BASE    = "https://api.hubapi.com";
const TIMEOUT_MS      = Math.max(parseInt(process.env.HUBSPOT_TIMEOUT_MS || "15000", 10) || 15000, 1000);
// Max number of associated objects (deals, notes, contacts) to fetch per entity
const MAX_ASSOCIATIONS = Math.max(parseInt(process.env.HUBSPOT_MAX_ASSOCIATIONS || "5", 10) || 5, 1);

if (!process.env.HUBSPOT_TOKEN) {
  console.warn("[hubspot] HUBSPOT_TOKEN not set — HubSpot features will be disabled");
}

// ── HubSpot association type IDs ─────────────────────────────
// These are HubSpot-defined numeric IDs for associating objects together.
// See: https://developers.hubspot.com/docs/api/crm/associations
const ASSOC_CONTACT_TO_DEAL   = 4;
const ASSOC_EMAIL_TO_DEAL     = 210;
const ASSOC_EMAIL_TO_CONTACT  = 198;

// ── HubSpot property name constants ──────────────────────────
const CONTACT_PROPS = ["firstname", "lastname", "email", "company", "jobtitle", "lifecyclestage", "hubspot_owner_id", "hs_object_id"];
const DEAL_PROPS    = "dealname,dealstage,amount,closedate,pipeline,hs_lastmodifieddate";

// ── In-process caches (5-minute TTL) ────────────────────────
const CACHE_TTL = 5 * 60 * 1000;
let _pipelinesCache = null;
let _pipelinesCacheExpiresAt = 0;
let _dealPropsCache = null;
let _dealPropsCacheExpiresAt = 0;
let _portalIdCache = null;
let _portalIdCacheExpiresAt = 0;
const NOTE_PROPS    = "hs_note_body,hs_timestamp";

/** Strip HTML tags and decode common entities to plain text */
function stripHtml(html) {
  if (!html) return "";
  return html
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

// Returns auth headers for all HubSpot API calls
function getHeaders() {
  return {
    Authorization: `Bearer ${process.env.HUBSPOT_TOKEN}`,
    "Content-Type": "application/json",
  };
}

// ── Rate-limit queue ─────────────────────────────────────────
// HubSpot enforces a 10-second rolling limit (free tier: ~100 calls/10s).
// This simple queue serialises calls with a minimum gap so we never burst.
const HS_MIN_GAP_MS = 100; // max ~10 calls/sec, well within limits
let _hsLastCall = 0;
const _hsQueue = [];
let _hsQueueRunning = false;

// Wraps an API call in a promise and adds it to the queue.
// Returns a promise that resolves when the call eventually executes.
function _hsEnqueue(fn) {
  return new Promise((resolve, reject) => {
    _hsQueue.push({ fn, resolve, reject });
    if (!_hsQueueRunning) _hsDrain();
  });
}

// Processes queued API calls one at a time with a minimum gap between calls.
async function _hsDrain() {
  _hsQueueRunning = true;
  while (_hsQueue.length > 0) {
    const { fn, resolve, reject } = _hsQueue.shift();
    const now = Date.now();
    const wait = Math.max(0, HS_MIN_GAP_MS - (now - _hsLastCall));
    if (wait > 0) await new Promise(r => setTimeout(r, wait));
    _hsLastCall = Date.now();
    try { resolve(await fn()); } catch (e) { reject(e); }
  }
  _hsQueueRunning = false;
}

// ── Unified fetch wrapper ─────────────────────────────────────
// All HubSpot API calls go through hsFetch, which routes through the rate-limit
// queue and handles retries on 429 (rate-limited) responses with exponential backoff.
const HS_MAX_RETRIES = 3;

async function hsFetch(url, options = {}) {
  return _hsEnqueue(() => _hsFetchInner(url, options));
}

async function _hsFetchInner(url, options, attempt = 0) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);
  try {
    const res = await fetch(url, {
      headers: getHeaders(),
      signal: controller.signal,
      ...options,
    });
    if (res.status === 429 && attempt < HS_MAX_RETRIES) {
      clearTimeout(timer);
      const retryAfter = parseInt(res.headers.get("retry-after") || "1", 10);
      const backoff = Math.max(retryAfter * 1000, 1000 * (attempt + 1));
      console.warn(`[hubspot] 429 rate-limited, retrying in ${backoff}ms (attempt ${attempt + 1}/${HS_MAX_RETRIES})`);
      await new Promise(r => setTimeout(r, backoff));
      return _hsFetchInner(url, options, attempt + 1);
    }
    if (!res.ok) {
      const body = await res.text();
      let message = `HubSpot ${options.method || "GET"} ${res.status}: ${body.substring(0, 200)}`;
      // Try to extract a cleaner message from HubSpot's JSON error response
      try {
        const parsed = JSON.parse(body);
        if (parsed.message) message = parsed.message;
      } catch (_) {}
      const error = new Error(message);
      error.status = res.status;
      throw error;
    }
    // Some HubSpot endpoints (PUT associations, DELETE) return empty bodies.
    // Parse JSON only when there is a body; return {} otherwise.
    const text = await res.text();
    if (!text) return {};
    try {
      return JSON.parse(text);
    } catch (parseErr) {
      throw new Error(`HubSpot response not valid JSON from ${url}`);
    }
  } finally {
    clearTimeout(timer);
  }
}

function hsGet(url) {
  return hsFetch(url);
}

function hsPost(url, body) {
  return hsFetch(url, { method: "POST", body: JSON.stringify(body) });
}

function hsPatch(url, body) {
  return hsFetch(url, { method: "PATCH", body: JSON.stringify(body) });
}

// ── Contact Lookup ────────────────────────────────────────────

// Looks up a HubSpot contact by email and fetches their associated deals and notes.
// Returns null if no contact is found. This is the primary entry point for CRM context.
async function getContactByEmail(email) {
  const searchRes = await hsPost(`${HUBSPOT_BASE}/crm/v3/objects/contacts/search`, {
    filterGroups: [{ filters: [{ propertyName: "email", operator: "EQ", value: email }] }],
    properties: CONTACT_PROPS,
    limit: 1,
  });

  if (!searchRes.results || searchRes.results.length === 0) return null;

  const raw = searchRes.results[0];
  const props = raw.properties || {};
  const contactId = raw.id;

  const company = props.company || "";
  const contact = {
    hubspotId:      contactId,
    name:           [props.firstname, props.lastname].filter(Boolean).join(" ") || email,
    email:          props.email || email,
    company,
    jobTitle:       props.jobtitle || "",
    lifecycleStage: props.lifecyclestage || "",
    hubspotOwner:   props.hubspot_owner_id || "",
  };

  const [dealsResult, notesResult] = await Promise.allSettled([
    getDeals(contactId, company),
    getNotes(contactId),
  ]);

  if (dealsResult.status === "rejected") console.warn("[hubspot] getDeals failed:", dealsResult.reason?.message);
  if (notesResult.status === "rejected") console.warn("[hubspot] getNotes failed:", notesResult.reason?.message);

  return {
    contact,
    deals: dealsResult.status === "fulfilled" ? dealsResult.value : [],
    notes: notesResult.status === "fulfilled" ? notesResult.value : [],
  };
}

// ── Deals ─────────────────────────────────────────────────────

// Fetches deals associated with a contact. Sorts them so the contact's
// company-relevant deals appear first (by matching company name in deal name).
async function getDeals(contactId, contactCompany) {
  try {
    const assocRes = await hsGet(
      `${HUBSPOT_BASE}/crm/v3/objects/contacts/${contactId}/associations/deals`
    );
    const dealIds = (assocRes.results || []).map((a) => a.id).slice(0, MAX_ASSOCIATIONS);
    if (dealIds.length === 0) return [];

    const dealResults = await Promise.allSettled(
      dealIds.map((id) =>
        hsGet(`${HUBSPOT_BASE}/crm/v3/objects/deals/${id}?properties=${DEAL_PROPS}`)
      )
    );
    const dealData = dealResults.filter(r => r.status === "fulfilled").map(r => r.value);

    // Sort: deals whose name contains the contact's company come first,
    // then by hs_lastmodifieddate descending within each group.
    const companyLower = (contactCompany || "").toLowerCase().trim();
    dealData.sort((a, b) => {
      const nameA = (a.properties?.dealname || "").toLowerCase();
      const nameB = (b.properties?.dealname || "").toLowerCase();
      const matchA = companyLower && nameA.includes(companyLower) ? 1 : 0;
      const matchB = companyLower && nameB.includes(companyLower) ? 1 : 0;
      if (matchA !== matchB) return matchB - matchA; // company-relevant deals first
      const ta = new Date(a.properties?.hs_lastmodifieddate || 0).getTime();
      const tb = new Date(b.properties?.hs_lastmodifieddate || 0).getTime();
      return tb - ta;
    });

    // Resolve custom-pipeline stage IDs to labels via the pipelines cache,
    // so badges show "Qualified to Buy" instead of raw "23155301".
    let stageIdToLabel = {};
    try {
      const pipelines = await getPipelines();
      for (const p of pipelines) {
        for (const s of (p.stages || [])) stageIdToLabel[s.id] = s.label;
      }
    } catch (e) {
      console.warn("[hubspot] getDeals: could not resolve stage labels:", e.message);
    }

    return dealData.map((d) => {
      const p = d.properties || {};
      const rawStage = p.dealstage || "";
      const stageLabel = stageIdToLabel[rawStage] || formatStage(rawStage);
      return {
        id:         d.id,
        name:       p.dealname || "Unnamed Deal",
        stage:      stageLabel,
        stageId:    rawStage,
        pipelineId: p.pipeline  || "",
        amount:     p.amount ? parseFloat(p.amount) : null,
        closeDate:  p.closedate ? String(p.closedate).substring(0, 10) : null,
      };
    });
  } catch (err) {
    console.error("[hubspot] getDeals failed:", err.message);
    return [];
  }
}

/** Convert HubSpot internal stage value to a human-readable label */
function formatStage(stage) {
  if (!stage) return "Unknown";
  const map = {
    appointmentscheduled:  "Appointment Scheduled",
    qualifiedtobuy:        "Qualified to Buy",
    presentationscheduled: "Presentation Scheduled",
    decisionmakerboughtin: "Decision Maker Bought In",
    contractsent:          "Contract Sent",
    closedwon:             "Closed Won",
    closedlost:            "Closed Lost",
  };
  const normalized = stage.toLowerCase().replace(/[\s_-]/g, "");
  return map[normalized] || stage.replace(/[_-]/g, " ").replace(/\b\w/g, (c) => c.toUpperCase());
}

// ── Notes ─────────────────────────────────────────────────────

// Fetches CRM notes associated with a contact, sorted by most recent first.
// Returns plain text snippets (HTML stripped) for display and AI context.
async function getNotes(contactId) {
  try {
    const assocRes = await hsGet(
      `${HUBSPOT_BASE}/crm/v3/objects/contacts/${contactId}/associations/notes`
    );
    const noteIds = (assocRes.results || []).map((a) => a.id).slice(0, MAX_ASSOCIATIONS);
    if (noteIds.length === 0) return [];

    const noteResults = await Promise.allSettled(
      noteIds.map((id) =>
        hsGet(`${HUBSPOT_BASE}/crm/v3/objects/notes/${id}?properties=${NOTE_PROPS}`)
      )
    );
    const noteData = noteResults.filter(r => r.status === "fulfilled").map(r => r.value);

    return noteData
      .filter((n) => n.properties?.hs_note_body)
      .sort((a, b) => {
        const ta = new Date(a.properties.hs_timestamp || 0);
        const tb = new Date(b.properties.hs_timestamp || 0);
        return tb - ta;
      })
      .map((n) => stripHtml(n.properties.hs_note_body).substring(0, 300));
  } catch (err) {
    console.error("[hubspot] getNotes failed:", err.message);
    return [];
  }
}

// ── Create Note ───────────────────────────────────────────────

// Creates a CRM note on a contact documenting that an email was sent.
// Automatically looks up the contact by email first.
async function createNote(contactEmail, subject, bodySummary) {
  const result = await getContactByEmail(contactEmail);
  if (!result) { const e = new Error(`No HubSpot contact found for ${contactEmail}`); e.status = 404; throw e; }

  const contactId = result.contact.hubspotId;
  const MAX_NOTE_LENGTH = 65000; // HubSpot property limit
  const timestamp = new Date().toISOString();
  const prefix    = `Email sent: "${subject}"\n\n`;
  const suffix    = `\n\n[${timestamp}]`;
  const maxBody   = MAX_NOTE_LENGTH - prefix.length - suffix.length;
  const noteBody  = prefix + (bodySummary || "").substring(0, maxBody) + suffix;

  return hsPost(`${HUBSPOT_BASE}/crm/v3/objects/notes`, {
    properties: {
      hs_note_body:  noteBody,
      hs_timestamp:  timestamp,
    },
    associations: [
      {
        to: { id: contactId },
        types: [{ associationCategory: "HUBSPOT_DEFINED", associationTypeId: 202 }],
      },
    ],
  });
}

// ── Pipelines ────────────────────────────────────────────────

// Fetches all deal pipelines and their stages from HubSpot.
// Cached for 5 minutes to avoid repeated API calls - stages are needed
// frequently to resolve numeric stage IDs to human-readable labels.
async function getPipelines() {
  const now = Date.now();
  if (_pipelinesCache && now < _pipelinesCacheExpiresAt) {
    return _pipelinesCache;
  }
  const data = await hsGet(`${HUBSPOT_BASE}/crm/v3/pipelines/deals`);
  const pipelines = (data.results || []).map((p) => ({
    id:    p.id,
    label: p.label,
    stages: (p.stages || [])
      .sort((a, b) => (a.displayOrder || 0) - (b.displayOrder || 0))
      .map((s) => ({
        id:    s.id,
        label: s.label,
      })),
  }));
  _pipelinesCache = pipelines;
  _pipelinesCacheExpiresAt = now + CACHE_TTL;
  return pipelines;
}

// ── Deal Properties (dynamic from HubSpot) ──────────────────
// Allowlist of deal properties shown in the Create Deal form.
// Only these HubSpot internal property names will appear as form fields.
// This prevents exposing system/internal properties to end users.
const ALLOWED_DEAL_PROPS = new Set([
  "industry",          // Industry
  "company_size",      // Size (custom)
  "deal_size",         // Size (alt custom name)
  "source_channel",    // Source Channel (custom)
  "dealtype",          // Deal Type
  "source_detail",     // Source Detail (custom)
  "hubspot_owner_id",  // Deal Owner
  "createdate",        // Create Date
]);

// Fetches deal property definitions and owner options from HubSpot.
// Cached for 5 minutes. Used to dynamically build the Create Deal form.
async function getDealProperties() {
  const now = Date.now();
  if (_dealPropsCache && now < _dealPropsCacheExpiresAt) {
    return _dealPropsCache;
  }
  // Fetch deal properties and owners in parallel — owners are needed
  // because hubspot_owner_id doesn't include options in the properties API.
  const [data, ownersData] = await Promise.all([
    hsGet(`${HUBSPOT_BASE}/crm/v3/properties/deals`),
    hsGet(`${HUBSPOT_BASE}/crm/v3/owners`).catch(() => ({ results: [] })),
  ]);
  const ownerOptions = (ownersData.results || [])
    .filter(o => !o.archived)
    .map(o => ({
      value: String(o.id),
      label: [o.firstName, o.lastName].filter(Boolean).join(" ") || o.email,
    }))
    .sort((a, b) => a.label.localeCompare(b.label));

  const props = (data.results || [])
    .filter((p) => {
      // Only show explicitly allowed properties
      return ALLOWED_DEAL_PROPS.has(p.name);
    })
    .map((p) => {
      // Inject owner options for hubspot_owner_id
      const options = p.name === "hubspot_owner_id" ? ownerOptions : (p.options || [])
        .filter(o => !o.hidden)
        .sort((a, b) => (a.displayOrder || 0) - (b.displayOrder || 0))
        .map(o => ({ value: o.value, label: o.label }));
      return {
        name:        p.name,
        label:       p.label,
        type:        p.name === "hubspot_owner_id" ? "enumeration" : p.type,
        fieldType:   p.fieldType,
        groupName:   p.groupName,
        description: p.description || "",
        required:    p.formField || false,
        options,
      };
    })
    .sort((a, b) => {
      // Put common fields first: amount, closedate, then rest alphabetically
      const priority = { amount: 0, closedate: 1, deal_currency_code: 2 };
      const pa = priority[a.name] ?? 99;
      const pb = priority[b.name] ?? 99;
      if (pa !== pb) return pa - pb;
      return a.label.localeCompare(b.label);
    });

  _dealPropsCache = props;
  _dealPropsCacheExpiresAt = now + CACHE_TTL;
  return props;
}

// ── Recent Deal Names (for AI naming patterns) ───────────────

// Fetches the 15 most recent deal names so Claude can learn the naming
// convention used in this CRM and suggest consistent new deal names.
async function getRecentDealNames() {
  try {
    const data = await hsPost(`${HUBSPOT_BASE}/crm/v3/objects/deals/search`, {
      sorts: [{ propertyName: "createdate", direction: "DESCENDING" }],
      properties: ["dealname"],
      limit: 15,
    });
    return (data.results || [])
      .map((d) => d.properties?.dealname)
      .filter(Boolean);
  } catch (err) {
    console.warn("[hubspot] getRecentDealNames failed:", err.message);
    return [];
  }
}

// ── Create Deal ──────────────────────────────────────────────

// Creates a new deal in HubSpot, associated with a contact.
// Auto-creates the contact if they don't already exist.
// Optionally attaches a note to the deal.
async function createDeal(contactEmail, dealProps) {
  // Look up contact — auto-create if not found so the user doesn't
  // have to create the contact as a separate step before creating a deal.
  let result = await getContactByEmail(contactEmail);
  let contactId;
  if (result) {
    contactId = result.contact.hubspotId;
  } else {
    console.log(`[hubspot] createDeal: contact ${contactEmail} not in HubSpot — auto-creating`);
    const created = await createContact({ email: contactEmail });
    contactId = created.contactId;
  }

  // Separate notes (handled as a separate object) from deal properties
  const { notes, ...rest } = dealProps;

  // Build properties — only allow known deal property names
  const properties = {};
  for (const [key, val] of Object.entries(rest)) {
    if (val === undefined || val === null || val === "") continue;
    // Enforce allowlist: dealname, dealstage, pipeline, amount, closedate + ALLOWED_DEAL_PROPS
    if (key !== "dealname" && key !== "dealstage" && key !== "pipeline" &&
        key !== "amount" && key !== "closedate" && !ALLOWED_DEAL_PROPS.has(key)) continue;
    properties[key] = String(val);
  }
  if (!properties.dealname) { const e = new Error("dealname is required"); e.status = 400; throw e; }

  const dealRes = await hsPost(`${HUBSPOT_BASE}/crm/v3/objects/deals`, {
    properties,
    associations: [
      {
        to: { id: contactId },
        types: [{ associationCategory: "HUBSPOT_DEFINED", associationTypeId: 3 }],
      },
    ],
  });

  // Optionally attach a note to the deal
  if (notes && notes.trim()) {
    try {
      await hsPost(`${HUBSPOT_BASE}/crm/v3/objects/notes`, {
        properties: {
          hs_note_body:  notes.substring(0, 65000),
          hs_timestamp:  new Date().toISOString(),
        },
        associations: [
          {
            to: { id: dealRes.id },
            types: [{ associationCategory: "HUBSPOT_DEFINED", associationTypeId: 214 }],
          },
          {
            to: { id: contactId },
            types: [{ associationCategory: "HUBSPOT_DEFINED", associationTypeId: 202 }],
          },
        ],
      });
    } catch (err) {
      console.warn("[hubspot] createDeal note failed (deal still created):", err.message);
    }
  }

  return { dealId: dealRes.id, dealname: properties.dealname };
}

// ── Search Deals ─────────────────────────────────────────────
// HubSpot's search endpoint caps each request at 100 results. To return all
// accessible deals (up to a safety cap) we follow the `paging.next.after`
// cursor. SEARCH_MAX_DEALS bounds the total to avoid unbounded loops on
// accounts with thousands of deals.
const SEARCH_MAX_DEALS = Math.max(parseInt(process.env.HUBSPOT_DEAL_SEARCH_MAX || "500", 10) || 500, 100);
const SEARCH_PAGE_SIZE = 100; // HubSpot CRM v3 per-request max

// Searches deals by name using HubSpot's CONTAINS_TOKEN operator (whole-word match).
// Used for the deal search autocomplete in the UI.
async function searchDeals(query) {
  try {
    // CONTAINS_TOKEN matches whole words/tokens in the deal name.
    // "SLB" matches "SLB North Africa" but not "Assembly".
    const sanitized = (query || "").replace(/[^a-zA-Z0-9\s\-_.]/g, "").trim();

    // Return only the first page of results — UI autocomplete, not bulk export.
    const data = await hsPost(`${HUBSPOT_BASE}/crm/v3/objects/deals/search`, {
      filterGroups: sanitized ? [{
        filters: [{ propertyName: "dealname", operator: "CONTAINS_TOKEN", value: sanitized }],
      }] : [],
      sorts: [{ propertyName: "createdate", direction: "DESCENDING" }],
      properties: ["dealname", "dealstage", "amount", "closedate"],
      limit: 20,
    });
    const all = data.results || [];
    // Resolve custom-pipeline stage IDs (e.g. "23155301") to their human label
    // via the pipelines cache. Hard-coded formatStage only knows the default
    // pipeline; custom ones come back as raw numeric ids without this step.
    let stageIdToLabel = {};
    try {
      const pipelines = await getPipelines();
      for (const p of pipelines) {
        for (const s of (p.stages || [])) stageIdToLabel[s.id] = s.label;
      }
    } catch (e) {
      console.warn("[hubspot] searchDeals: could not resolve stage labels:", e.message);
    }

    return all.slice(0, SEARCH_MAX_DEALS).map((d) => {
      const p = d.properties || {};
      const rawStage = p.dealstage || "";
      const stageLabel = stageIdToLabel[rawStage] || formatStage(rawStage);
      return {
        id:        d.id,
        name:      p.dealname || "Unnamed Deal",
        stage:     stageLabel,
        amount:    p.amount ? parseFloat(p.amount) : null,
        closeDate: p.closedate ? String(p.closedate).substring(0, 10) : null,
      };
    });
  } catch (err) {
    console.warn("[hubspot] searchDeals failed:", err.message);
    return [];
  }
}

// ── Move Deal Stage ──────────────────────────────────────────
/**
 * Move a deal to a different stage. Optionally verifies the stage belongs to
 * the deal's pipeline by checking against getPipelines().
 * Returns { dealId, stageId, stageLabel, pipelineId }.
 */
async function moveDealStage(dealId, stageId) {
  if (!dealId || !/^\d+$/.test(String(dealId))) {
    const e = new Error("Invalid dealId"); e.status = 400; throw e;
  }
  if (!stageId || typeof stageId !== "string") {
    const e = new Error("Invalid stageId"); e.status = 400; throw e;
  }

  // Fetch the deal first to know its current pipeline
  const deal = await hsGet(
    `${HUBSPOT_BASE}/crm/v3/objects/deals/${encodeURIComponent(dealId)}?properties=dealname,dealstage,pipeline`
  );
  const currentPipelineId = deal?.properties?.pipeline;
  if (!currentPipelineId) {
    const e = new Error("Deal has no pipeline set"); e.status = 404; throw e;
  }

  // Validate that stageId belongs to that pipeline
  const pipelines = await getPipelines();
  const pipeline  = pipelines.find((p) => p.id === currentPipelineId);
  if (!pipeline) {
    const e = new Error(`Unknown pipeline ${currentPipelineId}`); e.status = 404; throw e;
  }
  const stage = pipeline.stages.find((s) => s.id === stageId);
  if (!stage) {
    const e = new Error(`Stage ${stageId} does not belong to pipeline ${pipeline.label}`); e.status = 400; throw e;
  }

  await hsPatch(`${HUBSPOT_BASE}/crm/v3/objects/deals/${encodeURIComponent(dealId)}`, {
    properties: { dealstage: stageId },
  });

  // Use the raw pipeline label from HubSpot — getDeals and searchDeals now
  // also return the raw pipeline label (resolved via the same pipelines
  // cache), so this keeps the wire format consistent across all three
  // endpoints. Wrapping with formatStage() here would title-case custom
  // labels and drift from the other endpoints.
  return {
    dealId:     String(dealId),
    stageId,
    stageLabel: stage.label,
    pipelineId: currentPipelineId,
  };
}

// ── Link Contact to Existing Deal ────────────────────────────
async function linkContactToDeal(contactEmail, dealId) {
  let result = await getContactByEmail(contactEmail);
  let contactId;
  if (result) {
    contactId = result.contact.hubspotId;
  } else {
    console.log(`[hubspot] linkContactToDeal: contact ${contactEmail} not in HubSpot — auto-creating`);
    const created = await createContact({ email: contactEmail });
    contactId = created.contactId;
  }

  // Associate contact → deal
  await hsFetch(
    `${HUBSPOT_BASE}/crm/v3/objects/contacts/${contactId}/associations/deals/${dealId}/${ASSOC_CONTACT_TO_DEAL}`,
    { method: "PUT", body: JSON.stringify({}) }
  );

  return { contactId, dealId };
}

// ── Link Email to Deal ───────────────────────────────────────
// Creates an email engagement object in HubSpot and associates it with a deal
// (and optionally a contact). This makes the email appear in the deal's
// activity timeline. Three-step process: create email -> link to deal -> link to contact.
async function linkEmailToDeal({ subject, bodyHtml, bodyText, fromEmail, toEmail, ccEmail, sentAt }, dealId, contactEmail, { contactId: preResolvedContactId, userEmail } = {}) {
  // NOTE: hs_email_headers is intentionally omitted. When set, HubSpot
  // auto-associates the email to every deal that any from/to/cc contact is on,
  // causing cross-contamination (e.g. internal team members on many deals cause
  // unrelated emails to appear on all of them). Without headers, HubSpot only
  // associates the email where we explicitly tell it to (steps 2 and 3 below).
  // Trade-off: from/to shows as "Unknown" in HubSpot timeline, but subject,
  // body, direction, timestamp, and deal/contact associations are all preserved.

  // Use HTML body for formatting; fall back to plain text wrapped in <pre>
  const htmlContent = bodyHtml
    ? bodyHtml.substring(0, 65536)
    : `<pre style="white-space:pre-wrap;font-family:inherit">${(bodyText || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}</pre>`;

  // 1. Create the email engagement object in HubSpot
  const emailObj = await hsPost(`${HUBSPOT_BASE}/crm/v3/objects/emails`, {
    properties: {
      hs_timestamp: sentAt || new Date().toISOString(),
      // Determine direction: "EMAIL" = outbound (user sent it), "INCOMING_EMAIL" = inbound
      hs_email_direction: (userEmail && fromEmail && fromEmail.toLowerCase() === userEmail.toLowerCase()) ? "EMAIL" : "INCOMING_EMAIL",
      hs_email_subject: subject || "(no subject)",
      hs_email_html: htmlContent,
      hs_email_text: (bodyText || "").substring(0, 65536),
      hs_email_status: "SENT",
    },
  });
  const emailId = emailObj.id;

  // 2. Associate email → deal
  await hsFetch(
    `${HUBSPOT_BASE}/crm/v3/objects/emails/${emailId}/associations/deals/${dealId}/${ASSOC_EMAIL_TO_DEAL}`,
    { method: "PUT", body: JSON.stringify({}) }
  );

  // 3. Associate email → contact so HubSpot shows the contact name (not "Unknown Contact")
  if (contactEmail || preResolvedContactId) {
    try {
      let resolvedContactId = preResolvedContactId;
      if (!resolvedContactId && contactEmail) {
        const result = await getContactByEmail(contactEmail);
        resolvedContactId = result?.contact?.hubspotId;
      }
      if (resolvedContactId) {
        await hsFetch(
          `${HUBSPOT_BASE}/crm/v3/objects/emails/${emailId}/associations/contacts/${resolvedContactId}/${ASSOC_EMAIL_TO_CONTACT}`,
          { method: "PUT", body: JSON.stringify({}) }
        );
      }
    } catch (e) {
      console.warn("[hubspot] linkEmailToDeal: contact association failed:", e.message);
    }
  }

  return { emailId, dealId };
}

// ── Create Contact ───────────────────────────────────────────

// Creates a new contact in HubSpot. Optionally associates with a company.
async function createContact({ email, firstname, lastname, company, jobtitle, companyId }) {
  const properties = { email };
  if (firstname) properties.firstname = firstname;
  if (lastname)  properties.lastname  = lastname;
  if (company)   properties.company   = company;
  if (jobtitle)  properties.jobtitle  = jobtitle;

  const associations = [];
  if (companyId) {
    associations.push({
      to: { id: companyId },
      types: [{ associationCategory: "HUBSPOT_DEFINED", associationTypeId: 1 }], // contact-to-company
    });
  }

  const contactRes = await hsPost(`${HUBSPOT_BASE}/crm/v3/objects/contacts`, {
    properties,
    ...(associations.length > 0 ? { associations } : {}),
  });

  return { contactId: contactRes.id, email };
}

// ── Search Companies ─────────────────────────────────────────
async function searchCompanies(query) {
  try {
    const data = await hsPost(`${HUBSPOT_BASE}/crm/v3/objects/companies/search`, {
      filterGroups: query ? [{
        filters: [{ propertyName: "name", operator: "CONTAINS_TOKEN", value: query }],
      }] : [],
      sorts: [{ propertyName: "createdate", direction: "DESCENDING" }],
      properties: ["name", "domain"],
      limit: 20,
    });
    return (data.results || []).map((c) => ({
      id:     c.id,
      name:   c.properties?.name || "Unnamed",
      domain: c.properties?.domain || "",
    }));
  } catch (err) {
    console.warn("[hubspot] searchCompanies failed:", err.message);
    return [];
  }
}

// ── Create Company ───────────────────────────────────────────
async function createCompany({ name, domain }) {
  const properties = {};
  if (name)   properties.name   = name;
  if (domain) properties.domain = domain;

  const companyRes = await hsPost(`${HUBSPOT_BASE}/crm/v3/objects/companies`, { properties });
  return { companyId: companyRes.id, name: name || domain };
}

// ── Portal ID ────────────────────────────────────────────────

// Returns the HubSpot portal ID, needed for building deep links to HubSpot UI.
// Prefers the env var to avoid an API call; falls back to the account-info API.
async function getPortalId() {
  // Prefer env var if set (avoids an API call)
  if (process.env.HUBSPOT_PORTAL_ID) {
    return process.env.HUBSPOT_PORTAL_ID;
  }
  if (!process.env.HUBSPOT_TOKEN) return null;

  const now = Date.now();
  if (_portalIdCache && now < _portalIdCacheExpiresAt) {
    return _portalIdCache;
  }
  try {
    const data = await hsGet(`${HUBSPOT_BASE}/account-info/v3/details`);
    const portalId = data.portalId ? String(data.portalId) : null;
    if (portalId) {
      _portalIdCache = portalId;
      _portalIdCacheExpiresAt = now + CACHE_TTL;
    }
    return portalId;
  } catch (err) {
    console.warn("[hubspot] getPortalId failed:", err.message);
    return null;
  }
}

// ── Deal Contacts ─────────────────────────────────────────────

// Fetches all contacts associated with a deal. Used by autoTracker to verify
// that an email's participants overlap with the deal's contacts before linking.
async function getDealContacts(dealId) {
  try {
    const assocRes = await hsGet(
      `${HUBSPOT_BASE}/crm/v3/objects/deals/${dealId}/associations/contacts`
    );
    const contactIds = (assocRes.results || []).map((a) => a.id).slice(0, MAX_ASSOCIATIONS);
    if (contactIds.length === 0) return [];

    const contactResults = await Promise.allSettled(
      contactIds.map((id) =>
        hsGet(`${HUBSPOT_BASE}/crm/v3/objects/contacts/${id}?properties=email,firstname,lastname,company`)
      )
    );
    return contactResults
      .filter(r => r.status === "fulfilled")
      .map(r => {
        const p = r.value.properties || {};
        return {
          id:      r.value.id,
          email:   (p.email || "").toLowerCase(),
          name:    [p.firstname, p.lastname].filter(Boolean).join(" "),
          company: p.company || "",
        };
      })
      .filter(c => c.email);
  } catch (err) {
    console.error("[hubspot] getDealContacts failed:", err.message);
    return [];
  }
}

// ── Deal Emails (for idempotency check) ──────────────────────

// Fetches email engagements already linked to a deal.
// Used to check if an email was already logged, preventing duplicates.
async function getDealEmails(dealId) {
  try {
    const assocRes = await hsGet(
      `${HUBSPOT_BASE}/crm/v3/objects/deals/${dealId}/associations/emails`
    );
    const emailIds = (assocRes.results || []).map((a) => a.id).slice(0, 50);
    if (emailIds.length === 0) return [];

    const emailResults = await Promise.allSettled(
      emailIds.map((id) =>
        hsGet(`${HUBSPOT_BASE}/crm/v3/objects/emails/${id}?properties=hs_email_subject,hs_timestamp`)
      )
    );
    return emailResults
      .filter(r => r.status === "fulfilled" && r.value.properties)
      .map(r => ({
        id:        r.value.id,
        subject:   r.value.properties.hs_email_subject || "",
        timestamp: r.value.properties.hs_timestamp || "",
      }));
  } catch (err) {
    console.error("[hubspot] getDealEmails failed:", err.message);
    return [];
  }
}

// ── Follow-up Deals ─────────────────────────────────────────
// Fetches all open deals and calculates days since last real activity.
// Uses hs_lastactivitydate (logged calls/emails/meetings) first,
// then notes_last_updated, then createdate as fallback.
// Explicitly avoids hs_lastmodifieddate because it fires on automation
// changes (not real human activity), which would hide stale deals.
const FOLLOWUP_DEAL_PROPS = "dealname,dealstage,amount,closedate,pipeline,hs_lastactivitydate,notes_last_updated,createdate,hubspot_owner_id";
const FOLLOWUP_PAGE_SIZE = 100;
const FOLLOWUP_MAX_DEALS = 500;

// Cache for follow-up deals (2-minute TTL — shorter than general CACHE_TTL
// because follow-up data is more time-sensitive)
const FOLLOWUP_CACHE_TTL = 2 * 60 * 1000;
let _followupDealsCache = null;
let _followupDealsCacheExpiresAt = 0;

// Returns all open deals sorted by staleness (most stale first).
// Paginates through all results up to FOLLOWUP_MAX_DEALS. Cached for 2 minutes.
async function getFollowupDeals() {
  // Return cached result if still valid
  const now0 = Date.now();
  if (_followupDealsCache && now0 < _followupDealsCacheExpiresAt) {
    return _followupDealsCache;
  }

  try {
    let allDeals = [];
    let after = undefined;

    // Paginate through all deals (up to cap)
    while (allDeals.length < FOLLOWUP_MAX_DEALS) {
      const body = {
        filterGroups: [{ filters: [{ propertyName: "hs_is_closed", operator: "EQ", value: "false" }] }],
        sorts: [{ propertyName: "hs_lastactivitydate", direction: "ASCENDING" }],
        properties: FOLLOWUP_DEAL_PROPS.split(","),
        limit: FOLLOWUP_PAGE_SIZE,
      };
      if (after) body.after = after;

      const data = await hsPost(`${HUBSPOT_BASE}/crm/v3/objects/deals/search`, body);
      const results = data.results || [];
      allDeals = allDeals.concat(results);

      if (!data.paging?.next?.after || results.length < FOLLOWUP_PAGE_SIZE) break;
      after = data.paging.next.after;
    }

    // Resolve stage labels
    let stageIdToLabel = {};
    try {
      const pipelines = await getPipelines();
      for (const p of pipelines) {
        for (const s of (p.stages || [])) stageIdToLabel[s.id] = s.label;
      }
    } catch (e) {
      console.warn("[hubspot] getFollowupDeals: could not resolve stage labels:", e.message);
    }

    const now = Date.now();

    const result = allDeals.map((d) => {
      const p = d.properties || {};

      // Last activity fallback chain (same as Maxine)
      const lastActivity = p.hs_lastactivitydate || p.notes_last_updated || p.createdate || null;
      const lastActivityMs = lastActivity ? new Date(lastActivity).getTime() : 0;
      const daysSince = lastActivityMs ? Math.floor((now - lastActivityMs) / (1000 * 60 * 60 * 24)) : 999;

      const rawStage = p.dealstage || "";
      return {
        id:            d.id,
        name:          p.dealname || "Unnamed Deal",
        stage:         stageIdToLabel[rawStage] || formatStage(rawStage),
        stageId:       rawStage,
        pipelineId:    p.pipeline || "",
        amount:        p.amount ? parseFloat(p.amount) : null,
        closeDate:     p.closedate ? String(p.closedate).substring(0, 10) : null,
        lastActivity:  lastActivity,
        daysSince,
        ownerId:       p.hubspot_owner_id || null,
      };
    }).sort((a, b) => b.daysSince - a.daysSince); // stalest first

    // Store in cache
    _followupDealsCache = result;
    _followupDealsCacheExpiresAt = Date.now() + FOLLOWUP_CACHE_TTL;

    return result;
  } catch (err) {
    console.error("[hubspot] getFollowupDeals failed:", err.message);
    return [];
  }
}

// ── Owner Lookup by Email (cached) ──────────────────────────
// Finds a HubSpot owner ID by their email address. Caches the full
// owner list for 5 minutes. Used to match the logged-in user to their
// HubSpot owner record for filtering deals by ownership.
let _ownersCache = null;
let _ownersCacheExpiresAt = 0;

async function getOwnerIdByEmail(email) {
  const now = Date.now();
  if (!_ownersCache || now >= _ownersCacheExpiresAt) {
    try {
      const data = await hsGet(`${HUBSPOT_BASE}/crm/v3/owners`);
      _ownersCache = (data.results || []).filter(o => !o.archived);
      _ownersCacheExpiresAt = now + CACHE_TTL;
    } catch (err) {
      console.error("[hubspot] getOwnerIdByEmail: failed to fetch owners:", err.message);
      return null;
    }
  }
  const normalized = (email || "").toLowerCase();
  const match = _ownersCache.find(o => (o.email || "").toLowerCase() === normalized);
  return match ? String(match.id) : null;
}

module.exports = { getContactByEmail, createNote, getPipelines, getRecentDealNames, createDeal, searchDeals, linkContactToDeal, linkEmailToDeal, createContact, searchCompanies, createCompany, moveDealStage, getDealProperties, getPortalId, getDealContacts, getDealEmails, getFollowupDeals, getOwnerIdByEmail };
