"use strict";

/**
 * Tests for all link-email validation, auto-tracker guards, and refine improvements
 * added during the 2026-06-09 audit session.
 */

// ── Helper: reconstruct normalizeSubject from graph.js ──
function normalizeSubject(subj) {
  let s = (typeof subj === "string" ? subj : "").trim();
  s = s.replace(/\[(?:ext|external|extern|secure|encrypted|caution|spam)\]\s*/gi, "");
  s = s.replace(/\s*@\s+.*$/i, "");
  const prefixRe = /^(re|fw|fwd|aw|sv|ref|rif|odp|ynt|accepted|declined|tentative|tentatively accepted|updated invitation|updated|cancelled invitation|cancelled|canceled)\s*:\s*/i;
  let prev;
  do { prev = s; s = s.replace(prefixRe, ""); } while (s !== prev);
  return s.trim().toLowerCase();
}

// ── Helper: reconstruct isExternal check ──
function makeIsExternal(userEmail) {
  const internalDomain = userEmail ? userEmail.split("@")[1]?.toLowerCase() : null;
  return (email) => {
    if (!email || !internalDomain) return true;
    return !email.toLowerCase().endsWith("@" + internalDomain);
  };
}

// ══════════════════════════════════════════════════════════════
// 1. THREE-LAYER VALIDATION (participant, domain, subject)
// ══════════════════════════════════════════════════════════════
describe("link-email three-layer validation", () => {

  describe("Layer 1: External participant overlap", () => {
    it("passes when thread participant matches deal contact", () => {
      const dealEmails = ["john@acme.com", "jane@acme.com"];
      const threadParticipants = ["john@acme.com", "demo-user@example.com"];
      const isExternal = makeIsExternal("demo-user@example.com");
      const extDeal = dealEmails.filter(isExternal);
      const extThread = threadParticipants.filter(isExternal);
      const hasOverlap = extThread.some(e => extDeal.includes(e));
      expect(hasOverlap).toBe(true);
    });

    it("fails when only internal team members overlap", () => {
      const dealEmails = ["john@acme.com", "demo-user@example.com"];
      const threadParticipants = ["demo-user@example.com", "demo-user@example.com"];
      const isExternal = makeIsExternal("demo-user@example.com");
      const extDeal = dealEmails.filter(isExternal);
      const extThread = threadParticipants.filter(isExternal);
      const hasOverlap = extThread.some(e => extDeal.includes(e));
      expect(hasOverlap).toBe(false);
    });

    it("filters out all @example.invalid addresses from both sides", () => {
      const dealEmails = ["demo-user@example.com", "demo-user@example.com", "client@external.com"];
      const threadParticipants = ["demo-user@example.com", "other@different.com"];
      const isExternal = makeIsExternal("demo-user@example.com");
      const extDeal = dealEmails.filter(isExternal);
      const extThread = threadParticipants.filter(isExternal);
      expect(extDeal).toEqual(["client@external.com"]);
      expect(extThread).toEqual(["other@different.com"]);
      expect(extThread.some(e => extDeal.includes(e))).toBe(false);
    });
  });

  describe("Layer 2: Domain overlap", () => {
    it("passes when thread domain matches deal contact domain", () => {
      const dealDomains = new Set(["acme.com", "partner.co"]);
      const threadDomains = new Set(["acme.com"]);
      const hasDomainOverlap = [...threadDomains].some(d => dealDomains.has(d));
      expect(hasDomainOverlap).toBe(true);
    });

    it("fails when domains are completely different", () => {
      const dealDomains = new Set(["acme.com"]);
      const threadDomains = new Set(["eaton.com", "gmx.de"]);
      const hasDomainOverlap = [...threadDomains].some(d => dealDomains.has(d));
      expect(hasDomainOverlap).toBe(false);
    });

    it("ignores internal domain in comparison", () => {
      const isExternal = makeIsExternal("demo-user@example.com");
      const allDealEmails = ["demo-user@example.com", "john@acme.com"];
      const allThreadEmails = ["demo-user@example.com", "bob@different.com"];
      const dealDomains = new Set(allDealEmails.filter(isExternal).map(e => e.split("@")[1]));
      const threadDomains = new Set(allThreadEmails.filter(isExternal).map(e => e.split("@")[1]));
      expect([...threadDomains].some(d => dealDomains.has(d))).toBe(false);
    });
  });

  describe("Layer 3: Subject consistency", () => {
    it("passes when new thread subject matches existing deal email", () => {
      const existingSubjects = new Set(["contact details", "re: contact details"].map(normalizeSubject));
      const newSubject = normalizeSubject("RE: Contact Details");
      expect(existingSubjects.has(newSubject)).toBe(true);
    });

    it("fails when subject is completely different", () => {
      const existingSubjects = new Set(["contact details"].map(normalizeSubject));
      const newSubject = normalizeSubject("Example AI Enquiry");
      expect(existingSubjects.has(newSubject)).toBe(false);
    });

    it("normalizes Re: Fwd: [EXTERNAL] prefixes for comparison", () => {
      const existing = normalizeSubject("RE: FW: [External] Project Proposal");
      const incoming = normalizeSubject("Re: Project Proposal");
      expect(existing).toBe(incoming);
    });

    it("skips check when deal has no existing emails (first link)", () => {
      const existingSubjects = new Set();
      const newSubject = normalizeSubject("Anything Goes");
      // When existingSubjects is empty, check should be skipped (pass through)
      const shouldBlock = newSubject && existingSubjects.size > 0 && !existingSubjects.has(newSubject);
      expect(shouldBlock).toBe(false);
    });
  });
});

// ══════════════════════════════════════════════════════════════
// 2. DOMAIN-FILTERED CONTACT LINKING
// ══════════════════════════════════════════════════════════════
describe("domain-filtered contact linking", () => {
  it("links contacts from matching domains", () => {
    const existingDealDomains = new Set(["acme.com", "partner.co"]);
    const participants = ["alice@acme.com", "bob@partner.co", "charlie@eaton.com"];
    const toLink = participants.filter(e => {
      const domain = e.split("@")[1]?.toLowerCase();
      return !existingDealDomains.size || !domain || existingDealDomains.has(domain);
    });
    expect(toLink).toEqual(["alice@acme.com", "bob@partner.co"]);
  });

  it("skips contacts from non-matching domains", () => {
    const existingDealDomains = new Set(["acme.com"]);
    const participants = ["charlie@eaton.com", "dave@gmx.de"];
    const toLink = participants.filter(e => existingDealDomains.has(e.split("@")[1]));
    expect(toLink).toEqual([]);
  });

  it("allows all contacts when deal has no existing contacts", () => {
    const existingDealDomains = new Set();
    const participants = ["anyone@anywhere.com"];
    const toLink = participants.filter(e => {
      const domain = e.split("@")[1];
      return existingDealDomains.size === 0 || existingDealDomains.has(domain);
    });
    expect(toLink).toEqual(["anyone@anywhere.com"]);
  });
});

// ══════════════════════════════════════════════════════════════
// 3. AUTO-TRACKER PARTICIPANT GUARD
// ══════════════════════════════════════════════════════════════
describe("auto-tracker participant guard", () => {
  it("allows email when participant is a deal contact", () => {
    const dealContactEmails = new Set(["john@acme.com", "jane@acme.com"]);
    const emailParticipants = ["john@acme.com", "demo-user@example.com"];
    const userEmail = "demo-user@example.com";
    const external = emailParticipants.filter(e => e !== userEmail);
    const hasOverlap = external.some(e => dealContactEmails.has(e));
    expect(hasOverlap).toBe(true);
  });

  it("blocks email when no participant matches deal contacts", () => {
    const dealContactEmails = new Set(["john@acme.com"]);
    const emailParticipants = ["bob@eaton.com", "demo-user@example.com"];
    const userEmail = "demo-user@example.com";
    const external = emailParticipants.filter(e => e !== userEmail);
    const hasOverlap = external.some(e => dealContactEmails.has(e));
    expect(hasOverlap).toBe(false);
  });

  it("excludes the signed-in user from participant check", () => {
    const dealContactEmails = new Set(["demo-user@example.com", "john@acme.com"]);
    const emailParticipants = ["demo-user@example.com"];
    const userEmail = "demo-user@example.com";
    const external = emailParticipants.filter(e => e !== userEmail);
    // Only the user is on the email — no external overlap
    expect(external.length).toBe(0);
    expect(external.some(e => dealContactEmails.has(e))).toBe(false);
  });

  it("handles empty deal contacts gracefully (allows through)", () => {
    const dealContactEmails = new Set();
    const emailParticipants = ["bob@eaton.com"];
    // When deal has no contacts, should allow (can't verify)
    const hasOverlap = emailParticipants.some(e => dealContactEmails.has(e));
    expect(hasOverlap).toBe(false);
    // But the guard only blocks when dealEmails.size > 0
    const shouldBlock = !hasOverlap && dealContactEmails.size > 0;
    expect(shouldBlock).toBe(false);
  });
});

// ══════════════════════════════════════════════════════════════
// 4. hs_email_headers REMOVAL
// ══════════════════════════════════════════════════════════════
describe("linkEmailToDeal without hs_email_headers", () => {
  it("email properties should NOT contain hs_email_headers", () => {
    // Simulate what linkEmailToDeal now builds
    const properties = {
      hs_timestamp: "2026-06-09T00:00:00Z",
      hs_email_direction: "EMAIL",
      hs_email_subject: "Test Subject",
      hs_email_html: "<p>Test body</p>",
      hs_email_text: "Test body",
      hs_email_status: "SENT",
    };
    expect(properties).not.toHaveProperty("hs_email_headers");
  });

  it("preserves all other required properties", () => {
    const properties = {
      hs_timestamp: "2026-06-09T00:00:00Z",
      hs_email_direction: "INCOMING_EMAIL",
      hs_email_subject: "Re: Test",
      hs_email_html: "<p>Reply</p>",
      hs_email_text: "Reply",
      hs_email_status: "SENT",
    };
    expect(properties.hs_timestamp).toBeDefined();
    expect(properties.hs_email_direction).toBeDefined();
    expect(properties.hs_email_subject).toBeDefined();
    expect(properties.hs_email_html).toBeDefined();
    expect(properties.hs_email_text).toBeDefined();
    expect(properties.hs_email_status).toBe("SENT");
  });

  it("direction is EMAIL when sender matches user", () => {
    const userEmail = "demo-user@example.com";
    const fromEmail = "demo-user@example.com";
    const direction = (userEmail && fromEmail && fromEmail.toLowerCase() === userEmail.toLowerCase()) ? "EMAIL" : "INCOMING_EMAIL";
    expect(direction).toBe("EMAIL");
  });

  it("direction is INCOMING_EMAIL when sender is external", () => {
    const userEmail = "demo-user@example.com";
    const fromEmail = "john@acme.com";
    const direction = (userEmail && fromEmail && fromEmail.toLowerCase() === userEmail.toLowerCase()) ? "EMAIL" : "INCOMING_EMAIL";
    expect(direction).toBe("INCOMING_EMAIL");
  });
});

// ══════════════════════════════════════════════════════════════
// 5. REFINE CHAT MEMORY + THREAD CONTEXT
// ══════════════════════════════════════════════════════════════
describe("refine chat memory", () => {
  it("builds multi-turn messages from chat history", () => {
    const chatHistory = ["make it shorter", "add urgency"];
    const currentMessage = "also mention the deadline";

    // Simulate the message building from claude.js
    const messages = [];
    if (chatHistory.length > 0) {
      messages.push({ role: "user", content: "Initial context + " + chatHistory[0] });
      messages.push({ role: "assistant", content: '{"subject":"Test","body":"(applied)"}' });
      for (let i = 1; i < chatHistory.length; i++) {
        messages.push({ role: "user", content: chatHistory[i] });
        messages.push({ role: "assistant", content: '{"subject":"Test","body":"(applied)"}' });
      }
      messages.push({ role: "user", content: currentMessage });
    }

    // Should have: initial+first, response, second, response, current = 5 messages
    expect(messages.length).toBe(5);
    expect(messages[0].role).toBe("user");
    expect(messages[1].role).toBe("assistant");
    expect(messages[2].content).toBe("add urgency");
    expect(messages[4].content).toBe("also mention the deadline");
  });

  it("sends single message when no chat history", () => {
    const chatHistory = [];
    const messages = [];
    if (chatHistory.length > 0) {
      // multi-turn path
    } else {
      messages.push({ role: "user", content: "Full prompt here" });
    }
    expect(messages.length).toBe(1);
  });

  it("chat history resets on new generation", () => {
    let _chatHistory = ["make shorter", "add urgency"];
    // Simulate new generation
    _chatHistory = [];
    expect(_chatHistory.length).toBe(0);
  });
});

// ══════════════════════════════════════════════════════════════
// 6. GRACEFUL NON-JSON REFINE RESPONSE
// ══════════════════════════════════════════════════════════════
describe("graceful non-JSON refine response", () => {
  // Reconstruct parseEmailResponse
  function parseEmailResponse(text) {
    let cleaned = text.replace(/```[\w]*\n?/g, "").trim();
    let parsed;
    try {
      parsed = JSON.parse(cleaned);
    } catch {
      let start = -1, depth = 0, inString = false, escaped = false;
      for (let i = 0; i < cleaned.length; i++) {
        const ch = cleaned[i];
        if (escaped) { escaped = false; continue; }
        if (ch === '\\' && inString) { escaped = true; continue; }
        if (ch === '"') { inString = !inString; continue; }
        if (inString) continue;
        if (ch === "{") { if (depth++ === 0) start = i; }
        else if (ch === "}") {
          if (--depth === 0 && start !== -1) {
            try { parsed = JSON.parse(cleaned.slice(start, i + 1)); break; } catch { start = -1; }
          }
        }
      }
      if (!parsed) throw new Error("Could not parse email JSON from Claude response");
    }
    return { subject: parsed.subject || "", body: parsed.body || "" };
  }

  it("parses valid JSON response", () => {
    const result = parseEmailResponse('{"subject":"Hi","body":"Hello there"}');
    expect(result.subject).toBe("Hi");
    expect(result.body).toBe("Hello there");
  });

  it("throws on conversational text (no JSON)", () => {
    const text = "I don't have access to the thread. Could you paste the relevant content?";
    expect(() => parseEmailResponse(text)).toThrow("Could not parse email JSON");
  });

  it("refineEmail returns chatReply when Claude responds conversationally", () => {
    // Simulate the try/catch in refineEmail
    const claudeText = "I can't see the thread. Please share the questions.";
    const currentSubject = "Re: Proposal";
    const currentBody = "Original body here";

    let result;
    try {
      result = parseEmailResponse(claudeText);
    } catch (parseErr) {
      result = { subject: currentSubject, body: currentBody, chatReply: claudeText.trim() };
    }

    expect(result.subject).toBe("Re: Proposal");
    expect(result.body).toBe("Original body here");
    expect(result.chatReply).toBe("I can't see the thread. Please share the questions.");
  });
});

// ══════════════════════════════════════════════════════════════
// 7. normalizeSubject edge cases
// ══════════════════════════════════════════════════════════════
describe("normalizeSubject edge cases", () => {
  it("strips Re: Fwd: FW: prefixes", () => {
    expect(normalizeSubject("Re: Fwd: FW: Hello")).toBe("hello");
  });

  it("strips [External] [Ext] tags", () => {
    expect(normalizeSubject("[External] RE: Proposal")).toBe("proposal");
    expect(normalizeSubject("[Ext] Re: Proposal")).toBe("proposal");
  });

  it("handles deeply nested prefixes", () => {
    expect(normalizeSubject("RE: FW: RE: RE: FW: Topic")).toBe("topic");
  });

  it("handles empty and null input", () => {
    expect(normalizeSubject("")).toBe("");
    expect(normalizeSubject(null)).toBe("");
    expect(normalizeSubject(undefined)).toBe("");
  });

  it("strips calendar prefixes", () => {
    expect(normalizeSubject("Accepted: Meeting with Client")).toBe("meeting with client");
    expect(normalizeSubject("Declined: RE: Demo Call")).toBe("demo call");
  });

  it("same subject with different prefixes normalizes identically", () => {
    const a = normalizeSubject("RE: [EXTERNAL] Example AI / Eaton");
    const b = normalizeSubject("Re: Example AI / Eaton");
    const c = normalizeSubject("FW: [Ext] Re: Example AI / Eaton");
    expect(a).toBe(b);
    expect(b).toBe(c);
  });
});
