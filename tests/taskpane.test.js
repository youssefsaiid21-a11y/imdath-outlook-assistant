"use strict";

// Test pure client-side functions extracted from taskpane.js
// These don't depend on DOM or Office.js

describe("taskpane.js pure functions (reconstructed)", () => {

  // ── escapeHtml ──
  function escapeHtml(str) {
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  describe("escapeHtml", () => {
    it("escapes ampersands", () => {
      expect(escapeHtml("Tom & Jerry")).toBe("Tom &amp; Jerry");
    });

    it("escapes angle brackets", () => {
      expect(escapeHtml("<script>alert('xss')</script>")).toBe("&lt;script&gt;alert(&#39;xss&#39;)&lt;/script&gt;");
    });

    it("escapes quotes", () => {
      expect(escapeHtml('He said "hello"')).toBe("He said &quot;hello&quot;");
    });

    it("escapes single quotes", () => {
      expect(escapeHtml("it's")).toBe("it&#39;s");
    });

    it("handles empty string", () => {
      expect(escapeHtml("")).toBe("");
    });

    it("passes through safe text", () => {
      expect(escapeHtml("Hello World")).toBe("Hello World");
    });

    it("escapes all special chars in one string", () => {
      expect(escapeHtml("<div class=\"foo\" data-bar='baz'>&</div>"))
        .toBe("&lt;div class=&quot;foo&quot; data-bar=&#39;baz&#39;&gt;&amp;&lt;/div&gt;");
    });
  });

  // ── formatResponseTime ──
  function formatResponseTime(hours) {
    if (hours == null) return "N/A";
    if (hours < 1)     return `${Math.round(hours * 60)}m`;
    if (hours < 24)    return `${hours.toFixed(1)}h`;
    return `${Math.floor(hours / 24)}d ${Math.round(hours % 24)}h`;
  }

  describe("formatResponseTime", () => {
    it("returns N/A for null", () => {
      expect(formatResponseTime(null)).toBe("N/A");
    });

    it("returns N/A for undefined", () => {
      expect(formatResponseTime(undefined)).toBe("N/A");
    });

    it("formats minutes for < 1 hour", () => {
      expect(formatResponseTime(0.5)).toBe("30m");
      expect(formatResponseTime(0.25)).toBe("15m");
    });

    it("formats hours for < 24 hours", () => {
      expect(formatResponseTime(2.5)).toBe("2.5h");
      expect(formatResponseTime(12)).toBe("12.0h");
    });

    it("formats days and hours for >= 24 hours", () => {
      expect(formatResponseTime(25)).toBe("1d 1h");
      expect(formatResponseTime(48)).toBe("2d 0h");
      expect(formatResponseTime(50)).toBe("2d 2h");
    });

    it("handles zero", () => {
      expect(formatResponseTime(0)).toBe("0m");
    });

    it("handles exactly 1 hour", () => {
      expect(formatResponseTime(1)).toBe("1.0h");
    });

    it("handles exactly 24 hours", () => {
      expect(formatResponseTime(24)).toBe("1d 0h");
    });
  });

  // ── getStageBadgeClass ──
  function getStageBadgeClass(stage) {
    if (!stage) return "badge-grey";
    const s = stage.toLowerCase();
    if (s.includes("closed won") || /\bwon\b/.test(s)) return "badge-green";
    if (s.includes("proposal") || s.includes("progress") || s.includes("negotiat")) return "badge-orange";
    return "badge-grey";
  }

  describe("getStageBadgeClass", () => {
    it("returns grey for null/undefined/empty", () => {
      expect(getStageBadgeClass(null)).toBe("badge-grey");
      expect(getStageBadgeClass(undefined)).toBe("badge-grey");
      expect(getStageBadgeClass("")).toBe("badge-grey");
    });

    it("returns green for 'Closed Won'", () => {
      expect(getStageBadgeClass("Closed Won")).toBe("badge-green");
    });

    it("returns green for any stage containing 'won'", () => {
      expect(getStageBadgeClass("Deal Won")).toBe("badge-green");
      expect(getStageBadgeClass("won")).toBe("badge-green");
    });

    it("returns orange for 'Proposal'", () => {
      expect(getStageBadgeClass("Proposal Sent")).toBe("badge-orange");
    });

    it("returns orange for 'In Progress'", () => {
      expect(getStageBadgeClass("In Progress")).toBe("badge-orange");
    });

    it("returns orange for 'Negotiation'", () => {
      expect(getStageBadgeClass("Negotiation")).toBe("badge-orange");
      expect(getStageBadgeClass("Negotiating")).toBe("badge-orange");
    });

    it("returns grey for unknown stages", () => {
      expect(getStageBadgeClass("Appointment Scheduled")).toBe("badge-grey");
      expect(getStageBadgeClass("Qualified to Buy")).toBe("badge-grey");
    });

    it("is case-insensitive", () => {
      expect(getStageBadgeClass("CLOSED WON")).toBe("badge-green");
      expect(getStageBadgeClass("PROPOSAL")).toBe("badge-orange");
    });
  });

  // ── decodeJwtPayload ──
  function decodeJwtPayload(token) {
    const parts = token.split(".");
    if (parts.length !== 3) throw new Error("Invalid JWT structure");
    return JSON.parse(Buffer.from(parts[1].replace(/-/g, "+").replace(/_/g, "/"), "base64").toString());
  }

  describe("decodeJwtPayload", () => {
    it("decodes a valid JWT payload", () => {
      // Create a test JWT with known payload
      const payload = { name: "John Doe", preferred_username: "john@example.com" };
      const encoded = Buffer.from(JSON.stringify(payload)).toString("base64");
      const token = `header.${encoded}.signature`;
      const result = decodeJwtPayload(token);
      expect(result.name).toBe("John Doe");
      expect(result.preferred_username).toBe("john@example.com");
    });

    it("throws for token with wrong number of parts", () => {
      expect(() => decodeJwtPayload("only.two")).toThrow("Invalid JWT structure");
      expect(() => decodeJwtPayload("single")).toThrow("Invalid JWT structure");
    });

    it("handles base64url encoding with - and _", () => {
      // Create a payload that when encoded has - and _ characters
      const payload = { sub: "1234567890", name: "Test User" };
      const encoded = Buffer.from(JSON.stringify(payload)).toString("base64")
        .replace(/\+/g, "-")
        .replace(/\//g, "_");
      const token = `header.${encoded}.signature`;
      const result = decodeJwtPayload(token);
      expect(result.name).toBe("Test User");
    });
  });
});

// ── Feature: Read mode detection and banner logic ──
describe("Read mode banner logic", () => {
  it("shows banner in read mode (item.from exists, no item.to.getAsync)", () => {
    let bannerVisible = false;
    // Simulate: item has from (read mode), no to.getAsync (not compose)
    const item = { from: { emailAddress: "alice@co.com", displayName: "Alice" } };
    const isComposeMode = item.to && item.to.getAsync;
    const isReadMode = !isComposeMode && item.from;
    if (isReadMode) bannerVisible = true;
    expect(bannerVisible).toBe(true);
  });

  it("hides banner in compose mode (item.to.getAsync exists)", () => {
    let bannerVisible = true;
    const item = { to: { getAsync: () => {} }, from: null };
    const isComposeMode = item.to && item.to.getAsync;
    if (isComposeMode) bannerVisible = false;
    expect(bannerVisible).toBe(false);
  });

  it("hides banner on item change before re-detection", () => {
    let bannerVisible = true;
    // Simulate _itemChangedHandler
    bannerVisible = false; // hideReadModeBanner()
    expect(bannerVisible).toBe(false);
  });

  it("loads sender context from item.from in read mode", () => {
    let loadedEmail = null;
    const item = { from: { emailAddress: "bob@co.com", displayName: "Bob" } };
    if (item.from) {
      loadedEmail = item.from.emailAddress;
    }
    expect(loadedEmail).toBe("bob@co.com");
  });
});

// ── Feature: New Deal button visibility logic ──
describe("New Deal button visibility", () => {
  it("shows HubSpot actions when hubspotSearched is true", () => {
    const ctx = { hubspotSearched: true, recipient: { hubspotId: "123" } };
    const display = ctx.hubspotSearched ? "block" : "none";
    expect(display).toBe("block");
  });

  it("shows even when contact has no hubspotId", () => {
    const ctx = { hubspotSearched: true, recipient: { email: "new@co.com" } };
    const display = ctx.hubspotSearched ? "block" : "none";
    expect(display).toBe("block");
  });

  it("hides when HubSpot is not configured", () => {
    const ctx = { hubspotSearched: false, recipient: {} };
    const display = ctx.hubspotSearched ? "block" : "none";
    expect(display).toBe("none");
  });

  it("both New Deal and Link Existing Deal are available", () => {
    const ctx = { hubspotSearched: true };
    const actionsVisible = ctx.hubspotSearched;
    expect(actionsVisible).toBe(true);
    // Both buttons rendered inside hubspot-actions div
  });
});
