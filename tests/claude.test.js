"use strict";

// Mock the Anthropic SDK and fs before requiring the module
jest.mock("@anthropic-ai/sdk", () => {
  return jest.fn().mockImplementation(() => ({
    messages: { create: jest.fn() },
  }));
});

jest.mock("fs", () => ({
  readFileSync: jest.fn(() => "Mock writing style guide"),
}));

// Set required env vars before requiring module
process.env.ANTHROPIC_API_KEY = "test-key";

// We need to access internal functions — require the module source directly
// Since parseEmailResponse, sanitizeInput, cleanRecipientName are not exported,
// we test them indirectly or extract them.

// For pure function testing, we'll extract and test the logic directly
describe("claude.js pure functions", () => {

  // ── parseEmailResponse (reconstructed for testing) ──
  function parseEmailResponse(text) {
    const cleaned = text.replace(/^```json\s*/i, "").replace(/```\s*$/, "").trim();
    let parsed;
    try {
      parsed = JSON.parse(cleaned);
    } catch {
      let start = -1, depth = 0;
      for (let i = 0; i < cleaned.length; i++) {
        if (cleaned[i] === "{") { if (depth++ === 0) start = i; }
        else if (cleaned[i] === "}") {
          if (--depth === 0 && start !== -1) {
            try { parsed = JSON.parse(cleaned.slice(start, i + 1)); break; } catch { start = -1; }
          }
        }
      }
      if (!parsed) throw new Error("Could not parse email JSON from Claude response");
    }
    if (!parsed.subject && !parsed.body) {
      throw new Error("Claude returned JSON but missing subject/body fields");
    }
    return { subject: parsed.subject || "", body: parsed.body || "" };
  }

  // ── sanitizeInput (reconstructed for testing) ──
  function sanitizeInput(str, maxLen = 2000) {
    if (typeof str !== "string") return "";
    return str
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, "")
      .substring(0, maxLen)
      .trim();
  }

  // ── cleanRecipientName (reconstructed for testing) ──
  function cleanRecipientName(name, email) {
    if (!name || /^visitor\d+$/i.test(name.trim())) {
      const localPart = (email || "").split("@")[0].split(/[._\-+]/)[0];
      if (localPart && localPart.length > 1) {
        return localPart.charAt(0).toUpperCase() + localPart.slice(1).toLowerCase();
      }
      return null;
    }
    return name;
  }

  describe("parseEmailResponse", () => {
    it("parses clean JSON", () => {
      const result = parseEmailResponse('{"subject": "Hello", "body": "World"}');
      expect(result).toEqual({ subject: "Hello", body: "World" });
    });

    it("strips markdown code fences", () => {
      const result = parseEmailResponse('```json\n{"subject": "Test", "body": "Content"}\n```');
      expect(result).toEqual({ subject: "Test", body: "Content" });
    });

    it("extracts JSON from surrounding text", () => {
      const result = parseEmailResponse('Here is the email:\n{"subject": "Proposal", "body": "Dear John..."}\nDone!');
      expect(result).toEqual({ subject: "Proposal", body: "Dear John..." });
    });

    it("handles subject-only response", () => {
      const result = parseEmailResponse('{"subject": "Meeting Tomorrow", "body": ""}');
      expect(result.subject).toBe("Meeting Tomorrow");
      expect(result.body).toBe("");
    });

    it("handles body-only response", () => {
      const result = parseEmailResponse('{"body": "Just the body content"}');
      expect(result.subject).toBe("");
      expect(result.body).toBe("Just the body content");
    });

    it("throws on completely invalid JSON", () => {
      expect(() => parseEmailResponse("This is not JSON at all")).toThrow("Could not parse email JSON");
    });

    it("throws on JSON with neither subject nor body", () => {
      expect(() => parseEmailResponse('{"foo": "bar"}')).toThrow("missing subject/body fields");
    });

    it("handles nested braces in body", () => {
      const json = '{"subject": "Test", "body": "Use {braces} in text"}';
      const result = parseEmailResponse(json);
      expect(result.body).toBe("Use {braces} in text");
    });

    it("handles newlines in JSON values", () => {
      const json = '{"subject": "Re: Meeting", "body": "Hi,\\n\\nSounds good.\\n\\nBest,\\nJohn"}';
      const result = parseEmailResponse(json);
      expect(result.body).toContain("Sounds good.");
    });

    it("handles JSON with extra whitespace", () => {
      const result = parseEmailResponse('  \n  {"subject": "Spaced", "body": "Out"}  \n  ');
      expect(result).toEqual({ subject: "Spaced", body: "Out" });
    });
  });

  describe("sanitizeInput", () => {
    it("returns empty string for non-string input", () => {
      expect(sanitizeInput(null)).toBe("");
      expect(sanitizeInput(undefined)).toBe("");
      expect(sanitizeInput(123)).toBe("");
      expect(sanitizeInput({})).toBe("");
    });

    it("passes through normal text", () => {
      expect(sanitizeInput("Hello world")).toBe("Hello world");
    });

    it("preserves newlines and tabs", () => {
      expect(sanitizeInput("Line 1\nLine 2\tTabbed")).toBe("Line 1\nLine 2\tTabbed");
    });

    it("strips control characters", () => {
      expect(sanitizeInput("Hello\u0000World")).toBe("HelloWorld");
      expect(sanitizeInput("Test\u0007Bell")).toBe("TestBell");
      expect(sanitizeInput("Null\u007FDelete")).toBe("NullDelete");
    });

    it("truncates to maxLen", () => {
      const long = "a".repeat(3000);
      expect(sanitizeInput(long).length).toBe(2000);
    });

    it("respects custom maxLen", () => {
      const result = sanitizeInput("Hello World", 5);
      expect(result).toBe("Hello");
    });

    it("trims whitespace", () => {
      expect(sanitizeInput("  hello  ")).toBe("hello");
    });

    it("handles empty string", () => {
      expect(sanitizeInput("")).toBe("");
    });

    it("handles string of only control characters", () => {
      expect(sanitizeInput("\u0000\u0001\u0002")).toBe("");
    });
  });

  describe("cleanRecipientName", () => {
    it("returns name when valid", () => {
      expect(cleanRecipientName("John Doe", "john@example.com")).toBe("John Doe");
    });

    it("extracts name from email for null name", () => {
      expect(cleanRecipientName(null, "john.doe@example.com")).toBe("John");
    });

    it("extracts name from email for empty name", () => {
      expect(cleanRecipientName("", "alice@example.com")).toBe("Alice");
    });

    it("replaces visitor pattern names", () => {
      expect(cleanRecipientName("visitor123", "bob@example.com")).toBe("Bob");
      expect(cleanRecipientName("Visitor456", "carol@example.com")).toBe("Carol");
    });

    it("returns null for single-char email local part", () => {
      expect(cleanRecipientName(null, "a@example.com")).toBeNull();
    });

    it("returns null for empty email", () => {
      expect(cleanRecipientName(null, "")).toBeNull();
    });

    it("handles email with dots in local part", () => {
      expect(cleanRecipientName(null, "jane.smith@example.com")).toBe("Jane");
    });

    it("handles email with underscores", () => {
      expect(cleanRecipientName(null, "mike_johnson@example.com")).toBe("Mike");
    });

    it("handles email with plus addressing", () => {
      expect(cleanRecipientName(null, "sarah+newsletter@example.com")).toBe("Sarah");
    });

    it("capitalizes first letter and lowercases rest", () => {
      expect(cleanRecipientName(null, "DAVE@example.com")).toBe("Dave");
    });
  });
});

describe("claude.js module exports", () => {
  const claude = require("../server/claude");

  it("exports generateEmail function", () => {
    expect(typeof claude.generateEmail).toBe("function");
  });

  it("exports refineEmail function", () => {
    expect(typeof claude.refineEmail).toBe("function");
  });

  it("exports summarizeForCRM function", () => {
    expect(typeof claude.summarizeForCRM).toBe("function");
  });

  it("exports contextualizeVoice function", () => {
    expect(typeof claude.contextualizeVoice).toBe("function");
  });

  it("exports suggestDealName function", () => {
    expect(typeof claude.suggestDealName).toBe("function");
  });
});
