"use strict";

const { extractStyleProfile } = require("../server/styleProfile");

describe("extractStyleProfile", () => {
  it("returns default profile for empty input", () => {
    const result = extractStyleProfile([]);
    expect(result.avgLength).toBe("medium");
    expect(result.greeting).toBe("Hi [name],");
    expect(result.signOff).toBe("Talk soon,");
    expect(result.voice).toBe("warm and professional");
    expect(result.openerStyle).toBe("straight to the point");
    expect(result.samplePhrases).toEqual([]);
  });

  it("returns default profile for null input", () => {
    const result = extractStyleProfile(null);
    expect(result.avgLength).toBe("medium");
  });

  it("returns default profile for undefined input", () => {
    const result = extractStyleProfile(undefined);
    expect(result.greeting).toBe("Hi [name],");
  });

  it("detects short email length", () => {
    const emails = [
      { subject: "Quick update", bodyText: "Hey, just confirming the call tomorrow. Talk soon." },
      { subject: "Re: Meeting", bodyText: "Sounds good, see you then." },
    ];
    const result = extractStyleProfile(emails);
    expect(result.avgLength).toBe("short");
    expect(result.avgWords).toBeLessThan(60);
  });

  it("detects medium email length", () => {
    const body = "word ".repeat(100);
    const emails = [
      { subject: "Proposal", bodyText: body },
      { subject: "Follow up", bodyText: body },
    ];
    const result = extractStyleProfile(emails);
    expect(result.avgLength).toBe("medium");
  });

  it("detects long email length", () => {
    const body = "word ".repeat(200);
    const emails = [{ subject: "Report", bodyText: body }];
    const result = extractStyleProfile(emails);
    expect(result.avgLength).toBe("long");
    expect(result.avgWords).toBeGreaterThanOrEqual(150);
  });

  it("detects 'Hey' greeting pattern", () => {
    const emails = Array.from({ length: 5 }, () => ({
      subject: "Test",
      bodyText: "Hey John,\n\nJust a quick note about the project.\n\nBest,\nAlice",
    }));
    const result = extractStyleProfile(emails);
    expect(result.greeting).toBe("Hey [name],");
  });

  it("detects 'Hello' greeting pattern", () => {
    const emails = Array.from({ length: 5 }, () => ({
      subject: "Test",
      bodyText: "Hello team,\n\nPlease review the attached document.\n\nRegards,\nBob",
    }));
    const result = extractStyleProfile(emails);
    expect(result.greeting).toBe("Hello [name],");
  });

  it("detects 'Dear' greeting pattern", () => {
    const emails = Array.from({ length: 5 }, () => ({
      subject: "Test",
      bodyText: "Dear Mr. Smith,\n\nI am writing to confirm...\n\nKind regards,\nCarol",
    }));
    const result = extractStyleProfile(emails);
    expect(result.greeting).toBe("Dear [name],");
  });

  it("falls back to 'Hi [name],' when no greeting detected", () => {
    const emails = [{ subject: "Test", bodyText: "Just checking in on the proposal. Let me know." }];
    const result = extractStyleProfile(emails);
    expect(result.greeting).toBe("Hi [name],");
  });

  it("detects 'Cheers' sign-off", () => {
    const emails = Array.from({ length: 5 }, () => ({
      subject: "Test",
      bodyText: "Hi there,\n\nQuick update on the project.\n\nCheers,\nDave",
    }));
    const result = extractStyleProfile(emails);
    expect(result.signOff).toBe("Cheers,");
  });

  it("detects 'Best' sign-off", () => {
    const emails = Array.from({ length: 5 }, () => ({
      subject: "Test",
      bodyText: "Hi,\n\nHere's the document.\n\nBest,\nEve",
    }));
    const result = extractStyleProfile(emails);
    expect(result.signOff).toBe("Best,");
  });

  it("detects 'Kind regards' sign-off", () => {
    const emails = Array.from({ length: 5 }, () => ({
      subject: "Test",
      bodyText: "Dear Sir,\n\nPursuant to our conversation.\n\nKind regards,\nFrank",
    }));
    const result = extractStyleProfile(emails);
    expect(result.signOff).toBe("Kind regards,");
  });

  it("falls back to 'Talk soon,' when no sign-off detected", () => {
    const emails = [{ subject: "Test", bodyText: "Sounds good. See you at 3pm." }];
    const result = extractStyleProfile(emails);
    expect(result.signOff).toBe("Talk soon,");
  });

  it("assesses formal voice", () => {
    const emails = [{
      subject: "Test",
      bodyText: "Furthermore, I would like to inform you that pursuant to our agreement, the enclosed documents herewith are for your review. Regarding the matter at hand, we shall proceed accordingly.",
    }];
    const result = extractStyleProfile(emails);
    expect(result.voice).toBe("formal");
  });

  it("assesses warm and conversational voice", () => {
    const emails = [{
      subject: "Test",
      bodyText: "Great to hear from you! I'm really happy about the news. It's wonderful that we can move forward. I truly appreciate your help and I'm excited about what's next. Thank you so much!",
    }];
    const result = extractStyleProfile(emails);
    expect(result.voice).toBe("warm and conversational");
  });

  it("assesses direct and action-oriented voice", () => {
    const emails = [{
      subject: "Test",
      bodyText: "Can you send me the report? I need it by Friday. Please include the updated numbers. I will review and let me know if you want to discuss. Here is the link to the file.",
    }];
    const result = extractStyleProfile(emails);
    expect(result.voice).toBe("direct and action-oriented");
  });

  it("detects frequent exclamation marks", () => {
    const emails = [{
      subject: "Test",
      bodyText: "Amazing! Great work! Love it! Can't wait! So excited! This is fantastic! Wonderful!",
    }];
    const result = extractStyleProfile(emails);
    expect(result.punctuationHabits).toContain("frequent exclamation marks");
  });

  it("detects em dashes", () => {
    const emails = [{
      subject: "Test",
      bodyText: "The project — which started last month — is going well. We need to focus — more than ever — on the deliverables. The timeline — originally Q2 — has shifted.",
    }];
    const result = extractStyleProfile(emails);
    expect(result.punctuationHabits).toContain("uses em dashes");
  });

  it("detects ellipses", () => {
    const emails = [{
      subject: "Test",
      bodyText: "I was thinking... maybe we should try another approach... not sure yet though... let me get back to you... after I review the data...",
    }];
    const result = extractStyleProfile(emails);
    expect(result.punctuationHabits).toContain("uses ellipses");
  });

  it("returns 'no exclamation marks' for calm text", () => {
    const emails = [{
      subject: "Test",
      bodyText: "Please review the attached document and let me know your thoughts. I have included the updated figures. The deadline is Friday.",
    }];
    const result = extractStyleProfile(emails);
    // Short text with zero exclamation marks triggers "no exclamation marks"
    expect(result.punctuationHabits).toContain("no exclamation marks");
  });

  it("detects pleasantries-first opener style", () => {
    const emails = Array.from({ length: 5 }, () => ({
      subject: "Test",
      bodyText: "Hi John,\nHope you're doing well. Just wanted to check in on the project status.",
    }));
    const result = extractStyleProfile(emails);
    expect(result.openerStyle).toBe("pleasantries first");
  });

  it("detects straight-to-the-point opener style", () => {
    const emails = Array.from({ length: 5 }, () => ({
      subject: "Test",
      bodyText: "Hi John,\nThe quarterly numbers are in. Revenue is up 15% YoY.",
    }));
    const result = extractStyleProfile(emails);
    expect(result.openerStyle).toBe("straight to the point");
  });

  it("extracts sample phrases", () => {
    const emails = [{
      subject: "Test",
      bodyText: "Just wanted to check in. Let me know if you need anything. Feel free to reach out. Happy to help with anything else. Sounds good to me.",
    }];
    const result = extractStyleProfile(emails);
    expect(result.samplePhrases).toEqual(expect.arrayContaining(["let me know"]));
    expect(result.samplePhrases).toEqual(expect.arrayContaining(["happy to"]));
    expect(result.samplePhrases).toEqual(expect.arrayContaining(["feel free"]));
    // "sounds good" is in text but sample phrases are capped at 4, and earlier matches fill the list
  });

  it("limits sample phrases to 4", () => {
    const emails = [{
      subject: "Test",
      bodyText: "Just wanted to check in. Let me know. Happy to help. Feel free. Sounds good. Makes sense. Good to hear. Appreciate it.",
    }];
    const result = extractStyleProfile(emails);
    expect(result.samplePhrases.length).toBeLessThanOrEqual(4);
  });

  it("handles emails with empty bodyText", () => {
    const emails = [
      { subject: "Test", bodyText: "" },
      { subject: "Test2", bodyText: null },
    ];
    const result = extractStyleProfile(emails);
    expect(result).toBeDefined();
    expect(result.avgLength).toBeDefined();
  });

  it("handles single-word emails", () => {
    const emails = [{ subject: "Test", bodyText: "OK" }];
    const result = extractStyleProfile(emails);
    expect(result.avgLength).toBe("short");
    expect(result.avgWords).toBe(1);
  });
});
