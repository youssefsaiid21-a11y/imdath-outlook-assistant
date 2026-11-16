/**
 * styleProfile.js - Analyzes a user's sent emails to build a writing style profile.
 *
 * Extracts patterns from real sent emails including:
 * - Average email length (short/medium/long)
 * - Preferred greeting and sign-off
 * - Voice (formal, warm, direct, balanced)
 * - Punctuation habits (exclamation marks, em dashes, ellipses)
 * - Opener style (pleasantries first vs straight to the point)
 * - Characteristic phrases the sender commonly uses
 *
 * Claude uses this profile to ghostwrite emails that match the sender's voice.
 */
"use strict";

// Analyzes an array of sent emails and returns a style profile object.
// Each email should have { subject, bodyText } with HTML already stripped.
function extractStyleProfile(sentEmails) {
  if (!sentEmails || sentEmails.length === 0) {
    return defaultProfile();
  }

  const bodies = sentEmails.map((e) => e.bodyText || "").filter(Boolean);
  const allText = bodies.join("\n\n");

  const avgWords = Math.round(bodies.reduce((sum, b) => sum + countWords(b), 0) / bodies.length);
  return {
    avgLength: getAvgLength(avgWords), // pass pre-computed value to avoid re-calculation
    avgWords,
    greeting: detectGreeting(bodies),
    signOff: detectSignOff(bodies),
    voice: assessVoice(allText),
    punctuationHabits: assessPunctuation(allText),
    openerStyle: detectOpenerStyle(bodies),
    samplePhrases: extractSamplePhrases(bodies),
  };
}

// ── Length ────────────────────────────────────────────────────

// Categorizes average word count into a simple label for the style profile.
function getAvgLength(avgWords) {
  if (avgWords < 60)  return "short";
  if (avgWords < 150) return "medium";
  return "long";
}

function countWords(text) {
  return text.trim().split(/\s+/).filter(Boolean).length;
}

// ── Greeting Detection ────────────────────────────────────────

// Counts how often each greeting pattern appears in the first line of emails.
// Returns the most frequent one (e.g. "Hey [name]," or "Hi [name],").
function detectGreeting(bodies) {
  const patterns = [
    { re: /^hey\s/im, label: "Hey [name]," },
    { re: /^hi\s/im, label: "Hi [name]," },
    { re: /^hello\s/im, label: "Hello [name]," },
    { re: /^morning\s/im, label: "Morning [name]," },
    { re: /^dear\s/im, label: "Dear [name]," },
  ];

  const counts = {};
  for (const b of bodies) {
    const firstLine = b.split("\n")[0].trim();
    for (const p of patterns) {
      if (p.re.test(firstLine)) {
        counts[p.label] = (counts[p.label] || 0) + 1;
        break;
      }
    }
  }

  const top = Object.entries(counts).sort((a, b) => b[1] - a[1])[0];
  return top ? top[0] : "Hi [name],";
}

// ── Sign-off Detection ────────────────────────────────────────

// Scans the last 8 lines of each email for common sign-off patterns.
// Returns the most frequently used one (e.g. "Cheers," or "Best,").
function detectSignOff(bodies) {
  const patterns = [
    { re: /\btalks?\s+soon\b/i, label: "Talk soon," },
    { re: /\bcheers\b/i, label: "Cheers," },
    { re: /\bbest\b/i, label: "Best," },
    { re: /\bthanks\b/i, label: "Thanks," },
    { re: /\bkind\s+regards\b/i, label: "Kind regards," },
    { re: /\bregards\b/i, label: "Regards," },
    { re: /\blooking\s+forward\b/i, label: "Looking forward to it," },
    { re: /\bspeak\s+soon\b/i, label: "Speak soon," },
  ];

  const counts = {};
  for (const b of bodies) {
    const lastLines = b.split("\n").slice(-8).join(" ");
    for (const p of patterns) {
      if (p.re.test(lastLines)) {
        counts[p.label] = (counts[p.label] || 0) + 1;
        break;
      }
    }
  }

  const top = Object.entries(counts).sort((a, b) => b[1] - a[1])[0];
  return top ? top[0] : "Talk soon,";
}

// ── Voice Assessment ──────────────────────────────────────────

// Classifies the sender's overall tone by counting warm, direct, and formal
// keywords across all their emails. Whichever category dominates wins.
function assessVoice(text) {
  const warmWords = /\b(great|happy|glad|excited|pleasure|appreciate|thank|love|wonderful)\b/gi;
  const directWords = /\b(can you|please|need|want|let me|I will|here is|attached)\b/gi;
  const formalWords = /\b(furthermore|herewith|pursuant|regarding|enclosed)\b/gi;

  const warmCount = (text.match(warmWords) || []).length;
  const directCount = (text.match(directWords) || []).length;
  const formalCount = (text.match(formalWords) || []).length;

  if (formalCount > warmCount && formalCount > directCount) return "formal";
  if (warmCount > directCount) return "warm and conversational";
  if (directCount > warmCount) return "direct and action-oriented";
  return "balanced and professional";
}

// ── Punctuation Habits ────────────────────────────────────────

// Detects distinctive punctuation habits by measuring frequency per 100 words.
function assessPunctuation(text) {
  const habits = [];

  const exclamCount = (text.match(/!/g) || []).length;
  const wordCount = countWords(text);
  if (wordCount === 0) return "standard punctuation";
  const exclamRate = exclamCount / (wordCount / 100);

  if (exclamRate > 2) habits.push("frequent exclamation marks");
  else if (exclamRate < 0.5) habits.push("no exclamation marks");

  const emDashCount = (text.match(/—/g) || []).length;
  if (emDashCount > 2) habits.push("uses em dashes");

  const ellipsisCount = (text.match(/\.\.\./g) || []).length;
  if (ellipsisCount > 2) habits.push("uses ellipses");

  return habits.length > 0 ? habits.join(", ") : "standard punctuation";
}

// ── Opener Style ──────────────────────────────────────────────

// Determines if the sender typically starts with pleasantries ("Hope you're well")
// or goes straight to the point. Checks the second line of each email
// (first line is usually the greeting).
function detectOpenerStyle(bodies) {
  const pleasantriesRe =
    /^(hope|just wanted|thanks for|great to|good to|it was|wanted to reach)/im;

  let pleasantriesCount = 0;
  for (const b of bodies) {
    const lines = b.split("\n").filter(Boolean);
    const opener = lines.length > 1 ? lines[1] : lines[0] || "";
    if (pleasantriesRe.test(opener.trim())) pleasantriesCount++;
  }

  return pleasantriesCount > bodies.length * 0.4 ? "pleasantries first" : "straight to the point";
}

// ── Sample Phrases ────────────────────────────────────────────

// Checks for common conversational phrases the sender uses.
// These help Claude reproduce the sender's characteristic expressions.
function extractSamplePhrases(bodies) {
  const allText = bodies.join(" ").toLowerCase();
  const candidates = [
    "just wanted to",
    "let me know",
    "happy to",
    "feel free",
    "sounds good",
    "makes sense",
    "good to hear",
    "appreciate it",
  ];
  return candidates.filter((p) => allText.includes(p)).slice(0, 4);
}

// ── Default Profile ───────────────────────────────────────────

// Fallback profile used when no sent emails are available to analyze.
function defaultProfile() {
  return {
    avgLength: "medium",
    greeting: "Hi [name],",
    signOff: "Talk soon,",
    voice: "warm and professional",
    punctuationHabits: "standard punctuation",
    openerStyle: "straight to the point",
    samplePhrases: [],
  };
}

module.exports = { extractStyleProfile };
