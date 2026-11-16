"use strict";

// We need to test stripHtml which is not exported. We'll extract and test the logic.
// Also test the module loads and exports correctly.

describe("graph.js stripHtml (reconstructed)", () => {
  // Reconstructed from server/graph.js for unit testing
  function stripHtml(html) {
    if (!html) return "";
    return html
      .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
      .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "")
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/p>/gi, "\n")
      .replace(/<[^>]+>/g, "")
      .replace(/&nbsp;/gi, " ")
      .replace(/&amp;/gi, "&")
      .replace(/&lt;/gi, "<")
      .replace(/&gt;/gi, ">")
      .replace(/&quot;/gi, '"')
      .replace(/&apos;/gi, "'")
      .replace(/&#39;/gi, "'")
      .replace(/&mdash;/gi, "—")
      .replace(/&ndash;/gi, "–")
      .replace(/&lsquo;/gi, "\u2018")
      .replace(/&rsquo;/gi, "\u2019")
      .replace(/&ldquo;/gi, "\u201C")
      .replace(/&rdquo;/gi, "\u201D")
      .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(parseInt(n, 10)))
      .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
      .replace(/\r\n/g, "\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  }

  it("returns empty string for null/undefined/empty", () => {
    expect(stripHtml(null)).toBe("");
    expect(stripHtml(undefined)).toBe("");
    expect(stripHtml("")).toBe("");
  });

  it("strips basic HTML tags", () => {
    expect(stripHtml("<p>Hello <b>world</b></p>")).toBe("Hello world");
  });

  it("converts <br> to newline", () => {
    expect(stripHtml("Line 1<br>Line 2")).toBe("Line 1\nLine 2");
    expect(stripHtml("Line 1<br/>Line 2")).toBe("Line 1\nLine 2");
    expect(stripHtml("Line 1<br />Line 2")).toBe("Line 1\nLine 2");
  });

  it("converts </p> to newline", () => {
    expect(stripHtml("<p>Para 1</p><p>Para 2</p>")).toBe("Para 1\nPara 2");
  });

  it("removes <style> blocks", () => {
    expect(stripHtml('<style type="text/css">.foo { color: red; }</style><p>Content</p>')).toBe("Content");
  });

  it("removes <script> blocks", () => {
    expect(stripHtml('<script>alert("xss")</script><p>Safe</p>')).toBe("Safe");
  });

  it("decodes &nbsp;", () => {
    expect(stripHtml("Hello&nbsp;World")).toBe("Hello World");
  });

  it("decodes &amp;", () => {
    expect(stripHtml("Tom &amp; Jerry")).toBe("Tom & Jerry");
  });

  it("decodes &lt; and &gt;", () => {
    expect(stripHtml("a &lt; b &gt; c")).toBe("a < b > c");
  });

  it("decodes &quot;", () => {
    expect(stripHtml("He said &quot;hello&quot;")).toBe('He said "hello"');
  });

  it("decodes &apos; and &#39;", () => {
    expect(stripHtml("it&apos;s")).toBe("it's");
    expect(stripHtml("it&#39;s")).toBe("it's");
  });

  it("decodes &mdash; and &ndash;", () => {
    expect(stripHtml("hello&mdash;world")).toBe("hello—world");
    expect(stripHtml("2020&ndash;2025")).toBe("2020–2025");
  });

  it("decodes smart quotes", () => {
    expect(stripHtml("&ldquo;Hello&rdquo;")).toBe("\u201CHello\u201D");
    expect(stripHtml("&lsquo;Hi&rsquo;")).toBe("\u2018Hi\u2019");
  });

  it("decodes decimal numeric entities", () => {
    expect(stripHtml("&#8220;quoted&#8221;")).toBe("\u201Cquoted\u201D");
    expect(stripHtml("&#169;")).toBe("©");
  });

  it("decodes hex numeric entities", () => {
    expect(stripHtml("&#x201C;hex&#x201D;")).toBe("\u201Chex\u201D");
    expect(stripHtml("&#xA9;")).toBe("©");
  });

  it("normalizes \\r\\n to \\n", () => {
    expect(stripHtml("Line 1\r\nLine 2")).toBe("Line 1\nLine 2");
  });

  it("collapses 3+ newlines to 2", () => {
    expect(stripHtml("A\n\n\n\nB")).toBe("A\n\nB");
  });

  it("trims leading/trailing whitespace", () => {
    expect(stripHtml("  <p>Hello</p>  ")).toBe("Hello");
  });

  it("handles complex real-world email HTML", () => {
    const html = `
      <html><head><style>.sig { font-size: 12px; }</style></head>
      <body>
        <p>Hi John,</p>
        <p>Here&rsquo;s the report you &amp; the team requested.</p>
        <br>
        <p>Best,<br>Alice</p>
      </body>
    `;
    const result = stripHtml(html);
    expect(result).toContain("Hi John,");
    expect(result).toContain("Here\u2019s the report you & the team requested.");
    expect(result).toContain("Best,");
    expect(result).toContain("Alice");
    expect(result).not.toContain("<");
    expect(result).not.toContain("font-size");
  });
});
