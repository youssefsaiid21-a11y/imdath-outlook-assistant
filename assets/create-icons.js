/**
 * Generates icon PNGs from portfolio-logo.svg for the Sphinx Outlook Add-in.
 * Run from the project root: node assets/create-icons.js
 */

"use strict";

const sharp = require("sharp");
const fs    = require("fs");
const path  = require("path");

const SVG_PATH = path.join(__dirname, "portfolio-logo.svg");
const OUT_DIR  = __dirname;

const SIZES = [
  { size: 80, file: "icon-80.png" },
  { size: 32, file: "icon-32.png" },
  { size: 16, file: "icon-16.png" },
];

async function generate() {
  const svg = fs.readFileSync(SVG_PATH);

  for (const { size, file } of SIZES) {
    const dest = path.join(OUT_DIR, file);
    await sharp(svg)
      .resize(size, size, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } })
      .png()
      .toFile(dest);
    console.log(`✓ ${file} (${size}×${size}px, transparent background)`);
  }

  console.log("\nDone. Upload assets/ to the server to update the icons.");
}

generate().catch((err) => {
  console.error("Error generating icons:", err.message);
  process.exit(1);
});
