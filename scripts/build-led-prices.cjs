const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const prices = JSON.parse(fs.readFileSync(path.join(root, "LED/data/prices_cache.json"), "utf8"));
const output = "// Generated from LED/data/prices_cache.json. Run: node scripts/build-led-prices.cjs\n"
  + `globalThis.ImporteLEDPriceTable = ${JSON.stringify(prices)};\n`;
fs.writeFileSync(path.join(root, "assets/led-prices.js"), output);
