const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const { loadPtax } = require("../assets/ptax.js");
const { calculateQuote } = require("../assets/led-calculator.js");
const prices = require("../LED/data/prices_cache.json");
const input = { application: "indoor", wall: "yes", purpose: "use", pitch: "1.86", area: "12" };

test("Usa cotação fixa de 5,50 na fórmula origem × 2 × cotação", async () => {
  const ptax = await loadPtax();
  assert.equal(ptax.rate, 5.5);
  assert.equal(ptax.date, "2026-10-19");
  const quote = calculateQuote(input, prices, ptax);
  assert.equal(quote.originTotal, 6660);
  assert.ok(Math.abs(quote.nationalizedTotal - 73260) < 0.000001);
});

test("Carrega a referência no navegador sem consultar a rede", async () => {
  const context = vm.createContext({ fetch() { throw new Error("Network unavailable"); } });
  vm.runInContext(fs.readFileSync(require.resolve("../assets/ptax.js"), "utf8"), context);
  const ptax = await context.ImporteLEDPtax.loadPtax();
  assert.equal(ptax.rate, 5.5);
  assert.equal(ptax.date, "2026-10-19");
});

test("Um cálculo não altera a referência dos próximos cálculos", async () => {
  const first = await loadPtax();
  first.rate = 99;
  assert.equal((await loadPtax()).rate, 5.5);
});

test("Preço nacionalizado exige uma referência válida", () => {
  for (const ptax of [undefined, {}, { rate: 0, date: "2026-10-19" }, { rate: NaN, date: "2026-10-19" }, { rate: 5 }]) {
    assert.throws(() => calculateQuote(input, prices, ptax), /cotação válida/);
  }
});
