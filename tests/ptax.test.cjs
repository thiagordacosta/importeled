const test = require("node:test");
const assert = require("node:assert/strict");
const { parsePtax, loadPtax } = require("../assets/ptax.js");
const { calculateQuote } = require("../assets/led-calculator.js");
const prices = require("../LED/data/prices_cache.json");
const saturday = new Date("2026-10-17T15:00:00Z");
const published = [{ data: "16/10/2026", valor: "5.0826" }];

test("Usa a última PTAX de venda publicada no fim de semana e mantém a data original", () => {
  const ptax = parsePtax(published, saturday);
  assert.equal(ptax.rate, 5.0826);
  assert.equal(ptax.date, "2026-10-16");
  const quote = calculateQuote({ application: "indoor", wall: "yes", purpose: "use", pitch: "1.86", area: "12" }, prices, ptax);
  assert.equal(quote.originTotal, 6660);
  assert.ok(Math.abs(quote.nationalizedTotal - 71085.2436) < 0.000001);
});

test("Respeita o dia de Brasília mesmo depois de meia-noite em UTC", () => {
  assert.throws(() => parsePtax([{ data: "18/10/2026", valor: "5" }], new Date("2026-10-18T01:00:00Z")), /indisponível/);
  assert.equal(parsePtax(published, new Date("2026-10-18T01:00:00Z")).date, "2026-10-16");
});

test("Não aceita PTAX vazia, inválida, antiga ou futura", () => {
  for (const payload of [[], {}, [{ data: "16/10/2026", valor: "NaN" }], [{ data: "16/10/2026", valor: "-1" }],
    [{ data: "18/10/2026", valor: "5" }], [{ data: "16/09/2026", valor: "5" }], [{ data: "2026-10-16", valor: "5" }]]) {
    assert.throws(() => parsePtax(payload, saturday), /indisponível/);
  }
});

test("Consulta somente o endpoint oficial e busca a cotação novamente a cada cálculo", async () => {
  let calls = 0;
  const request = async (url, options) => {
    assert.equal(url, "https://api.bcb.gov.br/dados/serie/bcdata.sgs.1/dados/ultimos/1?formato=json");
    assert.equal(options.cache, "no-store");
    calls += 1;
    return { ok: true, json: async () => [{ data: "16/10/2026", valor: calls === 1 ? "5" : "5.1" }] };
  };
  assert.equal((await loadPtax(request, saturday)).rate, 5);
  assert.equal((await loadPtax(request, saturday)).rate, 5.1);
  assert.equal(calls, 2);
});

test("Falha do BCB não produz uma cotação substituta", async () => {
  await assert.rejects(() => loadPtax(async () => ({ ok: false }), saturday), /Não foi possível consultar/);
  await assert.rejects(() => loadPtax(async () => { throw new Error("Offline"); }, saturday), /Não foi possível consultar/);
  await assert.rejects(() => loadPtax(async () => ({ ok: true, json: async () => [] }), saturday), /Não foi possível consultar/);
});

test("Preço nacionalizado exige PTAX válida em vez de usar o fator antigo", () => {
  const input = { application: "indoor", wall: "yes", purpose: "use", pitch: "1.86", area: "12" };
  for (const ptax of [undefined, {}, { rate: 0, date: "2026-10-16" }, { rate: NaN, date: "2026-10-16" }, { rate: 5 }]) {
    assert.throws(() => calculateQuote(input, prices, ptax), /PTAX válida/);
  }
});
