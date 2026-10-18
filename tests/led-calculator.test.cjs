const test = require("node:test");
const assert = require("node:assert/strict");
const { calculateQuote: calculateWithPtax, availablePitches, sendCalculationAlert, normalizeWhatsApp } = require("../assets/led-calculator.js");
const prices = require("../LED/data/prices_cache.json");
const ptax = { rate: 5, date: "2026-10-16" }; // Fixed fixture, not the live exchange rate.
const calculateQuote = (input, prices) => calculateWithPtax(input, prices, ptax);
const base = { application: "indoor", pitch: "1.86", wall: "yes", purpose: "use", area: "12" };

test("Calcula origem em dólar e nacionalizado com PTAX, fator 2 e acréscimo de 5%, aceitando área com vírgula", () => {
  const quote = calculateQuote({ ...base, area: "1,25" }, prices);
  assert.equal(quote.usdPerM2, 555);
  assert.equal(quote.originTotal, 693.75);
  assert.equal(quote.nationalizedTotal, 7284.375);
  assert.equal(quote.nationalizedPerM2, 5827.5);
});

test("Lista apenas os pitches disponíveis, ordenados e sem duplicação", () => {
  assert.deepEqual(availablePitches(base, prices), [1.86, 1.95, 2, 2.5, 2.604, 2.976, 3, 3.076, 3.91, 4, 4.81, 5]);
  assert.deepEqual(availablePitches({ ...base, application: "outdoor" }, prices), [2.976, 3.91, 4.81, 6.25, 7.81, 10.42]);
  assert.deepEqual(availablePitches({ ...base, purpose: "rental" }, prices), [1.56, 1.95, 2.604, 2.976, 3.91]);
  assert.deepEqual(availablePitches({ ...base, application: "outdoor", wall: "no" }, prices), [2.604, 2.976, 3.91]);
});

test("Calcula o pitch escolhido e rejeita modelos indisponíveis sem substituí-los", () => {
  const quote = calculateQuote({ ...base, pitch: "2,5" }, prices);
  assert.equal(quote.pitch, 2.5);
  assert.equal(quote.usdPerM2, 396);
  assert.throws(() => calculateQuote({ ...base, application: "outdoor" }, prices), /Não há preço/);
  assert.throws(() => calculateQuote({ ...base, pitch: "2.501" }, prices), /Não há preço/);
});

test("Outdoor fixo usa tabela outdoor e não recebe margem de revenda", () => {
  const quote = calculateQuote({ ...base, application: "outdoor", pitch: "4.81", purpose: "resale", area: "12,5" }, prices);
  assert.equal(quote.family, "outdoor_fixed_front");
  assert.equal(quote.pitch, 4.81);
  assert.equal(quote.originTotal, 7687.5);
  assert.equal(quote.nationalizedTotal, 80718.75);
  assert.equal(quote.originTotal, calculateQuote({ ...quote, purpose: "use" }, prices).originTotal);
});

test("Locação ou montagem móvel seleciona a linha rental da aplicação", () => {
  const indoor = calculateQuote({ ...base, pitch: "2.604", wall: "no" }, prices);
  assert.equal(indoor.family, "rental_indoor");
  assert.equal(indoor.originTotal, 9612);
  const outdoor = calculateQuote({ ...base, pitch: "2.604", application: "outdoor", purpose: "rental" }, prices);
  assert.equal(outdoor.family, "rental_outdoor");
  assert.equal(outdoor.originTotal, 12756);
  assert.equal(outdoor.nationalizedTotal, 133938);
});

test("Rejeita área e pitch inválidos sem produzir estimativa", () => {
  for (const key of ["area", "pitch"]) {
    for (const value of ["", "0", "-1", "1,2,3", "12m", "Infinity", "NaN", "1e9"]) {
      assert.throws(() => calculateQuote({ ...base, [key]: value }, prices));
    }
  }
  assert.throws(() => calculateQuote({ ...base, application: "unknown" }, prices));
  assert.throws(() => calculateQuote({ ...base, wall: "unknown" }, prices));
  assert.throws(() => calculateQuote({ ...base, purpose: "toString" }, prices));
});

test("Tabela ausente ou preços inválidos não geram um preço falso", () => {
  for (const table of [{}, { indoor_fixed_front: {} }, { indoor_fixed_front: [null, { pitch: 1.86, usd_per_m2: -1 }] }]) {
    assert.throws(() => calculateQuote(base, table), /Não há preço disponível/);
  }
});

test("Cada opção listada tem preço em todas as combinações de aplicação, finalidade e instalação", () => {
  let combinations = 0;
  for (const application of ["indoor", "outdoor"])
    for (const wall of ["yes", "no"])
      for (const purpose of ["use", "resale", "rental"])
        for (const pitch of availablePitches({ ...base, application, wall, purpose }, prices)) {
          const quote = calculateQuote({ ...base, application, wall, purpose, pitch }, prices);
          assert.ok(quote.originTotal > 0);
          assert.equal(quote.nationalizedTotal, quote.originTotal * 2 * ptax.rate * 1.05);
          combinations += 1;
        }
  assert.equal(combinations, 68);
});

test("Dropdown muda com aplicação e linha, preservando apenas seleções compatíveis", async () => {
  const vm = require("node:vm");
  const fs = require("node:fs");
  const handlers = {}, nodes = {};
  const fields = { ...base };
  const node = (id) => nodes[id] ||= { hidden: false, value: "", textContent: "", addEventListener() {} };
  const select = node("calculator-pitch");
  select.replaceChildren = (...options) => { select.options = options; select.value = options[0]?.value || ""; };
  const form = {
    querySelector: () => node("button"),
    addEventListener: (event, handler) => (handlers[event] ||= []).push(handler),
  };
  class Fields extends Map { constructor() { super(Object.entries(fields)); } }
  class SelectOption { constructor(label, value) { this.label = label; this.value = value; } }
  vm.runInNewContext(fs.readFileSync(require.resolve("../assets/led-calculator.js"), "utf8"), {
    document: { querySelector: (selector) => selector === "#led-calculator" ? form : node(selector.slice(1)), getElementById: node },
    FormData: Fields, Option: SelectOption, AbortController, setTimeout, clearTimeout,
    fetch: async () => ({ ok: true, json: async () => prices }),
  });
  const settle = () => new Promise(setImmediate);
  const change = async (name, value) => {
    fields[name] = value;
    handlers.change.forEach((handler) => handler({ target: { name } }));
    await settle();
  };
  await settle();
  assert.equal(select.disabled, false);
  assert.equal(select.options.length, 13);
  select.value = "1.86";
  await change("application", "outdoor");
  assert.deepEqual(select.options.map((option) => option.value), ["", "2.976", "3.91", "4.81", "6.25", "7.81", "10.42"]);
  assert.equal(select.value, "");
  select.value = "3.91";
  await change("purpose", "rental");
  assert.deepEqual(select.options.map((option) => option.value), ["", "2.604", "2.976", "3.91"]);
  assert.equal(select.value, "3.91");
});

test("Tabela incluída na página mantém os preços originais e permite dropdown e cálculo sem fetch", async () => {
  const vm = require("node:vm");
  const fs = require("node:fs");
  const handlers = {}, nodes = {};
  const fields = { application: "indoor", wall: "yes", purpose: "use", area: "12", whatsapp: "5511999999999" };
  const node = (id) => nodes[id] ||= { hidden: false, value: "", textContent: "", addEventListener() {} };
  const select = node("calculator-pitch");
  select.replaceChildren = (...options) => { select.options = options; select.value = options[0]?.value || ""; };
  const form = {
    querySelector: () => node("button"), reportValidity: () => true,
    addEventListener: (event, handler) => (handlers[event] ||= []).push(handler),
    setAttribute() {}, removeAttribute() {},
  };
  class Fields extends Map { constructor() { super(Object.entries({ ...fields, pitch: select.value })); } }
  class SelectOption { constructor(label, value) { this.label = label; this.value = value; } }
  const context = vm.createContext({
    document: { querySelector: (selector) => selector === "#led-calculator" ? form : node(selector.slice(1)), getElementById: node },
    FormData: Fields, Option: SelectOption,
    ImporteLEDPtax: { loadPtax: async () => ptax },
    fetch() { throw new Error("Local preview must not depend on fetch"); },
  });
  vm.runInContext(fs.readFileSync(require.resolve("../assets/led-prices.js"), "utf8"), context);
  assert.deepEqual(JSON.parse(JSON.stringify(context.ImporteLEDPriceTable)), prices);
  vm.runInContext(fs.readFileSync(require.resolve("../assets/led-calculator.js"), "utf8"), context);
  await new Promise(setImmediate);
  assert.equal(select.disabled, false);
  assert.equal(select.options.length, 13);
  fields.application = "outdoor";
  handlers.change.forEach((handler) => handler({ target: { name: "application" } }));
  await new Promise(setImmediate);
  assert.equal(select.options.length, 7);
  select.value = "4.81";
  await handlers.submit[0]({ preventDefault() {} });
  assert.equal(node("calculator-values").hidden, false);
  assert.match(node("calculator-origin-total").textContent, /7\.380,00/);
  assert.match(node("calculator-nationalized-total").textContent, /77\.490,00/);
  assert.equal(node("button").disabled, false);

  // Email delivery is independent of rendering and invalid submissions send nothing.
  let alertRequests = 0, alertWarnings = 0;
  form.dataset = { alertEndpoint: "https://example.test/alerts" };
  Object.assign(context, {
    window: { location: { origin: "https://example.test", pathname: "/" } },
    AbortController, setTimeout, clearTimeout,
    console: { warn() { alertWarnings += 1; } },
    fetch: async () => { alertRequests += 1; throw new Error("Email provider unavailable"); },
  });
  await handlers.submit[0]({ preventDefault() {} });
  await new Promise(setImmediate);
  assert.equal(alertRequests, 1);
  assert.equal(alertWarnings, 1);
  assert.equal(node("calculator-values").hidden, false);
  assert.match(node("calculator-status").textContent, /Estimativa calculada/);
  fields.area = "0";
  await handlers.submit[0]({ preventDefault() {} });
  assert.equal(alertRequests, 1);
  fields.area = "12";
  fields.whatsapp = "telefone-invalido";
  await handlers.submit[0]({ preventDefault() {} });
  assert.equal(alertRequests, 1);
  assert.equal(node("calculator-values").hidden, true);
  assert.match(node("calculator-status").textContent, /número válido/);
});

test("Alerta envia os dados e valores da simulação para o endpoint configurado", async () => {
  const quote = calculateQuote({ ...base, whatsapp: "5511999999999" }, prices);
  let sent;
  await sendCalculationAlert(quote, "https://example.test/alerts", "https://importeled.club/", async (url, options) => {
    sent = { url, ...options, payload: JSON.parse(options.body) };
    return { ok: true, json: async () => ({ success: "true" }) };
  });
  assert.equal(sent.url, "https://example.test/alerts");
  assert.equal(sent.method, "POST");
  assert.equal(sent.keepalive, true);
  assert.equal(sent.payload.aplicacao, "Indoor / interno");
  assert.equal(sent.payload.fixa_na_parede, "Sim");
  assert.equal(sent.payload.pixel_pitch, "P1.86");
  assert.equal(sent.payload.area_m2, 12);
  assert.equal(sent.payload.preco_total_origem_usd, 6660);
  assert.equal(sent.payload.preco_total_nacionalizado_brl, 69930);
  assert.equal(sent.payload.ptax_venda, 5);
  assert.equal(sent.payload.data_ptax, "2026-10-16");
  assert.equal(sent.payload.whatsapp, "5511999999999");
  assert.equal(sent.payload.contato_whatsapp, "https://wa.me/5511999999999");
  assert.equal(Object.hasOwn(sent.payload, "email"), false);
  assert.equal(Object.hasOwn(sent.payload, "_replyto"), false);
});

test("Normaliza WhatsApp com DDD ou código de país e rejeita números inválidos", () => {
  assert.equal(normalizeWhatsApp("(11) 99999-9999"), "5511999999999");
  assert.equal(normalizeWhatsApp("(11) 3333-4444"), "551133334444");
  assert.equal(normalizeWhatsApp("+55 11 99999-9999"), "5511999999999");
  assert.equal(normalizeWhatsApp("+44 20 7946 0958"), "442079460958");
  for (const phone of ["", "telefone-invalido", "123", "11111111111", "1234567890123456", "11+999999999"]) {
    assert.throws(() => normalizeWhatsApp(phone), /número válido/);
  }
});

test("Alerta detecta falhas HTTP e rejeições do serviço de e-mail", async () => {
  const quote = calculateQuote(base, prices);
  await assert.rejects(() => sendCalculationAlert(quote, "https://example.test/alerts", "https://example.test/", async () => ({ ok: false })), /failed/);
  await assert.rejects(() => sendCalculationAlert(quote, "https://example.test/alerts", "https://example.test/", async () => ({ ok: true, json: async () => ({ success: false }) })), /rejected/);
});
