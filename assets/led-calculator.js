(() => {
  const NATIONALIZATION_FACTOR = 12;
  const FAMILY_LABELS = {
    indoor_fixed_front: "Fixo indoor · manutenção frontal",
    outdoor_fixed_front: "Fixo outdoor · manutenção frontal",
    rental_indoor: "Rental indoor",
    rental_outdoor: "Rental outdoor",
  };
  const PURPOSE_LABELS = { use: "Uso próprio", resale: "Revenda", rental: "Locação" };

  function positiveDecimal(value) {
    const text = String(value ?? "").trim();
    if (!/^\d+(?:[.,]\d+)?$/.test(text)) throw new Error("Informe um número maior que zero.");
    const number = Number(text.replace(",", "."));
    if (!Number.isFinite(number) || number <= 0) throw new Error("Informe um número maior que zero.");
    return number;
  }

  function priceFamily(input) {
    if (!["indoor", "outdoor"].includes(input.application) || !["yes", "no"].includes(input.wall) || !Object.hasOwn(PURPOSE_LABELS, input.purpose)) {
      throw new Error("Selecione a aplicação, a instalação e a finalidade do painel.");
    }
    const rental = input.wall === "no" || input.purpose === "rental";
    return rental ? `rental_${input.application}` : `${input.application}_fixed_front`;
  }

  function priceRows(input, prices) {
    const family = priceFamily(input);
    const familyRows = prices?.[family];
    return (Array.isArray(familyRows) ? familyRows : []).filter((row) => row && Number.isFinite(row.pitch) && row.pitch > 0 && Number.isFinite(row.usd_per_m2) && row.usd_per_m2 > 0);
  }

  function availablePitches(input, prices) {
    return [...new Set(priceRows(input, prices).map((row) => row.pitch))].sort((a, b) => a - b);
  }

  function calculateQuote(input, prices) {
    const family = priceFamily(input);
    const pitch = positiveDecimal(input.pitch);
    const area = positiveDecimal(input.area);
    // Price the exact user-selected pitch; never substitute another model.
    const selected = priceRows(input, prices).filter((row) => row.pitch === pitch).sort((a, b) => a.usd_per_m2 - b.usd_per_m2)[0];
    if (!selected) throw new Error("Não há preço disponível para essa configuração. Consulte nosso time.");
    const originTotal = selected.usd_per_m2 * area;
    const nationalizedTotal = originTotal * NATIONALIZATION_FACTOR;
    if (!Number.isFinite(nationalizedTotal)) throw new Error("Confira a área informada e tente novamente.");
    return { ...input, area, family, pitch: selected.pitch, lamp: selected.lamp || "Não informado", usdPerM2: selected.usd_per_m2,
      originTotal, nationalizedPerM2: selected.usd_per_m2 * NATIONALIZATION_FACTOR, nationalizedTotal };
  }

  // Allow the same price rules to be verified with Node, without DOM or network calls.
  if (typeof module !== "undefined" && module.exports) module.exports = { calculateQuote, positiveDecimal, availablePitches };
  if (typeof document === "undefined") return;
  const form = document.querySelector("#led-calculator");
  if (!form) return;

  const button = form.querySelector('button[type="submit"]');
  const status = document.querySelector("#calculator-status");
  const values = document.querySelector("#calculator-values");
  const empty = document.querySelector("#calculator-empty");
  const pitchSelect = document.querySelector("#calculator-pitch");
  const retryPrices = document.querySelector("#calculator-retry-prices");
  const formatNumber = (value) => new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 3 }).format(value);
  const usd = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "USD" });
  const brl = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });
  const text = (id, value) => { document.getElementById(id).textContent = value; };
  let pricesPromise;
  let busy = false;
  let revision = 0;
  let optionsRevision = 0;

  function clearResult() {
    revision += 1;
    values.hidden = true;
    empty.hidden = false;
    status.textContent = "";
  }
  form.addEventListener("input", clearResult);
  form.addEventListener("change", clearResult);

  async function loadPrices() {
    // Classic scripts work with file:// too; fetch of local JSON does not.
    if (globalThis.ImporteLEDPriceTable) return globalThis.ImporteLEDPriceTable;
    if (!pricesPromise) {
      pricesPromise = (async () => {
        const controller = new AbortController();
        const timeout = setTimeout(() => controller.abort(), 12000);
        try {
          const response = await fetch("LED/data/prices_cache.json", { signal: controller.signal });
          if (!response.ok) throw new Error("Price table unavailable");
          return await response.json();
        } finally {
          clearTimeout(timeout);
        }
      })().catch(() => {
          pricesPromise = undefined;
          throw new Error("Não foi possível carregar a tabela de preços. Abra esta página pelo site e tente novamente.");
        });
    }
    return pricesPromise;
  }

  async function refreshPitches() {
    const request = ++optionsRevision;
    const previousPitch = pitchSelect.value;
    const input = Object.fromEntries(new FormData(form));
    pitchSelect.disabled = true;
    pitchSelect.replaceChildren(new Option("Carregando opções…", ""));
    retryPrices.hidden = true;
    try {
      const pitches = availablePitches(input, await loadPrices());
      if (request !== optionsRevision) return;
      if (!pitches.length) throw new Error("Não há modelos disponíveis para esta configuração.");
      pitchSelect.replaceChildren(new Option("Selecione o pixel pitch", ""), ...pitches.map((pitch) => new Option(`P${formatNumber(pitch)}`, String(pitch))));
      pitchSelect.value = pitches.some((pitch) => String(pitch) === previousPitch) ? previousPitch : "";
      pitchSelect.disabled = false;
      status.textContent = "";
    } catch (error) {
      if (request !== optionsRevision) return;
      pitchSelect.replaceChildren(new Option("Opções indisponíveis", ""));
      retryPrices.hidden = false;
      status.textContent = error.message;
    }
  }
  form.addEventListener("change", (event) => {
    if (["application", "wall", "purpose"].includes(event.target.name)) refreshPitches();
  });
  retryPrices.addEventListener("click", refreshPitches);
  refreshPitches();

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    if (pitchSelect.disabled) {
      status.textContent = "Carregue as opções de pixel pitch antes de calcular.";
      return;
    }
    if (busy || !form.reportValidity()) return;
    const input = Object.fromEntries(new FormData(form));
    const currentRevision = revision;
    busy = true;
    button.disabled = true;
    button.textContent = "Calculando…";
    form.setAttribute("aria-busy", "true");
    status.textContent = "Consultando a tabela de preços…";
    try {
      // Validate numerical inputs before loading the table.
      positiveDecimal(input.pitch);
      positiveDecimal(input.area);
      const quote = calculateQuote(input, await loadPrices());
      if (revision !== currentRevision) return;
      text("calculator-origin-total", usd.format(quote.originTotal));
      text("calculator-origin-unit", `${usd.format(quote.usdPerM2)} por m²`);
      text("calculator-nationalized-total", brl.format(quote.nationalizedTotal));
      text("calculator-nationalized-unit", `${brl.format(quote.nationalizedPerM2)} por m²`);
      text("calculator-summary-area", `${formatNumber(quote.area)} m²`);
      text("calculator-summary-family", FAMILY_LABELS[quote.family]);
      text("calculator-summary-pitch", `P${formatNumber(quote.pitch)}`);
      text("calculator-summary-lamp", quote.lamp);
      text("calculator-summary-purpose", PURPOSE_LABELS[quote.purpose]);
      const message = `Olá, autorizo a ImporteLED a me atender sobre esta estimativa de painel de LED: ${FAMILY_LABELS[quote.family]}, ${formatNumber(quote.area)} m², P${formatNumber(quote.pitch)}, finalidade: ${PURPOSE_LABELS[quote.purpose]}, fixação na parede: ${quote.wall === "yes" ? "sim" : "não"}. Origem: ${usd.format(quote.originTotal)}. Nacionalizado estimado: ${brl.format(quote.nationalizedTotal)} (fator 12). Se eu pedir PARAR, não quero receber novos contatos.`;
      document.querySelector("#calculator-consultant").href = `https://wa.me/5551980295513?text=${encodeURIComponent(message)}`;
      empty.hidden = true;
      values.hidden = false;
      status.textContent = `Estimativa calculada para ${formatNumber(quote.area)} m². Na origem: ${usd.format(quote.originTotal)}. Nacionalizado: ${brl.format(quote.nationalizedTotal)}.`;
    } catch (error) {
      if (revision !== currentRevision) return;
      values.hidden = true;
      empty.hidden = false;
      status.textContent = error.message;
    } finally {
      busy = false;
      button.disabled = false;
      button.textContent = "Calcular preço";
      form.removeAttribute("aria-busy");
    }
  });
})();
