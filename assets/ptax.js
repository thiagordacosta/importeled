(() => {
  // Fixed commercial exchange rate defined by ImporteLED; not a BCB PTAX quote.
  const REFERENCE_PTAX = {
    rate: 5.5,
    date: "2026-10-19",
    source: "Cotação fixa definida pela ImporteLED",
  };

  async function loadPtax() {
    return { ...REFERENCE_PTAX };
  }

  const api = { loadPtax };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else globalThis.ImporteLEDPtax = api;
})();
