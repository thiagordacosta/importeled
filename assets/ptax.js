(() => {
  // BCB SGS series 1 is the daily USD selling PTAX closing rate.
  // This official endpoint permits browser requests (Access-Control-Allow-Origin: *).
  const ENDPOINT = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.1/dados/ultimos/1?formato=json";

  function brazilDate(now) {
    const parts = new Intl.DateTimeFormat("en-US", { timeZone: "America/Sao_Paulo", year: "numeric", month: "2-digit", day: "2-digit" }).formatToParts(now);
    const part = (type) => parts.find((item) => item.type === type).value;
    return `${part("year")}-${part("month")}-${part("day")}`;
  }

  function parsePtax(payload, now = new Date()) {
    const today = brazilDate(now);
    const oldest = new Date(`${today}T12:00:00Z`);
    oldest.setUTCDate(oldest.getUTCDate() - 14);
    const oldestDate = oldest.toISOString().slice(0, 10);
    const rows = Array.isArray(payload) ? payload : [];
    const latest = rows.filter((row) => /^\d{2}\/\d{2}\/\d{4}$/.test(row?.data))
      .map((row) => ({ rate: Number(row.valor), date: row.data.split("/").reverse().join("-") }))
      .filter((row) => Number.isFinite(row.rate) && row.rate > 0 && row.date <= today && row.date >= oldestDate
        && !Number.isNaN(Date.parse(`${row.date}T12:00:00Z`)) && new Date(`${row.date}T12:00:00Z`).toISOString().slice(0, 10) === row.date)
      .sort((a, b) => b.date.localeCompare(a.date))[0];
    if (!latest) throw new Error("PTAX de fechamento indisponível.");
    return { ...latest, source: ENDPOINT };
  }

  async function loadPtax(request = fetch, now = new Date()) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 12000);
    try {
      const response = await request(ENDPOINT, { signal: controller.signal, cache: "no-store" });
      if (!response.ok) throw new Error("BCB request failed");
      return parsePtax(await response.json(), now);
    } catch {
      throw new Error("Não foi possível consultar a PTAX do Banco Central. Tente calcular novamente em alguns instantes.");
    } finally {
      clearTimeout(timeout);
    }
  }

  const api = { parsePtax, loadPtax };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else globalThis.ImporteLEDPtax = api;
})();
