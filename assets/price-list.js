(() => {
  const form = document.querySelector("#price-list-form");
  if (!form) return;

  const button = form.querySelector('button[type="submit"]');
  const status = document.querySelector("#price-list-status");
  const downloadLink = document.querySelector("#price-list-download");
  const listUrl = form.dataset.priceListUrl.trim();
  let downloadUrl;
  let submitting = false;

  // Leave submissions disabled until the approved price-list file is configured.
  if (listUrl && form.dataset.leadEndpoint) {
    button.disabled = false;
    status.textContent = "";
  }

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    if (!listUrl || submitting || !form.reportValidity()) return;

    const values = new FormData(form);
    for (const [key, value] of values) values.set(key, value.trim());
    if (["nome", "email", "telefone", "cidade"].some((key) => !values.get(key))) {
      status.textContent = "Preencha todos os campos para continuar.";
      return;
    }
    const phoneDigits = values.get("telefone").replace(/\D/g, "");
    if (phoneDigits.length < 10 || phoneDigits.length > 15) {
      status.textContent = "Informe um telefone válido com DDD.";
      document.querySelector("#price-list-phone").focus();
      return;
    }

    submitting = true;
    button.disabled = true;
    button.textContent = "Preparando seu download…";
    form.setAttribute("aria-busy", "true");
    status.textContent = "Aguarde enquanto preparamos a lista de preços.";
    downloadLink.hidden = true;
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 25000);

    try {
      // Check the file before sending personal data or reporting a successful download.
      const fileResponse = await fetch(listUrl, { signal: controller.signal });
      if (!fileResponse.ok || fileResponse.headers.get("content-type")?.includes("text/html")) {
        throw new Error("Price list unavailable");
      }
      const file = await fileResponse.blob();
      if (!file.size) throw new Error("Empty price list");

      const response = await fetch(form.dataset.leadEndpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json", Accept: "application/json" },
        signal: controller.signal,
        body: JSON.stringify({
          ...Object.fromEntries(values),
          _subject: "ImporteLED — solicitação da lista de preços",
          _template: "table",
          origem: "Formulário de download da lista de preços",
        }),
      });
      if (!response.ok) throw new Error("Submission failed");
      const result = await response.json();
      if (result.success !== true && result.success !== "true") throw new Error("Submission rejected");

      if (downloadUrl) URL.revokeObjectURL(downloadUrl);
      downloadUrl = URL.createObjectURL(file);
      downloadLink.href = downloadUrl;
      downloadLink.download = new URL(listUrl, window.location.href).pathname.split("/").pop() || "lista-de-precos";
      downloadLink.hidden = false;
      downloadLink.click();
      status.textContent = "Seu download foi iniciado. Se precisar, use o link abaixo para baixar novamente.";
    } catch {
      status.textContent = "Não foi possível concluir. Tente novamente ou fale com nosso consultor pelo link no topo da página.";
    } finally {
      clearTimeout(timeout);
      submitting = false;
      button.disabled = false;
      button.textContent = "Baixar lista de preços";
      form.removeAttribute("aria-busy");
    }
  });
})();
