const canvas = document.querySelector("#led-canvas");
const ctx = canvas.getContext("2d");
const chatBody = document.querySelector("#chat-body");
const chatFooter = document.querySelector(".chat-footer");
const quickReplies = document.querySelector("#quick-replies");

const SHEET_ID = "1mHNTXO5nT57HZED-LdrfcfaWE1i0heuDNB4qofOMMxk";
const SHEET_GID = new URLSearchParams(window.location.search).get("gid") || "0";
const SHEET_JSONP_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?gid=${SHEET_GID}`;
const PRICE_CACHE_URL = "data/prices_cache.json";
const ZIP_LOOKUP_URL = "https://viacep.com.br/ws";
const SCHEDULE_EMAIL_ENDPOINT = "https://formsubmit.co/thiago@caring.com.br";
const LEAD_POLL_INTERVAL_MS = 30000;

const PRICE_TARGETS = {
  indoor: { near: 1.86, mid: 2.5, far: 3.91 },
  outdoor: { near: 2.976, mid: 3.91, far: 4.81 },
  rental: { near: 2.604, mid: 2.976, far: 3.91 },
};

const NATIONALIZATION_RATE = 11;
const QUALIFIED_AREA_M2 = 8;

const flow = [
  {
    key: "name",
    question:
      "Oi. Você chegou pelo anúncio para acessar o preço do painel nacionalizado e falar sobre instalação. Confirma seu nome para eu abrir a prévia?",
    placeholder: "Digite seu nome",
    type: "text",
  },
  {
    key: "application",
    question: "Perfeito, {name}. Para estimar o produto nacionalizado, onde esse painel vai ser usado?",
    type: "choice",
    options: [
      { label: "Ambiente interno", value: "indoor" },
      { label: "Área externa", value: "outdoor" },
      { label: "Eventos / rental", value: "rental" },
    ],
  },
  {
    key: "distance",
    question: "Qual é a distância média de visualização? Isso define o pixel pitch e muda bastante o preço final.",
    type: "choice",
    options: [
      { label: "Até 3 m", value: "near" },
      { label: "3 a 6 m", value: "mid" },
      { label: "Acima de 6 m", value: "far" },
    ],
  },
  {
    key: "area",
    question: "Qual é a área aproximada do painel em metros quadrados? Pode responder em m² ou como largura x altura.",
    placeholder: "Ex: 12 ou 3x4",
    type: "text",
  },
  {
    key: "install",
    question: "Sobre instalação: qual cenário parece mais próximo do seu projeto?",
    type: "choice",
    options: [
      { label: "Fixa frontal", value: "front" },
      { label: "Fixa traseira", value: "rear" },
      { label: "Móvel / rental", value: "mobile" },
    ],
  },
  {
    key: "phone",
    question: "Para um agente continuar a conversa sobre instalação e visita técnica, qual WhatsApp devemos usar?",
    placeholder: "(11) 99999-9999",
    type: "text",
  },
  {
    key: "city",
    question: "Em qual cidade será feita a instalação?",
    placeholder: "Ex: São Paulo, SP",
    type: "text",
  },
  {
    key: "visitPeriod",
    question: "Qual período é melhor para uma visita técnica?",
    type: "choice",
    options: [
      { label: "Manhã", value: "manhã" },
      { label: "Tarde", value: "tarde" },
      { label: "A combinar", value: "a combinar" },
    ],
  },
];

let width = 0;
let height = 0;
let pixelRatio = 1;
let time = 0;
let stepIndex = 0;
let answers = {};
let sheetRows = [];
let historyLoaded = false;
let priceCache = {};
let pricesLoaded = false;
let latestLeadSignature = "";
let pollingLead = false;
let autoScrollMessages = false;
let pendingLeadConfirmation = false;

function resizeCanvas() {
  pixelRatio = Math.min(window.devicePixelRatio || 1, 2);
  width = canvas.clientWidth;
  height = canvas.clientHeight;
  canvas.width = Math.floor(width * pixelRatio);
  canvas.height = Math.floor(height * pixelRatio);
  ctx.setTransform(pixelRatio, 0, 0, pixelRatio, 0, 0);
}

function drawLedWall() {
  time += 0.012;
  ctx.clearRect(0, 0, width, height);
  ctx.fillStyle = "#070b12";
  ctx.fillRect(0, 0, width, height);

  const gap = width < 640 ? 18 : 23;
  const dot = width < 640 ? 2.1 : 2.8;
  const startX = width < 980 ? width * 0.1 : width * 0.38;
  const colors = [
    [25, 179, 106],
    [49, 193, 214],
    [244, 185, 66],
    [223, 75, 63],
  ];

  for (let y = -gap; y < height + gap; y += gap) {
    for (let x = startX; x < width + gap; x += gap) {
      const wave = Math.sin(x * 0.011 + time * 3) + Math.cos(y * 0.018 - time * 2.1);
      const glow = Math.max(0.1, Math.min(1, 0.32 + wave * 0.24 + (x / width) * 0.28));
      const color = colors[Math.abs(Math.floor((x + y) / gap)) % colors.length];
      ctx.beginPath();
      ctx.fillStyle = `rgba(${color[0]}, ${color[1]}, ${color[2]}, ${glow})`;
      ctx.shadowColor = `rgba(${color[0]}, ${color[1]}, ${color[2]}, ${glow})`;
      ctx.shadowBlur = 11 * glow;
      ctx.arc(x, y, dot + glow * 1.2, 0, Math.PI * 2);
      ctx.fill();
    }
  }

  ctx.shadowBlur = 0;
  requestAnimationFrame(drawLedWall);
}

function currency(value) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 0,
  }).format(value);
}

function brlCurrency(value) {
  return new Intl.NumberFormat("pt-BR", {
    style: "currency",
    currency: "BRL",
    maximumFractionDigits: 0,
  }).format(value);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function normalizeKey(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function getField(row, candidates) {
  const normalized = Object.entries(row).map(([key, value]) => [normalizeKey(key), value]);
  for (const candidate of candidates) {
    const wanted = normalizeKey(candidate);
    const exact = normalized.find(([key]) => key === wanted);
    if (exact && String(exact[1]).trim()) {
      return String(exact[1]).trim();
    }
  }
  for (const candidate of candidates) {
    const wanted = normalizeKey(candidate);
    const fuzzy = normalized.find(([key, value]) => key.includes(wanted) && String(value).trim());
    if (fuzzy) {
      return String(fuzzy[1]).trim();
    }
  }
  return "";
}

function columnName(index) {
  let name = "";
  let cursor = index + 1;
  while (cursor > 0) {
    const remainder = (cursor - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    cursor = Math.floor((cursor - 1) / 26);
  }
  return name;
}

function getColumn(row, letter) {
  const wanted = String(letter || "").toUpperCase();
  const fallback = getField(row, [`col_${wanted.charCodeAt(0) - 64}`]);
  return String(row.__cellsByLetter?.[wanted] || fallback || "").trim();
}

function compactPhone(value) {
  return String(value || "").replace(/\D/g, "");
}

function compactZipCode(value) {
  return String(value || "").replace(/\D/g, "").slice(0, 8);
}

function formatZipCode(value) {
  const digits = compactZipCode(value);
  if (digits.length !== 8) {
    return value || "";
  }
  return `${digits.slice(0, 5)}-${digits.slice(5)}`;
}

function rowHasAnyValue(row) {
  return Object.entries(row).some(([key, value]) => key !== "__cellsByLetter" && String(value || "").trim());
}

function rowSignature(row) {
  return Object.entries(row)
    .filter(([key]) => key !== "__cellsByLetter")
    .map(([key, value]) => `${key}:${String(value || "").trim()}`)
    .join("|");
}

function formatHistoryLine(row, index) {
  const date = getField(row, ["created_time", "data", "timestamp", "date", "quando"]);
  const sender = getField(row, ["sender", "remetente", "origem", "canal", "author"]) || "Histórico";
  const message =
    getField(row, ["message", "mensagem", "historico", "histórico", "communication", "comunicacao", "observacao"]) ||
    getField(row, ["lead_status", "status"]) ||
    "Registro encontrado na planilha.";
  const prefix = date ? `${date} · ${sender}` : `${index + 1}. ${sender}`;
  return `${prefix}: ${message}`;
}

function inferApplication(value) {
  const text = normalizeKey(value);
  if (text.includes("outdoor") || text.includes("extern")) {
    return "outdoor";
  }
  if (text.includes("indoor") || text.includes("intern")) {
    return "indoor";
  }
  if (text.includes("rental") || text.includes("evento") || text.includes("movel") || text.includes("aluguel")) {
    return "rental";
  }
  if (text.includes("fixo") || text.includes("fixa") || text.includes("vitrine") || text.includes("loja")) {
    return "indoor";
  }
  return "";
}

function inferDistance(value) {
  const text = normalizeKey(value);
  const numbers = String(value || "")
    .replace(",", ".")
    .match(/\d+(?:\.\d+)?/g)
    ?.map(Number)
    .filter(Number.isFinite);

  if (text.includes("p1") || text.includes("ate_3") || text.includes("menor_3") || text.includes("proximo")) {
    return "near";
  }
  if (text.includes("p2") || text.includes("3_a_6") || text.includes("3_6") || text.includes("medio")) {
    return "mid";
  }
  if (text.includes("p3") || text.includes("p4") || text.includes("acima") || text.includes("maior_6")) {
    return "far";
  }
  if (numbers?.length) {
    const distance = Math.max(...numbers);
    if (distance <= 3) {
      return "near";
    }
    if (distance <= 6) {
      return "mid";
    }
    return "far";
  }
  return "";
}

function inferInstall(value) {
  const text = normalizeKey(value);
  if (text.includes("traseiro") || text.includes("rear")) {
    return "rear";
  }
  if (text.includes("movel") || text.includes("rental") || text.includes("aluguel") || text.includes("evento")) {
    return "mobile";
  }
  if (text.includes("sim") || text.includes("fixa") || text.includes("fixo") || text.includes("frontal") || text.includes("front")) {
    return "front";
  }
  return "";
}

function readAgentColumns(row) {
  const applicationText = getColumn(row, "N");
  const indoorDistanceText = getColumn(row, "O");
  const outdoorDistanceText = getColumn(row, "P");
  const installText = getColumn(row, "Q");
  const areaText = getColumn(row, "S");

  const application = inferApplication(applicationText);
  const distance = inferDistance(application === "outdoor" ? outdoorDistanceText || indoorDistanceText : indoorDistanceText || outdoorDistanceText);
  const install = inferInstall(installText || applicationText);
  const area = parseArea(areaText);

  return {
    applicationText,
    indoorDistanceText,
    outdoorDistanceText,
    installText,
    areaText,
    application,
    distance,
    install,
    area,
    qualified: area >= QUALIFIED_AREA_M2,
  };
}

function applyAgentAnswers(agent) {
  if (agent.areaText) {
    answers.area = agent.areaText;
  }
  if (agent.application) {
    answers.application = agent.application;
  }
  if (agent.distance) {
    answers.distance = agent.distance;
  }
  if (agent.install) {
    answers.install = agent.install;
  }
  answers.qualifiedOpportunity = agent.qualified;
}

function hydrateAnswersFromLatestRow(row) {
  const agent = readAgentColumns(row);
  const zipCodeFromColumn = getColumn(row, "Y");
  const fieldMap = {
    name: ["full name", "nome", "name"],
    email: ["email", "e-mail", "mail", "endereco_de_email", "endereço_de_email"],
    phone: ["phone", "telefone", "whatsapp", "phone_number"],
    city: ["cidade", "city", "local", "localização", "endereco", "endereço"],
    zipCode: ["cep", "zip", "zipcode", "zip_code", "codigo_postal", "código_postal", "postal_code"],
  };

  Object.entries(fieldMap).forEach(([key, candidates]) => {
    const value = getField(row, candidates);
    if (value && !answers[key]) {
      answers[key] = value;
    }
  });

  if (zipCodeFromColumn) {
    answers.zipCode = zipCodeFromColumn;
  }

  applyAgentAnswers(agent);

  if (!answers.area) {
    const area = getField(row, ["qual_o_tamanho_do_painel_em_metros_quadrados?", "tamanho_painel", "metros_quadrados", "area"]);
    if (area) {
      answers.area = area;
    }
  }

  const application =
    agent.application || inferApplication(getField(row, ["qual_a_aplicação_do_painel_de_led?", "aplicação", "application"]));
  if (application && !answers.application) {
    answers.application = application;
  }

  const distance =
    agent.distance ||
    inferDistance(getField(row, ["qual_a_distância_de_visualização?_in", "distância_visualização_in"])) ||
    inferDistance(getField(row, ["qual_a_distância_de_visualização?_out", "distância_visualização_out"]));
  if (distance && !answers.distance) {
    answers.distance = distance;
  }

  const install = agent.install || inferInstall(getField(row, ["é_instalação_fixa?", "instalação_fixa", "install"]));
  if (install && !answers.install) {
    answers.install = install;
  }

  answers.qualifiedOpportunity = agent.qualified;
}

function formatLoadedContext(row) {
  const parts = [];
  const agent = readAgentColumns(row);
  const name = getField(row, ["full name", "nome", "name"]);
  const application = agent.applicationText || getField(row, ["qual_a_aplicação_do_painel_de_led?", "aplicação", "application"]);
  const area = agent.areaText || getField(row, [
    "qual_o_tamanho_do_painel_em_metros_quadrados?",
    "tamanho_painel",
    "metros_quadrados",
    "area",
  ]);
  const phone = getField(row, ["phone", "telefone", "whatsapp", "phone_number"]);
  const city = getField(row, ["cidade", "city", "local", "localização", "endereco", "endereço"]);
  const zipCode =
    getColumn(row, "Y") || getField(row, ["cep", "zip", "zipcode", "zip_code", "codigo_postal", "código_postal", "postal_code"]);

  if (name) {
    parts.push(`nome: ${name}`);
  }
  if (application) {
    parts.push(`uso: ${application}`);
  }
  if (agent.distance) {
    parts.push(`distância: ${labelFor("distance", agent.distance)}`);
  }
  if (area) {
    parts.push(`área: ${area}`);
  }
  if (agent.install) {
    parts.push(`instalação: ${labelFor("install", agent.install)}`);
  }
  if (agent.area) {
    parts.push(agent.qualified ? "oportunidade >= 8 m²" : "oportunidade abaixo de 8 m²");
  }
  if (phone) {
    parts.push(`WhatsApp: ${phone}`);
  }
  if (city) {
    parts.push(`cidade: ${city}`);
  }
  if (zipCode) {
    parts.push(`CEP: ${zipCode}`);
  }

  return parts.length ? parts.join(" · ") : formatHistoryLine(row, 0);
}

function formatAddress(address) {
  if (!address) {
    return "";
  }
  return [address.street, address.neighborhood, address.city, address.state].filter(Boolean).join(" · ");
}

async function lookupZipCode(zipCode) {
  const digits = compactZipCode(zipCode);
  if (digits.length !== 8) {
    return null;
  }

  let response;
  try {
    response = await fetch(`${ZIP_LOOKUP_URL}/${digits}/json/`);
  } catch (error) {
    return null;
  }

  if (!response.ok) {
    return null;
  }

  let data;
  try {
    data = await response.json();
  } catch (error) {
    return null;
  }

  if (data.erro) {
    return null;
  }

  return {
    zipCode: formatZipCode(data.cep || digits),
    street: data.logradouro || "",
    neighborhood: data.bairro || "",
    city: data.localidade || "",
    state: data.uf || "",
  };
}

async function hydrateAddressFromZip() {
  if (!answers.zipCode || answers.address) {
    return answers.address || null;
  }

  const address = await lookupZipCode(answers.zipCode);
  if (address) {
    answers.zipCode = address.zipCode;
    answers.address = address;
    if (!answers.city && address.city) {
      answers.city = `${address.city}, ${address.state}`;
    }
  }
  return address;
}

function getLatestRow(rows) {
  return [...rows].reverse().find(rowHasAnyValue) || null;
}

async function loadSheetHistory(force = false) {
  if (historyLoaded && !force) {
    return sheetRows;
  }
  historyLoaded = true;
  sheetRows = await loadGoogleSheetRows();
  return sheetRows;
}

async function loadPriceCache() {
  if (pricesLoaded) {
    return priceCache;
  }
  pricesLoaded = true;
  const response = await fetch(`${PRICE_CACHE_URL}?v=${Date.now()}`, { cache: "no-store" });
  if (!response.ok) {
    pricesLoaded = false;
    throw new Error("Cache de preços indisponível");
  }
  try {
    priceCache = await response.json();
  } catch (error) {
    pricesLoaded = false;
    throw error;
  }
  return priceCache;
}

function sheetCellValue(cell) {
  if (!cell) {
    return "";
  }
  return String(cell.f ?? cell.v ?? "").trim();
}

function tableToRows(table) {
  const headers = table.cols.map((col, index) => String(col.label || `col_${index + 1}`).trim());
  return table.rows.map((row) => {
    const record = headers.reduce((acc, header, index) => {
      acc[header] = sheetCellValue(row.c[index]);
      return acc;
    }, {});

    record.__cellsByLetter = headers.reduce((acc, header, index) => {
      acc[columnName(index)] = sheetCellValue(row.c[index]);
      acc[header] = sheetCellValue(row.c[index]);
      return acc;
    }, {});

    return record;
  });
}

function pickPriceFamily(application, install) {
  if (install === "mobile" || application === "rental") {
    return application === "outdoor" ? "rental_outdoor" : "rental_indoor";
  }
  if (application === "outdoor") {
    return install === "rear" ? "outdoor_fixed_rear" : "outdoor_fixed_front";
  }
  return "indoor_fixed_front";
}

function formatPitch(value) {
  return `P${Number(value).toLocaleString("pt-BR", { maximumFractionDigits: 3 })}`;
}

function selectPriceItem(application, distance, install) {
  const family = pickPriceFamily(application, install);
  const rows = priceCache[family] || [];
  const target = PRICE_TARGETS[application]?.[distance] || PRICE_TARGETS.indoor.mid;

  if (!rows.length) {
    throw new Error(`Família de preço não encontrada: ${family}`);
  }

  const closestPitch = rows
    .map((item) => Number(item.pitch))
    .filter(Number.isFinite)
    .sort((a, b) => Math.abs(a - target) - Math.abs(b - target))[0];

  const candidates = rows
    .filter((item) => Number(item.pitch) === closestPitch)
    .sort((a, b) => Number(a.usd_per_m2) - Number(b.usd_per_m2));

  return { ...candidates[0], family };
}

function parseArea(value) {
  const text = String(value).toLowerCase().replace(",", ".").trim();
  const dimensions = text.match(/(\d+(?:\.\d+)?)\s*(x|por)\s*(\d+(?:\.\d+)?)/);
  if (dimensions) {
    return Number(dimensions[1]) * Number(dimensions[3]);
  }
  const number = text.match(/\d+(?:\.\d+)?/);
  return number ? Number(number[0]) : 0;
}

function getQuote() {
  const application = answers.application || "indoor";
  const distance = answers.distance || "mid";
  const install = answers.install || "front";
  const area = parseArea(answers.area);
  const selected = selectPriceItem(application, distance, install);
  const usdPerM2 = Number(selected.usd_per_m2) || 0;
  const usdTotal = usdPerM2 * Math.max(area, 0);

  return {
    application,
    distance,
    install,
    area,
    family: selected.family,
    pitch: formatPitch(selected.pitch),
    lamp: selected.lamp,
    cabinet: selected.cabinet || "",
    usdPerM2,
    usdTotal,
    nationalizedTotal: usdTotal * NATIONALIZATION_RATE,
    qualifiedOpportunity: area >= QUALIFIED_AREA_M2,
  };
}

function getLeadReadiness() {
  const required = ["application", "distance", "area", "install"];
  return required.every((key) => Boolean(answers[key]));
}

function loadGoogleSheetRows() {
  return new Promise((resolve, reject) => {
    const callbackName = `importeledSheet_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const script = document.createElement("script");
    const timeout = window.setTimeout(() => {
      cleanup();
      reject(new Error("Tempo esgotado ao carregar a planilha"));
    }, 12000);

    function cleanup() {
      window.clearTimeout(timeout);
      delete window[callbackName];
      script.remove();
    }

    window[callbackName] = (payload) => {
      cleanup();
      if (!payload?.table?.cols?.length) {
        reject(new Error("Planilha sem dados públicos"));
        return;
      }
      resolve(tableToRows(payload.table));
    };

    script.onerror = () => {
      cleanup();
      reject(new Error("Google bloqueou o carregamento da planilha"));
    };

    const params = new URLSearchParams({
      tqx: `responseHandler:${callbackName}`,
      _: String(Date.now()),
    });
    script.src = `${SHEET_JSONP_URL}&${params.toString()}`;
    document.head.append(script);
  });
}

async function showInitialHistory() {
  try {
    const rows = await loadSheetHistory();
    const latestRow = getLatestRow(rows);

    if (!latestRow) {
      addMessage("Não encontrei uma última linha preenchida. Vou começar uma nova conversa.");
      return;
    }

    hydrateAnswersFromLatestRow(latestRow);
    latestLeadSignature = rowSignature(latestRow);

    if (getLeadReadiness()) {
      askLeadConfirmation();
      return true;
    }

    addMessage("Recebi seus dados e vou completar as informações que faltam para calcular o valor dos painéis em reais.");
    return false;
  } catch (error) {
    addMessage(
      "Não consegui carregar a planilha agora. Se ela não estiver pública como visualização, o Google bloqueia o acesso pelo site."
    );
    return false;
  }
}

function askLeadConfirmation() {
  pendingLeadConfirmation = true;
  quickReplies.innerHTML = "";
  const lastName = getLeadLastName(answers.name);
  addMessage("Para confirmar sua identidade, selecione o seu sobrenome.");

  getLastNameOptions(lastName).forEach((option) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "reply-button";
    button.textContent = option;
    button.addEventListener("click", () => confirmLeadIdentity(option === lastName, option));
    quickReplies.append(button);
  });
}

function getLeadLastName(name) {
  const parts = String(name || "")
    .trim()
    .split(/\s+/)
    .filter(Boolean);
  return parts.length > 1 ? parts.at(-1) : parts[0] || "Cliente";
}

function getLastNameOptions(lastName) {
  const decoys = ["Silva", "Santos", "Oliveira", "Souza", "Costa", "Pereira", "Almeida", "Ferreira", "Rodrigues", "Lima"];
  const options = [lastName];
  decoys.forEach((item) => {
    if (normalizeKey(item) !== normalizeKey(lastName) && options.length < 4) {
      options.push(item);
    }
  });

  const seed = normalizeKey(answers.name || lastName).length % options.length;
  return [...options.slice(seed), ...options.slice(0, seed)];
}

function confirmLeadIdentity(isConfirmed, selectedLastName) {
  pendingLeadConfirmation = false;
  quickReplies.innerHTML = "";
  addMessage(selectedLastName, "user");

  if (isConfirmed) {
    if (!getQuote().qualifiedOpportunity) {
      addDisqualifiedSummary();
      return;
    }
    finishChat();
    return;
  }

  addMessage("Não consegui confirmar sua identidade com essa opção. Por segurança, não vou mostrar este orçamento.");
}

function labelFor(key, value) {
  const step = flow.find((item) => item.key === key);
  const option = step?.options?.find((item) => item.value === value);
  return option ? option.label : value;
}

function botText(text) {
  return text.replaceAll("{name}", answers.name || "");
}

function addMessage(text, owner = "bot", extraClass = "") {
  const message = document.createElement("div");
  message.className = `message ${owner} ${extraClass}`.trim();
  message.textContent = text;
  chatBody.append(message);
  if (autoScrollMessages) {
    chatBody.scrollTop = chatBody.scrollHeight;
  }
  return message;
}

function addSummary() {
  const quote = getQuote();
  const prospect = escapeHtml(answers.name || "cliente");
  const location = [answers.city, answers.zipCode].filter(Boolean).join(" · ") || "Local a confirmar";
  const technicalVisit = escapeHtml(`${location} · ${answers.visitPeriod || "A combinar"}`);
  const ledDetails = escapeHtml(`${quote.lamp || "Cache"}${quote.cabinet ? ` · ${quote.cabinet}` : ""}`);
  const areaStatus = quote.qualifiedOpportunity ? "Oportunidade qualificada para instalação" : "Oportunidade abaixo de 8 m²";
  const summary = document.createElement("div");
  summary.className = "message bot summary quote-card";
  summary.innerHTML = `
    <div class="quote-header">
      <span class="quote-kicker">Orçamento preliminar</span>
      <strong>${prospect}, este é o valor estimado do produto nacionalizado.</strong>
    </div>
    <div class="quote-price">
      <span>Produto nacionalizado antes da instalação</span>
      <strong>${brlCurrency(quote.nationalizedTotal)}</strong>
    </div>
    <div class="summary-grid">
      <div><span>Aplicação interpretada</span><strong>${escapeHtml(labelFor("application", quote.application))}</strong></div>
      <div><span>Distância de visualização</span><strong>${escapeHtml(labelFor("distance", quote.distance))}</strong></div>
      <div><span>Pixel pitch definido</span><strong>${escapeHtml(quote.pitch)}</strong></div>
      <div><span>Área do painel</span><strong>${quote.area || 0} m²</strong></div>
      <div><span>Preço por metro quadrado</span><strong>${currency(quote.usdPerM2)} / m²</strong></div>
      <div><span>Produto em dólar</span><strong>${currency(quote.usdTotal)}</strong></div>
      <div><span>LED</span><strong>${ledDetails}</strong></div>
      <div><span>Instalação</span><strong>${escapeHtml(labelFor("install", quote.install))}</strong></div>
      <div><span>Qualificação</span><strong>${areaStatus}</strong></div>
      <div><span>Visita técnica</span><strong>${technicalVisit}</strong></div>
    </div>
    <p class="quote-note">A instalação é avaliada em visita técnica para confirmar estrutura, acesso, fixação, elétrica e acabamento.</p>
  `;
  chatBody.append(summary);
  chatBody.scrollTop = summary.offsetTop - chatBody.offsetTop;
}

function addDisqualifiedSummary() {
  const area = parseArea(answers.area);
  addMessage(
    `Obrigado pelo interesse, ${answers.name || "cliente"}. No momento não conseguimos seguir com oportunidades de painéis abaixo de ${QUALIFIED_AREA_M2} m².`
  );
  addMessage(`A área informada foi de ${area || 0} m². Para avançar com orçamento e visita técnica, o projeto precisa ter pelo menos ${QUALIFIED_AREA_M2} m².`);
}

function nextUnansweredStep() {
  const index = flow.findIndex((step) => !answers[step.key] && !shouldSkipStep(step));
  stepIndex = index === -1 ? flow.length : index;
}

function renderQuickReplies(step) {
  quickReplies.innerHTML = "";
  if (step.type !== "choice") {
    return;
  }
  step.options.forEach((option) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "reply-button";
    button.textContent = option.label;
    button.addEventListener("click", () => handleAnswer(option.value, option.label));
    quickReplies.append(button);
  });
}

function askCurrentStep() {
  while (flow[stepIndex] && shouldSkipStep(flow[stepIndex])) {
    stepIndex += 1;
  }

  const step = flow[stepIndex];
  if (!step) {
    finishChat();
    return;
  }

  window.setTimeout(() => addMessage(botText(step.question)), 240);
  renderQuickReplies(step);
}

function shouldSkipStep(step) {
  return ["city", "visitPeriod"].includes(step.key) && getLeadReadiness();
}

function handleAnswer(value, displayValue = value) {
  const step = flow[stepIndex];
  const cleanValue = String(value).trim();
  if (!cleanValue) {
    return;
  }

  answers[step.key] = cleanValue;
  addMessage(displayValue, "user");
  stepIndex += 1;
  quickReplies.innerHTML = "";
  askCurrentStep();
}

function cleanPhone(phone) {
  const digits = String(phone).replace(/\D/g, "");
  if (digits.startsWith("55")) {
    return digits;
  }
  return `55${digits}`;
}

function whatsappMessage() {
  const quote = getQuote();
  return [
    `Olá, aqui é ${answers.name}.`,
    "Vim do anúncio para acessar o preço nacionalizado do painel e conversar sobre instalação.",
    `Quero cotar um painel ${labelFor("application", quote.application)} ${quote.pitch}.`,
    `Área aproximada: ${quote.area} m².`,
    `Preço base do cache: ${currency(quote.usdPerM2)} por m², ${quote.lamp || "sem marca informada"}.`,
    `Instalação: ${labelFor("install", quote.install)}.`,
    `Oportunidade: ${quote.qualifiedOpportunity ? "maior ou igual a 8 m²" : "abaixo de 8 m²"}.`,
    `Cidade da instalação: ${answers.city}.`,
    `Melhor período para visita técnica: ${answers.visitPeriod}.`,
    `Preço estimado do produto nacionalizado antes da instalação: ${brlCurrency(quote.nationalizedTotal)}.`,
    "Quero agendar uma visita técnica para consultar o preço da instalação.",
  ].join(" ");
}

function submitScheduleEmail() {
  const quote = getQuote();
  const addressLine = [
    formatAddress(answers.address),
    answers.addressNumber,
    answers.addressComplement,
    answers.addressReference ? `Ref: ${answers.addressReference}` : "",
  ]
    .filter(Boolean)
    .join(" · ");
  const payload = {
    _subject: `Nova visita técnica ImporteLED - ${answers.name || "Lead"}`,
    _template: "table",
    _captcha: "false",
    nome: answers.name || "",
    email: answers.email || "",
    whatsapp: answers.phone || "",
    cep: answers.zipCode || "",
    endereco: addressLine,
    data_desejada: answers.visitDate || "",
    periodo: answers.visitPeriod || "",
    aplicacao: labelFor("application", quote.application),
    distancia_visualizacao: labelFor("distance", quote.distance),
    area_m2: quote.area || 0,
    pixel_pitch: quote.pitch,
    led: quote.lamp || "",
    instalacao: labelFor("install", quote.install),
    valor_produto_nacionalizado: brlCurrency(quote.nationalizedTotal),
  };
  const iframeName = `scheduleEmail_${Date.now()}`;
  const iframe = document.createElement("iframe");
  iframe.name = iframeName;
  iframe.hidden = true;
  document.body.append(iframe);

  const form = document.createElement("form");
  form.action = SCHEDULE_EMAIL_ENDPOINT;
  form.method = "POST";
  form.target = iframeName;
  form.hidden = true;

  Object.entries(payload).forEach(([name, value]) => {
    const input = document.createElement("input");
    input.type = "hidden";
    input.name = name;
    input.value = String(value ?? "");
    form.append(input);
  });

  document.body.append(form);
  form.submit();
  window.setTimeout(() => {
    form.remove();
    iframe.remove();
  }, 5000);
}

async function showNativeSchedule() {
  quickReplies.innerHTML = "";
  addMessage("Estou buscando o endereço pelo CEP para agilizar o agendamento.");
  await hydrateAddressFromZip();

  const hasAddress = Boolean(formatAddress(answers.address));
  const zipField = hasAddress
    ? `<input name="zipCode" autocomplete="postal-code" value="${escapeHtml(answers.zipCode || "")}" readonly />`
    : `<input name="zipCode" autocomplete="postal-code" value="${escapeHtml(answers.zipCode || "")}" required />`;
  const address = answers.address || {};

  const container = document.createElement("div");
  container.className = "message bot schedule-card";
  container.innerHTML = `
    <strong>Agendar visita técnica</strong>
    <form class="schedule-form">
      <label>
        <span>Nome</span>
        <input name="name" autocomplete="name" value="${escapeHtml(answers.name || "")}" required />
      </label>
      <label>
        <span>E-mail</span>
        <input name="email" type="email" autocomplete="email" value="${escapeHtml(answers.email || "")}" required />
      </label>
      <label>
        <span>WhatsApp</span>
        <input name="phone" autocomplete="tel" value="${escapeHtml(answers.phone || "")}" required />
      </label>
      <label>
        <span>CEP</span>
        ${zipField}
      </label>
      <label>
        <span>Endereço</span>
        <input name="street" value="${escapeHtml(address.street || "")}" readonly />
      </label>
      <label>
        <span>Bairro</span>
        <input name="neighborhood" value="${escapeHtml(address.neighborhood || "")}" readonly />
      </label>
      <label>
        <span>Cidade</span>
        <input name="addressCity" value="${escapeHtml(address.city || "")}" readonly />
      </label>
      <label>
        <span>UF</span>
        <input name="addressState" value="${escapeHtml(address.state || "")}" readonly />
      </label>
      <label>
        <span>Número</span>
        <input name="addressNumber" autocomplete="address-line2" value="${escapeHtml(answers.addressNumber || "")}" required />
      </label>
      <label>
        <span>Complemento</span>
        <input name="addressComplement" autocomplete="address-line3" value="${escapeHtml(answers.addressComplement || "")}" placeholder="Sala, bloco, referência interna" />
      </label>
      <label>
        <span>Referência para acesso</span>
        <input name="addressReference" value="${escapeHtml(answers.addressReference || "")}" placeholder="Portaria, doca, estacionamento..." />
      </label>
      <label>
        <span>Data desejada</span>
        <input name="date" type="date" required />
      </label>
      <label>
        <span>Período</span>
        <select name="period" required>
          <option value="">Selecione</option>
          <option value="Manhã">Manhã</option>
          <option value="Tarde">Tarde</option>
          <option value="A combinar">A combinar</option>
        </select>
      </label>
    </form>
  `;
  chatBody.append(container);

  const form = container.querySelector(".schedule-form");
  const zipInput = form.elements.zipCode;

  function fillAddressFields(address) {
    form.elements.street.value = address?.street || "";
    form.elements.neighborhood.value = address?.neighborhood || "";
    form.elements.addressCity.value = address?.city || "";
    form.elements.addressState.value = address?.state || "";
  }

  zipInput.addEventListener("blur", async () => {
    const address = await lookupZipCode(zipInput.value);
    if (!address) {
      answers.address = null;
      fillAddressFields(null);
      return;
    }

    answers.zipCode = address.zipCode;
    answers.address = address;
    if (!answers.city && address.city) {
      answers.city = `${address.city}, ${address.state}`;
    }
    zipInput.value = address.zipCode;
    fillAddressFields(address);
  });

  async function confirmSchedule() {
    if (!form.reportValidity()) {
      return;
    }

    const formData = new FormData(form);
    const address = answers.address || (await lookupZipCode(formData.get("zipCode")));
    answers.name = String(formData.get("name") || "").trim();
    answers.email = String(formData.get("email") || "").trim();
    answers.phone = String(formData.get("phone") || "").trim();
    answers.zipCode = formatZipCode(formData.get("zipCode"));
    answers.address = address;
    answers.addressNumber = String(formData.get("addressNumber") || "").trim();
    answers.addressComplement = String(formData.get("addressComplement") || "").trim();
    answers.addressReference = String(formData.get("addressReference") || "").trim();
    answers.visitDate = String(formData.get("date") || "").trim();
    answers.visitPeriod = String(formData.get("period") || "").trim();
    submitScheduleEmail();
    container.remove();
    quickReplies.innerHTML = "";
    addMessage(
      `Solicitação recebida, ${answers.name}. Vamos confirmar a visita técnica para ${answers.visitDate} no período ${answers.visitPeriod} em ${formatAddress(answers.address)}${answers.addressNumber ? `, ${answers.addressNumber}` : ""}.`
    );
  }

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    await confirmSchedule();
  });

  quickReplies.innerHTML = "";
  const confirmButton = document.createElement("button");
  confirmButton.type = "button";
  confirmButton.className = "reply-button whatsapp schedule-button";
  confirmButton.textContent = "Confirmar solicitação de visita";
  confirmButton.addEventListener("click", confirmSchedule);
  quickReplies.append(confirmButton);

  chatBody.scrollTop = container.offsetTop - chatBody.offsetTop;
}

function finishChat() {
  if (!getQuote().qualifiedOpportunity) {
    addDisqualifiedSummary();
    quickReplies.innerHTML = "";

    const restart = document.createElement("button");
    restart.type = "button";
    restart.className = "reply-button";
    restart.textContent = "Nova cotação";
    restart.addEventListener("click", restartChat);
    quickReplies.append(restart);
    return;
  }

  addSummary();
  quickReplies.innerHTML = "";

  const button = document.createElement("button");
  button.type = "button";
  button.className = "reply-button whatsapp schedule-button";
  button.textContent = "Agendar visita técnica";
  button.addEventListener("click", showNativeSchedule);
  quickReplies.append(button);
}

async function restartChat() {
  stepIndex = 0;
  answers = {};
  autoScrollMessages = false;
  chatBody.innerHTML = "";
  quickReplies.innerHTML = "";
  try {
    await loadPriceCache();
  } catch (error) {
    addMessage("Não consegui carregar o cache de preços local. Rode a página por um servidor a partir da pasta windsurf para liberar o arquivo JSON.");
  }
  const handled = await showInitialHistory();
  if (!handled) {
    nextUnansweredStep();
    askCurrentStep();
  }
  chatBody.scrollTop = 0;
  window.scrollTo(0, 0);
  autoScrollMessages = true;
}

async function pollForNewLead() {
  if (pollingLead) {
    return;
  }
  pollingLead = true;
  try {
    const rows = await loadSheetHistory(true);
    const latestRow = getLatestRow(rows);
    const signature = latestRow ? rowSignature(latestRow) : "";
    if (signature && signature !== latestLeadSignature) {
      latestLeadSignature = signature;
      answers = {};
      chatBody.innerHTML = "";
      quickReplies.innerHTML = "";
      hydrateAnswersFromLatestRow(latestRow);
      if (getLeadReadiness()) {
        askLeadConfirmation();
      } else {
        nextUnansweredStep();
        askCurrentStep();
      }
    }
  } catch (error) {
    // Mantem a tela atual se a consulta pontual ao Google falhar.
  } finally {
    pollingLead = false;
  }
}

window.addEventListener("resize", resizeCanvas);
window.setInterval(pollForNewLead, LEAD_POLL_INTERVAL_MS);

resizeCanvas();
drawLedWall();
restartChat();
