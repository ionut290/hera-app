"use strict";

const { authenticateEvent } = require("./_firebase-token");

const MAX_IMAGE_BYTES = 4 * 1024 * 1024;
const MAX_TEXT_LENGTH = 24000;
const DEFAULT_GEMINI_MODELS = ["gemini-2.5-flash", "gemini-2.5-flash-lite"];

function json(statusCode, payload) {
  return {
    statusCode,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Authorization, Content-Type"
    },
    body: JSON.stringify(payload)
  };
}

function cleanText(value, maxLength = 240) {
  return String(value || "").replace(/[\u0000-\u001f\u007f]/g, " ").replace(/\s+/g, " ").trim().slice(0, maxLength);
}

function parseBody(event) {
  let parsed;
  try {
    parsed = JSON.parse(event.body || "{}");
  } catch (_) {
    throw new Error("Dati della richiesta non validi.");
  }
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) throw new Error("Dati della richiesta non validi.");
  return parsed;
}

function configured() {
  return {
    plantnet: Boolean(String(process.env.PLANTNET_API_KEY || "").trim()),
    trefle: Boolean(String(process.env.TREFLE_API_TOKEN || "").trim()),
    gemini: Boolean(String(process.env.GEMINI_API_KEY || "").trim())
  };
}

function requireSecret(name, label) {
  const value = String(process.env[name] || "").trim();
  if (!value) {
    const error = new Error(`${label} non configurato. Aggiungi ${name} nelle variabili protette di Netlify.`);
    error.statusCode = 503;
    error.code = "PROVIDER_NOT_CONFIGURED";
    throw error;
  }
  return value;
}

function parseImage(value) {
  const input = String(value || "");
  const match = input.match(/^data:(image\/(?:jpeg|png));base64,([A-Za-z0-9+/=\r\n]+)$/i);
  if (!match) throw new Error("Immagine non valida. Usa una fotografia JPG o PNG.");
  const buffer = Buffer.from(match[2].replace(/\s/g, ""), "base64");
  if (!buffer.length || buffer.length > MAX_IMAGE_BYTES) throw new Error("La fotografia deve pesare meno di 4 MB dopo la compressione.");
  return { buffer, mimeType: match[1].toLowerCase(), extension: match[1].toLowerCase().includes("png") ? "png" : "jpg" };
}

async function fetchJson(url, options = {}, timeoutMs = 30000) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(url, { ...options, signal: controller.signal });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      const message = cleanText(data?.message || data?.error?.message || data?.error || `Servizio esterno HTTP ${response.status}`, 500);
      const error = new Error(message || `Servizio esterno HTTP ${response.status}`);
      error.statusCode = response.status === 429 ? 429 : 502;
      error.providerStatus = response.status;
      throw error;
    }
    return data;
  } finally {
    clearTimeout(timeout);
  }
}

async function plantNetRequest(body, disease = false) {
  const apiKey = requireSecret("PLANTNET_API_KEY", "Pl@ntNet");
  const image = parseImage(body.image);
  const organ = ["leaf", "flower", "fruit", "bark", "auto"].includes(body.organ) ? body.organ : "auto";
  const form = new FormData();
  form.append(disease ? "image" : "images", new Blob([image.buffer], { type: image.mimeType }), `foto.${image.extension}`);
  form.append("organs", organ);
  const endpoint = disease ? "diseases/identify" : "identify/all";
  const params = new URLSearchParams({
    "api-key": apiKey,
    lang: "it",
    "nb-results": "5",
    "include-related-images": "true"
  });
  const data = await fetchJson(`https://my-api.plantnet.org/v2/${endpoint}?${params}`, {
    method: "POST",
    body: form,
    headers: { Accept: "application/json" }
  }, 45000);
  if (disease) {
    return {
      source: "Pl@ntNet Diseases",
      limitedCoverage: true,
      remainingRequests: Number.isFinite(data.remainingIdentificationRequests) ? data.remainingIdentificationRequests : null,
      results: (Array.isArray(data.results) ? data.results : []).slice(0, 5).map((item) => ({
        code: cleanText(item.name, 80),
        name: cleanText(item.description || item.name, 240),
        score: Number(item.score || 0),
        image: cleanText(item.images?.[0]?.url?.m || item.images?.[0]?.url?.s || "", 1000),
        citation: cleanText(item.images?.[0]?.citation || "Pl@ntNet", 500)
      }))
    };
  }
  return {
    source: "Pl@ntNet",
    bestMatch: cleanText(data.bestMatch, 240),
    remainingRequests: Number.isFinite(data.remainingIdentificationRequests) ? data.remainingIdentificationRequests : null,
    results: (Array.isArray(data.results) ? data.results : []).slice(0, 5).map((item) => ({
      score: Number(item.score || 0),
      scientificName: cleanText(item.species?.scientificNameWithoutAuthor || item.species?.scientificName || "", 240),
      fullScientificName: cleanText(item.species?.scientificName || "", 300),
      commonNames: (Array.isArray(item.species?.commonNames) ? item.species.commonNames : []).map((name) => cleanText(name, 160)).filter(Boolean).slice(0, 8),
      family: cleanText(item.species?.family?.scientificNameWithoutAuthor || item.species?.family?.scientificName || "", 160),
      genus: cleanText(item.species?.genus?.scientificNameWithoutAuthor || item.species?.genus?.scientificName || "", 160),
      image: cleanText(item.images?.[0]?.url?.m || item.images?.[0]?.url?.s || "", 1000),
      citation: cleanText(item.images?.[0]?.citation || "Pl@ntNet", 500)
    }))
  };
}

function pickTreflePlant(item) {
  return {
    id: Number(item?.id || 0) || null,
    slug: cleanText(item?.slug, 240),
    scientificName: cleanText(item?.scientific_name, 240),
    commonName: cleanText(item?.common_name, 240),
    family: cleanText(item?.family, 180),
    genus: cleanText(item?.genus, 180),
    image: cleanText(item?.image_url, 1000),
    status: cleanText(item?.status, 80)
  };
}

async function searchTrefle(body) {
  const token = requireSecret("TREFLE_API_TOKEN", "Trefle");
  const query = cleanText(body.query, 180);
  if (query.length < 2) throw new Error("Inserisci almeno due caratteri per cercare la pianta.");
  const params = new URLSearchParams({ token, q: query });
  const data = await fetchJson(`https://trefle.io/api/v1/plants/search?${params}`, {
    headers: { Accept: "application/json" }
  }, 20000);
  return {
    source: "Trefle",
    attributionRequired: true,
    results: (Array.isArray(data.data) ? data.data : []).slice(0, 10).map(pickTreflePlant)
  };
}

async function trefleDetails(body) {
  const token = requireSecret("TREFLE_API_TOKEN", "Trefle");
  const slug = cleanText(body.slug, 240).toLowerCase();
  if (!/^[a-z0-9-]+$/.test(slug)) throw new Error("Identificativo della pianta non valido.");
  const data = await fetchJson(`https://trefle.io/api/v1/plants/${encodeURIComponent(slug)}?token=${encodeURIComponent(token)}`, {
    headers: { Accept: "application/json" }
  }, 20000);
  const item = data?.data || {};
  return {
    source: "Trefle",
    attributionRequired: true,
    plant: {
      ...pickTreflePlant(item),
      duration: Array.isArray(item.duration) ? item.duration.map((value) => cleanText(value, 100)).filter(Boolean) : [],
      edible: typeof item.edible === "boolean" ? item.edible : null,
      edibleParts: Array.isArray(item.edible_part) ? item.edible_part.map((value) => cleanText(value, 100)).filter(Boolean) : [],
      observations: cleanText(item.observations, 1600),
      commonNames: item.common_names && typeof item.common_names === "object" ? item.common_names : {},
      specifications: item.specifications && typeof item.specifications === "object" ? item.specifications : {},
      growth: item.growth && typeof item.growth === "object" ? item.growth : {},
      flower: item.flower && typeof item.flower === "object" ? item.flower : {},
      foliage: item.foliage && typeof item.foliage === "object" ? item.foliage : {},
      sources: (Array.isArray(item.sources) ? item.sources : []).slice(0, 10).map((source) => ({
        name: cleanText(source.name, 180),
        citation: cleanText(source.citation, 500),
        url: cleanText(source.url, 1000)
      }))
    }
  };
}

function geminiModels() {
  const configuredModel = cleanText(process.env.GEMINI_MODEL, 120);
  return [...new Set([configuredModel, ...DEFAULT_GEMINI_MODELS].filter(Boolean))];
}

function equipmentPrompt(body) {
  const equipment = {
    type: cleanText(body.type, 120),
    brand: cleanText(body.brand, 120),
    model: cleanText(body.model, 160),
    year: cleanText(body.year, 20),
    serial: cleanText(body.serial, 120),
    manualText: String(body.manualText || "").slice(0, MAX_TEXT_LENGTH)
  };
  if (!equipment.brand || !equipment.model) throw new Error("Inserisci almeno marca e modello del mezzo o utensile.");
  return `Sei un assistente tecnico per mezzi e utensili usati nella manutenzione del verde.\n
Prepara una scheda tecnica prudente in italiano per questo elemento:\n${JSON.stringify(equipment, null, 2)}\n
Regole obbligatorie:\n
- Non inventare valori tecnici. Se un dato non e certo, usa null e stato \"da_verificare\".\n
- Se e presente testo del manuale, usalo come fonte primaria e marca i dati trovati come \"confermato_manual\".\n
- Senza manuale, qualunque miscela, olio, capacita, ricambio o intervallo deve essere \"da_verificare\", anche se sembra noto.\n
- Non dichiarare di avere consultato internet o un manuale non fornito.\n
- Non proporre modifiche ai dispositivi di sicurezza.\n
- Le istruzioni devono restare generali e invitare a seguire il manuale ufficiale.\n
- Rispondi esclusivamente nel JSON richiesto.`;
}

const equipmentSchema = {
  type: "object",
  properties: {
    identification: {
      type: "object",
      properties: {
        type: { type: "string", nullable: true },
        brand: { type: "string", nullable: true },
        model: { type: "string", nullable: true },
        probableVariant: { type: "string", nullable: true },
        summary: { type: "string", nullable: true }
      }
    },
    technicalData: {
      type: "array",
      items: {
        type: "object",
        properties: {
          label: { type: "string" },
          value: { type: "string", nullable: true },
          status: { type: "string", enum: ["confermato_manual", "da_verificare", "non_disponibile"] },
          sourceNote: { type: "string", nullable: true }
        },
        required: ["label", "value", "status", "sourceNote"]
      }
    },
    maintenance: { type: "array", items: { type: "string" } },
    safety: { type: "array", items: { type: "string" } },
    commonChecks: { type: "array", items: { type: "string" } },
    missingInformation: { type: "array", items: { type: "string" } },
    warning: { type: "string" }
  },
  required: ["identification", "technicalData", "maintenance", "safety", "commonChecks", "missingInformation", "warning"]
};

function normalizeGeminiResult(value, hasManual) {
  const result = value && typeof value === "object" ? value : {};
  const identification = result.identification && typeof result.identification === "object" ? result.identification : {};
  return {
    source: "Google Gemini",
    modelGenerated: true,
    manualProvided: hasManual,
    identification: {
      type: cleanText(identification.type, 160) || null,
      brand: cleanText(identification.brand, 160) || null,
      model: cleanText(identification.model, 200) || null,
      probableVariant: cleanText(identification.probableVariant, 240) || null,
      summary: cleanText(identification.summary, 1200) || null
    },
    technicalData: (Array.isArray(result.technicalData) ? result.technicalData : []).slice(0, 30).map((item) => {
      const requestedStatus = ["confermato_manual", "da_verificare", "non_disponibile"].includes(item?.status) ? item.status : "da_verificare";
      return {
        label: cleanText(item?.label, 180),
        value: cleanText(item?.value, 500) || null,
        status: !hasManual && requestedStatus === "confermato_manual" ? "da_verificare" : requestedStatus,
        sourceNote: cleanText(item?.sourceNote, 500) || null
      };
    }).filter((item) => item.label),
    maintenance: (Array.isArray(result.maintenance) ? result.maintenance : []).slice(0, 20).map((item) => cleanText(item, 500)).filter(Boolean),
    safety: (Array.isArray(result.safety) ? result.safety : []).slice(0, 20).map((item) => cleanText(item, 500)).filter(Boolean),
    commonChecks: (Array.isArray(result.commonChecks) ? result.commonChecks : []).slice(0, 20).map((item) => cleanText(item, 500)).filter(Boolean),
    missingInformation: (Array.isArray(result.missingInformation) ? result.missingInformation : []).slice(0, 20).map((item) => cleanText(item, 300)).filter(Boolean),
    warning: cleanText(result.warning, 800) || "Verifica sempre i dati tecnici sul manuale ufficiale del modello esatto."
  };
}

async function equipmentInfo(body) {
  const apiKey = requireSecret("GEMINI_API_KEY", "Gemini");
  const prompt = equipmentPrompt(body);
  const image = body.image ? parseImage(body.image) : null;
  const parts = [{ text: prompt }];
  if (image) parts.push({ inlineData: { mimeType: image.mimeType, data: image.buffer.toString("base64") } });
  let lastError;
  for (const model of geminiModels()) {
    try {
      const data = await fetchJson(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(apiKey)}`, {
        method: "POST",
        headers: { Accept: "application/json", "Content-Type": "application/json" },
        body: JSON.stringify({
          contents: [{ role: "user", parts }],
          generationConfig: {
            temperature: 0.1,
            maxOutputTokens: 4096,
            responseMimeType: "application/json",
            responseSchema: equipmentSchema
          }
        })
      }, 45000);
      const output = data?.candidates?.[0]?.content?.parts?.map((part) => part.text || "").join("") || "";
      if (!output) throw new Error("Gemini non ha restituito una scheda utilizzabile.");
      let parsed;
      try { parsed = JSON.parse(output); } catch (_) { throw new Error("Gemini ha restituito una scheda non valida."); }
      return { ...normalizeGeminiResult(parsed, Boolean(cleanText(body.manualText, MAX_TEXT_LENGTH))), model };
    } catch (error) {
      lastError = error;
      if (![400, 404, 429, 502].includes(Number(error?.statusCode || 0))) throw error;
    }
  }
  throw lastError || new Error("Gemini non disponibile.");
}

exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") return json(204, {});
  if (event.httpMethod !== "POST") return json(405, { ok: false, error: "Metodo non consentito." });
  try {
    const user = await authenticateEvent(event);
    const body = parseBody(event);
    const action = cleanText(body.action, 80);
    if (action === "status") {
      return json(200, {
        ok: true,
        configured: configured(),
        freeOnly: true,
        authenticatedUid: user.sub,
        notice: "Le API sono configurate per usare esclusivamente le rispettive quote gratuite."
      });
    }
    let result;
    if (action === "identifyPlant") result = await plantNetRequest(body, false);
    else if (action === "identifyDisease") result = await plantNetRequest(body, true);
    else if (action === "searchPlant") result = await searchTrefle(body);
    else if (action === "plantDetails") result = await trefleDetails(body);
    else if (action === "equipmentInfo") result = await equipmentInfo(body);
    else return json(400, { ok: false, error: "Azione assistente non riconosciuta." });
    return json(200, { ok: true, result });
  } catch (error) {
    const statusCode = Number(error?.statusCode || 0) || (/Token Firebase|Sessione Firebase|Utente Firebase/i.test(error?.message || "") ? 401 : 500);
    console.error("green-assistant:", error?.code || error?.message || error);
    return json(statusCode, {
      ok: false,
      code: error?.code || (statusCode === 429 ? "FREE_QUOTA_REACHED" : "GREEN_ASSISTANT_ERROR"),
      error: cleanText(error?.message || "Assistente temporaneamente non disponibile.", 600)
    });
  }
};
