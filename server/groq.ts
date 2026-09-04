import { ENV } from "./_core/env";

const GROQ_BASE_URL = "https://api.groq.com/openai/v1";
const MAX_RECORDING_BYTES = 180 * 1024 * 1024;

function groqKey() {
  if (!ENV.groqApiKey) throw new Error("GROQ_API_KEY no está configurada. Añádela al archivo .env local.");
  return ENV.groqApiKey;
}
function extensionForMime(mimeType: string) {
  if (mimeType.includes("mp4")) return "m4a";
  if (mimeType.includes("mpeg") || mimeType.includes("mp3")) return "mp3";
  if (mimeType.includes("wav")) return "wav";
  if (mimeType.includes("ogg")) return "ogg";
  return "webm";
}
export function decodeRecordingDataUrl(value: string) {
  const comma = value.indexOf(",");
  const encoded = comma >= 0 ? value.slice(comma + 1) : value;
  if (!encoded || encoded.length < 100) throw new Error("La grabación recibida está vacía o incompleta.");
  const buffer = Buffer.from(encoded, "base64");
  if (buffer.length < 100) throw new Error("La grabación recibida está vacía o incompleta.");
  if (buffer.length > MAX_RECORDING_BYTES) throw new Error("La grabación supera el límite local de 180 MB.");
  return buffer;
}

export async function transcribeWithGroq(input: { buffer: Buffer; mimeType: string; language?: string; title?: string }) {
  if (ENV.deepgramApiKey) return transcribeWithDeepgram(input);
  const form = new FormData();
  const extension = extensionForMime(input.mimeType);
  form.append("file", new Blob([new Uint8Array(input.buffer)], { type: input.mimeType }), `teams-recording.${extension}`);
  form.append("model", ENV.groqTranscriptionModel);
  // The minutes pipeline only needs the text. Avoid returning segment
  // timestamps and metadata that are never sent to the summarizer.
  form.append("response_format", "json");
  form.append("language", input.language || "es");
  form.append("temperature", "0");
  form.append("prompt", `Transcripción de una reunión corporativa en español. Título: ${input.title || "Reunión de Teams"}. Conserva nombres, fechas, responsables y frases de compromiso.`);

  const response = await fetch(`${GROQ_BASE_URL}/audio/transcriptions`, {
    method: "POST",
    headers: { Authorization: `Bearer ${groqKey()}` },
    body: form,
  });
  if (!response.ok) {
    const detail = await response.text().catch(() => "");
    if (response.status === 413) {
      throw new Error("Groq rechazó el archivo por tamaño. La grabación local quedó guardada en Grabaciones. Para procesar una grabación de video grande, usa Modo largo · solo audio o divide la reunión en partes menores antes de volver a procesarla.");
    }
    if (response.status === 429) {
      const retry = detail.match(/try again in\s+([\d.]+)\s*m(?:in(?:ute)?s?)?\s*([\d.]+)?\s*s?/i);
      const secondsOnly = detail.match(/try again in\s+([\d.]+)\s*s/i);
      const waitLabel = retry ? `${retry[1]} min${retry[2] ? ` ${retry[2]} s` : ""}` : secondsOnly ? `${secondsOnly[1]} s` : "el tiempo indicado por Groq";
      throw new Error(`Groq alcanzó la cuota horaria de audio (7.200 segundos por hora). Esta organización ya usó parte de la cuota con otras transcripciones; la reunión no está dañada. La grabación local quedó guardada. Espera aproximadamente ${waitLabel} y pulsa Procesar nuevamente. Para ampliar la cuota debes cambiar el plan o límites de la organización en la consola de Groq. Detalle: ${detail}`);
    }
    throw new Error(`Groq Whisper falló (${response.status}): ${detail}`);
  }
  const payload = await response.json() as { text?: string; language?: string; duration?: number };
  if (!payload.text || payload.text.trim().length < 10) throw new Error("Groq no devolvió una transcripción utilizable.");
  return { text: payload.text, language: payload.language || "es", duration: payload.duration || null };
}

async function transcribeWithDeepgram(input: { buffer: Buffer; mimeType: string; language?: string; title?: string }) {
  const params = new URLSearchParams({ model: ENV.deepgramModel, language: input.language || "es", smart_format: "true", punctuate: "true", diarize: "true", utterances: "true" });
  const response = await fetch(`https://api.deepgram.com/v1/listen?${params}`, { method: "POST", headers: { Authorization: `Token ${ENV.deepgramApiKey}`, "Content-Type": input.mimeType || "audio/webm" }, body: new Uint8Array(input.buffer) });
  if (!response.ok) { const detail = await response.text().catch(() => ""); throw new Error(`Deepgram falló (${response.status}). La grabación quedó guardada; revisa DEEPGRAM_API_KEY. ${detail}`); }
  const payload = await response.json() as { results?: { channels?: Array<{ alternatives?: Array<{ transcript?: string }>; detected_language?: string }>; utterances?: Array<{ speaker?: number; transcript?: string }> }; metadata?: { duration?: number } };
  const channel = payload.results?.channels?.[0];
  const utterances = payload.results?.utterances || [];
  const text = utterances.length ? utterances.map((item) => `[Hablante ${typeof item.speaker === "number" ? item.speaker + 1 : "?"}] ${item.transcript || ""}`).join("\n").trim() : channel?.alternatives?.[0]?.transcript?.trim() || "";
  if (text.length < 10) throw new Error("Deepgram no devolvió una transcripción utilizable. La grabación permanece guardada.");
  return { text, language: channel?.detected_language || input.language || "es", duration: payload.metadata?.duration || null };
}

export async function generateWithGroq(input: { system: string; user: string; schema: Record<string, unknown> }) {
  const key = groqKey();
  // The organization already accepted this model (the previous error was a
  // JSON validation failure, not a model access failure). Do not use the
  // retired llama-3.1-8b-instant identifier here.
  const model = ENV.groqChatModel;
  const body: Record<string, unknown> = {
    model,
    temperature: 0,
    max_completion_tokens: 8192,
    response_format: { type: "json_object" },
    messages: [
      { role: "system", content: `${input.system}\nResponde solo con un objeto JSON válido, sin markdown ni texto adicional. Usa exactamente las claves executiveSummary, objective, decisions, openTopics, risks y commitments. Los arreglos pueden estar vacíos.` },
      { role: "user", content: `${input.user}\n\nDevuelve únicamente JSON válido. Si no hay compromisos que cumplan las reglas, usa commitments: [].` },
    ],
  };
  if (model.startsWith("openai/gpt-oss")) { body.reasoning_effort = "low"; body.reasoning_format = "hidden"; }
  const request = () => fetch(`${GROQ_BASE_URL}/chat/completions`, { method: "POST", headers: { Authorization: `Bearer ${key}`, "Content-Type": "application/json" }, body: JSON.stringify(body) });
  let response = await request();
  if (response.status === 429) {
    const detail = await response.text().catch(() => "");
    const seconds = Number(detail.match(/try again in\s+([\d.]+)\s*s/i)?.[1] || 0);
    const milliseconds = Number(detail.match(/try again in\s+([\d.]+)\s*ms/i)?.[1] || 0);
    await new Promise((resolve) => setTimeout(resolve, Math.min(15000, Math.max(1000, seconds * 1000 || milliseconds || 1000))));
    response = await request();
  }
  if (!response.ok) {
    const detail = await response.text().catch(() => "");
    throw new Error(`Groq no pudo generar el acta (${response.status}). La grabación permanece guardada; pulsa Procesar nuevamente. ${detail}`);
  }
  const payload = await response.json() as { choices?: Array<{ message?: { content?: string | null } }>; usage?: { prompt_tokens?: number; completion_tokens?: number; total_tokens?: number; prompt_tokens_details?: { cached_tokens?: number } } };
  if (payload.usage) console.info(`[Groq] acta modelo=${model} prompt=${payload.usage.prompt_tokens ?? 0} completion=${payload.usage.completion_tokens ?? 0} total=${payload.usage.total_tokens ?? 0}`);
  let content = payload.choices?.[0]?.message?.content?.trim();
  // GPT-OSS can spend the whole completion budget reasoning. Retry once with
  // a larger controlled budget instead of losing the saved recording.
  if (!content) {
    const retryBody = { ...body, max_completion_tokens: 8192 };
    const retryResponse = await fetch(`${GROQ_BASE_URL}/chat/completions`, { method: "POST", headers: { Authorization: `Bearer ${key}`, "Content-Type": "application/json" }, body: JSON.stringify(retryBody) });
    if (retryResponse.ok) {
      const retryPayload = await retryResponse.json() as { choices?: Array<{ message?: { content?: string | null } }> };
      content = retryPayload.choices?.[0]?.message?.content?.trim();
    }
  }
  if (!content) throw new Error("Groq no generó contenido visible. La grabación permanece guardada; pulsa Procesar nuevamente.");
  const jsonStart = content.indexOf("{");
  const jsonEnd = content.lastIndexOf("}");
  if (jsonStart < 0 || jsonEnd <= jsonStart) throw new Error("Groq no devolvió JSON válido. Pulsa Procesar nuevamente; la grabación permanece guardada.");
  try { return JSON.parse(content.slice(jsonStart, jsonEnd + 1)) as Record<string, unknown>; } catch { throw new Error("Groq devolvió JSON incompleto. Pulsa Procesar nuevamente; la grabación permanece guardada."); }
}
