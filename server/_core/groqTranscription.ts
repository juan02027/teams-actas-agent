import { ENV } from "./env";
import type { TranscriptionError, WhisperResponse } from "./voiceTranscription";

export async function transcribeWithGroq(audioBuffer: Buffer, mimeType: string, language = "es", prompt?: string): Promise<WhisperResponse | TranscriptionError> {
  if (!ENV.groqApiKey) return { error: "GROQ_API_KEY no está configurada", code: "SERVICE_ERROR" };
  if (audioBuffer.length > 25 * 1024 * 1024) return { error: "El archivo supera el límite de 25 MB de Groq", code: "FILE_TOO_LARGE" };
  const form = new FormData();
  form.append("file", new Blob([new Uint8Array(audioBuffer)], { type: mimeType }), mimeType.includes("video") ? "reunion.webm" : "reunion.webm");
  form.append("model", "whisper-large-v3-turbo");
  form.append("response_format", "verbose_json");
  form.append("language", language);
  if (prompt) form.append("prompt", prompt);
  const response = await fetch("https://api.groq.com/openai/v1/audio/transcriptions", { method: "POST", headers: { Authorization: `Bearer ${ENV.groqApiKey}` }, body: form });
  if (!response.ok) return { error: "Groq no pudo transcribir la grabación", code: "TRANSCRIPTION_FAILED", details: `${response.status}: ${await response.text()}` };
  const result = await response.json() as WhisperResponse;
  if (!result.text) return { error: "Groq devolvió una transcripción vacía", code: "SERVICE_ERROR" };
  return result;
}
