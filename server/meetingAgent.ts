import { invokeLLM } from "./_core/llm";
import { ENV } from "./_core/env";
import { generateWithGroq } from "./groq";

export const meetingOutputSchema = {
  type: "object",
  properties: {
    executiveSummary: { type: "string" },
    objective: { type: "string" },
    decisions: { type: "array", items: { type: "string" } },
    openTopics: { type: "array", items: { type: "string" } },
    risks: { type: "array", items: { type: "string" } },
    commitments: {
      type: "array",
      items: {
        type: "object",
        properties: {
          personName: { type: "string" },
          personEmail: { type: "string" },
          action: { type: "string" },
          dueDate: { type: "string" },
          evidence: { type: "string" },
          confidence: { type: "string", enum: ["high", "medium", "low"] },
        },
        required: ["personName", "personEmail", "action", "dueDate", "evidence", "confidence"],
        additionalProperties: false,
      },
    },
  },
  required: ["executiveSummary", "objective", "decisions", "openTopics", "risks", "commitments"],
  additionalProperties: false,
} as const;

export type MeetingOutput = {
  executiveSummary: string;
  objective: string;
  decisions: string[];
  openTopics: string[];
  risks: string[];
  commitments: Array<{ personName: string; personEmail: string; action: string; dueDate: string; evidence: string; confidence: "high" | "medium" | "low" }>;
};

const systemPrompt = `Eres un secretario corporativo. Analiza TODA la transcripción y redacta un acta completa, clara y específica. El executiveSummary debe tener varios párrafos o viñetas e incluir los temas realmente tratados, avances, explicaciones relevantes, problemas, decisiones y próximos pasos; no lo reduzcas a una frase. Ignora únicamente saludos, silencios, repeticiones y conversación casual. No inventes información.

COMPROMISOS/TAREAS: extrae cada tarea o compromiso que se haya expresado de forma clara, con su responsable y una evidencia breve de la transcripción. Incluye acciones como entregar, enviar, preparar, validar, revisar, programar o hacer seguimiento cuando estén asignadas explícitamente. Si se menciona una fecha o plazo, consérvalo; si no se menciona, usa dueDate = "Por definir" para que el operador pueda completarlo. No conviertas preguntas, deseos, opiniones o temas generales en tareas. Describe la acción con suficiente detalle para que otra persona pueda ejecutarla. Devuelve únicamente JSON que cumpla el esquema. `;

const MAX_TRANSCRIPT_CHARS = Number(process.env.GROQ_MAX_TRANSCRIPT_CHARS || 60000);
function compactTranscript(transcript: string) {
  // Mantener el contenido completo evita que el acta quede reducida solo a
  // las líneas que contienen palabras como "tarea" o "compromiso".
  return transcript.length <= MAX_TRANSCRIPT_CHARS ? transcript : `${transcript.slice(0, MAX_TRANSCRIPT_CHARS)}\n[Transcripción adicional omitida por límite de tamaño.]`;
}

function textFromContent(content: unknown) {
  if (typeof content === "string") return content;
  if (Array.isArray(content)) return content.map((part) => typeof part === "string" ? part : part && typeof part === "object" && "text" in part ? String((part as { text: unknown }).text) : "").join("");
  return "{}";
}
function stringArray(value: unknown) { return Array.isArray(value) ? value.filter((item): item is string => typeof item === "string" && item.trim().length > 0).map((item) => item.trim()) : []; }
function fallbackCommitments(transcript: string, attendees: Array<{ name: string; email: string; role?: string }> = []) {
  const actionPattern = /\b(debe(?:mos|rán)?|se debe|quedó en|acordamos|compromiso|tarea|entregar|enviar|preparar|validar|revisar|programar|coordinar|documentar|actualizar|completar|hacer seguimiento|dar seguimiento|agendar|compartir|elaborar|definir)\b/i;
  const datePattern = /\b(hoy|mañana|lunes|martes|miércoles|jueves|viernes|sábado|domingo|esta semana|la próxima semana|antes del [^,.;\n]+|el [^,.;\n]+|\d{1,2}[/-]\d{1,2}(?:[/-]\d{2,4})?|\d{1,2}\s+de\s+[a-záéíóú]+)\b/i;
  const seen = new Set<string>();
  return transcript.split(/\r?\n|(?<=[.!?])\s+/).flatMap((line) => {
    const clean = line.replace(/^[-*•\d.)\s]+/, "").trim();
    if (clean.length < 12 || !actionPattern.test(clean)) return [];
    const key = clean.toLowerCase(); if (seen.has(key)) return []; seen.add(key);
    const colon = clean.indexOf(":");
    const possibleName = colon > 0 ? clean.slice(0, colon).trim() : "";
    const attendee = attendees.find((person) => possibleName && person.name.toLowerCase().includes(possibleName.toLowerCase()));
    const personName = attendee?.name || (possibleName && possibleName.split(" ").length <= 5 ? possibleName : "Por definir");
    const dueDate = clean.match(datePattern)?.[0] || "Por definir";
    const action = (colon > 0 ? clean.slice(colon + 1) : clean).trim();
    if (action.length < 8) return [];
    return [{ personName, personEmail: attendee?.email || "", action, dueDate, evidence: clean, confidence: "medium" as const }];
  }).slice(0, 20);
}
function normalizeOutput(raw: Record<string, unknown>, transcript = "", attendees: Array<{ name: string; email: string; role?: string }> = []): MeetingOutput {
  const rawCommitments = Array.isArray(raw.commitments) ? raw.commitments : [];
  const commitments = rawCommitments.flatMap((item) => {
    if (!item || typeof item !== "object") return [];
    const candidate = item as Record<string, unknown>;
    const rawPersonName = typeof candidate.personName === "string" ? candidate.personName.trim() : "";
    const action = typeof candidate.action === "string" ? candidate.action.trim() : "";
    const dueDate = typeof candidate.dueDate === "string" ? candidate.dueDate.trim() : "";
    const evidence = typeof candidate.evidence === "string" ? candidate.evidence.trim() : "";
    if (!action || /^(no definido|no definida|n\/a|none|null)$/i.test(rawPersonName)) return [];
    const confidence: "high" | "medium" | "low" = candidate.confidence === "high" || candidate.confidence === "medium" || candidate.confidence === "low" ? candidate.confidence : "medium";
    const sourceText = `${rawPersonName} ${action} ${evidence}`.toLowerCase();
    const attendee = attendees.find((person) => person.name.trim().length > 3 && sourceText.includes(person.name.toLowerCase()));
    const personName = attendee?.name || rawPersonName || "Por definir";
    const personEmail = attendee?.email || (typeof candidate.personEmail === "string" ? candidate.personEmail.trim() : "");
    return [{ personName, personEmail, action, dueDate: dueDate || "Por definir", evidence: evidence || action, confidence }];
  });
  const finalCommitments = commitments.length ? commitments : fallbackCommitments(transcript, attendees);
  return {
    executiveSummary: typeof raw.executiveSummary === "string" && raw.executiveSummary.trim() ? raw.executiveSummary.trim() : "No se identificaron notas relevantes en la transcripción.",
    objective: typeof raw.objective === "string" && raw.objective.trim() ? raw.objective.trim() : "No se identificó un objetivo explícito en la transcripción.",
    decisions: stringArray(raw.decisions).slice(0, 8), openTopics: stringArray(raw.openTopics).slice(0, 8), risks: stringArray(raw.risks).slice(0, 8), commitments: finalCommitments.slice(0, 20),
  };
}

export async function generateMeetingDocuments(input: { meetingTitle: string; transcript: string; attendees?: Array<{ name: string; email: string; role?: string }> }) {
  const attendeeText = input.attendees?.length ? input.attendees.map((person) => `${person.name} <${person.email}>${person.role ? ` (${person.role})` : ""}`).join("\n") : "[No hay asistentes disponibles]";
  const user = `Título: ${input.meetingTitle}\n\nAsistentes invitados (inclúyelos en el acta, aunque no hayan hablado):\n${attendeeText}\n\nTranscripción:\n${compactTranscript(input.transcript)}`;
  if (ENV.groqApiKey) return normalizeOutput(await generateWithGroq({ system: systemPrompt, user, schema: meetingOutputSchema }), input.transcript, input.attendees);

  const response = await invokeLLM({
    model: "gpt-5-mini",
    messages: [{ role: "system", content: `${systemPrompt} Si un campo no aplica, usa un arreglo vacío.` }, { role: "user", content: user }],
    response_format: { type: "json_schema", json_schema: { name: "teams_meeting_documents", strict: true, schema: meetingOutputSchema } },
    maxTokens: 8000,
  });
  return normalizeOutput(JSON.parse(textFromContent(response.choices?.[0]?.message?.content)) as Record<string, unknown>, input.transcript, input.attendees);
}
