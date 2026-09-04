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

const systemPrompt = `Eres un secretario corporativo preciso. Analiza únicamente la conversación relacionada con compromisos, tareas asignadas, responsables, fechas de entrega y el contexto inmediato de esas acciones. Ignora saludos, opiniones, conversaciones generales, explicaciones que no cambien el compromiso y cualquier tema no relacionado. El resumen, decisiones, temas abiertos y riesgos deben referirse solo a esos compromisos.

REGLA OBLIGATORIA PARA COMPROMISOS: incluye un elemento en commitments únicamente cuando la transcripción diga claramente que existe un compromiso, tarea asignada, acción acordada o promesa de entrega, Y también aparezca una fecha o plazo concreto. Cada elemento debe tener un responsable identificable y dueDate no vacío. Si falta responsable, fecha o evidencia explícita, NO lo incluyas. No conviertas opiniones, deseos, preguntas, ideas, tareas genéricas ni acuerdos sin fecha en compromisos. Conserva una cita breve como evidence. Devuelve únicamente JSON que cumpla el esquema.`;

const MAX_TRANSCRIPT_CHARS = Number(process.env.GROQ_MAX_TRANSCRIPT_CHARS || 16000);
const COMMITMENT_CONTEXT = /comprom|tarea|acción|accion|entregar|entrega|enviar|preparar|validar|revisar|asign|responsable|dueño|dueno|fecha|plazo|vence|viernes|lunes|martes|miércoles|miercoles|jueves|sábado|sabado|domingo|antes del|para el|pendiente|acordamos|acuerdo|decidimos|decisión|decision|seguimiento|área|area/i;
function compactTranscript(transcript: string) {
  const lines = transcript.split(/\r?\n/).filter(Boolean);
  const relevantIndexes = new Set<number>();
  lines.forEach((line, index) => {
    if (COMMITMENT_CONTEXT.test(line)) {
      // Keep one adjacent line before and after each hit so the owner/action
      // pair is not separated from its date or assignment.
      if (index > 0) relevantIndexes.add(index - 1);
      relevantIndexes.add(index);
      if (index + 1 < lines.length) relevantIndexes.add(index + 1);
    }
  });
  const relevant = Array.from(relevantIndexes).sort((a, b) => a - b).map((index) => lines[index]).join("\n");
  if (!relevant) return "[No se encontraron frases relacionadas con compromisos, tareas asignadas, responsables o fechas de entrega.]";
  return relevant.length <= MAX_TRANSCRIPT_CHARS ? relevant : `${relevant.slice(0, MAX_TRANSCRIPT_CHARS)}\n[Contexto adicional omitido por límite de consumo.]`;
}

function textFromContent(content: unknown) {
  if (typeof content === "string") return content;
  if (Array.isArray(content)) return content.map((part) => typeof part === "string" ? part : part && typeof part === "object" && "text" in part ? String((part as { text: unknown }).text) : "").join("");
  return "{}";
}
function stringArray(value: unknown) { return Array.isArray(value) ? value.filter((item): item is string => typeof item === "string" && item.trim().length > 0).map((item) => item.trim()) : []; }
function normalizeOutput(raw: Record<string, unknown>): MeetingOutput {
  const rawCommitments = Array.isArray(raw.commitments) ? raw.commitments : [];
  const commitments = rawCommitments.flatMap((item) => {
    if (!item || typeof item !== "object") return [];
    const candidate = item as Record<string, unknown>;
    const personName = typeof candidate.personName === "string" ? candidate.personName.trim() : "";
    const action = typeof candidate.action === "string" ? candidate.action.trim() : "";
    const dueDate = typeof candidate.dueDate === "string" ? candidate.dueDate.trim() : "";
    const evidence = typeof candidate.evidence === "string" ? candidate.evidence.trim() : "";
    if (!personName || !action || !dueDate || !evidence || /^(no definido|no definida|n\/a|none|null)$/i.test(personName) || /^(no definido|no definida|n\/a|none|null)$/i.test(dueDate)) return [];
    const confidence: "high" | "medium" | "low" = candidate.confidence === "high" || candidate.confidence === "medium" || candidate.confidence === "low" ? candidate.confidence : "medium";
    return [{ personName, personEmail: typeof candidate.personEmail === "string" ? candidate.personEmail.trim() : "", action, dueDate, evidence, confidence }];
  });
  return {
    executiveSummary: typeof raw.executiveSummary === "string" ? raw.executiveSummary.trim() : "No se pudo generar un resumen ejecutivo.",
    objective: typeof raw.objective === "string" ? raw.objective.trim() : "No definido",
    decisions: stringArray(raw.decisions).slice(0, 8), openTopics: stringArray(raw.openTopics).slice(0, 8), risks: stringArray(raw.risks).slice(0, 8), commitments: commitments.slice(0, 20),
  };
}

export async function generateMeetingDocuments(input: { meetingTitle: string; transcript: string }) {
  const user = `Título: ${input.meetingTitle}\n\nTranscripción:\n${compactTranscript(input.transcript)}`;
  if (ENV.groqApiKey) return normalizeOutput(await generateWithGroq({ system: systemPrompt, user, schema: meetingOutputSchema }));

  const response = await invokeLLM({
    model: "gpt-5-mini",
    messages: [{ role: "system", content: `${systemPrompt} Si un campo no aplica, usa un arreglo vacío.` }, { role: "user", content: user }],
    response_format: { type: "json_schema", json_schema: { name: "teams_meeting_documents", strict: true, schema: meetingOutputSchema } },
    maxTokens: 8000,
  });
  return normalizeOutput(JSON.parse(textFromContent(response.choices?.[0]?.message?.content)) as Record<string, unknown>);
}
