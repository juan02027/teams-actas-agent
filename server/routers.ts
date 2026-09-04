import fs from "node:fs";
import { TRPCError } from "@trpc/server";
import { z } from "zod";
import { COOKIE_NAME } from "@shared/const";
import { getSessionCookieOptions } from "./_core/cookies";
import { systemRouter } from "./_core/systemRouter";
import { publicProcedure, protectedProcedure, router } from "./_core/trpc";
import { processMeetingTranscript } from "./meetingPipeline";
import { generateMeetingDocuments, inferAttendeesFromTranscript } from "./meetingAgent";
import { makeLocalDocuments } from "./localDocuments";
import { decodeRecordingDataUrl, transcribeWithGroq } from "./groq";
import { sendMinutesEmail } from "./email";
import { getAgentSettings, listCommitments, listMeetings, saveAgentSettings, updateMeetingControls } from "./db";
import { clearLocalRecordingReference, createLocalMeeting, deleteLocalCommitment, deleteLocalDocument, getLocalDocument, getLocalMeetingById, getLocalRecording, getLocalSettings, listLocalCommitments, listLocalDocuments, listLocalMeetings, listLocalRecordings, localDataDirectory, saveLocalDocuments, saveLocalRecording, saveLocalSettings, updateLocalCommitment, updateLocalCommitmentStatus, updateLocalMeeting, updateLocalMeetingControls } from "./localStore";

const operatorProcedure = protectedProcedure.use(async ({ ctx, next }) => {
  if (ctx.user.role !== "admin") throw new TRPCError({ code: "FORBIDDEN", message: "Solo el operador autorizado puede administrar este agente." });
  return next();
});
const statusSchema = z.enum(["open", "in_progress", "done", "blocked"]);

async function processStoredRecording(meetingId: number, input: { buffer: Buffer; mimeType: string; title: string; attendees?: Array<{ name: string; email: string; role?: string }> }) {
  const target = getLocalMeetingById(meetingId) || createLocalMeeting({ title: input.title });
  if (!target) throw new TRPCError({ code: "NOT_FOUND", message: "No se encontró la reunión de la grabación." });
  updateLocalMeeting(target.id, { status: "processing", errorMessage: null });
  try {
    const transcript = await transcribeWithGroq({ buffer: input.buffer, mimeType: input.mimeType, title: target.title });
    const attendees = input.attendees?.length ? input.attendees : inferAttendeesFromTranscript(transcript.text);
    const output = await generateMeetingDocuments({ meetingTitle: target.title, transcript: transcript.text, attendees });
    const generated = await makeLocalDocuments({ meetingTitle: target.title, output, attendees });
    const documents = saveLocalDocuments({ meetingId: target.id, documents: generated.documents, commitments: output.commitments });
    updateLocalMeeting(target.id, { status: "review" });
    return { meetingId: target.id, graphMeetingId: target.graphMeetingId, attendees, meetingTitle: target.title, transcript: transcript.text, output, recordingUrl: `/api/local-recordings/${target.id}/download`, documents, dataDirectory: localDataDirectory() };
  } catch (error) {
    updateLocalMeeting(target.id, { status: "error", errorMessage: error instanceof Error ? error.message : "Error de procesamiento" });
    throw new TRPCError({ code: "INTERNAL_SERVER_ERROR", message: error instanceof Error ? error.message : "No se pudo procesar la grabación.", cause: error });
  }
}

export const appRouter = router({
  system: systemRouter,
  auth: router({
    me: publicProcedure.query((opts) => opts.ctx.user),
    logout: publicProcedure.mutation(({ ctx }) => {
      const cookieOptions = getSessionCookieOptions(ctx.req);
      ctx.res.clearCookie(COOKIE_NAME, { ...cookieOptions, maxAge: -1 });
      return { success: true } as const;
    }),
  }),
  agent: router({
    meetings: operatorProcedure.query(async () => { const dbMeetings = await listMeetings(); return dbMeetings.length ? dbMeetings : listLocalMeetings(); }),
    recordings: operatorProcedure.query(() => listLocalRecordings()),
    documents: operatorProcedure.query(() => listLocalDocuments()),
    commitments: operatorProcedure.query(async () => { const local = listLocalCommitments(); return local.length ? local : listCommitments(); }),
    settings: operatorProcedure.query(async () => getLocalSettings() || await getAgentSettings()),
    updateSettings: operatorProcedure.input(z.object({ companyName: z.string().min(1).max(180), recordByDefault: z.boolean(), processByDefault: z.boolean(), requireReview: z.boolean(), destination: z.string().min(1).max(255) })).mutation(async ({ input }) => { saveLocalSettings({ companyName: input.companyName, recordByDefault: Number(input.recordByDefault), processByDefault: Number(input.processByDefault), requireReview: Number(input.requireReview), destination: input.destination, aiModel: "groq/compound-mini" }); return saveAgentSettings({ companyName: input.companyName, recordByDefault: Number(input.recordByDefault), processByDefault: Number(input.processByDefault), requireReview: Number(input.requireReview), destination: input.destination, aiModel: "groq/compound-mini" }); }),
    setMeetingControls: operatorProcedure.input(z.object({ id: z.number().int().positive(), recordingEnabled: z.boolean(), processingEnabled: z.boolean() })).mutation(async ({ input }) => { updateLocalMeetingControls(input.id, input.recordingEnabled, input.processingEnabled); return updateMeetingControls(input.id, input.recordingEnabled, input.processingEnabled); }),
    processTranscript: operatorProcedure.input(z.object({ meetingId: z.number().int().positive().optional(), graphMeetingId: z.string().optional(), title: z.string().optional(), organizerName: z.string().optional(), organizerEmail: z.string().optional(), scheduledAt: z.string().optional(), attendeesCount: z.number().optional(), transcript: z.string().min(20).max(500000) })).mutation(({ input }) => processMeetingTranscript(input)),
    processTextTranscript: operatorProcedure.input(z.object({ title: z.string().min(1).max(255), transcript: z.string().min(20).max(500000) })).mutation(async ({ input }) => { const meeting = createLocalMeeting({ title: input.title }); const output = await generateMeetingDocuments({ meetingTitle: meeting.title, transcript: input.transcript }); const generated = await makeLocalDocuments({ meetingTitle: meeting.title, output }); const documents = saveLocalDocuments({ meetingId: meeting.id, documents: generated.documents, commitments: output.commitments }); return { meetingId: meeting.id, meetingTitle: meeting.title, transcript: input.transcript, output, documents }; }),
    processLocalRecording: operatorProcedure.input(z.object({ audioBase64: z.string().min(100), recordingBase64: z.string().min(100).optional(), mimeType: z.string().min(3).max(120), recordingMimeType: z.string().min(3).max(120).optional(), title: z.string().min(1).max(255), meetingId: z.number().int().positive().optional(), graphMeetingId: z.string().min(1).optional(), durationSeconds: z.number().nonnegative().optional(), attendees: z.array(z.object({ name: z.string(), email: z.string(), role: z.string().optional() })).optional() })).mutation(async ({ input }) => { const audioBuffer = decodeRecordingDataUrl(input.audioBase64); const recordingBuffer = input.recordingBase64 ? decodeRecordingDataUrl(input.recordingBase64) : audioBuffer; const stored = saveLocalRecording({ meetingId: input.meetingId, graphMeetingId: input.graphMeetingId, title: input.title, mimeType: input.recordingMimeType || input.mimeType, buffer: recordingBuffer, processingBuffer: audioBuffer, processingMimeType: input.mimeType, durationSeconds: input.durationSeconds }); return processStoredRecording(stored.meeting.id, { buffer: audioBuffer, mimeType: input.mimeType, title: stored.meeting.title, attendees: input.attendees }); }),
    reprocessRecording: operatorProcedure.input(z.object({ meetingId: z.number().int().positive(), recordingUrl: z.string().optional(), attendees: z.array(z.object({ name: z.string(), email: z.string(), role: z.string().optional() })).optional() })).mutation(async ({ input }) => { const recording = getLocalRecording(input.meetingId); const meeting = getLocalMeetingById(input.meetingId); if (!recording || !meeting) throw new TRPCError({ code: "NOT_FOUND", message: "No hay un archivo local para reprocesar." }); const processingPath = recording.processingFilePath && fs.existsSync(recording.processingFilePath) ? recording.processingFilePath : recording.filePath; return processStoredRecording(input.meetingId, { buffer: fs.readFileSync(processingPath), mimeType: recording.processingMimeType || recording.mimeType, title: meeting.title, attendees: input.attendees }); }),
    clearRecordingReference: operatorProcedure.input(z.object({ meetingId: z.number().int().positive() })).mutation(({ input }) => clearLocalRecordingReference(input.meetingId)),
    deleteDocument: operatorProcedure.input(z.object({ id: z.number().int().positive() })).mutation(({ input }) => { if (!getLocalDocument(input.id)) throw new TRPCError({ code: "NOT_FOUND", message: "El acta ya no existe." }); return deleteLocalDocument(input.id); }),
    updateCommitmentStatus: operatorProcedure.input(z.object({ id: z.number().int().positive(), status: statusSchema })).mutation(({ input }) => { const result = updateLocalCommitmentStatus(input.id, input.status); if (!result) throw new TRPCError({ code: "NOT_FOUND", message: "El compromiso ya no existe." }); return result; }),
    updateCommitment: operatorProcedure.input(z.object({ id: z.number().int().positive(), personName: z.string().min(1).max(180), action: z.string().min(1).max(2000), dueDate: z.string().min(1).max(64) })).mutation(({ input }) => { const result = updateLocalCommitment(input.id, input); if (!result) throw new TRPCError({ code: "NOT_FOUND", message: "El compromiso ya no existe." }); return result; }),
    sendMinutesEmail: operatorProcedure.input(z.object({ recipients: z.array(z.string().email()).min(1).max(100), subject: z.string().min(1).max(200), text: z.string().min(1).max(100000) })).mutation(({ input }) => sendMinutesEmail(input)),
    deleteCommitment: operatorProcedure.input(z.object({ id: z.number().int().positive() })).mutation(({ input }) => deleteLocalCommitment(input.id)),
  }),
});

export type AppRouter = typeof appRouter;
